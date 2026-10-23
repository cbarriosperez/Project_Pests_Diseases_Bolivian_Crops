#!/usr/bin/env python3
"""
Pronóstico climático a 16 días por municipio productor — SAT Plagas y Enfermedades Bolivia.

Versión automatizada (GitHub Actions) del notebook Pronostico_Municipios_Bolivia.ipynb, Pasos 3–9.
Descarga de Open-Meteo precipitación, Tmax, Tmin y humedad relativa diarias para los municipios
productores de los 15 cultivos, en 7 modelos + un multimodelo (promedio simple), y escribe un JSON
compacto que consume la página web (index.html → #explorar → "Pronóstico 16 días").

Entrada : data/geo/municipios_bolivia.geojson   (punto de consulta lat/lon por municipio)
          data/geo/cultivos_municipios.json     (municipios productores por cultivo)
Salida  : data/pronostico/pronostico_latest.json

Uso     : python Scripts/pronostico/pronostico_municipios.py
          python Scripts/pronostico/pronostico_municipios.py --limite-municipios 20   (prueba rápida)

Variable de entorno opcional: OPEN_METEO_API_KEY (plan comercial de Open-Meteo). Si existe, se usan
los endpoints "customer-" con la clave; si no, la API gratuita.
"""
import argparse
import json
import os
import re
import sys
import time
import warnings
from datetime import datetime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

RAIZ = Path(__file__).resolve().parents[2]
RUTA_MUNICIPIOS = RAIZ / "data" / "geo" / "municipios_bolivia.geojson"
RUTA_CULTIVOS = RAIZ / "data" / "geo" / "cultivos_municipios.json"
RUTA_SALIDA = RAIZ / "data" / "pronostico" / "pronostico_latest.json"

# ----------------------------------------------------------------------------------------------
# Configuración (equivale al Paso 3 del notebook)
# ----------------------------------------------------------------------------------------------
CONFIG = {
    "dias_pronostico": 16,
    "timezone": "America/La_Paz",     # Bolivia: UTC-4, sin horario de verano
    "tam_lote": 15,
    "pausa_entre_lotes_seg": 2.0,
    "pausa_entre_modelos_seg": 10.0,
    "intentos_por_lote": 4,
    "decimales": 1,
}

API_KEY = os.environ.get("OPEN_METEO_API_KEY", "").strip()
if API_KEY:
    URL_FORECAST = "https://customer-api.open-meteo.com/v1/forecast"
    URL_ENSEMBLE = "https://customer-ensemble-api.open-meteo.com/v1/ensemble"
else:
    URL_FORECAST = "https://api.open-meteo.com/v1/forecast"
    URL_ENSEMBLE = "https://ensemble-api.open-meteo.com/v1/ensemble"

VARIABLES_DIARIAS = ["precipitation_sum", "temperature_2m_max", "temperature_2m_min"]
VARIABLE_HR = "relative_humidity_2m"
VARIABLES_SALIDA = VARIABLES_DIARIAS + [VARIABLE_HR]
METADATA_VARIABLES = {
    "precipitation_sum": {"nombre": "Precipitación diaria", "unidad": "mm"},
    "temperature_2m_max": {"nombre": "Temperatura máxima", "unidad": "°C"},
    "temperature_2m_min": {"nombre": "Temperatura mínima", "unidad": "°C"},
    "relative_humidity_2m": {"nombre": "Humedad relativa", "unidad": "%"},
}

MODELOS = {
    "GFS":         {"endpoint": "forecast", "candidatos": ["gfs_seamless", "gfs_global"]},
    "GEFS":        {"endpoint": "ensemble", "candidatos": ["gfs_seamless", "gfs025", "gfs05"]},
    "ECMWF-ENS":   {"endpoint": "ensemble", "candidatos": ["ecmwf_ifs025", "ecmwf_ifs04"]},
    "AIFS-Single": {"endpoint": "forecast", "candidatos": ["ecmwf_aifs025_single", "ecmwf_aifs025", "ecmwf_aifs_0_25"]},
    "AIFS-ENS":    {"endpoint": "ensemble", "candidatos": ["ecmwf_aifs025", "ecmwf_aifs025_ensemble"]},
    "GEPS":        {"endpoint": "ensemble", "candidatos": ["gem_global", "gem_ensemble"]},
    "ICON Global": {"endpoint": "forecast", "candidatos": ["icon_global", "icon_seamless"]},
}
MULTIMODELO = "MULTIMODELO"

# ----------------------------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------------------------
session = requests.Session()
session.mount("https://", HTTPAdapter(max_retries=Retry(
    total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504], allowed_methods=["GET"],
)))


def log(msg):
    print(msg, flush=True)


def aviso_gha(msg):
    """Anotación visible en la interfaz de GitHub Actions (y log normal fuera de Actions)."""
    print(f"::warning::{msg}" if os.environ.get("GITHUB_ACTIONS") else f"AVISO: {msg}", flush=True)


def limpiar(texto):
    """Evita que la API key aparezca en logs (los mensajes de error de requests incluyen la URL)."""
    texto = str(texto)
    return texto.replace(API_KEY, "***") if API_KEY else texto


def url_de(endpoint):
    return URL_FORECAST if endpoint == "forecast" else URL_ENSEMBLE


def con_clave(params):
    if API_KEY:
        params["apikey"] = API_KEY
    return params


# ----------------------------------------------------------------------------------------------
# Paso 5 — Resolver códigos de modelo y si admiten HR diaria
# ----------------------------------------------------------------------------------------------
def _consulta_minima(url, codigo, daily):
    params = con_clave({
        "latitude": -16.5, "longitude": -68.15, "daily": daily,
        "timezone": CONFIG["timezone"], "forecast_days": 1, "models": codigo,
    })
    try:
        r = session.get(url, params=params, timeout=30)
    except requests.RequestException as e:
        return False, limpiar(e), None
    if r.status_code == 200:
        return True, "OK", r.json()
    try:
        return False, r.json().get("reason", r.text[:200]), None
    except ValueError:
        return False, r.text[:200], None


def resolver_modelos():
    resueltos = {}
    for nombre, spec in MODELOS.items():
        url = url_de(spec["endpoint"])
        encontrado, errores = None, []
        for candidato in spec["candidatos"]:
            ok, msg, _ = _consulta_minima(url, candidato, ",".join(VARIABLES_DIARIAS))
            if ok:
                encontrado = candidato
                break
            errores.append(f"{candidato}: {msg}")
            time.sleep(0.3)
        if not encontrado:
            aviso_gha(f"Modelo {nombre}: ningún código candidato funcionó ({'; '.join(errores)})")
            continue
        ok, _, datos = _consulta_minima(url, encontrado, f"{VARIABLE_HR}_mean")
        hr_diaria = bool(ok and f"{VARIABLE_HR}_mean" in (datos or {}).get("daily", {}))
        resueltos[nombre] = {"endpoint": spec["endpoint"], "codigo": encontrado, "hr_diaria": hr_diaria}
        log(f"[OK] {nombre:<12} -> {encontrado:<22} endpoint={spec['endpoint']:<8} "
            f"HR={'diaria' if hr_diaria else 'horaria→diaria'}")
        time.sleep(0.3)
    return resueltos


# ----------------------------------------------------------------------------------------------
# Paso 6 — Descarga resiliente por lotes
# ----------------------------------------------------------------------------------------------
def _descargar_lote(lats, lons, url, codigo, fecha_ini, fecha_fin, necesita_hr_horaria):
    daily = VARIABLES_DIARIAS + ([] if necesita_hr_horaria else [f"{VARIABLE_HR}_mean"])
    params = con_clave({
        "latitude": ",".join(f"{v:.4f}" for v in lats),
        "longitude": ",".join(f"{v:.4f}" for v in lons),
        "daily": ",".join(daily),
        "timezone": CONFIG["timezone"],
        "start_date": fecha_ini.isoformat(), "end_date": fecha_fin.isoformat(),
        "models": codigo,
    })
    if necesita_hr_horaria:
        params["hourly"] = VARIABLE_HR

    ultimo_error = None
    intentos = CONFIG["intentos_por_lote"]
    for intento in range(intentos):
        try:
            r = session.get(url, params=params, timeout=180)
            if r.status_code == 429:
                raise requests.HTTPError(f"429 Too Many Requests: {r.text[:200]}", response=r)
            if r.status_code == 400:  # error de parámetros: reintentar no sirve
                raise ValueError(f"400 Bad Request: {r.text[:300]}")
            r.raise_for_status()
            datos = r.json()
            return [datos] if isinstance(datos, dict) else datos
        except requests.RequestException as e:
            ultimo_error = limpiar(e)
            if intento < intentos - 1:
                espera = 30 * (intento + 1)
                log(f"    reintento {intento + 1}/{intentos - 1} en {espera}s: {ultimo_error[:160]}")
                time.sleep(espera)
    raise RuntimeError(f"Lote fallido tras {intentos} intentos: {ultimo_error}")


def descargar(lats, lons, url, codigo, fecha_ini, fecha_fin, necesita_hr_horaria):
    resultados, n, tam = [], len(lats), CONFIG["tam_lote"]
    for i in range(0, n, tam):
        resultados.extend(_descargar_lote(lats[i:i + tam], lons[i:i + tam], url, codigo,
                                          fecha_ini, fecha_fin, necesita_hr_horaria))
        if i + tam < n:
            time.sleep(CONFIG["pausa_entre_lotes_seg"])
    if len(resultados) != n:
        raise RuntimeError(f"Se esperaban {n} respuestas y llegaron {len(resultados)}")
    return resultados


# ----------------------------------------------------------------------------------------------
# Paso 8 — De la respuesta JSON a un valor "oficial" por municipio/día
#          (determinista, o media de todos los miembros del ensamble incluido el control)
# ----------------------------------------------------------------------------------------------
def _patron(nombre):
    return re.compile(rf"^{re.escape(nombre)}(?:_member\d+)?$")


PATRONES_DIARIOS = {v: _patron(v) for v in VARIABLES_DIARIAS}
PATRON_HR_DIARIA = _patron(f"{VARIABLE_HR}_mean")
PATRON_HR_HORARIA = _patron(VARIABLE_HR)


def _media_miembros(bloque, patron, n_dias):
    series = [np.asarray(v, dtype=float) for k, v in bloque.items() if patron.match(k)]
    series = [s for s in series if s.shape == (n_dias,)]
    if not series:
        return np.full(n_dias, np.nan)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=RuntimeWarning)
        return np.nanmean(np.vstack(series), axis=0)


def _hr_horaria_a_diaria(bloque_horario, fechas):
    if not bloque_horario or "time" not in bloque_horario:
        return np.full(len(fechas), np.nan)
    dias = pd.to_datetime(bloque_horario["time"]).strftime("%Y-%m-%d")
    miembros = {k: v for k, v in bloque_horario.items() if PATRON_HR_HORARIA.match(k)}
    if not miembros:
        return np.full(len(fechas), np.nan)
    df = pd.DataFrame({k: pd.to_numeric(pd.Series(v), errors="coerce") for k, v in miembros.items()})
    df["dia"] = dias
    diario = df.groupby("dia").mean().mean(axis=1)  # media diaria por miembro → media entre miembros
    return diario.reindex(fechas).to_numpy(dtype=float)


def respuesta_a_matriz(resp, fechas, necesita_hr_horaria):
    """Devuelve dict variable -> array(len(fechas)) para un municipio."""
    diario = resp.get("daily", {}) or {}
    tiempos = diario.get("time", [])
    n = len(tiempos)
    salida = {}
    for var in VARIABLES_DIARIAS:
        salida[var] = _media_miembros(diario, PATRONES_DIARIOS[var], n)
    if necesita_hr_horaria:
        hr = _hr_horaria_a_diaria(resp.get("hourly"), tiempos)
    else:
        hr = _media_miembros(diario, PATRON_HR_DIARIA, n)
    salida[VARIABLE_HR] = hr

    # alinear a las fechas esperadas (por si la API devuelve un rango distinto)
    s_idx = {t: i for i, t in enumerate(tiempos)}
    alineado = {}
    for var, arr in salida.items():
        alineado[var] = np.array([arr[s_idx[f]] if f in s_idx else np.nan for f in fechas], dtype=float)
    return alineado


# ----------------------------------------------------------------------------------------------
# Programa principal
# ----------------------------------------------------------------------------------------------
def cargar_municipios(limite=None):
    with open(RUTA_CULTIVOS, encoding="utf-8") as f:
        cultivos = json.load(f)
    productores = sorted({g for c in cultivos.values() for g in c["municipios"]})
    with open(RUTA_MUNICIPIOS, encoding="utf-8") as f:
        geo = json.load(f)
    coords = {ft["properties"]["GID_3"]: (ft["properties"]["lat"], ft["properties"]["lon"])
              for ft in geo["features"]}
    faltan = [g for g in productores if g not in coords]
    if faltan:
        raise SystemExit(f"GID_3 productores sin coordenadas en {RUTA_MUNICIPIOS.name}: {faltan}")
    if limite:
        productores = productores[:limite]
    return productores, [coords[g][0] for g in productores], [coords[g][1] for g in productores]


def redondear_matriz(m):
    d = CONFIG["decimales"]
    return [[None if np.isnan(v) else round(float(v), d) for v in fila] for fila in m]


def escribir_resumen_gha(lineas):
    ruta = os.environ.get("GITHUB_STEP_SUMMARY")
    if ruta:
        with open(ruta, "a", encoding="utf-8") as f:
            f.write("\n".join(lineas) + "\n")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--limite-municipios", type=int, default=None, help="solo los N primeros (prueba)")
    ap.add_argument("--salida", type=Path, default=RUTA_SALIDA)
    args = ap.parse_args()

    t0 = time.time()
    fecha_ini = datetime.now(ZoneInfo(CONFIG["timezone"])).date()   # fecha local de Bolivia
    fecha_fin = fecha_ini + timedelta(days=CONFIG["dias_pronostico"] - 1)
    fechas = [(fecha_ini + timedelta(days=d)).isoformat() for d in range(CONFIG["dias_pronostico"])]

    gids, lats, lons = cargar_municipios(args.limite_municipios)
    log(f"Pronóstico {fecha_ini} → {fecha_fin} | {len(gids)} municipios | "
        f"API {'comercial' if API_KEY else 'gratuita'}")

    modelos = resolver_modelos()
    if not modelos:
        aviso_gha("Ningún modelo disponible; no se actualiza el pronóstico.")
        return 1

    datos = {v: {} for v in VARIABLES_SALIDA}     # var -> modelo -> array(n_mun, n_dias)
    fallidos, tiempos_modelo = {}, {}
    for nombre, info in modelos.items():
        t_m = time.time()
        necesita_hr_horaria = not info["hr_diaria"]
        log(f"Descargando {nombre} ({info['codigo']})...")
        try:
            respuestas = descargar(lats, lons, url_de(info["endpoint"]), info["codigo"],
                                   fecha_ini, fecha_fin, necesita_hr_horaria)
            filas = [respuesta_a_matriz(r, fechas, necesita_hr_horaria) for r in respuestas]
            for var in VARIABLES_SALIDA:
                datos[var][nombre] = np.vstack([f[var] for f in filas])
        except Exception as e:  # un modelo caído no detiene la corrida
            fallidos[nombre] = limpiar(e)[:300]
            aviso_gha(f"Modelo {nombre} omitido: {fallidos[nombre]}")
        tiempos_modelo[nombre] = round(time.time() - t_m)
        time.sleep(CONFIG["pausa_entre_modelos_seg"])

    exitosos = [m for m in modelos if m not in fallidos]
    if not exitosos:
        aviso_gha("Todos los modelos fallaron; se conserva el pronóstico anterior.")
        return 1

    # Multimodelo: promedio simple de los modelos disponibles en cada celda + cuántos contribuyen
    n_modelos_multi = {}
    for var in VARIABLES_SALIDA:
        pila = np.stack([datos[var][m] for m in exitosos])
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=RuntimeWarning)
            datos[var][MULTIMODELO] = np.nanmean(pila, axis=0)
        n_modelos_multi[var] = np.sum(~np.isnan(pila), axis=0).max(axis=0).astype(int).tolist()  # por día

    # Días con dato por modelo (los modelos tienen horizontes distintos, p. ej. ICON ~7 días)
    cobertura = {m: int(np.sum(~np.all(np.isnan(datos["temperature_2m_max"][m]), axis=0)))
                 for m in exitosos}

    orden_modelos = exitosos + [MULTIMODELO]
    salida = {
        "version": 1,
        "generado_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "fecha_inicio": fechas[0],
        "fechas": fechas,
        "fuente": "Open-Meteo (open-meteo.com), CC BY 4.0",
        "variables": {v: METADATA_VARIABLES[v] for v in VARIABLES_SALIDA},
        "modelos": orden_modelos,
        "codigos_modelo": {m: modelos[m]["codigo"] for m in exitosos},
        "modelos_fallidos": fallidos,
        "dias_con_dato": cobertura,
        "n_modelos_multimodelo": n_modelos_multi,
        "municipios": gids,
        "datos": {v: {m: redondear_matriz(datos[v][m]) for m in orden_modelos} for v in VARIABLES_SALIDA},
    }

    args.salida.parent.mkdir(parents=True, exist_ok=True)
    tmp = args.salida.with_suffix(".json.tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(salida, f, ensure_ascii=False, separators=(",", ":"))
    tmp.replace(args.salida)   # escritura atómica: nunca queda un JSON a medias

    dur = round((time.time() - t0) / 60, 1)
    kb = args.salida.stat().st_size / 1024
    log(f"Listo en {dur} min → {args.salida} ({kb:.0f} KB). "
        f"Modelos OK: {len(exitosos)}/{len(MODELOS)}")

    resumen = [f"### Pronóstico {fechas[0]} → {fechas[-1]}", "",
               f"{len(gids)} municipios · {dur} min · {kb:.0f} KB", "",
               "| Modelo | Código | Días con dato | Tiempo (s) | Estado |", "|---|---|---|---|---|"]
    for m in MODELOS:
        if m in exitosos:
            resumen.append(f"| {m} | `{modelos[m]['codigo']}` | {cobertura[m]} | {tiempos_modelo[m]} | ✅ |")
        elif m in fallidos:
            resumen.append(f"| {m} | `{modelos[m]['codigo']}` | – | {tiempos_modelo[m]} | ❌ {fallidos[m][:80]} |")
        else:
            resumen.append(f"| {m} | – | – | – | ❌ código no resuelto |")
    escribir_resumen_gha(resumen)
    return 0


if __name__ == "__main__":
    sys.exit(main())
