#!/usr/bin/env python3
## ============================================================================
## INDICE AGROCLIMATICO DE RIESGO - DURAZNO / ARAÑUELA ROJA
## Sistema de Alerta Temprana de Plagas y Enfermedades - Bolivia
## Versión Python para el PRONÓSTICO a 16 días por municipio
## ============================================================================
##
## QUE HACE ESTE SCRIPT
## ---------------------
## Calcula, para cada municipio productor de durazno, el mismo índice de riesgo que
## la hoja "Calculadora_Indice de riesgo" del Excel calcula para un solo punto
## (fila 9: Durazno / Arañuela roja), usando como entrada el pronóstico diario de
## Open-Meteo que genera Scripts/pronostico/pronostico_municipios.py.
##
## Es la traducción a Python de:
##   Scripts/Modelos_Agroclimaticos_R/Durazno/Aranuela_roja/Script_Indice_Riesgo_Durazno_Aranuela_Espacial.R
## Fórmula, pesos y semáforo son idénticos a los del script en R.
##
## Umbral de alerta (columna P9 del Excel):
##   ALERTA si: T media > 25 °C, >= 7 días secos consecutivos y HR < 50 %.
##
## La fórmula (celda AC9) suma, con sus pesos:
##     (T media > 25 °C) * 2
##     (>= 7 días secos consecutivos) * 1
##     (HR media < 50 %) * 1
##   ... y se divide entre la suma de los pesos (4).
##
## Semáforo de alerta (columna S del Excel):
##   indice = 0    -> "Normal"
##   indice < 0.5  -> "Alerta baja"
##   indice < 1    -> "Alerta media"
##   indice = 1    -> "Alerta alta"
##
## ============================================================================
## DIFERENCIAS CON LA VERSION EN R
## --------------------------------
## 1. Unidad espacial: municipio en lugar de píxel. El pronóstico trae un valor
##    diario por municipio (punto lat/lon de consulta), así que la "región
##    productora" ya no es el shapefile del cultivo sino la lista de municipios
##    de data/geo/cultivos_municipios.json (clave "DURAZNO").
## 2. Periodo: en R se recorre un año en bloques de 10 días. Aquí se evalúa
##    UNA ventana fija con los primeros VENTANA_DIAS días del pronóstico (desde
##    hoy). Como el pronóstico se actualiza a diario, la alerta también.
## 3. Varios modelos: el índice se calcula para cada modelo del pronóstico y
##    para el MULTIMODELO. Un modelo cuyo horizonte no cubre la ventana completa
##    (p. ej. ICON, ~7 días) queda sin índice (NaN), igual que en R cuando falta
##    un día en el periodo.
## 4. Racha de días secos: se cuenta desde el primer día del pronóstico (sin
##    días previos), igual que el script en R con dias_buffer = 0.
## ============================================================================
##
## USO
##   python Script_Indice_Riesgo_Durazno_Aranuela_Pronostico.py
##   python Script_Indice_Riesgo_Durazno_Aranuela_Pronostico.py --pronostico ruta/pronostico_latest.json --salida carpeta/
##
## Requiere solo numpy y pandas.
## ============================================================================

import argparse
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd


# ============================================================================
# 1. PARAMETROS DEL USUARIO
# ============================================================================

ID_MODELO   = "DURAZNO_ARANUELA"           # identificador en alertas_latest.json
CLAVE_CULTIVO = "DURAZNO"             # clave en data/geo/cultivos_municipios.json
CULTIVO     = "Durazno"
PLAGA       = "Arañuela roja"

PERIODO_DIAS = 10                # periodo del modelo original en R (días)
VENTANA_DIAS = PERIODO_DIAS             # ventana evaluada (primeros N días desde hoy)
UMBRAL_DIA_SECO_MM = 1                   # día seco: precipitación <= 1 mm (igual que en R)
UMBRAL_DIA_LLUVIA_MM = 1                 # día de lluvia: precipitación > 1 mm (igual que en R)

# Rutas por defecto (relativas a la raíz del repositorio)
RAIZ = Path(__file__).resolve().parents[4]
RUTA_PRONOSTICO = RAIZ / "data" / "pronostico" / "pronostico_latest.json"
RUTA_CULTIVOS   = RAIZ / "data" / "geo" / "cultivos_municipios.json"
CARPETA_SALIDA  = RAIZ / "Salidas_Modelos_Agroclimaticos" / "Pronostico"

# Variables del pronóstico (nombres de Open-Meteo)
VAR_PRECIP, VAR_TMAX, VAR_TMIN, VAR_HR = (
    "precipitation_sum", "temperature_2m_max", "temperature_2m_min", "relative_humidity_2m")

CATEGORIAS = ["Normal", "Alerta baja", "Alerta media", "Alerta alta"]


# ============================================================================
# 2. FORMULA DEL INDICE (replica la celda AC9 del Excel)
# ============================================================================

UMBRAL_TEXTO = "T media > 25 °C, >= 7 días secos consecutivos y HR < 50 %"

# condición: (descripción, peso)
CONDICIONES = {
    "cond_t": ("T media > 25 °C", 2),
    "cond_secos": (">= 7 días secos consecutivos", 1),
    "cond_hr": ("HR media < 50 %", 1),
}


def calcular_indice_durazno_aranuela(t_media_periodo, dias_secos_periodo, hr_media_periodo):
    """Índice de riesgo Durazno / Arañuela roja. Devuelve (índice, condiciones)."""
    cond_t     = t_media_periodo > 25   # T media > 25 °C
    cond_secos = dias_secos_periodo >= 7   # >= 7 días secos consecutivos
    cond_hr    = hr_media_periodo < 50   # HR media < 50 %

    # pesos: 2, 1, 1 (tal cual columnas Z9:AB9 del Excel)
    indice = (cond_t * 2 + cond_secos * 1 + cond_hr * 1) / 4
    return indice, {"cond_t": cond_t, "cond_secos": cond_secos, "cond_hr": cond_hr}


# Variables agregadas que necesita la fórmula (en el mismo orden de los argumentos)
ENTRADAS = ["t_media_periodo", "dias_secos_periodo", "hr_media_periodo"]


# ============================================================================
# 3. FUNCIONES AUXILIARES (idénticas en todos los modelos)
# ============================================================================

def calcular_racha_seca(dia_seco):
    """Días secos consecutivos, día a día (misma lógica que calcular_racha_seca() en R).
    dia_seco: matriz municipios x días con 1 = seco, 0 = lluvia, NaN = sin dato.
    Un dato faltante no cuenta como seco; un municipio sin ningún dato queda NaN."""
    racha = np.zeros_like(dia_seco, dtype=float)
    actual = np.zeros(dia_seco.shape[0])
    seco = np.nan_to_num(dia_seco, nan=0.0)
    for d in range(dia_seco.shape[1]):
        actual = (actual + 1) * seco[:, d]
        racha[:, d] = actual
    racha[np.all(np.isnan(dia_seco), axis=1), :] = np.nan
    return racha


def variables_del_periodo(tmax, tmin, precip, hr):
    """Agrega las series diarias de la ventana (municipios x días) como en la sección 6 del
    script en R. Igual que terra::mean()/sum() sin na.rm, un día faltante deja el valor en NaN.
    Solo se calculan las variables que usa este modelo, para que un faltante en otra variable
    (p. ej. HR) no anule el índice."""
    salida = {}
    if "t_media_periodo" in ENTRADAS:
        salida["t_media_periodo"] = ((tmax + tmin) / 2).mean(axis=1)          # promedio del periodo
    if "precip_acum_periodo" in ENTRADAS:
        salida["precip_acum_periodo"] = precip.sum(axis=1)                    # suma del periodo
    if "hr_media_periodo" in ENTRADAS:
        salida["hr_media_periodo"] = hr.mean(axis=1)                          # promedio del periodo
    if "dias_lluvia_periodo" in ENTRADAS:
        dia_lluvia = np.where(np.isnan(precip), np.nan, (precip > UMBRAL_DIA_LLUVIA_MM).astype(float))
        salida["dias_lluvia_periodo"] = dia_lluvia.sum(axis=1)                # conteo de días de lluvia
    if "dias_secos_periodo" in ENTRADAS:
        dia_seco = np.where(np.isnan(precip), np.nan, (precip <= UMBRAL_DIA_SECO_MM).astype(float))
        salida["dias_secos_periodo"] = calcular_racha_seca(dia_seco)[:, -1]   # racha vigente al ÚLTIMO día
    return salida


def clasificar_alerta(indice):
    """Semáforo (columna S del Excel). Devuelve 0..3 (índice en CATEGORIAS) o -1 si no hay dato."""
    i = np.round(indice, 9)
    cat = np.select([i == 0, i < 0.5, i < 1], [0, 1, 2], default=3)
    return np.where(np.isnan(indice), -1, cat).astype(int)


def _matriz(pronostico, variable, modelo, filas, n_dias):
    datos = pronostico["datos"][variable][modelo]
    m = [[np.nan if v is None else v for v in datos[f][:n_dias]] for f in filas]
    return np.array(m, dtype=float).reshape(len(filas), n_dias)


def calcular_alertas(pronostico, cultivos):
    """Calcula índice, semáforo, variables y condiciones para cada modelo del pronóstico.
    Devuelve un diccionario listo para serializar en JSON."""
    municipios_cultivo = cultivos[CLAVE_CULTIVO]["municipios"]
    pos = {g: i for i, g in enumerate(pronostico["municipios"])}
    gids = [g for g in municipios_cultivo if g in pos]
    sin_pronostico = [g for g in municipios_cultivo if g not in pos]
    filas = [pos[g] for g in gids]

    fechas = pronostico["fechas"]
    n = min(VENTANA_DIAS, len(fechas))
    if n < VENTANA_DIAS:
        print(f"AVISO {ID_MODELO}: el pronóstico trae {len(fechas)} días; ventana reducida a {n}.",
              file=sys.stderr)

    resultado = {
        "cultivo": CULTIVO, "clave_cultivo": CLAVE_CULTIVO, "plaga": PLAGA,
        "periodo_dias_modelo": PERIODO_DIAS,
        "ventana": {"inicio": fechas[0], "fin": fechas[n - 1], "dias": n},
        "umbral": UMBRAL_TEXTO,
        "condiciones": {c: {"descripcion": d, "peso": p} for c, (d, p) in CONDICIONES.items()},
        "municipios": gids, "municipios_sin_pronostico": sin_pronostico,
        "modelos": {},
    }
    for modelo in pronostico["modelos"]:
        tmax = _matriz(pronostico, VAR_TMAX, modelo, filas, n)
        tmin = _matriz(pronostico, VAR_TMIN, modelo, filas, n)
        precip = _matriz(pronostico, VAR_PRECIP, modelo, filas, n)
        hr = _matriz(pronostico, VAR_HR, modelo, filas, n)

        variables = variables_del_periodo(tmax, tmin, precip, hr)
        with np.errstate(invalid="ignore"):
            indice, condiciones = calcular_indice_durazno_aranuela(*[variables[v] for v in ENTRADAS])
        indice = np.asarray(indice, dtype=float)
        # numpy compara NaN como False; en R el resultado sería NA -> se replica aquí
        faltante = np.any([np.isnan(variables[v]) for v in ENTRADAS], axis=0)
        indice[faltante] = np.nan

        resultado["modelos"][modelo] = {
            "indice": indice,
            "alerta": clasificar_alerta(indice),
            "variables": variables,
            "condiciones": {c: np.where(faltante, np.nan, np.asarray(v, dtype=float))
                            for c, v in condiciones.items()},
        }
    return resultado


def a_tabla(resultado, cultivos):
    """Convierte el resultado en una tabla larga (una fila por municipio y modelo)."""
    nombres = cultivos[CLAVE_CULTIVO]["municipios"]
    filas = []
    for modelo, r in resultado["modelos"].items():
        for i, gid in enumerate(resultado["municipios"]):
            fila = {"GID_3": gid, "municipio": ", ".join(nombres[gid]["nombres"]), "modelo": modelo,
                    "ventana_inicio": resultado["ventana"]["inicio"], "ventana_fin": resultado["ventana"]["fin"]}
            fila.update({v: r["variables"][v][i] for v in ENTRADAS})
            fila.update({c: r["condiciones"][c][i] for c in CONDICIONES})
            fila["indice"] = r["indice"][i]
            fila["alerta"] = CATEGORIAS[r["alerta"][i]] if r["alerta"][i] >= 0 else "Sin dato"
            filas.append(fila)
    return pd.DataFrame(filas)


# ============================================================================
# 4. EJECUCION INDIVIDUAL (genera un CSV para este cultivo/plaga)
# ============================================================================

def main():
    ap = argparse.ArgumentParser(description=f"Índice de riesgo {CULTIVO} / {PLAGA} sobre el pronóstico")
    ap.add_argument("--pronostico", type=Path, default=RUTA_PRONOSTICO)
    ap.add_argument("--cultivos", type=Path, default=RUTA_CULTIVOS)
    ap.add_argument("--salida", type=Path, default=CARPETA_SALIDA)
    args = ap.parse_args()

    with open(args.pronostico, encoding="utf-8") as f:
        pronostico = json.load(f)
    with open(args.cultivos, encoding="utf-8") as f:
        cultivos = json.load(f)

    resultado = calcular_alertas(pronostico, cultivos)
    tabla = a_tabla(resultado, cultivos)

    args.salida.mkdir(parents=True, exist_ok=True)
    ruta = args.salida / f"alerta_{ID_MODELO.lower()}_{pronostico['fecha_inicio']}.csv"
    tabla.round(3).to_csv(ruta, index=False, encoding="utf-8-sig")

    v = resultado["ventana"]
    print(f"{CULTIVO} / {PLAGA} | ventana {v['inicio']} a {v['fin']} ({v['dias']} días) | "
          f"{len(resultado['municipios'])} municipios")
    resumen = (tabla.groupby(["modelo", "alerta"]).size().unstack(fill_value=0)
               .reindex(columns=CATEGORIAS + ["Sin dato"], fill_value=0))
    print(resumen.to_string())
    if resultado["municipios_sin_pronostico"]:
        print(f"AVISO: {len(resultado['municipios_sin_pronostico'])} municipio(s) del cultivo sin pronóstico.")
    print(f"Tabla guardada en: {ruta}")


if __name__ == "__main__":
    main()
