#!/usr/bin/env python3
"""
Corre los modelos agroclimáticos (los mismos scripts de correr_modelos.py) a nivel de COMUNIDAD.

La web calcula estas alertas directamente en el navegador (la misma lógica, validada contra este
script); este programa sirve para obtener las tablas en CSV, revisar resultados o replicar el
cálculo fuera de la plataforma. No se corre en la automatización diaria.

Cómo se arma el pronóstico de cada comunidad:
  1. Cada comunidad toma el pronóstico del punto de su celda (pronostico_comunidades_latest.json,
     generado con: python Scripts/pronostico/pronostico_municipios.py --nivel comunidades).
  2. La temperatura máxima y mínima se ajustan por la diferencia de altitud entre el punto de la celda
     y la comunidad, con un gradiente de 6,5 °C por km:  T_comunidad = T_celda + 0,0065 x (z_celda - z_comunidad).
     La precipitación y la humedad relativa se toman tal cual de la celda.
  3. Las comunidades de cada cultivo son las de sus municipios productores.

Entrada : data/pronostico/pronostico_comunidades_latest.json
          data/geo/comunidades/comunidades_bolivia.geojson
          data/geo/comunidades/elevacion_comunidades.json   (opcional; sin él no se corrige la temperatura)
          data/geo/cultivos_municipios.json
Salida  : Salidas_Modelos_Agroclimaticos/Pronostico_Comunidades/alerta_<indicador>_<fecha>.csv

Uso     : python Scripts/Modelos_Agroclimaticos_Python/correr_modelos_comunidades.py
          python Scripts/Modelos_Agroclimaticos_Python/correr_modelos_comunidades.py --indicador PAPA
"""
import argparse
import importlib.util
import json
import math
import sys
from pathlib import Path

AQUI = Path(__file__).resolve().parent
RAIZ = AQUI.parents[1]
RUTA_PRONOSTICO = RAIZ / "data" / "pronostico" / "pronostico_comunidades_latest.json"
RUTA_COMUNIDADES = RAIZ / "data" / "geo" / "comunidades" / "comunidades_bolivia.geojson"
RUTA_ELEVACION = RAIZ / "data" / "geo" / "comunidades" / "elevacion_comunidades.json"
RUTA_CULTIVOS = RAIZ / "data" / "geo" / "cultivos_municipios.json"
CARPETA_SALIDA = RAIZ / "Salidas_Modelos_Agroclimaticos" / "Pronostico_Comunidades"
VARIABLES_TEMPERATURA = ("temperature_2m_max", "temperature_2m_min")


def cargar_modulo(ruta):
    spec = importlib.util.spec_from_file_location(ruta.stem, ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def clave_celda(lat, lon, res):
    return f"{math.floor(lat / res)}_{math.floor(lon / res)}" if res > 0 else None


def pronostico_por_comunidad(pc, comunidades, elevacion):
    """Convierte el pronóstico por celdas en un pronóstico por comunidad con el mismo formato que
    pronostico_latest.json (una fila por comunidad en lugar de por municipio)."""
    res = pc["resolucion_grados"]
    gradiente = pc.get("gradiente_temp_c_por_km", 6.5) / 1000.0
    fila_punto = {k: i for i, k in enumerate(pc["municipios"])}
    z_punto = pc.get("puntos_elevacion_m") or [None] * len(pc["municipios"])

    ids, filas, ajustes = [], [], []
    for p in comunidades:
        clave = clave_celda(p["y"], p["x"], res) or p["id"]
        if clave not in fila_punto:
            continue
        i = fila_punto[clave]
        zc = elevacion.get(p["id"])
        ajuste = gradiente * (z_punto[i] - zc) if (zc is not None and z_punto[i] is not None) else 0.0
        ids.append(p["id"])
        filas.append(i)
        ajustes.append(ajuste)

    datos = {}
    for var, por_modelo in pc["datos"].items():
        datos[var] = {}
        for modelo, matriz in por_modelo.items():
            if var in VARIABLES_TEMPERATURA:
                datos[var][modelo] = [[None if v is None else v + a for v in matriz[i]] for i, a in zip(filas, ajustes)]
            else:
                datos[var][modelo] = [matriz[i] for i in filas]
    return {**pc, "municipios": ids, "datos": datos}, dict(zip(ids, ajustes))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--pronostico", type=Path, default=RUTA_PRONOSTICO)
    ap.add_argument("--indicador", default=None, help="solo este ID_MODELO (p. ej. PAPA)")
    ap.add_argument("--salida", type=Path, default=CARPETA_SALIDA)
    args = ap.parse_args()

    with open(args.pronostico, encoding="utf-8") as f:
        pc = json.load(f)
    with open(RUTA_COMUNIDADES, encoding="utf-8") as f:
        comunidades = [ft["properties"] for ft in json.load(f)["features"]]
    elevacion = {}
    if RUTA_ELEVACION.exists():
        with open(RUTA_ELEVACION, encoding="utf-8") as f:
            elevacion = json.load(f)["z"]
    else:
        print("AVISO: sin elevacion_comunidades.json; la temperatura no se ajusta por altitud.", file=sys.stderr)
    with open(RUTA_CULTIVOS, encoding="utf-8") as f:
        cultivos_mun = json.load(f)

    pronostico, ajustes = pronostico_por_comunidad(pc, comunidades, elevacion)
    nombre = {p["id"]: p["n"] for p in comunidades}
    cultivos = {}
    for clave, c in cultivos_mun.items():
        cultivos[clave] = {"etiqueta": c["etiqueta"], "municipios": {
            p["id"]: {"nombres": [nombre[p["id"]]], "sup_ha": 0.0, "prod_tn": 0.0}
            for p in comunidades if p["g"] in c["municipios"] and p["id"] in ajustes}}

    args.salida.mkdir(parents=True, exist_ok=True)
    scripts = sorted(AQUI.rglob("Script_Indice_Riesgo_*_Pronostico.py"))
    for ruta in scripts:
        mod = cargar_modulo(ruta)
        if args.indicador and mod.ID_MODELO != args.indicador:
            continue
        res = mod.calcular_alertas(pronostico, cultivos)
        tabla = mod.a_tabla(res, cultivos)
        tabla.insert(1, "ajuste_temp_c", tabla["GID_3"].map(ajustes).round(2))
        tabla = tabla.rename(columns={"GID_3": "id_comunidad", "municipio": "comunidad"})
        destino = args.salida / f"alerta_{mod.ID_MODELO.lower()}_{pc['fecha_inicio']}.csv"
        tabla.round(3).to_csv(destino, index=False, encoding="utf-8-sig")
        multi = tabla[tabla["modelo"] == "MULTIMODELO"]["alerta"].value_counts().to_dict()
        print(f"[OK] {mod.ID_MODELO:<17} {len(res['municipios']):>6} comunidades | multimodelo {multi}")
    print(f"Tablas en {args.salida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
