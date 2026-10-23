#!/usr/bin/env python3
"""
Corre todos los modelos agroclimáticos (un script por cultivo/plaga) sobre el pronóstico diario
y escribe un JSON compacto para la página web.

Entrada : data/pronostico/pronostico_latest.json   (Scripts/pronostico/pronostico_municipios.py)
          data/geo/cultivos_municipios.json         (municipios productores por cultivo)
Salida  : data/alertas/alertas_latest.json

Uso     : python Scripts/Modelos_Agroclimaticos_Python/correr_modelos.py

Cada modelo vive en su propio archivo (*/Script_Indice_Riesgo_*_Pronostico.py) y también se puede
correr solo. Este script los descubre automáticamente: para agregar un modelo basta con agregar
su archivo en una carpeta nueva.
"""
import argparse
import importlib.util
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import numpy as np

AQUI = Path(__file__).resolve().parent
RAIZ = AQUI.parents[1]
RUTA_PRONOSTICO = RAIZ / "data" / "pronostico" / "pronostico_latest.json"
RUTA_CULTIVOS = RAIZ / "data" / "geo" / "cultivos_municipios.json"
RUTA_SALIDA = RAIZ / "data" / "alertas" / "alertas_latest.json"

CATEGORIAS = [
    {"id": 0, "nombre": "Normal", "color": "#228b22"},
    {"id": 1, "nombre": "Alerta baja", "color": "#ffd700"},
    {"id": 2, "nombre": "Alerta media", "color": "#ffa500"},
    {"id": 3, "nombre": "Alerta alta", "color": "#cd0000"},
]   # -1 = sin dato (modelo sin horizonte suficiente o dato faltante)


def aviso(msg):
    print(f"::warning::{msg}" if os.environ.get("GITHUB_ACTIONS") else f"AVISO: {msg}", flush=True)


def cargar_modulo(ruta):
    spec = importlib.util.spec_from_file_location(ruta.stem, ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def lista(arr, decimales):
    return [None if np.isnan(v) else round(float(v), decimales) for v in arr]


def a_json(mod, res):
    modelos = res["modelos"]
    return {
        "cultivo": res["cultivo"], "clave_cultivo": res["clave_cultivo"], "plaga": res["plaga"],
        "periodo_dias_modelo": res["periodo_dias_modelo"], "ventana": res["ventana"],
        "umbral": res["umbral"], "condiciones": res["condiciones"],
        "script": str(Path(mod.__file__).relative_to(RAIZ)).replace("\\", "/"),
        "municipios": res["municipios"],
        "indice": {m: lista(r["indice"], 3) for m, r in modelos.items()},
        "alerta": {m: r["alerta"].tolist() for m, r in modelos.items()},
        "variables": {v: {m: lista(r["variables"][v], 1) for m, r in modelos.items()}
                      for v in mod.ENTRADAS},
        "condiciones_cumplidas": {c: {m: [None if np.isnan(x) else int(x) for x in r["condiciones"][c]]
                                      for m, r in modelos.items()} for c in mod.CONDICIONES},
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--pronostico", type=Path, default=RUTA_PRONOSTICO)
    ap.add_argument("--cultivos", type=Path, default=RUTA_CULTIVOS)
    ap.add_argument("--salida", type=Path, default=RUTA_SALIDA)
    args = ap.parse_args()

    with open(args.pronostico, encoding="utf-8") as f:
        pronostico = json.load(f)
    with open(args.cultivos, encoding="utf-8") as f:
        cultivos = json.load(f)

    scripts = sorted(AQUI.rglob("Script_Indice_Riesgo_*_Pronostico.py"))
    indicadores, errores = {}, {}
    for ruta in scripts:
        try:
            mod = cargar_modulo(ruta)
            res = mod.calcular_alertas(pronostico, cultivos)
            indicadores[mod.ID_MODELO] = a_json(mod, res)
            multi = res["modelos"].get("MULTIMODELO")
            conteo = np.bincount(multi["alerta"] + 1, minlength=5).tolist() if multi else "-"
            print(f"[OK] {mod.ID_MODELO:<17} ventana {res['ventana']['dias']:>2} d | "
                  f"{len(res['municipios']):>3} mun. | multimodelo [sin dato, N, B, M, A] = {conteo}")
        except Exception as e:   # un modelo con error no detiene a los demás
            errores[ruta.parent.name] = f"{type(e).__name__}: {e}"[:300]
            aviso(f"Modelo {ruta.name} omitido: {errores[ruta.parent.name]}")

    if not indicadores:
        aviso("Ningún modelo pudo calcularse; se conserva el archivo anterior.")
        return 1

    salida = {
        "version": 1,
        "generado_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "pronostico_generado_utc": pronostico.get("generado_utc"),
        "fecha_inicio": pronostico["fecha_inicio"],
        "modelos": pronostico["modelos"],
        "categorias": CATEGORIAS,
        "indicadores": indicadores,
        "errores": errores,
    }
    args.salida.parent.mkdir(parents=True, exist_ok=True)
    tmp = args.salida.with_suffix(".json.tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(salida, f, ensure_ascii=False, separators=(",", ":"))
    tmp.replace(args.salida)   # escritura atómica
    print(f"Listo: {len(indicadores)}/{len(scripts)} modelos → {args.salida} "
          f"({args.salida.stat().st_size / 1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
