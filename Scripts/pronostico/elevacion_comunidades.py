#!/usr/bin/env python3
"""
Altitud del punto de consulta de cada comunidad (se corre una sola vez, después de
preparar_comunidades.py). GitHub Actions lo corre solo si el archivo de salida no existe.

La altitud se toma del modelo digital de elevación Copernicus DEM GLO-90 (90 m), el mismo que
usa Open-Meteo para ajustar la temperatura a la altitud del punto consultado. Se leen solo los
bloques necesarios de cada tesela (Cloud Optimized GeoTIFF) desde el repositorio público de AWS,
sin descargar las teselas completas.

Con estas altitudes, el pronóstico de cada comunidad corrige la temperatura de su celda de
pronóstico por la diferencia de altitud (gradiente de 6,5 °C por km; ver pronostico_municipios.py).

Entrada : data/geo/comunidades/comunidades_bolivia.geojson
Salida  : data/geo/comunidades/elevacion_comunidades.json   {"fuente": ..., "z": {id: metros}}

Uso     : python Scripts/pronostico/elevacion_comunidades.py
Requiere: rasterio
Variable de entorno opcional DEM_BASE: carpeta o URL alternativa con las teselas (mismo nombre).
"""
import json
import math
import os
import sys
import time
from collections import defaultdict
from pathlib import Path

import numpy as np

RAIZ = Path(__file__).resolve().parents[2]
RUTA_COMUNIDADES = RAIZ / "data" / "geo" / "comunidades" / "comunidades_bolivia.geojson"
RUTA_SALIDA = RAIZ / "data" / "geo" / "comunidades" / "elevacion_comunidades.json"
DEM_BASE = os.environ.get("DEM_BASE", "https://copernicus-dem-90m.s3.amazonaws.com").rstrip("/")


def nombre_tesela(lat0, lon0):
    """Teselas de 1° nombradas por su esquina suroeste, p. ej. S17_00_W067_00 cubre -17..-16, -67..-66."""
    ns = f"{'N' if lat0 >= 0 else 'S'}{abs(lat0):02d}_00"
    ew = f"{'E' if lon0 >= 0 else 'W'}{abs(lon0):03d}_00"
    nombre = f"Copernicus_DSM_COG_30_{ns}_{ew}_DEM"
    return f"{DEM_BASE}/{nombre}/{nombre}.tif"


def muestrear(ruta, puntos):
    """Altitud (m) en cada (lon, lat); NaN si el punto cae fuera de la tesela o sin dato."""
    import rasterio
    ruta_gdal = ruta if not ruta.startswith("http") else f"/vsicurl/{ruta}"
    with rasterio.open(ruta_gdal) as ds:
        b = ds.bounds
        dentro = [(b.left <= x <= b.right and b.bottom <= y <= b.top) for x, y in puntos]
        valores = np.full(len(puntos), np.nan)
        coords = [p for p, d in zip(puntos, dentro) if d]
        if coords:
            muestras = np.array([v[0] for v in ds.sample(coords)], dtype=float)
            if ds.nodata is not None:
                muestras[muestras == ds.nodata] = np.nan
            valores[np.array(dentro)] = muestras
    return valores


def main():
    os.environ.setdefault("GDAL_DISABLE_READDIR_ON_OPEN", "EMPTY_DIR")   # sin listar el bucket
    os.environ.setdefault("CPL_VSIL_CURL_ALLOWED_EXTENSIONS", ".tif")
    with open(RUTA_COMUNIDADES, encoding="utf-8") as f:
        feats = json.load(f)["features"]
    ids = [ft["properties"]["id"] for ft in feats]
    xy = [(ft["properties"]["x"], ft["properties"]["y"]) for ft in feats]

    grupos = defaultdict(list)
    for i, (x, y) in enumerate(xy):
        grupos[(math.floor(y), math.floor(x))].append(i)

    z = np.full(len(ids), np.nan)
    t0 = time.time()
    for k, ((lat0, lon0), idx) in enumerate(sorted(grupos.items()), 1):
        ruta = nombre_tesela(lat0, lon0)
        for intento in range(3):
            try:
                z[idx] = muestrear(ruta, [xy[i] for i in idx])
                break
            except Exception as e:  # tesela inexistente o corte de red
                if intento == 2:
                    print(f"  AVISO tesela {ruta.rsplit('/', 1)[-1]}: {e}", file=sys.stderr)
                time.sleep(3 * (intento + 1))
        print(f"  [{k}/{len(grupos)}] {ruta.rsplit('/', 1)[-1]}: {len(idx)} comunidades", flush=True)

    ok = ~np.isnan(z)
    salida = {
        "fuente": "Copernicus DEM GLO-90 (© DLR e.V. 2010-2014 y © Airbus Defence and Space GmbH 2014-2018, "
                  "proporcionado bajo COPERNICUS por la Unión Europea y ESA)",
        "unidad": "m",
        "z": {i: int(round(v)) for i, v in zip(ids, z) if not math.isnan(v)},
    }
    RUTA_SALIDA.parent.mkdir(parents=True, exist_ok=True)
    tmp = RUTA_SALIDA.with_suffix(".json.tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(salida, f, ensure_ascii=False, separators=(",", ":"))
    tmp.replace(RUTA_SALIDA)
    print(f"Altitud de {ok.sum()} de {len(ids)} comunidades en {time.time() - t0:.0f} s → {RUTA_SALIDA.name}"
          f" (rango {np.nanmin(z):.0f}–{np.nanmax(z):.0f} m)")
    return 0 if ok.mean() > 0.95 else 1


if __name__ == "__main__":
    sys.exit(main())
