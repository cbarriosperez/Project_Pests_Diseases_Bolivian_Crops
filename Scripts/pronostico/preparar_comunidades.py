#!/usr/bin/env python3
"""
Preparación de las comunidades para el nivel «Comunidad» de la plataforma
(se corre una sola vez, o cuando cambien el shapefile de comunidades o los de cultivos).

Corta (filtra) las comunidades del Censo a las zonas productoras actuales, es decir, a los
municipios productores de cada cultivo (data/geo/cultivos_municipios.json). Cada comunidad se
asigna a su municipio GADM por la ubicación de su punto representativo (un punto que siempre cae
dentro del polígono; en comunidades de forma regular coincide prácticamente con el centroide).

Entrada : data/geo/shapefiles/Shapefiles_comunidades/Comunidades_2025_Pol.shp
          data/geo/shapefiles/Shapefiles_municipios_bolivia/gadm41_BOL_3.shp
          data/geo/cultivos_municipios.json
Salida  : data/geo/comunidades/comunidades_bolivia.geojson
            Comunidades de los municipios productores, simplificadas para la web. Propiedades con
            nombres cortos para aligerar el archivo:
              id  código INE de la comunidad      n  nombre de la comunidad
              g   GID_3 del municipio GADM         y, x  latitud y longitud del punto de consulta
              p   población (Censo)
            El municipio y el departamento se obtienen del GID_3 (municipios_bolivia.geojson).

Uso     : python Scripts/pronostico/preparar_comunidades.py
Requiere: geopandas >= 1.0, shapely >= 2.1 (coverage_simplify)

Después de correrlo, generar la altitud de cada comunidad (una sola vez):
          python Scripts/pronostico/elevacion_comunidades.py
"""
import json
from pathlib import Path

import geopandas as gpd
import shapely

RAIZ = Path(__file__).resolve().parents[2]
DIR_GEO = RAIZ / "data" / "geo"
SHP_COMUNIDADES = DIR_GEO / "shapefiles" / "Shapefiles_comunidades" / "Comunidades_2025_Pol.shp"
SHP_GADM = DIR_GEO / "shapefiles" / "Shapefiles_municipios_bolivia" / "gadm41_BOL_3.shp"
RUTA_CULTIVOS = DIR_GEO / "cultivos_municipios.json"
DIR_SALIDA = DIR_GEO / "comunidades"

TOLERANCIA_GRADOS = 0.004   # ~450 m: suficiente para ver comunidades a escala municipal
DECIMALES_GEOM = 3          # ~110 m en la geometría de la web (el GeoJSON pesa ~7 MB, ~0,9 MB comprimido)
DECIMALES_PUNTO = 5         # ~1 m en el punto de consulta

MINUSCULAS = {"de", "del", "la", "las", "los", "el", "y", "e", "en"}


def nombre_propio(texto):
    """'SANTA ROSA DE LIMA' -> 'Santa Rosa de Lima' (los nombres del censo vienen en mayúsculas)."""
    palabras = str(texto or "").strip().lower().split()
    return " ".join(p if (i > 0 and p in MINUSCULAS) else p[:1].upper() + p[1:] for i, p in enumerate(palabras))


def main():
    com = gpd.read_file(SHP_COMUNIDADES).to_crs(4326)
    gadm = gpd.read_file(SHP_GADM).to_crs(4326)[["GID_3", "NAME_1", "NAME_3", "geometry"]]
    with open(RUTA_CULTIVOS, encoding="utf-8") as f:
        cultivos = json.load(f)
    productores = {g for c in cultivos.values() for g in c["municipios"]}

    com = com.rename(columns={"Cod INE 20": "id", "nombre_2": "comunidad", "Poblacion": "poblacion"})
    assert com["id"].is_unique, "El código INE de las comunidades debe ser único"
    com["comunidad"] = com["comunidad"].map(nombre_propio)

    # Punto de consulta: siempre dentro del polígono, calculado sobre la geometría original
    puntos = com[["id"]].copy()
    puntos = gpd.GeoDataFrame(puntos, geometry=com.geometry.representative_point(), crs=4326)
    unidos = gpd.sjoin(puntos, gadm, how="left", predicate="within").drop_duplicates("id")
    sin_mun = unidos["GID_3"].isna()
    if sin_mun.any():   # puntos en el borde exacto: municipio más cercano
        faltan = puntos[puntos["id"].isin(unidos.loc[sin_mun, "id"])].to_crs(32720)
        cercano = gpd.sjoin_nearest(faltan, gadm.to_crs(32720), how="left")
        print(f"  {len(faltan)} comunidades asignadas al municipio más cercano (punto en el borde)")
        unidos = unidos.set_index("id")
        unidos.update(cercano.drop_duplicates("id").set_index("id")[["GID_3", "NAME_1", "NAME_3"]])
        unidos = unidos.reset_index()
    unidos = unidos.set_index("id")

    com["GID_3"] = com["id"].map(unidos["GID_3"])
    com["municipio"] = com["id"].map(unidos["NAME_3"])
    com["departamento"] = com["id"].map(unidos["NAME_1"])
    com["lat"] = puntos.geometry.y.round(DECIMALES_PUNTO).values
    com["lon"] = puntos.geometry.x.round(DECIMALES_PUNTO).values

    # Corte a las zonas productoras actuales (municipios productores de al menos un cultivo)
    total = len(com)
    com = com[com["GID_3"].isin(productores)].copy()

    # Simplificación de cobertura (bordes compartidos sin huecos ni solapes) y precisión para la web
    simpl = shapely.coverage_simplify(com.geometry.values, tolerance=TOLERANCIA_GRADOS)
    com["geometry"] = shapely.set_precision(simpl, 10 ** -DECIMALES_GEOM)
    com = com[~com.geometry.is_empty]

    salida = com[["id", "comunidad", "GID_3", "lat", "lon", "poblacion", "geometry"]]
    salida = salida.sort_values(["GID_3", "id"]).rename(
        columns={"comunidad": "n", "GID_3": "g", "lat": "y", "lon": "x", "poblacion": "p"})
    DIR_SALIDA.mkdir(parents=True, exist_ok=True)
    ruta = DIR_SALIDA / "comunidades_bolivia.geojson"
    salida.to_file(ruta, driver="GeoJSON",
                   layer_options={"COORDINATE_PRECISION": DECIMALES_GEOM, "RFC7946": "YES"})

    print(f"Comunidades: {total} | en municipios productores: {len(salida)} "
          f"| municipios con comunidades: {salida['g'].nunique()} de {len(productores)}")
    for clave, c in sorted(cultivos.items()):
        n = int(salida["g"].isin(c["municipios"]).sum())
        print(f"  {clave:<14} {n:>6} comunidades")
    print(f"  {ruta.relative_to(RAIZ)}: {ruta.stat().st_size / 1e6:.1f} MB")


if __name__ == "__main__":
    main()
