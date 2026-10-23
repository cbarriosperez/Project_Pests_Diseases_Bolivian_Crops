"""
Fase 1 — Preparación de geometría (se corre una sola vez, o cuando cambien los shapefiles).

Entrada : data/geo/shapefiles/Shapefiles_municipios_bolivia/gadm41_BOL_3.shp
          data/geo/shapefiles/Shapefiles_cultivos/<CULTIVO>/<CULTIVO>.shp
Salida  : data/geo/municipios_bolivia.geojson   344 municipios GADM simplificados (topología preservada)
                                                 + punto representativo (lat/lon) calculado sobre la
                                                 geometría ORIGINAL, que es el que consulta el pronóstico.
          data/geo/departamentos_bolivia.geojson límites departamentales (para el mapa)
          data/geo/cultivos_municipios.json      índice cultivo -> municipios productores (GID_3)

Uso     : python scripts/pronostico/preparar_geometria.py
Requiere: geopandas >= 1.0, shapely >= 2.1 (coverage_simplify)
"""
import json
from pathlib import Path

import geopandas as gpd
import shapely

RAIZ = Path(__file__).resolve().parents[2]
DIR_SHP = RAIZ / "data" / "geo" / "shapefiles"
DIR_SALIDA = RAIZ / "data" / "geo"

TOLERANCIA_GRADOS = 0.006   # ~650 m; suficiente para escala nacional/departamental
DECIMALES = 4               # ~11 m de precisión en coordenadas

CULTIVOS = {  # carpeta -> etiqueta para la web
    "AJO": "Ajo", "BANANO": "Banano", "CAFE": "Café", "CEBOLLA": "Cebolla",
    "DURAZNO": "Durazno", "FRUTILLA": "Frutilla", "LECHUGA": "Lechuga",
    "MAIZ_BERZA": "Maíz Berza", "MAIZ_CHOCLO": "Maíz Choclo", "MAIZ_EN_GRANO": "Maíz en Grano",
    "NARANJA": "Naranja", "PAPA": "Papa", "QUINUA": "Quinua", "TOMATE": "Tomate", "UVA": "Uva",
}


# Correcciones de asignación municipio (censo) -> polígono GADM detectadas en los shapefiles de cultivos.
# Clave: (MUNICIPIO, DEPARTAMEN) tal como vienen en el shapefile; valor: GID_3 correcto.
# San Pedro de Buena Vista es la capital de la provincia Chárcas (norte de Potosí) y venía asignado a
# San Pablo de Lípez (BOL.7.15.3_2, Sur Lípez), ~420 km al sur.
CORRECCIONES_GID3 = {
    ("San Pedro De Buena Vista", "Potosi"): "BOL.7.3.1_2",
}


def redondear(geom):
    return shapely.set_precision(geom, 10 ** -DECIMALES)


def main():
    gadm = gpd.read_file(DIR_SHP / "Shapefiles_municipios_bolivia" / "gadm41_BOL_3.shp").to_crs(4326)
    assert gadm["GID_3"].is_unique

    # Punto de consulta del pronóstico: siempre dentro del polígono, sobre la geometría original
    rp = gadm.geometry.representative_point()
    gadm["lat"] = rp.y.round(4)
    gadm["lon"] = rp.x.round(4)

    # Simplificación de cobertura: preserva bordes compartidos (sin huecos ni solapes entre vecinos)
    simpl = shapely.coverage_simplify(gadm.geometry.values, tolerance=TOLERANCIA_GRADOS)
    gadm["geometry"] = [redondear(g) for g in simpl]

    municipios = gadm.rename(columns={"NAME_1": "departamento", "NAME_2": "provincia", "NAME_3": "municipio"})
    municipios = municipios[["GID_3", "departamento", "provincia", "municipio", "lat", "lon", "geometry"]]
    municipios.to_file(DIR_SALIDA / "municipios_bolivia.geojson", driver="GeoJSON",
                       layer_options={"COORDINATE_PRECISION": DECIMALES, "RFC7946": "YES"})

    deptos = municipios.dissolve("departamento", as_index=False)[["departamento", "geometry"]]
    deptos["geometry"] = [redondear(g) for g in deptos.geometry]
    deptos.to_file(DIR_SALIDA / "departamentos_bolivia.geojson", driver="GeoJSON",
                   layer_options={"COORDINATE_PRECISION": DECIMALES, "RFC7946": "YES"})

    # Índice cultivo -> municipios. Algunos municipios del censo agropecuario comparten un mismo
    # polígono GADM (p. ej. Jesús de Machaca -> Santiago de Machaca); se agregan por GID_3.
    indice, avisos = {}, []
    for carpeta, etiqueta in CULTIVOS.items():
        df = gpd.read_file(DIR_SHP / "Shapefiles_cultivos" / carpeta / f"{carpeta}.shp").drop(columns="geometry")
        for (mun, dep), gid_ok in CORRECCIONES_GID3.items():
            mascara = (df["MUNICIPIO"] == mun) & (df["DEPARTAMEN"] == dep) & (df["GID_3"] != gid_ok)
            if mascara.any():
                avisos.append(f"{carpeta}: corregido {mun} ({dep}) {df.loc[mascara, 'GID_3'].iloc[0]} -> {gid_ok}")
                df.loc[mascara, "GID_3"] = gid_ok
        faltan = set(df["GID_3"]) - set(municipios["GID_3"])
        assert not faltan, f"{carpeta}: GID_3 que no existen en GADM: {faltan}"
        municipios_cultivo = {}
        for gid, grupo in df.groupby("GID_3", sort=True):
            if len(grupo) > 1:
                avisos.append(f"{carpeta}: {gid} agrupa {list(grupo['MUNICIPIO'])}")
            municipios_cultivo[gid] = {
                "nombres": sorted(grupo["MUNICIPIO"].unique().tolist()),
                "sup_ha": round(float(grupo["SUP_HA"].sum()), 1),
                "prod_tn": round(float(grupo["PROD_TN"].sum()), 1),
            }
        indice[carpeta] = {"etiqueta": etiqueta, "municipios": municipios_cultivo}

    with open(DIR_SALIDA / "cultivos_municipios.json", "w", encoding="utf-8") as f:
        json.dump(indice, f, ensure_ascii=False, separators=(",", ":"))

    unicos = sorted({g for c in indice.values() for g in c["municipios"]})
    print(f"Municipios GADM: {len(municipios)} | productores únicos (a consultar): {len(unicos)}")
    for a in avisos:
        print("  Aviso:", a)
    for p in ("municipios_bolivia.geojson", "departamentos_bolivia.geojson", "cultivos_municipios.json"):
        print(f"  {p}: {(DIR_SALIDA / p).stat().st_size / 1024:.0f} KB")


if __name__ == "__main__":
    main()
