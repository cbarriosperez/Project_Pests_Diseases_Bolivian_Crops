################################################################################
## Nombre del Script: Generar mapas de temperatura mínima diaria para Bolivia ##
## Objetivo:          Recortar y remuestrear (resample) los rasters diarios   ##
##                    de temperatura mínima de AgERA5_V2 (1981-2025) para la  ##
##                    extensión territorial de Bolivia, convirtiendo las      ##
##                    unidades de Kelvin a grados Celsius y guardando los     ##
##                    resultados a 5 km de resolución, alineados con la       ##
##                    cuadrícula de CHIRPS.                                   ##
## Autor:             Camilo Barrios-Perez (Ph.D)                             ##
## Correo:            c.barrios@cgiar.org                                     ##
## Institución:       Alianza Bioversity International - CIAT (CGIAR)         ##
################################################################################

# ###### 1. Cargar librerías requeridas ########################################
library(terra)  # Para el manejo de datos raster y análisis espacial
library(sp)     # Para el manejo de objetos espaciales heredados
library(sf)     # Para leer, manipular y transformar datos vectoriales (Simple Features)

# ###### 2. Cargar y preparar límites vectoriales ##############################

## Extraer el polígono de Bolivia para usarlo como mapa base
country_map = st_read("\\\\Catalogue/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/shapefile_Bolivia/Nivel_3_Municipio/gadm41_BOL_3.shp")

# ###### 3. Visualizar las capas espaciales para verificar la alineación #######

## Graficar Bolivia como capa base
plot(
  st_geometry(country_map),
  main   = "Área de Estudio: Bolivia",
  col    = "lightgrey",
  border = "white",
  reset  = FALSE
)

# ###### 4. Cargar y preparar el raster de referencia a 5 km (Grid CHIRPS) #####

## Raster de CHIRPS V3 utilizado como plantilla espacial para el remuestreo:
## Todos los rasters de Tmin de salida coincidirán con la extensión, resolución y CRS de esta cuadrícula
reference.5km = terra::rast("\\\\CATALOGUE.CGIARAD.ORG/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/Climate/Daily/CHIRPS_v3/chirps-v3.0.1981.01.01.tif")

plot(reference.5km, main = "Cuadrícula de referencia CHIRPS - Bolivia")

# ###### 5. Establecer directorio de trabajo de Temperatura Mínima AgERA5 ######
setwd("\\\\ALLIANCEDFS.ALLIANCE.CGIAR.ORG/data_cluster17V2/Observed_Climate data/AgERA5_V2/Minimum Temperature")

# ###### 6. Procesar rasters diarios de Tmin año por año (1981-2025) ###########
#
#  Por cada año, el bucle:
#    (a) Lee todos los GeoTIFFs diarios de Tmin para ese año
#    (b) Recorta y enmascara el stack de rasters a la extensión de Bolivia
#    (c) Convierte las unidades de temperatura de Kelvin a grados Celsius
#    (d) Remuestrea a la cuadrícula de referencia CHIRPS de 5 km (interpolación bilineal)
#    (e) Escribe un GeoTIFF por día en la carpeta de salida
#
################################################################################

## Listar todos los subdirectorios de años disponibles en el directorio de AgERA5
years = list.dirs(getwd(), full.names = FALSE, recursive = FALSE)

## Iterar sobre cada año (excluyendo la última entrada, que suele estar incompleta)
for (i in 1:(length(years) - 1)) {
  # i = 1  # descomentar para depurar una sola iteración
  
  # ###### 6a. Importar todos los rasters diarios de Tmin para el año actual ###
  
  ## Construir la lista de archivos que coincidan con la nomenclatura de AgERA5 para este año
  rastlist = list.files(
    path       = file.path(getwd(), 1980 + i),
    pattern    = paste0("Temperature-Air-2m_Min-24h_C3S-glob-agric_AgERA5_", 1980 + i),
    full.names = TRUE,
    recursive  = FALSE
  )
  
  ## Apilar todos los rasters diarios en un único SpatRaster multicapa
  allrasters = terra::rast(rastlist)
  
  # ###### 6b. Recortar (crop) y enmascarar (mask) el stack a Bolivia ##########
  Tmin = crop(allrasters, country_map)
  Tmin = mask(Tmin, country_map)
  
  # ###### 6c. Convertir Tmin de Kelvin a Celsius (K - 273.15 = °C) ############
  T.min.C = Tmin - 273.15
  
  # ###### 6d. Remuestrear a la cuadrícula de referencia CHIRPS de 5 km ########
  ## Se utiliza interpolación bilineal para datos continuos de temperatura
  Tmin.5km = resample(T.min.C, reference.5km, method = "bilinear")
  
  ## Verificar que la resolución de salida coincida con la cuadrícula de referencia
  stopifnot(all(res(Tmin.5km) == res(reference.5km)))
  
  # ###### 6e. Escribir un GeoTIFF por día en la carpeta de salida #############
  for (j in 1:nlyr(Tmin.5km)) {
    # j = 1  # descomentar para depurar una sola capa
    
    ## Recuperar la fecha incrustada en los metadatos de la capa raster
    raster.name = time(Tmin.5km)[j]
    
    ## Guardar la capa diaria de Tmin usando la fecha como nombre de archivo
    terra::writeRaster(Tmin.5km[[j]], filename  = paste0("\\\\Catalogue/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/Climate/Daily/Tmin/Tmin_", raster.name, ".tif"), overwrite = TRUE)
  }
}