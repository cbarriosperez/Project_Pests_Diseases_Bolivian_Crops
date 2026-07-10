################################################################################
## Nombre del Script: Generar mapas de Evapotranspiración diaria para Bolivia ##
## Objetivo:          Recortar y remuestrear (resample) los rasters diarios   ##
##                    de Evapotranspiración de Referencia (ETo) de AgERA5_V2  ##
##                    (1981-2025) para la extensión territorial de Bolivia,   ##
##                    guardando los resultados a 5 km de resolución,          ##
##                    alineados con la cuadrícula de CHIRPS.                  ##
##                    Unidades: mm/día                                        ##
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
## Todos los rasters de ETo de salida coincidirán con la extensión, resolución y CRS de esta cuadrícula
reference.5km = terra::rast("\\\\CATALOGUE.CGIARAD.ORG/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/Climate/Daily/CHIRPS_v3/chirps-v3.0.1981.01.01.tif")

plot(reference.5km, main = "Cuadrícula de referencia CHIRPS - Bolivia")

# ###### 5. Establecer directorio de trabajo de Evapotranspiración (ETo) #######
setwd("\\\\ALLIANCEDFS.ALLIANCE.CGIAR.ORG/data_cluster17V2/Observed_Climate data/AgERA5_V2/ETo")

# ###### 6. Procesar rasters diarios de ETo año por año (1981-2025) ############
#
#  Por cada año, el bucle:
#    (a) Lee todos los archivos diarios de ETo para ese año
#    (b) Recorta y enmascara el stack de rasters a la extensión de Bolivia
#    (c) No se requiere conversión de unidades (AgERA5 ETo ya está en mm/día)
#    (d) Remuestrea a la cuadrícula de referencia CHIRPS de 5 km (interpolación bilineal)
#    (e) Escribe un GeoTIFF por día en la carpeta de salida
#
################################################################################

## Listar todos los subdirectorios de años disponibles en el directorio de ETo
years = list.dirs(getwd(), full.names = FALSE, recursive = FALSE)

## Iterar sobre cada año disponible
for (i in seq_along(years)) {
  # i = 1 # descomentar para depurar una sola iteración
  year_i = years[i]
  
  # ###### 6a. Importar todos los rasters diarios de ETo para el año actual ####
  ## Nomenclatura real:
  ## ReferenceET-PenmanMonteith-FAO56_C3S-glob-agric_AgERA5_YYYYMMDD_final-v2.0.0.nc
  rastlist = list.files(
    path       = file.path(getwd(), year_i),
    pattern    = paste0("ReferenceET-PenmanMonteith-FAO56_C3S-glob-agric_AgERA5_", year_i),
    full.names = TRUE,
    recursive  = FALSE
  )
  
  ## Si la carpeta está vacía, saltar al siguiente año
  if (length(rastlist) == 0) {
    message("Sin archivos en el año: ", year_i)
    next
  }
  
  ## Apilar archivos NetCDF - 'terra' selecciona la primera variable automáticamente.
  ## Si hay más de una variable en el .nc, se puede especificar con: terra::rast(f, subds = "et0_pm")
  allrasters = terra::rast(rastlist)
  
  # ###### 6b. Recortar (crop) y enmascarar (mask) el stack a Bolivia ##########
  ETo = crop(allrasters, country_map)
  ETo = mask(ETo, country_map)
  
  # ###### 6c. Unidades ########################################################
  ## AgERA5 V2 ReferenceET viene por defecto en mm/día - No requiere conversión.
  
  # ###### 6d. Remuestrear a la cuadrícula de referencia CHIRPS de 5 km ########
  ## Se utiliza interpolación bilineal
  ETo.5km = resample(ETo, reference.5km, method = "bilinear")
  
  ## Verificar que la resolución de salida coincida con la cuadrícula de referencia
  stopifnot(all(res(ETo.5km) == res(reference.5km)))
  
  # ###### 6e. Escribir un GeoTIFF por día en la carpeta de salida #############
  for (j in 1:nlyr(ETo.5km)) {
    
    ## Recuperar la fecha incrustada en los metadatos de la capa raster
    raster.name = time(ETo.5km)[j]
    
    ## Guardar la capa diaria de ETo usando la fecha como nombre de archivo
    terra::writeRaster(
      ETo.5km[[j]],
      filename  = paste0(
        "\\\\Catalogue/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/Climate/Daily/ETo/ETo_",
        raster.name, ".tif"
      ),
      overwrite = TRUE
    )
  }
}