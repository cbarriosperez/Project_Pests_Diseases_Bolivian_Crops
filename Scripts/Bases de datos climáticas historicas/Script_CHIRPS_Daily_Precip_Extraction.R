################################################################################
## Nombre del Script: Generar mapas de precipitación diaria para Bolivia      ##
## Objetivo:          Extraer, recortar, enmascarar y limpiar los datos       ##
##                    globales de precipitación diaria CHIRPS_v3 (1981-2025)  ##
##                    para la extensión territorial de Bolivia. Los           ##
##                    resultados se guardan como archivos GeoTIFF diarios.    ##
## Autor:             Camilo Barrios-Perez (Ph.D)                             ##
## Correo:            c.barrios@cgiar.org                                     ##
## Institución:       Alianza Bioversity International - CIAT (CGIAR)         ##
################################################################################

######################
## Cargar librerías ##
######################

library(terra) # Para el manejo de datos raster y análisis espacial (reemplazo moderno de 'raster')
library(sp)    # Para el manejo de objetos de datos espaciales heredados
library(sf)    # Para leer, manipular y transformar datos vectoriales (Simple Features)

# ###### 1. Cargar y preparar límites vectoriales ##############################

##-- Cargar shapefile de Bolivia --##

# Cargar el shapefile del país (Nivel 3 - Municipios)
country_map = st_read("\\\\Catalogue/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/shapefile_Bolivia/Nivel_3_Municipio/gadm41_BOL_3.shp")

########################################################
## Visualizar el shapefile de la región de estudio    ##
########################################################

# Graficar Bolivia como capa base para verificar la extensión espacial
plot(st_geometry(country_map), main = "Área de Estudio: Bolivia", col = "lightgrey", border = "white", reset = FALSE)


###########################################################################################
### Esta sección del script creará un bucle para extraer la precipitación de CHIRPS     ###
###########################################################################################

##-- Establecer el directorio de trabajo para los datos globales crudos de CHIRPS --##

setwd("\\\\ALLIANCEDFS.ALLIANCE.CGIAR.ORG/data_cluster17V2/Observed_Climate data/CHIRPS_V3/CHIRPS_V3_ERA5_Daily")

#######################################################################################################
## Este bucle importa los mapas globales diarios de precipitación de CHIRPS, los recorta y enmascara ##
## usando el shapefile de Bolivia, aplica limpieza de datos y guarda cada mapa diario en su destino. ##
#######################################################################################################

##--- Listar subdirectorios que representan los años disponibles en el conjunto de datos CHIRPS ---##

years = list.dirs(getwd(), full.names = FALSE, recursive = FALSE)

##--- Inicio del bucle anual ---##

for(i in 1:(length(years)-1)){
  #i=1
  
  ##-- Identificar y listar todos los archivos .tif para el año de procesamiento actual (comenzando en 1981) --##
  # La ruta se construye asumiendo que las carpetas siguen una secuencia cronológica desde 1981
  rastlist = list.files(path = paste(getwd(),"/", 1980 + i, sep =""), pattern="tif$", full.names = TRUE, recursive = FALSE)
  
  ##-- Cargar todos los archivos raster diarios del año en un objeto SpatRaster multicapa --##
  allrasters = terra::rast(rastlist)
  
  ##-- Procesamiento espacial: Recortar (crop) y enmascarar (mask) los mapas globales a la extensión de Bolivia --##
  precipitation = crop(allrasters, country_map)
  precipitation = mask(precipitation, country_map)
  
  ##-- Limpieza de datos: Reemplazar valores negativos (usados a menudo para NoData/Errores) con NA --##
  precipitation = ifel(precipitation < 0, NA, precipitation)
  
  ##--- Inicio del bucle diario: Exportar cada capa individualmente ---##
  
  for(j in 1:nlyr(precipitation)){
    
    # Extraer el nombre original de la capa para el nombre del archivo de salida
    raster.name = names(precipitation)[j]
    
    # Escribir el raster diario procesado en la carpeta de destino local
    writeRaster(precipitation[[j]], 
                paste("\\\\Catalogue/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/Climate/Daily/CHIRPS_v3/", raster.name,".tif", sep =""),
                overwrite = TRUE)
    
  } # Fin del bucle diario
  
} # Fin del bucle anual