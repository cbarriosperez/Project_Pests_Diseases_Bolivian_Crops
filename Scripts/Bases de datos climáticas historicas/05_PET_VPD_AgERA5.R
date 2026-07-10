################################################################################
## Nombre del Script: Generar mapas de Evapotranspiración y VPD para Bolivia  ##
## Objetivo:          Recortar y remuestrear (resample) los rasters diarios   ##
##                    de PET y VPD de AgERA5_V2 para Bolivia (1981-2025)      ##
## Autor:             Camilo Barrios-Perez (Ph.D)                             ##
## Correo:            c.barrios@cgiar.org                                     ##
## Institución:       Alianza Bioversity International - CIAT (CGIAR)         ##
################################################################################

# ###### 1. Cargar librerías requeridas ########################################
library(terra)  # Para el manejo de datos raster y análisis espacial
library(sp)     # Para el manejo de objetos espaciales heredados
library(sf)     # Para leer, manipular y transformar datos vectoriales

# ###### 2. Función para procesar variables AgERA5 #############################
procesar_variable_agera5 <- function(var_name, input_dir, output_dir, file_pattern, out_prefix, country_map, reference_grid) {
  message("Iniciando procesamiento de variable: ", var_name)
  
  years = list.dirs(input_dir, full.names = FALSE, recursive = FALSE)
  
  for (i in seq_along(years)) {
    year_i = years[i]
    
    rastlist = list.files(
      path       = file.path(input_dir, year_i),
      pattern    = paste0(file_pattern, "_", year_i),
      full.names = TRUE,
      recursive  = FALSE
    )
    
    if (length(rastlist) == 0) {
      message("Sin archivos en el año: ", year_i)
      next
    }
    
    allrasters = terra::rast(rastlist)
    
    var_crop = crop(allrasters, country_map)
    var_mask = mask(var_crop, country_map)
    
    var_5km = resample(var_mask, reference_grid, method = "bilinear")
    
    stopifnot(all(res(var_5km) == res(reference_grid)))
    
    for (j in 1:nlyr(var_5km)) {
      raster.name = time(var_5km)[j]
      terra::writeRaster(
        var_5km[[j]],
        filename  = file.path(output_dir, paste0(out_prefix, "_", raster.name, ".tif")),
        overwrite = TRUE
      )
    }
    message("Procesado año ", year_i)
  }
}

# ###### 3. Preparación de Entornos y Referencias ##############################
country_map = st_read("\\\\Catalogue/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/shapefile_Bolivia/Nivel_3_Municipio/gadm41_BOL_3.shp")
reference.5km = terra::rast("\\\\CATALOGUE.CGIARAD.ORG/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/Climate/Daily/CHIRPS_v3/chirps-v3.0.1981.01.01.tif")

# ###### 4. Ejecutar procesamiento para PET (ETo) ##############################
procesar_variable_agera5(
  var_name     = "PET (ETo)",
  input_dir    = "\\\\ALLIANCEDFS.ALLIANCE.CGIAR.ORG/data_cluster17V2/Observed_Climate data/AgERA5_V2/ETo",
  output_dir   = "\\\\Catalogue/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/Climate/Daily/ETo",
  file_pattern = "ReferenceET-PenmanMonteith-FAO56_C3S-glob-agric_AgERA5",
  out_prefix   = "ETo",
  country_map  = country_map,
  reference_grid = reference.5km
)

# ###### 5. Ejecutar procesamiento para VPD ####################################
# (Rutas y nombres ilustrativos basados en la estructura estándar de AgERA5)
procesar_variable_agera5(
  var_name     = "VPD",
  input_dir    = "\\\\ALLIANCEDFS.ALLIANCE.CGIAR.ORG/data_cluster17V2/Observed_Climate data/AgERA5_V2/Vapour_Pressure_Deficit",
  output_dir   = "\\\\Catalogue/AgriLACRes_WP2/1.Data/Bolivia/1.Data/Raw/Climate/Daily/VPD",
  file_pattern = "Vapour-Pressure-Deficit_C3S-glob-agric_AgERA5",
  out_prefix   = "VPD",
  country_map  = country_map,
  reference_grid = reference.5km
)