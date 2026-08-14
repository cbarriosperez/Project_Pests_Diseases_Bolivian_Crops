## ============================================================================
## INDICE AGROCLIMATICO DE RIESGO - TOMATE / POLILLA (TUTA ABSOLUTA)
## Sistema de Alerta Temprana de Plagas y Enfermedades - Bolivia
## ============================================================================
##
## QUE HACE ESTE SCRIPT
## ---------------------
## Calcula, a nivel espacial (pixel a pixel, no puntual), el mismo indice de
## riesgo que la hoja "Calculadora_Indice de riesgo" del archivo Excel calcula
## para un solo punto (fila 18: Tomate / Polilla (Tuta absoluta)).
##
## Umbral de alerta (columna P18 del Excel):
##   ALERTA si: T media > 22 C  y  >= 7 dias secos consecutivos  y  HR < 60%.
##
## Indicador propuesto (columna R18):
##   Indice seco-calido = coincidencia de T media > 22 C, >= 7 dias secos consecutivos y HR < 60%.
##
## La formula de esa hoja (celda AC18) suma, con sus pesos:
##     (T media > 22 C) * 2
##     (>= 7 dias secos consecutivos) * 1
##     (HR media < 60%) * 1
##   ... y se divide entre la suma de los pesos (4).
##
## y el semaforo de alerta (misma logica que la celda S18) es:
##   indice = 0    -> "Normal"
##   indice < 0.5  -> "Alerta baja"
##   indice < 1    -> "Alerta media"
##   indice = 1    -> "Alerta alta"
## (los pixeles fuera de la region de tomate, o sin dato climatico,
## quedan NA y se pintan transparentes en el mapa; no se les asigna categoria)
##
## Este script replica la misma arquitectura del script de Ajo (mismas
## funciones auxiliares, mismo flujo de carga/recorte/periodizacion/mapeo);
## solo cambian los parametros, la formula del indice y el area de estudio.
##
## ============================================================================
## POR QUE PERIODOS DE 10 DIAS
## ----------------------------
## La columna K18 exige "≥ 7 dias secos consecutivos"; se usa un periodo de 10
## dias, con la misma logica que Aranuela roja en Durazno, para dar holgura a esa
## racha.
## Por eso PERIODO_DIAS = 10 es el valor recomendado por defecto.
## El script deja este valor como parametro para que el usuario lo pruebe con
## otras duraciones si lo prefiere (ver PARAMETROS DEL USUARIO).
## ============================================================================


# ============================================================================
# 1. PARAMETROS DEL USUARIO -- editar solo esta seccion para cada corrida
# ============================================================================

# --- 1.1 Periodizacion --------------------------------------------------
periodo_dias   <- 10                       # largo de cada periodo (dias). Recomendado: 10
fecha_inicio   <- as.Date("2025-01-01")   # primer dia desde el cual se arman los periodos
fecha_fin      <- as.Date("2025-12-31")   # ultimo dia a considerar (prueba: año 2025)

# Buffer opcional de dias ANTES de fecha_inicio, solo para calcular bien la
# racha de dias secos del primer periodo (si no se dispone de datos previos,
# dejar en 0; la racha del primer periodo se calculara desde fecha_inicio).
# Necesario porque la formula usa dias secos consecutivos (columna W).
dias_buffer    <- 10

# --- 1.2 Carpetas con los GeoTIFF diarios --------------------------------
# Se asume: un archivo .tif por dia y por variable, en una carpeta por
# variable (pueden tener subcarpetas por año, el script busca recursivamente).
# OJO: cada variable puede tener su propio formato de fecha en el nombre de
# archivo. En este proyecto:
#   Tmax_2025-01-01.tif           -> fecha con guiones (AAAA-MM-DD)
#   Tmin_2025-01-01.tif           -> fecha con guiones (AAAA-MM-DD)
#   RHum_2025-01-01.tif           -> fecha con guiones (AAAA-MM-DD)
#   chirps-v3.0.2025.01.01.tif    -> fecha con PUNTOS  (AAAA.MM.DD)  (CHIRPS)
# Por eso cargar_serie_diaria() recibe DOS argumentos de fecha:
#   patron_fecha  = expresion regular para EXTRAER la fecha del nombre
#   formato_fecha = formato para CONVERTIR ese texto en fecha (as.Date)
# Si tus archivos usan otro patron, solo hay que ajustar estos dos argumentos
# en las llamadas de la seccion 4. Estas carpetas son las MISMAS para todos
# los cultivos (son series climaticas de Bolivia); lo que cambia por cultivo
# es el shapefile de la region productora (seccion 1.3).

carpeta_tmax   <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Datos_clima_diarios_2025/Tmax"     # temperatura maxima diaria (°C)
carpeta_tmin   <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Datos_clima_diarios_2025/Tmin"     # temperatura minima diaria (°C)
carpeta_precip <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Datos_clima_diarios_2025/Precipitacion"   # precipitacion diaria acumulada (mm)
carpeta_hr     <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Datos_clima_diarios_2025/HR"       # humedad relativa media diaria (%)

# --- 1.3 Shapefiles: region productora de tomate y mapa base de municipios ---
ruta_shp_cultivo     <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Shapefiles/TOMATE/TOMATE.shp"

# Limites municipales de Bolivia, solo para dar contexto geografico al mapa
# (se dibujan como fondo; el indice de riesgo se sigue calculando y pintando
# unicamente dentro de la region de tomate).
ruta_shp_municipios  <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Shapefiles/Municipios/gadm41_BOL_3.shp"

# margen (en grados) que se deja alrededor de la region al graficar, para que
# el mapa muestre algo de contexto y no quede recortado justo al borde
buffer_mapa_grados  <- 0.4

# --- 1.4 Carpeta de salida -------------------------------------------------
carpeta_salida <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/salidas_tomate_2025"
dir.create(carpeta_salida, showWarnings = FALSE, recursive = TRUE)

# ============================================================================
# 2. LIBRERIAS
# ============================================================================

library(terra)      # manejo de rasters (mas rapido y moderno que 'raster')
library(sf)          # manejo de shapefiles
library(dplyr)       # manipulacion de tablas
library(lubridate)   # manejo de fechas
library(ggplot2)     # mapas
library(tidyterra)   # geom_spatraster() para graficar SpatRaster con ggplot2


# ============================================================================
# 3. FUNCIONES AUXILIARES
# ============================================================================

## ----------------------------------------------------------------------
## 3.1 Cargar una serie diaria de GeoTIFF como un SpatRaster con fecha
##     asignada a cada capa, ya filtrado al rango de fechas que interesa.
##     (identica a la usada en el script de Ajo: cada archivo se relaciona
##     con su fecha real extraida del nombre, no con su posicion/orden).
## ----------------------------------------------------------------------
cargar_serie_diaria <- function(carpeta, fecha_desde, fecha_hasta,
                                 patron_fecha  = "\\d{4}-\\d{2}-\\d{2}",
                                 formato_fecha = "%Y-%m-%d") {

  archivos <- list.files(carpeta, pattern = "\\.tif$", full.names = TRUE,
                          recursive = TRUE)
  if (length(archivos) == 0) {
    stop("No se encontraron archivos .tif en: ", carpeta)
  }

  fechas_txt <- regmatches(basename(archivos),
                            regexpr(patron_fecha, basename(archivos)))
  fechas <- as.Date(fechas_txt, format = formato_fecha)

  validos <- !is.na(fechas)
  if (any(!validos)) {
    warning(sum(!validos), " archivo(s) de ", carpeta,
            " no se pudieron leer como fecha con el patron/formato indicado y se ignoraron.")
  }
  archivos <- archivos[validos]
  fechas   <- fechas[validos]

  en_rango <- which(fechas >= fecha_desde & fechas <= fecha_hasta)
  if (length(en_rango) == 0) {
    stop("Ningun archivo de ", carpeta, " cae dentro del rango de fechas pedido.")
  }
  archivos <- archivos[en_rango]
  fechas   <- fechas[en_rango]

  orden    <- order(fechas)
  archivos <- archivos[orden]
  fechas   <- fechas[orden]

  if (anyDuplicated(fechas)) {
    stop("Hay fechas duplicadas en ", carpeta, ": revisar nombres de archivo.")
  }
  dias_esperados <- seq(min(fechas), max(fechas), by = "day")
  faltantes <- as.character(dias_esperados[!dias_esperados %in% fechas])
  if (length(faltantes) > 0) {
    warning(length(faltantes), " dia(s) faltantes en la serie de ", carpeta, ": ",
            paste(head(faltantes, 5), collapse = ", "),
            if (length(faltantes) > 5) ", ..." else "")
  }

  r <- rast(archivos)
  terra::time(r) <- fechas
  names(r) <- as.character(fechas)
  r
}

## ----------------------------------------------------------------------
## 3.2 Racha de dias secos consecutivos, pixel a pixel, dia a dia.
##     Un dia seco = precipitacion <= 1 mm. La racha se reinicia en 0 apenas
##     hay un dia de lluvia (> 1 mm). (identica a la del script de Ajo)
## ----------------------------------------------------------------------
calcular_racha_seca <- function(dia_seco_r) {

  racha_por_pixel <- function(x) {
    if (all(is.na(x))) return(rep(NA_real_, length(x)))
    x[is.na(x)] <- 0
    grupo <- cumsum(x == 0)
    racha <- ave(x, grupo, FUN = seq_along) * x
    racha
  }

  app(dia_seco_r, racha_por_pixel)
}

## ----------------------------------------------------------------------
## 3.3 Formula EXACTA del indice de riesgo (replica celda AC18 del
##     Excel) para Tomate / Polilla (Tuta absoluta).
## ----------------------------------------------------------------------
calcular_indice_tomate <- function(t_media_periodo, dias_secos_periodo, hr_media_periodo) {

  cond_t <- t_media_periodo > 22   # T media > 22 C
  cond_secos <- dias_secos_periodo >= 7   # >= 7 dias secos consecutivos
  cond_hr <- hr_media_periodo < 60   # HR media < 60%

  # pesos: 2, 1, 1 (tal cual columnas Z18:AB18 del Excel)
  (cond_t * 2 + cond_secos * 1 + cond_hr * 1) / 4
}

## ----------------------------------------------------------------------
## 3.4 Clasificacion de alerta tipo semaforo (misma logica en las 15 fichas
##     del Excel: columna S). Los pixeles NA (fuera de la region, o sin dato
##     climatico) se dejan NA -> se pintan transparentes en el mapa, no se
##     les asigna una categoria "Sin datos".
## ----------------------------------------------------------------------
clasificar_alerta <- function(indice_r) {

  estado <- ifel(indice_r == 0,   1,
            ifel(indice_r < 0.5,  2,
            ifel(indice_r < 1,    3,
                                   4)))

  levels(estado) <- data.frame(
    id     = 1:4,
    alerta = c("Normal", "Alerta baja", "Alerta media", "Alerta alta")
  )
  estado
}

## paleta de colores tipo semaforo (NA = transparente, ver na.value mas abajo)
colores_alerta <- c("Normal"      = "forestgreen",
                    "Alerta baja" = "gold",
                    "Alerta media"= "orange",
                    "Alerta alta" = "red3")


# ============================================================================
# 4. CARGA DE SHAPEFILES Y DE LAS SERIES CLIMATICAS DIARIAS
# ============================================================================

region_cultivo <- st_read(ruta_shp_cultivo, quiet = TRUE)

# mapa base de municipios: solo da contexto geografico, no participa en el
# calculo del indice.
municipios <- st_read(ruta_shp_municipios, quiet = TRUE)
municipios <- st_transform(municipios, st_crs(region_cultivo))

# extension del mapa: la region del cultivo + un margen de contexto alrededor
bbox_region <- st_bbox(region_cultivo)
xlim_mapa   <- c(bbox_region["xmin"] - buffer_mapa_grados, bbox_region["xmax"] + buffer_mapa_grados)
ylim_mapa   <- c(bbox_region["ymin"] - buffer_mapa_grados, bbox_region["ymax"] + buffer_mapa_grados)

# rango real de datos a cargar (incluye el buffer para la racha del 1er periodo)
fecha_carga_desde <- fecha_inicio - dias_buffer

# Tmax, Tmin y HR usan fecha con guiones (AAAA-MM-DD) -> sirven los valores
# por defecto de cargar_serie_diaria(). Precipitacion (CHIRPS) usa fecha con
# puntos (AAAA.MM.DD), asi que se le indica el patron/formato propio.
tmax   <- cargar_serie_diaria(carpeta_tmax,   fecha_carga_desde, fecha_fin)
tmin   <- cargar_serie_diaria(carpeta_tmin,   fecha_carga_desde, fecha_fin)
hr     <- cargar_serie_diaria(carpeta_hr,     fecha_carga_desde, fecha_fin)
precip <- cargar_serie_diaria(carpeta_precip, fecha_carga_desde, fecha_fin,
                               patron_fecha  = "\\d{4}\\.\\d{2}\\.\\d{2}",
                               formato_fecha = "%Y.%m.%d")

# recortar y enmascarar todo a la region productora de tomate
region_vect <- vect(region_cultivo)
tmax   <- mask(crop(tmax,   region_vect), region_vect)
tmin   <- mask(crop(tmin,   region_vect), region_vect)
precip <- mask(crop(precip, region_vect), region_vect)
hr     <- mask(crop(hr,     region_vect), region_vect)

# ============================================================================
# 5. VARIABLES DIARIAS DERIVADAS
# ============================================================================

t_media_diaria <- (tmax + tmin) / 2                 # T media diaria (°C)
dia_seco       <- precip <= 1                       # 1 = seco, 0 = lluvia
racha_seca     <- calcular_racha_seca(dia_seco)      # dias secos consecutivos, dia a dia

fechas_disponibles <- terra::time(precip)


# ============================================================================
# 6. CALCULO DEL INDICE POR PERIODOS
# ============================================================================

inicios_periodo <- seq(fecha_inicio, fecha_fin, by = periodo_dias)

resultados        <- list()
resultados_alerta  <- list()
resumen_tabla      <- data.frame()

for (i in seq_along(inicios_periodo)) {

  ini <- inicios_periodo[i]
  fin <- min(ini + periodo_dias - 1, fecha_fin)

  idx <- which(fechas_disponibles >= ini & fechas_disponibles <= fin)
  if (length(idx) == 0) next

  # variables agregadas del periodo, pixel a pixel
  t_media_periodo    <- mean(t_media_diaria[[idx]])          # promedio del periodo
  hr_media_periodo    <- mean(hr[[idx]])                       # promedio del periodo
  dias_secos_periodo  <- racha_seca[[max(idx)]]                # racha vigente al ULTIMO dia del periodo

  indice <- calcular_indice_tomate(t_media_periodo, dias_secos_periodo, hr_media_periodo)
  names(indice) <- paste0("indice_", format(ini, "%Y%m%d"), "_", format(fin, "%Y%m%d"))

  alerta <- clasificar_alerta(indice)
  names(alerta) <- paste0("alerta_", format(ini, "%Y%m%d"), "_", format(fin, "%Y%m%d"))

  resultados[[i]]        <- indice
  resultados_alerta[[i]] <- alerta

  # tabla resumen: superficie (%) de la region en cada categoria de alerta
  frecuencias  <- freq(alerta)
  total_celdas <- sum(frecuencias$count)
  fila <- data.frame(
    periodo_inicio = ini,
    periodo_fin    = fin,
    categoria      = frecuencias$value,
    pct_area       = round(100 * frecuencias$count / total_celdas, 1)
  )

  # ---- diagnostico: que tan seguido se cumple CADA condicion por separado ----
  cond_t_r <- t_media_periodo > 22
  cond_secos_r <- dias_secos_periodo >= 7
  cond_hr_r <- hr_media_periodo < 60
  fila$pct_area_cond_t <- round(100 * as.numeric(global(cond_t_r, "mean", na.rm = TRUE)), 1)
  fila$pct_area_cond_secos <- round(100 * as.numeric(global(cond_secos_r, "mean", na.rm = TRUE)), 1)
  fila$pct_area_cond_hr <- round(100 * as.numeric(global(cond_hr_r, "mean", na.rm = TRUE)), 1)

  resumen_tabla <- rbind(resumen_tabla, fila)

  # ----- mapa del periodo -----
  mapa <- ggplot() +
    geom_sf(data = municipios, fill = "grey97", color = "grey60", linewidth = 0.2) +
    geom_spatraster(data = alerta) +
    geom_sf(data = region_cultivo, fill = NA, color = "black", linewidth = 0.5) +
    scale_fill_manual(values = colores_alerta, na.translate = FALSE,
                       na.value = "transparent", name = "Estado de alerta") +
    coord_sf(xlim = xlim_mapa, ylim = ylim_mapa, expand = FALSE) +
    labs(title = "Riesgo de Polilla (Tuta absoluta) en Tomate",
         subtitle = paste0("Periodo: ", format(ini, "%d-%b-%Y"), " a ", format(fin, "%d-%b-%Y")),
         caption = "Indice = coincidencia de T media >22C, >=7 dias secos consecutivos y HR <60%") +
    theme_minimal()

  # ---- mostrar el mapa en pantalla ANTES de guardarlo ----
  message(sprintf("Periodo %d/%d: %s a %s",
                   i, length(inicios_periodo), format(ini, "%d-%b-%Y"), format(fin, "%d-%b-%Y")))
  print(mapa)
  Sys.sleep(0.3)

  ggsave(file.path(carpeta_salida,
                    paste0("mapa_alerta_tomate_", format(ini, "%Y%m%d"), "_", format(fin, "%Y%m%d"), ".png")),
         mapa, width = 7, height = 6, dpi = 150)
}


# ============================================================================
# 7. GUARDAR RESULTADOS
# ============================================================================

writeRaster(rast(resultados),        file.path(carpeta_salida, "indice_riesgo_tomate_2025.tif"), overwrite = TRUE)
writeRaster(rast(resultados_alerta), file.path(carpeta_salida, "estado_alerta_tomate_2025.tif"),  overwrite = TRUE)

write.csv(resumen_tabla, file.path(carpeta_salida, "resumen_superficie_alerta_tomate_2025.csv"),
          row.names = FALSE)

message("Listo. Mapas, rasters y tabla resumen guardados en: ", carpeta_salida)
