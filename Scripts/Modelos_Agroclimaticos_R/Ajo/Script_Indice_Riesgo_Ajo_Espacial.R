## ============================================================================
## INDICE AGROCLIMATICO DE RIESGO - AJO / FUSARIOSIS (pudricion basal)
## Sistema de Alerta Temprana de Plagas y Enfermedades - Bolivia
## ============================================================================
##
## QUE HACE ESTE SCRIPT
## ---------------------
## Calcula, a nivel espacial (pixel a pixel, no puntual), el mismo indice de
## riesgo que la hoja "Calculadora_Indice de riesgo" del archivo Excel
## calcula para un solo punto (fila 5: Ajo / Fusariosis).
##
## La formula de esa hoja (celda AC5) es:
##
##   indice = ( (T_media >= 25)*1 + (Dias_secos_consecutivos >= 10)*1 +
##              (Precip_acumulada < 20)*1 ) / 3
##
## y el semaforo de alerta (celda S5) es:
##   indice = 0    -> "Normal"
##   indice < 0.5  -> "Alerta baja"
##   indice < 1    -> "Alerta media"
##   indice = 1    -> "Alerta alta"
## (los pixeles fuera de la region de ajo, o sin dato climatico, quedan NA
## y se pintan transparentes en el mapa; no se les asigna una categoria)
##
## Este script reproduce EXACTAMENTE esa formula, pero calculando T_media,
## Dias_secos_consecutivos y Precip_acumulada para cada pixel del area
## productora de ajo, a partir de series climaticas diarias.
##
## Nota: las variables Humedad Relativa (U) y Dias de lluvia/nublados (X) se
## calculan tambien (quedan disponibles) porque la calculadora general las
## pide, pero para Ajo/Fusariosis NO entran en la formula del indice (ver
## columna Y5 del Excel: las variables determinantes son solo T, dias secos
## y precipitacion).
##
## ============================================================================
## POR QUE PERIODOS DE 15 DIAS
## ----------------------------
## El umbral definido en el Excel para esta plaga usa dos referencias de
## tiempo distintas:
##   - Columna I5 (precipitacion favorable): "< 20 mm / 15 dias"
##   - Columna P5 (umbral de alerta): ">= 10 dias secos consecutivos"
## Un periodo de 15 dias es el que hace compatibles ambas condiciones: permite
## acumular la precipitacion exactamente en la misma ventana que define el
## indicador bibliografico (15 dias) y dentro de esa ventana es perfectamente
## posible observar una racha de 10 o mas dias secos consecutivos.
## Por eso PERIODO_DIAS = 15 es el valor recomendado por defecto. El script
## deja este valor como parametro para que el usuario lo pruebe con 10 o 20
## dias si lo prefiere (ver PARAMETROS DEL USUARIO).
## ============================================================================


# ============================================================================
# 1. PARAMETROS DEL USUARIO -- editar solo esta seccion para cada corrida
# ============================================================================

# --- 1.1 Periodizacion --------------------------------------------------
periodo_dias   <- 15                      # largo de cada periodo (dias). Recomendado: 15
fecha_inicio   <- as.Date("2025-01-01")   # primer dia desde el cual se arman los periodos
fecha_fin      <- as.Date("2025-12-31")   # ultimo dia a considerar (prueba: año 2025)

# Buffer opcional de dias ANTES de fecha_inicio, solo para calcular bien la
# racha de dias secos del primer periodo (si no se dispone de datos previos,
# dejar en 0; la racha del primer periodo se calculara desde fecha_inicio).
dias_buffer    <- 15

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
# en las llamadas de la seccion 4.

carpeta_tmax   <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Datos_clima_diarios_2025/Tmax"     # temperatura maxima diaria (°C)
carpeta_tmin   <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Datos_clima_diarios_2025/Tmin"     # temperatura minima diaria (°C)
carpeta_precip <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Datos_clima_diarios_2025/Precipitacion"   # precipitacion diaria acumulada (mm)
carpeta_hr     <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Datos_clima_diarios_2025/HR"       # humedad relativa media diaria (%)

# --- 1.3 Shapefiles: region productora de ajo y mapa base de municipios ---
ruta_shp_ajo         <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Shapefiles/AJO/AJO.shp"

# Limites municipales de Bolivia, solo para dar contexto geografico al mapa
# (se dibujan como fondo; el indice de riesgo se sigue calculando y pintando
# unicamente dentro de la region de ajo). << COMPLETAR con la ruta real.
ruta_shp_municipios  <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/Shapefiles/Municipios/gadm41_BOL_3.shp"

# margen (en grados) que se deja alrededor de la region de ajo al graficar,
# para que el mapa muestre algo de contexto y no quede recortado justo al
# borde de la region
buffer_mapa_grados  <- 0.4

# --- 1.4 Carpeta de salida -------------------------------------------------
carpeta_salida <- "D:/OneDrive - CGIAR/Otras Colaboraciones/Proyectos Formulados/2024/Yapu/Proyecto_SAT de Plagas y enfermedades en Bolivia/Scripts_Modelos agroclimaticos por cultivo/salidas_ajo_2025"
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
##
##     Como se hace la relacion archivo <-> fecha (siguiendo la logica que
##     propuso Camilo): a cada archivo se le extrae su fecha real desde el
##     nombre (no se asume que el orden alfabetico ya es el orden cronologico),
##     se descartan fechas fuera del rango pedido, se ordena explicitamente
##     por fecha y RECIEN AHI se arma el stack. Asi, cuando mas adelante se
##     pida "el periodo del 1 al 15 de enero", el script sabe con certeza que
##     capas corresponden a esos 15 dias exactos (y no solo "los primeros 15
##     archivos en el orden en que los devolvio el sistema operativo").
##     Ademas se avisa si faltan dias o si hay fechas repetidas.
## ----------------------------------------------------------------------
cargar_serie_diaria <- function(carpeta, fecha_desde, fecha_hasta,
                                 patron_fecha  = "\\d{4}-\\d{2}-\\d{2}",
                                 formato_fecha = "%Y-%m-%d") {

  archivos <- list.files(carpeta, pattern = "\\.tif$", full.names = TRUE,
                          recursive = TRUE)
  if (length(archivos) == 0) {
    stop("No se encontraron archivos .tif en: ", carpeta)
  }

  # extraer la fecha desde el nombre del archivo, usando el patron y formato
  # propios de esta variable (ver seccion 1.2)
  fechas_txt <- regmatches(basename(archivos),
                            regexpr(patron_fecha, basename(archivos)))
  fechas <- as.Date(fechas_txt, format = formato_fecha)

  # archivos cuyo nombre no se pudo interpretar como fecha -> se descartan y se avisa
  validos <- !is.na(fechas)
  if (any(!validos)) {
    warning(sum(!validos), " archivo(s) de ", carpeta,
            " no se pudieron leer como fecha con el patron/formato indicado y se ignoraron.")
  }
  archivos <- archivos[validos]
  fechas   <- fechas[validos]

  # quedarnos solo con lo que esta en el rango de interes, y ordenar por fecha
  en_rango <- which(fechas >= fecha_desde & fechas <= fecha_hasta)
  if (length(en_rango) == 0) {
    stop("Ningun archivo de ", carpeta, " cae dentro del rango de fechas pedido.")
  }
  archivos <- archivos[en_rango]
  fechas   <- fechas[en_rango]

  orden    <- order(fechas)
  archivos <- archivos[orden]
  fechas   <- fechas[orden]

  # controles de calidad: fechas duplicadas o dias faltantes en la serie
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
##     Un dia seco = precipitacion <= 1 mm (definicion pedida por el usuario)
##     La racha se reinicia en 0 apenas hay un dia de lluvia (> 1 mm).
##     dia_seco_r: SpatRaster binario (1 = seco, 0 = lluvia) en orden temporal
## ----------------------------------------------------------------------
calcular_racha_seca <- function(dia_seco_r) {

  racha_por_pixel <- function(x) {
    if (all(is.na(x))) return(rep(NA_real_, length(x)))
    x[is.na(x)] <- 0            # dato faltante = no se cuenta como seco
    # cada vez que aparece un 0 (dia de lluvia) se "corta" la racha:
    grupo <- cumsum(x == 0)
    # dentro de cada grupo, se numera 1,2,3... y se multiplica por x para que
    # los dias de lluvia (x=0) queden en racha=0
    racha <- ave(x, grupo, FUN = seq_along) * x
    racha
  }

  app(dia_seco_r, racha_por_pixel)
}

## ----------------------------------------------------------------------
## 3.3 Formula EXACTA del indice de riesgo (replica celda AC5 del Excel)
##     para Ajo / Fusariosis (pudricion basal).
## ----------------------------------------------------------------------
calcular_indice_ajo_fusariosis <- function(t_media_periodo, dias_secos_periodo,
                                            precip_acum_periodo) {

  cond_t      <- t_media_periodo    >= 25   # T media >= 25 °C
  cond_secos  <- dias_secos_periodo >= 10   # >= 10 dias secos consecutivos
  cond_precip <- precip_acum_periodo < 20   # precipitacion acumulada < 20 mm

  # pesos: 1, 1, 1 (iguales, tal cual columnas Z5:AB5 del Excel)
  (cond_t * 1 + cond_secos * 1 + cond_precip * 1) / 3
}

## ----------------------------------------------------------------------
## 3.4 Clasificacion de alerta tipo semaforo (replica celda S5 del Excel)
##
##     OJO con los pixeles NA: ya no se convierten en la categoria "Sin
##     datos" (como en una version anterior de este script). terra::ifel()
##     deja un pixel en NA si su valor de entrada era NA, y aqui eso es
##     justamente lo que queremos: los pixeles fuera de la region de ajo
##     (NA por el mask() de la seccion 4) deben quedar transparentes en el
##     mapa, no pintados de un color solido. Si dentro de la region llegara
##     a faltar un dato climatico puntual, tambien quedara NA/transparente
##     por la misma razon (en vez de aparentar que ahi "no hay riesgo").
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

## paleta de colores tipo semaforo, en el mismo orden de las categorias
## (ya no incluye "Sin datos": esos pixeles quedan NA -> transparentes,
## ver na.value = "transparent" en scale_fill_manual() mas abajo)
colores_alerta <- c("Normal"      = "forestgreen",
                    "Alerta baja" = "gold",
                    "Alerta media"= "orange",
                    "Alerta alta" = "red3")


# ============================================================================
# 4. CARGA DE SHAPEFILES Y DE LAS SERIES CLIMATICAS DIARIAS
# ============================================================================

region_ajo <- st_read(ruta_shp_ajo, quiet = TRUE)

# mapa base de municipios: solo da contexto geografico, no participa en el
# calculo del indice. Se reproyecta al CRS de la region de ajo por si vienen
# en sistemas de referencia distintos.
municipios <- st_read(ruta_shp_municipios, quiet = TRUE)
municipios <- st_transform(municipios, st_crs(region_ajo))

# extension del mapa: la region de ajo + un margen de contexto alrededor
bbox_ajo  <- st_bbox(region_ajo)
xlim_mapa <- c(bbox_ajo["xmin"] - buffer_mapa_grados, bbox_ajo["xmax"] + buffer_mapa_grados)
ylim_mapa <- c(bbox_ajo["ymin"] - buffer_mapa_grados, bbox_ajo["ymax"] + buffer_mapa_grados)

# rango real de datos a cargar (incluye el buffer para la racha del 1er periodo)
fecha_carga_desde <- fecha_inicio - dias_buffer

# Tmax, Tmin y HR usan fecha con guiones (AAAA-MM-DD) -> sirven los valores
# por defecto de cargar_serie_diaria(). Precipitacion (CHIRPS) usa fecha con
# puntos (AAAA.MM.DD), asi que se le indica el patron/formato propio.
tmax   <- cargar_serie_diaria(carpeta_tmax,   fecha_carga_desde, fecha_fin)
tmin   <- cargar_serie_diaria(carpeta_tmin,   fecha_carga_desde, fecha_fin)
hr     <- cargar_serie_diaria(carpeta_hr,     fecha_carga_desde, fecha_fin)  # no entra en la formula de Ajo, se guarda para referencia
precip <- cargar_serie_diaria(carpeta_precip, fecha_carga_desde, fecha_fin,
                               patron_fecha  = "\\d{4}\\.\\d{2}\\.\\d{2}",
                               formato_fecha = "%Y.%m.%d")

# recortar y enmascarar todo a la region productora de ajo
region_vect <- vect(region_ajo)
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

# construir la secuencia de periodos: desde fecha_inicio hasta fecha_fin,
# en bloques de 'periodo_dias' dias (el ultimo periodo puede quedar mas corto)
inicios_periodo <- seq(fecha_inicio, fecha_fin, by = periodo_dias)

resultados       <- list()   # aqui se guarda el raster de indice de cada periodo
resultados_alerta<- list()   # aqui se guarda el raster de alerta de cada periodo
resumen_tabla     <- data.frame()

for (i in seq_along(inicios_periodo)) {

  ini <- inicios_periodo[i]
  fin <- min(ini + periodo_dias - 1, fecha_fin)

  idx <- which(fechas_disponibles >= ini & fechas_disponibles <= fin)
  if (length(idx) == 0) next

  # variables agregadas del periodo, pixel a pixel
  t_media_periodo    <- mean(t_media_diaria[[idx]])          # promedio del periodo
  precip_acum_periodo<- sum(precip[[idx]])                    # suma del periodo
  dias_secos_periodo <- racha_seca[[max(idx)]]                 # racha vigente al ULTIMO dia del periodo

  indice <- calcular_indice_ajo_fusariosis(t_media_periodo, dias_secos_periodo,
                                            precip_acum_periodo)
  names(indice) <- paste0("indice_", format(ini, "%Y%m%d"), "_", format(fin, "%Y%m%d"))

  alerta <- clasificar_alerta(indice)
  names(alerta) <- paste0("alerta_", format(ini, "%Y%m%d"), "_", format(fin, "%Y%m%d"))

  resultados[[i]]        <- indice
  resultados_alerta[[i]] <- alerta

  # tabla resumen: superficie (%) de la region en cada categoria de alerta
  frecuencias <- freq(alerta)
  total_celdas <- sum(frecuencias$count)
  fila <- data.frame(
    periodo_inicio = ini,
    periodo_fin    = fin,
    categoria      = frecuencias$value,
    pct_area       = round(100 * frecuencias$count / total_celdas, 1)
  )

  # ---- diagnostico: que tan seguido se cumple CADA condicion por separado ----
  # Esto ayuda a entender por que "Alerta alta" (las 3 condiciones a la vez)
  # puede ser rara: basta con que UNA condicion casi nunca se cumpla (por
  # ejemplo, si la zona de ajo rara vez supera los 25 C de T media) para que
  # el indice nunca llegue a 1, aunque la formula este bien programada.
  cond_t_r      <- t_media_periodo     >= 25
  cond_secos_r  <- dias_secos_periodo  >= 10
  cond_precip_r <- precip_acum_periodo <  20
  fila$pct_area_cond_temp   <- round(100 * as.numeric(global(cond_t_r,      "mean", na.rm = TRUE)), 1)
  fila$pct_area_cond_secos  <- round(100 * as.numeric(global(cond_secos_r,  "mean", na.rm = TRUE)), 1)
  fila$pct_area_cond_precip <- round(100 * as.numeric(global(cond_precip_r,"mean", na.rm = TRUE)), 1)

  resumen_tabla <- rbind(resumen_tabla, fila)

  # ----- mapa del periodo -----
  # Orden de las capas: primero el fondo de municipios (contexto), despues
  # el raster de alerta (transparente fuera de la region de ajo, porque sus
  # pixeles NA usan na.value = "transparent"), y por ultimo el contorno de
  # la region de ajo encima de todo. coord_sf() recorta la vista a la region
  # de ajo + el margen de contexto definido en buffer_mapa_grados.
  mapa <- ggplot() +
    geom_sf(data = municipios, fill = "grey97", color = "grey60", linewidth = 0.2) +
    geom_spatraster(data = alerta) +
    geom_sf(data = region_ajo, fill = NA, color = "black", linewidth = 0.5) +
    scale_fill_manual(values = colores_alerta, na.translate = FALSE,
                       na.value = "transparent", name = "Estado de alerta") +
    coord_sf(xlim = xlim_mapa, ylim = ylim_mapa, expand = FALSE) +
    labs(title = "Riesgo de Fusariosis (pudricion basal) en Ajo",
         subtitle = paste0("Periodo: ", format(ini, "%d-%b-%Y"), " a ", format(fin, "%d-%b-%Y")),
         caption = "Indice = coincidencia de T media >= 25C, >=10 dias secos consecutivos y precip. < 20 mm") +
    theme_minimal()

  # ---- mostrar el mapa en pantalla ANTES de guardarlo ----
  # Esto permite ver cada periodo aparecer en la pestaña "Plots" de RStudio
  # a medida que el bucle avanza, sin esperar a que termine todo el proceso.
  message(sprintf("Periodo %d/%d: %s a %s",
                   i, length(inicios_periodo), format(ini, "%d-%b-%Y"), format(fin, "%d-%b-%Y")))
  print(mapa)
  Sys.sleep(0.3)   # pequeña pausa para poder ver cada mapa antes de pasar al siguiente
                   # (ajustar o poner en 0 si se quiere maxima velocidad)

  ggsave(file.path(carpeta_salida,
                    paste0("mapa_alerta_ajo_", format(ini, "%Y%m%d"), "_", format(fin, "%Y%m%d"), ".png")),
         mapa, width = 7, height = 6, dpi = 150)
}


# ============================================================================
# 7. GUARDAR RESULTADOS
# ============================================================================

# rasters de indice y de alerta, todos los periodos apilados en un solo archivo
writeRaster(rast(resultados),        file.path(carpeta_salida, "indice_riesgo_ajo_2025.tif"), overwrite = TRUE)
writeRaster(rast(resultados_alerta), file.path(carpeta_salida, "estado_alerta_ajo_2025.tif"),  overwrite = TRUE)

# tabla resumen de superficie en riesgo por periodo
write.csv(resumen_tabla, file.path(carpeta_salida, "resumen_superficie_alerta_ajo_2025.csv"),
          row.names = FALSE)

message("Listo. Mapas, rasters y tabla resumen guardados en: ", carpeta_salida)
