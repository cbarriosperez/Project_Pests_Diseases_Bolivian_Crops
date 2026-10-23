# Sistema de Alerta Temprana para Plagas y Enfermedades en Cultivos Bolivianos 🌱🇧🇴

[![GitHub Pages](https://img.shields.io/badge/Despliegue-GitHub%20Pages-success)](https://cbarriosperez.github.io/Project_Pests_Diseases_Bolivian_Crops/)

**[Visita la página web oficial del proyecto aquí](https://cbarriosperez.github.io/Project_Pests_Diseases_Bolivian_Crops/)**

## 📖 Descripción del Proyecto

Este repositorio aloja una plataforma web informativa y educativa dedicada a documentar, describir y concientizar sobre las principales **plagas y enfermedades que afectan a los cultivos agrícolas en Bolivia**. 

El objetivo principal de este proyecto es proporcionar un recurso accesible y centralizado para agricultores, estudiantes de agronomía, investigadores y profesionales del sector agrícola, facilitando la identificación de amenazas fitosanitarias y promoviendo el conocimiento sobre la sanidad vegetal adaptada a la realidad boliviana.

## ✨ Características Principales

- **Presentación Clara:** Una introducción detallada ([ver sección](#presentacion)) sobre el contexto agrícola de Bolivia y la importancia del control de plagas.
- **Catálogo de Afectaciones:** Información estructurada sobre diferentes patógenos, insectos y enfermedades. *(Nota: ajusta esto según lo que tenga tu página)*
- **Enfoque Local:** Datos y descripciones contextualizadas a las regiones geográficas y cultivos específicos de Bolivia.
- **Diseño Responsivo:** Interfaz accesible desde dispositivos móviles y computadoras de escritorio.

## 🌦️ Pronóstico a 16 días (actualización diaria)

La sección **Explorar Datos → Pronóstico 16 días** muestra precipitación, temperatura máxima y mínima y humedad relativa pronosticadas para los municipios productores de los 15 cultivos priorizados.

- **Fuente:** [Open-Meteo](https://open-meteo.com/) (CC BY 4.0). Modelos deterministas: GFS, ECMWF IFS, AIFS-Single, ICON Global y GEM Global, más un multimodelo (promedio simple). Los modelos de ensamble (GEFS, ECMWF-ENS, AIFS-ENS, GEPS) se activan automáticamente al configurar el *secret* `OPEN_METEO_API_KEY`.
- **Automatización:** el workflow [`.github/workflows/pronostico_diario.yml`](.github/workflows/pronostico_diario.yml) corre todos los días a las 10:00 UTC (06:00 en La Paz), ejecuta [`Scripts/pronostico/pronostico_municipios.py`](Scripts/pronostico/pronostico_municipios.py) y hace commit de `data/pronostico/pronostico_latest.json`. El historial de commits de ese archivo funciona como archivo de pronósticos pasados.
- **Geometría:** `data/geo/` contiene los municipios simplificados y el índice cultivo → municipios, generados con [`Scripts/pronostico/preparar_geometria.py`](Scripts/pronostico/preparar_geometria.py) a partir de los shapefiles de `data/geo/shapefiles/`.
- **Corrida manual:** pestaña *Actions → Pronóstico diario 16 días → Run workflow*.

## 👨‍💻 Autor
Desarrollado y mantenido por **[Camilo Barrios-Pérez](https://github.com/cbarriosperez)**.

Si esta plataforma te ha sido útil, ¡no dudes en darle una ⭐ al repositorio!



