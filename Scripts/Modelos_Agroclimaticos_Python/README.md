# Modelos Agroclimáticos en Python (pronóstico a 16 días) 🌾🌡️

## 📖 Descripción

Versión en **Python** de los [Modelos Agroclimáticos en R](../Modelos_Agroclimaticos_R/), adaptada para trabajar con el **pronóstico diario por municipio** que genera [`pronostico_municipios.py`](../pronostico/pronostico_municipios.py).

Cada script calcula el índice de riesgo de una plaga o enfermedad con la **misma fórmula, pesos y semáforo** que su par en R (hoja "Calculadora_Indice de riesgo" del Excel). Cambia lo siguiente:

| | Versión R | Versión Python |
|---|---|---|
| Datos de entrada | GeoTIFF diarios (CHIRPS, AgERA5) | Pronóstico Open-Meteo (`data/pronostico/pronostico_latest.json`) |
| Unidad espacial | Píxel, dentro del shapefile del cultivo | Municipio productor (`data/geo/cultivos_municipios.json`) |
| Periodo | Bloques de N días a lo largo del año | Una ventana fija: los primeros N días del pronóstico, desde hoy |
| Modelos climáticos | Uno (serie histórica) | Cada modelo del pronóstico + MULTIMODELO |

Como el pronóstico se actualiza cada día, las alertas también.

## 📂 Scripts por cultivo

| Cultivo | Plaga / enfermedad | Ventana | Script |
|---|---|---|---|
| 🧄 Ajo | Fusariosis | 15 días | [`Script_Indice_Riesgo_Ajo_Pronostico.py`](./Ajo/Script_Indice_Riesgo_Ajo_Pronostico.py) |
| 🍌 Banano | Sigatoka negra | 10 días | [`Script_Indice_Riesgo_Banano_Pronostico.py`](./Banano/Script_Indice_Riesgo_Banano_Pronostico.py) |
| ☕ Café | Broca del café | 14 días* | [`Script_Indice_Riesgo_Cafe_Pronostico.py`](./Cafe/Script_Indice_Riesgo_Cafe_Pronostico.py) |
| 🧅 Cebolla | Mildiu (polvillo negro) | 10 días | [`Script_Indice_Riesgo_Cebolla_Pronostico.py`](./Cebolla/Script_Indice_Riesgo_Cebolla_Pronostico.py) |
| 🍑 Durazno | Arañuela roja | 10 días | [`Script_Indice_Riesgo_Durazno_Aranuela_Pronostico.py`](./Durazno/Aranuela_roja/Script_Indice_Riesgo_Durazno_Aranuela_Pronostico.py) |
| 🍑 Durazno | Oídio | 10 días | [`Script_Indice_Riesgo_Durazno_Oidio_Pronostico.py`](./Durazno/Oidio/Script_Indice_Riesgo_Durazno_Oidio_Pronostico.py) |
| 🍓 Frutilla | Botrytis (moho gris) | 10 días | [`Script_Indice_Riesgo_Frutilla_Pronostico.py`](./Frutilla/Script_Indice_Riesgo_Frutilla_Pronostico.py) |
| 🥬 Lechuga | Sclerotinia sclerotiorum | 15 días | [`Script_Indice_Riesgo_Lechuga_Pronostico.py`](./Lechuga/Script_Indice_Riesgo_Lechuga_Pronostico.py) |
| 🌽 Maíz choclo | Gusano cogollero | 7 días | [`Script_Indice_Riesgo_Maiz_Choclo_Pronostico.py`](./Maiz_Choclo/Script_Indice_Riesgo_Maiz_Choclo_Pronostico.py) |
| 🌽 Maíz de grano | Roya | 10 días | [`Script_Indice_Riesgo_Maiz_Grano_Pronostico.py`](./Maiz_de_Grano/Script_Indice_Riesgo_Maiz_Grano_Pronostico.py) |
| 🍊 Naranja | Mosca de la fruta | 15 días | [`Script_Indice_Riesgo_Naranja_Pronostico.py`](./Naranja/Script_Indice_Riesgo_Naranja_Pronostico.py) |
| 🥔 Papa | Gorgojo de los Andes | 10 días | [`Script_Indice_Riesgo_Papa_Pronostico.py`](./Papa/Script_Indice_Riesgo_Papa_Pronostico.py) |
| 🌾 Quinua | Ticona | 10 días | [`Script_Indice_Riesgo_Quinua_Pronostico.py`](./Quinua/Script_Indice_Riesgo_Quinua_Pronostico.py) |
| 🍅 Tomate | Polilla (Tuta absoluta) | 10 días | [`Script_Indice_Riesgo_Tomate_Pronostico.py`](./Tomate/Script_Indice_Riesgo_Tomate_Pronostico.py) |
| 🍇 Uva | Mildiu de la uva | 7 días | [`Script_Indice_Riesgo_Uva_Pronostico.py`](./Uva/Script_Indice_Riesgo_Uva_Pronostico.py) |

\* El modelo de café usa un periodo de 30 días, más largo que el pronóstico. Se evalúan 14 días y el umbral de precipitación acumulada se escala en proporción (30–80 mm / 30 días → 14–37,3 mm / 14 días).

## 🚀 Uso

Requiere solo `numpy` y `pandas`.

```bash
# Un cultivo: genera un CSV en Salidas_Modelos_Agroclimaticos/Pronostico/
python Scripts/Modelos_Agroclimaticos_Python/Papa/Script_Indice_Riesgo_Papa_Pronostico.py

# Todos los cultivos: genera data/alertas/alertas_latest.json para la página web
python Scripts/Modelos_Agroclimaticos_Python/correr_modelos.py
```

En GitHub Actions, `correr_modelos.py` corre a diario justo después del pronóstico (`.github/workflows/pronostico_diario.yml`).

## ⚠️ Notas

- **Horizonte de cada modelo.** Si un modelo no cubre la ventana completa (p. ej. ICON, ~7 días), su índice queda sin dato (`-1` en el JSON), igual que en R cuando falta un día del periodo.
- **Racha de días secos** (Ajo, Durazno–arañuela, Naranja, Tomate): se cuenta desde el primer día del pronóstico, sin días previos.

## 👨‍💻 Autor
Desarrollado y mantenido por **[Camilo Barrios-Pérez](https://github.com/cbarriosperez)**.
