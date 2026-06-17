# Temperatura Fundos Aquanqa — Predicción y Monitoreo

Dashboard Streamlit para monitoreo y predicción de temperatura máxima (Tmax) y mínima (Tmin) en los fundos de Aquanqa, usando datos meteorológicos de sensores y normales climatológicas SENAMHI (1991–2020).

---

## Estructura del proyecto

```
PREDICCION_TEMPERATURA/
├── temperatura_senamhi_Tmax-Tmin.py   ← Script principal (este README lo documenta)
├── requirements.txt
├── assets/
│   ├── Metereologia_Prize.xlsx         ← Datos meteorológicos del sensor (15 min)
│   ├── NORMALES_1991_2020.xlsx         ← Normales climatológicas SENAMHI
│   └── .prophet_cache/                 ← Caché automática de modelos Prophet
│       ├── forecasts.pkl
│       └── data_hash.txt
```

---

## Instalación

### 1. Requisitos previos

- Python 3.10 o superior
- `pip` actualizado

### 2. Instalar dependencias

```bash
pip install -r requirements.txt
```

El `requirements.txt` incluye:

| Paquete | Versión mínima | Uso |
|---|---|---|
| streamlit | 1.28.0 | Framework de la app |
| pandas | 2.1.0 | Manejo de datos |
| numpy | 1.24.0 | Cálculos numéricos |
| scipy | 1.11.0 | Interpolación spline |
| prophet | 1.1.0 | Modelo de predicción |
| cmdstanpy | 1.2.0 | Backend de Prophet |
| pystan | 2.19.0 | Backend alternativo de Prophet |
| plotly | 5.18.0 | Gráficos interactivos |
| openpyxl | 3.0.0 | Lectura/escritura Excel |

### 3. Instalar paquetes adicionales (no están en requirements.txt)

```bash
pip install folium streamlit-folium
```

Estos son necesarios para el mapa interactivo de distritos de Perú.

### 4. Instalar CmdStan (requerido por Prophet)

Prophet usa CmdStan como motor de cómputo. Si no se instala solo, ejecutar:

```python
import cmdstanpy
cmdstanpy.install_cmdstan()
```

O desde terminal:

```bash
python -c "import cmdstanpy; cmdstanpy.install_cmdstan()"
```

---

## Archivos de datos requeridos

Los archivos deben estar ubicados en `assets/` dentro del mismo directorio del script.

### `Metereologia_Prize.xlsx`

Datos del sensor meteorológico cada 15 minutos. Debe tener la hoja llamada **`Prize_Climatology`** con las columnas:

| Columna | Descripción |
|---|---|
| `Fecha-Hora` | Fecha y hora en formato `M/DD/YYYY HH:MM` (americano) |
| `Fundo` | Nombre del fundo (ej. `Ayllu Allpa`, `Vivadis`) |
| `Empresa` | Nombre de la empresa |
| `Temp-C` | Temperatura promedio (°C) |
| `TempAlta-C` | Temperatura alta / Tmax del intervalo (°C) |
| `TempBaja-C` | Temperatura baja / Tmin del intervalo (°C) |
| `ET-mm` | Evapotranspiración (mm) — opcional, se rellena con 0 si falta |

### `NORMALES_1991_2020.xlsx`

Normales climatológicas SENAMHI 1991–2020. Debe tener dos hojas:
- **`TMAX`**: normales de temperatura máxima
- **`TMIN`**: normales de temperatura mínima

Cada hoja debe tener una fila de encabezado con la columna `DISTRITO` y columnas con los nombres de los meses (Enero, Febrero, … Diciembre).

---

## Ejecución

```bash
streamlit run temperatura_senamhi_Tmax-Tmin.py
```

La app abre en el navegador en `http://localhost:8501`.

---

## Cómo funciona la app

### Flujo principal al iniciar

1. Lee `NORMALES_1991_2020.xlsx` desde disco y carga normales de **Guadalupe** (referencia fija).
2. Lee `Metereologia_Prize.xlsx` desde disco y detecta los fundos disponibles.
3. Agrega los datos de 15 min a diarios (Tmax, Tmin, ET).
4. Verifica si existe caché de Prophet válida (por hash MD5 del Excel). Si los datos no cambiaron, carga los modelos del caché sin reentrenar. Si cambiaron o es la primera vez, entrena todos los modelos.
5. Muestra los 4 tabs de la interfaz.

### Tabs disponibles

| Tab | Contenido |
|---|---|
| 🔴 Temperatura Máxima | Gráfico Tmax real + climatología SENAMHI + predicción Prophet + mapa distritos |
| 🔵 Temperatura Mínima | Ídem para Tmin |
| 💧 ET | Evapotranspiración diaria + predicción del mes siguiente por fundo |
| 📋 Datos | Tabla filtrable + descarga Excel/CSV |

### Sidebar

- **Fundos**: selección múltiple de los fundos a mostrar.
- **Climatología SENAMHI**: siempre muestra Guadalupe como referencia. Permite seleccionar una segunda estación por Sector → Departamento → Distrito para comparar.
- **Corrección BIAS**: muestra el sesgo sistemático detectado en la sección de validación (se activa después de abrir la validación Prophet).

---

## Documentación de funciones

### Utilidades

#### `_hex_to_rgba(hex_color, alpha)`
Convierte un color hexadecimal (`#RRGGBB`) a formato `rgba(r,g,b,alpha)` para usar en Plotly.

#### `_normalizar(s)`
Elimina tildes, convierte a mayúsculas y elimina espacios de un string. Usado para comparar nombres de distritos/meses sin sensibilidad a acentos.

---

### Lectura de datos

#### `leer_meteo_bytes_optimizado(file_bytes, filename, sheet)`
Lee el archivo meteorológico desde bytes en memoria (caché de Streamlit).
- Si es `.csv`: detecta separador automáticamente, prueba UTF-8 y Latin-1.
- Si es `.xlsx`: lee la hoja especificada sin parsear fechas (se controlan en el paso siguiente).

#### `cargar_meteoro_optimizado(_file_bytes, filename, sheet, fundos_sel, min_reg)`
Pipeline completo de carga y procesamiento de datos de sensor:
1. Parsea `Fecha-Hora` en formato americano (`M/DD/YYYY`).
2. Convierte columnas de temperatura a numérico.
3. Filtra rangos válidos: Tmax 10–45°C, Tmin 5–35°C.
4. Cuenta registros por fundo+fecha (día completo = 96 registros de 15 min).
5. Incluye siempre el último día de cada fundo aunque tenga menos de `min_reg` registros.
6. Agrega a diario: `Tmax` (máximo), `Tmin` (mínimo), `ET` (suma).
7. Aplica suavizado rolling de 3 días (columnas `Tmax_smooth`, `Tmin_smooth`).

#### `_leer_normales_desde_disco(path)`
Lee el archivo de normales SENAMHI desde disco con caché de Streamlit.

#### `cargar_normales(_file_bytes, hoja)`
Carga normales climatológicas del distrito **Guadalupe** (fijo) desde la hoja indicada (`TMAX` o `TMIN`). Devuelve valores medios mensuales y bandas Q1/Q3 (±0.6745 × 1.5 σ).

#### `cargar_catalogo_normales(_file_bytes, hoja)`
Carga el catálogo completo de estaciones SENAMHI estructurado como `{sector: {departamento: [distritos]}}`, excluyendo Guadalupe (que siempre es la referencia fija).

#### `cargar_normales_dinamico(_file_bytes, distrito_sel)`
Versión flexible de `cargar_normales`: carga normales para cualquier distrito seleccionado dinámicamente desde el sidebar. Devuelve medias y bandas Q1/Q3 para Tmax y Tmin.

#### `cargar_geojson_peru()`
Descarga desde GitHub el GeoJSON de distritos de Perú (primera vez — luego queda en caché de Streamlit). Necesita conexión a internet.

---

### Spline climatológico

#### `spline_diario_cached(fecha_inicio_str, fecha_fin_str, valores_tuple)`
Interpola 12 valores mensuales (medias o percentiles) a una serie diaria usando **CubicSpline** de scipy. La interpolación se extiende un año antes y después del rango pedido para evitar efectos de borde. Resultado cacheado por Streamlit según los parámetros de entrada.

---

### Modelos Prophet

#### `entrenar_prophet_opt(_serie_hash, _serie_bytes, dias_pred, variable, fundo)`
Entrena un modelo **Facebook Prophet** para una serie temporal de temperatura o ET:
- Aplica clipping de outliers (±3σ) antes de entrenar.
- Asigna peso 2.5× a los últimos 180 días para que el modelo priorice el comportamiento reciente.
- Realiza **validación cruzada walk-forward** (15 pliegues) y calcula MAE por horizonte.
- Entrena el modelo final con todos los datos e incluye estacionalidades mensual (Fourier 5), semanal (Fourier 4) y bimensual (Fourier 3).
- Devuelve: forecast futuro, forecast histórico completo, MAE real, MAE por día de horizonte.

#### `entrenar_todos_optimizado(dia, dias_pred)`
Itera sobre todos los fundos × variables (`Tmax`, `Tmin`, `ET`) y llama a `entrenar_prophet_opt` con barra de progreso. Recorta cada serie hasta el último día del mes anterior para que las predicciones arranquen el 1 del mes actual.

#### `guardar_cache_prophet(forecasts, data_hash)` / `cargar_cache_prophet(data_hash)`
Persisten el diccionario de forecasts en `assets/.prophet_cache/forecasts.pkl`. La caché es invalidada automáticamente si el MD5 del archivo Excel cambia (datos nuevos).

---

### Visualización

#### `generar_figura(variable, dia, media_mensual, q1_mensual, q3_mensual, dias_pred, forecasts_cache, ...)`
Genera la figura principal de Tmax o Tmin con subplots por fundo. Por cada fundo:
- Dibuja la banda climatológica SENAMHI (Q1–Q3) y la media como spline diario.
- Dibuja opcionalmente una segunda climatología de la estación seleccionada en el sidebar.
- Dibuja los datos reales suavizados con rolling 3 días.
- Aplica ajuste de residuos a la predicción Prophet (patrón de los últimos 21 días) y corrección de BIAS si está disponible.
- Muestra la predicción como línea punteada con banda IC 95%.
- Añade línea vertical en el último dato real.
- Soporta modo `linea` (por defecto) y modo `candlestick` (range plot con error bars).

#### `generar_range_plot(variable, historico_df, prediccion_df, fundo_name, row, fig)`
Agrega trazas de tipo Range Plot (barras de error asimétricas) al subplot indicado, mostrando la amplitud Tmax–Tmin alrededor de la curva suavizada.

#### `generar_tab_et(dia, forecasts_cache, dias_pred_ui)`
Genera el tab de evapotranspiración diaria por fundo con:
- Serie real con área rellena desde cero.
- Línea de promedio histórico.
- Predicción Prophet del mes actual ajustada con el patrón cíclico de los últimos 30 días.
- Banda IC 95% y conector desde el último dato real.
- KPIs de ET total y promedio diario por fundo.

#### `generar_mapa_distritos(geojson, distrito_fijo, distrito_din, variable, label_clim2)`
Genera un mapa **Folium** con imagen satelital Esri. Resalta en rojo/azul el distrito fijo (Guadalupe) y en naranja el distrito dinámico seleccionado. Marca los tres fundos Aquanqa con íconos azules.

---

### Validación

#### `generar_seccion_validacion(variable, dia, forecasts_cache, dias_pred_ui)`
Muestra una sección expandible de validación Prophet con dos tabs:
- **Último mes**: entrena solo con datos anteriores al mes anterior y predice ese mes completo.
- **Últimos 3 meses**: ventana ampliada.

#### `_validar_ventana(variable, dia, ventana, key_suffix)`
Motor interno de validación para una ventana temporal específica. Por cada fundo:
1. Entrena con datos anteriores al período de validación.
2. Predice el período.
3. Calcula MAE, RMSE, MBE (BIAS), MAPE y precisión.
4. Guarda el BIAS en `st.session_state['bias_correccion']` para corregir automáticamente las predicciones futuras.
5. Muestra gráfico de error diario firmado (verde = sobreestima, rojo = subestima).
6. Muestra box plot y gráfico apilado de rangos de error por fundo.

---

### Exportación

#### `exportar_excel(df_hist, df_pred)`
Genera un Excel con dos hojas formateadas:
- **Por_Fechas**: datos diarios reales y predichos por fundo (amarillo = real, naranja = predicción).
- **Por_Horas**: predicción horaria generada a partir de la diaria mediante spline + offset sinusoidal (Tmax pico al mediodía, Tmin pico de madrugada).

#### `_generar_pred_horaria(df_pred)`
Convierte predicciones diarias en horarias interpolando con CubicSpline y añadiendo un offset sinusoidal proporcional al ancho del IC 95%, simulando el ciclo diurno.

---

## Parámetros clave (constantes)

| Constante | Valor | Descripción |
|---|---|---|
| `SHEET_NAME` | `Prize_Climatology` | Hoja del Excel meteorológico |
| `MIN_REGISTROS` | 80 | Mínimo de registros 15-min por día (de 96 posibles) |
| `ROLLING_DIAS` | 3 | Ventana de suavizado rolling |
| `SIGMA_CLIM` | 1.5 | Número de sigmas para banda climatológica |
| `TMAX_MIN / TMAX_MAX` | 10 / 45 °C | Rango válido para temperatura máxima |
| `TMIN_MIN / TMIN_MAX` | 5 / 35 °C | Rango válido para temperatura mínima |
| `DISTRITO_BUSCAR` | `GUADALUPE` | Distrito fijo de referencia SENAMHI |
| `dias_pred_ui` | 30 | Días de predicción Prophet (hardcodeado en la app) |

---

## Solución de problemas frecuentes

**Prophet no entrena / error de CmdStan**
```bash
python -c "import cmdstanpy; cmdstanpy.install_cmdstan()"
```

**`streamlit_folium` no encontrado**
```bash
pip install folium streamlit-folium
```

**"Archivo no encontrado" al iniciar**
Verificar que `assets/Metereologia_Prize.xlsx` y `assets/NORMALES_1991_2020.xlsx` existen en la misma carpeta que el script.

**Modelos no se actualizan con datos nuevos**
El caché Prophet se invalida automáticamente cuando el archivo Excel cambia. Si aun así no actualiza, borrar manualmente `assets/.prophet_cache/`.

**El mapa no carga**
La app necesita acceso a internet para descargar el GeoJSON de distritos de Perú. Si hay restricción de red, el mapa se omite silenciosamente.

---

*Elaborado por la Gerencia de Planificación, Control Operacional y Gestión — Aquanqa*
