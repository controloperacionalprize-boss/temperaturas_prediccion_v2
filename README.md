# Predicción de Temperatura — Fundos Aquanqa

Aplicación **Streamlit** (`temperatura_v2.py`) para monitoreo, análisis y predicción de temperatura máxima (Tmax), temperatura mínima (Tmin) y evapotranspiración (ET) en los fundos agrícolas Arena Azul, Vivadis, Santa Teresa y Ayllu Allpa (Paijan, La Libertad).

---

## Índice

1. [Arquitectura general](#arquitectura-general)
2. [Flujo de datos](#flujo-de-datos)
3. [Constantes y parámetros](#constantes-y-parámetros)
4. [Módulo ENFEN — ajuste climático](#módulo-enfen--ajuste-climático)
5. [Funciones clave](#funciones-clave)
6. [Modelo Prophet — pipeline completo](#modelo-prophet--pipeline-completo)
7. [Validación walk-forward y métricas](#validación-walk-forward-y-métricas)
8. [Visualizaciones](#visualizaciones)
9. [Mapas Folium](#mapas-folium)
10. [Exportación Excel](#exportación-excel)
11. [Monitor ENFEN automático](#monitor-enfen-automático)
12. [Instalación y dependencias](#instalación-y-dependencias)
13. [Configuración de secretos](#configuración-de-secretos)
14. [Estructura de archivos requerida](#estructura-de-archivos-requerida)
15. [Casuísticas y comportamiento esperado](#casuísticas-y-comportamiento-esperado)

---

## Arquitectura general

```
temperatura_v2.py  (Streamlit backend Python)
│
├── Fuente de datos: Excel subido por el usuario
│   └── Hoja: Prize_Climatology — registros cada 15 min (Fecha-Hora, TempAlta-C, TempBaja-C, ET-mm)
│
├── Climatología de referencia: NORMALES_1991_2020.xlsx (SENAMHI)
│   ├── Hoja TMAX — normales mensuales de Tmax por distrito
│   └── Hoja TMIN — normales mensuales de Tmin por distrito
│   └── Distrito fijo: GUADALUPE (La Libertad)
│
├── Geometría: KMZ de módulos AQ1/AQ2 (GitHub privado o subido por usuario)
│   └── Polígonos de fundos fusionados con Shapely
│
├── Motor de predicción: Facebook Prophet (serie temporal diaria)
│   ├── Anomalía = temperatura - climatología armónica
│   ├── Corrección de bias: tendencia reciente + estacional + ENFEN
│   └── Validación walk-forward honesta (1 mes / 3 meses)
│
└── Visualizaciones:
    ├── Plotly — gráficos de temperatura y ET con predicción
    ├── Folium — mapa de riesgo térmico y ET por fundo
    └── Excel exportado con datos diarios + predicción horaria interpolada
```

---

## Flujo de datos

```
1. Usuario sube Excel meteorológico (.xlsx o .csv)
        │
        ▼
2. cargar_meteoro_optimizado()
   ├── Parseo fecha M/DD/YYYY HH:MM (formato americano)
   ├── Filtro de rango válido (Tmax: 10-45°C, Tmin: 5-35°C)
   ├── Mínimo 80 registros/día (equivale a ≥20h de cobertura)
   ├── Siempre incluye el último día por fundo (aunque esté incompleto)
   └── Agregación diaria: Tmax=max, Tmin=min, ET=sum
        │
        ▼
3. cargar_normales() — Climatología SENAMHI Guadalupe
   ├── 12 valores mensuales de Tmax y Tmin para referencia
   ├── Banda de confianza: ±DELTA_Q = ±(1.5 × 0.6745) ≈ ±1.01°C
   └── Spline cúbico diario (CubicSpline) para graficar
        │
        ▼
4. entrenar_todos_optimizado()
   ├── Para cada (fundo, variable): Tmax, Tmin, ET
   │   ├── calcular_climatologia_armonica() — regresión Fourier multi-año
   │   ├── Serie de anomalías (temperatura - climatología)
   │   ├── Validación cruzada walk-forward (últimos N pasos)
   │   ├── Modelo Prophet final sobre toda la serie de anomalías
   │   ├── Re-añadir climatología → temperatura absoluta
   │   └── Corrección de bias (tendencia lineal + estacional + ENFEN)
        │
        ▼
5. generar_figura() — gráficos Plotly por variable y fundo
   ├── Datos reales con suavizado rolling 3 días
   ├── Climatología SENAMHI con banda Q1-Q3
   ├── Segunda climatología dinámica (distrito opcional)
   └── Predicción con banda IC 95%
        │
        ▼
6. Mapas Folium
   ├── generar_mapa_distritos() — mapa de riesgo térmico por fundo
   └── generar_mapa_et()        — mapa de riesgo ET por fundo
        │
        ▼
7. exportar_excel()
   ├── Hoja "Por_Fechas" — histórico + predicción diaria
   └── Hoja "Por_Horas" — predicción horaria (spline cúbico)
```

---

## Constantes y parámetros

### Parámetros de calidad de datos

| Constante | Valor | Descripción |
|---|---|---|
| `SHEET_NAME` | `"Prize_Climatology"` | Nombre de hoja del Excel meteorológico |
| `MIN_REGISTROS` | `80` | Registros mínimos por día (de 96 posibles a 15min) |
| `ROLLING_DIAS` | `3` | Ventana suavizado rolling centrado |
| `TMAX_MIN/MAX` | `10.0 / 45.0 °C` | Rango válido para temperatura máxima |
| `TMIN_MIN/MAX` | `5.0 / 35.0 °C` | Rango válido para temperatura mínima |

### Hiperparámetros Prophet (`PROPHET_PARAMS`)

| Parámetro | Valor | Descripción |
|---|---|---|
| `changepoint_prior_scale` | `0.15` | Flexibilidad para detectar cambios de tendencia |
| `changepoint_range` | `0.95` | % del historial donde se buscan changepoints |
| `seasonality_prior_scale` | `2.0` | Fuerza de la estacionalidad |
| `yearly_seasonality` | `8` | Armónicos anuales |
| `fourier_monthly` | `3` | Armónicos de estacionalidad mensual adicional |
| `bias_window_dias` | `21` | Ventana para estimar tendencia de residuos recientes |
| `interval_width` | `0.90` | Amplitud del intervalo de confianza (90%) |
| `n_harmonics` | `4` | Armónicos para climatología Fourier multi-año |
| `clim_halflife_anios` | `2.0` | Semivida exponencial (años recientes pesan más) |

### Parámetros climatología SENAMHI

| Constante | Valor | Descripción |
|---|---|---|
| `SIGMA_CLIM` | `1.5` | Sigmas para banda de confianza climatológica |
| `Z_Q` | `0.6745` | Cuantil z para Q1/Q3 (50% de probabilidad) |
| `DELTA_Q` | `≈1.01°C` | Ancho de la banda Q1-Q3 en temperatura |
| `DISTRITO_BUSCAR` | `"GUADALUPE"` | Distrito SENAMHI fijo de referencia |
| `H_MAX_PENDIENTE` | `14 días` | Máximo horizonte para extrapolación de pendiente de bias |

---

## Módulo ENFEN — ajuste climático

El diccionario `AJUSTE_ENFEN` aplica una corrección adicional al bias del modelo para compensar eventos climáticos declarados oficialmente por el [ENFEN](https://enfen.imarpe.gob.pe/comunicados/).

```python
AJUSTE_ENFEN = {
    (5, 2026): {'Tmax': 1.9, 'Tmin': 1.0},   # El Niño Costero moderado
    (6, 2026): {'Tmax': 2.5, 'Tmin': 1.0},
    (7, 2026): {'Tmax': 2.5, 'Tmin': 1.0},
    (8, 2026): {'Tmax': 2.5, 'Tmin': 1.0},
}
```

- **Cuándo actualizar:** Cuando ENFEN publica un nuevo comunicado oficial en `enfen.imarpe.gob.pe`.
- Los valores de Tmax están calibrados con validación empírica walk-forward (MBE calculado en mayo-2026 sobre Arena Azul, Ayllu Allpa y Vivadis).
- La función `chequear_comunicado_enfen()` monitorea automáticamente la web de ENFEN cada 12 horas y genera una alerta visual si detecta un nuevo comunicado.

### Monitor ENFEN automático

`chequear_comunicado_enfen()`:
- Hace scraping de la página de comunicados ENFEN.
- Extrae el número y fecha del comunicado más reciente.
- Lo compara con el último visto (guardado en `assets/enfen_ultimo_visto.json`).
- Si hay uno nuevo, muestra un `st.warning` con el número y la fecha para que el usuario actualice `AJUSTE_ENFEN`.

---

## Funciones clave

### `cargar_meteoro_optimizado(file_bytes, filename, sheet, fundos_sel, min_reg)`
1. Lee el Excel con `leer_meteo_bytes_optimizado()`.
2. Parsea `Fecha-Hora` con `dayfirst=False` (formato M/DD/YYYY americano).
3. Filtra valores fuera de rango de los sensores.
4. Incluye siempre el último día por fundo aunque tenga pocos registros.
5. Agrega a diario: `Tmax=max(TempAlta-C)`, `Tmin=min(TempBaja-C)`, `ET=sum(ET-mm)`.
6. Aplica suavizado rolling de 3 días a `Tmax_smooth` y `Tmin_smooth`.

### `calcular_climatologia_armonica(df_clim, n_harmonics, halflife_anios)`
Ajusta una regresión Fourier (serie de senos/cosenos) a toda la serie histórica con pesos exponenciales por antigüedad. Los años más recientes pesan más (half-life configurable). Devuelve los coeficientes del modelo armónico.

### `predecir_climatologia_armonica(fechas, coef, n_harmonics)`
Aplica los coeficientes Fourier a un rango de fechas y devuelve la temperatura climatológica esperada para cada día del año.

### `spline_diario_cached(fecha_inicio_str, fecha_fin_str, valores_tuple)`
Genera una interpolación spline cúbica a partir de los 12 valores mensuales de climatología SENAMHI. El resultado es una serie diaria suave. Cachado con `@st.cache_data` usando los valores como clave determinística.

### `cargar_normales(file_bytes, hoja)`
Lee el Excel SENAMHI de normales 1991-2020, busca las filas del distrito GUADALUPE y devuelve los 12 valores mensuales + la banda Q1-Q3 (±DELTA_Q).

### `cargar_catalogo_normales(file_bytes, hoja)`
Construye un catálogo jerárquico `{sector: {departamento: [distritos]}}` excluyendo GUADALUPE (que es el fijo). Permite al usuario seleccionar un segundo distrito de referencia.

### `load_kmz_bytes(kmz_bytes)`
Parsea el KMZ de módulos AQ1/AQ2 y devuelve polígonos con `{name, coords, mod_n, fundo_aq}`.

### `disolver_modulos(kmz_polygons)`
Fusiona todos los polígonos de cada `(fundo_aq, mod_n)` en un único contorno usando `shapely.ops.unary_union`. Devuelve los módulos disueltos por fundo para usarlos en los mapas Folium.

### `asignar_fundo(fundo_aq, mod_n)`
Mapea el código AQ y número de módulo al nombre del fundo agrícola:

| fundo_aq | mod_n | Fundo |
|---|---|---|
| AQ1 | cualquier | Arena Azul |
| AQ2 | 1 | Vivadis |
| AQ2 | 2 | Santa Teresa |
| AQ2 | 3 | Ayllu Allpa |

---

## Modelo Prophet — pipeline completo

```
Serie de temperatura diaria (fundo, variable)
        │
        ▼
1. Clip ±3σ — eliminar outliers extremos
        │
        ▼
2. calcular_climatologia_armonica() — regresión Fourier multi-año (4 armónicos)
        │
        ▼
3. Anomalía = temperatura - climatología diaria
        │
        ▼
4. Peso 2.5× para datos de los últimos 6 meses (pesos recientes)
        │
        ▼
5. Validación cruzada walk-forward (N últimos pasos)
   └── MAE por horizonte h (1..30 días)
        │
        ▼
6. Prophet final entrenado en toda la anomalía:
   - Estacionalidad anual (8 armónicos)
   - Estacionalidad mensual adicional (3 armónicos, periodo 30.5d)
   - Sin estacionalidad semanal ni diaria
        │
        ▼
7. Predicción → añadir climatología → temperatura absoluta
        │
        ▼
8. Corrección de bias:
   ├── a) Tendencia lineal de residuos recientes (ventana 42 días, cap 14d)
   ├── b) Bias estacional multi-año (mismo mes, años anteriores, peso exp decay 1.5 años)
   └── c) Ajuste ENFEN (si el mes/año está en AJUSTE_ENFEN)
        │
        ▼
9. Clip al rango histórico ±5°C
```

---

## Validación walk-forward y métricas

La sección de validación **no usa** el modelo cacheado; entrena un modelo nuevo cortando los datos antes de la ventana de evaluación (honesto, sin data leakage).

### Ventanas disponibles

| Ventana | Descripción |
|---|---|
| Último mes | Entrena hasta fin del mes anterior al actual. Predice ese mes completo. |
| Últimos 3 meses | Entrena hasta 3 meses atrás. Predice los 3 meses siguientes. |

### Métricas mostradas

| Métrica | Descripción |
|---|---|
| **MAE** | Error absoluto medio diario en °C |
| **RMSE** | Raíz del error cuadrático medio. Si ratio RMSE/MAE > 1.3: hay días problemáticos |
| **BIAS (MBE)** | Sesgo sistemático. `+ subestima` / `- sobreestima` |
| **Precisión ±1.5°C** | % de días con error dentro del umbral operacional ±1.5°C |

### Diagnóstico adicional (captions)

- **Techo teórico de precisión:** límite alcanzable si el bias fuera 0 (función CDF de la normal con σ=desviación residual). La brecha entre techo y precisión actual indica cuánto queda por ganar con calibración.
- **Componentes del bias:** muestra por separado `a_int`, `pendiente×h`, `bias_estacional` y `ENFEN` para facilitar el diagnóstico.
- **Diagnóstico double-counting:** MBE a cada paso de la corrección: sin corrección → +tendencia → +estacional → +ENFEN.

### Grid search de hiperparámetros

Herramienta integrada (expander "Calibración de hiperparámetros"):
- **Tipo 1 — Hiperparámetros Prophet:** grid sobre `changepoint_prior_scale` y `seasonality_prior_scale`.
- **Tipo 2 — Climatología:** grid sobre `n_harmonics` y `clim_halflife_anios`.
- Evalúa MAE, Precisión, MBE y Std para cada combo en la ventana seleccionada.

---

## Visualizaciones

### Gráficos de temperatura (Plotly)

Cada fundo tiene un subplot con:
- **Banda climatológica SENAMHI** — relleno Q1-Q3 semitransparente + línea media.
- **Segunda climatología dinámica** — distrito opcional seleccionado por el usuario (naranja punteado).
- **Datos reales** — línea negra con suavizado rolling 3 días + anotación del último valor.
- **Predicción** — línea de color (rojo=Tmax, azul=Tmin) con banda IC 90%.
- **Línea vertical** en el último dato real (corte real/predicción).
- **Conector** entre último real y primer predicho.

Tipos de visualización disponibles:
- **Línea:** curvas estándar.
- **Candlestick / Range Plot:** barras de error mostrando rango Tmax-Tmin real + predicción.

### Gráficos de ET (Plotly)

- Climatología ET histórica (spline de promedios mensuales) + banda ±std.
- Serie real suavizada + cuartiles Q1/Q3 horizontales.
- Predicción Prophet de ET (con patrón cíclico de los últimos 30 días aplicado).

### Gráficos de validación

- **Gráfico de error diario:** barras verdes (subestimación) y rojas (sobreestimación), con banda ±1.5°C.
- **Box plot por fundo:** distribución de errores con puntos individuales.
- **Gráfico apilado:** % de días en cada rango de error (< -1.5°C / -1.5 a 0 / 0 a 1.5 / > 1.5°C).

---

## Mapas Folium

### `generar_mapa_distritos()` — Mapa de riesgo térmico

- Base satellite ESRI + capa de etiquetas.
- Distritos SENAMHI coloreados según temperatura climatológica mensual (colormap azul-rojo).
- Distrito fijo (GUADALUPE) y distrito dinámico resaltados con relleno y marcador.
- Módulos KMZ disueltos por fundo, **coloreados por nivel de riesgo térmico**:
  - El riesgo se calcula comparando el último valor real del fundo contra sus propios percentiles P50/P75/P95.

| Nivel | Umbral | Color |
|---|---|---|
| Bajo | ≤ P50 propio | Verde |
| Medio | ≤ P75 propio | Amarillo |
| Alto | ≤ P95 propio | Naranja |
| Muy alto | > P95 propio | Rojo |

### `generar_mapa_et()` — Mapa de riesgo ET

- Polígonos de módulos KMZ coloreados según nivel de riesgo de evapotranspiración (escala de azules).
- Leyenda discreta con los fundos y sus valores ET semanal.
- Riesgo basado en percentiles Q1/Q2/Q3/P95 del propio historial ET del fundo.

---

## Exportación Excel

Botón **"Exportar Excel"** genera un `.xlsx` con:

### Hoja `Por_Fechas`
- Columnas: Empresa, Fundo, Fecha, `max` (real), `min` (real), `max pred`, `min pred`, status.
- Filas amarillas = datos reales; filas naranja = datos predichos.

### Hoja `Por_Horas`
- Predicción interpolada hora a hora mediante `CubicSpline` sobre las predicciones diarias.
- Simula el ciclo diario: Tmax alcanza su pico a media tarde, Tmin en la madrugada, con amplitud proporcional al IC 95%.
- Columnas: Empresa, Fundo, Fecha, Hora (HH:00), `max pred`, `min pred`.

---

## Instalación y dependencias

### Python >= 3.10 requerido

```bash
pip install streamlit pandas numpy scipy matplotlib plotly
pip install prophet
pip install openpyxl lxml shapely folium branca
pip install msal pyodbc requests
```

### Instalación de Prophet (requiere Stan)

```bash
# Opción recomendada (más rápida):
conda install -c conda-forge prophet

# Alternativa pip:
pip install prophet
# Si falla, instala antes: pip install pystan==2.19.1.1
```

### Tabla completa de dependencias

| Paquete | Uso |
|---|---|
| `streamlit` | Framework de la app web |
| `pandas` | Manipulación de datos |
| `numpy` | Cálculos numéricos, percentiles, regresión |
| `scipy` | `CubicSpline`, `norm.cdf`, `gaussian_filter1d` |
| `matplotlib` | Solo para compatibilidad (no se renderiza) |
| `plotly` | Gráficos interactivos (subplots, scatter, box) |
| `prophet` | Modelo de predicción de series temporales |
| `openpyxl` | Lectura y escritura de Excel con estilos |
| `lxml` | Parseo de KML dentro del KMZ |
| `shapely` | Unión de polígonos, geometría vectorial |
| `folium` | Mapas interactivos basados en Leaflet |
| `branca` | Elementos HTML personalizados en Folium (leyendas) |
| `msal` | Autenticación Azure AD (opcional, si se usa SQL Server) |
| `pyodbc` | Conexión ODBC (opcional) |
| `requests` | Scraping del monitor ENFEN |

---

## Configuración de secretos

El archivo `.streamlit/secrets.toml` debe contener (solo si se usa la conexión SQL Server):

```toml
GITHUB_TOKEN_KMZ = "ghp_xxxxxxxxxxxx"   # Token para descargar el KMZ privado

[database]
server   = "tu-servidor.database.windows.net"
database = "nombre-base-de-datos"
username = "usuario@dominio.com"
```

> Si no se usa SQL Server, los secretos son opcionales. La app funciona completamente con el Excel subido manualmente y el KMZ subido por el usuario.

---

## Estructura de archivos requerida

```
PREDICCION_TEMPERATURA/
├── temperatura_v2.py              ← Este script
└── assets/
    ├── NORMALES_1991_2020.xlsx    ← Normales climatológicas SENAMHI (OBLIGATORIO)
    │   ├── Hoja TMAX              ← Columnas: Departamento, Provincia, Distrito, Ene..Dic
    │   └── Hoja TMIN
    └── enfen_ultimo_visto.json    ← Creado automáticamente por el monitor ENFEN
```

### Formato del Excel meteorológico (entrada del usuario)

Hoja: `Prize_Climatology`

| Columna | Formato | Descripción |
|---|---|---|
| `Fecha-Hora` | M/DD/YYYY HH:MM | Timestamp cada 15 minutos (formato americano) |
| `Empresa` | texto | Nombre de la empresa |
| `Fundo` | texto | Nombre del fundo (debe coincidir con los datos KMZ) |
| `TempAlta-C` | numérico | Temperatura máxima del intervalo (°C) |
| `TempBaja-C` | numérico | Temperatura mínima del intervalo (°C) |
| `Temp-C` | numérico | Temperatura media del intervalo (°C) |
| `ET-mm` | numérico | Evapotranspiración del intervalo (mm) |

---

## Casuísticas y comportamiento esperado

### 1. Formato de fecha incorrecto en el Excel

- El parser usa `dayfirst=False` para interpretar `6/01/2026` como 1 de junio (no 6 de enero).
- Si la columna `Fecha-Hora` tiene un formato diferente, `pd.to_datetime` producirá `NaT` y esas filas se descartan silenciosamente.
- Si muchas filas se descartan, el sidebar mostrará un conteo bajo de registros.

### 2. Día con pocos registros (sensor apagado o caído)

- El filtro `N_registros >= min_reg` (80 de 96 posibles) descarta días incompletos.
- **Excepción:** el último día disponible por fundo **siempre se incluye**, aunque tenga menos de 80 registros. Esto evita perder el dato más reciente cuando el usuario exporta a mediodía.

### 3. No hay columna ET-mm

- Si la columna no existe en el Excel, se crea con valor 0.0.
- El modelo ET no se entrena si hay menos de 30 días con ET > 0.

### 4. Menos de 30 días de historial

- Prophet requiere mínimo 30 datos. Si hay menos, el fundo/variable se salta sin error.
- Para ET se exige además que los datos sean > 0.

### 5. Nuevo comunicado ENFEN detectado

- El monitor muestra un `st.warning` naranja con el número y fecha del comunicado.
- El usuario debe actualizar manualmente el diccionario `AJUSTE_ENFEN` en el código con los nuevos valores de ajuste.
- El archivo `assets/enfen_ultimo_visto.json` guarda el último comunicado visto para no repetir la alerta.

### 6. KMZ no disponible (GitHub caído o sin token)

- La descarga intenta primero `raw.githubusercontent.com` (sin autenticación) y luego la API de GitHub (con token).
- Si ambas fallan, el usuario puede subir el KMZ manualmente desde la interfaz (file uploader en la sidebar).
- Sin KMZ: los mapas Folium se generan sin polígonos de módulos (solo distritos SENAMHI).

### 7. Segunda climatología SENAMHI (distrito dinámico)

- El catálogo se construye desde el Excel de normales excluyendo GUADALUPE.
- El usuario puede seleccionar sector → departamento → distrito en cascada.
- Si selecciona `— Ninguna —`, solo se muestra la climatología de Guadalupe.
- La segunda climatología aparece en naranja punteado en los gráficos y en el mapa de distritos.

### 8. Validación walk-forward con bias double-counting

- El diagnóstico de MBE muestra cada componente por separado para verificar que no se está sumando el mismo efecto dos veces.
- Si `MBE final ≈ 0`, el sistema está bien calibrado para esa ventana y esa variable.
- Si el bias sigue siendo alto, la sección "Calibración de hiperparámetros" permite hacer grid search para mejorar.

### 9. Predicción ET con patrón cíclico

- La predicción ET no es una línea plana: hereda el patrón de variabilidad de los últimos 30 días reales.
- El factor de escala del patrón está capado a 1.5 para evitar amplificaciones excesivas.
- El IC de ET se construye sumando ±0.8×std_histórica a la predicción (no usa el IC de Prophet directamente).

### 10. Exportación Excel horaria

- La predicción horaria simula el ciclo día/noche aplicando un seno desplazado sobre el IC diario.
- Para Tmax: pico a las ~14:00 (offset de 5h sobre el ciclo de 24h).
- Para Tmin: valle a las ~05:00.
- La amplitud del ciclo es `(IC_upper - IC_lower) / 4`, lo que produce oscilaciones moderadas.

### 11. Cambio de año en la predicción

- La app muestra datos desde el 1 de enero del año anterior para contexto histórico.
- La predicción siempre alcanza hasta el fin del mes actual del sensor.
- Si `dias_pred` del usuario es menor que los días restantes del mes, se usa el fin de mes como mínimo.

### 12. Bias estacional "sin años anteriores"

- Si hay menos de 15 días del mismo mes en años anteriores al de predicción, `bias_estacional = 0.0`.
- Esto ocurre en años iniciales de operación del sensor o cuando solo hay datos de un año.

---

## Ejecutar la aplicación

```bash
cd PREDICCION_TEMPERATURA
streamlit run temperatura_v2.py
```

La aplicación queda disponible en `http://localhost:8501`.

### Pasos de uso típico

1. Subir el Excel meteorológico (`.xlsx` o `.csv`) desde el sidebar.
2. Seleccionar los fundos a analizar (multiselect).
3. Ajustar el mínimo de registros por día si los sensores tienen lagunas.
4. Presionar **"Entrenar Prophet"** para generar las predicciones.
5. Navegar entre las pestañas: **Tmax**, **Tmin**, **ET**, **Mapa**, **Validación**.
6. Exportar Excel con el botón **"Exportar Excel"**.