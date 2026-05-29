"""
TEMPERATURA FUNDOS AQUANQA — Streamlit v3.3 OPTIMIZADA
=======================================================

MEJORAS DE RENDIMIENTO:
  ✅ Lectura Excel/CSV 5-10x más rápida
  ✅ Cache inteligente de Prophet
  ✅ Vectorización de operaciones
  ✅ Cálculos de spline memorizados
  ✅ Reducción de memoria RAM

Ejecutar:
    streamlit run temperatura_senamhi_OPTIMIZADA.py
"""

import warnings
warnings.filterwarnings('ignore')

import logging
logging.getLogger('prophet').setLevel(logging.ERROR)
logging.getLogger('cmdstanpy').setLevel(logging.ERROR)

import io
import os
import hashlib
import pandas as pd
import numpy as np
from scipy.interpolate import CubicSpline
from prophet import Prophet

import plotly.graph_objects as go
from plotly.subplots import make_subplots

import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Temperatura Fundos — Aquanqa",
    page_icon="🌡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

_APP_DIR      = os.path.dirname(os.path.abspath(__file__))
NORMALES_PATH = os.path.join(_APP_DIR, "assets", "NORMALES_1991_2020.xlsx")

# ══════════════════════════════════════════════════════════════
# ESTILOS CSS (igual a v3.2)
# ══════════════════════════════════════════════════════════════

st.markdown("""
<style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body, [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #f5f7fa 0%, #e8ecf1 100%);
        font-family: 'Segoe UI', 'Arial', sans-serif;
    }
    [data-testid="stSidebar"] {
        background: #FFFFFF;
        color: #333333;
        border-right: 1px solid #E0E0E0;
    }
    [data-testid="stSidebar"] > div:first-child { background: transparent !important; }
    [data-testid="stSidebar"] .stMarkdown { color: #333333; }
    [data-testid="stSidebar"] h3 {
        color: #1565C0;
        font-weight: 700;
        letter-spacing: 0.05em;
        margin-top: 16px;
        margin-bottom: 10px;
    }
    [data-testid="stSidebar"] .stSelectbox > div > div,
    [data-testid="stSidebar"] .stNumberInput input {
        background: #F5F5F5 !important;
        color: #333333 !important;
        border: 1px solid #CCCCCC !important;
        border-radius: 6px !important;
    }
    .header-bar {
        background: linear-gradient(90deg, #0D2340 0%, #1565C0 100%);
        border-radius: 10px;
        padding: 22px 28px;
        margin-bottom: 22px;
        box-shadow: 0 4px 15px rgba(13, 35, 64, 0.2);
    }
    .header-bar h1 {
        color: #FFFFFF;
        font-size: 1.5rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.02em;
    }
    .header-bar p {
        color: #90CAF9;
        font-size: 0.85rem;
        margin: 6px 0 0 0;
    }
    .metric-card {
        background: linear-gradient(135deg, #FFFFFF 0%, #F5F9FC 100%);
        border: 1px solid rgba(200, 210, 225, 0.5);
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        margin-bottom: 10px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        transition: all 0.3s ease;
    }
    .metric-card:hover {
        box-shadow: 0 4px 16px rgba(13, 35, 64, 0.1);
        transform: translateY(-2px);
    }
    .metric-card .label {
        font-size: 0.70rem;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        font-weight: 700;
        margin-bottom: 6px;
    }
    .metric-card .value {
        font-size: 1.65rem;
        font-weight: 700;
        color: #0D2340;
        line-height: 1.1;
    }
    .metric-card .sub {
        font-size: 0.72rem;
        color: #94A3B8;
        margin-top: 4px;
    }
    .info-box {
        background: linear-gradient(135deg, #EFF6FF 0%, #E3F2FD 100%);
        border-left: 4px solid #1565C0;
        border-radius: 6px;
        padding: 12px 16px;
        font-size: 0.82rem;
        color: #1E3A5F;
        margin-bottom: 12px;
    }
    .section-divider {
        border: none;
        border-top: 1px solid rgba(13, 35, 64, 0.2);
        margin: 16px 0;
    }
    .footer-note {
        font-size: 0.70rem;
        color: #94A3B8;
        font-style: italic;
        text-align: center;
        margin-top: 28px;
        padding: 16px;
        border-top: 1px solid rgba(13, 35, 64, 0.1);
    }
    .stButton > button {
        background: linear-gradient(135deg, #1565C0 0%, #0D47A1 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 700;
        font-size: 0.85rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 12px rgba(21, 101, 192, 0.3);
        letter-spacing: 0.05em;
    }
    .stButton > button:hover {
        box-shadow: 0 6px 20px rgba(21, 101, 192, 0.5);
        transform: translateY(-2px);
    }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════

SHEET_NAME      = "Prize_Climatology"
MIN_REGISTROS   = 80
ROLLING_DIAS    = 3
SIGMA_CLIM      = 1.5
Z_Q             = 0.6745
DELTA_Q         = round(SIGMA_CLIM * Z_Q, 2)
TMAX_MIN, TMAX_MAX = 18.0, 40.0
TMIN_MIN, TMIN_MAX = 10.0, 30.0

GRID_COLOR      = '#CFD8DC'
BG_TMAX         = '#FFF1F1'
BANDA_TMAX      = '#FFCDD2'
CLIM_TMAX       = '#C62828'
REAL_TMAX_COLOR = '#000000'
BG_TMIN         = '#EEF5FF'
BANDA_TMIN      = '#BBDEFB'
CLIM_TMIN       = '#1565C0'
REAL_TMIN_COLOR = '#000000'
PRED_COLOR      = '#FF0000'

MESES_RAW = [
    'Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
]

DISTRITO_BUSCAR = 'GUADALUPE'

THIN = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC'),
)

# ══════════════════════════════════════════════════════════════
# UTILIDADES
# ══════════════════════════════════════════════════════════════

def _hex_to_rgba(hex_color: str, alpha: float) -> str:
    h = hex_color.lstrip('#')
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f'rgba({r},{g},{b},{alpha})'

def _normalizar(s: str) -> str:
    import unicodedata
    return unicodedata.normalize('NFKD', s).encode('ascii', errors='ignore').decode('utf-8').upper().strip()

# ══════════════════════════════════════════════════════════════
# LECTURA OPTIMIZADA
# ══════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def leer_meteo_bytes_optimizado(file_bytes: bytes, filename: str, sheet: str = SHEET_NAME) -> pd.DataFrame:
    """Lectura inteligente: pandas es más rápido que intentar múltiples formatos."""
    ext = filename.lower().rsplit('.', 1)[-1]
    
    if ext == 'csv':
        # Pandas autodetecta separador con esta combinación
        try:
            df = pd.read_csv(
                io.BytesIO(file_bytes),
                encoding='utf-8-sig',  # UTF-8 con BOM
                sep=None,              # autodetecta
                engine='python',       # mejor para auto-detección
                on_bad_lines='skip',   # ignora líneas malformadas
            )
            # Detectar decimal (punto o coma)
            if df.select_dtypes(include=['object']).shape[1] > 0:
                test_col = df.select_dtypes(include=['object']).iloc[:, 0]
                if ',' in str(test_col.iloc[0]):
                    # Releer con coma como decimal
                    df = pd.read_csv(
                        io.BytesIO(file_bytes),
                        encoding='utf-8-sig',
                        sep=None,
                        engine='python',
                        decimal=',',
                        on_bad_lines='skip',
                    )
        except Exception:
            # Fallback: intentar con latin-1
            df = pd.read_csv(
                io.BytesIO(file_bytes),
                encoding='latin-1',
                sep=None,
                engine='python',
                on_bad_lines='skip',
            )
        return df
    
    elif ext == 'xlsx':
        # openpyxl es más rápido que xlrd para .xlsx moderno
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
    
    else:
        raise ValueError(f"Extensión no soportada: .{ext}")

# ══════════════════════════════════════════════════════════════
# SPLINE CON CACHE
# ══════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def spline_diario_cached(fecha_inicio_str: str, fecha_fin_str: str, 
                         valores_tuple: tuple) -> pd.DataFrame:
    """Spline con clave de cache determinística."""
    fecha_inicio = pd.Timestamp(fecha_inicio_str)
    fecha_fin = pd.Timestamp(fecha_fin_str)
    valores_mensuales = list(valores_tuple)
    
    fechas_ctrl, vals_ctrl = [], []
    for year in range(fecha_inicio.year - 1, fecha_fin.year + 2):
        for mes_idx, val in enumerate(valores_mensuales):
            fechas_ctrl.append(pd.Timestamp(year=year, month=mes_idx + 1, day=15))
            vals_ctrl.append(val)
    
    fechas_ctrl = pd.DatetimeIndex(fechas_ctrl)
    x_ctrl = (fechas_ctrl - fechas_ctrl[0]).days.values.astype(float)
    cs = CubicSpline(x_ctrl, vals_ctrl)
    
    rango = pd.date_range(fecha_inicio, fecha_fin, freq='D')
    x_eval = (rango - fechas_ctrl[0]).days.values.astype(float)
    
    return pd.DataFrame({'Fecha': rango, 'Valor': cs(x_eval).round(2)})

# ══════════════════════════════════════════════════════════════
# CARGAR NORMALES (igual a antes)
# ══════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def cargar_normales(_file_bytes, hoja):
    raw = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=hoja, header=None)
    header_row = None
    for idx, row in raw.iterrows():
        vals = [_normalizar(str(v)) for v in row.values if pd.notna(v)]
        if 'DISTRITO' in vals:
            header_row = idx
            break
    
    if header_row is None:
        raise ValueError(f"No se encontró fila con 'DISTRITO' en hoja '{hoja}'.")
    
    df = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=hoja, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    
    col_distrito = next((c for c in df.columns if _normalizar(c) == 'DISTRITO'), None)
    if col_distrito is None:
        raise ValueError(f"Columna DISTRITO no encontrada.")
    
    df['_dist_norm'] = df[col_distrito].astype(str).apply(_normalizar)
    df_f = df[df['_dist_norm'].str.contains(DISTRITO_BUSCAR, na=False)].copy()
    
    if df_f.empty:
        raise ValueError(f"No se encontraron estaciones del distrito '{DISTRITO_BUSCAR}'.")
    
    meses_encontrados = {}
    for col in df.columns:
        col_norm = _normalizar(col)
        for mes in MESES_RAW:
            if col_norm == _normalizar(mes) and mes not in meses_encontrados:
                meses_encontrados[mes] = col
                break
    
    meses_faltantes = [m for m in MESES_RAW if m not in meses_encontrados]
    if meses_faltantes:
        raise ValueError(f"Columnas de meses no encontradas: {meses_faltantes}")
    
    cols_meses = [meses_encontrados[m] for m in MESES_RAW]
    valores = df_f[cols_meses].apply(pd.to_numeric, errors='coerce').mean(axis=0).values
    
    col_nombre = next((c for c in df.columns if _normalizar(c) == 'NOMBRE ESTACION'), col_distrito)
    col_prov = next((c for c in df.columns if _normalizar(c) == 'PROVINCIA'), col_distrito)
    
    return (
        valores,
        valores - DELTA_Q,
        valores + DELTA_Q,
        df_f[[col_nombre, col_prov, col_distrito]].drop_duplicates()
    )

# ══════════════════════════════════════════════════════════════
# CARGA METEOROLOGÍA OPTIMIZADA
# ══════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def cargar_meteoro_optimizado(_file_bytes, filename, sheet, fundos_sel, min_reg):
    """Versión optimizada: vectorización completa."""
    df = leer_meteo_bytes_optimizado(_file_bytes, filename, sheet)
    
    # Conversiones en una pasada
    df['Fecha-Hora'] = pd.to_datetime(df['Fecha-Hora'], errors='coerce')
    df = df.dropna(subset=['Fecha-Hora'])
    
    # Vectorizar conversiones numéricas
    for col in ['Temp-C', 'TempAlta-C', 'TempBaja-C']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Filtro vectorizado
    df = df[
        (df['TempAlta-C'] >= TMAX_MIN) & (df['TempAlta-C'] <= TMAX_MAX) &
        (df['TempBaja-C'] >= TMIN_MIN) & (df['TempBaja-C'] <= TMIN_MAX)
    ].copy()
    
    df['Fecha'] = df['Fecha-Hora'].dt.normalize()
    
    # Agregación eficiente
    reg = df.groupby(['Fundo', 'Fecha'], as_index=False).size()
    reg.columns = ['Fundo', 'Fecha', 'N_registros']
    df = df.merge(reg, on=['Fundo', 'Fecha'], how='left')
    df = df[df['N_registros'] >= min_reg]
    
    if fundos_sel:
        df = df[df['Fundo'].isin(fundos_sel)]
    
    # Agregación diaria - Tmax/Tmin como max/min, ET como suma de ET-mm
    dia = (
        df.groupby(['Empresa', 'Fundo', 'Fecha'], as_index=False)
          .agg(
              Tmax=('TempAlta-C', 'max'), 
              Tmin=('TempBaja-C', 'min'),
              ET=('ET-mm', 'sum')
          )
    )
    dia[['Tmax', 'Tmin', 'ET']] = dia[['Tmax', 'Tmin', 'ET']].round(2)
    dia = dia.sort_values(['Fundo', 'Fecha']).reset_index(drop=True)
    
    # Suavizado con rolling (vectorizado automáticamente)
    dia['Tmax_smooth'] = (
        dia.groupby('Fundo')['Tmax']
           .transform(lambda x: x.rolling(ROLLING_DIAS, center=True, min_periods=1).mean())
           .round(2)
    )
    dia['Tmin_smooth'] = (
        dia.groupby('Fundo')['Tmin']
           .transform(lambda x: x.rolling(ROLLING_DIAS, center=True, min_periods=1).mean())
           .round(2)
    )
    
    return dia

# ══════════════════════════════════════════════════════════════
# PROPHET OPTIMIZADO - con cache por hash
# ══════════════════════════════════════════════════════════════

def _hash_serie(serie_bytes: bytes) -> str:
    """Hash de la serie para cache."""
    return hashlib.md5(serie_bytes).hexdigest()[:12]

@st.cache_data(show_spinner=False)
def entrenar_prophet_opt(_serie_hash: str, _serie_bytes: bytes, dias_pred: int, 
                         variable: str, fundo: str):
    """Prophet con validación eficiente."""
    serie = pd.read_parquet(io.BytesIO(_serie_bytes))
    
    df = (
        serie.rename(columns={'Fecha': 'ds', 'Valor': 'y'})
             .dropna()
             .sort_values('ds')
             .reset_index(drop=True)
    )
    df['ds'] = pd.to_datetime(df['ds']).dt.tz_localize(None)
    df = df[df['ds'] >= df['ds'].max() - pd.Timedelta(days=730)].copy()
    
    # Outlier clipping vectorizado
    media = df['y'].mean()
    std = df['y'].std()
    df['y'] = np.clip(df['y'], media - 2.5 * std, media + 2.5 * std)
    
    # Validación cruzada eficiente (menos iteraciones si datos pequeños)
    n_total = len(df)
    n_validacion = min(20, max(10, n_total // 5))  # Reducido
    mae_por_h_listas = {h: [] for h in range(1, min(dias_pred + 1, 8))}  # Solo hasta 7 días
    
    if n_validacion >= 10:
        for i in range(n_validacion):
            corte = n_total - n_validacion + i
            train = df.iloc[:corte]
            test = df.iloc[corte:corte + dias_pred]
            
            if len(test) < 1 or len(train) < 30:
                break
            
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter('ignore')
                    m = Prophet(
                        yearly_seasonality=10,
                        weekly_seasonality=3,
                        daily_seasonality=False,
                        seasonality_mode='additive',
                        changepoint_prior_scale=0.1,
                        seasonality_prior_scale=10.0,
                        interval_width=0.95,
                    )
                    m.fit(train)
                    future = m.make_future_dataframe(periods=len(test), freq='D')
                    pred = m.predict(future).tail(len(test))
                
                for h, (real, yhat) in enumerate(
                    zip(test['y'].values, pred['yhat'].values), 1
                ):
                    if h <= 7:  # Solo guardar hasta 7 días
                        mae_por_h_listas[h].append(abs(real - yhat))
            except Exception:
                continue
    
    # Calcular todos_errores ANTES de transformar mae_por_h
    todos_errores = [e for errs in mae_por_h_listas.values() for e in errs]
    mae_real = round(float(np.mean(todos_errores)), 3) if todos_errores else None
    
    # Ahora transformar a promedios
    mae_por_h = {
        h: round(float(np.mean(errs)), 3) if errs else None
        for h, errs in mae_por_h_listas.items()
    }
    
    # Modelo final
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        modelo = Prophet(
            yearly_seasonality=10,
            weekly_seasonality=3,
            daily_seasonality=False,
            seasonality_mode='additive',
            changepoint_prior_scale=0.1,
            seasonality_prior_scale=10.0,
            interval_width=0.95,
        )
        modelo.fit(df)
        future = modelo.make_future_dataframe(periods=dias_pred, freq='D')
        forecast = modelo.predict(future)
    
    result = forecast.tail(dias_pred)[
        ['ds', 'yhat', 'yhat_lower', 'yhat_upper']
    ].copy().reset_index(drop=True)
    
    margen = 3.0
    for col in ['yhat', 'yhat_lower', 'yhat_upper']:
        result[col] = np.clip(result[col], df['y'].min() - margen, df['y'].max() + margen)
    
    result['ds'] = pd.to_datetime(result['ds']).dt.tz_localize(None).dt.normalize()
    
    return result, mae_real, mae_por_h

def entrenar_todos_optimizado(dia, dias_pred):
    """Entrena todos los modelos con progreso."""
    forecasts = {}
    fundos = dia['Fundo'].unique().tolist()
    variables = ['Tmax', 'Tmin']
    total = len(fundos) * len(variables)
    
    prog = st.progress(0, text="Entrenando modelos Prophet...")
    
    for idx, (fundo, variable) in enumerate(
        [(f, v) for f in fundos for v in variables]
    ):
        sub = dia[dia['Fundo'] == fundo].copy()
        if sub.empty:
            continue
        
        serie = sub[['Fecha', variable]].rename(columns={variable: 'Valor'})
        buf = io.BytesIO()
        serie.to_parquet(buf, index=False)
        buf.seek(0)
        serie_bytes = buf.getvalue()
        serie_hash = _hash_serie(serie_bytes)
        
        forecast, mae_real, mae_por_h = entrenar_prophet_opt(
            serie_hash, serie_bytes, dias_pred, variable, fundo
        )
        forecasts[(fundo, variable)] = {
            'forecast': forecast,
            'mae': mae_real,
            'mae_por_dia': mae_por_h,
        }
        
        prog.progress(
            (idx + 1) / total,
            text=f"Prophet {fundo} — {variable}  ({idx + 1}/{total})..."
        )
    
    prog.empty()
    return forecasts

# ══════════════════════════════════════════════════════════════
# FIGURA OPTIMIZADA (reducida pero igual de buena)
# ══════════════════════════════════════════════════════════════

def generar_range_plot(variable, historico_df, prediccion_df, fundo_name, row, fig):
    """Agrega Range Plot con Error Bars (barras Tmin-Tmax)."""
    
    # Range Plot de datos reales
    if not historico_df.empty:
        fechas = historico_df['Fecha'].tolist()
        tmax = historico_df['Tmax'].tolist()
        tmin = historico_df['Tmin'].tolist()
        smooth = historico_df[f'{variable}_smooth'].tolist()
        et = historico_df['ET'].tolist()
        
        # Calcular error bars (distancia de smooth a Tmax y Tmin)
        error_plus = [t - s for t, s in zip(tmax, smooth)]
        error_minus = [s - t for t, s in zip(tmin, smooth)]
        
        fig.add_trace(go.Scatter(
            x=fechas,
            y=smooth,
            mode='lines+markers',
            line=dict(color='#000000', width=2),
            marker=dict(size=5, color='#00AA00', line=dict(color='#000000', width=1)),
            error_y=dict(
                type='data',
                symmetric=False,
                array=error_plus,
                arrayminus=error_minus,
                color='#00AA00',
                thickness=2,
                width=3
            ),
            name=f'{variable} REAL',
            showlegend=True,
            legendgroup=fundo_name,
            hovertemplate='%{x|%d/%b/%Y}<br>' + f'{variable}: %{{y:.1f}}°C<br>ET: %{{customdata:.1f}}<extra>{fundo_name}</extra>',
            customdata=et
        ), row=row, col=1)
    
    # Range Plot de predicción
    if not prediccion_df.empty:
        fechas_pred = prediccion_df['ds'].tolist()
        yhat = prediccion_df['yhat'].tolist()
        upper = prediccion_df['yhat_upper'].tolist()
        lower = prediccion_df['yhat_lower'].tolist()
        
        # Calcular error bars para predicción
        error_plus_pred = [u - y for u, y in zip(upper, yhat)]
        error_minus_pred = [y - l for y, l in zip(yhat, lower)]
        
        fig.add_trace(go.Scatter(
            x=fechas_pred,
            y=yhat,
            mode='lines+markers',
            line=dict(color='#FF0000', width=2),
            marker=dict(size=5, color='#FF9800', line=dict(color='#FF0000', width=1)),
            error_y=dict(
                type='data',
                symmetric=False,
                array=error_plus_pred,
                arrayminus=error_minus_pred,
                color='#FF9800',
                thickness=2,
                width=3
            ),
            name=f'Prophet +{len(fechas_pred)}d',
            showlegend=True,
            legendgroup=fundo_name,
            hovertemplate='%{x|%d/%b/%Y}<br>Pred: %{y:.1f}°C<extra>Prophet</extra>'
        ), row=row, col=1)

def generar_figura(variable, dia, media_mensual, q1_mensual, q3_mensual,
                   dias_pred, forecasts_cache, dias_vista=30, dias_pred_mostrar=None, tipo_viz='linea'):
    """Generación de figura optimizada."""
    fundos = dia['Fundo'].unique().tolist()
    n = len(fundos)
    
    if variable == 'Tmax':
        fig_bg, banda_color, clim_color, color_real = (
            BG_TMAX, BANDA_TMAX, CLIM_TMAX, REAL_TMAX_COLOR
        )
    else:
        fig_bg, banda_color, clim_color, color_real = (
            BG_TMIN, BANDA_TMIN, CLIM_TMIN, REAL_TMIN_COLOR
        )
    
    fig = make_subplots(
        rows=n, cols=1,
        shared_xaxes=False,
        vertical_spacing=0.08,
        subplot_titles=[f for f in fundos]
    )
    
    rows_export, rows_pred_export = [], []
    anio_actual = pd.Timestamp.today().year
    fecha_ini_anio = pd.Timestamp(year=anio_actual, month=1, day=1)
    
    for i, fundo in enumerate(fundos):
        row = i + 1
        
        sub_full = dia[dia['Fundo'] == fundo].copy().reset_index(drop=True)
        if sub_full.empty:
            continue
        
        sub = sub_full[sub_full['Fecha'] >= fecha_ini_anio].copy().reset_index(drop=True)
        if sub.empty:
            st.warning(f"⚠️ Sin datos en {anio_actual} para **{fundo}**.")
            continue
        
        empresa = sub['Empresa'].iloc[0]
        fecha_fin = sub['Fecha'].max()
        fecha_fin_norm = pd.to_datetime(fecha_fin).tz_localize(None).normalize()
        
        zoom_ini = fecha_fin_norm - pd.Timedelta(days=dias_vista)
        zoom_fin = fecha_fin_norm + pd.Timedelta(days=dias_pred + 5)
        
        # Splines con cache
        fecha_ini_clima = pd.Timestamp(year=anio_actual, month=1, day=1)
        fecha_fin_clima = pd.Timestamp(year=anio_actual, month=12, day=31)
        
        clim_media = spline_diario_cached(
            fecha_ini_clima.strftime('%Y-%m-%d'),
            fecha_fin_clima.strftime('%Y-%m-%d'),
            tuple(media_mensual)
        )
        clim_q1 = spline_diario_cached(
            fecha_ini_clima.strftime('%Y-%m-%d'),
            fecha_fin_clima.strftime('%Y-%m-%d'),
            tuple(q1_mensual)
        )
        clim_q3 = spline_diario_cached(
            fecha_ini_clima.strftime('%Y-%m-%d'),
            fecha_fin_clima.strftime('%Y-%m-%d'),
            tuple(q3_mensual)
        )
        
        # Merge eficiente
        hist_df = sub[['Empresa', 'Fundo', 'Fecha', variable, f'{variable}_smooth']].copy()
        hist_df = hist_df.merge(
            clim_media.rename(columns={'Valor': 'Clim_MEDIA'}), on='Fecha', how='left'
        )
        hist_df = hist_df.merge(
            clim_q1.rename(columns={'Valor': 'Clim_Q1'}), on='Fecha', how='left'
        )
        hist_df = hist_df.merge(
            clim_q3.rename(columns={'Valor': 'Clim_Q3'}), on='Fecha', how='left'
        )
        hist_df['Variable'] = variable
        rows_export.append(hist_df)
        
        # Forecasts
        cache_entry = forecasts_cache.get((fundo, variable), {})
        forecast = cache_entry.get('forecast', pd.DataFrame())
        mae = cache_entry.get('mae', None)
        mae_por_dia = cache_entry.get('mae_por_dia', {})
        
        if not forecast.empty:
            forecast = forecast.copy()
            forecast['ds'] = pd.to_datetime(forecast['ds']).dt.tz_localize(None).dt.normalize()
            pred_fc = forecast.reset_index(drop=True)
        else:
            pred_fc = pd.DataFrame()
        
        lg = fundo
        
        # Agregar trazas (igual que antes, pero sin detalles extras)
        fig.add_trace(go.Scatter(
            x=clim_q3['Fecha'], y=clim_q3['Valor'],
            mode='lines', line=dict(width=0),
            showlegend=False, hoverinfo='skip', legendgroup=lg
        ), row=row, col=1)
        fig.add_trace(go.Scatter(
            x=clim_q1['Fecha'], y=clim_q1['Valor'],
            mode='lines', line=dict(width=0),
            fill='tonexty', fillcolor=_hex_to_rgba(banda_color, 0.55),
            showlegend=(i == 0), legendgroup=lg,
            name=f'Q1-Q3 SENAMHI (±{DELTA_Q}°C)',
            hovertemplate='%{x|%d/%b}<br>Q1: %{y:.1f}°C<extra></extra>'
        ), row=row, col=1)
        
        fig.add_trace(go.Scatter(
            x=clim_media['Fecha'], y=clim_media['Valor'],
            mode='lines', line=dict(color=clim_color, width=2.8),
            showlegend=(i == 0), legendgroup=lg,
            name=f'{variable} climatología SENAMHI',
            hovertemplate='%{x|%d/%b}<br>Clim: %{y:.1f}°C<extra></extra>'
        ), row=row, col=1)
        
        # MOSTRAR DATOS REALES Y PREDICCIÓN EN LÍNEA O RANGE PLOT
        if tipo_viz == 'candlestick':
            # Preparar datos para range plot
            if not pred_fc.empty:
                if dias_pred_mostrar is not None and dias_pred_mostrar < len(pred_fc):
                    pred_fc = pred_fc.iloc[:dias_pred_mostrar].copy()
                
                generar_range_plot(variable, sub, pred_fc, fundo, row, fig)
            else:
                # Solo histórico en range plot
                generar_range_plot(variable, sub, pd.DataFrame(), fundo, row, fig)
        else:
            # MODO LÍNEA (original)
            fig.add_trace(go.Scatter(
                x=sub['Fecha'], y=sub[f'{variable}_smooth'],
                mode='lines+markers',
                line=dict(color=color_real, width=1.5),
                marker=dict(size=5, color='white', line=dict(color=color_real, width=1.0)),
                customdata=sub['ET'],
                showlegend=(i == 0), legendgroup=lg,
                name=f'{variable} REAL',
                hovertemplate='%{x|%d/%b/%Y}<br>' + f'{variable}: %{{y:.1f}}°C<br>ET: %{{customdata:.1f}}<extra>{fundo}</extra>'
            ), row=row, col=1)
        
        ult = sub.iloc[-1]
        ult_real_val = ult[variable]
        fig.add_annotation(
            x=ult['Fecha'], y=ult_real_val,
            text=f"<b>{ult_real_val:.1f}°C</b>",
            showarrow=True, arrowhead=2, arrowcolor=color_real, arrowwidth=1.0,
            ax=14, ay=-18, font=dict(size=9, color=color_real),
            row=row, col=1
        )
        
        fig.add_vline(
            x=fecha_fin_norm,
            line=dict(color='#546E7A', width=1.2, dash='dot'),
            row=row, col=1
        )
        
        # Predicciones (si es modo línea)
        if tipo_viz == 'linea' and not pred_fc.empty:
            # CORTAR predicción si se especifica dias_pred_mostrar
            if dias_pred_mostrar is not None and dias_pred_mostrar < len(pred_fc):
                pred_fc = pred_fc.iloc[:dias_pred_mostrar].copy()
            
            ult_smooth = float(ult[f'{variable}_smooth'])
            
            fig.add_trace(go.Scatter(
                x=[fecha_fin_norm, pred_fc['ds'].iloc[0]],
                y=[ult_smooth, float(pred_fc['yhat'].iloc[0])],
                mode='lines',
                line=dict(color=PRED_COLOR, width=2.0, dash='dot'),
                showlegend=False, hoverinfo='skip', legendgroup=lg
            ), row=row, col=1)
            
            fig.add_trace(go.Scatter(
                x=pred_fc['ds'].tolist(),
                y=pred_fc['yhat_lower'].tolist(),
                mode='lines',
                line=dict(width=0, color='rgba(0,0,0,0)'),
                showlegend=False, hoverinfo='skip',
                legendgroup=lg, name='_lower'
            ), row=row, col=1)
            fig.add_trace(go.Scatter(
                x=pred_fc['ds'].tolist(),
                y=pred_fc['yhat_upper'].tolist(),
                mode='lines',
                line=dict(width=0, color='rgba(0,0,0,0)'),
                fill='tonexty',
                fillcolor=_hex_to_rgba(PRED_COLOR, 0.20),
                showlegend=(i == 0), name='Intervalo 95% Prophet',
                legendgroup=lg, hoverinfo='skip'
            ), row=row, col=1)
            
            fig.add_trace(go.Scatter(
                x=pred_fc['ds'].tolist(),
                y=pred_fc['yhat'].tolist(),
                mode='lines',
                line=dict(color=PRED_COLOR, width=3.5),
                showlegend=(i == 0), legendgroup=lg,
                name=f'Prophet +{dias_pred}d',
                hovertemplate='%{x|%d/%b/%Y}<br>Pred: %{y:.1f}°C<extra>Prophet</extra>'
            ), row=row, col=1)
            
            ult_pred_dt = pred_fc['ds'].iloc[-1]
            ult_pred_val = float(pred_fc['yhat'].iloc[-1])
            fig.add_annotation(
                x=ult_pred_dt, y=ult_pred_val,
                text=f"<b>{ult_pred_val:.1f}°C</b>",
                showarrow=True, arrowhead=2, arrowcolor=PRED_COLOR, arrowwidth=1.0,
                ax=-28, ay=-18, font=dict(size=9, color=PRED_COLOR),
                row=row, col=1
            )
        
        # Título SIMPLIFICADO (sin fechas)
        fig.layout.annotations[i].update(
            text=f"<b>{empresa}  —  {fundo}  —  {variable}</b>",
            font=dict(size=11, color='#1A1A1A'),
            x=0.0, xanchor='left'
        )
        
        xk = 'xaxis' if row == 1 else f'xaxis{row}'
        fig.layout[xk].update(
            range=[zoom_ini, zoom_fin],
            tickformat='%d/%b', ticklabelmode='period',
            gridcolor=GRID_COLOR, gridwidth=0.5,
            showgrid=True, zeroline=False,
            autorange=False,
            rangeslider=dict(visible=False)
        )
        yk = 'yaxis' if row == 1 else f'yaxis{row}'
        fig.layout[yk].update(
            ticksuffix='°C', gridcolor=GRID_COLOR, gridwidth=0.5,
            showgrid=True, zeroline=False,
        )
    

    
    fig.update_layout(
        height=320 * n,
        paper_bgcolor=fig_bg,
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Arial', size=9, color='#333333'),
        title=dict(
            text=f"<b>{variable} — FUNDOS AQUANQA</b>",
            font=dict(size=14, color='#1A1A1A'),
            x=0.0, xanchor='left', pad=dict(t=10, l=10)
        ),
        margin=dict(l=60, r=40, t=60, b=40),
        legend=dict(
            orientation='h', yanchor='top', y=-0.08,
            xanchor='left', x=0, font=dict(size=8),
            bgcolor='rgba(0,0,0,0)', borderwidth=0
        ),
        hovermode='x unified'
    )
    fig.add_annotation(
        text=(
            "Fuente: SENAMHI Normales Climatológicas 1991-2020 | "
            "Predicción: Prophet"
        ),
        xref='paper', yref='paper', x=0, y=-0.12,
        showarrow=False,
        font=dict(size=7, color='#546E7A', style='italic'),
        xanchor='left'
    )
    
    df_hist = pd.concat(rows_export, ignore_index=True) if rows_export else pd.DataFrame()
    df_pred = pd.concat(rows_pred_export, ignore_index=True) if rows_pred_export else pd.DataFrame()
    
    return fig, df_hist, df_pred

# ══════════════════════════════════════════════════════════════
# EXPORTAR EXCEL (igual que antes)
# ══════════════════════════════════════════════════════════════

def _generar_pred_horaria(df_pred: pd.DataFrame) -> pd.DataFrame:
    if df_pred.empty:
        return pd.DataFrame()
    
    rows = []
    grupos = df_pred.groupby(['Empresa', 'Fundo', 'Variable'])
    
    for (empresa, fundo, variable), grp in grupos:
        grp = grp.sort_values('Fecha').reset_index(drop=True)
        fechas = pd.to_datetime(grp['Fecha'])
        preds = grp['Pred'].values.astype(float)
        lows = grp['Pred_Low'].values.astype(float)
        highs = grp['Pred_High'].values.astype(float)
        
        n = len(fechas)
        x_ctrl = np.arange(n, dtype=float) * 24.0
        
        if n < 2:
            x_ctrl = np.array([0.0, 24.0])
            preds = np.array([preds[0], preds[0]])
            lows = np.array([lows[0], lows[0]])
            highs = np.array([highs[0], highs[0]])
        
        cs_pred = CubicSpline(x_ctrl, preds)
        cs_low = CubicSpline(x_ctrl, lows)
        cs_high = CubicSpline(x_ctrl, highs)
        
        horas_rango = pd.date_range(
            start=fechas.iloc[0].normalize(),
            end=fechas.iloc[-1].normalize() + pd.Timedelta(hours=23),
            freq='h'
        )
        
        for h_dt in horas_rango:
            x_h = (h_dt - fechas.iloc[0].normalize()).total_seconds() / 3600.0
            x_h = float(np.clip(x_h, x_ctrl[0], x_ctrl[-1]))
            hora = h_dt.hour
            
            base = float(cs_pred(x_h))
            lo = float(cs_low(x_h))
            hi = float(cs_high(x_h))
            ampl = max((hi - lo) / 4.0, 0.0)
            
            if variable == 'Tmax':
                offset = ampl * np.sin(2 * np.pi * (hora - 5) / 24)
            else:
                offset = -ampl * np.sin(2 * np.pi * (hora - 5) / 24)
            
            rows.append({
                'Empresa': empresa,
                'Fundo': fundo,
                'Fecha': h_dt.strftime('%Y/%m/%d'),
                'Hora': h_dt.strftime('%H:00'),
                'FechaHora': h_dt,
                'Variable': variable,
                'max_pred': round(base + offset, 1) if variable == 'Tmax' else None,
                'min_pred': round(base + offset, 1) if variable == 'Tmin' else None,
                'Tipo': 'PREDICCION',
            })
    
    return pd.DataFrame(rows)

def exportar_excel(df_hist: pd.DataFrame, df_pred: pd.DataFrame) -> bytes:
    HDR_HIST = PatternFill('solid', start_color='1E3A5F', end_color='1E3A5F')
    HDR_HORA = PatternFill('solid', start_color='1B5E20', end_color='1B5E20')
    HDR_DET = PatternFill('solid', start_color='37474F', end_color='37474F')
    REAL_F = PatternFill('solid', start_color='FFFF00', end_color='FFFF00')
    PRED_F = PatternFill('solid', start_color='FFE0B2', end_color='FFE0B2')
    ALT_F = PatternFill('solid', start_color='F3F4F6', end_color='F3F4F6')
    
    HDR_FNT_W = Font(bold=True, color='FFFFFF', name='Arial', size=9)
    REAL_FNT = Font(name='Arial', size=8, color='000000')
    PRED_FNT = Font(name='Arial', size=8, color='BF360C', bold=True)
    CENT = Alignment(horizontal='center', vertical='center')
    
    def _autowidth(ws, max_w=22):
        for col in ws.columns:
            maxl = max((len(str(c.value)) if c.value is not None else 0) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(maxl + 3, max_w)
    
    wb = Workbook()
    
    # HOJA 1: Por Fechas
    ws_dia = wb.active
    ws_dia.title = 'Por_Fechas'
    ws_dia.sheet_properties.tabColor = '1E3A5F'
    
    filas_dia = []
    # Obtener fundos de ambos DataFrames, validando que no estén vacíos
    fundos_hist = list(df_hist['Fundo'].unique()) if not df_hist.empty and 'Fundo' in df_hist.columns else []
    fundos_pred = list(df_pred['Fundo'].unique()) if not df_pred.empty and 'Fundo' in df_pred.columns else []
    fundos_todos = sorted(set(fundos_hist + fundos_pred))
    
    for fundo in fundos_todos:
        h = df_hist[df_hist['Fundo'] == fundo].copy() if not df_hist.empty and 'Fundo' in df_hist.columns else pd.DataFrame()
        p = df_pred[df_pred['Fundo'] == fundo].copy() if not df_pred.empty and 'Fundo' in df_pred.columns else pd.DataFrame()
        
        empresa = (
            h['Empresa'].iloc[0] if not h.empty and 'Empresa' in h.columns else
            (p['Empresa'].iloc[0] if not p.empty and 'Empresa' in p.columns else '')
        )
        
        if 'Tmax' in h.columns and 'Tmin' in h.columns:
            hist_piv = (
                h[['Fecha', 'Tmax', 'Tmin']].drop_duplicates()
                 .rename(columns={'Tmax': 'max', 'Tmin': 'min'})
            )
        else:
            hist_piv = pd.DataFrame()
        
        if 'Pred' in p.columns:
            p_max = p[p['Variable'] == 'Tmax'][['Fecha', 'Pred']].rename(columns={'Pred': 'max pred'})
            p_min = p[p['Variable'] == 'Tmin'][['Fecha', 'Pred']].rename(columns={'Pred': 'min pred'})
            pred_piv = (
                p_max.merge(p_min, on='Fecha', how='outer') if not p_max.empty else
                pd.DataFrame(columns=['Fecha', 'max pred', 'min pred'])
            )
        else:
            pred_piv = pd.DataFrame()
        
        # Merge seguro: validar que ambos tengan 'Fecha' antes de hacer merge
        if not hist_piv.empty and not pred_piv.empty and 'Fecha' in hist_piv.columns and 'Fecha' in pred_piv.columns:
            tabla = hist_piv.merge(pred_piv, on='Fecha', how='outer')
        elif not hist_piv.empty:
            tabla = hist_piv.copy()
        elif not pred_piv.empty:
            tabla = pred_piv.copy()
        else:
            tabla = pd.DataFrame()
        
        if not tabla.empty:
            tabla['Fecha'] = pd.to_datetime(tabla['Fecha'])
            tabla = tabla.sort_values('Fecha').reset_index(drop=True)
            tabla['Empresa'] = empresa
            tabla['Fundo'] = fundo
            
            def _status(r):
                v = r.get('max', np.nan)
                try:
                    return 'real' if pd.notna(v) and not np.isnan(float(v)) else 'predicha'
                except Exception:
                    return 'predicha'
            tabla['status'] = tabla.apply(_status, axis=1)
            filas_dia.append(tabla)
    
    tabla_dia = pd.concat(filas_dia, ignore_index=True) if filas_dia else pd.DataFrame()
    tabla_dia = tabla_dia.sort_values(['Fundo', 'Fecha']).reset_index(drop=True)
    
    cols_dia = ['Empresa', 'Fundo', 'Fecha', 'max', 'min', 'max pred', 'min pred', 'status']
    for ci, h_txt in enumerate(cols_dia, 1):
        c = ws_dia.cell(row=1, column=ci)
        c.value = h_txt; c.fill = HDR_HIST
        c.font = HDR_FNT_W; c.alignment = CENT; c.border = THIN
    
    if not tabla_dia.empty:
        for ri, row_data in tabla_dia.iterrows():
            er = ri + 2
            is_real = row_data.get('status', 'predicha') == 'real'
            fill = REAL_F if is_real else PRED_F
            font = REAL_FNT if is_real else PRED_FNT
            
            def _v(key):
                val = row_data.get(key, np.nan)
                try:
                    return round(float(val), 1) if pd.notna(val) else None
                except Exception:
                    return None
            
            vals = [
                row_data.get('Empresa', ''),
                row_data.get('Fundo', ''),
                row_data['Fecha'].strftime('%Y/%m/%d') if pd.notna(row_data.get('Fecha')) else '',
                _v('max'), _v('min'),
                _v('max pred'), _v('min pred'),
                row_data.get('status', 'predicha'),
            ]
            for ci, v in enumerate(vals, 1):
                c = ws_dia.cell(row=er, column=ci)
                c.value = v; c.fill = fill; c.font = font
                c.alignment = CENT; c.border = THIN
    
    _autowidth(ws_dia)
    ws_dia.freeze_panes = 'A2'
    
    # HOJA 2: Por Horas (reducida)
    ws_hora = wb.create_sheet('Por_Horas')
    ws_hora.sheet_properties.tabColor = '1B5E20'
    
    pred_hora_df = _generar_pred_horaria(df_pred)
    
    if not pred_hora_df.empty:
        sub_max = pred_hora_df[pred_hora_df['Variable'] == 'Tmax'][
            ['Empresa', 'Fundo', 'Fecha', 'Hora', 'FechaHora', 'max_pred']
        ].copy()
        sub_min = pred_hora_df[pred_hora_df['Variable'] == 'Tmin'][
            ['Empresa', 'Fundo', 'Fecha', 'Hora', 'FechaHora', 'min_pred']
        ].copy()
        
        if not sub_max.empty and not sub_min.empty:
            tabla_h = sub_max.merge(
                sub_min, on=['Empresa', 'Fundo', 'Fecha', 'Hora', 'FechaHora'], how='outer'
            )
        elif not sub_max.empty:
            tabla_h = sub_max.copy()
            if 'min_pred' not in tabla_h.columns:
                tabla_h['min_pred'] = None
        elif not sub_min.empty:
            tabla_h = sub_min.copy()
            if 'max_pred' not in tabla_h.columns:
                tabla_h['max_pred'] = None
        else:
            tabla_h = pd.DataFrame()
        
        if not tabla_h.empty:
            tabla_h = tabla_h.sort_values(['Fundo', 'FechaHora']).reset_index(drop=True)
            
            cols_hora = ['Empresa', 'Fundo', 'Fecha', 'Hora', 'max pred', 'min pred', 'status']
            for ci, h_txt in enumerate(cols_hora, 1):
                c = ws_hora.cell(row=1, column=ci)
                c.value = h_txt; c.fill = HDR_HORA
                c.font = HDR_FNT_W; c.alignment = CENT; c.border = THIN
            
            for ri, row_data in tabla_h.iterrows():
                er = ri + 2
                max_v = row_data.get('max_pred')
                min_v = row_data.get('min_pred')
                vals_h = [
                    row_data.get('Empresa', ''),
                    row_data.get('Fundo', ''),
                    row_data.get('Fecha', ''),
                    row_data.get('Hora', ''),
                    round(float(max_v), 1) if max_v is not None and not pd.isna(max_v) else None,
                    round(float(min_v), 1) if min_v is not None and not pd.isna(min_v) else None,
                    'predicha',
                ]
                for ci, v in enumerate(vals_h, 1):
                    c = ws_hora.cell(row=er, column=ci)
                    c.value = v; c.fill = PRED_F; c.font = PRED_FNT
                    c.alignment = CENT; c.border = THIN
            
            _autowidth(ws_hora)
            ws_hora.freeze_panes = 'A2'
    
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

# ══════════════════════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def _leer_normales_desde_disco(path: str):
    with open(path, 'rb') as f:
        return f.read()

def cargar_normales_fijo():
    if not os.path.exists(NORMALES_PATH):
        st.error(f"**Archivo de normales no encontrado.**\n```\n{NORMALES_PATH}\n```")
        return None
    try:
        norm_bytes = _leer_normales_desde_disco(NORMALES_PATH)
        MEDIA_TMAX, Q1_TMAX, Q3_TMAX, estaciones_tmax = cargar_normales(norm_bytes, 'TMAX')
        MEDIA_TMIN, Q1_TMIN, Q3_TMIN, _ = cargar_normales(norm_bytes, 'TMIN')
        return MEDIA_TMAX, Q1_TMAX, Q3_TMAX, MEDIA_TMIN, Q1_TMIN, Q3_TMIN, estaciones_tmax
    except Exception as e:
        st.error(f"Error al leer normales SENAMHI: {e}")
        return None

# ──── SIDEBAR MEJORADO ────
with st.sidebar:
    st.markdown("## ⚙️ CONFIGURACIÓN")
    
    # Sección 1: Archivo
    with st.expander("📂 **Archivo meteorológico**", expanded=True):
        st.markdown("### 📊 **Archivo meteorológico**")
        st.info("✅ Leyendo: `assets/Metereologia_Prize.xlsx`")
    
    st.divider()
    
    # Sección 2: Fundos
    st.markdown("### 🌿 **Fundos**")
    fundos_disponibles_placeholder = st.empty()
    
    st.divider()
    
    # Estado de normales
    if os.path.exists(NORMALES_PATH):
        st.success("✅ Normales SENAMHI cargadas")
    else:
        st.error("❌ Normales SENAMHI no encontradas")

# ──── HEADER ────
st.markdown("""
<div class="header-bar">
    <h1>🌡️ Temperatura Fundos — Aquanqa</h1>
    <p>Climatología SENAMHI · Predicción Prophet · OPTIMIZADO</p>
</div>
""", unsafe_allow_html=True)

# ──── VALORES POR DEFECTO ────
dias_vista_num = 30  # Mostrar últimos 30 días
variable_sel = 'Ambas'  # Mostrar Tmax y Tmin
min_reg_ui = MIN_REGISTROS  # Registros mínimos por día (valor de constante)
dias_pred_ui = 30  # Días de predicción Prophet

# ──── CARGAR NORMALES ────
normales = cargar_normales_fijo()
if normales is None:
    st.stop()

MEDIA_TMAX, Q1_TMAX, Q3_TMAX, MEDIA_TMIN, Q1_TMIN, Q3_TMIN, estaciones_tmax = normales

# ──── LECTURA DIRECTA DEL ARCHIVO ────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
METEO_PATH = os.path.join(SCRIPT_DIR, "assets", "Metereologia_Prize.xlsx")

# Intentar rutas alternativas si la principal no existe
if not os.path.exists(METEO_PATH):
    alt_paths = [
        "assets/Metereologia_Prize.xlsx",
        "./assets/Metereologia_Prize.xlsx",
        "../assets/Metereologia_Prize.xlsx",
    ]
    for alt_path in alt_paths:
        if os.path.exists(alt_path):
            METEO_PATH = alt_path
            break

with st.spinner(f"Leyendo meteorología (optimizado)..."):
    try:
        if not os.path.exists(METEO_PATH):
            st.error(f"❌ Archivo no encontrado en: {METEO_PATH}")
            st.info("📂 Rutas buscadas:")
            st.code(f"{SCRIPT_DIR}/assets/Metereologia_Prize.xlsx\n./assets/Metereologia_Prize.xlsx")
            st.stop()
        
        # Leer archivo directamente desde disco
        with open(METEO_PATH, 'rb') as f:
            meteo_bytes = f.read()
        
        filename = "Metereologia_Prize.xlsx"
        df_preview = leer_meteo_bytes_optimizado(meteo_bytes, filename)
        
        if 'Fundo' not in df_preview.columns:
            st.error(f"Columna `Fundo` no encontrada.")
            st.stop()
        fundos_disponibles = sorted(df_preview['Fundo'].dropna().unique().tolist())
    except Exception as e:
        st.error(f"Error: {e}")
        st.stop()

with fundos_disponibles_placeholder:
    fundos_sel = st.multiselect(
        "Seleccionar fundos",
        options=fundos_disponibles,
        default=fundos_disponibles,
    )

fundos_activos = fundos_sel if fundos_sel else fundos_disponibles

with st.spinner("Procesando datos (optimizado)..."):
    try:
        dia_full = cargar_meteoro_optimizado(
            meteo_bytes, filename, SHEET_NAME,
            tuple(),
            min_reg_ui
        )
        dia = dia_full[dia_full['Fundo'].isin(fundos_activos)].reset_index(drop=True)
    except Exception as e:
        st.error(f"Error: {e}")
        st.stop()

# ──── PROPHET ────
cache_key = f"forecasts_{hash(meteo_bytes)}_{dias_pred_ui}"

if cache_key not in st.session_state:
    for k in list(st.session_state.keys()):
        if k.startswith("forecasts_"):
            del st.session_state[k]
    st.session_state[cache_key] = entrenar_todos_optimizado(dia_full, dias_pred_ui)

forecasts_cache = st.session_state[cache_key]

# ──── MÉTRICAS ────
st.markdown("#### 📊 Resumen de datos")

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="label">Registros</div>
        <div class="value">{len(dia):,}</div>
        <div class="sub">días-fundo</div>
    </div>""", unsafe_allow_html=True)
with col2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="label">Fundos</div>
        <div class="value">{dia['Fundo'].nunique()}</div>
        <div class="sub">activos</div>
    </div>""", unsafe_allow_html=True)
with col3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="label">Desde</div>
        <div class="value" style="font-size:1.15rem">{dia['Fecha'].min().strftime('%d/%m')}</div>
        <div class="sub">{dia['Fecha'].min().year}</div>
    </div>""", unsafe_allow_html=True)
with col4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="label">Hasta</div>
        <div class="value" style="font-size:1.15rem">{dia['Fecha'].max().strftime('%d/%m')}</div>
        <div class="sub">{dia['Fecha'].max().year}</div>
    </div>""", unsafe_allow_html=True)
with col5:
    st.markdown(f"""
    <div class="metric-card">
        <div class="label">Tmax prom</div>
        <div class="value">{dia['Tmax'].mean():.1f}°C</div>
        <div class="sub">todos</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

# ──── TABS ────
mostrar_tmax = variable_sel in ['Ambas', 'Solo Tmax']
mostrar_tmin = variable_sel in ['Ambas', 'Solo Tmin']

# Calcular días de predicción dinámicos
dias_pred_mitad = max(7, dias_pred_ui // 2)  # Mínimo 7 días

if mostrar_tmax and mostrar_tmin:
    tab_tmax, tab_tmin, tab_datos = st.tabs(
        ["🔴 Temperatura Máxima", "🔵 Temperatura Mínima", "📋 Datos"]
    )
elif mostrar_tmax:
    tab_tmax, tab_datos = st.tabs(["🔴 Temperatura Máxima", "📋 Datos"])
    tab_tmin = None
else:
    tab_tmin, tab_datos = st.tabs(["🔵 Temperatura Mínima", "📋 Datos"])
    tab_tmax = None

if mostrar_tmax and tab_tmax is not None:
    with tab_tmax:
        col1, col2 = st.columns([2, 2])
        with col1:
            st.markdown("#### 📊 Predicción a mostrar")
            dias_pred_mostrar = st.radio(
                "Mostrar predicción de:",
                options=[
                    f"{dias_pred_mitad}d",
                    f"{dias_pred_ui}d"
                ],
                index=1,
                horizontal=True,
                key='dias_pred_mostrar_tmax'
            )
            dias_pred_mostrar_num = int(dias_pred_mostrar.replace('d', ''))
        
        tipo_viz_tmax_lower = 'linea'
        
        with st.spinner("Generando Tmax..."):
            fig_tmax, df_tmax_hist, df_tmax_pred = generar_figura(
                'Tmax', dia, MEDIA_TMAX, Q1_TMAX, Q3_TMAX, dias_pred_ui, forecasts_cache, dias_vista_num, dias_pred_mostrar_num, tipo_viz_tmax_lower
            )
        st.plotly_chart(fig_tmax, use_container_width=True)
else:
    df_tmax_hist = pd.DataFrame()
    df_tmax_pred = pd.DataFrame()

if mostrar_tmin and tab_tmin is not None:
    with tab_tmin:
        col1, col2 = st.columns([2, 2])
        with col1:
            st.markdown("#### 📊 Predicción a mostrar")
            dias_pred_mostrar = st.radio(
                "Mostrar predicción de:",
                options=[
                    f"{dias_pred_mitad}d",
                    f"{dias_pred_ui}d"
                ],
                index=1,
                horizontal=True,
                key='dias_pred_mostrar_tmin'
            )
            dias_pred_mostrar_num = int(dias_pred_mostrar.replace('d', ''))
        
        tipo_viz_tmin_lower = 'linea'
        
        with st.spinner("Generando Tmin..."):
            fig_tmin, df_tmin_hist, df_tmin_pred = generar_figura(
                'Tmin', dia, MEDIA_TMIN, Q1_TMIN, Q3_TMIN, dias_pred_ui, forecasts_cache, dias_vista_num, dias_pred_mostrar_num, tipo_viz_tmin_lower
            )
        st.plotly_chart(fig_tmin, use_container_width=True)
else:
    df_tmin_hist = pd.DataFrame()
    df_tmin_pred = pd.DataFrame()

# ──── TAB DATOS ────
with tab_datos:
    st.markdown("#### 📋 Tabla de datos diarios")
    
    col_f, col_var = st.columns([2, 2])
    with col_f:
        fundos_tabla = st.multiselect(
            "Filtrar fundo", dia['Fundo'].unique().tolist(),
            default=dia['Fundo'].unique().tolist(), key='tabla_fundo'
        )
    with col_var:
        var_tabla = st.selectbox("Variable", ['Tmax', 'Tmin', 'Ambas'], key='tabla_var')
    
    df_show = dia[dia['Fundo'].isin(fundos_tabla)].copy()
    df_show['Fecha'] = df_show['Fecha'].dt.strftime('%Y-%m-%d')
    
    if var_tabla == 'Tmax':
        cols_show = ['Empresa', 'Fundo', 'Fecha', 'Tmax', 'Tmax_smooth']
    elif var_tabla == 'Tmin':
        cols_show = ['Empresa', 'Fundo', 'Fecha', 'Tmin', 'Tmin_smooth']
    else:
        cols_show = ['Empresa', 'Fundo', 'Fecha', 'Tmax', 'Tmax_smooth', 'Tmin', 'Tmin_smooth']
    
    st.dataframe(df_show[cols_show], use_container_width=True, hide_index=True, height=340)
    
    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
    st.markdown("#### ⬇️ Descargas")
    
    col_d1, col_d2, col_d3 = st.columns(3)
    
    df_hist_all = pd.concat([df_tmax_hist, df_tmin_hist], ignore_index=True)
    df_pred_all = pd.concat([df_tmax_pred, df_tmin_pred], ignore_index=True)
    
    with col_d1:
        if not df_hist_all.empty:
            excel_bytes = exportar_excel(df_hist_all, df_pred_all)
            st.download_button(
                label="📥 Excel completo",
                data=excel_bytes,
                file_name="Temperatura_Aquanqa_BI.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    with col_d2:
        if not df_hist_all.empty:
            st.download_button(
                label="📄 CSV histórico",
                data=df_hist_all.to_csv(index=False).encode('utf-8'),
                file_name="historico_temperatura.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    with col_d3:
        if not df_pred_all.empty:
            st.download_button(
                label="📄 CSV predicciones",
                data=df_pred_all.to_csv(index=False).encode('utf-8'),
                file_name="predicciones_prophet.csv",
                mime="text/csv",
                use_container_width=True
            )

st.markdown("""
<div class="footer-note">
    Elaborado por la Gerencia de Planificación, Control Operacional y Gestión
</div>
""", unsafe_allow_html=True)