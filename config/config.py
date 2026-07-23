import os
from pathlib import Path

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
# Ruta absoluta relativa al config.py
NORMALES_PATH = os.path.join(
    os.path.dirname(__file__),  # carpeta config/
    '..', 'assets', 'NORMALES_1991_2020.xlsx'
)
H_MAX_PENDIENTE = 14  # días: más allá de esto, la pendiente deja de extrapolarse

SHEET_NAME    = "Prize_Climatology"
MIN_REGISTROS = 80
ROLLING_DIAS  = 3
SIGMA_CLIM    = 1.5
Z_Q           = 0.6745
DELTA_Q       = round(SIGMA_CLIM * Z_Q, 2)

PROPHET_PARAMS = {
    "changepoint_prior_scale" : 0.15,
    "changepoint_range"       : 0.95,
    "seasonality_prior_scale" : 2.0,
    "yearly_seasonality"      : 8,
    "fourier_monthly"         : 3,
    "bias_window_dias"        : 21,
    "interval_width"          : 0.90,
    "anomaly_based"           : "v4",
    "clim_halflife_days"      : 90,
    "n_harmonics"             : 4,
    "clim_halflife_anios"     : 2.0,
}

CODE_VERSION = "v2.4-fabric-secrets"

CACHE_DIR  = Path(_APP_DIR) / "assets" / ".prophet_cache"
CACHE_PKL  = CACHE_DIR / "forecasts.pkl"
CACHE_HASH = CACHE_DIR / "data_hash.txt"

TMAX_MIN, TMAX_MAX = 10.0, 45.0
TMIN_MIN, TMIN_MAX =  5.0, 35.0

# ── Colores ────────────────────────────────────────────────────
GRID_COLOR      = '#CFD8DC'
BG_TMAX         = '#FFFFFF'
BANDA_TMAX      = '#FFCDD2'
CLIM_TMAX       = '#C62828'
REAL_TMAX_COLOR = '#000000'
BG_TMIN         = '#FFFFFF'
BANDA_TMIN      = '#BBDEFB'
CLIM_TMIN       = '#1565C0'
REAL_TMIN_COLOR = '#000000'
PRED_COLOR_TMAX = '#FF0000'
PRED_COLOR_TMIN = '#1565C0'
CLIM2_COLOR     = '#FF6F00'
BANDA2_TMAX     = '#FFE0B2'
BANDA2_TMIN     = '#E3F2FD'

MESES_RAW = [
    'Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
]

DISTRITO_BUSCAR  = 'GUADALUPE'
DISTRITO_NINGUNA = '— Ninguna —'


# ── Mapa / riesgo ──────────────────────────────────────────────
COLORES_TEMP_SENAMHI = [
    '#0D47A1', '#1976D2', '#42A5F5', '#81D4FA',
    '#FFF59D', '#FFB74D', '#FB8C00', '#E64A19', '#B71C1C',
]
DIAS_ET_PROMEDIO   = 7
COLORES_ET         = ['#E1F5FE', '#81D4FA', '#29B6F6', '#0288D1', '#01579B']
DIAS_RIESGO_VENTANA = 7

RIESGO_COLOR = {
    'Bajo': '#4CAF50', 'Medio': '#FFC107', 'Alto': '#FB8C00',
    'Muy alto': '#C62828', 'Sin datos': '#90A4AE',
}
RIESGO_COLOR_ET = {
    'Bajo':     '#E1F5FE',
    'Medio':    '#81D4FA',
    'Alto':     '#42A5F5',
    'Muy alto': '#0D47A1',
    'Sin datos':'#90A4AE',
}
