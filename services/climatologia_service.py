import numpy as np
import pandas as pd
from scipy.interpolate import CubicSpline

import streamlit as st

from OPERACIONES.PREDICCION_TEMPERATURA.config.config import PROPHET_PARAMS


def calcular_climatologia_armonica(df_clim: pd.DataFrame, n_harmonics: int, halflife_anios: float) -> np.ndarray:
    """
    Ajusta climatología diaria vía regresión armónica (Fourier)
    sobre todos los años disponibles, con peso exponencial por antigüedad.
    df_clim: columnas ['ds','y'] — 'y' en temperatura absoluta.
    """
    d = df_clim[['ds', 'y']].dropna().copy()
    doy = d['ds'].dt.dayofyear
    theta = 2 * np.pi * doy / 365.25

    cols = {'const': np.ones(len(d))}
    for k in range(1, n_harmonics + 1):
        cols[f'sin{k}'] = np.sin(k * theta)
        cols[f'cos{k}'] = np.cos(k * theta)
    X = np.column_stack(list(cols.values()))

    edad_anios = ((d['ds'].max() - d['ds']).dt.days / 365.25).to_numpy()
    w = np.exp(-edad_anios / halflife_anios)

    Xw = X * w[:, None]
    coef, *_ = np.linalg.lstsq(Xw, d['y'].to_numpy() * w, rcond=None)
    return coef


def predecir_climatologia_armonica(fechas, coef: np.ndarray, n_harmonics: int) -> np.ndarray:
    doy = pd.to_datetime(fechas).dayofyear.to_numpy().astype(float)
    theta = 2 * np.pi * doy / 365.25
    cols = [np.ones(len(doy))]
    for k in range(1, n_harmonics + 1):
        cols.append(np.sin(k * theta))
        cols.append(np.cos(k * theta))
    X = np.column_stack(cols)
    return X @ coef


@st.cache_data(show_spinner=False)
def spline_diario_cached(fecha_inicio_str: str, fecha_fin_str: str,
                         valores_tuple: tuple) -> pd.DataFrame:
    """Spline cúbico mensual → diario, con clave de cache determinística."""
    fecha_inicio = pd.Timestamp(fecha_inicio_str)
    fecha_fin    = pd.Timestamp(fecha_fin_str)
    valores_mensuales = list(valores_tuple)

    fechas_ctrl, vals_ctrl = [], []
    for year in range(fecha_inicio.year - 1, fecha_fin.year + 2):
        for mes_idx, val in enumerate(valores_mensuales):
            fechas_ctrl.append(pd.Timestamp(year=year, month=mes_idx + 1, day=15))
            vals_ctrl.append(val)

    fechas_ctrl = pd.DatetimeIndex(fechas_ctrl)
    x_ctrl = (fechas_ctrl - fechas_ctrl[0]).days.values.astype(float)
    cs = CubicSpline(x_ctrl, vals_ctrl)

    rango  = pd.date_range(fecha_inicio, fecha_fin, freq='D')
    x_eval = (rango - fechas_ctrl[0]).days.values.astype(float)

    return pd.DataFrame({'Fecha': rango, 'Valor': cs(x_eval).round(2)})
