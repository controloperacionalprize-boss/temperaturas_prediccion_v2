import calendar
import hashlib
import io
import json
import pickle
import warnings

import numpy as np
import pandas as pd
import streamlit as st
from prophet import Prophet

from config.config import PROPHET_PARAMS, H_MAX_PENDIENTE, CACHE_DIR, CACHE_PKL, CACHE_HASH, CODE_VERSION
from services.climatologia_service import calcular_climatologia_armonica, predecir_climatologia_armonica
from services.enfen_service import obtener_ajuste_enfen


def _hash_serie(serie_bytes: bytes) -> str:
    return hashlib.md5(serie_bytes).hexdigest()[:12]


@st.cache_data(show_spinner=False)
def entrenar_prophet_opt(_serie_hash: str, _serie_bytes: bytes, dias_pred: int,
                         variable: str, fundo: str):
    """Prophet con historial completo + peso a datos recientes + corrección de bias."""
    serie = pd.read_parquet(io.BytesIO(_serie_bytes))

    df = (
        serie.rename(columns={'Fecha': 'ds', 'Valor': 'y'})
             .dropna()
             .sort_values('ds')
             .reset_index(drop=True)
    )
    df['ds'] = pd.to_datetime(df['ds']).dt.tz_localize(None)

    media = df['y'].mean()
    std   = df['y'].std()
    df['y'] = np.clip(df['y'], media - 3.0 * std, media + 3.0 * std)

    # ── Climatología armónica ──────────────────────────────────
    coef_clim = calcular_climatologia_armonica(
        df[['ds', 'y']], PROPHET_PARAMS['n_harmonics'], PROPHET_PARAMS['clim_halflife_anios']
    )
    _fi_c = df['ds'].min()
    _ff_c = df['ds'].max() + pd.Timedelta(days=dias_pred + 60)
    _fechas_clim = pd.date_range(_fi_c, _ff_c, freq='D')
    _clim_daily = pd.DataFrame({
        'ds'  : _fechas_clim,
        'clim': predecir_climatologia_armonica(
            _fechas_clim, coef_clim, PROPHET_PARAMS['n_harmonics']
        ).round(2)
    })
    df = df.merge(_clim_daily, on='ds', how='left')
    _y_abs_min, _y_abs_max = df['y'].min(), df['y'].max()
    df['y'] = (df['y'] - df['clim']).round(4)

    df['weight'] = 1.0
    mask_reciente = df['ds'] >= df['ds'].max() - pd.Timedelta(days=180)
    df.loc[mask_reciente, 'weight'] = 2.5

    # ── Walk-forward cross-validation ─────────────────────────
    n_total     = len(df)
    n_validacion = min(15, max(10, n_total // 8))
    mae_por_h_listas = {h: [] for h in range(1, min(dias_pred + 1, 31))}

    if n_validacion >= 10:
        for i in range(n_validacion):
            corte = n_total - n_validacion + i
            train = df.iloc[:corte].copy()
            test  = df.iloc[corte:corte + dias_pred].copy()

            if len(test) < 1 or len(train) < 30:
                break

            train['weight'] = 1.0
            mask_rec_train = train['ds'] >= train['ds'].max() - pd.Timedelta(days=180)
            train.loc[mask_rec_train, 'weight'] = 2.5

            try:
                with warnings.catch_warnings():
                    warnings.simplefilter('ignore')
                    m = Prophet(
                        yearly_seasonality=PROPHET_PARAMS['yearly_seasonality'],
                        weekly_seasonality=False,
                        daily_seasonality=False,
                        seasonality_mode='additive',
                        changepoint_prior_scale=PROPHET_PARAMS['changepoint_prior_scale'],
                        changepoint_range=PROPHET_PARAMS['changepoint_range'],
                        seasonality_prior_scale=PROPHET_PARAMS['seasonality_prior_scale'],
                        interval_width=PROPHET_PARAMS['interval_width'],
                    )
                    m.add_seasonality(name='monthly', period=30.5,
                                      fourier_order=PROPHET_PARAMS['fourier_monthly'])
                    m.fit(train)
                    future = m.make_future_dataframe(periods=len(test), freq='D')
                    pred   = m.predict(future).tail(len(test))
                    _test_clim = _clim_daily[_clim_daily['ds'].isin(
                        pd.to_datetime(test['ds']).dt.normalize()
                    )].set_index('ds')['clim']

                    for h, (ds_val, real_anom, yhat_anom) in enumerate(
                        zip(test['ds'], test['y'].values, pred['yhat'].values), 1
                    ):
                        ds_norm  = pd.Timestamp(ds_val).normalize()
                        clim_val = float(_test_clim.get(ds_norm, _clim_daily['clim'].mean()))
                        real_abs = real_anom + clim_val
                        yhat_abs = yhat_anom + clim_val
                        if h <= 30:
                            mae_por_h_listas[h].append(abs(real_abs - yhat_abs))
            except Exception:
                continue

    todos_errores = [e for errs in mae_por_h_listas.values() for e in errs]
    mae_real  = round(float(np.mean(todos_errores)), 3) if todos_errores else None
    mae_por_h = {
        h: round(float(np.mean(errs)), 3) if errs else None
        for h, errs in mae_por_h_listas.items()
    }

    # ── Modelo final ───────────────────────────────────────────
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        modelo = Prophet(
            yearly_seasonality=PROPHET_PARAMS['yearly_seasonality'],
            weekly_seasonality=False,
            daily_seasonality=False,
            seasonality_mode='additive',
            changepoint_prior_scale=PROPHET_PARAMS['changepoint_prior_scale'],
            changepoint_range=PROPHET_PARAMS['changepoint_range'],
            seasonality_prior_scale=PROPHET_PARAMS['seasonality_prior_scale'],
            interval_width=PROPHET_PARAMS['interval_width'],
        )
        modelo.add_seasonality(name='monthly', period=30.5,
                               fourier_order=PROPHET_PARAMS['fourier_monthly'])
        modelo.fit(df)
        future   = modelo.make_future_dataframe(periods=dias_pred, freq='D')
        forecast = modelo.predict(future)

    forecast = forecast.merge(_clim_daily[['ds', 'clim']], on='ds', how='left')
    forecast['clim'] = forecast['clim'].ffill().bfill()
    for col in ['yhat', 'yhat_lower', 'yhat_upper']:
        forecast[col] = (forecast[col] + forecast['clim']).round(2)

    result_futuro    = forecast.tail(dias_pred)[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].copy().reset_index(drop=True)
    result_historico = forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].copy().reset_index(drop=True)

    insample = forecast[['ds', 'yhat']].merge(
        df[['ds', 'y', 'clim']].assign(y=lambda d: d['y'] + d['clim']),
        on='ds', how='inner'
    ).sort_values('ds').reset_index(drop=True)

    ventana_bias = min(PROPHET_PARAMS['bias_window_dias'] * 2, len(insample))
    ult = insample.tail(ventana_bias).copy().reset_index(drop=True)
    ult['residuo'] = ult['y'] - ult['yhat']
    ult['t'] = np.arange(len(ult))

    if len(ult) >= 5:
        b_slope, a_int = np.polyfit(ult['t'], ult['residuo'], 1)
    else:
        b_slope, a_int = 0.0, (ult['residuo'].mean() if len(ult) else 0.0)

    insample['anio']          = insample['ds'].dt.year
    insample['residuo_total'] = insample['y'] - insample['yhat']

    mes_objetivo  = result_futuro['ds'].iloc[0].month
    anio_objetivo = result_futuro['ds'].iloc[0].year
    datos_mes_prev = insample[
        (insample['ds'].dt.month == mes_objetivo) & (insample['anio'] < anio_objetivo)
    ]
    if len(datos_mes_prev) >= 15:
        edad_anios = anio_objetivo - datos_mes_prev['anio']
        pesos = np.exp(-edad_anios / 1.5)
        bias_estacional = float(np.average(datos_mes_prev['residuo_total'], weights=pesos))
    else:
        bias_estacional = 0.0

    h_arr      = np.arange(1, len(result_futuro) + 1)
    h_efectivo = np.minimum(h_arr, H_MAX_PENDIENTE)
    ajuste_enfen = obtener_ajuste_enfen(mes_objetivo, anio_objetivo, variable)

    bias_h_arr = (
        a_int + b_slope * (len(ult) - 1) + b_slope * h_efectivo
        + bias_estacional + ajuste_enfen
    )
    for col in ['yhat', 'yhat_lower', 'yhat_upper']:
        result_futuro[col] = (result_futuro[col] + bias_h_arr).round(2)

    margen = 5.0
    for col in ['yhat', 'yhat_lower', 'yhat_upper']:
        result_futuro[col] = np.clip(result_futuro[col], _y_abs_min - margen, _y_abs_max + margen)

    result_futuro['ds']    = pd.to_datetime(result_futuro['ds']).dt.tz_localize(None).dt.normalize()
    result_historico['ds'] = pd.to_datetime(result_historico['ds']).dt.tz_localize(None).dt.normalize()

    return result_futuro, result_historico, mae_real, mae_por_h


def entrenar_todos_optimizado(dia: pd.DataFrame, dias_pred: int) -> dict:
    """Entrena todos los modelos Prophet con barra de progreso."""
    forecasts = {}
    fundos    = dia['Fundo'].unique().tolist()
    variables = ['Tmax', 'Tmin', 'ET', 'RadSolarAlta']
    total     = len(fundos) * len(variables)

    prog = st.progress(0, text="Entrenando modelos Prophet...")

    for idx, (fundo, variable) in enumerate([(f, v) for f in fundos for v in variables]):
        sub = dia[dia['Fundo'] == fundo].copy()
        if sub.empty:
            continue

        if variable == 'ET':
            sub = sub[sub['ET'] > 0].copy()
            if len(sub) < 30:
                continue

        if len(sub) < 30:
            continue

        serie = sub[['Fecha', variable]].rename(columns={variable: 'Valor'})
        buf   = io.BytesIO()
        serie.to_parquet(buf, index=False)
        buf.seek(0)
        serie_bytes = buf.getvalue()
        serie_hash  = _hash_serie(serie_bytes)

        fecha_ultimo_dato = pd.to_datetime(serie['Fecha'].max()).normalize()
        primer_dia_mes    = fecha_ultimo_dato.replace(day=1)
        ultimo_dia_mes    = pd.Timestamp(
            year=primer_dia_mes.year,
            month=primer_dia_mes.month,
            day=calendar.monthrange(primer_dia_mes.year, primer_dia_mes.month)[1]
        )
        dias_pred_real = max(dias_pred, (ultimo_dia_mes - fecha_ultimo_dato).days)

        forecast, forecast_hist, mae_real, mae_por_h = entrenar_prophet_opt(
            serie_hash, serie_bytes, dias_pred_real, variable, fundo
        )

        if variable == 'ET':
            for col in ['yhat', 'yhat_lower', 'yhat_upper']:
                forecast[col]      = forecast[col].clip(lower=0)
                forecast_hist[col] = forecast_hist[col].clip(lower=0)
        if variable == 'RadSolarAlta':
            for col in ['yhat', 'yhat_lower', 'yhat_upper']:
                forecast[col]      = forecast[col].clip(lower=0)
                forecast_hist[col] = forecast_hist[col].clip(lower=0)
        forecasts[(fundo, variable)] = {
            'forecast'     : forecast,
            'forecast_hist': forecast_hist,
            'mae'          : mae_real,
            'mae_por_dia'  : mae_por_h,
        }

        prog.progress((idx + 1) / total,
                      text=f"Prophet {fundo} — {variable}  ({idx + 1}/{total})...")

    prog.empty()
    return forecasts


# ── Cache en disco ─────────────────────────────────────────────

def guardar_cache_prophet(forecasts: dict, data_hash: str) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    with open(CACHE_PKL, "wb") as f:
        pickle.dump(forecasts, f)
    CACHE_HASH.write_text(data_hash)


def cargar_cache_prophet(data_hash: str) -> dict | None:
    if not CACHE_PKL.exists() or not CACHE_HASH.exists():
        return None
    if CACHE_HASH.read_text().strip() != data_hash:
        return None
    with open(CACHE_PKL, "rb") as f:
        return pickle.load(f)


def calcular_hash_meteo(dia_full: pd.DataFrame) -> str:
    params_bytes  = json.dumps(PROPHET_PARAMS, sort_keys=True).encode()
    version_bytes = CODE_VERSION.encode()
    try:
        df_hash = hashlib.md5(
            pd.util.hash_pandas_object(dia_full, index=True).values
        ).hexdigest().encode()
    except Exception:
        df_hash = b"fabric_hash"
    return hashlib.md5(df_hash + params_bytes + version_bytes).hexdigest()
