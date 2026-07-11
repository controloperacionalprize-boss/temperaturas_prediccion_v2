import calendar
import itertools
import warnings

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from prophet import Prophet
from scipy.stats import norm

from config.config import PROPHET_PARAMS, H_MAX_PENDIENTE
from services.climatologia_service import calcular_climatologia_armonica, predecir_climatologia_armonica
from services.enfen_service import obtener_ajuste_enfen


def _score_combo(sub_fundo: pd.DataFrame, variable: str, ventana: dict, params_combo: dict):
    """Entrena con params_combo y devuelve MAE/Precisión/MBE/Std para la ventana dada."""
    fecha_inicio = ventana['inicio']
    fecha_fin    = ventana['fin']
    fecha_corte  = ventana['fecha_corte']

    train = sub_fundo[sub_fundo['Fecha'] <= fecha_corte].copy()
    if len(train) < 30:
        return None

    serie_train = train[['Fecha', variable]].rename(columns={'Fecha': 'ds', variable: 'y'}).dropna()
    serie_train['ds'] = pd.to_datetime(serie_train['ds']).dt.tz_localize(None).dt.normalize()

    media, std = serie_train['y'].mean(), serie_train['y'].std()
    serie_train['y'] = serie_train['y'].clip(lower=media - 3*std, upper=media + 3*std)

    coef_clim   = calcular_climatologia_armonica(
        serie_train[['ds', 'y']], params_combo['n_harmonics'], params_combo['clim_halflife_anios']
    )
    fechas_clim = pd.date_range(serie_train['ds'].min(), fecha_fin + pd.Timedelta(days=30), freq='D')
    clim_daily  = pd.DataFrame({
        'ds'  : fechas_clim,
        'clim': predecir_climatologia_armonica(fechas_clim, coef_clim, params_combo['n_harmonics']).round(2)
    })

    serie_train = serie_train.merge(clim_daily, on='ds', how='left')
    serie_train['y_abs'] = serie_train['y']
    serie_train['y']     = (serie_train['y'] - serie_train['clim']).round(4)

    try:
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            m = Prophet(
                yearly_seasonality=params_combo['yearly_seasonality'],
                weekly_seasonality=False, daily_seasonality=False,
                seasonality_mode='additive',
                changepoint_prior_scale=params_combo['changepoint_prior_scale'],
                changepoint_range=params_combo['changepoint_range'],
                seasonality_prior_scale=params_combo['seasonality_prior_scale'],
                interval_width=params_combo['interval_width'],
            )
            m.add_seasonality(name='monthly', period=30.5, fourier_order=params_combo['fourier_monthly'])
            m.fit(serie_train[['ds', 'y']])
            dias_pred_period = (fecha_fin - fecha_inicio).days + 1
            future   = m.make_future_dataframe(periods=dias_pred_period, freq='D')
            forecast = m.predict(future)
    except Exception:
        return None

    forecast['ds'] = pd.to_datetime(forecast['ds']).dt.tz_localize(None).dt.normalize()
    forecast = forecast.merge(clim_daily, on='ds', how='left')
    forecast['clim'] = forecast['clim'].ffill().bfill()
    forecast['yhat'] = (forecast['yhat'] + forecast['clim']).round(2)

    pred_periodo = forecast[
        (forecast['ds'] >= fecha_inicio) & (forecast['ds'] <= fecha_fin)
    ][['ds', 'yhat']]

    real_periodo = sub_fundo[
        (sub_fundo['Fecha'] >= fecha_inicio) & (sub_fundo['Fecha'] <= fecha_fin)
    ][['Fecha', variable]].rename(columns={'Fecha': 'ds', variable: 'Real'})

    comp = real_periodo.merge(pred_periodo, on='ds', how='inner')
    if comp.empty:
        return None

    insample = forecast[['ds', 'yhat']].merge(
        serie_train[['ds', 'y_abs']], on='ds', how='inner'
    ).sort_values('ds').reset_index(drop=True)
    insample['residuo'] = insample['y_abs'] - insample['yhat']

    ventana_bias = min(params_combo['bias_window_dias'] * 2, len(insample))
    ult = insample.tail(ventana_bias).reset_index(drop=True)
    ult['t'] = np.arange(len(ult))
    if len(ult) >= 5:
        b_slope, a_int = np.polyfit(ult['t'], ult['residuo'], 1)
    else:
        b_slope, a_int = 0.0, (ult['residuo'].mean() if len(ult) else 0.0)

    insample['anio']    = insample['ds'].dt.year
    mes_obj             = fecha_inicio.month
    datos_mes_prev      = insample[
        (insample['ds'].dt.month == mes_obj) & (insample['anio'] < fecha_inicio.year)
    ]
    if len(datos_mes_prev) >= 15:
        edad  = fecha_inicio.year - datos_mes_prev['anio']
        pesos = np.exp(-edad / 1.5)
        bias_est = float(np.average(datos_mes_prev['residuo'], weights=pesos))
    else:
        bias_est = 0.0

    comp = comp.copy()
    comp['h'] = (comp['ds'] - pd.Timestamp(fecha_corte)).dt.days
    h_ef = np.minimum(comp['h'].clip(lower=0), H_MAX_PENDIENTE)
    comp['ajuste_enfen'] = comp['ds'].apply(lambda f: obtener_ajuste_enfen(f.month, f.year, variable))
    comp['bias_h']   = a_int + b_slope * (len(ult) - 1) + b_slope * h_ef + bias_est + comp['ajuste_enfen']
    comp['pred_corr'] = comp['yhat'] + comp['bias_h']
    comp['error']     = comp['Real'] - comp['pred_corr']

    return {
        'MAE'      : round(comp['error'].abs().mean(), 2),
        'Precision': round((comp['error'].abs() <= 1.5).mean() * 100, 1),
        'MBE'      : round(comp['error'].mean(), 2),
        'Std'      : round(comp['error'].std(), 2),
    }


def grid_search_hiperparams(dia: pd.DataFrame, ventana: dict, grid: dict) -> pd.DataFrame:
    fundos    = dia['Fundo'].unique().tolist()
    variables = ['Tmax', 'Tmin']
    nombres   = list(grid.keys())
    combos    = list(itertools.product(*grid.values()))

    total = len(fundos) * len(variables) * len(combos)
    prog  = st.progress(0, text="Grid search...")
    c     = 0
    filas = []

    for fundo in fundos:
        sub_fundo = dia[dia['Fundo'] == fundo].copy().reset_index(drop=True)
        for variable in variables:
            for combo in combos:
                params_combo = {**PROPHET_PARAMS, **dict(zip(nombres, combo))}
                r = _score_combo(sub_fundo, variable, ventana, params_combo)
                c += 1
                prog.progress(c / total, text=f"{fundo} — {variable} — {dict(zip(nombres, combo))}")
                if r is None:
                    continue
                filas.append({'Fundo': fundo, 'Variable': variable,
                              **dict(zip(nombres, combo)), **r})

    prog.empty()
    return pd.DataFrame(filas)


def _forecast_mes_walkforward(sub_fundo: pd.DataFrame, variable: str,
                              fecha_inicio: pd.Timestamp, fecha_fin: pd.Timestamp,
                              fecha_corte: pd.Timestamp) -> pd.DataFrame:
    """Entrena solo con datos hasta fecha_corte (sin ver el mes objetivo) y devuelve
    la predicción corregida día a día para [fecha_inicio, fecha_fin]. Igual mecánica
    que _score_combo pero devuelve la serie en vez de métricas agregadas."""
    train = sub_fundo[sub_fundo['Fecha'] <= fecha_corte].copy()
    if len(train) < 30:
        return pd.DataFrame()

    serie_train = train[['Fecha', variable]].rename(columns={'Fecha': 'ds', variable: 'y'}).dropna()
    serie_train['ds'] = pd.to_datetime(serie_train['ds']).dt.tz_localize(None).dt.normalize()

    media, std = serie_train['y'].mean(), serie_train['y'].std()
    serie_train['y'] = serie_train['y'].clip(lower=media - 3 * std, upper=media + 3 * std)

    coef_clim = calcular_climatologia_armonica(
        serie_train[['ds', 'y']], PROPHET_PARAMS['n_harmonics'], PROPHET_PARAMS['clim_halflife_anios']
    )
    fechas_clim = pd.date_range(serie_train['ds'].min(), fecha_fin + pd.Timedelta(days=30), freq='D')
    clim_daily = pd.DataFrame({
        'ds'  : fechas_clim,
        'clim': predecir_climatologia_armonica(fechas_clim, coef_clim, PROPHET_PARAMS['n_harmonics']).round(2)
    })

    serie_train = serie_train.merge(clim_daily, on='ds', how='left')
    serie_train['y_abs'] = serie_train['y']
    serie_train['y']     = (serie_train['y'] - serie_train['clim']).round(4)

    try:
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            m = Prophet(
                yearly_seasonality=PROPHET_PARAMS['yearly_seasonality'],
                weekly_seasonality=False, daily_seasonality=False,
                seasonality_mode='additive',
                changepoint_prior_scale=PROPHET_PARAMS['changepoint_prior_scale'],
                changepoint_range=PROPHET_PARAMS['changepoint_range'],
                seasonality_prior_scale=PROPHET_PARAMS['seasonality_prior_scale'],
                interval_width=PROPHET_PARAMS['interval_width'],
            )
            m.add_seasonality(name='monthly', period=30.5, fourier_order=PROPHET_PARAMS['fourier_monthly'])
            m.fit(serie_train[['ds', 'y']])
            dias_pred_period = (fecha_fin - fecha_inicio).days + 1
            future   = m.make_future_dataframe(periods=dias_pred_period, freq='D')
            forecast = m.predict(future)
    except Exception:
        return pd.DataFrame()

    forecast['ds'] = pd.to_datetime(forecast['ds']).dt.tz_localize(None).dt.normalize()
    forecast = forecast.merge(clim_daily, on='ds', how='left')
    forecast['clim'] = forecast['clim'].ffill().bfill()
    forecast['yhat'] = (forecast['yhat'] + forecast['clim']).round(2)

    pred_periodo = forecast[
        (forecast['ds'] >= fecha_inicio) & (forecast['ds'] <= fecha_fin)
    ][['ds', 'yhat']].copy()
    if pred_periodo.empty:
        return pd.DataFrame()

    insample = forecast[['ds', 'yhat']].merge(
        serie_train[['ds', 'y_abs']], on='ds', how='inner'
    ).sort_values('ds').reset_index(drop=True)
    insample['residuo'] = insample['y_abs'] - insample['yhat']

    ventana_bias = min(PROPHET_PARAMS['bias_window_dias'] * 2, len(insample))
    ult = insample.tail(ventana_bias).reset_index(drop=True)
    ult['t'] = np.arange(len(ult))
    if len(ult) >= 5:
        b_slope, a_int = np.polyfit(ult['t'], ult['residuo'], 1)
    else:
        b_slope, a_int = 0.0, (ult['residuo'].mean() if len(ult) else 0.0)

    insample['anio']    = insample['ds'].dt.year
    mes_obj             = fecha_inicio.month
    datos_mes_prev      = insample[
        (insample['ds'].dt.month == mes_obj) & (insample['anio'] < fecha_inicio.year)
    ]
    if len(datos_mes_prev) >= 15:
        edad  = fecha_inicio.year - datos_mes_prev['anio']
        pesos = np.exp(-edad / 1.5)
        bias_est = float(np.average(datos_mes_prev['residuo'], weights=pesos))
    else:
        bias_est = 0.0

    pred_periodo['h'] = (pred_periodo['ds'] - pd.Timestamp(fecha_corte)).dt.days
    h_ef = np.minimum(pred_periodo['h'].clip(lower=0), H_MAX_PENDIENTE)
    pred_periodo['bias_h']  = a_int + b_slope * (len(ult) - 1) + b_slope * h_ef + bias_est
    pred_periodo[variable]  = (pred_periodo['yhat'] + pred_periodo['bias_h']).round(2)

    return pred_periodo[['ds', variable]]


def generar_export_prediccion_historica(dia: pd.DataFrame) -> pd.DataFrame:
    """Tabla plana (Empresa, Fundo, Fecha, Tmax, Tmin, Tipo='Predicho') con walk-forward
    puro para los últimos 3 meses cerrados MÁS el mes en curso: cada uno de los 4 meses
    se entrena solo con datos reales hasta el día antes de que empiece ese mes, sin ver
    ni un día del mes que predice — incluido el mes en curso, aunque ya haya días reales
    transcurridos de él (se ignoran a propósito, todo sale como predicción del modelo)."""
    fundos = dia['Fundo'].unique().tolist()
    filas  = []

    for fundo in fundos:
        sub = dia[dia['Fundo'] == fundo].copy().reset_index(drop=True)
        if sub.empty:
            continue
        empresa = sub['Empresa'].iloc[0]

        fecha_max      = pd.to_datetime(sub['Fecha'].max()).normalize()
        mes_actual_ini = fecha_max.replace(day=1)

        piezas = {'Tmax': [], 'Tmin': []}

        # ── 3 meses cerrados + el mes en curso (k=0), todos walk-forward puro ──
        for k in range(3, -1, -1):
            mes_ini = (mes_actual_ini - pd.DateOffset(months=k)).replace(day=1)
            mes_fin = pd.Timestamp(
                year=mes_ini.year, month=mes_ini.month,
                day=calendar.monthrange(mes_ini.year, mes_ini.month)[1]
            )
            fecha_corte = mes_ini - pd.Timedelta(days=1)

            for variable in ('Tmax', 'Tmin'):
                res = _forecast_mes_walkforward(sub, variable, mes_ini, mes_fin, fecha_corte)
                if not res.empty:
                    piezas[variable].append(res)

        tmax_df = pd.concat(piezas['Tmax'], ignore_index=True) if piezas['Tmax'] else pd.DataFrame(columns=['ds', 'Tmax'])
        tmin_df = pd.concat(piezas['Tmin'], ignore_index=True) if piezas['Tmin'] else pd.DataFrame(columns=['ds', 'Tmin'])

        tabla = tmax_df.merge(tmin_df, on='ds', how='outer').sort_values('ds').reset_index(drop=True)
        if tabla.empty:
            continue

        tabla['Empresa'] = empresa
        tabla['Fundo']   = fundo
        tabla['Tipo']    = 'Predicho'
        tabla = tabla.rename(columns={'ds': 'Fecha'})
        filas.append(tabla[['Empresa', 'Fundo', 'Fecha', 'Tmax', 'Tmin', 'Tipo']])

    if not filas:
        return pd.DataFrame(columns=['Empresa', 'Fundo', 'Fecha', 'Tmax', 'Tmin', 'Tipo'])

    resultado = pd.concat(filas, ignore_index=True).sort_values(['Fundo', 'Fecha']).reset_index(drop=True)
    resultado['Fecha'] = resultado['Fecha'].dt.strftime('%Y/%m/%d')
    return resultado


def _validar_ventana(variable: str, dia: pd.DataFrame, ventana: dict, key_suffix: str) -> None:
    fundos       = dia['Fundo'].unique().tolist()
    fecha_inicio = ventana['inicio']
    fecha_fin    = ventana['fin']
    fecha_corte  = ventana['fecha_corte']
    datos_boxplot = []

    for fundo in fundos:
        sub = dia[dia['Fundo'] == fundo].copy().reset_index(drop=True)
        if sub.empty:
            continue

        real_periodo = sub[
            (sub['Fecha'] >= fecha_inicio) & (sub['Fecha'] <= fecha_fin)
        ].copy().reset_index(drop=True)

        if real_periodo.empty:
            continue

        train = sub[sub['Fecha'] <= fecha_corte].copy()
        if len(train) < 30:
            st.warning(f"⚠️ {fundo}: historial insuficiente")
            continue

        serie_train = train[['Fecha', variable]].rename(
            columns={'Fecha': 'ds', variable: 'y'}
        ).dropna()
        serie_train['ds'] = pd.to_datetime(serie_train['ds']).dt.tz_localize(None).dt.normalize()

        media = serie_train['y'].mean()
        std   = serie_train['y'].std()
        serie_train['y'] = serie_train['y'].clip(lower=media - 3*std, upper=media + 3*std)

        coef_clim_val = calcular_climatologia_armonica(
            serie_train[['ds', 'y']], PROPHET_PARAMS['n_harmonics'], PROPHET_PARAMS['clim_halflife_anios']
        )
        _val_fi = serie_train['ds'].min()
        _val_ff = fecha_fin + pd.Timedelta(days=30)
        _val_fechas = pd.date_range(_val_fi, _val_ff, freq='D')
        _val_clim_daily = pd.DataFrame({
            'ds'  : _val_fechas,
            'clim': predecir_climatologia_armonica(
                _val_fechas, coef_clim_val, PROPHET_PARAMS['n_harmonics']
            ).round(2)
        })
        serie_train = serie_train.merge(_val_clim_daily, on='ds', how='left')
        serie_train['y'] = (serie_train['y'] - serie_train['clim']).round(4)

        with st.spinner(f"Entrenando {fundo} {variable}..."):
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
                    m.fit(serie_train)
                    dias_pred_period = (fecha_fin - fecha_inicio).days + 1
                    future       = m.make_future_dataframe(periods=dias_pred_period, freq='D')
                    forecast_wf  = m.predict(future)
            except Exception as e:
                st.error(f"❌ Error {fundo}: {e}")
                continue

        forecast_wf['ds'] = pd.to_datetime(forecast_wf['ds']).dt.tz_localize(None).dt.normalize()
        forecast_wf = forecast_wf.merge(_val_clim_daily[['ds', 'clim']], on='ds', how='left')
        forecast_wf['clim'] = forecast_wf['clim'].ffill().bfill()
        for col in ['yhat', 'yhat_lower', 'yhat_upper']:
            forecast_wf[col] = (forecast_wf[col] + forecast_wf['clim']).round(2)

        pred_periodo = forecast_wf[
            (forecast_wf['ds'] >= fecha_inicio) & (forecast_wf['ds'] <= fecha_fin)
        ][['ds', 'yhat', 'yhat_lower', 'yhat_upper']].copy().reset_index(drop=True)

        if pred_periodo.empty:
            continue

        real_periodo = real_periodo.rename(columns={variable: 'Real'})
        comparacion  = real_periodo[['Fecha', 'Real']].merge(
            pred_periodo.rename(columns={
                'ds': 'Fecha', 'yhat': 'Pred',
                'yhat_lower': 'Pred_Low', 'yhat_upper': 'Pred_High',
            }),
            on='Fecha', how='inner'
        )

        if comparacion.empty:
            continue

        insample_val = forecast_wf[['ds', 'yhat']].merge(
            serie_train[['ds', 'y', 'clim']].assign(y=serie_train['y'] + serie_train['clim']),
            on='ds', how='inner'
        ).sort_values('ds').reset_index(drop=True)

        ventana_bias = min(PROPHET_PARAMS['bias_window_dias'] * 2, len(insample_val))
        ult = insample_val.tail(ventana_bias).copy().reset_index(drop=True)
        ult['residuo'] = ult['y'] - ult['yhat']
        ult['t'] = np.arange(len(ult))

        if len(ult) >= 5:
            b_slope, a_int = np.polyfit(ult['t'], ult['residuo'], 1)
        else:
            b_slope, a_int = 0.0, (ult['residuo'].mean() if len(ult) else 0.0)

        insample_val['anio']          = insample_val['ds'].dt.year
        insample_val['residuo_total'] = insample_val['y'] - insample_val['yhat']

        mes_objetivo   = fecha_inicio.month
        datos_mes_prev = insample_val[
            (insample_val['ds'].dt.month == mes_objetivo) &
            (insample_val['anio'] < fecha_inicio.year)
        ]
        if len(datos_mes_prev) >= 15:
            edad_anios = fecha_inicio.year - datos_mes_prev['anio']
            pesos = np.exp(-edad_anios / 1.5)
            bias_estacional = float(np.average(datos_mes_prev['residuo_total'], weights=pesos))
        else:
            bias_estacional = 0.0

        fecha_corte_dt = pd.Timestamp(fecha_corte)
        comparacion    = comparacion.copy()
        comparacion['h'] = (comparacion['Fecha'] - fecha_corte_dt).dt.days
        h_efectivo = np.minimum(comparacion['h'].clip(lower=0), H_MAX_PENDIENTE)

        comparacion['ajuste_enfen'] = comparacion['Fecha'].apply(
            lambda f: obtener_ajuste_enfen(f.month, f.year, variable)
        )
        comparacion['bias_h'] = (
            a_int + b_slope * (len(ult) - 1) + b_slope * h_efectivo
            + bias_estacional + comparacion['ajuste_enfen']
        )

        mbe_sin_correccion = (comparacion['Real'] - comparacion['Pred']).mean()
        tendencia_h        = a_int + b_slope * (len(ult) - 1) + b_slope * h_efectivo
        mbe_solo_tendencia = (comparacion['Real'] - (comparacion['Pred'] + tendencia_h)).mean()
        mbe_tend_estacional = (
            comparacion['Real'] - (comparacion['Pred'] + tendencia_h + bias_estacional)
        ).mean()

        st.caption(
            f"🩺 Diagnóstico MBE — sin corrección: {mbe_sin_correccion:+.2f}°C | "
            f"+tendencia: {mbe_solo_tendencia:+.2f}°C | "
            f"+estacional: {mbe_tend_estacional:+.2f}°C | "
            f"+ENFEN (final): -- (ver MBE abajo)"
        )
        st.caption(
            f"🔍 Componentes del bias — tendencia (a_int)={a_int:.2f}°C, "
            f"pendiente·h promedio={float((b_slope * h_efectivo).mean()):.2f}°C, "
            f"estacional={bias_estacional:.2f}°C, "
            f"ajuste ENFEN promedio={comparacion['ajuste_enfen'].mean():+.2f}°C "
            f"→ bias_h promedio={comparacion['bias_h'].mean():.2f}°C"
        )

        comparacion['Pred_corr']    = comparacion['Pred'] + comparacion['bias_h']
        comparacion['Error_abs']    = (comparacion['Real'] - comparacion['Pred_corr']).abs().round(2)
        comparacion['Error_signed'] = (comparacion['Real'] - comparacion['Pred_corr']).round(2)

        mae_mes        = comparacion['Error_abs'].mean().round(2)
        mbe_mes        = comparacion['Error_signed'].mean().round(2)
        rmse_mes       = round(float(np.sqrt((comparacion['Error_signed']**2).mean())), 2)
        dias_dentro    = int((comparacion['Error_abs'] <= 1.5).sum())
        dias_total     = len(comparacion)
        precision_mae  = round(dias_dentro / dias_total * 100, 2)
        dias_fuera     = dias_total - dias_dentro
        ratio_rmse_mae = round(rmse_mes / mae_mes, 2) if mae_mes > 0 else None

        std_residual = comparacion['Error_signed'].std()
        if std_residual > 0:
            precision_teorica = (
                norm.cdf(1.5, loc=0, scale=std_residual) -
                norm.cdf(-1.5, loc=0, scale=std_residual)
            ) * 100
        else:
            precision_teorica = 100.0

        brecha      = precision_teorica - precision_mae
        diagnostico = (
            "el resto es ruido irreducible — no lo resuelve calibración"
            if abs(mbe_mes) < 0.3 else
            "todavía hay bias corregible — calibración SÍ puede ayudar"
        )
        st.caption(
            f"📐 **Techo teórico (σ={std_residual:.2f}°C, bias=0):** "
            f"{precision_teorica:.0f}% — vs actual {precision_mae:.0f}% "
            f"(brecha {brecha:.0f} pts). {diagnostico}"
        )

        for _, r in comparacion.iterrows():
            datos_boxplot.append({'Fundo': fundo, 'Error': r['Error_signed']})

        if 'bias_correccion' not in st.session_state:
            st.session_state['bias_correccion'] = {}
        st.session_state['bias_correccion'][(fundo, variable)] = float(mbe_mes)

        # ── Métricas ──
        st.markdown(f"##### {fundo} — {ventana['label']}")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("MAE", f"{mae_mes:.2f}°C", help="Error promedio diario del mes")
        k2.metric("RMSE", f"{rmse_mes:.2f}°C",
                  delta=f"ratio {ratio_rmse_mae}x" if ratio_rmse_mae else None,
                  help="Si ratio > 1.3: hay días con errores grandes")
        k3.metric("BIAS (MBE)", f"{mbe_mes:+.2f}°C",
                  delta="subestima" if mbe_mes > 0 else "sobreestima",
                  delta_color="inverse" if mbe_mes > 0 else "normal",
                  help="Sesgo sistemático — se corrige automáticamente")
        k4.metric("Precisión (±1.5°C)", f"{precision_mae:.1f}%",
                  delta=f"{dias_dentro}/{dias_total} días dentro",
                  delta_color="normal" if precision_mae >= 60 else "inverse",
                  help="% de días con error dentro de ±1.5°C — umbral operacional")

        # ── Gráfico error ──
        fig_val = go.Figure()
        fig_val.add_hrect(y0=-1.5, y1=1.5, fillcolor='rgba(76, 175, 80, 0.10)', line_width=0)
        fig_val.add_hline(y=0, line=dict(color='#546E7A', width=1.5))

        for j in range(len(comparacion) - 1):
            x0 = comparacion['Fecha'].iloc[j]
            x1 = comparacion['Fecha'].iloc[j + 1]
            e0 = comparacion['Error_signed'].iloc[j]
            e1 = comparacion['Error_signed'].iloc[j + 1]

            if e0 >= 0 and e1 >= 0:
                fig_val.add_trace(go.Scatter(
                    x=[x0, x1, x1, x0], y=[e0, e1, 0, 0],
                    fill='toself', fillcolor='rgba(76,175,80,0.30)',
                    line=dict(color='rgba(0,0,0,0)'), showlegend=False, hoverinfo='skip'
                ))
            elif e0 < 0 and e1 < 0:
                fig_val.add_trace(go.Scatter(
                    x=[x0, x1, x1, x0], y=[e0, e1, 0, 0],
                    fill='toself', fillcolor='rgba(244,67,54,0.30)',
                    line=dict(color='rgba(0,0,0,0)'), showlegend=False, hoverinfo='skip'
                ))
            else:
                frac  = abs(e0) / (abs(e0) + abs(e1))
                rango = (x1 - x0).total_seconds()
                x_cross = x0 + pd.Timedelta(seconds=rango * frac)
                color_izq = 'rgba(76,175,80,0.30)' if e0 >= 0 else 'rgba(244,67,54,0.30)'
                fig_val.add_trace(go.Scatter(
                    x=[x0, x_cross, x_cross, x0], y=[e0, 0, 0, 0],
                    fill='toself', fillcolor=color_izq,
                    line=dict(color='rgba(0,0,0,0)'), showlegend=False, hoverinfo='skip'
                ))
                color_der = 'rgba(76,175,80,0.30)' if e1 >= 0 else 'rgba(244,67,54,0.30)'
                fig_val.add_trace(go.Scatter(
                    x=[x_cross, x1, x1, x_cross], y=[0, e1, 0, 0],
                    fill='toself', fillcolor=color_der,
                    line=dict(color='rgba(0,0,0,0)'), showlegend=False, hoverinfo='skip'
                ))

        fig_val.add_trace(go.Scatter(
            x=comparacion['Fecha'], y=comparacion['Error_signed'],
            mode='lines+markers',
            line=dict(color='#1A237E', width=2.5),
            marker=dict(size=7,
                        color=comparacion['Error_signed'].apply(lambda e: '#4CAF50' if e >= 0 else '#F44336'),
                        line=dict(color='white', width=1)),
            name='Error',
            hovertemplate='%{x|%d/%b}<br>Error: %{y:+.2f}°C<extra></extra>'
        ))
        fig_val.add_annotation(
            x=comparacion['Fecha'].iloc[len(comparacion) // 2],
            y=2.8,
            text=f"<b>MBE: {mbe_mes:+.2f}°C</b><br>{'🔺 subestima' if mbe_mes > 0 else '🔻 sobreestima'}<br>MAE: {mae_mes:.2f}°C",
            showarrow=False,
            bgcolor='rgba(255,255,255,0.85)', bordercolor='#546E7A', borderwidth=1,
            font=dict(size=10, color='#1A1A1A'), align='center'
        )
        fig_val.update_layout(
            height=350,
            title=dict(text=f"<b>Error diario — {fundo} — {variable}</b>", font=dict(size=12), x=0.0),
            xaxis=dict(tickformat='%d/%b', gridcolor='#CFD8DC'),
            yaxis=dict(title='Error (°C)', ticksuffix='°C', gridcolor='#CFD8DC',
                       range=[-3.5, 3.5], dtick=1, zeroline=False),
            legend=dict(orientation='h', y=-0.15, x=0),
            hovermode='x unified', paper_bgcolor='#FFFFFF',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(family='Arial', size=9),
            margin=dict(l=50, r=30, t=60, b=40),
        )
        st.plotly_chart(fig_val, use_container_width=True,
                        key=f"fig_val_{variable}_{fundo}{key_suffix}")
        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Box plot de errores ────────────────────────────────────
    if datos_boxplot:
        df_box = pd.DataFrame(datos_boxplot)
        fig_box = go.Figure()
        colores_box = ['#1565C0', '#C62828', '#2E7D32', '#F57F17', '#6A1B9A', '#00838F']

        for idx_f, fundo in enumerate(fundos):
            df_f  = df_box[df_box['Fundo'] == fundo]['Error']
            if df_f.empty:
                continue
            color = colores_box[idx_f % len(colores_box)]
            fig_box.add_trace(go.Box(
                y=df_f, name=fundo, boxpoints='all', jitter=0.4, pointpos=0,
                marker=dict(color=color, size=5, opacity=0.6, line=dict(color='white', width=0.5)),
                line=dict(color=color, width=2),
                fillcolor=f'rgba({int(color[1:3],16)},{int(color[3:5],16)},{int(color[5:7],16)},0.15)',
                hovertemplate=f'<b>{fundo}</b><br>Error: %{{y:+.2f}}°C<extra></extra>'
            ))

        fig_box.add_hrect(y0=-1.5, y1=1.5, fillcolor='rgba(76, 175, 80, 0.08)', line_width=0)
        fig_box.add_hline(y=1.5,  line=dict(color='#4CAF50', width=1.5, dash='dot'))
        fig_box.add_hline(y=-1.5, line=dict(color='#4CAF50', width=1.5, dash='dot'))
        fig_box.update_layout(
            height=420,
            title=dict(text=f"<b>📦 Distribución Error — {variable}</b>",
                       font=dict(size=13, color='#1A1A1A'), x=0.0),
            xaxis=dict(title='Fundo', gridcolor='#CFD8DC'),
            yaxis=dict(title='Error (°C)', ticksuffix='°C', gridcolor='#CFD8DC',
                       tickvals=[-3, -2, -1, 0, 1, 2, 3], range=[-3.5, 3.5]),
            paper_bgcolor='#FFFFFF', plot_bgcolor='rgba(0,0,0,0)',
            font=dict(family='Arial', size=9, color='#333333'),
            showlegend=False, margin=dict(l=60, r=40, t=60, b=40),
        )
        st.plotly_chart(fig_box, use_container_width=True,
                        key=f"fig_box_{variable}{key_suffix}")

        # ── Gráfico apilado de rangos ──────────────────────────
        def categorizar_error(e):
            if e < -1.5: return '< -1.5°C'
            elif e < 0:  return '-1.5 a 0°C'
            elif e <= 1.5: return '0 a 1.5°C'
            else: return '> 1.5°C'

        df_box['Categoria']   = df_box['Error'].apply(categorizar_error)
        categorias_orden      = ['< -1.5°C', '-1.5 a 0°C', '0 a 1.5°C', '> 1.5°C']
        colores_cat = {
            '< -1.5°C': '#E53935', '-1.5 a 0°C': '#FFAB40',
            '0 a 1.5°C': '#66BB6A', '> 1.5°C': '#1E88E5',
        }

        tabla_pct = df_box.groupby(['Fundo', 'Categoria']).size().reset_index(name='N')
        total_por_fundo     = tabla_pct.groupby('Fundo')['N'].transform('sum')
        tabla_pct['Pct']    = (tabla_pct['N'] / total_por_fundo * 100).round(1)
        fundos_orden        = sorted(df_box['Fundo'].unique().tolist())

        fig_stack = go.Figure()
        for cat in categorias_orden:
            df_cat = tabla_pct[tabla_pct['Categoria'] == cat]
            pcts, ns = [], []
            for fundo in fundos_orden:
                row_cat = df_cat[df_cat['Fundo'] == fundo]
                if not row_cat.empty:
                    pcts.append(float(row_cat['Pct'].iloc[0]))
                    ns.append(int(row_cat['N'].iloc[0]))
                else:
                    pcts.append(0.0); ns.append(0)

            fig_stack.add_trace(go.Bar(
                name=cat, x=fundos_orden, y=pcts, marker_color=colores_cat[cat],
                text=[f'{p:.1f}%' if p > 4 else '' for p in pcts],
                textposition='inside', textfont=dict(size=11, color='white', family='Arial'),
                customdata=ns,
                hovertemplate='<b>%{x}</b><br>' + cat + '<br>Días: %{customdata}<br>%{y:.1f}%<extra></extra>'
            ))

        fig_stack.update_layout(
            barmode='stack', height=380,
            title=dict(text=f"<b>📊 Rangos Error — {variable}</b>", font=dict(size=13), x=0.0),
            xaxis_title='Fundo',
            yaxis=dict(title='% de días', ticksuffix='%', range=[0, 100],
                       gridcolor='#CFD8DC', dtick=25),
            legend=dict(orientation='h', y=1.02, x=0),
            paper_bgcolor='#FFFFFF', plot_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=60, r=40, t=80, b=40),
        )
        st.plotly_chart(fig_stack, use_container_width=True,
                        key=f"fig_stack_{variable}{key_suffix}")


def generar_seccion_validacion(variable: str, dia: pd.DataFrame,
                               forecasts_cache: dict, dias_pred_ui: int) -> None:
    import calendar

    fecha_max           = pd.to_datetime(dia['Fecha'].max()).normalize()
    ultimo_dia_mes_ant  = fecha_max.replace(day=1) - pd.Timedelta(days=1)
    primer_dia_mes_ant  = ultimo_dia_mes_ant.replace(day=1)
    fecha_corte_1mes    = primer_dia_mes_ant - pd.Timedelta(days=1)
    fecha_inicio_3m     = primer_dia_mes_ant.replace(day=1) - pd.DateOffset(months=2)
    fecha_corte_3m      = fecha_inicio_3m - pd.Timedelta(days=1)

    ventanas = {
        '1mes': {
            'label'      : f"{primer_dia_mes_ant.strftime('%B %Y')}",
            'inicio'     : primer_dia_mes_ant,
            'fin'        : ultimo_dia_mes_ant,
            'fecha_corte': fecha_corte_1mes,
            'dias'       : (ultimo_dia_mes_ant - primer_dia_mes_ant).days + 1,
        },
        '3mes': {
            'label'      : f"{fecha_inicio_3m.strftime('%b %Y')} → {ultimo_dia_mes_ant.strftime('%b %Y')}",
            'inicio'     : fecha_inicio_3m,
            'fin'        : ultimo_dia_mes_ant,
            'fecha_corte': fecha_corte_3m,
            'dias'       : (ultimo_dia_mes_ant - fecha_inicio_3m).days + 1,
        },
    }

    with st.expander(f"📊 Validación Prophet — {variable} (walk-forward honesto)", expanded=False):
        tab_1mes, tab_3mes = st.tabs(["Último mes", "Últimos 3 meses"])

        with tab_1mes:
            st.info(
                f"Entrenamiento: datos hasta {ventanas['1mes']['fecha_corte'].strftime('%d/%b/%Y')} | "
                f"Predicción: {ventanas['1mes']['dias']} días"
            )
            _validar_ventana(variable, dia, ventanas['1mes'], '_1mes')

        with tab_3mes:
            st.info(
                f"Entrenamiento: datos hasta {ventanas['3mes']['fecha_corte'].strftime('%d/%b/%Y')} | "
                f"Predicción: {ventanas['3mes']['dias']} días"
            )
            _validar_ventana(variable, dia, ventanas['3mes'], '_3mes')

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
    with st.expander("🔬 Calibración de hiperparámetros (grid search)", expanded=False):
        st.caption(
            "Reentrena cada combinación para cada fundo/variable. "
            "El grid de climatología (n_harmonics/halflife) puede tardar más en '3 meses'."
        )
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            ventana_grid_sel = st.radio(
                "Ventana", options=['1mes', '3mes'], horizontal=True,
                key=f"grid_ventana_{variable}"
            )
        with col_g2:
            grid_tipo = st.radio(
                "Tipo de grid",
                options=['Hiperparámetros Prophet', 'Climatología (n_harmonics/halflife)'],
                horizontal=True, key=f"grid_tipo_{variable}"
            )

        if st.button(f"Ejecutar grid search — {variable}", key=f"grid_btn_{variable}"):
            if grid_tipo == 'Hiperparámetros Prophet':
                grid = {
                    'changepoint_prior_scale': [0.05, 0.15, 0.30],
                    'seasonality_prior_scale': [1.0, 2.0, 5.0],
                }
            else:
                grid = {
                    'n_harmonics'      : [2, 3, 4, 6],
                    'clim_halflife_anios': [1.0, 1.5, 2.0, 3.0],
                }

            df_grid = grid_search_hiperparams(dia, ventanas[ventana_grid_sel], grid)

            if df_grid.empty:
                st.warning("Sin resultados — revisa que haya suficiente historial.")
            else:
                st.markdown("##### 📋 Resultados completos")
                st.dataframe(
                    df_grid.sort_values(['Fundo', 'Variable', 'Precision'],
                                        ascending=[True, True, False]),
                    use_container_width=True
                )
                mejores = df_grid.loc[df_grid.groupby(['Fundo', 'Variable'])['Precision'].idxmax()]
                st.markdown("##### 🏆 Mejor combo por fundo/variable")
                st.dataframe(mejores, use_container_width=True)
                st.session_state['grid_search_resultado'] = df_grid
