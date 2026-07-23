import calendar

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots
from scipy.ndimage import gaussian_filter1d

from config.config import (
    GRID_COLOR, ROLLING_DIAS, Z_Q,
    BG_TMAX, BANDA_TMAX, CLIM_TMAX, REAL_TMAX_COLOR, PRED_COLOR_TMAX,
    BG_TMIN, BANDA_TMIN, CLIM_TMIN, REAL_TMIN_COLOR, PRED_COLOR_TMIN,
    BANDA2_TMAX, BANDA2_TMIN, DELTA_Q, MESES_RAW,
)
from services.climatologia_service import spline_diario_cached
from services.utils import _hex_to_rgba


def _fin_mes_prediccion(fecha_corte_real: pd.Timestamp) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Primer y último día del mes a pronosticar.

    Si fecha_corte_real ya es el último día de su mes (mes cerrado), avanza al
    mes siguiente completo; de lo contrario la ventana [corte+1, fin_mes] queda
    invertida y la predicción sale vacía.
    """
    primer_dia = fecha_corte_real.replace(day=1)
    ultimo_dia = pd.Timestamp(
        year=primer_dia.year, month=primer_dia.month,
        day=calendar.monthrange(primer_dia.year, primer_dia.month)[1]
    )
    if fecha_corte_real >= ultimo_dia:
        primer_dia = (primer_dia + pd.DateOffset(months=1)).replace(day=1)
        ultimo_dia = pd.Timestamp(
            year=primer_dia.year, month=primer_dia.month,
            day=calendar.monthrange(primer_dia.year, primer_dia.month)[1]
        )
    return primer_dia, ultimo_dia


def ajustar_prediccion_patron(pred_fc: pd.DataFrame, sub: pd.DataFrame,
                              variable: str) -> pd.DataFrame:
    """Inyecta el patrón de residuos recientes (real vs suavizado) sobre una
    predicción cruda de Prophet. Misma lógica que se dibuja en el gráfico,
    factorizada para que cualquier export coincida con él.

    No aplica la corrección BIAS de session_state: ese valor lo recalcula
    _validar_ventana en cada rerun de Streamlit (incluyendo clics sin relación),
    lo que hacía que la predicción ya mostrada cambiara sola entre una corrida y
    otra. La predicción queda fija una vez calculada — solo cambia si cambian los
    datos reales (nuevo fecha_max) o se reentrena el caché de Prophet.
    """
    if pred_fc.empty or len(sub) < 14:
        return pred_fc

    ventana      = min(21, len(sub))
    sub_reciente = sub.tail(ventana).copy().reset_index(drop=True)
    residuos     = (sub_reciente[variable] - sub_reciente[f'{variable}_smooth']).values
    residuos_suaves = gaussian_filter1d(residuos, sigma=0.8)

    n_pred  = len(pred_fc)
    n_res   = len(residuos_suaves)
    indices = [ii % n_res for ii in range(n_pred)]
    patron  = residuos_suaves[indices]

    std_hist = residuos_suaves.std()
    std_pred = pred_fc['yhat'].std()
    factor   = min(1.0, std_hist / (std_pred + 1e-6)) * 0.6

    pred_fc = pred_fc.copy()
    cols_ajustar = [c for c in ('yhat', 'yhat_lower', 'yhat_upper') if c in pred_fc.columns]
    for col in cols_ajustar:
        pred_fc[col] = (pred_fc[col] + patron * factor).round(2)

    return pred_fc


def generar_range_plot(variable: str, historico_df: pd.DataFrame, prediccion_df: pd.DataFrame,
                       fundo_name: str, row: int, fig) -> None:
    """Agrega Range Plot con Error Bars (barras Tmin-Tmax)."""
    if not historico_df.empty:
        fechas  = historico_df['Fecha'].tolist()
        tmax    = historico_df['Tmax'].tolist()
        tmin    = historico_df['Tmin'].tolist()
        smooth  = historico_df[f'{variable}_smooth'].tolist()
        et      = historico_df['ET'].tolist()

        error_plus  = [t - s for t, s in zip(tmax, smooth)]
        error_minus = [s - t for t, s in zip(tmin, smooth)]

        fig.add_trace(go.Scatter(
            x=fechas, y=smooth,
            mode='lines+markers',
            line=dict(color='#000000', width=2),
            marker=dict(size=5, color='#00AA00', line=dict(color='#000000', width=1)),
            error_y=dict(type='data', symmetric=False, array=error_plus,
                         arrayminus=error_minus, color='#00AA00', thickness=2, width=3),
            name=f'{variable} REAL', showlegend=True, legendgroup=fundo_name,
            hovertemplate='%{x|%d/%b/%Y}<br>' + f'{variable}: %{{y:.1f}}°C<br>ET: %{{customdata:.1f}}<extra>{fundo_name}</extra>',
            customdata=et
        ), row=row, col=1)

    if not prediccion_df.empty:
        fechas_pred    = prediccion_df['ds'].tolist()
        yhat           = prediccion_df['yhat'].tolist()
        upper          = prediccion_df['yhat_upper'].tolist()
        lower          = prediccion_df['yhat_lower'].tolist()
        error_plus_pred  = [u - y for u, y in zip(upper, yhat)]
        error_minus_pred = [y - l for y, l in zip(yhat, lower)]

        fig.add_trace(go.Scatter(
            x=fechas_pred, y=yhat,
            mode='lines+markers',
            line=dict(color='#FF0000', width=2),
            marker=dict(size=5, color='#FF9800', line=dict(color='#FF0000', width=1)),
            error_y=dict(type='data', symmetric=False, array=error_plus_pred,
                         arrayminus=error_minus_pred, color='#FF9800', thickness=2, width=3),
            name=f'Predicción +{len(fechas_pred)}d', showlegend=True, legendgroup=fundo_name,
            hovertemplate='%{x|%d/%b/%Y}<br>Pred: %{y:.1f}°C<extra>Prophet</extra>'
        ), row=row, col=1)


def generar_figura(variable: str, dia: pd.DataFrame,
                   media_mensual, q1_mensual, q3_mensual,
                   dias_pred: int, forecasts_cache: dict, dias_vista: int = 30,
                   dias_pred_mostrar: int | None = None, tipo_viz: str = 'linea',
                   media_mensual2=None, q1_mensual2=None, q3_mensual2=None,
                   label_clim2: str | None = None):
    """Genera figura Plotly con datos reales + predicción Prophet + climatología."""
    fundos = dia['Fundo'].unique().tolist()
    n      = len(fundos)

    if variable == 'Tmax':
        fig_bg, banda_color, clim_color, color_real, pred_color = (
            BG_TMAX, BANDA_TMAX, CLIM_TMAX, REAL_TMAX_COLOR, PRED_COLOR_TMAX
        )
    else:
        fig_bg, banda_color, clim_color, color_real, pred_color = (
            BG_TMIN, BANDA_TMIN, CLIM_TMIN, REAL_TMIN_COLOR, PRED_COLOR_TMIN
        )

    fig = make_subplots(rows=n, cols=1, shared_xaxes=False,
                        vertical_spacing=0.08, subplot_titles=fundos)

    rows_export, rows_pred_export = [], []
    anio_actual   = pd.Timestamp.today().year
    fecha_ini_anio = pd.Timestamp(year=anio_actual - 1, month=1, day=1)

    for i, fundo in enumerate(fundos):
        row = i + 1

        sub_full = dia[dia['Fundo'] == fundo].copy().reset_index(drop=True)
        if sub_full.empty:
            continue

        sub_full = sub_full[sub_full['Fecha'] >= fecha_ini_anio].copy().reset_index(drop=True)
        if sub_full.empty:
            st.warning(f"⚠️ Sin datos en {anio_actual} para **{fundo}**.")
            continue

        fecha_corte_real      = pd.to_datetime(sub_full['Fecha'].max()).tz_localize(None).normalize()
        primer_dia_mes_actual, ultimo_dia_mes_actual = _fin_mes_prediccion(fecha_corte_real)

        sub    = sub_full.copy().reset_index(drop=True)
        if sub.empty:
            continue

        empresa  = sub['Empresa'].iloc[0]
        zoom_ini = sub['Fecha'].min()
        zoom_fin = ultimo_dia_mes_actual + pd.Timedelta(days=2)

        fecha_ini_clima = sub['Fecha'].min()
        if hasattr(fecha_ini_clima, 'tz_localize'):
            fecha_ini_clima = pd.Timestamp(fecha_ini_clima).tz_localize(None).normalize()
        fecha_fin_clima = pd.Timestamp(zoom_fin)

        clim_media = spline_diario_cached(
            fecha_ini_clima.strftime('%Y-%m-%d'), fecha_fin_clima.strftime('%Y-%m-%d'), tuple(media_mensual)
        )
        clim_q1 = spline_diario_cached(
            fecha_ini_clima.strftime('%Y-%m-%d'), fecha_fin_clima.strftime('%Y-%m-%d'), tuple(q1_mensual)
        )
        clim_q3 = spline_diario_cached(
            fecha_ini_clima.strftime('%Y-%m-%d'), fecha_fin_clima.strftime('%Y-%m-%d'), tuple(q3_mensual)
        )

        hist_df = sub[['Empresa', 'Fundo', 'Fecha', variable, f'{variable}_smooth']].copy()
        hist_df = hist_df.merge(clim_media.rename(columns={'Valor': 'Clim_MEDIA'}), on='Fecha', how='left')
        hist_df = hist_df.merge(clim_q1.rename(columns={'Valor': 'Clim_Q1'}), on='Fecha', how='left')
        hist_df = hist_df.merge(clim_q3.rename(columns={'Valor': 'Clim_Q3'}), on='Fecha', how='left')
        hist_df['Variable'] = variable
        rows_export.append(hist_df)

        cache_entry = forecasts_cache.get((fundo, variable), {})
        forecast    = cache_entry.get('forecast', pd.DataFrame())
        mae         = cache_entry.get('mae', None)
        mae_por_dia = cache_entry.get('mae_por_dia', {})

        if not forecast.empty:
            forecast = forecast.copy()
            forecast['ds'] = pd.to_datetime(forecast['ds']).dt.tz_localize(None).dt.normalize()
            primer_dia_pred = fecha_corte_real + pd.Timedelta(days=1)
            pred_fc = forecast[
                (forecast['ds'] >= primer_dia_pred) & (forecast['ds'] <= ultimo_dia_mes_actual)
            ].reset_index(drop=True)
        else:
            pred_fc = pd.DataFrame()

        pred_fc = ajustar_prediccion_patron(pred_fc, sub, variable)

        lg = fundo

        # ── Climatología Guadalupe ─────────────────────────────
        fig.add_trace(go.Scatter(
            x=clim_q3['Fecha'], y=clim_q3['Valor'], mode='lines', line=dict(width=0),
            showlegend=False, hoverinfo='skip', legendgroup=lg
        ), row=row, col=1)
        fig.add_trace(go.Scatter(
            x=clim_q1['Fecha'], y=clim_q1['Valor'], mode='lines', line=dict(width=0),
            fill='tonexty', fillcolor=_hex_to_rgba(banda_color, 0.55),
            showlegend=(i == 0), legendgroup=lg,
            name=f'Q1-Q3 SENAMHI Guadalupe (±{DELTA_Q}°C)',
            hovertemplate='%{x|%d/%b}<br>Q1: %{y:.1f}°C<extra></extra>'
        ), row=row, col=1)
        fig.add_trace(go.Scatter(
            x=clim_media['Fecha'], y=clim_media['Valor'], mode='lines',
            line=dict(color=clim_color, width=2.8),
            showlegend=(i == 0), legendgroup=lg,
            name=f'{variable} climatología SENAMHI Guadalupe',
            hovertemplate='%{x|%d/%b}<br>Clim Guadalupe: %{y:.1f}°C<extra></extra>'
        ), row=row, col=1)

        # ── Segunda climatología dinámica ──────────────────────
        if media_mensual2 is not None and label_clim2:
            clim_media2 = spline_diario_cached(
                fecha_ini_clima.strftime('%Y-%m-%d'), fecha_fin_clima.strftime('%Y-%m-%d'), tuple(media_mensual2)
            )
            clim_q1_2 = spline_diario_cached(
                fecha_ini_clima.strftime('%Y-%m-%d'), fecha_fin_clima.strftime('%Y-%m-%d'), tuple(q1_mensual2)
            )
            clim_q3_2 = spline_diario_cached(
                fecha_ini_clima.strftime('%Y-%m-%d'), fecha_fin_clima.strftime('%Y-%m-%d'), tuple(q3_mensual2)
            )
            fig.add_trace(go.Scatter(
                x=clim_q3_2['Fecha'], y=clim_q3_2['Valor'], mode='lines', line=dict(width=0),
                showlegend=False, hoverinfo='skip', legendgroup=lg
            ), row=row, col=1)
            fig.add_trace(go.Scatter(
                x=clim_q1_2['Fecha'], y=clim_q1_2['Valor'], mode='lines', line=dict(width=0),
                fill='tonexty',
                fillcolor=_hex_to_rgba(BANDA2_TMAX if variable == 'Tmax' else BANDA2_TMIN, 0.15),
                showlegend=(i == 0), legendgroup=lg,
                name=f'Q1-Q3 SENAMHI {label_clim2} (±{DELTA_Q}°C)',
                hovertemplate='%{x|%d/%b}<br>Q1: %{y:.1f}°C<extra></extra>'
            ), row=row, col=1)
            fig.add_trace(go.Scatter(
                x=clim_media2['Fecha'], y=clim_media2['Valor'], mode='lines',
                line=dict(color='rgba(255,111,0,0.70)', width=1.8, dash='dot'),
                showlegend=(i == 0), legendgroup=lg,
                name=f'{variable} climatología SENAMHI {label_clim2}',
                hovertemplate='%{x|%d/%b}<br>Clim ' + label_clim2 + ': %{y:.1f}°C<extra></extra>'
            ), row=row, col=1)

        # ── Datos reales ───────────────────────────────────────
        if tipo_viz == 'candlestick':
            if not pred_fc.empty:
                if dias_pred_mostrar is not None and dias_pred_mostrar < len(pred_fc):
                    pred_fc = pred_fc.iloc[:dias_pred_mostrar].copy()
                generar_range_plot(variable, sub, pred_fc, fundo, row, fig)
            else:
                generar_range_plot(variable, sub, pd.DataFrame(), fundo, row, fig)
        else:
            fig.add_trace(go.Scatter(
                x=sub['Fecha'], y=sub[f'{variable}_smooth'],
                mode='lines+markers',
                line=dict(color=color_real, width=1.5),
                marker=dict(size=5, color='white', line=dict(color=color_real, width=1.0)),
                customdata=sub['ET'], showlegend=(i == 0), legendgroup=lg,
                name=f'{variable} REAL',
                hovertemplate='%{x|%d/%b/%Y}<br>' + f'{variable}: %{{y:.1f}}°C<br>ET: %{{customdata:.1f}}<extra>{fundo}</extra>'
            ), row=row, col=1)

        ult          = sub.iloc[-1]
        ult_real_val = ult[variable]
        fig.add_annotation(
            x=ult['Fecha'], y=ult_real_val,
            text=f"<b>{ult_real_val:.1f}°C</b>",
            showarrow=True, arrowhead=2, arrowcolor=color_real, arrowwidth=1.0,
            ax=14, ay=-18, font=dict(size=9, color=color_real),
            row=row, col=1
        )
        fig.add_vline(x=fecha_corte_real, line=dict(color='#546E7A', width=1.2, dash='dot'),
                      row=row, col=1)

        # ── Predicciones modo línea ────────────────────────────
        if tipo_viz == 'linea' and not pred_fc.empty:
            if dias_pred_mostrar is not None and dias_pred_mostrar < len(pred_fc):
                pred_fc = pred_fc.iloc[:dias_pred_mostrar].copy()

            ult_smooth = float(ult[f'{variable}_smooth'])
            fig.add_trace(go.Scatter(
                x=[ult['Fecha'], pred_fc['ds'].iloc[0]],
                y=[ult_smooth, float(pred_fc['yhat'].iloc[0])],
                mode='lines',
                line=dict(
                    color=f'rgba({int(pred_color[1:3],16)},{int(pred_color[3:5],16)},{int(pred_color[5:7],16)},0.70)',
                    width=1.5, dash='dot'
                ),
                showlegend=False, hoverinfo='skip', legendgroup=lg
            ), row=row, col=1)

            fig.add_trace(go.Scatter(
                x=pred_fc['ds'].tolist(), y=pred_fc['yhat_lower'].tolist(),
                mode='lines', line=dict(width=0, color='rgba(0,0,0,0)'),
                showlegend=False, hoverinfo='skip', legendgroup=lg, name='_lower'
            ), row=row, col=1)
            fig.add_trace(go.Scatter(
                x=pred_fc['ds'].tolist(), y=pred_fc['yhat_upper'].tolist(),
                mode='lines', line=dict(width=0, color='rgba(0,0,0,0)'),
                fill='tonexty', fillcolor=_hex_to_rgba(pred_color, 0.20),
                showlegend=(i == 0), name='IC 95%', legendgroup=lg, hoverinfo='skip'
            ), row=row, col=1)
            fig.add_trace(go.Scatter(
                x=pred_fc['ds'].tolist(), y=pred_fc['yhat'].tolist(),
                mode='lines',
                line=dict(
                    color=f'rgba({int(pred_color[1:3],16)},{int(pred_color[3:5],16)},{int(pred_color[5:7],16)},0.70)',
                    width=1.8, dash='dot'
                ),
                showlegend=(i == 0), legendgroup=lg,
                name='Predicción mes siguiente',
                hovertemplate='%{x|%d/%b/%Y}<br>Pred: %{y:.1f}°C<extra>Predicción</extra>'
            ), row=row, col=1)

            ult_pred_dt  = pred_fc['ds'].iloc[-1]
            ult_pred_val = float(pred_fc['yhat'].iloc[-1])
            fig.add_annotation(
                x=ult_pred_dt, y=ult_pred_val,
                text=f"<b>{ult_pred_val:.1f}°C</b>",
                showarrow=True, arrowhead=2, arrowcolor=pred_color, arrowwidth=1.0,
                ax=-28, ay=-18, font=dict(size=9, color=pred_color),
                row=row, col=1
            )

        bg_color = '#C62828' if variable == 'Tmax' else '#1565C0'
        fig.layout.annotations[i].update(
            text=f"  <b>{fundo}</b>  ",
            font=dict(size=12, color='#FFFFFF'),
            bgcolor=bg_color, borderpad=4,
            bordercolor=bg_color, borderwidth=2,
            x=0.0, xanchor='left'
        )

        xk = 'xaxis' if row == 1 else f'xaxis{row}'
        fig.layout[xk].update(
            range=[zoom_ini, zoom_fin], tickformat='%d/%b', ticklabelmode='period',
            gridcolor=GRID_COLOR, gridwidth=0.5, showgrid=True, zeroline=False,
            autorange=False, fixedrange=False, rangeslider=dict(visible=False)
        )
        yk = 'yaxis' if row == 1 else f'yaxis{row}'
        fig.layout[yk].update(ticksuffix='°C', gridcolor=GRID_COLOR, gridwidth=0.5,
                               showgrid=True, zeroline=False)

    fig.update_layout(
        height=320 * n,
        paper_bgcolor=fig_bg,
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Arial', size=9, color='#333333'),
        title=dict(text="", font=dict(size=14, color='#1A1A1A'),
                   x=0.0, xanchor='left', pad=dict(t=10, l=10)),
        margin=dict(l=60, r=40, t=60, b=40),
        legend=dict(orientation='h', yanchor='top', y=-0.08, xanchor='left', x=0,
                    font=dict(size=8), bgcolor='rgba(0,0,0,0)', borderwidth=0),
        hovermode='x unified'
    )

    df_hist = pd.concat(rows_export, ignore_index=True) if rows_export else pd.DataFrame()
    df_pred = pd.concat(rows_pred_export, ignore_index=True) if rows_pred_export else pd.DataFrame()
    return fig, df_hist, df_pred


def generar_tab_et(dia: pd.DataFrame, forecasts_cache: dict, dias_pred_ui: int,
                   modulos_kmz=None, riesgos_et_pre: dict | None = None,
                   df_et_mensual: pd.DataFrame | None = None) -> None:
    """Renderiza el tab completo de Evapotranspiración."""
    from services.map_service import (
        calcular_et_suma_semanal_fundos, crear_colormap_y_leyenda_et, generar_mapa_et
    )

    fundos         = dia['Fundo'].unique().tolist()
    anio_actual    = pd.Timestamp.today().year
    fecha_ini_anio = pd.Timestamp(year=anio_actual - 1, month=1, day=1)
    n              = len(fundos)
    altura_et      = 300 * n

    col_graf_et, col_mapa_et = st.columns([3, 1])

    fig = make_subplots(rows=n, cols=1, shared_xaxes=False,
                        vertical_spacing=0.08, subplot_titles=fundos)

    for i, fundo in enumerate(fundos):
        row = i + 1
        sub = dia[
            (dia['Fundo'] == fundo) & (dia['Fecha'] >= fecha_ini_anio)
        ].copy().reset_index(drop=True)

        if sub.empty or 'ET' not in sub.columns:
            continue

        et_df = sub[['Fecha', 'ET']].dropna().copy()
        if et_df.empty or et_df['ET'].sum() == 0:
            continue

        fecha_corte_real      = pd.to_datetime(et_df['Fecha'].max()).normalize()
        primer_dia_mes_actual, ultimo_dia_mes_actual = _fin_mes_prediccion(fecha_corte_real)

        et_real = et_df[et_df['Fecha'] <= fecha_corte_real].copy().reset_index(drop=True)
        et_real['ET_smooth'] = (
            et_real['ET'].rolling(ROLLING_DIAS, center=True, min_periods=1).mean().round(2)
        )

        # ── Climatología ET diaria ─────────────────────────────
        if df_et_mensual is not None and not df_et_mensual.empty:
            sub_et_clim = (
                df_et_mensual[df_et_mensual['Fundo'] == fundo]
                .sort_values('Mes').reset_index(drop=True)
            )
            if len(sub_et_clim) == 12:
                valores_et_clim  = sub_et_clim['EToPromedioDiaria'].values.tolist()
                delta_et         = round(np.std(valores_et_clim) * Z_Q, 2)
                fecha_ini_et_str = pd.Timestamp(et_real['Fecha'].min()).strftime('%Y-%m-%d')
                fecha_fin_et_str = (ultimo_dia_mes_actual + pd.Timedelta(days=2)).strftime('%Y-%m-%d')

                clim_et_media = spline_diario_cached(fecha_ini_et_str, fecha_fin_et_str, tuple(valores_et_clim))
                clim_et_q1    = spline_diario_cached(fecha_ini_et_str, fecha_fin_et_str,
                                                      tuple([max(0.0, v - delta_et) for v in valores_et_clim]))
                clim_et_q3    = spline_diario_cached(fecha_ini_et_str, fecha_fin_et_str,
                                                      tuple([v + delta_et for v in valores_et_clim]))

                fig.add_trace(go.Scatter(
                    x=clim_et_q3['Fecha'], y=clim_et_q3['Valor'],
                    mode='lines', line=dict(width=0),
                    showlegend=False, hoverinfo='skip'
                ), row=row, col=1)
                fig.add_trace(go.Scatter(
                    x=clim_et_q1['Fecha'], y=clim_et_q1['Valor'],
                    mode='lines', line=dict(width=0),
                    fill='tonexty', fillcolor='rgba(2, 136, 209, 0.12)',
                    showlegend=(i == 0),
                    name=f'Banda ET histórica (±{delta_et:.2f} mm/día)', hoverinfo='skip'
                ), row=row, col=1)
                fig.add_trace(go.Scatter(
                    x=clim_et_media['Fecha'], y=clim_et_media['Valor'],
                    mode='lines', line=dict(color='#0277BD', width=2.2),
                    showlegend=(i == 0), name='ET climatología histórica',
                    hovertemplate='%{x|%d/%b}<br>ET clim: %{y:.2f} mm/día<extra></extra>'
                ), row=row, col=1)

        # ── Línea real ─────────────────────────────────────────
        fig.add_trace(go.Scatter(
            x=et_real['Fecha'], y=et_real['ET_smooth'],
            mode='lines+markers',
            line=dict(color='#000000', width=1.5),
            marker=dict(size=5, color='white', line=dict(color='#000000', width=1.0)),
            customdata=et_real['ET'], name='ET real', showlegend=(i == 0),
            hovertemplate='%{x|%d/%b/%Y}<br>ET: %{customdata:.2f} mm<br>ET suav: %{y:.2f} mm<extra>' + fundo + '</extra>'
        ), row=row, col=1)

        # ── Cuartiles horizontales ─────────────────────────────
        if not et_real.empty:
            q1 = et_real['ET_smooth'].quantile(0.25)
            q3 = et_real['ET_smooth'].quantile(0.75)
            fig.add_hline(y=q1, line=dict(color='#B3E5FC', width=1, dash='dash'),
                          annotation_text=f"Q1: {q1:.2f}", annotation_position="right",
                          annotation_font=dict(size=7, color='#0288D1'), row=row, col=1)
            fig.add_hline(y=q3, line=dict(color='#01579B', width=1, dash='dash'),
                          annotation_text=f"Q3: {q3:.2f}", annotation_position="right",
                          annotation_font=dict(size=7, color='#01579B'), row=row, col=1)

        if not et_real.empty:
            fig.add_annotation(
                x=et_real['Fecha'].iloc[-1], y=float(et_real['ET'].iloc[-1]),
                text=f"<b>{float(et_real['ET'].iloc[-1]):.2f}</b>",
                showarrow=True, arrowhead=2, arrowcolor='#000000', arrowwidth=1.0,
                ax=14, ay=-18, font=dict(size=9, color='#000000'), row=row, col=1
            )

        fig.add_vline(x=fecha_corte_real, line=dict(color='#546E7A', width=1.2, dash='dot'),
              row=row, col=1)

        # ── Predicción Prophet ET ──────────────────────────────
        cache_entry = forecasts_cache.get((fundo, 'ET'), {})
        forecast_et = cache_entry.get('forecast', pd.DataFrame())

        if not forecast_et.empty:
            forecast_et = forecast_et.copy()
            forecast_et['ds'] = pd.to_datetime(forecast_et['ds']).dt.tz_localize(None).dt.normalize()
            primer_dia_pred = fecha_corte_real + pd.Timedelta(days=1)
            pred_et = forecast_et[
                (forecast_et['ds'] >= primer_dia_pred) &
                (forecast_et['ds'] <= ultimo_dia_mes_actual)
            ].reset_index(drop=True)

            mes_actual_num = primer_dia_mes_actual.month
            et_mismo_mes   = et_df[pd.to_datetime(et_df['Fecha']).dt.month == mes_actual_num]['ET']
            if len(et_mismo_mes) >= 5:
                std_historica = et_mismo_mes.std()
                amp_max       = et_mismo_mes.max()
            else:
                std_historica = et_df['ET'].std()
                amp_max       = et_df['ET'].max()

            if not pred_et.empty and not et_real.empty:
                ventana_patron = min(30, len(et_real))
                et_ventana     = et_real['ET'].values[-ventana_patron:]
                desviacion     = et_ventana - et_ventana.mean()
                n_pred         = len(pred_et)
                patron_ciclado = desviacion[[j % len(desviacion) for j in range(n_pred)]]
                std_actual     = np.std(desviacion)
                factor         = min(std_historica / (std_actual + 1e-6), 1.5)

                pred_et = pred_et.copy()
                pred_et['yhat']       = (pred_et['yhat'] + patron_ciclado * factor).clip(lower=0).round(2)
                pred_et['yhat_upper'] = (pred_et['yhat'] + std_historica * 0.8).clip(upper=float(amp_max)).round(2)
                pred_et['yhat_lower'] = (pred_et['yhat'] - std_historica * 0.8).clip(lower=0.0).round(2)

                fig.add_trace(go.Scatter(
                    x=[et_real['Fecha'].iloc[-1], pred_et['ds'].iloc[0]],
                    y=[float(et_real['ET_smooth'].iloc[-1]), float(pred_et['yhat'].iloc[0])],
                    mode='lines', line=dict(color='#43A047', width=1.8, dash='dot'),
                    showlegend=False, hoverinfo='skip'
                ), row=row, col=1)
                fig.add_trace(go.Scatter(
                    x=pred_et['ds'].tolist(), y=pred_et['yhat_lower'].tolist(),
                    mode='lines', line=dict(width=0, color='rgba(0,0,0,0)'),
                    showlegend=False, hoverinfo='skip'
                ), row=row, col=1)
                fig.add_trace(go.Scatter(
                    x=pred_et['ds'].tolist(), y=pred_et['yhat_upper'].tolist(),
                    mode='lines', line=dict(width=0, color='rgba(0,0,0,0)'),
                    fill='tonexty', fillcolor='rgba(67,160,71,0.20)',
                    showlegend=False, hoverinfo='skip'
                ), row=row, col=1)
                fig.add_trace(go.Scatter(
                    x=pred_et['ds'].tolist(), y=pred_et['yhat'].tolist(),
                    mode='lines', line=dict(color='#43A047', width=1.8, dash='dot'),
                    showlegend=(i == 0), name='ET predicha',
                    hovertemplate='%{x|%d/%b/%Y}<br>ET pred: %{y:.2f} mm<extra>Prophet</extra>'
                ), row=row, col=1)
                fig.add_annotation(
                    x=pred_et['ds'].iloc[-1], y=float(pred_et['yhat'].iloc[-1]),
                    text=f"<b>{float(pred_et['yhat'].iloc[-1]):.2f}</b>",
                    showarrow=True, arrowhead=2, arrowcolor='#43A047', arrowwidth=1.0,
                    ax=-28, ay=-18, font=dict(size=9, color='#43A047'), row=row, col=1
                )

        fig.layout.annotations[i].update(
            text=f"<b>💧 {fundo}</b>",
            font=dict(size=13, color='#00838F'),
            bgcolor='rgba(0,0,0,0)', borderpad=4,
            bordercolor='rgba(0,0,0,0)', x=0.0, xanchor='left'
        )
        xk = 'xaxis' if row == 1 else f'xaxis{row}'
        fig.layout[xk].update(
        range=[et_real['Fecha'].min() if not et_real.empty else fecha_ini_anio,
            ultimo_dia_mes_actual + pd.Timedelta(days=2)],
            tickformat='%d/%b', gridcolor=GRID_COLOR, gridwidth=0.5,
            showgrid=True, zeroline=False,
        )
        yk = 'yaxis' if row == 1 else f'yaxis{row}'
        fig.layout[yk].update(title='mm/día', gridcolor=GRID_COLOR, gridwidth=0.5,
                               showgrid=True, zeroline=True, rangemode='tozero')

    fig.update_layout(
        height=300 * n, paper_bgcolor='#FFFFFF', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Arial', size=9, color='#333333'),
        title=dict(text="<b>💧 Evapotranspiración diaria — FUNDOS AQUANQA</b>",
                   font=dict(size=14, color='#1A1A1A'), x=0.0, xanchor='left', pad=dict(t=10, l=10)),
        margin=dict(l=60, r=40, t=60, b=40),
        legend=dict(orientation='h', yanchor='top', y=-0.05, xanchor='left', x=0,
                    font=dict(size=8), bgcolor='rgba(0,0,0,0)', borderwidth=0),
        hovermode='x unified'
    )

    with col_graf_et:
        st.plotly_chart(fig, use_container_width=True, key="fig_et_main")

    with col_mapa_et:
        et_por_fundo = calcular_et_suma_semanal_fundos(dia)
        colormap_et, leyenda_et_data = crear_colormap_y_leyenda_et(et_por_fundo)
        mapa_et = generar_mapa_et(modulos_kmz, et_por_fundo, colormap_et, df_et_mensual, riesgos_et_pre)
        if mapa_et is not None:
            from streamlit_folium import st_folium
            st_folium(mapa_et, width=None, height=altura_et,
                      returned_objects=[], key='mapa_et')

    # ── KPIs resumen ───────────────────────────────────────────
    st.markdown("#### 📊 Resumen ET por fundo")
    cols = st.columns(len(fundos))
    for i, fundo in enumerate(fundos):
        sub_et = dia[(dia['Fundo'] == fundo) & (dia['Fecha'] >= fecha_ini_anio)]['ET'].dropna()
        if not sub_et.empty:
            cols[i].metric(fundo, f"{sub_et.sum():.1f} mm", f"Prom {sub_et.mean():.2f} mm/día")
def generar_tab_rad(dia: pd.DataFrame, forecasts_cache: dict, dias_pred_ui: int,
                    modulos_kmz=None) -> None:
    """Renderiza el tab completo de Radiación Solar."""
    from services.map_service import calcular_rad_promedio_fundos, generar_mapa_rad

    fundos         = dia['Fundo'].unique().tolist()
    anio_actual    = pd.Timestamp.today().year
    fecha_ini_anio = pd.Timestamp(year=anio_actual - 1, month=1, day=1)
    n              = len(fundos)
    altura_rad     = 300 * n

    col_graf_rad, col_mapa_rad = st.columns([3, 1])

    fig = make_subplots(rows=n, cols=1, shared_xaxes=False,
                        vertical_spacing=0.08, subplot_titles=fundos)

    for i, fundo in enumerate(fundos):
        row = i + 1
        sub = dia[
            (dia['Fundo'] == fundo) & (dia['Fecha'] >= fecha_ini_anio)
        ].copy().reset_index(drop=True)

        if sub.empty or 'RadSolarAlta' not in sub.columns:
            continue

        rad_df = sub[['Fecha', 'RadSolarAlta']].dropna(subset=['RadSolarAlta']).copy()
        if rad_df.empty:
            continue

        fecha_corte_real      = pd.to_datetime(rad_df['Fecha'].max()).normalize()
        primer_dia_mes_actual, ultimo_dia_mes_actual = _fin_mes_prediccion(fecha_corte_real)

        rad_real = rad_df[rad_df['Fecha'] <= fecha_corte_real].copy().reset_index(drop=True)
        rad_real['RadSolarAlta_smooth'] = (
            rad_real['RadSolarAlta'].rolling(ROLLING_DIAS, center=True, min_periods=1).mean().round(1)
        )

        # ── Línea RadSolarAlta ────────────────────────────────
        fig.add_trace(go.Scatter(
            x=rad_real['Fecha'], y=rad_real['RadSolarAlta_smooth'],
            mode='lines+markers',
            line=dict(color='#E65100', width=2.0),
            marker=dict(size=4, color='white', line=dict(color='#E65100', width=1.0)),
            showlegend=(i == 0), name='RadSolar Máx (suav)',
            hovertemplate='%{x|%d/%b/%Y}<br>Rad Máx: %{y:.0f} W/m²<extra>' + fundo + '</extra>',
        ), row=row, col=1)

        # ── Anotación último valor ────────────────────────────
        if not rad_real.empty:
            ult_val = float(rad_real['RadSolarAlta'].iloc[-1])
            fig.add_annotation(
                x=rad_real['Fecha'].iloc[-1], y=ult_val,
                text=f"<b>{ult_val:.0f}</b>",
                showarrow=True, arrowhead=2, arrowcolor='#E65100', arrowwidth=1.0,
                ax=14, ay=-18, font=dict(size=9, color='#E65100'),
                row=row, col=1
            )
            fig.add_vline(x=fecha_corte_real, line=dict(color='#546E7A', width=1.2, dash='dot'),
                          row=row, col=1)

        # ── Predicción Prophet RadSolarAlta ───────────────────
        cache_entry  = forecasts_cache.get((fundo, 'RadSolarAlta'), {})
        forecast_rad = cache_entry.get('forecast', pd.DataFrame())

        if not forecast_rad.empty and not rad_real.empty:
            forecast_rad = forecast_rad.copy()
            forecast_rad['ds'] = pd.to_datetime(forecast_rad['ds']).dt.tz_localize(None).dt.normalize()
            primer_dia_pred = fecha_corte_real + pd.Timedelta(days=1)
            pred_rad = forecast_rad[
                (forecast_rad['ds'] >= primer_dia_pred) &
                (forecast_rad['ds'] <= ultimo_dia_mes_actual)
            ].reset_index(drop=True)

            if not pred_rad.empty:
                # ── Banda IC predicción ────────────────────────
                # ── Línea conector ─────────────────────────────
                fig.add_trace(go.Scatter(
                    x=[rad_real['Fecha'].iloc[-1], pred_rad['ds'].iloc[0]],
                    y=[float(rad_real['RadSolarAlta_smooth'].iloc[-1]), float(pred_rad['yhat'].iloc[0])],
                    mode='lines', line=dict(color='rgba(230,81,0,0.70)', width=1.5, dash='dot'),
                    showlegend=False, hoverinfo='skip',
                ), row=row, col=1)

                # ── Banda IC ───────────────────────────────────
                fig.add_trace(go.Scatter(
                    x=pred_rad['ds'].tolist(), y=pred_rad['yhat_lower'].tolist(),
                    mode='lines', line=dict(width=0, color='rgba(0,0,0,0)'),
                    showlegend=False, hoverinfo='skip',
                ), row=row, col=1)
                fig.add_trace(go.Scatter(
                    x=pred_rad['ds'].tolist(), y=pred_rad['yhat_upper'].tolist(),
                    mode='lines', line=dict(width=0, color='rgba(0,0,0,0)'),
                    fill='tonexty', fillcolor='rgba(230,81,0,0.20)',
                    showlegend=False, hoverinfo='skip',
                ), row=row, col=1)

                # ── Línea predicción ───────────────────────────
                fig.add_trace(go.Scatter(
                    x=pred_rad['ds'].tolist(), y=pred_rad['yhat'].tolist(),
                    mode='lines', line=dict(color='rgba(230,81,0,0.70)', width=1.8, dash='dot'),
                    showlegend=(i == 0), name='RadSolarAlta predicha',
                    hovertemplate='%{x|%d/%b/%Y}<br>Rad Máx pred: %{y:.0f} W/m²<extra>Prophet</extra>',
                ), row=row, col=1)

                fig.add_annotation(
                    x=pred_rad['ds'].iloc[-1], y=float(pred_rad['yhat'].iloc[-1]),
                    text=f"<b>{float(pred_rad['yhat'].iloc[-1]):.0f}</b>",
                    showarrow=True, arrowhead=2, arrowcolor='#E65100', arrowwidth=1.0,
                    ax=-28, ay=-18, font=dict(size=9, color='#E65100'),
                    row=row, col=1
                )

        fig.layout.annotations[i].update(
            text=f"<b>☀️ {fundo}</b>",
            font=dict(size=13, color='#E65100'),
            bgcolor='rgba(0,0,0,0)', borderpad=4,
            bordercolor='rgba(0,0,0,0)', x=0.0, xanchor='left'
        )
        xk = 'xaxis' if row == 1 else f'xaxis{row}'
        fig.layout[xk].update(
            range=[rad_real['Fecha'].min() if not rad_real.empty else fecha_ini_anio,
                   ultimo_dia_mes_actual + pd.Timedelta(days=2)],
            tickformat='%d/%b', gridcolor=GRID_COLOR, gridwidth=0.5,
            showgrid=True, zeroline=False,
        )
        yk = 'yaxis' if row == 1 else f'yaxis{row}'
        fig.layout[yk].update(
            title='W/m²', gridcolor=GRID_COLOR, gridwidth=0.5,
            showgrid=True, zeroline=True, rangemode='tozero'
        )

    fig.update_layout(
        height=300 * n, paper_bgcolor='#FFFDE7', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Arial', size=9, color='#333333'),
        title=dict(text="<b>☀️ Radiación Solar Máxima diaria — FUNDOS AQUANQA</b>",
                   font=dict(size=14, color='#E65100'), x=0.0, xanchor='left',
                   pad=dict(t=10, l=10)),
        margin=dict(l=60, r=40, t=60, b=40),
        legend=dict(orientation='h', yanchor='top', y=-0.05, xanchor='left', x=0,
                    font=dict(size=8), bgcolor='rgba(0,0,0,0)', borderwidth=0),
        hovermode='x unified'
    )

    with col_graf_rad:
        st.plotly_chart(fig, use_container_width=True, key="fig_rad_main")

    with col_mapa_rad:
        rad_por_fundo = calcular_rad_promedio_fundos(dia)
        mapa_rad      = generar_mapa_rad(modulos_kmz, rad_por_fundo)
        if mapa_rad is not None:
            from streamlit_folium import st_folium
            st_folium(mapa_rad, width=None, height=altura_rad,
                      returned_objects=[], key='mapa_rad')

    # ── KPIs resumen ───────────────────────────────────────────
    st.markdown("#### 📊 Resumen Radiación Solar Máxima por fundo")
    cols = st.columns(len(fundos))
    for i, fundo in enumerate(fundos):
        sub_rad = dia[
            (dia['Fundo'] == fundo) & (dia['Fecha'] >= fecha_ini_anio)
        ]['RadSolarAlta'].dropna()
        if not sub_rad.empty:
            cols[i].metric(
                fundo,
                f"{sub_rad.mean():.0f} W/m²",
                f"Máx {sub_rad.max():.0f} W/m²"
            )