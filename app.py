import sys, os
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
sys.path.insert(0, BASE_DIR)
sys.path.insert(0, os.path.abspath(os.path.join(BASE_DIR, '..')))
sys.path.insert(0, os.path.abspath(os.path.join(BASE_DIR, '..', '..')))

import pandas as pd
import streamlit as st
from services.visualization_service import generar_figura, generar_tab_et, generar_tab_rad

# ── Page config (must be first Streamlit call) ────────────────
st.set_page_config(
    page_title="Temperatura Fundos — Aquanqa",
    page_icon="🌡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

from config.config import (
    NORMALES_PATH, TMAX_MIN, TMAX_MAX, TMIN_MIN, TMIN_MAX,
    MIN_REGISTROS, DISTRITO_BUSCAR, DISTRITO_NINGUNA, MESES_RAW,
)
from styles.styles import MAIN_CSS
from services.enfen_service import (
    chequear_comunicado_enfen, _leer_ultimo_visto_enfen, _guardar_ultimo_visto_enfen,
)
from services.data_service import (
    _leer_normales_desde_disco, cargar_catalogo_normales, cargar_normales_dinamico,
    obtener_distritos_senamhi, cargar_temperaturas_distritos,
    cargar_geojson_peru, get_conn_fabric, cargar_et_mensual_promedio, cargar_datos_clima,
)
from services.map_service import (
    load_kmz_bytes, download_kmz_from_github, disolver_modulos,
    crear_colormap_temperatura, generar_mapa_distritos, calcular_riesgo_et_fundo,
)
from services.prophet_service import (
    entrenar_todos_optimizado, guardar_cache_prophet,
    cargar_cache_prophet, calcular_hash_meteo,
)
from services.visualization_service import generar_figura, generar_tab_et
from services.validation_service import generar_seccion_validacion, generar_export_prediccion_historica
from services.export_service import exportar_excel_tabla
from services.climatologia_service import generar_climatologia_diaria_anual

st.markdown(MAIN_CSS, unsafe_allow_html=True)

# ──── SIDEBAR ─────────────────────────────────────────────────
norm_bytes_sidebar = _leer_normales_desde_disco(NORMALES_PATH) if os.path.exists(NORMALES_PATH) else None
# ← DIAGNÓSTICO TEMPORAL
if norm_bytes_sidebar:
    import zipfile, io
    print(f"Bytes leídos: {len(norm_bytes_sidebar)}")
    print(f"Header hex: {norm_bytes_sidebar[:8].hex()}")
    try:
        zipfile.ZipFile(io.BytesIO(norm_bytes_sidebar))
        print("✅ ZIP válido")
    except Exception as e:
        print(f"❌ No es ZIP válido: {e}")
else:
    print("❌ norm_bytes_sidebar es None")
with st.sidebar:
    st.markdown("### 🌊 Monitor ENFEN")
    comunicado_actual = chequear_comunicado_enfen()

    if comunicado_actual is None:
        st.caption("⚠️ No se pudo verificar ENFEN (sin conexión o cambió el formato de la página).")
    else:
        ultimo_visto = _leer_ultimo_visto_enfen()
        try:
            from config.config import ENFEN_URL
        except Exception:
            ENFEN_URL = None
        if comunicado_actual['id'] != ultimo_visto.get('id'):
            st.warning(
                f"📢 **Nuevo comunicado ENFEN**\n\n"
                f"N°{comunicado_actual['numero']}-{comunicado_actual['anio']} "
                f"({comunicado_actual['fecha']})\n\n"
                f"Revisa si `AJUSTE_ENFEN` necesita actualizarse."
            )
            st.markdown(f"[📄 Ver comunicados ENFEN]({ENFEN_URL})")
            if st.button("✅ Marcar como revisado"):
                _guardar_ultimo_visto_enfen(comunicado_actual)
                st.rerun()
        else:
            st.caption(
                f"✅ Al día — último: N°{comunicado_actual['numero']}-{comunicado_actual['anio']} "
                f"({comunicado_actual['fecha']})"
            )
            st.markdown(f"[📄 Ver comunicados ENFEN]({ENFEN_URL or '#'})")

with st.sidebar:
    st.markdown("## ⚙️ CONFIGURACIÓN")

    with st.expander("📂 **Archivo meteorológico**", expanded=True):
        st.markdown("### 📊 **Archivo meteorológico**")
        st.info("✅ Leyendo: `assets/Metereologia_Prize.xlsx`")

    st.divider()

    st.markdown("### 🌿 **Fundos**")
    fundos_disponibles_placeholder = st.empty()

    st.divider()

    bias_dict = st.session_state.get('bias_correccion', {})
    if bias_dict:
        st.markdown("### 🎯 Corrección BIAS")
        for (f, v), b in bias_dict.items():
            icono = '🔺' if b > 0 else '🔻'
            st.caption(f"{icono} {f} {v}: {b:+.2f}°C")
    else:
        st.caption("⚠️ Abre validación para activar BIAS")

    st.markdown("### 🌦️ **Climatología SENAMHI**")
    st.info(f"📌 Referencia fija: **{DISTRITO_BUSCAR}**")
    st.markdown("**Comparar con otra estación:**")

    distrito_sel2 = DISTRITO_NINGUNA
    if norm_bytes_sidebar is not None:
        catalogo = cargar_catalogo_normales(norm_bytes_sidebar, 'TMAX')
        if catalogo:
            sectores_lista = sorted(catalogo.keys())
            sector_sel = st.selectbox("Sector", options=sectores_lista, index=0, key='sector_sel')
            deptos_lista = sorted(catalogo.get(sector_sel, {}).keys())
            depto_sel = st.selectbox("Departamento", options=deptos_lista, index=0, key='depto_sel')
            distritos_lista = [DISTRITO_NINGUNA] + catalogo.get(sector_sel, {}).get(depto_sel, [])
            distrito_sel2 = st.selectbox(
                "Estación / Distrito", options=distritos_lista, index=0, key='distrito_sel2'
            )
            if distrito_sel2 != DISTRITO_NINGUNA:
                st.success(f"✅ Comparando: **{distrito_sel2}**")
            else:
                st.caption("Sin comparación activa.")
        else:
            st.error("❌ No se pudo leer el catálogo.")
    else:
        st.error("❌ Normales SENAMHI no encontradas")

# ──── DATOS AUXILIARES ────────────────────────────────────────
geojson_peru      = cargar_geojson_peru()
distritos_senamhi = obtener_distritos_senamhi(norm_bytes_sidebar, 'TMAX') if norm_bytes_sidebar is not None else set()

_kmz_bytes_modulos = download_kmz_from_github()
kmz_polygons       = load_kmz_bytes(_kmz_bytes_modulos) if _kmz_bytes_modulos else []
modulos_kmz        = disolver_modulos(kmz_polygons) if kmz_polygons else []

# ──── HEADER ─────────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
    <h1>🌡️ Temperatura Fundos — Aquanqa</h1>
    <p>Climatología SENAMHI · Predicción Prophet · OPTIMIZADO</p>
</div>
""", unsafe_allow_html=True)

# ──── VALORES POR DEFECTO ─────────────────────────────────────
dias_vista_num = 30
variable_sel   = 'Ambas'
min_reg_ui     = MIN_REGISTROS
dias_pred_ui   = 30

# ──── NORMALES GUADALUPE (fijo) ───────────────────────────────
normales_guadalupe = cargar_normales_dinamico(norm_bytes_sidebar, DISTRITO_BUSCAR)
if normales_guadalupe is None:
    st.stop()

MEDIA_TMAX, Q1_TMAX, Q3_TMAX, MEDIA_TMIN, Q1_TMIN, Q3_TMIN, estaciones_tmax = normales_guadalupe

# ──── NORMALES DINÁMICAS ──────────────────────────────────────
normales_din = None
if distrito_sel2 != DISTRITO_NINGUNA:
    normales_din = cargar_normales_dinamico(norm_bytes_sidebar, distrito_sel2)

if normales_din is not None:
    MEDIA_TMAX2, Q1_TMAX2, Q3_TMAX2, MEDIA_TMIN2, Q1_TMIN2, Q3_TMIN2, _ = normales_din
else:
    MEDIA_TMAX2 = Q1_TMAX2 = Q3_TMAX2 = None
    MEDIA_TMIN2 = Q1_TMIN2 = Q3_TMIN2 = None

# ──── CONEXIÓN FABRIC ─────────────────────────────────────────
if 'conn_fabric' not in st.session_state or st.session_state['conn_fabric'] is None:
    with st.spinner("Conectando a Fabric..."):
        st.session_state['conn_fabric'] = get_conn_fabric()

conn_fabric = st.session_state['conn_fabric']

if conn_fabric is None:
    st.error("❌ No se pudo conectar a Fabric")
    st.stop()

# ← DEBE ESTAR AQUÍ, justo después de conn_fabric
df_et_mensual = cargar_et_mensual_promedio(conn_fabric)

# ──── CARGAR DATOS DESDE FABRIC ───────────────────────────────
with st.spinner("Cargando..."):
    try:
        df_fabric = cargar_datos_clima(conn_fabric)
        if df_fabric.empty:
            st.error("❌ La vista vw_Clima no tiene datos")
            st.stop()

        fecha_min = pd.to_datetime(df_fabric['Fecha-Hora'], format='mixed', dayfirst=True).min()
        fecha_max = pd.to_datetime(df_fabric['Fecha-Hora'], format='mixed', dayfirst=True).max()
        st.info(
            f"✅ {len(df_fabric):,} registros desde Fabric\n"
            f"📅 {fecha_min.strftime('%d/%m/%Y')} → {fecha_max.strftime('%d/%m/%Y')}"
        )
    except Exception as e:
        st.error(f"Error al leer Fabric: {e}")
        st.stop()

# ──── FUNDOS MULTISELECT ──────────────────────────────────────
fundos_disponibles = sorted(df_fabric['Fundo'].dropna().unique().tolist())

with fundos_disponibles_placeholder:
    fundos_sel = st.multiselect(
        "Seleccionar fundos",
        options=fundos_disponibles,
        default=fundos_disponibles,
    )

fundos_activos = fundos_sel if fundos_sel else fundos_disponibles

# ──── PROCESAR DATOS ─────────────────────────────────────────
with st.spinner("Procesando datos desde Fabric..."):
    try:
        df = df_fabric.copy()

        df['Fecha-Hora'] = pd.to_datetime(df['Fecha-Hora'], format='mixed', dayfirst=True, errors='coerce')
        df = df.dropna(subset=['Fecha-Hora'])

        for col in ['TempAlta-C', 'TempBaja-C']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df[
            (df['TempAlta-C'] >= TMAX_MIN) & (df['TempAlta-C'] <= TMAX_MAX) &
            (df['TempBaja-C'] >= TMIN_MIN) & (df['TempBaja-C'] <= TMIN_MAX)
        ].copy()

        if 'ET-mm' not in df.columns:
            df['ET-mm'] = 0.0
        df['ET-mm'] = pd.to_numeric(df['ET-mm'], errors='coerce').fillna(0.0)

        df['Fecha'] = df['Fecha-Hora'].dt.normalize()

        reg = df.groupby(['Fundo', 'Fecha'], as_index=False).size()
        reg.columns = ['Fundo', 'Fecha', 'N_registros']
        df = df.merge(reg, on=['Fundo', 'Fecha'], how='left')

        fecha_max_por_fundo = (
            df.groupby('Fundo')['Fecha']
              .max()
              .reset_index()
              .rename(columns={'Fecha': 'Fecha_max_fundo'})
        )
        df = df.merge(fecha_max_por_fundo, on='Fundo', how='left')

        df = df[
            (df['N_registros'] >= min_reg_ui) |
            (df['Fecha'] == df['Fecha_max_fundo'])
        ].copy()
        df = df.drop(columns=['Fecha_max_fundo'])
        df = df[df['Fundo'].isin(fundos_activos)]

        dia_full = (
            df.groupby(['Empresa', 'Fundo', 'Fecha'], as_index=False)
              .agg(Tmax=('TempAlta-C', 'max'), Tmin=('TempBaja-C', 'min'), ET=('ET-mm', 'sum'), RadSolar=('RadSolar-W/m2', 'mean'), RadSolarAlta=('RadSolarAlta-W/m2', 'max'))
        )
        dia_full[['Tmax', 'Tmin', 'ET']] = dia_full[['Tmax', 'Tmin', 'ET']].round(2)
        dia_full = dia_full.sort_values(['Fundo', 'Fecha']).reset_index(drop=True)

        dia_full['Tmax_smooth'] = (
            dia_full.groupby('Fundo')['Tmax']
                    .transform(lambda x: x.rolling(3, center=True, min_periods=1).mean())
                    .round(2)
        )
        dia_full['Tmin_smooth'] = (
            dia_full.groupby('Fundo')['Tmin']
                    .transform(lambda x: x.rolling(3, center=True, min_periods=1).mean())
                    .round(2)
        )

        dia = dia_full.copy()

    except Exception as e:
        st.error(f"Error procesando datos: {e}")
        st.stop()

# ──── PROPHET CACHE ───────────────────────────────────────────
_data_hash      = calcular_hash_meteo(dia_full)
forecasts_cache = cargar_cache_prophet(_data_hash)

if forecasts_cache is None:
    with st.spinner("Entrenando modelos Prophet..."):
        forecasts_cache = entrenar_todos_optimizado(dia_full, dias_pred_ui)
    guardar_cache_prophet(forecasts_cache, _data_hash)
    st.success("Modelos entrenados y guardados en caché.")
else:
    st.info("Modelos cargados desde caché.")

# ──── MÉTRICAS ────────────────────────────────────────────────
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

# ──── TABS ────────────────────────────────────────────────────
mostrar_tmax = variable_sel in ['Ambas', 'Solo Tmax']
mostrar_tmin = variable_sel in ['Ambas', 'Solo Tmin']

dias_pred_mitad = max(7, dias_pred_ui // 2)

if mostrar_tmax and mostrar_tmin:
    tab_tmax, tab_tmin, tab_et, tab_rad, tab_datos = st.tabs(
        ["🔴 Temperatura Máxima", "🔵 Temperatura Mínima", "💧 ET", "☀️ Radiación Solar", "📋 Datos"],
        key="main_tabs"
    )
elif mostrar_tmax:
    tab_tmax, tab_tmin, tab_et, tab_rad, tab_datos = st.tabs(
        ["🔴 Temperatura Máxima", "🔵 Temperatura Mínima", "💧 ET", "☀️ Radiación Solar", "📋 Datos"],
        key="main_tabs"
    )
    tab_tmin = None
else:
    tab_tmax, tab_tmin, tab_et, tab_rad, tab_datos = st.tabs(
        ["🔴 Temperatura Máxima", "🔵 Temperatura Mínima", "💧 ET", "☀️ Radiación Solar", "📋 Datos"],
        key="main_tabs"
    )
    tab_tmax = None

n_fundos         = len(dia['Fundo'].unique())
altura_graficos  = 320 * n_fundos
ALTURA_SELECTBOX = 78

# Pre-calcular riesgos ET
fundos_activos_list = list(fundos_activos) if fundos_activos else []
riesgos_et_cache    = {f: calcular_riesgo_et_fundo(df_et_mensual, f) for f in fundos_activos_list}

# ── TAB TMAX ──────────────────────────────────────────────────
if mostrar_tmax and tab_tmax is not None:
    with tab_tmax:
        col1, col2 = st.columns([2, 2])
        with col1:
            st.markdown("#### 📊 Predicción a mostrar")
            dias_pred_mostrar = st.radio(
                "Mostrar predicción de:",
                options=[f"{dias_pred_mitad}d", f"{dias_pred_ui}d"],
                index=1, horizontal=True,
                key='dias_pred_mostrar_tmax'
            )
            dias_pred_mostrar_num = int(dias_pred_mostrar.replace('d', ''))

        with st.spinner("Generando Tmax..."):
            fig_tmax, df_tmax_hist, df_tmax_pred = generar_figura(
                'Tmax', dia, MEDIA_TMAX, Q1_TMAX, Q3_TMAX,
                dias_pred_ui, forecasts_cache, dias_vista_num,
                dias_pred_mostrar_num, 'linea',
                media_mensual2=MEDIA_TMAX2, q1_mensual2=Q1_TMAX2, q3_mensual2=Q3_TMAX2,
                label_clim2=distrito_sel2 if distrito_sel2 != DISTRITO_NINGUNA else None
            )

        col_graf_tmax, col_mapa_tmax = st.columns([3, 1])
        with col_graf_tmax:
            st.plotly_chart(fig_tmax, use_container_width=True)
        with col_mapa_tmax:
            from streamlit_folium import st_folium

            mes_actual    = pd.Timestamp.today().month
            mes_sel_tmax  = st.selectbox(
                "Mes", options=list(range(1, 13)),
                format_func=lambda mm: MESES_RAW[mm - 1],
                index=mes_actual - 1, key='mes_sel_anomalia_tmax',
                label_visibility='collapsed'
            )
            temps_distritos_tmax = (
                cargar_temperaturas_distritos(norm_bytes_sidebar, 'TMAX', mes_sel_tmax)
                if norm_bytes_sidebar is not None else {}
            )
            temp_colormap_tmax = crear_colormap_temperatura(temps_distritos_tmax, 'Tmax', mes_sel_tmax)
            mapa_tmax = generar_mapa_distritos(
                geojson_peru,
                distrito_fijo=DISTRITO_BUSCAR,
                distrito_din=distrito_sel2 if distrito_sel2 != DISTRITO_NINGUNA else None,
                variable='Tmax', dia=dia, mes_sel=mes_sel_tmax,
                media_clim_fijo=MEDIA_TMAX, media_clim_din=MEDIA_TMAX2,
                label_clim2=distrito_sel2 if distrito_sel2 != DISTRITO_NINGUNA else None,
                modulos_kmz=modulos_kmz, distritos_senamhi=distritos_senamhi,
                temps_distritos=temps_distritos_tmax, temp_colormap=temp_colormap_tmax,
            )
            if mapa_tmax:
                st_folium(mapa_tmax, width=None,
                          height=max(300, altura_graficos - ALTURA_SELECTBOX),
                          returned_objects=[], key='mapa_tmax')

        generar_seccion_validacion('Tmax', dia, forecasts_cache, dias_pred_ui)
else:
    df_tmax_hist = pd.DataFrame()
    df_tmax_pred = pd.DataFrame()

# ── TAB TMIN ──────────────────────────────────────────────────
if mostrar_tmin and tab_tmin is not None:
    with tab_tmin:
        col1, col2 = st.columns([2, 2])
        with col1:
            st.markdown("#### 📊 Predicción a mostrar")
            dias_pred_mostrar = st.radio(
                "Mostrar predicción de:",
                options=[f"{dias_pred_mitad}d", f"{dias_pred_ui}d"],
                index=1, horizontal=True,
                key='dias_pred_mostrar_tmin'
            )
            dias_pred_mostrar_num = int(dias_pred_mostrar.replace('d', ''))

        with st.spinner("Generando Tmin..."):
            fig_tmin, df_tmin_hist, df_tmin_pred = generar_figura(
                'Tmin', dia, MEDIA_TMIN, Q1_TMIN, Q3_TMIN,
                dias_pred_ui, forecasts_cache, dias_vista_num,
                dias_pred_mostrar_num, 'linea',
                media_mensual2=MEDIA_TMIN2, q1_mensual2=Q1_TMIN2, q3_mensual2=Q3_TMIN2,
                label_clim2=distrito_sel2 if distrito_sel2 != DISTRITO_NINGUNA else None
            )

        col_graf_tmin, col_mapa_tmin = st.columns([3, 1])
        with col_graf_tmin:
            st.plotly_chart(fig_tmin, use_container_width=True)
        with col_mapa_tmin:
            from streamlit_folium import st_folium

            mes_actual    = pd.Timestamp.today().month
            mes_sel_tmin  = st.selectbox(
                "Mes", options=list(range(1, 13)),
                format_func=lambda mm: MESES_RAW[mm - 1],
                index=mes_actual - 1, key='mes_sel_anomalia_tmin',
                label_visibility='collapsed'
            )
            temps_distritos_tmin = (
                cargar_temperaturas_distritos(norm_bytes_sidebar, 'TMIN', mes_sel_tmin)
                if norm_bytes_sidebar is not None else {}
            )
            temp_colormap_tmin = crear_colormap_temperatura(temps_distritos_tmin, 'Tmin', mes_sel_tmin)
            mapa_tmin = generar_mapa_distritos(
                geojson_peru,
                distrito_fijo=DISTRITO_BUSCAR,
                distrito_din=distrito_sel2 if distrito_sel2 != DISTRITO_NINGUNA else None,
                variable='Tmin', dia=dia, mes_sel=mes_sel_tmin,
                media_clim_fijo=MEDIA_TMIN, media_clim_din=MEDIA_TMIN2,
                label_clim2=distrito_sel2 if distrito_sel2 != DISTRITO_NINGUNA else None,
                modulos_kmz=modulos_kmz, distritos_senamhi=distritos_senamhi,
                temps_distritos=temps_distritos_tmin, temp_colormap=temp_colormap_tmin,
            )
            if mapa_tmin:
                st_folium(mapa_tmin, width=None,
                          height=max(300, altura_graficos - ALTURA_SELECTBOX),
                          returned_objects=[], key='mapa_tmin')

        generar_seccion_validacion('Tmin', dia, forecasts_cache, dias_pred_ui)
else:
    df_tmin_hist = pd.DataFrame()
    df_tmin_pred = pd.DataFrame()

# ── TAB ET ────────────────────────────────────────────────────
if tab_et is not None:
    with tab_et:
        generar_tab_et(
            dia, forecasts_cache, dias_pred_ui,
            modulos_kmz, riesgos_et_cache,
            df_et_mensual=df_et_mensual
        )
# ── TAB RAD ───────────────────────────────────────────────────
if tab_rad is not None:
    with tab_rad:
        generar_tab_rad(
            dia, forecasts_cache, dias_pred_ui,
            modulos_kmz,
        )
# ── TAB DATOS ─────────────────────────────────────────────────
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

    df_show         = dia[dia['Fundo'].isin(fundos_tabla)].copy()
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
    st.caption(
        "La predicción exportable cubre los últimos 3 meses cerrados más el mes en curso, "
        "los 4 reconstruidos con walk-forward puro: cada mes se entrena solo con datos "
        "reales hasta el día antes de que empiece, sin ver ni un día de ese mes (Tipo="
        "'Predicho' en todas las filas, incluso donde ya existe dato real)."
    )

    if st.button("🔄 Generar predicción histórica para exportar"):
        with st.spinner("Entrenando walk-forward (4 meses) por fundo/variable..."):
            st.session_state['df_pred_export'] = generar_export_prediccion_historica(dia)

    df_pred_export = st.session_state.get('df_pred_export', pd.DataFrame())
    df_hist_all    = pd.concat([df_tmax_hist, df_tmin_hist], ignore_index=True)
    if not df_hist_all.empty:
        df_hist_all['Fecha'] = pd.to_datetime(df_hist_all['Fecha']).dt.strftime('%Y/%m/%d')

    anio_clima  = pd.Timestamp.today().year
    df_clima_anual = generar_climatologia_diaria_anual(
        anio_clima, MEDIA_TMAX, Q1_TMAX, Q3_TMAX, MEDIA_TMIN, Q1_TMIN, Q3_TMIN
    )

    col_d1, col_d2, col_d3, col_d4 = st.columns(4)

    with col_d1:
        if not df_pred_export.empty:
            st.download_button(
                label="📥 Excel predicción (walk-forward)",
                data=exportar_excel_tabla(df_pred_export, 'Prediccion'),
                file_name="Prediccion_walkforward_Aquanqa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.caption("Presiona 'Generar predicción histórica' para habilitar la descarga.")
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
        if not df_pred_export.empty:
            st.download_button(
                label="📄 CSV predicciones",
                data=df_pred_export.to_csv(index=False).encode('utf-8'),
                file_name="predicciones_walkforward.csv",
                mime="text/csv",
                use_container_width=True
            )
    with col_d4:
        st.download_button(
            label=f"📥 Excel climatología diaria {anio_clima}",
            data=exportar_excel_tabla(df_clima_anual, 'Climatologia'),
            file_name=f"Climatologia_SENAMHI_diaria_{anio_clima}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

st.markdown("""
<div class="footer-note">
    Elaborado por la Gerencia de Planificación, Control Operacional y Gestión
</div>
""", unsafe_allow_html=True)
