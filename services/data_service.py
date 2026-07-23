import io
import json
import os

import pandas as pd
import streamlit as st
from sqlalchemy import create_engine

from config.config import (
    SHEET_NAME, MIN_REGISTROS, ROLLING_DIAS,
    TMAX_MIN, TMAX_MAX, TMIN_MIN, TMIN_MAX,
    DELTA_Q, MESES_RAW, DISTRITO_BUSCAR,
)
from services.utils import _normalizar


# ── Lectura de archivos ────────────────────────────────────────

@st.cache_data(show_spinner=False)
def leer_meteo_bytes_optimizado(file_bytes: bytes, filename: str, sheet: str = SHEET_NAME) -> pd.DataFrame:
    ext = filename.lower().rsplit('.', 1)[-1]

    if ext == 'csv':
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8-sig',
                             sep=None, engine='python', on_bad_lines='skip')
        except Exception:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding='latin-1',
                             sep=None, engine='python', on_bad_lines='skip')
        return df

    elif ext == 'xlsx':
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, parse_dates=False)

    raise ValueError(f"Extensión no soportada: .{ext}")


@st.cache_data(show_spinner=False)
def cargar_meteoro_optimizado(_file_bytes: bytes, filename: str, sheet: str,
                              fundos_sel: tuple, min_reg: int) -> pd.DataFrame:
    df = leer_meteo_bytes_optimizado(_file_bytes, filename, sheet)

    df['Fecha-Hora'] = pd.to_datetime(df['Fecha-Hora'], dayfirst=False, errors='coerce')
    df = df.dropna(subset=['Fecha-Hora'])

    for col in ['Temp-C', 'TempAlta-C', 'TempBaja-C']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df[
        (df['TempAlta-C'] >= TMAX_MIN) & (df['TempAlta-C'] <= TMAX_MAX) &
        (df['TempBaja-C'] >= TMIN_MIN) & (df['TempBaja-C'] <= TMIN_MAX)
    ].copy()

    if 'ET-mm' not in df.columns:
        df['ET-mm'] = 0.0

    df['Fecha'] = df['Fecha-Hora'].dt.normalize()

    reg = df.groupby(['Fundo', 'Fecha'], as_index=False).size()
    reg.columns = ['Fundo', 'Fecha', 'N_registros']
    df = df.merge(reg, on=['Fundo', 'Fecha'], how='left')

    fecha_max_por_fundo = (
        df.groupby('Fundo')['Fecha'].max().reset_index()
          .rename(columns={'Fecha': 'Fecha_max_fundo'})
    )
    df = df.merge(fecha_max_por_fundo, on='Fundo', how='left')
    df = df[
        (df['N_registros'] >= min_reg) | (df['Fecha'] == df['Fecha_max_fundo'])
    ].copy()
    df = df.drop(columns=['Fecha_max_fundo'])

    if fundos_sel:
        df = df[df['Fundo'].isin(fundos_sel)]

    dia = (
        df.groupby(['Empresa', 'Fundo', 'Fecha'], as_index=False)
          .agg(Tmax=('TempAlta-C', 'max'), Tmin=('TempBaja-C', 'min'), ET=('ET-mm', 'sum'))
    )
    dia[['Tmax', 'Tmin', 'ET']] = dia[['Tmax', 'Tmin', 'ET']].round(2)
    dia = dia.sort_values(['Fundo', 'Fecha']).reset_index(drop=True)

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


# ── Normales SENAMHI ───────────────────────────────────────────

@st.cache_data(show_spinner=False)
def cargar_catalogo_normales(_file_bytes: bytes, hoja: str = 'TMAX') -> dict:
    raw = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=hoja, header=None)
    header_row = None
    for idx, row in raw.iterrows():
        vals = [_normalizar(str(v)) for v in row.values if pd.notna(v)]
        if 'DISTRITO' in vals:
            header_row = idx
            break
    if header_row is None:
        return {}

    df = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=hoja, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    col_sector   = next((c for c in df.columns if _normalizar(c) == 'SECTOR'), None)
    col_depto    = next((c for c in df.columns if _normalizar(c) == 'DEPARTAMENTO'), None)
    col_distrito = next((c for c in df.columns if _normalizar(c) == 'DISTRITO'), None)

    if col_depto is None or col_distrito is None:
        return {}

    df['_sector']   = df[col_sector].astype(str).apply(_normalizar) if col_sector else 'SIN SECTOR'
    df['_depto']    = df[col_depto].astype(str).apply(_normalizar)
    df['_distrito'] = df[col_distrito].astype(str).apply(_normalizar)

    df = df[
        (df['_depto']    != 'NAN') & (df['_depto']    != '') &
        (df['_distrito'] != 'NAN') & (df['_distrito'] != '') &
        (df['_sector']   != 'NAN') & (df['_sector']   != '')
    ].copy()

    df = df[~df['_distrito'].str.contains(DISTRITO_BUSCAR, na=False)].copy()

    catalogo = {}
    for sector, g_sector in df.groupby('_sector'):
        catalogo[sector] = {}
        for depto, g_depto in g_sector.groupby('_depto'):
            catalogo[sector][depto] = sorted(g_depto['_distrito'].unique().tolist())
    return catalogo


@st.cache_data(show_spinner=False)
def cargar_normales(_file_bytes: bytes, hoja: str):
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
        raise ValueError("Columna DISTRITO no encontrada.")

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
    col_prov   = next((c for c in df.columns if _normalizar(c) == 'PROVINCIA'), col_distrito)

    return (
        valores,
        valores - DELTA_Q,
        valores + DELTA_Q,
        df_f[[col_nombre, col_prov, col_distrito]].drop_duplicates()
    )


def cargar_normales_dinamico(_file_bytes: bytes, distrito_sel: str):
    """Carga normales SENAMHI para cualquier distrito seleccionado dinámicamente."""
    try:
        def _cargar_hoja(hoja):
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
                raise ValueError("Columna DISTRITO no encontrada.")

            df['_dist_norm'] = df[col_distrito].astype(str).apply(_normalizar)
            df_f = df[df['_dist_norm'].str.contains(distrito_sel, na=False)].copy()

            if df_f.empty:
                raise ValueError(f"No se encontraron datos para '{distrito_sel}'.")

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
            col_prov   = next((c for c in df.columns if _normalizar(c) == 'PROVINCIA'), col_distrito)

            return (
                valores,
                valores - DELTA_Q,
                valores + DELTA_Q,
                df_f[[col_nombre, col_prov, col_distrito]].drop_duplicates()
            )

        MEDIA_TMAX, Q1_TMAX, Q3_TMAX, estaciones_tmax = _cargar_hoja('TMAX')
        MEDIA_TMIN, Q1_TMIN, Q3_TMIN, _               = _cargar_hoja('TMIN')

        return MEDIA_TMAX, Q1_TMAX, Q3_TMAX, MEDIA_TMIN, Q1_TMIN, Q3_TMIN, estaciones_tmax

    except Exception as e:
        st.error(f"Error al cargar normales para '{distrito_sel}': {e}")
        return None


@st.cache_data(show_spinner=False)
def obtener_distritos_senamhi(_file_bytes: bytes, hoja: str = 'TMAX') -> set:
    raw = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=hoja, header=None)
    header_row = None
    for idx, row in raw.iterrows():
        vals = [_normalizar(str(v)) for v in row.values if pd.notna(v)]
        if 'DISTRITO' in vals:
            header_row = idx
            break
    if header_row is None:
        return set()

    df = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=hoja, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    col_distrito = next((c for c in df.columns if _normalizar(c) == 'DISTRITO'), None)
    if col_distrito is None:
        return set()

    distritos = df[col_distrito].dropna().astype(str).apply(_normalizar)
    return set(distritos[(distritos != 'NAN') & (distritos != '')].unique())


@st.cache_data(show_spinner=False)
def cargar_temperaturas_distritos(_file_bytes: bytes, hoja: str, mes_sel: int) -> dict:
    """Devuelve {distrito_normalizado: temperatura_promedio_mes} para todos los distritos."""
    raw = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=hoja, header=None)
    header_row = None
    for idx, row in raw.iterrows():
        vals = [_normalizar(str(v)) for v in row.values if pd.notna(v)]
        if 'DISTRITO' in vals:
            header_row = idx
            break
    if header_row is None:
        return {}

    df = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=hoja, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    col_distrito = next((c for c in df.columns if _normalizar(c) == 'DISTRITO'), None)
    if col_distrito is None:
        return {}

    mes_nombre = MESES_RAW[mes_sel - 1]
    col_mes = next((c for c in df.columns if _normalizar(c) == _normalizar(mes_nombre)), None)
    if col_mes is None:
        return {}

    df = df.copy()
    df['_dist_norm'] = df[col_distrito].astype(str).apply(_normalizar)
    df['_valor'] = pd.to_numeric(df[col_mes], errors='coerce')
    df = df[(df['_dist_norm'] != 'NAN') & (df['_dist_norm'] != '') & df['_valor'].notna()]

    return df.groupby('_dist_norm')['_valor'].mean().round(1).to_dict()


@st.cache_data(show_spinner=False)
def _leer_normales_desde_disco(path: str) -> bytes:
    with open(path, 'rb') as f:
        return f.read()


@st.cache_data(show_spinner=False)
def cargar_geojson_peru() -> dict | None:
    """Descarga GeoJSON de distritos de Perú desde GitHub."""
    url = (
        "https://raw.githubusercontent.com/juaneladio/peru-geojson/"
        "master/peru_distrital_simple.geojson"
    )
    try:
        import urllib.request
        with urllib.request.urlopen(url) as r:
            return json.loads(r.read().decode('utf-8'))
    except Exception as e:
        st.warning(f"⚠️ No se pudo cargar el mapa de Perú: {e}")
        return None


# ── Conexión PostgreSQL ────────────────────────────────────────

def _crear_engine_pg():
    host = st.secrets["PG_HOST"]
    port = st.secrets["PG_PORT"]
    db   = st.secrets["PG_DB"]
    user = st.secrets["PG_USER"]
    pwd  = st.secrets["PG_PASS"]
    url  = f"postgresql+psycopg2://{user}:{pwd}@{host}:{port}/{db}"
    return create_engine(url, connect_args={"connect_timeout": 15})


@st.cache_resource(ttl=3600, show_spinner=False)
def get_conn_fabric():
    try:
        return _crear_engine_pg()
    except Exception as e:
        st.error(f"❌ No se pudo conectar a PostgreSQL: {e}")
        return None


def ejecutar_query(query: str) -> pd.DataFrame:
    """Ejecuta una query con reconexión automática en caso de fallo."""
    import time
    for intento in range(3):
        try:
            conn = get_conn_fabric()
            if conn is None:
                raise Exception("No se pudo obtener conexión")
            return pd.read_sql(query, conn)
        except Exception as e:
            err = str(e)
            if 'connection' in err.lower() or 'timeout' in err.lower():
                if intento < 2:
                    get_conn_fabric.clear()
                    time.sleep(2)
                    continue
            raise
    raise Exception("Falló tras 3 intentos de reconexión")


@st.cache_data(ttl=3600, show_spinner=False)
def cargar_datos_clima(_conn_fabric) -> pd.DataFrame:
    """Carga la vista vw_clima completa y normaliza nombres de columnas."""
    df = pd.read_sql("""
        SELECT empresa, fundo, fecha_hora, temp_alta_c, temp_baja_c,
               et_mm, rad_solar_alta_wm2, rad_solar_wm2
        FROM public.rpt_climatologia_prize
    """, _conn_fabric)
    df = df.rename(columns={
        "empresa":            "Empresa",
        "fundo":              "Fundo",
        "fecha_hora":         "Fecha-Hora",
        "temp_alta_c":        "TempAlta-C",
        "temp_baja_c":        "TempBaja-C",
        "et_mm":              "ET-mm",
        "rad_solar_wm2":      "RadSolar-W/m2",
        "rad_solar_alta_wm2": "RadSolarAlta-W/m2",
    })
    return df


@st.cache_data(ttl=3600, show_spinner=False)
def cargar_et_mensual_promedio(_conn_fabric) -> pd.DataFrame:
    """Carga ET promedio diario mensual por fundo desde Fabric."""
    try:
        query = """
        WITH base AS (
            SELECT
                fundo,
                EXTRACT(YEAR  FROM fecha_hora::date) AS anio,
                EXTRACT(MONTH FROM fecha_hora::date) AS mes,
                SUM(et_mm)                           AS eto_mensual,
                COUNT(DISTINCT fecha_hora::date)     AS dias_con_datos
            FROM public.rpt_climatologia_prize
            WHERE fecha_hora IS NOT NULL
            GROUP BY fundo,
                     EXTRACT(YEAR  FROM fecha_hora::date),
                     EXTRACT(MONTH FROM fecha_hora::date)
        )
        SELECT
            fundo AS "Fundo",
            mes   AS "Mes",
            ROUND(AVG(eto_mensual * 1.0 / dias_con_datos)::numeric, 2) AS "EToPromedioDiaria"
        FROM base
        GROUP BY fundo, mes
        ORDER BY fundo, mes
        """
        return pd.read_sql(query, _conn_fabric)
    except Exception as e:
        st.error(f"Error al cargar ET mensual: {e}")
        return pd.DataFrame()
