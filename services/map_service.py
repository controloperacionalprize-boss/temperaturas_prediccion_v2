import re

import numpy as np
import pandas as pd
import streamlit as st
import folium

from OPERACIONES.PREDICCION_TEMPERATURA.config.config import (
    RIESGO_COLOR, RIESGO_COLOR_ET, COLORES_ET, COLORES_TEMP_SENAMHI,
    DIAS_ET_PROMEDIO, DIAS_RIESGO_VENTANA, MESES_RAW,
)
from services.utils import _normalizar


# ── KMZ / polígonos ────────────────────────────────────────────

@st.cache_data(show_spinner="Cargando polígonos KMZ…")
def load_kmz_bytes(_kmz_bytes: bytes) -> list:
    """Parsea polígonos de módulos AQ1/AQ2 desde un KMZ (en bytes)."""
    import zipfile
    import io
    from lxml import etree

    try:
        with zipfile.ZipFile(io.BytesIO(_kmz_bytes)) as kmz:
            kml_files = [f for f in kmz.namelist() if f.lower().endswith('.kml')]
            if not kml_files:
                return []
            kml_content = kmz.read(kml_files[0])

        parser = etree.XMLParser(recover=True, encoding='utf-8')
        root   = etree.fromstring(kml_content, parser=parser)
        nsmap  = {'kml': 'http://www.opengis.net/kml/2.2'}

        folders = root.xpath('.//kml:Folder', namespaces=nsmap) \
                  or root.xpath('.//*[local-name()="Folder"]')

        target_folders = []
        for folder in folders:
            fname_xpath = folder.xpath('.//kml:name/text()', namespaces=nsmap) \
                          or folder.xpath('.//*[local-name()="name"]/text()')
            if fname_xpath:
                fname = fname_xpath[0].upper()
                if ('AQ1' in fname or 'AQ2' in fname) and 'MODULO' in fname:
                    target_folders.append((folder, fname_xpath[0]))

        polygons = []
        for target_folder, folder_name in target_folders:
            placemarks = target_folder.xpath('.//kml:Placemark', namespaces=nsmap) \
                         or target_folder.xpath('.//*[local-name()="Placemark"]')

            fundo_match = re.search(r'(AQ\d+)', folder_name, re.IGNORECASE)
            fundo_aq    = fundo_match.group(1).upper() if fundo_match else None
            mod_match   = re.search(r'MODULO\s*0*(\d+)', folder_name, re.IGNORECASE)
            mod_n       = int(mod_match.group(1)) if mod_match else None

            if not fundo_aq or not mod_n:
                continue

            for placemark in placemarks:
                name_xpath = placemark.xpath('.//kml:name/text()', namespaces=nsmap) \
                              or placemark.xpath('.//*[local-name()="name"]/text()')
                name = name_xpath[0].strip() if name_xpath else ""

                coord_xpath = placemark.xpath('.//kml:Polygon//kml:coordinates/text()', namespaces=nsmap) \
                              or placemark.xpath('.//*[local-name()="Polygon"]//*[local-name()="coordinates"]/text()')
                if not coord_xpath:
                    continue

                coords = []
                for pair in coord_xpath[0].strip().split():
                    parts = pair.split(',')
                    if len(parts) >= 2:
                        try:
                            lon, lat = float(parts[0]), float(parts[1])
                            coords.append([lat, lon])
                        except ValueError:
                            pass

                if len(coords) < 3:
                    continue

                polygons.append({
                    "name"    : name or f"Polígono {len(polygons)+1}",
                    "coords"  : coords,
                    "mod_n"   : mod_n,
                    "fundo_aq": fundo_aq,
                })
        return polygons

    except Exception:
        return []


@st.cache_data(show_spinner="Descargando KMZ desde GitHub…", ttl=3600)
def download_kmz_from_github() -> bytes | None:
    import urllib.request, urllib.error

    token   = st.secrets.get("GITHUB_TOKEN_KMZ", "")
    raw_url = (
        "https://raw.githubusercontent.com/"
        "controloperacionalprize-boss/CAMPO_RENDIMIENTO/"
        "main/MODULOS_PRIZE_PAIJAN.kmz"
    )
    api_url = (
        "https://api.github.com/repos/"
        "controloperacionalprize-boss/CAMPO_RENDIMIENTO/"
        "contents/MODULOS_PRIZE_PAIJAN.kmz"
    )

    for url, use_token in [(raw_url, False), (api_url, True)]:
        try:
            headers = {"Accept": "application/vnd.github.v3.raw"} if use_token else {}
            if token and use_token:
                headers["Authorization"] = f"token {token}"
            req  = urllib.request.Request(url, headers=headers)
            resp = urllib.request.urlopen(req, timeout=30)
            return resp.read()
        except urllib.error.HTTPError as e:
            st.write(f"❌ HTTP {e.code}: {url}")
        except urllib.error.URLError as e:
            st.write(f"❌ URL Error: {e.reason}")
        except Exception as e:
            st.write(f"❌ Error: {str(e)}")

    st.error("❌ No se pudo descargar KMZ desde GitHub")
    return None


@st.cache_data(show_spinner=False)
def disolver_modulos(_kmz_polygons: list) -> list:
    """Une los polígonos de cada (fundo_aq, mod_n) en un solo contorno."""
    from shapely.geometry import Polygon as ShPolygon
    from shapely.ops import unary_union

    grupos = {}
    for p in _kmz_polygons:
        key  = (p['fundo_aq'], p['mod_n'])
        anillo = [(lon, lat) for lat, lon in p['coords']]
        if len(anillo) < 3:
            continue
        try:
            poly = ShPolygon(anillo)
            if not poly.is_valid:
                poly = poly.buffer(0)
            grupos.setdefault(key, []).append(poly)
        except Exception:
            continue

    resultado = []
    for (fundo_aq, mod_n), polys in grupos.items():
        union = unary_union(polys)
        geoms = [union] if union.geom_type == 'Polygon' else list(union.geoms)
        for geom in geoms:
            if geom.is_empty:
                continue
            coords = [[lat, lon] for lon, lat in geom.exterior.coords]
            resultado.append({'fundo_aq': fundo_aq, 'mod_n': mod_n, 'coords': coords})
    return resultado


def asignar_fundo(fundo_aq: str, mod_n: int) -> str:
    if fundo_aq == 'AQ1':
        return 'Arena Azul'
    if fundo_aq == 'AQ2':
        if 1 <= mod_n <= 5:
            return 'Vivadis'
        if mod_n in (6, 7, 8, 9, 10, 11, 16, 17, 18):
            return 'Santa Teresa'
        return 'Ayllu Allpa'
    return 'Desconocido'


# ── Colormaps ──────────────────────────────────────────────────

def crear_colormap_temperatura(temps_dict: dict, variable: str, mes_sel: int):
    from branca.colormap import LinearColormap

    if not temps_dict:
        return None

    valores = np.array(list(temps_dict.values()))
    vmin    = np.floor(np.percentile(valores, 5))
    vmax    = np.ceil(np.percentile(valores, 95))
    if vmin == vmax:
        vmin -= 1; vmax += 1

    return LinearColormap(
        colors=COLORES_TEMP_SENAMHI,
        vmin=vmin, vmax=vmax,
        caption=f'{variable} climatología SENAMHI — {MESES_RAW[mes_sel-1]} [°C]'
    )


def crear_colormap_y_leyenda_et(et_por_fundo: dict):
    from branca.colormap import LinearColormap

    if not et_por_fundo:
        return None, None

    valores = np.array(list(et_por_fundo.values()))
    vmin    = float(np.floor(valores.min() * 10) / 10)
    vmax    = float(np.ceil(valores.max() * 10) / 10)
    if vmin == vmax:
        vmin = max(0.0, vmin - 0.5); vmax = vmax + 0.5

    colormap     = LinearColormap(colors=COLORES_ET, vmin=vmin, vmax=vmax,
                                  caption=f'ET suma semanal (últ. {DIAS_ET_PROMEDIO}d) [mm]')
    leyenda_data = {'vmin': vmin, 'vmax': vmax, 'fundos': sorted(et_por_fundo.items())}
    return colormap, leyenda_data


# ── Riesgo ─────────────────────────────────────────────────────

def calcular_umbrales_riesgo(media_clim) -> dict:
    p50, p75, p95 = np.percentile(media_clim, [50, 75, 95])
    return {'p50': float(p50), 'p75': float(p75), 'p95': float(p95)}


def clasificar_riesgo(valor, umbrales: dict) -> str:
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return 'Sin datos'
    if valor <= umbrales['p50']: return 'Bajo'
    elif valor <= umbrales['p75']: return 'Medio'
    elif valor <= umbrales['p95']: return 'Alto'
    else: return 'Muy alto'


def calcular_metricas_riesgo(forecasts_cache: dict, fundo: str, variable: str,
                              umbrales: dict, dias_ventana: int = DIAS_RIESGO_VENTANA) -> dict:
    cache_entry = forecasts_cache.get((fundo, variable), {}) if forecasts_cache else {}
    forecast    = cache_entry.get('forecast', pd.DataFrame())

    if forecast.empty:
        return {'valor_prox': None, 'nivel': 'Sin datos',
                'color': RIESGO_COLOR['Sin datos'], 'dias_ventana': None}

    fc          = forecast.sort_values('ds').reset_index(drop=True)
    ventana     = fc.head(dias_ventana)
    valor_prox  = round(float(ventana['yhat'].mean()), 1)
    nivel       = clasificar_riesgo(valor_prox, umbrales)

    return {'valor_prox': valor_prox, 'nivel': nivel,
            'color': RIESGO_COLOR.get(nivel, RIESGO_COLOR['Sin datos']),
            'dias_ventana': len(ventana)}


def calcular_riesgo_real_fundo(dia: pd.DataFrame, fundo: str, variable: str) -> dict:
    sub    = dia[dia['Fundo'] == fundo].sort_values('Fecha')
    serie  = sub[variable].dropna()

    if serie.empty:
        return {'valor_actual': None, 'fecha': None, 'nivel': 'Sin datos',
                'color': RIESGO_COLOR['Sin datos'], 'umbrales': None}

    umbrales     = calcular_umbrales_riesgo(serie.values)
    valor_actual = float(serie.iloc[-1])
    fecha_actual = sub.loc[serie.index[-1], 'Fecha']
    nivel        = clasificar_riesgo(valor_actual, umbrales)

    return {'valor_actual': round(valor_actual, 1), 'fecha': fecha_actual,
            'nivel': nivel, 'color': RIESGO_COLOR.get(nivel, RIESGO_COLOR['Sin datos']),
            'umbrales': umbrales}


@st.cache_data(show_spinner=False)
def calcular_riesgo_et_fundo(df_et_mensual: pd.DataFrame, fundo: str) -> dict:
    sub_et = df_et_mensual[df_et_mensual['Fundo'] == fundo]['EToPromedioDiaria'].dropna()

    if sub_et.empty:
        return {'valor_actual': None, 'nivel': 'Sin datos',
                'color': RIESGO_COLOR_ET['Sin datos'], 'umbrales': None}

    umbrales = {
        'p25': float(np.percentile(sub_et.values, 25)),
        'p50': float(np.percentile(sub_et.values, 50)),
        'p75': float(np.percentile(sub_et.values, 75)),
        'p95': float(np.percentile(sub_et.values, 95)),
    }
    valor_actual = float(sub_et.iloc[-1])
    nivel        = clasificar_riesgo(valor_actual, umbrales)

    return {'valor_actual': round(valor_actual, 2), 'nivel': nivel,
            'color': RIESGO_COLOR_ET.get(nivel, RIESGO_COLOR_ET['Sin datos']),
            'umbrales': umbrales}


def calcular_et_suma_semanal_fundos(dia: pd.DataFrame, dias: int = DIAS_ET_PROMEDIO) -> dict:
    resultado = {}
    for fundo in dia['Fundo'].unique():
        sub     = dia[dia['Fundo'] == fundo].sort_values('Fecha')
        et_vals = sub['ET'].dropna()
        et_vals = et_vals[et_vals > 0].tail(dias)
        if not et_vals.empty:
            resultado[fundo] = round(float(et_vals.sum()), 2)
    return resultado


# ── Leyendas Folium ────────────────────────────────────────────

def agregar_leyenda_vertical(m, vmin: float, vmax: float, colors: list, caption: str,
                              n_ticks: int = 6, position: str = 'topleft',
                              height_px: int = 260, width_px: int = 22,
                              font_caption: int = 13, font_ticks: int = 12,
                              decimales: int = 0) -> None:
    from branca.element import MacroElement
    from jinja2 import Template

    colores_ab = list(reversed(colors))
    n          = len(colores_ab)
    stops      = [f"{c} {i/(n-1)*100:.0f}%" for i, c in enumerate(colores_ab)]
    gradiente  = ", ".join(stops)

    ticks_html = ""
    for i in range(n_ticks):
        frac      = i / (n_ticks - 1)
        valor     = vmax - frac * (vmax - vmin)
        valor_str = f"{valor:.{decimales}f}"
        ticks_html += (
            f'<div style="position:absolute;left:{width_px + 6}px;top:{frac*100:.1f}%;'
            f'transform:translateY(-50%);font-size:{font_ticks}px;font-weight:600;'
            f'color:#263238;white-space:nowrap;">{valor_str}</div>'
        )

    leyenda = MacroElement()
    leyenda._template = Template(f"""
    {{% macro script(this, kwargs) %}}
    var legend_{{{{ this.get_name() }}}} = L.control({{position: '{position}'}});
    legend_{{{{ this.get_name() }}}}.onAdd = function (map) {{
        var div = L.DomUtil.create('div', 'info legend');
        div.style.background = 'rgba(255,255,255,0.88)';
        div.style.padding = '10px 32px 10px 10px';
        div.style.borderRadius = '10px';
        div.style.boxShadow = '0 2px 8px rgba(0,0,0,0.25)';
        div.style.fontFamily = "'Segoe UI', Arial, sans-serif";
        div.innerHTML = `
            <div style="font-size:{font_caption}px;font-weight:700;color:#263238;
                         margin-bottom:8px;max-width:130px;">{caption}</div>
            <div style="position:relative;width:{width_px}px;height:{height_px}px;
                         background:linear-gradient(to bottom, {gradiente});
                         border-radius:4px;border:1px solid #B0BEC5;">
                {ticks_html}
            </div>
        `;
        return div;
    }};
    legend_{{{{ this.get_name() }}}}.addTo({{{{ this._parent.get_name() }}}});
    {{% endmacro %}}
    """)
    m.add_child(leyenda)


def agregar_leyenda_riesgo(m, riesgos_por_fundo=None, unidad: str = '°C',
                            position: str = 'bottomleft', titulo: str = 'Riesgo térmico') -> None:
    from branca.element import MacroElement
    from jinja2 import Template

    filas_color = [
        ('Bajo',     RIESGO_COLOR['Bajo']),
        ('Medio',    RIESGO_COLOR['Medio']),
        ('Alto',     RIESGO_COLOR['Alto']),
        ('Muy alto', RIESGO_COLOR['Muy alto']),
    ]
    items_html = "".join(
        f'<div style="display:flex;align-items:center;gap:6px;margin:2px 0;">'
        f'<span style="width:11px;height:11px;border-radius:50%;background:{color};display:inline-block;"></span>'
        f'<span style="font-size:11px;color:#263238;">{nivel}</span></div>'
        for nivel, color in filas_color
    )

    leyenda = MacroElement()
    leyenda._template = Template(f"""
    {{% macro script(this, kwargs) %}}
    var legend_{{{{ this.get_name() }}}} = L.control({{position: '{position}'}});
    legend_{{{{ this.get_name() }}}}.onAdd = function (map) {{
        var div = L.DomUtil.create('div', 'info legend');
        div.style.background = 'rgba(255,255,255,0.88)';
        div.style.padding = '8px 12px';
        div.style.borderRadius = '8px';
        div.style.boxShadow = '0 2px 8px rgba(0,0,0,0.25)';
        div.style.fontFamily = "'Segoe UI', Arial, sans-serif";
        div.innerHTML = `
            <div style="font-size:11px;font-weight:700;color:#263238;margin-bottom:4px;">{titulo}</div>
            {items_html}
        `;
        return div;
    }};
    legend_{{{{ this.get_name() }}}}.addTo({{{{ this._parent.get_name() }}}});
    {{% endmacro %}}
    """)
    m.add_child(leyenda)


def agregar_leyenda_riesgo_et(m, position: str = 'bottomleft') -> None:
    from branca.element import MacroElement
    from jinja2 import Template

    filas_color = [
        ('Bajo',     RIESGO_COLOR_ET['Bajo']),
        ('Medio',    RIESGO_COLOR_ET['Medio']),
        ('Alto',     RIESGO_COLOR_ET['Alto']),
        ('Muy alto', RIESGO_COLOR_ET['Muy alto']),
    ]
    items_html = "".join(
        f'<div style="display:flex;align-items:center;gap:6px;margin:2px 0;">'
        f'<span style="width:11px;height:11px;border-radius:50%;background:{color};display:inline-block;border:1px solid white;"></span>'
        f'<span style="font-size:11px;color:#263238;font-weight:600;">{nivel}</span></div>'
        for nivel, color in filas_color
    )

    leyenda = MacroElement()
    leyenda._template = Template(f"""
    {{% macro script(this, kwargs) %}}
    var legend_{{{{ this.get_name() }}}} = L.control({{position: '{position}'}});
    legend_{{{{ this.get_name() }}}}.onAdd = function (map) {{
        var div = L.DomUtil.create('div', 'info legend');
        div.style.background = 'rgba(255,255,255,0.92)';
        div.style.padding = '8px 12px';
        div.style.borderRadius = '8px';
        div.style.boxShadow = '0 2px 8px rgba(0,0,0,0.25)';
        div.style.fontFamily = "'Segoe UI', Arial, sans-serif";
        div.innerHTML = `
            <div style="font-size:11px;font-weight:700;color:#00838F;margin-bottom:4px;">⚠️ Riesgo ET</div>
            {items_html}
        `;
        return div;
    }};
    legend_{{{{ this.get_name() }}}}.addTo({{{{ this._parent.get_name() }}}});
    {{% endmacro %}}
    """)
    m.add_child(leyenda)


def agregar_leyenda_et_discreta(m, et_por_fundo_sorted: list, vmin: float, vmax: float,
                                 position: str = 'topleft') -> None:
    from branca.element import MacroElement
    from jinja2 import Template

    filas_html = ""
    for fundo, valor in et_por_fundo_sorted:
        norm_val  = (valor - vmin) / (vmax - vmin) if vmax > vmin else 0.5
        idx_color = norm_val * (len(COLORES_ET) - 1)
        idx_bajo  = int(np.floor(idx_color))
        color     = COLORES_ET[min(idx_bajo, len(COLORES_ET)-1)]
        filas_html += (
            f'<div style="display:flex;align-items:center;gap:8px;margin:4px 0;">'
            f'<span style="width:14px;height:14px;border-radius:3px;background:{color};display:inline-block;border:1px solid white;"></span>'
            f'<span style="font-size:11px;color:#263238;font-weight:600;">{fundo}: <b>{valor:.2f}</b> mm</span></div>'
        )

    leyenda = MacroElement()
    leyenda._template = Template(f"""
    {{% macro script(this, kwargs) %}}
    var legend_{{{{ this.get_name() }}}} = L.control({{position: '{position}'}});
    legend_{{{{ this.get_name() }}}}.onAdd = function (map) {{
        var div = L.DomUtil.create('div', 'info legend');
        div.style.background = 'rgba(255,255,255,0.92)';
        div.style.padding = '10px 14px';
        div.style.borderRadius = '8px';
        div.style.boxShadow = '0 2px 8px rgba(0,0,0,0.25)';
        div.style.fontFamily = "'Segoe UI', Arial, sans-serif";
        div.innerHTML = `
            <div style="font-size:12px;font-weight:700;color:#00838F;margin-bottom:6px;">💧 ET Semanal</div>
            {filas_html}
            <div style="font-size:9px;color:#78909C;margin-top:6px;font-style:italic;">Últimos {DIAS_ET_PROMEDIO} días</div>
        `;
        return div;
    }};
    legend_{{{{ this.get_name() }}}}.addTo({{{{ this._parent.get_name() }}}});
    {{% endmacro %}}
    """)
    m.add_child(leyenda)


# ── Mapas Folium ───────────────────────────────────────────────

def generar_mapa_distritos(geojson, distrito_fijo: str, distrito_din: str | None,
                            variable: str, dia: pd.DataFrame | None = None,
                            mes_sel: int | None = None,
                            media_clim_fijo=None, media_clim_din=None,
                            label_clim2: str | None = None, modulos_kmz: list | None = None,
                            distritos_senamhi: set | None = None,
                            temps_distritos: dict | None = None, temp_colormap=None,
                            forecasts_cache: dict | None = None):
    if geojson is None:
        return None

    import folium
    from folium.plugins import Fullscreen

    color_fijo     = '#C62828' if variable == 'Tmax' else '#1565C0'
    color_din      = '#FF6F00'
    dist_fijo_norm = _normalizar(distrito_fijo)
    dist_din_norm  = _normalizar(distrito_din) if distrito_din else None

    m = folium.Map(location=[-7.27, -79.45], zoom_start=7, tiles=None,
                   control_scale=False, zoom_control=False)
    folium.TileLayer(
        tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
        attr='Esri WorldImagery', name='Satellite', overlay=False, control=False,
    ).add_to(m)
    folium.TileLayer(
        tiles='https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',
        attr='Esri Labels', name='Labels', overlay=True, control=False, opacity=0.7,
    ).add_to(m)
    Fullscreen(position='topright', title='Pantalla completa',
               title_cancel='Salir', force_separate_button=True).add_to(m)

    # ── Distritos SENAMHI con temperatura ─────────────────────
    if distritos_senamhi:
        for feat in geojson['features']:
            p       = feat['properties']
            nombre  = _normalizar(p['NOMBDIST'])
            if nombre == dist_fijo_norm or (dist_din_norm and nombre == dist_din_norm):
                continue
            if nombre not in distritos_senamhi:
                continue
            label_txt  = p['NOMBDIST'] + ' — ' + p.get('NOMBDEP', '')
            valor_temp = (temps_distritos or {}).get(nombre)
            if valor_temp is not None and temp_colormap is not None:
                fill_color   = temp_colormap(valor_temp)
                fill_opacity = 0.70
                borde        = '#FFFFFF'; grosor = 0.8
                tooltip_txt  = (
                    f"🌡️ {label_txt}<br>{variable} clim.: {valor_temp:.1f}°C "
                    f"({MESES_RAW[mes_sel-1] if mes_sel else ''})"
                )
            else:
                fill_color   = '#90A4AE'; fill_opacity = 0.06
                borde        = '#90A4AE'; grosor = 1
                tooltip_txt  = f"📡 {label_txt}"

            folium.GeoJson(
                feat,
                style_function=(lambda f, fc=fill_color, op=fill_opacity, bc=borde, w=grosor: {
                    'fillColor': fc, 'color': bc, 'weight': w,
                    'dashArray': None if op > 0.1 else '3,3', 'fillOpacity': op,
                }),
                highlight_function=lambda f: {'fillOpacity': 0.75, 'weight': 1.5},
                tooltip=folium.Tooltip(tooltip_txt, sticky=False),
            ).add_to(m)

        if temp_colormap is not None:
            agregar_leyenda_vertical(m, vmin=temp_colormap.vmin, vmax=temp_colormap.vmax,
                                     colors=COLORES_TEMP_SENAMHI,
                                     caption=temp_colormap.caption, position='bottomleft')

    # ── Distritos fijo / dinámico ──────────────────────────────
    def _style_fijo(f):
        return {'fillColor': color_fijo, 'color': 'white', 'weight': 1.5, 'fillOpacity': 0.55}
    def _style_din(f):
        return {'fillColor': color_din, 'color': 'white', 'weight': 1.5, 'fillOpacity': 0.55}

    for feat in geojson['features']:
        p       = feat['properties']
        nombre  = _normalizar(p['NOMBDIST'])
        es_fijo = nombre == dist_fijo_norm
        es_din  = dist_din_norm and nombre == dist_din_norm
        if not es_fijo and not es_din:
            continue

        label_txt = p['NOMBDIST'] + ' — ' + p.get('NOMBDEP', '')
        if mes_sel is not None:
            valor_clim = None
            if es_fijo and media_clim_fijo is not None:
                valor_clim = float(media_clim_fijo[mes_sel - 1])
            elif es_din and media_clim_din is not None:
                valor_clim = float(media_clim_din[mes_sel - 1])
            if valor_clim is not None:
                label_txt += f"<br>🌡️ {variable} clim.: {valor_clim:.1f}°C ({MESES_RAW[mes_sel-1]})"

        folium.GeoJson(feat, style_function=_style_fijo if es_fijo else _style_din,
                       tooltip=folium.Tooltip(label_txt, sticky=False)).add_to(m)

        geom   = feat['geometry']
        coords = []
        if geom['type'] == 'Polygon':
            coords = geom['coordinates'][0]
        elif geom['type'] == 'MultiPolygon':
            coords = geom['coordinates'][0][0]
        if coords:
            lon_c = sum(c[0] for c in coords) / len(coords)
            lat_c = sum(c[1] for c in coords) / len(coords)
            folium.Marker(
                location=[lat_c, lon_c],
                popup=folium.Popup(label_txt, max_width=200),
                icon=folium.Icon(color='red' if es_fijo else 'orange',
                                 icon='tint' if variable == 'Tmax' else 'info-sign',
                                 prefix='glyphicon'),
            ).add_to(m)

    # ── Módulos KMZ ───────────────────────────────────────────
    if modulos_kmz:
        from shapely.geometry import Polygon as ShPolygon
        from shapely.ops import unary_union

        for mod in modulos_kmz:
            mod['fundo'] = asignar_fundo(mod['fundo_aq'], mod['mod_n'])

        grupos_fundo = {}
        for mod in modulos_kmz:
            anillo = [(lon, lat) for lat, lon in mod['coords']]
            try:
                poly = ShPolygon(anillo)
                if not poly.is_valid: poly = poly.buffer(0)
                grupos_fundo.setdefault(mod['fundo'], []).append(poly)
            except Exception:
                continue

        COLOR_FUNDO = {
            'Arena Azul': '#1565C0', 'Vivadis': '#2E7D32',
            'Santa Teresa': '#F57F17', 'Ayllu Allpa': '#6A1B9A',
        }

        for fundo_nombre, polys in grupos_fundo.items():
            union      = unary_union(polys)
            geoms      = [union] if union.geom_type == 'Polygon' else list(union.geoms)
            centroide  = union.centroid
            lat_c, lon_c = centroide.y, centroide.x
            color_acento = COLOR_FUNDO.get(fundo_nombre, '#37474F')

            riesgo = (
                calcular_riesgo_real_fundo(dia, fundo_nombre, variable)
                if dia is not None else
                {'valor_actual': None, 'fecha': None, 'nivel': 'Sin datos',
                 'color': RIESGO_COLOR['Sin datos'], 'umbrales': None}
            )
            color_riesgo = riesgo['color']
            nivel        = riesgo['nivel']
            valor_txt    = f"{riesgo['valor_actual']:.1f}°C" if riesgo['valor_actual'] is not None else "—"
            fecha_txt    = riesgo['fecha'].strftime('%d/%b') if riesgo['fecha'] is not None else ""

            tooltip_html = (
                f"<b>{fundo_nombre}</b><br>"
                f"{variable} ({fecha_txt}): <b>{valor_txt}</b><br>"
                f"<span style='color:{color_riesgo};font-weight:700;'>Riesgo: {nivel.upper()}</span>"
            )
            popup_html = f"<b>{fundo_nombre}</b><br>"
            if riesgo['valor_actual'] is not None:
                popup_html += f"{variable} ({fecha_txt}): <b>{riesgo['valor_actual']:.1f}°C</b><br>"
            if riesgo['umbrales'] is not None:
                u = riesgo['umbrales']
                popup_html += f"<small>P50={u['p50']:.1f} | P75={u['p75']:.1f} | P95={u['p95']:.1f}</small><br>"
            popup_html += f"<b>Riesgo: <span style='color:{color_riesgo}'>{nivel.upper()}</span></b>"

            for geom in geoms:
                if geom.is_empty: continue
                coords = [[lat, lon] for lon, lat in geom.exterior.coords]
                folium.Polygon(locations=coords, color='white', weight=2.5, fill=True,
                               fill_color=color_riesgo, fill_opacity=0.60,
                               tooltip=folium.Tooltip(tooltip_html, sticky=False),
                               popup=folium.Popup(popup_html, max_width=250)).add_to(m)

            folium.map.Marker(
                [lat_c, lon_c],
                icon=folium.DivIcon(
                    icon_size=(0, 0), icon_anchor=(0, 0),
                    html=(
                        '<div style="transform:translate(-50%,-50%);display:inline-flex;align-items:center;gap:6px;'
                        'background:rgba(255,255,255,0.92);padding:4px 10px 4px 8px;border-radius:14px;'
                        f'border-left:4px solid {color_acento};box-shadow:0 2px 6px rgba(0,0,0,0.25);'
                        'font-family:"Segoe UI",Arial,sans-serif;font-size:12px;font-weight:700;'
                        f'color:#263238;letter-spacing:0.02em;white-space:nowrap;pointer-events:auto;">'
                        f'{fundo_nombre}</div>'
                    )
                ),
                tooltip=folium.Tooltip(tooltip_html, sticky=False),
                popup=folium.Popup(popup_html, max_width=250),
            ).add_to(m)

        agregar_leyenda_riesgo(m, position='bottomleft')

    return m


def generar_mapa_et(modulos_kmz: list | None, et_por_fundo: dict,
                    colormap_et=None, df_et_mensual: pd.DataFrame | None = None,
                    riesgos_et: dict | None = None):
    """Mapa con los módulos KMZ coloreados según ET por fundo."""
    import folium
    from folium.plugins import Fullscreen
    from shapely.geometry import Polygon as ShPolygon
    from shapely.ops import unary_union

    m = folium.Map(location=[-7.65, -79.36], zoom_start=12, tiles=None,
                   control_scale=False, zoom_control=False)
    folium.TileLayer(
        tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
        attr='Esri WorldImagery', overlay=False, control=False,
    ).add_to(m)
    folium.TileLayer(
        tiles='https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',
        attr='Esri Labels', overlay=True, control=False, opacity=0.7,
    ).add_to(m)
    Fullscreen(position='topright', title='Pantalla completa',
               title_cancel='Salir', force_separate_button=True).add_to(m)

    if not modulos_kmz:
        return m

    for mod in modulos_kmz:
        if 'fundo' not in mod:
            mod['fundo'] = asignar_fundo(mod['fundo_aq'], mod['mod_n'])

    grupos_fundo = {}
    for mod in modulos_kmz:
        anillo = [(lon, lat) for lat, lon in mod['coords']]
        try:
            poly = ShPolygon(anillo)
            if not poly.is_valid: poly = poly.buffer(0)
            grupos_fundo.setdefault(mod['fundo'], []).append(poly)
        except Exception:
            continue

    bounds_all = []

    for fundo_nombre, polys in grupos_fundo.items():
        union      = unary_union(polys)
        geoms      = [union] if union.geom_type == 'Polygon' else list(union.geoms)
        et_val     = et_por_fundo.get(fundo_nombre)
        centroide  = union.centroid
        et_txt     = f"{et_val:.2f} mm/día" if et_val is not None else "—"

        minx, miny, maxx, maxy = union.bounds
        bounds_all.append((miny, minx, maxy, maxx))

        if riesgos_et and fundo_nombre in riesgos_et:
            riesgo_et = riesgos_et[fundo_nombre]
        else:
            riesgo_et = calcular_riesgo_et_fundo(df_et_mensual, fundo_nombre) if df_et_mensual is not None else {}

        color_riesgo_et = riesgo_et.get('color', RIESGO_COLOR_ET['Sin datos'])
        nivel_et        = riesgo_et.get('nivel', 'Sin datos')
        valor_et_txt    = f"{riesgo_et['valor_actual']:.2f} mm/día" if riesgo_et.get('valor_actual') is not None else "—"

        tooltip_et = (
            f"<b>{fundo_nombre}</b><br>"
            f"💧 ET promedio: <b>{valor_et_txt}</b><br>"
            f"<span style='color:{color_riesgo_et};font-weight:700;'>Riesgo: {nivel_et.upper()}</span>"
        )

        popup_content = f"<b>{fundo_nombre}</b><br>"
        if et_val is not None:
            popup_content += f"💧 <b>ET semanal:</b> {et_val:.2f} mm<br>"
        popup_content += f"💧 <b>ET promedio:</b> {valor_et_txt}<br>"
        if riesgo_et.get('umbrales') is not None:
            u = riesgo_et['umbrales']
            popup_content += (
                f"<small>Q1={u.get('p25', 'N/A'):.2f} | Q2={u['p50']:.2f} | Q3={u['p75']:.2f}</small><br>"
            )
        popup_content += f"<b>Riesgo: <span style='color:{color_riesgo_et};'>{nivel_et.upper()}</span></b>"
        popup_html = (
            f'<div style="background:rgba(255,255,255,0.75);padding:8px;border-radius:4px;'
            f'font-family:Arial,sans-serif;">{popup_content}</div>'
        )

        for geom in geoms:
            if geom.is_empty: continue
            coords = [[lat, lon] for lon, lat in geom.exterior.coords]
            folium.Polygon(locations=coords, color='white', weight=2, fill=True,
                           fill_color=color_riesgo_et, fill_opacity=0.55,
                           tooltip=folium.Tooltip(tooltip_et, sticky=False),
                           popup=folium.Popup(popup_html, max_width=280)).add_to(m)

        folium.map.Marker(
            [centroide.y, centroide.x],
            icon=folium.DivIcon(
                icon_size=(0, 0), icon_anchor=(0, 0),
                html=(
                    f'<div style="font-size:10px;font-weight:700;color:#263238;'
                    f'text-align:center;pointer-events:none;">{fundo_nombre}</div>'
                )
            ),
            tooltip=folium.Tooltip(et_txt, sticky=False)
        ).add_to(m)

    if bounds_all:
        lat_min = min(b[0] for b in bounds_all); lon_min = min(b[1] for b in bounds_all)
        lat_max = max(b[2] for b in bounds_all); lon_max = max(b[3] for b in bounds_all)
        m.fit_bounds([[lat_min, lon_min], [lat_max, lon_max]])

    if colormap_et is not None and et_por_fundo:
        agregar_leyenda_et_discreta(m, sorted(et_por_fundo.items()),
                                    vmin=colormap_et.vmin, vmax=colormap_et.vmax, position='topleft')
    agregar_leyenda_riesgo_et(m, position='bottomleft')
    return m
def calcular_rad_promedio_fundos(dia: pd.DataFrame, dias: int = 7) -> dict:
    """Promedio de RadSolar de los últimos N días por fundo."""
    resultado = {}
    for fundo in dia['Fundo'].unique():
        sub = dia[dia['Fundo'] == fundo].sort_values('Fecha')
        vals = sub['RadSolar'].dropna().tail(dias)
        if not vals.empty:
            resultado[fundo] = round(float(vals.mean()), 1)
    return resultado


def generar_mapa_rad(modulos_kmz: list | None, rad_por_fundo: dict) -> 'folium.Map | None':
    """Mapa con módulos KMZ coloreados por RadSolar promedio."""
    import folium
    from folium.plugins import Fullscreen
    from shapely.geometry import Polygon as ShPolygon
    from shapely.ops import unary_union
    from branca.colormap import LinearColormap

    m = folium.Map(location=[-7.65, -79.36], zoom_start=12, tiles=None,
                   control_scale=False, zoom_control=False)
    folium.TileLayer(
        tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
        attr='Esri WorldImagery', overlay=False, control=False,
    ).add_to(m)
    folium.TileLayer(
        tiles='https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',
        attr='Esri Labels', overlay=True, control=False, opacity=0.7,
    ).add_to(m)
    Fullscreen(position='topright', title='Pantalla completa',
               title_cancel='Salir', force_separate_button=True).add_to(m)

    if not modulos_kmz or not rad_por_fundo:
        return m

    # Colormap amarillo → naranja → rojo
    valores = list(rad_por_fundo.values())
    vmin    = float(np.floor(min(valores)))
    vmax    = float(np.ceil(max(valores)))
    if vmin == vmax:
        vmin = max(0.0, vmin - 50); vmax = vmax + 50

    colormap = LinearColormap(
        colors=['#FFF9C4', '#FFB300', '#E65100'],
        vmin=vmin, vmax=vmax,
        caption='RadSolar promedio 7d [W/m²]'
    )

    for mod in modulos_kmz:
        if 'fundo' not in mod:
            mod['fundo'] = asignar_fundo(mod['fundo_aq'], mod['mod_n'])

    grupos_fundo = {}
    for mod in modulos_kmz:
        anillo = [(lon, lat) for lat, lon in mod['coords']]
        try:
            poly = ShPolygon(anillo)
            if not poly.is_valid:
                poly = poly.buffer(0)
            grupos_fundo.setdefault(mod['fundo'], []).append(poly)
        except Exception:
            continue

    bounds_all = []
    for fundo_nombre, polys in grupos_fundo.items():
        union     = unary_union(polys)
        geoms     = [union] if union.geom_type == 'Polygon' else list(union.geoms)
        centroide = union.centroid
        rad_val   = rad_por_fundo.get(fundo_nombre)
        rad_txt   = f"{rad_val:.1f} W/m²" if rad_val is not None else "—"

        fill_color = colormap(rad_val) if rad_val is not None else '#90A4AE'

        tooltip_html = (
            f"<b>{fundo_nombre}</b><br>"
            f"☀️ RadSolar prom. 7d: <b>{rad_txt}</b>"
        )

        minx, miny, maxx, maxy = union.bounds
        bounds_all.append((miny, minx, maxy, maxx))

        for geom in geoms:
            if geom.is_empty:
                continue
            coords = [[lat, lon] for lon, lat in geom.exterior.coords]
            folium.Polygon(
                locations=coords, color='white', weight=2.5,
                fill=True, fill_color=fill_color, fill_opacity=0.65,
                tooltip=folium.Tooltip(tooltip_html, sticky=False),
                popup=folium.Popup(tooltip_html, max_width=250),
            ).add_to(m)

        folium.map.Marker(
            [centroide.y, centroide.x],
            icon=folium.DivIcon(
                icon_size=(0, 0), icon_anchor=(0, 0),
                html=(
                    f'<div style="transform:translate(-50%,-50%);display:inline-flex;'
                    f'align-items:center;gap:4px;background:rgba(255,255,255,0.90);'
                    f'padding:3px 8px;border-radius:12px;border-left:4px solid #FFB300;'
                    f'box-shadow:0 2px 6px rgba(0,0,0,0.25);font-family:"Segoe UI",Arial,sans-serif;'
                    f'font-size:11px;font-weight:700;color:#263238;white-space:nowrap;">'
                    f'{fundo_nombre}<br><span style="color:#E65100;">{rad_txt}</span></div>'
                )
            ),
            tooltip=folium.Tooltip(tooltip_html, sticky=False),
        ).add_to(m)

    if bounds_all:
        lat_min = min(b[0] for b in bounds_all); lon_min = min(b[1] for b in bounds_all)
        lat_max = max(b[2] for b in bounds_all); lon_max = max(b[3] for b in bounds_all)
        m.fit_bounds([[lat_min, lon_min], [lat_max, lon_max]])

    agregar_leyenda_vertical(
        m, vmin=vmin, vmax=vmax,
        colors=['#FFF9C4', '#FFB300', '#E65100'],
        caption='RadSolar 7d [W/m²]',
        position='bottomleft', decimales=0
    )
    return m