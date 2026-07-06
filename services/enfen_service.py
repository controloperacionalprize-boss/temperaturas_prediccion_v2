import json
import re

import streamlit as st

try:
    import requests
    REQUESTS_DISPONIBLE = True
except ImportError:
    REQUESTS_DISPONIBLE = False

from OPERACIONES.PREDICCION_TEMPERATURA.config.config import AJUSTE_ENFEN, ENFEN_URL, ENFEN_CACHE_FILE


def obtener_ajuste_enfen(mes: int, anio: int, variable: str) -> float:
    return AJUSTE_ENFEN.get((mes, anio), {}).get(variable, 0.0)


@st.cache_data(ttl=3600 * 12, show_spinner=False)
def chequear_comunicado_enfen():
    """Revisa la página de comunicados ENFEN y extrae el último número/fecha."""
    if not REQUESTS_DISPONIBLE:
        return None
    try:
        resp = requests.get(ENFEN_URL, timeout=10, headers={'User-Agent': 'Mozilla/5.0'})
        resp.raise_for_status()
        texto = resp.text

        match = re.search(
            r'Comunicado\s+Oficial\s+Enfen\s*N[°º]\s*(\d+)\s*-\s*(\d{4}).{0,60}?'
            r'(\d{1,2}\s+\w+,?\s+\d{4})',
            texto, re.IGNORECASE | re.DOTALL
        )
        if match:
            numero, anio, fecha = match.group(1), match.group(2), match.group(3).strip()
            return {'numero': numero, 'anio': anio, 'fecha': fecha, 'id': f"{numero}-{anio}"}
    except Exception:
        return None
    return None


def _leer_ultimo_visto_enfen() -> dict:
    if ENFEN_CACHE_FILE.exists():
        try:
            return json.loads(ENFEN_CACHE_FILE.read_text())
        except Exception:
            return {'id': None}
    return {'id': None}


def _guardar_ultimo_visto_enfen(comunicado: dict) -> None:
    ENFEN_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    ENFEN_CACHE_FILE.write_text(json.dumps(comunicado))
