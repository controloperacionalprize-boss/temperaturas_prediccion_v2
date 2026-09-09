import io
import logging
from pathlib import Path
from urllib.parse import quote

import msal
import requests
import streamlit as st

logger = logging.getLogger(__name__)

CLIENT_ID  = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
AUTHORITY  = "https://login.microsoftonline.com/aquanqape.onmicrosoft.com"
SCOPES     = ["https://graph.microsoft.com/.default"]
GRAPH_URL  = "https://graph.microsoft.com/v1.0"
SP_SITE    = "aquanqape.sharepoint.com:/sites/OficinasPrizePeru"

_ASSETS_DIR      = Path(__file__).parent.parent / "assets"
TOKEN_CACHE_FILE = _ASSETS_DIR / ".msal_token_cache_sp.json"

# Archivos de SharePoint: (nombre, ruta_en_drive, hoja_excel)
SP_ARCHIVOS = [
    (
        "Metereologia_Prize.xlsx",
        "PowerBI Global/04.-BI-Division Administrativo/11.-Control Operacional"
        "/4.-Reporte Clima/Metereologia_Prize.xlsx",
        "Prize_Climatology",
    ),
    (
        "Metereologia_Prize_V2.xlsx",
        "PowerBI Global/04.-BI-Division Administrativo/11.-Control Operacional"
        "/4.-Reporte Clima/Metereologia_Prize_V2.xlsx",
        "Prize_Climatology_v2",
    ),
]


# ── Token cache ────────────────────────────────────────────────

def _cargar_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text(encoding="utf-8"))
    return cache


def _guardar_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        TOKEN_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")


def _build_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )


# ── Autenticación ──────────────────────────────────────────────

def get_sp_token_silente() -> str | None:
    """Intenta obtener un token Graph silenciosamente desde el caché."""
    cache = _cargar_cache()
    app   = _build_app(cache)
    cuentas = app.get_accounts()
    st.toast(f"🔍 Cache existe: {TOKEN_CACHE_FILE.exists()} | Cuentas: {len(cuentas)}")
    if not cuentas:
        return None
    st.toast(f"👤 Cuenta: {cuentas[0].get('username', '?')}")
    result = app.acquire_token_silent(SCOPES, account=cuentas[0], force_refresh=True)
    if result and "access_token" in result:
        _guardar_cache(cache)
        token = result["access_token"]
        st.toast(f"✅ Token silente OK (len={len(token)}, scopes={result.get('scope', '?')})")
        return token
    error = result.get("error", "?") if result else "None"
    error_desc = result.get("error_description", "") if result else ""
    st.toast(f"❌ Token silente falló: {error} - {error_desc[:100]}")
    return None


def iniciar_device_flow() -> dict:
    """Inicia el flujo de dispositivo para autenticación interactiva."""
    cache = _cargar_cache()
    app   = _build_app(cache)
    flow = app.initiate_device_flow(scopes=SCOPES)
    st.toast(f"🔄 Device flow iniciado: {'user_code' in flow}")
    return flow


def completar_device_flow(flow: dict) -> str | None:
    """Completa el device flow; retorna access_token si fue exitoso."""
    cache = _cargar_cache()
    app   = _build_app(cache)
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        _guardar_cache(cache)
        token = result["access_token"]
        st.toast(f"✅ Device flow OK (len={len(token)}, scopes={result.get('scope', '?')})")
        return token
    error = result.get("error", "?")
    error_desc = result.get("error_description", "")
    st.toast(f"❌ Device flow falló: {error} - {error_desc[:200]}")
    return None


# ── Descarga desde SharePoint ──────────────────────────────────

@st.cache_data(ttl=60, show_spinner=False)
def _get_site_id(token: str) -> str:
    headers = {"Authorization": f"Bearer {token}"}
    url  = f"{GRAPH_URL}/sites/{SP_SITE}"
    st.toast(f"🌐 GET site_id: {url}")
    resp = requests.get(url, headers=headers, timeout=30)
    st.toast(f"📡 Site response: {resp.status_code}")
    if resp.status_code != 200:
        st.toast(f"❌ Site error body: {resp.text[:300]}")
    resp.raise_for_status()
    site_id = resp.json()["id"]
    st.toast(f"✅ Site ID: {site_id[:50]}...")
    return site_id


def descargar_excel_sp(token: str, nombre: str, ruta: str) -> bytes:
    """Descarga un archivo Excel de SharePoint y retorna sus bytes."""
    headers  = {"Authorization": f"Bearer {token}"}
    site_id  = _get_site_id(token)
    encoded  = quote(ruta)
    url      = f"{GRAPH_URL}/sites/{site_id}/drive/root:/{encoded}:/content"
    st.toast(f"📥 Descargando {nombre}: {url[:100]}...")
    resp     = requests.get(url, headers=headers, timeout=120)
    st.toast(f"📡 {nombre}: status={resp.status_code}, len={len(resp.content)}")
    if resp.status_code != 200:
        st.toast(f"❌ {nombre} error: {resp.text[:300]}")
    resp.raise_for_status()
    return resp.content


@st.cache_data(ttl=3600, show_spinner=False)
def descargar_todos_excels(token: str) -> list[tuple[str, bytes, str]]:
    """Descarga ambos Excels de SharePoint. Retorna [(nombre, bytes, hoja), ...]."""
    resultados = []
    for nombre, ruta, hoja in SP_ARCHIVOS:
        try:
            datos = descargar_excel_sp(token, nombre, ruta)
            resultados.append((nombre, datos, hoja))
        except requests.exceptions.HTTPError as e:
            if e.response is not None and e.response.status_code in (401, 403):
                raise
            st.warning(f"⚠️ No se pudo descargar '{nombre}': {e}")
        except Exception as e:
            st.warning(f"⚠️ No se pudo descargar '{nombre}': {e}")
    return resultados
