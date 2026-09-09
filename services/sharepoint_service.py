import io
from pathlib import Path
from urllib.parse import quote

import msal
import requests
import streamlit as st

CLIENT_ID  = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
AUTHORITY  = "https://login.microsoftonline.com/aquanqape.onmicrosoft.com"
SCOPES     = ["Files.Read.All", "Sites.Read.All"]
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
    if not cuentas:
        return None
    result = app.acquire_token_silent(SCOPES, account=cuentas[0], force_refresh=True)
    if result and "access_token" in result:
        _guardar_cache(cache)
        return result["access_token"]
    return None


def iniciar_device_flow() -> dict:
    """Inicia el flujo de dispositivo para autenticación interactiva."""
    cache = _cargar_cache()
    app   = _build_app(cache)
    return app.initiate_device_flow(scopes=SCOPES)


def completar_device_flow(flow: dict) -> str | None:
    """Completa el device flow; retorna access_token si fue exitoso."""
    cache = _cargar_cache()
    app   = _build_app(cache)
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        _guardar_cache(cache)
        return result["access_token"]
    return None


# ── Descarga desde SharePoint ──────────────────────────────────

@st.cache_data(ttl=60, show_spinner=False)
def _get_site_id(token: str) -> str:
    headers = {"Authorization": f"Bearer {token}"}
    url  = f"{GRAPH_URL}/sites/{SP_SITE}"
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json()["id"]


def descargar_excel_sp(token: str, nombre: str, ruta: str) -> bytes:
    """Descarga un archivo Excel de SharePoint y retorna sus bytes."""
    headers  = {"Authorization": f"Bearer {token}"}
    site_id  = _get_site_id(token)
    encoded  = quote(ruta)
    url      = f"{GRAPH_URL}/sites/{site_id}/drive/root:/{encoded}:/content"
    resp     = requests.get(url, headers=headers, timeout=120)
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
