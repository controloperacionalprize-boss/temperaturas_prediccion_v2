import json
import logging
import time
from pathlib import Path
from urllib.parse import quote, urlencode

import requests
import streamlit as st

logger = logging.getLogger(__name__)

CLIENT_ID  = "d3590ed6-52b3-4102-aeff-aad2292ab01c"   # Microsoft Office
AUTHORITY  = "https://login.microsoftonline.com/aquanqape.onmicrosoft.com"
SCOPE      = "https://graph.microsoft.com/.default offline_access"
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


# ── Token cache (JSON simple) ─────────────────────────────────

def _cargar_tokens() -> dict:
    if TOKEN_CACHE_FILE.exists():
        try:
            return json.loads(TOKEN_CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _guardar_tokens(data: dict) -> None:
    TOKEN_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_CACHE_FILE.write_text(json.dumps(data), encoding="utf-8")


# ── Autenticación via device flow (sin MSAL) ──────────────────

def get_sp_token_silente() -> str | None:
    """Intenta refrescar el token con el refresh_token guardado."""
    tokens = _cargar_tokens()
    refresh_token = tokens.get("refresh_token")
    if not refresh_token:
        st.toast("🔍 No hay refresh token en cache")
        return None

    st.toast("🔄 Intentando refresh token...")
    resp = requests.post(f"{AUTHORITY}/oauth2/v2.0/token", data={
        "client_id": CLIENT_ID,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": SCOPE,
    }, timeout=30)

    data = resp.json()
    if "access_token" in data:
        _guardar_tokens({
            "access_token": data["access_token"],
            "refresh_token": data.get("refresh_token", refresh_token),
        })
        st.toast(f"✅ Token refreshed OK (len={len(data['access_token'])})")
        return data["access_token"]

    st.toast(f"❌ Refresh falló: {data.get('error', '?')} - {data.get('error_description', '')[:100]}")
    TOKEN_CACHE_FILE.unlink(missing_ok=True)
    return None


def iniciar_device_flow() -> dict:
    """Inicia el device flow directamente con la API de Microsoft."""
    resp = requests.post(f"{AUTHORITY}/oauth2/v2.0/devicecode", data={
        "client_id": CLIENT_ID,
        "scope": SCOPE,
    }, timeout=30)
    data = resp.json()
    st.toast(f"🔄 Device flow: {'user_code' in data}")
    if not resp.ok or "device_code" not in data:
        return {"error": data.get("error_description", "Error al conectar con Microsoft")}
    return {
        "device_code": data["device_code"],
        "user_code": data["user_code"],
        "verification_uri": data.get("verification_uri", "https://microsoft.com/devicelogin"),
        "expires_in": data.get("expires_in", 900),
        "interval": data.get("interval", 5),
    }


def completar_device_flow(flow: dict) -> str | None:
    """Hace polling hasta que el usuario complete el login."""
    device_code = flow.get("device_code")
    interval = flow.get("interval", 5)
    expires_in = flow.get("expires_in", 900)
    deadline = time.time() + expires_in

    while time.time() < deadline:
        resp = requests.post(f"{AUTHORITY}/oauth2/v2.0/token", data={
            "client_id": CLIENT_ID,
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": device_code,
        }, timeout=30)
        data = resp.json()

        if data.get("error") in ("authorization_pending", "slow_down"):
            time.sleep(interval)
            continue

        if "access_token" in data:
            _guardar_tokens({
                "access_token": data["access_token"],
                "refresh_token": data.get("refresh_token", ""),
            })
            st.toast(f"✅ Login OK (scopes={data.get('scope', '?')[:80]})")
            return data["access_token"]

        st.toast(f"❌ Device flow error: {data.get('error', '?')} - {data.get('error_description', '')[:150]}")
        return None

    st.toast("❌ Device flow expiró")
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
