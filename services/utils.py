import unicodedata
import pandas as pd


def _hex_to_rgba(hex_color: str, alpha: float) -> str:
    h = hex_color.lstrip('#')
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f'rgba({r},{g},{b},{alpha})'


def _normalizar(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ''
    s = str(s).strip()
    if s.lower() == 'nan' or s == '':
        return ''
    return unicodedata.normalize('NFKD', s).encode('ascii', errors='ignore').decode('utf-8').upper().strip()
