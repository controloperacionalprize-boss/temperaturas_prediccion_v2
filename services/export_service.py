import io

import pandas as pd


def exportar_excel_tabla(df: pd.DataFrame, sheet_name: str = 'Datos') -> bytes:
    """Exporta una tabla plana a un Excel de una sola hoja."""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    out.seek(0)
    return out.read()
