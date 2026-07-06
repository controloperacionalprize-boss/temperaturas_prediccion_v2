import io

import pandas as pd


def exportar_excel_prediccion(df_pred: pd.DataFrame) -> bytes:
    """Exporta la tabla plana de predicción (Fecha, Tmax, Tmin, Tipo) a Excel."""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df_pred.to_excel(writer, index=False, sheet_name='Prediccion')
    out.seek(0)
    return out.read()
