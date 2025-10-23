import streamlit as st
import pandas as pd
import re
from io import BytesIO

# (Opcional) Solo si tu app principal NO llama set_page_config:
try:
    st.set_page_config(page_title="Reporte de Cuentas", layout="wide")
except Exception:
    pass

st.title("Reporte de Cuentas → Limpieza automática")
st.caption(
    "Sube el Excel (.xls o .xlsx) tal como lo descargas. "
    "La página eliminará encabezados/hipervínculos/sumarios, propagará la cuenta y te dará un archivo limpio."
)

def _read_excel_any(uploaded_file):
    """
    Intenta leer .xls/.xlsx con o sin xlrd. Siempre como tabla cruda (header=None).
    """
    # Primero probamos con xlrd (útil para .xls antiguos).
    try:
        return pd.read_excel(uploaded_file, sheet_name=0, engine="xlrd", header=None)
    except Exception:
        # Fallback al engine por defecto
        return pd.read_excel(uploaded_file, sheet_name=0, header=None)

def _guess_header(df: pd.DataFrame):
    """
    Detecta la fila de encabezados buscando 'Cuenta' y 'Concepto' y también 'Saldo/Cargos/Abonos'
    en las primeras ~10 filas. Devuelve (df_sin_header, columnas_asignadas).
    Si no encuentra, asigna columnas por defecto.
    """
    header_idx = None
    for i in range(min(10, len(df))):
        row_vals = df.iloc[i].astype(str).str.strip().tolist()
        row_join = " ".join(v for v in row_vals if v and v.lower() != "nan")
        if re.search(r"Cuenta.*Concepto", row_join, re.IGNORECASE) and re.search(
            r"Saldo|Cargos|Abonos", row_join, re.IGNORECASE
        ):
            header_idx = i
            break

    if header_idx is not None:
        new_cols = df.iloc[header_idx].astype(str).str.strip().tolist()
        new_cols = [c if c and c.lower() != "nan" else f"col{j}" for j, c in enumerate(new_cols)]
        df2 = df.iloc[header_idx + 1 :].reset_index(drop=True)
        df2.columns = new_cols[: df2.shape[1]]
        return df2, df2.columns.tolist()

    # Fallback
    base_cols = ["Cuenta / Concepto", "Cheque", "Trafico", "Factura", "Fecha", "Cargos", "Abonos", "Saldo"]
    cols = base_cols + [f"col{j}" for j in range(len(base_cols), df.shape[1])]
    df2 = df.copy()
    df2.columns = cols[: df.shape[1]]
    return df2, df2.columns.tolist()

def _find_cuenta_concepto_col(colnames):
    for c in colnames:
        if re.search(r"cuenta.*concepto", str(c), re.IGNORECASE):
            return c
    # fallback: primera no-auxiliar razonable
    return colnames[3] if len(colnames) > 3 else colnames[-1]

def process_report(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Implementa la transformación descrita por Sasha.
    """
    df = df_raw.copy()

    # 1) Eliminar siempre la primera fila (banner)
    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    # 2) Detectar encabezados
    df, cols = _guess_header(df)

    # 3) Agregar 3 columnas al inicio (Cuenta, Aux1, Aux2)
    df.insert(0, "Cuenta", "")
    df.insert(1, "Aux1", "")
    df.insert(2, "Aux2", "")

    # 4) Propagar cuenta y marcar filas a eliminar
    cuenta_pat = re.compile(
        r"^\s*\d{3}-\d{2}-\d{2}-\d{3}-\d{2}-\d{3}-\d{4}\s+-\s+.+",  # 200-02-01-200-20-001-0001 - TEXTO
        re.IGNORECASE,
    )
    col_cc = _find_cuenta_concepto_col(df.columns)

    last_cuenta = None
    rows_to_drop = []

    for idx, val in df[col_cc].astype(str).items():
        text = val.strip()

        # renglones que siempre quitamos
        if text == "" or text.lower() == "saldo" or text.lower().startswith("sumas totales"):
            rows_to_drop.append(idx)
            continue

        # encabezado de cuenta
        if cuenta_pat.match(text):
            last_cuenta = text
            rows_to_drop.append(idx)  # quitamos el encabezado
        else:
            # detalle
            df.at[idx, "Cuenta"] = last_cuenta if last_cuenta else "__SIN_CUENTA_DETECTADA__"

    # 5) Eliminar filas no deseadas
    df_clean = df.drop(index=rows_to_drop).reset_index(drop=True)

    # 6) Quitar filas completamente vacías (ignorando Aux1/Aux2/Cuenta)
    non_aux_cols = [c for c in df_clean.columns if c not in ["Cuenta", "Aux1", "Aux2"]]
    def _row_is_empty(series):
        return all((str(v).strip() == "" or str(v).strip().lower() == "nan") for v in series.values)

    df_clean["__empty__"] = df_clean[non_aux_cols].astype(str).apply(_row_is_empty, axis=1)
    df_clean = df_clean.loc[~df_clean["__empty__"]].drop(columns="__empty__").reset_index(drop=True)

    # 7) Tipos (Fecha y montos)
    for col in df_clean.columns:
        if re.search(r"fecha", str(col), re.IGNORECASE):
            try:
                df_clean[col] = pd.to_datetime(df_clean[col], errors="coer_
