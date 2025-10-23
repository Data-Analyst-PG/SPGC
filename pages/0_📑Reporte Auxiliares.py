# pages/10_Reporte_de_Cuentas.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO

try:
    st.set_page_config(page_title="Reporte de Cuentas", layout="wide")
except Exception:
    pass

st.title("Reporte de Cuentas → Limpieza automática")
st.caption(
    "Sube el Excel (.xls o .xlsx) tal como lo descargas. "
    "La página eliminará encabezados/sumarios, propagará la cuenta y te dará un archivo limpio."
)

def _read_excel_any(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, sheet_name=0, engine="xlrd", header=None)
    except Exception:
        return pd.read_excel(uploaded_file, sheet_name=0, header=None)

def _guess_header(df):
    header_idx = None
    limit = min(10, len(df))
    for i in range(limit):
        row_vals = df.iloc[i].astype(str).str.strip().tolist()
        row_join = " ".join([v for v in row_vals if v and v.lower() != "nan"])
        if re.search(r"Cuenta.*Concepto", row_join, re.IGNORECASE) and re.search(
            r"Saldo|Cargos|Abonos", row_join, re.IGNORECASE
        ):
            header_idx = i
            break
    if header_idx is not None:
        new_cols = df.iloc[header_idx].astype(str).str.strip().tolist()
        new_cols = [c if c and c.lower() != "nan" else "col{}".format(j) for j, c in enumerate(new_cols)]
        df2 = df.iloc[header_idx + 1 :].reset_index(drop=True)
        df2.columns = new_cols[: df2.shape[1]]
        return df2, list(df2.columns)
    base_cols = ["Cuenta / Concepto", "Cheque", "Trafico", "Factura", "Fecha", "Cargos", "Abonos", "Saldo"]
    cols = base_cols + ["col{}".format(j) for j in range(len(base_cols), df.shape[1])]
    df2 = df.copy()
    df2.columns = cols[: df.shape[1]]
    return df2, list(df2.columns)

def _find_cuenta_concepto_col(colnames):
    for c in colnames:
        if re.search(r"cuenta.*concepto", str(c), re.IGNORECASE):
            return c
    return colnames[3] if len(colnames) > 3 else colnames[-1]

def process_report(df_raw):
    df = df_raw.copy()
    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)
    df, _ = _guess_header(df)
    df.insert(0, "Cuenta", "")
    df.insert(1, "Aux1", "")
    df.insert(2, "Aux2", "")
    cuenta_pat = re.compile(r"^\s*\d{3}-\d{2}-\d{2}-\d{3}-\d{2}-\d{3}-\d{4}\s+-\s+.+", re.IGNORECASE)
    col_cc = _find_cuenta_concepto_col(df.columns)
    last_cuenta = None
    rows_to_drop = []
    for idx, val in df[col_cc].astype(str).items():
        text = val.strip()
        if text == "" or text.lower() == "saldo" or text.lower().startswith("sumas totales"):
            rows_to_drop.append(idx)
            continue
        if cuenta_pat.match(text):
            last_cuenta = text
            rows_to_drop.append(idx)
        else:
            df.at[idx, "Cuenta"] = last_cuenta if last_cuenta else "__SIN_CUENTA_DETECTADA__"
    df_clean = df.drop(index=rows_to_drop).reset_index(drop=True)
    non_aux_cols = [c for c in df_clean.columns if c not in ["Cuenta", "Aux1", "Aux2"]]
    def _row_is_empty(series):
        for v in series.values:
            s = str(v).strip().lower()
            if s != "" and s != "nan":
                return False
        return True
    df_clean["__empty__"] = df_clean[non_aux_cols].astype(str).apply(_row_is_empty, axis=1)
    df_clean = df_clean.loc[~df_clean["__empty__"]].drop(columns="__empty__").reset_index(drop=True)
    for col in df_clean.columns:
        if re.search(r"fecha", str(col), re.IGNORECASE):
            try:
                df_clean[col] = pd.to_datetime(df_clean[col], errors="coerce").dt.date
            except Exception:
                pass
        if re.search(r"(cargos|abonos|saldo)", str(col), re.IGNORECASE):
            try:
                df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce")
            except Exception:
                pass
    return df_clean

uploaded = st.file_uploader("Sube el archivo (.xls o .xlsx)", type=["xls", "xlsx"], accept_multiple_files=False)

if uploaded is None:
    st.info("Esperando archivo…")
else:
    try:
        df0 = _read_excel_any(uploaded)
        df_clean = process_report(df0)
        st.success("Listo. Filas finales: {:,}".format(len(df_clean)))
        st.dataframe(df_clean.head(500), use_container_width=True)
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_clean.to_excel(writer, index=False, sheet_name="REPORTE")
        st.download_button(
            "Descargar Excel procesado",
            data=buffer.getvalue(),
            file_name="Reporte_procesado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error("Ocurrió un error procesando el archivo: {}".format(e))
        st.exception(e)
