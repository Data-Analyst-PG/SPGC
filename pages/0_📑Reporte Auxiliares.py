import streamlit as st
import pandas as pd
import re
from io import BytesIO
import io

# =====================================================
# --- UTILIDADES BÁSICAS ---
# =====================================================

def _to_num_safe(x):
    """Convierte a número sin romper por comas o símbolos."""
    if pd.isna(x):
        return 0.0
    s = str(x).replace(",", "").replace("$", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def _read_excel_any(file):
    """Lee .xls/.xlsx/.html de STAR 1 o STAR 2.0 en DataFrame."""
    name = file.name.lower()
    if name.endswith((".xls", ".xlsx")):
        return pd.read_excel(file, header=None, dtype=str)
    elif name.endswith((".html", ".htm")):
        tables = pd.read_html(file, header=None, dtype=str)
        return tables[0]
    else:
        raise ValueError("Formato no soportado: debe ser .xls, .xlsx o .html")


# =====================================================
# --- DETECCIÓN DEL MODO ---
# =====================================================

def _detect_mode(df_raw: pd.DataFrame) -> str:
    """Detecta STAR 1 o STAR 2.0 según encabezados o celdas iniciales."""
    try:
        df_guess, _ = _guess_header(df_raw.copy())
    except Exception:
        df_guess = df_raw.copy()

    cols_norm = [str(c).strip().lower().replace("\xa0", " ") for c in df_guess.columns]

    # Señal directa
    if any(c == "poliza" for c in cols_norm):
        return "star2"

    # A2 con cuenta
    if len(df_guess.index) >= 1 and len(df_guess.columns) >= 1:
        a2 = str(df_guess.iloc[0, 0])
        if a2.replace("\xa0", " ").strip().startswith(":"):
            return "star2"

    # Encabezado típico
    header_join = " ".join(cols_norm)
    if ("poliza" in header_join and "concepto" in header_join):
        return "star2"

    return "star1"


def _guess_header(df):
    """Encuentra la fila de encabezados tanto para STAR 1 como STAR 2.0."""
    header_idx = None
    for i in range(min(12, len(df))):
        row_vals = df.iloc[i].astype(str).str.replace("\xa0", " ", regex=False).str.strip().tolist()
        row_join = " ".join(row_vals)
        if re.search(r"cuenta.*concepto", row_join, re.IGNORECASE):
            header_idx = i
            break
        if re.search(r"poliza", row_join, re.IGNORECASE) and re.search(r"(concepto|fecha|cargos|abonos)", row_join, re.IGNORECASE):
            header_idx = i
            break

    if header_idx is not None:
        new_cols = df.iloc[header_idx].astype(str).str.replace("\xa0", " ", regex=False).str.strip().tolist()
        df2 = df.iloc[header_idx + 1:].reset_index(drop=True)
        df2.columns = new_cols[:df2.shape[1]]
        return df2, list(df2.columns)

    base_cols = ["Cuenta / Concepto", "Cheque", "Trafico", "Factura", "Fecha", "Cargos", "Abonos", "Saldo"]
    df.columns = base_cols[:df.shape[1]]
    return df, list(df.columns)


# =====================================================
# --- STAR 2.0 ---
# =====================================================

def process_star2_single(df_raw: pd.DataFrame) -> pd.DataFrame:
    df, _ = _guess_header(df_raw.copy())

    def _norm(c):
        s = str(c).strip().replace("\xa0", " ")
        mapping = {
            "poliza": "Poliza",
            "concepto": "Concepto",
            "cheque": "Cheque",
            "trafico": "Trafico",
            "tráfico": "Trafico",
            "factura": "Factura",
            "fecha": "Fecha",
            "cargos": "Cargos",
            "abonos": "Abonos",
            "saldo": "Saldo",
        }
        return mapping.get(s.lower(), s)

    df = df.rename(columns={c: _norm(c) for c in df.columns})

    # Cuenta en A2
    cuenta_text = ""
    if len(df) > 0:
        a2 = str(df.iloc[0, 0]).replace("\xa0", " ").strip()
        a2 = re.sub(r"^(cuenta\s*:|:)\s*", "", a2, flags=re.IGNORECASE).strip()
        cuenta_text = a2

    df_det = df.iloc[1:].reset_index(drop=True)
    df_det.insert(0, "Cuenta", cuenta_text)

    if "Concepto" in df_det.columns:
        df_det = df_det[df_det["Concepto"].astype(str).str.strip().ne("")]

    for col in ["Cargos", "Abonos", "Saldo"]:
        if col in df_det.columns:
            df_det[col] = df_det[col].apply(_to_num_safe)

    amt_cols = [c for c in ["Cargos", "Abonos", "Saldo"] if c in df_det.columns]
    if amt_cols:
        df_det = df_det[df_det[amt_cols].fillna(0).abs().sum(axis=1) > 0]

    order = ["Cuenta", "Poliza", "Concepto", "Cheque", "Trafico", "Factura", "Fecha", "Cargos", "Abonos", "Saldo"]
    return df_det[[c for c in order if c in df_det.columns]]


def process_star2_many(raws):
    all_data = []
    for df_raw in raws:
        df = process_star2_single(df_raw)
        all_data.append(df)
    return pd.concat(all_data, ignore_index=True)


# =====================================================
# --- STAR 1 ---
# =====================================================

def _normalize_date_series(s: pd.Series) -> pd.Series:
    s2 = s.copy()
    as_num = pd.to_numeric(s2, errors="coerce")
    mask_num = as_num.notna()
    if mask_num.any():
        s2.loc[mask_num] = pd.to_datetime(as_num[mask_num], unit="d", origin="1899-12-30").dt.strftime("%d/%m/%Y")
    mask_txt = ~mask_num
    if mask_txt.any():
        parsed = pd.to_datetime(s2[mask_txt].astype(str).str.strip(), errors="coerce", dayfirst=True)
        need_retry = parsed.isna()
        if need_retry.any():
            parsed2 = pd.to_datetime(s2[mask_txt][need_retry], errors="coerce", dayfirst=False)
            parsed.loc[need_retry] = parsed2
        s2.loc[mask_txt] = parsed.dt.strftime("%d/%m/%Y")
        s2 = s2.replace({"NaT": ""})
    return s2


def process_report(df_raw):
    df = df_raw.copy()
    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)
    df, _ = _guess_header(df)
    if "Cuenta" not in df.columns:
        df.insert(0, "Cuenta", "")

    def find_col(pat, default=None):
        for c in df.columns:
            if re.search(pat, str(c), re.IGNORECASE):
                return c
        return default

    col_cc = find_col(r"cuenta.*concepto", df.columns[1] if len(df.columns) > 1 else df.columns[0])
    col_cheque = find_col(r"cheq")
    col_traf = find_col(r"traf")
    col_fact = find_col(r"fact")
    col_cargos = find_col(r"cargos")
    col_abonos = find_col(r"abonos")
    col_saldo = find_col(r"saldo")

    cuenta_pat = re.compile(r"^\s*\d{3}-\d{2}-\d{2}-\d{3}-\d{2}-\d{3}-\d{4}\s+-\s+.+", re.IGNORECASE)

    last_cuenta = None
    rows_to_drop = []
    for idx, val in df[col_cc].astype(str).items():
        text = val.replace("\xa0", " ").strip()
        if cuenta_pat.match(text):
            last_cuenta = text
            rows_to_drop.append(idx)
            continue
        if text.lower() in {"saldo", "sumas totales"}:
            rows_to_drop.append(idx)
            continue
        df.at[idx, "Cuenta"] = last_cuenta or "__SIN_CUENTA__"

    df = df.drop(index=rows_to_drop).reset_index(drop=True)

    for col in [col_cargos, col_abonos, col_saldo]:
        if col:
            df[col] = df[col].apply(_to_num_safe)

    for c in df.columns:
        if re.search(r"fecha", str(c), re.IGNORECASE):
            df[c] = _normalize_date_series(df[c])

    amt_cols = [c for c in [col_cargos, col_abonos, col_saldo] if c]
    if amt_cols:
        df = df[df[amt_cols].fillna(0).abs().sum(axis=1) > 0]

    first_cols = ["Cuenta"]
    rest = [c for c in df.columns if c not in first_cols]
    return df[first_cols + rest]


# =====================================================
# --- INTERFAZ STREAMLIT ---
# =====================================================

st.subheader("Tipo de archivo")
mode = st.selectbox(
    "Selecciona el modo de procesamiento",
    ["Auto", "STAR 1 (todas las cuentas en un archivo)", "STAR 2.0 (por cuenta, múltiples archivos)"],
    index=0
)

uploaded_files = st.file_uploader(
    "Sube uno o varios archivos (.xls, .xlsx, .html, .htm)",
    type=["xls", "xlsx", "html", "htm"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Sube tus archivos para procesar.\n\n• **STAR 1**: un solo archivo con todas las cuentas.\n• **STAR 2.0**: varios archivos (uno por cuenta) y los consolidamos.")
else:
    try:
        raws = [_read_excel_any(up) for up in uploaded_files]

        if mode.startswith("Auto"):
            eff_mode = _detect_mode(raws[0])
        elif mode.startswith("STAR 1"):
            eff_mode = "star1"
        else:
            eff_mode = "star2"

        if eff_mode == "star1":
            if len(raws) > 1:
                st.warning("Modo STAR 1: se tomará solo el primer archivo.")
            df_clean = process_report(raws[0])
        else:
            df_clean = process_star2_many(raws)

        st.success(f"✅ Listo. Filas finales: {len(df_clean):,}")
        st.dataframe(df_clean.head(1000), use_container_width=True)

        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_clean.to_excel(writer, index=False, sheet_name="REPORTE")
        st.download_button(
            "⬇️ Descargar Excel procesado",
            data=buf.getvalue(),
            file_name="Reporte_procesado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Ocurrió un error procesando: {e}")
        st.exception(e)
