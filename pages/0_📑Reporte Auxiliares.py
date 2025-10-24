import streamlit as st
import pandas as pd
import re
from io import BytesIO
import io

# =====================================================
# --- UTILIDADES BÁSICAS ---
# =====================================================

def _to_num_safe(x):
    """Convierte a float tolerando comas/$. NaN -> 0.0"""
    if pd.isna(x):
        return 0.0
    s = str(x).replace(",", "").replace("$", "").strip()
    try:
        return float(s)
    except Exception:
        return 0.0

def _drop_summary_rows(df: pd.DataFrame, cols: list[str] | None = None) -> pd.DataFrame:
    """
    Elimina renglones de sumatorias/globales donde cualquier columna de `cols`
    tenga valores como 'Sumas Totales', 'Suma Total', 'Total', 'Totales', 'Saldo', 'Saldos'.
    También limpia NBSP y espacios.
    """
    if cols is None:
        cols = list(df.columns)

    summary_re = re.compile(
        r"^\s*(sumas?\s+totales?|suma\s+total|totales?|saldo|saldos?)\s*$",
        re.IGNORECASE
    )

    mask = pd.Series(False, index=df.index)
    for c in cols:
        if c in df.columns:
            mask = mask | (
                df[c].astype(str)
                     .str.replace("\xa0", " ", regex=False)
                     .str.strip()
                     .str.match(summary_re)
            )
    return df.loc[~mask].reset_index(drop=True)

def _read_excel_any(uploaded):
    """
    Lee .xls, .xlsx, .html, .htm desde un UploadedFile de Streamlit.
    Usa engines modernos y mensajes claros si falta soporte para .xls.
    Devuelve DataFrame con dtype=str y sin encabezado (header=None).
    """
    # Conserva el buffer del archivo de Streamlit
    data = uploaded.read()
    bio = BytesIO(data)

    name = (uploaded.name or "").lower()

    # HTML/HTM
    if name.endswith((".html", ".htm")):
        bio.seek(0)
        tables = pd.read_html(bio, header=None, dtype=str)
        return tables[0]

    # XLSX (openpyxl)
    if name.endswith(".xlsx"):
        bio.seek(0)
        return pd.read_excel(bio, header=None, dtype=str, engine="openpyxl")

    # XLS (xlrd)
    if name.endswith(".xls"):
        try:
            bio.seek(0)
            return pd.read_excel(bio, header=None, dtype=str, engine="xlrd")
        except Exception as e:
            raise ValueError(
                "No pude leer .xls. Instala xlrd (soporta .xls) o convierte a .xlsx. "
                f"Detalle: {e}"
            )

    # Intento por contenido si la extensión viene rara
    # 1) probar openpyxl
    try:
        bio.seek(0)
        return pd.read_excel(bio, header=None, dtype=str, engine="openpyxl")
    except Exception:
        pass
    # 2) probar xlrd
    try:
        bio.seek(0)
        return pd.read_excel(bio, header=None, dtype=str, engine="xlrd")
    except Exception:
        pass
    # 3) probar como HTML
    try:
        bio.seek(0)
        tables = pd.read_html(bio, header=None, dtype=str)
        return tables[0]
    except Exception:
        pass

    raise ValueError("Formato no soportado. Usa .xlsx (recomendado), .xls (con xlrd), o .html/.htm.")


# =====================================================
# --- DETECCIÓN DEL MODO ---
# =====================================================

def _detect_mode(df_raw: pd.DataFrame) -> str:
    """Detecta STAR 1 o STAR 2.0 por encabezados o contenido A2."""
    try:
        df_guess, _ = _guess_header(df_raw.copy())
    except Exception:
        df_guess = df_raw.copy()

    cols_norm = [str(c).strip().lower().replace("\xa0", " ") for c in df_guess.columns]

    if any(c == "poliza" for c in cols_norm):
        return "star2"

    if len(df_guess.index) >= 1 and len(df_guess.columns) >= 1:
        a2 = str(df_guess.iloc[0, 0] if df_guess.shape[1] > 0 else "")
        if a2.replace("\xa0", " ").strip().startswith(":"):
            return "star2"

    header_join = " ".join(cols_norm)
    if ("poliza" in header_join and "concepto" in header_join) or ("poliza" in header_join and "fecha" in header_join):
        return "star2"

    return "star1"


def _guess_header(df):
    """Encuentra la fila de encabezados tanto para STAR 1 como STAR 2.0."""
    header_idx = None
    limit = min(12, len(df))

    for i in range(limit):
        row_vals = df.iloc[i].astype(str).str.replace("\xa0", " ", regex=False).str.strip().tolist()
        row_join = " ".join([v for v in row_vals if v and v.lower() != "nan"])

        if re.search(r"cuenta.*concepto", row_join, re.IGNORECASE) and re.search(
            r"(saldo|cargos|abonos)", row_join, re.IGNORECASE
        ):
            header_idx = i
            break

        if re.search(r"\bpoliza\b", row_join, re.IGNORECASE) and re.search(
            r"\b(concepto|fecha|saldo|cargos|abonos)\b", row_join, re.IGNORECASE
        ):
            header_idx = i
            break

    if header_idx is not None:
        new_cols = df.iloc[header_idx].astype(str).str.replace("\xa0", " ", regex=False).str.strip().tolist()
        new_cols = [c if c and c.lower() != "nan" else f"col{j}" for j, c in enumerate(new_cols)]
        df2 = df.iloc[header_idx + 1:].reset_index(drop=True)
        df2.columns = new_cols[: df2.shape[1]]
        return df2, list(df2.columns)

    base_cols = ["Cuenta / Concepto", "Cheque", "Trafico", "Factura", "Fecha", "Cargos", "Abonos", "Saldo"]
    cols = base_cols + [f"col{j}" for j in range(len(base_cols), df.shape[1])]
    df2 = df.copy()
    df2.columns = cols[: df.shape[1]]
    return df2, list(df2.columns)


# =====================================================
# --- STAR 2.0 ---
# =====================================================

def process_star2_single(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    STAR 2.0:
    - Detectar encabezados.
    - Tomar 'Cuenta' desde A2 (primera fila de datos, primera columna) removiendo ': ' / NBSP / 'Cuenta:' etc.
    - Eliminar esa fila y propagar 'Cuenta' como primera columna.
    - OMITIR filas de SUMAS TOTALES / SALDO(S).
    - Limpiar montos; ordenar columnas exactamente.
    """
    df, _ = _guess_header(df_raw.copy())

    def _norm_name(c: str) -> str:
        s = str(c).strip().replace("\xa0", " ")
        s_l = s.lower()
        if s_l == "poliza":   return "Poliza"
        if s_l == "concepto": return "Concepto"
        if s_l == "cheque":   return "Cheque"
        if s_l in ("trafico", "tráfico"): return "Trafico"
        if s_l == "factura":  return "Factura"
        if s_l == "fecha":    return "Fecha"
        if s_l == "cargos":   return "Cargos"
        if s_l == "abonos":   return "Abonos"
        if s_l == "saldo":    return "Saldo"
        return s
    df = df.rename(columns={c: _norm_name(c) for c in df.columns})

    # Cuenta en A2 (primera fila de datos, primera columna)
    cuenta_text = ""
    if len(df) > 0 and df.shape[1] > 0:
        a2 = str(df.iloc[0, 0])
        a2 = a2.replace("\xa0", " ").strip()
        a2 = re.sub(r"^(cuenta\s*:|:)\s*", "", a2, flags=re.IGNORECASE).strip()
        cuenta_text = a2

    # Remover esa fila y reset
    df_det = df.iloc[1:].reset_index(drop=True)

    # 🔴 Omitir filas de “SUMAS TOTALES / SALDO(S)”
    summary_re = re.compile(r"^\s*(sumas?\s+totales?|suma\s+total|saldo|saldos?)\s*$", re.IGNORECASE)
    mask_summary = pd.Series(False, index=df_det.index)
    for col in ["Concepto", "Poliza"]:
        if col in df_det.columns:
            mask_summary |= (
                df_det[col]
                .astype(str)
                .str.replace("\xa0", " ", regex=False)
                .str.strip()
                .str.match(summary_re)
            )
    df_det = df_det.loc[~mask_summary].reset_index(drop=True)

    # Insertar Cuenta primera
    if "Cuenta" not in df_det.columns:
        df_det.insert(0, "Cuenta", cuenta_text)
    else:
        df_det["Cuenta"] = cuenta_text

    # Filtrar conceptos vacíos si existiera la columna
    if "Concepto" in df_det.columns:
        df_det = df_det[df_det["Concepto"].astype(str).str.strip().ne("")]

    # Montos a numérico
    for col in ["Cargos", "Abonos", "Saldo"]:
        if col in df_det.columns:
            df_det[col] = df_det[col].apply(_to_num_safe)

    # Mantener filas con algún monto distinto de cero
    amt_cols = [c for c in ["Cargos", "Abonos", "Saldo"] if c in df_det.columns]
    if amt_cols:
        mask_nonzero = (df_det[amt_cols].fillna(0).abs().sum(axis=1) > 0)
        df_det = df_det.loc[mask_nonzero].reset_index(drop=True)

    # Orden exacto
    desired = ["Cuenta", "Poliza", "Concepto", "Cheque", "Trafico", "Factura", "Fecha", "Cargos", "Abonos", "Saldo"]
    ordered = [c for c in desired if c in df_det.columns]
    rest = [c for c in df_det.columns if c not in ordered]
    return df_det[ordered + rest]

def process_star2_many(raws):
    frames = [process_star2_single(df_raw) for df_raw in raws]
    if not frames:
        return pd.DataFrame(columns=["Cuenta", "Poliza", "Concepto", "Cheque", "Trafico", "Factura", "Fecha", "Cargos", "Abonos", "Saldo"])
    return pd.concat(frames, ignore_index=True)


# =====================================================
# --- STAR 1 ---
# =====================================================

def _normalize_date_series(s: pd.Series) -> pd.Series:
    """
    Convierte a dd/mm/yyyy respetando dayfirst e incluye seriales de Excel.
    """
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
    """
    STAR 1: limpia encabezados/sumarios, propaga Cuenta,
    normaliza montos y FECHA (dd/mm/yyyy).
    """
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

    col_cc     = find_col(r"cuenta.*concepto", df.columns[1] if len(df.columns) > 1 else df.columns[0])
    col_cheque = find_col(r"cheq")
    col_traf   = find_col(r"traf")
    col_fact   = find_col(r"fact")
    col_cargos = find_col(r"cargos")
    col_abonos = find_col(r"abonos")
    col_saldo  = find_col(r"saldo")

    cuenta_pat = re.compile(r"^\s*\d{3}-\d{2}-\d{2}-\d{3}-\d{2}-\d{3}-\d{4}\s+-\s+.+", re.IGNORECASE)

    last_cuenta = None
    rows_to_drop = []

    for idx, val in df[col_cc].astype(str).items():
        text = val.replace("\xa0", " ").strip()

        if cuenta_pat.match(text):
            last_cuenta = text
            rows_to_drop.append(idx)
            continue

        is_summary_word = text.lower() in {"saldo", "sumas totales"}
        has_detail_refs = any(
            c and str(df.at[idx, c]).strip() not in {"", "nan", "None"}
            for c in [col_cheque, col_traf, col_fact]
        )

        if is_summary_word and not has_detail_refs:
            rows_to_drop.append(idx)
            continue

        df.at[idx, "Cuenta"] = last_cuenta if last_cuenta else "__SIN_CUENTA_DETECTADA__"

    df = df.drop(index=rows_to_drop).reset_index(drop=True)

    for col in [col_cargos, col_abonos, col_saldo]:
        if col:
            df[col] = df[col].apply(_to_num_safe)

    for c in df.columns:
        if re.search(r"fecha", str(c), re.IGNORECASE):
            df[c] = _normalize_date_series(df[c])

    amt_cols = [c for c in [col_cargos, col_abonos, col_saldo] if c]
    if amt_cols:
        for c in amt_cols:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        if col_cc in df.columns:
            df = df[df[col_cc].astype(str).str.strip().ne("")]
        df = df[df[amt_cols].fillna(0).abs().sum(axis=1) > 0].reset_index(drop=True)

    non_cuenta_cols = [c for c in df.columns if c != "Cuenta"]
    def _row_is_empty(series):
        for v in series.values:
            s = str(v).replace("\xa0", " ").strip().lower()
            if s not in {"", "nan", "none"}:
                return False
        return True

    df["__empty__"] = df[non_cuenta_cols].astype(str).apply(_row_is_empty, axis=1)
    df = df.loc[~df["__empty__"]].drop(columns="__empty__").reset_index(drop=True)

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
        # Leer todos a DF crudos
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
            df_clean = process_report(raws[0])      # STAR 1
        else:
            df_clean = process_star2_many(raws)     # STAR 2.0 (consolidado)

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
