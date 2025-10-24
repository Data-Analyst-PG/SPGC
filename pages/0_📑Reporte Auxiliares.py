import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
from lxml import etree
import re
from io import BytesIO
import io

# --- Detector simple del modo por DataFrame ---
def _detect_mode(df_raw: pd.DataFrame) -> str:
    """
    Detecta STAR 2.0 si, tras adivinar encabezados, existe una columna Poliza
    o si la primera columna luce como 'Poliza' y la primera fila de datos (A2)
    contiene algo tipo ': 200-...' (la Cuenta).
    En otro caso asume STAR 1.
    """
    try:
        df_guess, _ = _guess_header(df_raw.copy())
    except Exception:
        df_guess = df_raw

    cols_norm = [str(c).strip().lower() for c in df_guess.columns]

    # Señal directa por nombre de columna
    if "poliza" in cols_norm:
        return "star2"

    # Señal por estructura: primera columna tipo poliza y A2 con "Cuenta" iniciando por ": "
    if len(df_guess.columns) >= 1 and len(df_guess) >= 1:
        first_col_name = str(df_guess.columns[0]).strip().lower()
        a2 = str(df_guess.iloc[0, 0])  # primera fila de datos en la 0 tras guess_header
        if ("poliza" in first_col_name) or (a2.startswith(": ") or a2.startswith(":\u00a0")):
            return "star2"

    # fallback: STAR 1 si vemos "Cuenta / Concepto"
    if any(re.search(r"cuenta.*concepto", c, re.IGNORECASE) for c in cols_norm):
        return "star1"

    return "star1"

# --- Limpieza de montos a numérico seguro ---
def _to_num_safe(x):
    if pd.isna(x): 
        return pd.NA
    s = str(x)
    s = s.replace("\xa0", "").replace(" ", "")
    s = re.sub(r"[^\d,\.-]", "", s)  # quita $, etc.
    s = s.replace(",", "")           # separador miles
    return pd.to_numeric(s, errors="coerce")

# --- STAR 2.0: procesa 1 archivo por cuenta y devuelve DF detalle con columna Cuenta agregada ---
def process_star2_single(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    STAR 2.0:
    - Encabezados reales con _guess_header.
    - 'Poliza' es la PRIMERA columna.
    - La 'Cuenta' viene en A2 (primera fila de datos, primera columna) con prefijo ': '.
    - Quitar esa primera fila (la que trae la cuenta) y propagarla como nueva columna 'Cuenta'.
    - NO tocar formato de 'Fecha' (se deja tal cual).
    - Limpiar montos y filtrar filas sin montos reales.
    - Entregar columnas en orden exacto pedido.
    """
    # 1) Encabezados correctos
    df, _ = _guess_header(df_raw.copy())

    # 2) Normalizar nombres (por si hay mayúsculas/acentos/espacios)
    def _norm_name(c: str) -> str:
        s = str(c).strip()
        s_l = s.lower()
        # mapeos simples sin dependencias
        if s_l == "poliza":   return "Poliza"
        if s_l == "concepto": return "Concepto"
        if s_l == "cheque":   return "Cheque"
        if s_l == "trafico" or "tráfico" in s_l: return "Trafico"
        if s_l == "factura":  return "Factura"
        if s_l == "fecha":    return "Fecha"
        if s_l == "cargos":   return "Cargos"
        if s_l == "abonos":   return "Abonos"
        if s_l == "saldo":    return "Saldo"
        return s  # lo demás se conserva
    df = df.rename(columns={c: _norm_name(c) for c in df.columns})

    # 3) Tomar la Cuenta desde A2 (primera fila de datos, col 0 tras guess_header)
    cuenta_text = ""
    if len(df) > 0:
        cuenta_text = str(df.iloc[0, 0] if df.shape[1] > 0 else "")
    # quitar prefijo ": " o ": " (NBSP)
    cuenta_text = re.sub(r"^\s*:\s*", "", cuenta_text or "").strip()

    # 4) Quitar esa primera fila (la que contenía la cuenta) y resetear
    df_det = df.iloc[1:].reset_index(drop=True)

    # 5) Insertar columna 'Cuenta' al inicio y propagar
    if "Cuenta" not in df_det.columns:
        df_det.insert(0, "Cuenta", cuenta_text)
    else:
        df_det["Cuenta"] = cuenta_text

    # 6) Eliminar filas con Concepto vacío (si existe la col)
    if "Concepto" in df_det.columns:
        df_det = df_det[df_det["Concepto"].astype(str).str.strip().ne("")]

    # 7) Limpiar montos y filtrar por montos reales
    for col in ["Cargos", "Abonos", "Saldo"]:
        if col in df_det.columns:
            df_det[col] = df_det[col].apply(_to_num_safe)

    amt_cols = [c for c in ["Cargos", "Abonos", "Saldo"] if c in df_det.columns]
    if amt_cols:
        mask_nonzero = (df_det[amt_cols].fillna(0).abs().sum(axis=1) > 0)
        df_det = df_det.loc[mask_nonzero].reset_index(drop=True)

    # 8) NO tocar 'Fecha': se deja tal cual viene (texto o número según Excel).
    #    Si deseas vaciar NaN visualmente:
    if "Fecha" in df_det.columns:
        df_det["Fecha"] = df_det["Fecha"].astype(str).replace({"nan": ""})

    # 9) Orden EXACTO de columnas para STAR 2.0
    desired = ["Cuenta", "Poliza", "Concepto", "Cheque", "Trafico", "Factura", "Fecha", "Cargos", "Abonos", "Saldo"]
    ordered = [c for c in desired if c in df_det.columns]
    rest = [c for c in df_det.columns if c not in ordered]
    df_det = df_det[ordered + rest]

    return df_det

# --- STAR 2.0: consolida varios archivos ---
def process_star2_many(dfs_raw: list[pd.DataFrame]) -> pd.DataFrame:
    partes = [process_star2_single(df_raw) for df_raw in dfs_raw]
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

    # Orden de columnas sugerido
    ordered = [c for c in ["Cuenta", "Poliza", "Cuenta / Concepto", "Concepto" "Cheque", "Trafico", "Factura", "Fecha", "Cargos", "Abonos", "Saldo"] if c in out.columns]
    rest = [c for c in out.columns if c not in ordered]
    out = out[ordered + rest]
    return out

try:
    st.set_page_config(page_title="Reporte de Cuentas", layout="wide")
except Exception:
    pass

st.title("Reporte de Cuentas → Limpieza automática")
st.caption(
    "Sube el Excel (.xls o .xlsx) tal como lo descargas. "
    "La página eliminará encabezados/sumarios, propagará la cuenta y te dará un archivo limpio."
)

# Detecta por firma binaria: ZIP -> xlsx/xlsm; CFB -> xls
def _detect_excel_format(fileobj) -> str:
    """
    Devuelve 'xlsx' | 'xls' | 'html' | 'unknown' según contenido real.
    No consume el stream.
    """
    pos = fileobj.tell()
    head = fileobj.read(1024)  # leemos un poco más para detectar HTML
    fileobj.seek(pos)

    # ZIP → .xlsx /.xlsm
    if head.startswith(b'PK'):
        return "xlsx"

    # CFBF → .xls
    if head.startswith(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'):
        return "xls"

    # HTML → muchos exportadores lo guardan como .xls
    h = head.lstrip().lower()
    if h.startswith(b'<!doctype html') or h.startswith(b'<html') or h.startswith(b'<table'):
        return "html"

    return "unknown"

def _read_excel_any(uploaded_file):
    """
    Carga .xlsx/.xls reales, HTML "tipo Excel" (con o sin <table> estándar),
    y Excel 2003 XML (SpreadsheetML), incluso si viene incrustado en HTML.
    Devuelve un DataFrame (sin encabezados).
    """
    # Normaliza a BytesIO
    if hasattr(uploaded_file, "read"):
        raw = uploaded_file.read()
    else:
        raw = uploaded_file
    bio = io.BytesIO(raw)

    head = raw[:4096]
    head_stripped = head.lstrip().lower()

    # 1) XLSX (ZIP)
    if head.startswith(b"PK"):
        bio.seek(0)
        return pd.read_excel(bio, sheet_name=0, engine="openpyxl", header=None)

    # 2) XLS (CFBF/BIFF)
    if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
        bio.seek(0)
        try:
            return pd.read_excel(bio, sheet_name=0, engine="xlrd", header=None)
        except Exception:
            bio.seek(0)
            return pd.read_excel(bio, sheet_name=0, header=None)

    # 3) ¿HTML?
    is_html = (
        head_stripped.startswith(b"<!doctype html")
        or head_stripped.startswith(b"<html")
        or (b"<table" in head_stripped[:1024])
        or (b"xmlns:x=\"urn:schemas-microsoft-com:office:excel\"" in head)  # HTML "tipo Excel"
    )
    if is_html:
        # 3.a Intento rápido con read_html (lxml)
        bio.seek(0)
        try:
            tables = pd.read_html(bio, header=None, flavor="lxml")
            if tables:
                return tables[0]
        except Exception:
            pass

        # 3.b BeautifulSoup: buscar cualquier <table> (con o sin namespace)
        bio.seek(0)
        soup = BeautifulSoup(bio.read(), "lxml")

        # Buscar <table> sin/ con namespace (e.g., x:table, ss:table)
        def _is_table(tag):
            if not getattr(tag, "name", None):
                return False
            name = tag.name.lower()
            return name == "table" or name.endswith(":table")

        table = soup.find(_is_table)
        if table:
            # Convertir filas/ celdas (TR / TD-TH) con o sin namespace
            def _match(tag, names):
                if not getattr(tag, "name", None):
                    return False
                n = tag.name.lower()
                return (n in names) or any(n.endswith(":" + nm) for nm in names)

            rows = []
            for tr in table.find_all(lambda t: _match(t, {"tr"})):
                cells = [td.get_text(strip=True) for td in tr.find_all(lambda t: _match(t, {"td", "th"}))]
                if cells:
                    rows.append(cells)
            if rows:
                width = max(len(r) for r in rows)
                rows = [r + [""] * (width - len(r)) for r in rows]
                return pd.DataFrame(rows)

        # 3.c SpreadsheetML (Excel 2003 XML) incrustado dentro de HTML en <xml>…</xml>
        #    Buscar un bloque XML con el namespace de SpreadsheetML
        xml_block = soup.find("xml")
        if xml_block and ("urn:schemas-microsoft-com:office:spreadsheet" in xml_block.text):
            xml_bytes = xml_block.text.encode("utf-8", errors="ignore")
            try:
                tree = etree.fromstring(xml_bytes)
            except Exception:
                # Si falla, intenta parsear el documento completo como XML
                try:
                    tree = etree.XML(xml_bytes)
                except Exception:
                    tree = None
            if tree is not None:
                ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
                # Buscar la primera Worksheet/Table
                table = tree.find(".//ss:Worksheet/ss:Table", namespaces=ns)
                if table is not None:
                    rows = []
                    for row in table.findall("ss:Row", namespaces=ns):
                        row_vals = []
                        cur_col = 1
                        for cell in row.findall("ss:Cell", namespaces=ns):
                            idx = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
                            if idx is not None:
                                idx = int(idx)
                                while cur_col < idx:
                                    row_vals.append("")
                                    cur_col += 1
                            data_el = cell.find("ss:Data", namespaces=ns)
                            val = data_el.text if data_el is not None else ""
                            row_vals.append(val if val is not None else "")
                            cur_col += 1
                        rows.append(row_vals)
                    if rows:
                        width = max(len(r) for r in rows)
                        rows = [r + [""] * (width - len(r)) for r in rows]
                        return pd.DataFrame(rows)

        # Si llegamos aquí, es HTML pero sin tabla explotable
        raise ValueError("El archivo es HTML pero no contiene una tabla utilizable.")

    # 4) SpreadsheetML (XML plano, no HTML)
    if (b"<Workbook" in head) or (b"urn:schemas-microsoft-com:office:spreadsheet" in head):
        bio.seek(0)
        tree = etree.parse(bio)
        ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
        table = tree.find(".//ss:Worksheet/ss:Table", namespaces=ns)
        if table is None:
            raise ValueError("XML SpreadsheetML sin <Worksheet>/<Table>.")
        rows = []
        for row in table.findall("ss:Row", namespaces=ns):
            row_vals = []
            cur_col = 1
            for cell in row.findall("ss:Cell", namespaces=ns):
                idx = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
                if idx is not None:
                    idx = int(idx)
                    while cur_col < idx:
                        row_vals.append("")
                        cur_col += 1
                data_el = cell.find("ss:Data", namespaces=ns)
                val = data_el.text if data_el is not None else ""
                row_vals.append(val if val is not None else "")
                cur_col += 1
            rows.append(row_vals)
        if rows:
            width = max(len(r) for r in rows)
            rows = [r + [""] * (width - len(r)) for r in rows]
            return pd.DataFrame(rows)

    # 5) Desconocido: último intento con motores estándar
    bio.seek(0)
    try:
        return pd.read_excel(bio, sheet_name=0, engine="openpyxl", header=None)
    except Exception:
        bio.seek(0)
        return pd.read_excel(bio, sheet_name=0, engine="xlrd", header=None)

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

    # 1) Quitar banner
    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    # 2) Encabezados
    df, _ = _guess_header(df)

    # 3) Columna Cuenta
    if "Cuenta" not in df.columns:
        df.insert(0, "Cuenta", "")

    # --- localizar columnas clave por nombre ---
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

    # 4) Propagar Cuenta y marcar headers/sumarios puros para eliminar
    last_cuenta = None
    rows_to_drop = []

    for idx, val in df[col_cc].astype(str).items():
        text = val.strip()

        if cuenta_pat.match(text):
            last_cuenta = text
            rows_to_drop.append(idx)        # quitar encabezado
            continue

        # sumario si dice "Saldo" o "Sumas Totales"
        is_summary_word = text.lower() in {"saldo", "sumas totales"}

        # ¿tiene referencias de detalle?
        has_detail_refs = any(
            c and str(df.at[idx, c]).strip() not in {"", "nan", "None"}
            for c in [col_cheque, col_traf, col_fact]
        )

        if is_summary_word and not has_detail_refs:
            rows_to_drop.append(idx)        # sumario puro -> fuera
            continue

        # detalle → propagar cuenta
        df.at[idx, "Cuenta"] = last_cuenta if last_cuenta else "__SIN_CUENTA_DETECTADA__"

    df = df.drop(index=rows_to_drop).reset_index(drop=True)

    # 5) Normalizar montos (para no perder nada por formato)
    def to_num_safe(x):
        if pd.isna(x):
            return pd.NA
        s = str(x)
        s = s.replace("\xa0", "").replace(" ", "")
        s = re.sub(r"[^\d,\.-]", "", s)  # quita $ y otros
        s = s.replace(",", "")           # quita separador de miles
        return pd.to_numeric(s, errors="coerce")

    for col in [col_cargos, col_abonos, col_saldo]:
        if col:
            df[col] = df[col].apply(to_num_safe)

    # 6) Fecha → formato dd/mm/yyyy
    for c in df.columns:
        if re.search(r"fecha", str(c), re.IGNORECASE):
            try:
                # Convertir a datetime y luego formatear al texto dd/mm/yyyy
                df[c] = pd.to_datetime(df[c], errors="coerce").dt.strftime("%d/%m/%Y")
            except Exception:
                pass

    # 7) QUEDARME SOLO CON DETALLE QUE TENGA MONTOS REALES (>0 o <0)
    amt_cols = [c for c in [col_cargos, col_abonos, col_saldo] if c]
    if amt_cols:
        # Asegura que sean numéricos
        for c in amt_cols:
            df[c] = pd.to_numeric(df[c], errors="coerce")

        # 7.1) Eliminar filas con Cuenta/Concepto vacío (solo montos, sin detalle)
        if col_cc in df.columns:
            df = df[df[col_cc].astype(str).str.strip().ne("")]

        # Mantener solo filas con al menos un monto distinto de cero
        mask_nonzero = (df[amt_cols].fillna(0).abs().sum(axis=1) > 0)
        df = df.loc[mask_nonzero].reset_index(drop=True)

    # 8) Filas totalmente vacías (ignorando Cuenta) → fuera
    non_cuenta_cols = [c for c in df.columns if c != "Cuenta"]
    def _row_is_empty(series):
        for v in series.values:
            s = str(v).replace("\xa0", " ").strip().lower()
            if s not in {"", "nan", "none"}:
                return False
        return True
    df["__empty__"] = df[non_cuenta_cols].astype(str).apply(_row_is_empty, axis=1)
    df = df.loc[~df["__empty__"]].drop(columns="__empty__").reset_index(drop=True)

    return df

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
            df_clean = process_report(raws[0])      # tu función STAR 1 ya afinada
        else:
            df_clean = process_star2_many(raws)     # consolida y ordena columnas

        st.success(f"✅ Listo. Filas finales: {len(df_clean):,}")
        st.dataframe(df_clean.head(1000), use_container_width=True)

        # 4) descarga
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
