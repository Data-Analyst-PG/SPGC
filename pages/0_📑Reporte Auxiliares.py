import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
from lxml import etree
import re
from io import BytesIO
import io

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

    # 6) Fecha → date
    for c in df.columns:
        if re.search(r"fecha", str(c), re.IGNORECASE):
            try:
                df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
            except Exception:
                pass

    # 7) QUEDARME SOLO CON DETALLE QUE TENGA MONTOS REALES (>0 o <0)
    amt_cols = [c for c in [col_cargos, col_abonos, col_saldo] if c]
    if amt_cols:
        # Asegura que sean numéricos
        for c in amt_cols:
            df[c] = pd.to_numeric(df[c], errors="coerce")

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

uploaded = st.file_uploader(
    "Sube el archivo (.xls, .xlsx o .html)",
    type=["xls", "xlsx", "html", "htm"],
    accept_multiple_files=False
)

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
