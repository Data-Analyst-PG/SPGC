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
    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)
    df, _ = _guess_header(df)
    df.insert(0, "Cuenta", "")
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
    non_aux_cols = [c for c in df_clean.columns if c != "Cuenta"]
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
