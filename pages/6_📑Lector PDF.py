# app.py
import re
import io
import pandas as pd
import pdfplumber
import streamlit as st

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Lector PDF → Excel", layout="wide")

OUTPUT_COLS = [
    "EMPRESA", "#FACTURA", "UUID", "FECHA FACTURA",
    "FECHA Y HR SERVICIO", "#UNIDAD",
    "ACTIVIDAD", "CANTIDAD", "SUBTOTAL", "IVA", "TOTAL"
]

# =========================
# HELPERS
# =========================
def norm_num(s):
    if not s:
        return 0.0
    return float(str(s).replace(",", "").replace("$", "").strip())

def extract_pages_text(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        return [(p.extract_text() or "") for p in pdf.pages]

def find_first(pattern, text, flags=0):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""

def clean_k9_service_dt(raw):
    if not raw:
        return ""
    raw = re.sub(r"HORA", "", raw, flags=re.I)
    raw = re.sub(r"(\d{1,2})\.(\d{2})", r"\1:\2", raw)
    raw = re.sub(r"\s+", " ", raw).strip()
    raw = raw.replace("AM", "am").replace("PM", "pm")
    return raw

def build_df(rows, iva_rate):
    out = []
    for r in rows:
        subtotal = norm_num(r.get("SUBTOTAL"))
        iva = round(subtotal * iva_rate, 2)
        total = round(subtotal + iva, 2)

        out.append({
            "EMPRESA": r.get("EMPRESA", ""),
            "#FACTURA": r.get("#FACTURA", ""),
            "UUID": r.get("UUID", ""),
            "FECHA FACTURA": r.get("FECHA FACTURA", ""),
            "FECHA Y HR SERVICIO": r.get("FECHA Y HR SERVICIO", ""),
            "#UNIDAD": r.get("#UNIDAD", ""),
            "ACTIVIDAD": r.get("ACTIVIDAD", ""),
            "CANTIDAD": r.get("CANTIDAD", 1),
            "SUBTOTAL": round(subtotal, 2),
            "IVA": iva,
            "TOTAL": total
        })

    return pd.DataFrame(out, columns=OUTPUT_COLS)

# =========================
# PARSER K9
# =========================
def parse_k9(pdf_bytes):
    pages = extract_pages_text(pdf_bytes)
    full = "\n".join(pages)

    empresa = find_first(r"NOMBRE COMERCIAL:\s*(.+)", full)
    uuid = find_first(r"UUID\s*\n\s*([0-9a-fA-F-]{36})", full)
    fecha_factura = find_first(
        r"(\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}\s*[ap]\.m\.)",
        full, re.I
    )

    comentarios = find_first(r"Comentarios:\s*(.+)", full)

    factura = find_first(r"ORDEN\s+(K9\s*\d+)", comentarios, re.I)

    # unidad = lo que va después de la primera palabra
    unidad = ""
    m = re.search(r"^\s*\w+\s+([A-Z0-9\-]+)", comentarios)
    if m:
        unidad = m.group(1)

    servicio_raw = find_first(r"SERVICIO REALIZADO\s+(.+)$", comentarios, re.I)
    servicio = clean_k9_service_dt(servicio_raw)

    header = {
        "EMPRESA": empresa,
        "#FACTURA": factura,
        "UUID": uuid,
        "FECHA FACTURA": fecha_factura,
        "FECHA Y HR SERVICIO": servicio,
        "#UNIDAD": unidad
    }

    items = []
    pending_desc = ""
    pending = False

    for line in full.splitlines():
        s = re.sub(r"\s+", " ", line.strip())
        if not s:
            continue

        m = re.match(
            r"^\d{8}\s+(.+?)\s+SERVICIO\S*\s+(\d+)\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})$",
            s, re.I
        )
        if m:
            items.append({
                "ACTIVIDAD": m.group(1).strip(),
                "CANTIDAD": int(m.group(2)),
                "SUBTOTAL": norm_num(m.group(3))
            })
            pending = False
            continue

        if pending:
            pending_desc += " " + s
        elif re.match(r"^\d{8}\s+", s):
            pending_desc = s
            pending = True

    return header, items

# =========================
# PARSER ROYAN
# =========================
def parse_royan(pdf_bytes):
    pages = extract_pages_text(pdf_bytes)
    full = "\n".join(pages)

    header = {
        "EMPRESA": find_first(r"Cliente:\s*\n?([A-Z0-9 ]+)", full),
        "#FACTURA": find_first(r"(ROYAN-\d+)", full),
        "UUID": find_first(r"([0-9a-f]{8}-[0-9a-f\-]{27})", full, re.I),
        "FECHA FACTURA": find_first(r"(\d{2}/\d{2}/\d{4})", full),
        "FECHA Y HR SERVICIO": "",
        "#UNIDAD": find_first(r"Caja:\s*([A-Z0-9\-]+)", full)
    }

    items = []
    for page in pages:
        for line in page.splitlines():
            m = re.match(
                r"([\d,]+\.\d{2})\s+Actividad\s+(.+?)\s+ACT",
                line.strip(), re.I
            )
            if m:
                items.append({
                    "ACTIVIDAD": m.group(2).strip(),
                    "CANTIDAD": 1,
                    "SUBTOTAL": norm_num(m.group(1))
                })

    return header, items

# =========================
# PARSER WASH N CROSS
# =========================
def parse_wash(pdf_bytes):
    pages = extract_pages_text(pdf_bytes)
    full = "\n".join(pages)

    header = {
        "EMPRESA": find_first(r"\n(PICUS)\n", full),
        "#FACTURA": find_first(r"SERIE Y FOLIO\s+([A-Z0-9\-]+)", full),
        "UUID": find_first(r"UUID\)\s*\n([0-9A-F-]{36})", full, re.I),
        "FECHA FACTURA": find_first(
            r"FECHA DE EMISION\s+(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})",
            full
        )
    }

    items = []
    for line in full.splitlines():
        s = re.sub(r"\s+", " ", line.strip())
        m = re.match(
            r"^(\d+).+?\d{8}\s+(.+?)\s+\d+\s+([A-Z0-9\- ]+)\s+(\d{4}-\d{2}-\d{2}).+?([\d,]+\.\d{2})$",
            s
        )
        if m:
            items.append({
                "EMPRESA": header["EMPRESA"],
                "#FACTURA": header["#FACTURA"],
                "UUID": header["UUID"],
                "FECHA FACTURA": header["FECHA FACTURA"],
                "FECHA Y HR SERVICIO": m.group(4),
                "#UNIDAD": m.group(3),
                "ACTIVIDAD": m.group(2),
                "CANTIDAD": int(m.group(1)),
                "SUBTOTAL": norm_num(m.group(5))
            })

    return header, items

# =========================
# STREAMLIT UI
# =========================
st.title("📄 Lector de Facturas PDF → Excel")

formato = st.selectbox(
    "Selecciona el formato",
    ["K9", "ROYAN", "WASH N CROSS"]
)

pdf = st.file_uploader("Sube la factura PDF", type="pdf")

if st.button("Procesar") and pdf:
    pdf_bytes = pdf.read()

    if formato == "K9":
        header, items = parse_k9(pdf_bytes)
        rows = [{**header, **i} for i in items]
        df = build_df(rows, 0.08)

    elif formato == "ROYAN":
        header, items = parse_royan(pdf_bytes)
        rows = [{**header, **i} for i in items]
        df = build_df(rows, 0.16)

    else:
        header, items = parse_wash(pdf_bytes)
        df = build_df(items, 0.08)

    st.success(f"Listo: {len(df)} registros")
    st.dataframe(df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="FACTURAS")

    st.download_button(
        "⬇️ Descargar Excel",
        data=output.getvalue(),
        file_name=f"{formato}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
