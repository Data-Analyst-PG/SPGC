import io
import re
import unicodedata
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import pdfplumber
try:
    import streamlit as st
except ModuleNotFoundError:
    st = None

FINAL_COLUMNS = [
    "EMPRESA",
    "# FACTURA",
    "UUID",
    "FECHA FACTURA",
    "FECHA Y HR SERVICIO REALIZADO",
    "#REPORTE",
    "# DE UNIDAD",
    "ACTIVIDAD",
    "CANTIDAD",
    "SUBTOTAL",
    "IVA",
    "TOTAL",
]

PROVEEDORES = {
    "K9": ["MA. DEL CARMEN BALDERAS ESCAMILLA", "MA DEL CARMEN BALDERAS ESCAMILLA", "BAEM890616HW5"],
    "ROYAN": ["ALLAN ADRIAN NAVARRO MACIAS", "NAMA820330G3A"],
    "WASH N CROSS": ["WASH N CROSS", "WNC070608P43"],
}
SAT_GENERICO = "SAT GENERICO"


def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s or "") if unicodedata.category(c) != "Mn")


def find_text(pattern: str, text: str, flags: int = 0) -> str:
    m = re.search(pattern, text or "", flags)
    if not m:
        return ""
    return (m.group(1) if m.lastindex else m.group(0)).strip()


def norm_money(s: Any) -> float:
    s = str(s or "").replace("$", "").replace(",", "").strip()
    try:
        return round(float(s), 2)
    except ValueError:
        return 0.0


def D(value: Any, default: str = "0") -> Decimal:
    if value is None or value == "":
        value = default
    try:
        return Decimal(str(value).replace(",", "")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return Decimal(default).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def q2(value: Decimal) -> float:
    return float(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def norm(text: Optional[str]) -> str:
    return (text or "").upper().strip()


def clean_k9_service_dt(raw: str) -> str:
    if not raw:
        return ""
    raw = re.sub(r"\bHORA\b", "", raw, flags=re.I).strip()
    raw = re.sub(r"(\d{1,2})\.(\d{2})", r"\1:\2", raw)
    raw = re.sub(r"\s+", " ", raw).strip()
    return raw.replace("AM", "am").replace("PM", "pm")


def make_row(
    empresa: str = "",
    factura: str = "",
    uuid: str = "",
    fecha_factura: str = "",
    fecha_servicio: str = "",
    reporte: str = "",
    unidad: str = "",
    actividad: str = "",
    cantidad: Any = 1,
    subtotal: Any = 0,
    iva: Any = "",
    iva_rate: float = 0.08,
) -> Dict[str, Any]:
    subtotal_f = norm_money(subtotal)
    iva_f = round(subtotal_f * iva_rate, 2) if iva in (None, "") else norm_money(iva)
    return {
        "EMPRESA": empresa or "",
        "# FACTURA": factura or "",
        "UUID": uuid or "",
        "FECHA FACTURA": fecha_factura or "",
        "FECHA Y HR SERVICIO REALIZADO": fecha_servicio or "",
        "#REPORTE": reporte or "",
        "# DE UNIDAD": unidad or "",
        "ACTIVIDAD": actividad or "",
        "CANTIDAD": cantidad,
        "SUBTOTAL": subtotal_f,
        "IVA": iva_f,
        "TOTAL": round(subtotal_f + iva_f, 2),
    }


def rows_to_df(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    return pd.DataFrame(rows, columns=FINAL_COLUMNS)


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "FACTURAS") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 65)
    return output.getvalue()


# ===================== PDF =====================

def extract_pages_text(pdf_bytes: bytes) -> List[str]:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        return [(p.extract_text() or "") for p in pdf.pages]


def autodetect_pdf_format(full_text: str) -> str:
    t = strip_accents(full_text).upper()
    if "LOGA8509108NA" in t:
        return "ANA_CECILIA"
    if "WNC070608P43" in t or "WASH N CROSS" in t:
        return "WASH"
    if "NAMA820330G3A" in t or "ROYAN-" in t:
        return "ROYAN"
    if "BAEM890616HW5" in t or "COMENTARIOS:" in t or "ORDEN K9" in t:
        return "K9"
    return "K9"


def parse_pdf_k9(pdf_bytes: bytes) -> List[Dict[str, Any]]:
    pages = extract_pages_text(pdf_bytes)
    full = "\n".join(pages)
    empresa = find_text(r"NOMBRE COMERCIAL:\s*([^\n]+)", full)
    empresa = re.split(r"\s+CERTIFICADO\s+SAT", empresa, flags=re.I)[0].strip()
    uuid = find_text(r"\bUUID\s*\n\s*([0-9a-fA-F-]{36})", full, flags=re.I)
    fecha_factura = find_text(r"TEL\.?\s*\n\s*(\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}\s*[ap]\.m\.)", full, flags=re.I)
    if not fecha_factura:
        fecha_factura = find_text(r"(\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}\s*[ap]\.m\.)", full, flags=re.I)
    comentarios = find_text(r"Comentarios:\s*(.+)", full)
    factura = find_text(r"\bORDEN\s+(K9\s*\d+)\b", comentarios, flags=re.I).upper().replace("  ", " ")
    unidad = ""
    m = re.search(r"^\s*([A-ZÁÉÍÓÚÑ]+)\s+([A-Z0-9\-]+)\b", comentarios.strip(), flags=re.I)
    if m:
        unidad = m.group(2).strip()
    servicio = clean_k9_service_dt(find_text(r"\bSERVICIO REALIZADO\s+(.+)$", comentarios, flags=re.I))

    items = []
    pat_full = re.compile(r"^(?P<clave>\d{8})\s+(?P<desc>.+?)\s+(?P<unidad>[A-ZÁÉÍÓÚÑ]+)\s+(?P<cant>\d+)\s+(?P<precio>[\d,]+\.\d{2})\s+(?P<importe>[\d,]+\.\d{2})$", re.I)
    pat_close = re.compile(r"^(?P<desc2>.+?)\s+(?P<unidad>[A-ZÁÉÍÓÚÑ]+)\s+(?P<cant>\d+)\s+(?P<precio>[\d,]+\.\d{2})\s+(?P<importe>[\d,]+\.\d{2})$", re.I)
    pending_desc = ""
    for line in full.splitlines():
        s = re.sub(r"\s+", " ", line.strip())
        if not s:
            continue
        m1 = pat_full.match(s)
        if m1:
            items.append((m1.group("desc"), int(m1.group("cant")), m1.group("importe")))
            pending_desc = ""
            continue
        if re.match(r"^\d{8}\s+", s):
            pending_desc = re.sub(r"^\d{8}\s+", "", s).strip()
            continue
        if pending_desc:
            m2 = pat_close.match(s)
            if m2:
                items.append(((pending_desc + " " + m2.group("desc2")).strip(), int(m2.group("cant")), m2.group("importe")))
                pending_desc = ""
            else:
                pending_desc = (pending_desc + " " + s).strip()
    return [make_row(empresa, factura, uuid, fecha_factura, servicio, "", unidad, d, c, imp, "", 0.08) for d, c, imp in items]


def parse_pdf_royan(pdf_bytes: bytes) -> List[Dict[str, Any]]:
    pages = extract_pages_text(pdf_bytes)
    full = "\n".join(pages)
    empresa = find_text(r"\nCliente:\s*\n?([A-Z0-9ÁÉÍÓÚÑ ]+)\n", full).strip()
    factura = find_text(r"\b(ROYAN-\d+)\b", full)
    uuid = find_text(r"\b([0-9a-f]{8}-[0-9a-f\-]{27})\b", full, flags=re.I)
    fecha = find_text(r"\b(\d{2}/\d{2}/\d{4})\b", full)
    unidad = find_text(r"\bCaja:\s*([A-Z0-9\-]+)\b", full, flags=re.I)
    rows = []
    current_importe = None
    current_desc_parts: List[str] = []
    pat_start = re.compile(r"^(?P<importe>[\d,]+\.\d{2})\s+Actividad\s+(?P<desc>.+)$", re.I)
    for page_text in pages:
        for line in (page_text or "").splitlines():
            line = line.strip()
            if not line:
                continue
            m = pat_start.match(line)
            if m:
                if current_importe is not None and current_desc_parts:
                    rows.append(make_row(empresa, factura, uuid, fecha, "", "", unidad, " ".join(current_desc_parts), 1, current_importe, "", 0.16))
                current_importe = m.group("importe")
                current_desc_parts = [m.group("desc").strip()]
                continue
            if current_importe is not None:
                if line.upper() == "ACT" or re.match(r".*\bACT$", line, re.I):
                    cleaned = re.sub(r"\bACT\b", "", line, flags=re.I).strip()
                    if cleaned:
                        current_desc_parts.append(cleaned)
                    rows.append(make_row(empresa, factura, uuid, fecha, "", "", unidad, " ".join(current_desc_parts), 1, current_importe, "", 0.16))
                    current_importe = None
                    current_desc_parts = []
                else:
                    current_desc_parts.append(line)
    if current_importe is not None and current_desc_parts:
        rows.append(make_row(empresa, factura, uuid, fecha, "", "", unidad, " ".join(current_desc_parts), 1, current_importe, "", 0.16))
    return rows


def parse_pdf_wash(pdf_bytes: bytes) -> List[Dict[str, Any]]:
    pages = extract_pages_text(pdf_bytes)
    if not pages:
        return []
    p1 = strip_accents(pages[0])
    empresa = find_text(r"REGIMEN\s+FISCAL\s+\d+\s+([A-Z0-9 ,.&'\-]+)", p1, flags=re.I)
    if not empresa:
        # Alternativa: línea justo después del bloque Regimen Fiscal <num>
        empresa = find_text(r"Regimen\s+Fiscal\s+\d+\s*\n\s*([^\n]+)", p1, flags=re.I)
    factura = find_text(r"SERIE\s+Y\s+FOLIO\s+([A-Z0-9\-]+)", p1, flags=re.I)
    uuid = find_text(r"FOLIO\s+FISCAL\s*\(UUID\)\s*([0-9A-F\-]{36})", p1, flags=re.I)
    fecha_factura = find_text(r"FECHA\s+DE\s+EMISION\s+(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}(?::\d{2})?)", p1, flags=re.I)

    line_pat = re.compile(
        r"(?P<cant>\d+)\s*E48-?Unidad\s*de\s*servicio\s*(?P<prod>\d{8})\s*SERVICIOS\s*"
        r"(?P<desc>.+?)\s+(?P<traf>\d{6})\s+(?P<ref>[A-Z0-9\-]+(?:\s+[A-Z0-9\-]+)?)\s+"
        r"(?P<obs>\d{4}-\d{2}-\d{2})\s+(?P<precio>[\d,]+\.\d{2})\s+(?P<importe>[\d,]+\.\d{2})",
        flags=re.I,
    )
    rows: List[Dict[str, Any]] = []
    for raw in pages:
        t = strip_accents(raw or "")
        if "OBSERVACIONES" in t.upper() and not re.search(r"\b\d+\s*E48", t, flags=re.I):
            continue
        t = re.sub(r"servicio(\d{8})", r"servicio \1", t, flags=re.I)
        t = re.sub(r"\s+(?=\d+\s*E48-?Unidad\s*de\s*servicio)", "\n", t)
        for ln in [re.sub(r"\s+", " ", x).strip() for x in t.splitlines() if x.strip()]:
            m = line_pat.search(ln)
            if not m:
                continue
            rows.append(make_row(
                empresa=empresa.strip(),
                factura=factura.strip(),
                uuid=uuid.strip(),
                fecha_factura=fecha_factura.strip(),
                fecha_servicio=m.group("obs").strip(),
                reporte=m.group("traf").strip(),
                unidad=m.group("ref").strip(),
                actividad=re.sub(r"\s+", " ", m.group("desc")).strip(),
                cantidad=int(m.group("cant")),
                subtotal=m.group("importe"),
                iva="",
                iva_rate=0.08,
            ))
    return rows


def parse_pdf_ana_cecilia(pdf_bytes: bytes) -> List[Dict[str, Any]]:
    pages = extract_pages_text(pdf_bytes)
    t = strip_accents("\n".join(pages))
    t1 = re.sub(r"\s+", " ", t).strip()
    empresa = find_text(r"Nombre\s*receptor:\s*([^\n]+)", t, flags=re.I).strip()
    empresa = re.split(r"\s+emisi[oó]n:?", empresa, flags=re.I)[0].strip()
    if empresa.upper().replace(" ", "") == "LINCOLNFREIGHTCOMPANYLLC":
        empresa = "LINCOLN FREIGHT COMPANY LLC"
    factura = find_text(r"Folio:\s*(\d+)", t, flags=re.I)
    uuid = find_text(r"Folio\s*fiscal:\s*([0-9A-F-]{36})", t, flags=re.I)
    mdt = re.search(r"Codigo\s*postal,?fechayhorade.*?(\d{5})\s*(\d{4}-\d{2}-\d{2})\s*(\d{2}:\d{2}:\d{2})", t1, flags=re.I)
    fecha_factura = f"{mdt.group(2)} {mdt.group(3)}" if mdt else ""
    rows = []
    concept_pat = re.compile(
        r"(?P<clave>\d{8})\s+(?P<cant>\d+\.\d+)\s+E48\s+Unidaddeservicio\s+(?P<valor_unit>\d+)\s+"
        r"(?P<imp_concepto>\d+\.\d+)\s+Siobjetodeimpuesto\.\s+.*?Descripcion\s+(?P<desc>.+?)\s+"
        r"IVA\s+Traslado\s+(?P<base>\d+\.\d+)\s+Tasa\s+(?P<tasa>\d+\.\d+)%\s+(?P<iva>\d+\.\d+)\s+Numerodepedimento",
        flags=re.I | re.S,
    )
    for m in concept_pat.finditer(t1):
        desc = re.sub(r"\b(Factor|Cuota)\b", " ", m.group("desc"), flags=re.I)
        for w in ["PARA", "CAJA", "REVISION", "LUCES", "GENERAL", "REPARAR", "CORTO", "GALIBO", "LIMPIEZA", "SERVICIO", "DOMICILIO"]:
            desc = re.sub(w, " " + w + " ", desc, flags=re.I)
        desc = re.sub(r"\s+", " ", desc).strip()
        mu = re.search(r":\s*([A-Z0-9\-]+)\b", desc.upper())
        rows.append(make_row(empresa, factura, uuid, fecha_factura, "", "", mu.group(1) if mu else "", desc, int(float(m.group("cant"))), m.group("base"), m.group("iva"), 0.08))
    return rows


def parse_pdf_file(file_name: str, pdf_bytes: bytes, force_format: str = "Autodetectar") -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    try:
        pages = extract_pages_text(pdf_bytes)
        full = "\n".join(pages)
        fmt = autodetect_pdf_format(full) if force_format == "Autodetectar" else force_format
        if fmt == "K9":
            rows = parse_pdf_k9(pdf_bytes)
        elif fmt == "ROYAN":
            rows = parse_pdf_royan(pdf_bytes)
        elif fmt == "WASH":
            rows = parse_pdf_wash(pdf_bytes)
        elif fmt == "ANA_CECILIA":
            rows = parse_pdf_ana_cecilia(pdf_bytes)
        else:
            rows = []
        return rows, {"archivo": file_name, "tipo": fmt, "filas": len(rows), "estatus": "OK" if rows else "SIN FILAS"}
    except Exception as e:
        return [], {"archivo": file_name, "tipo": "ERROR", "filas": 0, "estatus": str(e)}


# ===================== XML =====================

def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def find_first_elem(root: ET.Element, name: str) -> Optional[ET.Element]:
    return next((elem for elem in root.iter() if local_name(elem.tag) == name), None)


def find_all_elem(root: ET.Element, name: str) -> List[ET.Element]:
    return [elem for elem in root.iter() if local_name(elem.tag) == name]


def attr(elem: Optional[ET.Element], key: str, default: str = "") -> str:
    return elem.attrib.get(key, default) if elem is not None else default


def get_uuid(root: ET.Element) -> str:
    return attr(find_first_elem(root, "TimbreFiscalDigital"), "UUID")


def get_emisor_receptor(root: ET.Element) -> Tuple[Dict[str, str], Dict[str, str]]:
    emisor = find_first_elem(root, "Emisor")
    receptor = find_first_elem(root, "Receptor")
    return (emisor.attrib if emisor is not None else {}, receptor.attrib if receptor is not None else {})


def get_concepts(root: ET.Element) -> List[ET.Element]:
    return find_all_elem(root, "Concepto")


def get_iva_from_concept(concepto: ET.Element) -> Decimal:
    iva = Decimal("0")
    for traslado in find_all_elem(concepto, "Traslado"):
        if traslado.attrib.get("Impuesto") == "002" or "IVA" in norm(traslado.attrib.get("Impuesto")):
            iva += D(traslado.attrib.get("Importe"))
    return iva.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def get_base_from_concept(concepto: ET.Element) -> Decimal:
    for traslado in find_all_elem(concepto, "Traslado"):
        if traslado.attrib.get("Base"):
            return D(traslado.attrib.get("Base"))
    return D(concepto.attrib.get("Importe"))


def serie_folio(root: ET.Element) -> str:
    serie = attr(root, "Serie")
    folio = attr(root, "Folio")
    return f"{serie}-{folio}" if serie and folio else (folio or serie)


def detect_xml_format(root: ET.Element) -> str:
    emisor, _ = get_emisor_receptor(root)
    emisor_text = norm(" ".join([emisor.get("Nombre", ""), emisor.get("Rfc", "")]))
    for formato, needles in PROVEEDORES.items():
        if any(norm(n) in emisor_text for n in needles):
            return formato
    return SAT_GENERICO if local_name(root.tag) == "Comprobante" and get_concepts(root) else "NO DETECTADO"


def common_xml_header(root: ET.Element) -> Dict[str, str]:
    _, receptor = get_emisor_receptor(root)
    return {"empresa": receptor.get("Nombre", ""), "factura": serie_folio(root), "uuid": get_uuid(root), "fecha": attr(root, "Fecha")}


def all_xml_text(root: ET.Element) -> str:
    parts = []
    for elem in root.iter():
        parts.extend([str(v) for v in elem.attrib.values()])
        if elem.text and elem.text.strip():
            parts.append(elem.text.strip())
    return " ".join(parts)


def parse_xml_k9(root: ET.Element) -> Tuple[List[Dict[str, Any]], str]:
    h = common_xml_header(root)
    text = all_xml_text(root)
    factura = find_text(r"ORDEN\s+(K9\s*\d+)", text, flags=re.I).upper() or h["factura"]
    unidad = ""
    mu = re.search(r"\b(CAJA|TRACTOR|CAMION)\s+([^\s]+)", text, flags=re.I)
    if mu:
        unidad = mu.group(2).strip()
    servicio = clean_k9_service_dt(find_text(r"SERVICIO\s+REALIZADO\s+(.+?)(?:\s+CAJA|\s+TRACTOR|\s+CAMION|$)", text, flags=re.I))
    rows = []
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = get_iva_from_concept(c) or subtotal * Decimal("0.08")
        rows.append(make_row(h["empresa"], factura, h["uuid"], h["fecha"], servicio, "", unidad, c.attrib.get("Descripcion", ""), q2(D(c.attrib.get("Cantidad"))), q2(subtotal), q2(iva), 0.08))
    msg = ""
    if not (factura and unidad and servicio):
        msg = "El XML K9 no trae todos los comentarios; si faltan unidad/servicio, normalmente vienen solo en el PDF."
    return rows, msg


def parse_xml_royan(root: ET.Element) -> Tuple[List[Dict[str, Any]], str]:
    h = common_xml_header(root)
    rows = []
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = get_iva_from_concept(c) or subtotal * Decimal("0.16")
        rows.append(make_row(h["empresa"], h["factura"], h["uuid"], h["fecha"], "", "", "", c.attrib.get("Descripcion", ""), q2(D(c.attrib.get("Cantidad"))), q2(subtotal), q2(iva), 0.16))
    return rows, "XML ROYAN procesado; el desglose detallado de actividades de la hoja 2 no viene dentro del XML."


def parse_xml_wash(root: ET.Element) -> Tuple[List[Dict[str, Any]], str]:
    h = common_xml_header(root)
    if re.fullmatch(r"A-\d{4}", h["factura"]):
        h["factura"] = "A-0" + h["factura"].split("-", 1)[1]
    rows = []
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = get_iva_from_concept(c) or subtotal * Decimal("0.08")
        rows.append(make_row(h["empresa"], h["factura"], h["uuid"], h["fecha"], "", "", "", c.attrib.get("Descripcion", ""), q2(D(c.attrib.get("Cantidad"))), q2(subtotal), q2(iva), 0.08))
    return rows, "XML WASH procesado. #REPORTE, # DE UNIDAD y fecha de servicio no vienen en el XML CFDI; esos datos salen de la representacion impresa PDF."


def parse_xml_sat_generico(root: ET.Element) -> Tuple[List[Dict[str, Any]], str]:
    h = common_xml_header(root)
    rows = []
    for c in get_concepts(root):
        desc = c.attrib.get("Descripcion", "")
        subtotal = get_base_from_concept(c)
        iva = get_iva_from_concept(c)
        cantidad = D(c.attrib.get("Cantidad"))
        cantidad_out = int(cantidad) if cantidad == cantidad.to_integral() else q2(cantidad)
        unidad = desc.rsplit(":", 1)[-1].split()[0].strip(".,;:") if ":" in desc else ""
        rows.append(make_row(h["empresa"], h["factura"], h["uuid"], h["fecha"], "", "", unidad, desc, cantidad_out, q2(subtotal), q2(iva), 0.08))
    return rows, ""


def parse_xml_file(file_name: str, xml_bytes: bytes) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    debug = {"archivo": file_name, "tipo": "", "filas": 0, "estatus": "OK", "mensaje": ""}
    try:
        root = ET.fromstring(xml_bytes)
        fmt = detect_xml_format(root)
        debug["tipo"] = fmt
        if fmt == "K9":
            rows, msg = parse_xml_k9(root)
        elif fmt == "ROYAN":
            rows, msg = parse_xml_royan(root)
        elif fmt == "WASH N CROSS":
            rows, msg = parse_xml_wash(root)
        elif fmt == SAT_GENERICO:
            rows, msg = parse_xml_sat_generico(root)
        else:
            rows, msg = [], "No se detecto como CFDI valido."
            debug["estatus"] = "ERROR"
        debug.update({"filas": len(rows), "mensaje": msg})
        if msg and debug["estatus"] == "OK":
            debug["estatus"] = "OK CON AVISO"
        return rows, debug
    except Exception as e:
        debug.update({"tipo": "ERROR", "estatus": "ERROR", "mensaje": str(e)})
        return [], debug


# ===================== STREAMLIT UI =====================

def render_result(rows: List[Dict[str, Any]], debug_rows: List[Dict[str, Any]], file_name: str):
    df = rows_to_df(rows)
    st.success(f"Listo: {len(df)} registros generados.")
    st.dataframe(df, use_container_width=True)
    with st.expander("Debug de procesamiento", expanded=False):
        st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)
    if not df.empty:
        st.download_button(
            "Descargar Excel consolidado",
            data=dataframe_to_excel_bytes(df),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def page_pdf():
    st.title("Lector de PDF")
    st.caption("Formatos soportados: K9, ROYAN, WASH N CROSS y ANA CECILIA.")
    files = st.file_uploader("Sube tus facturas PDF", type=["pdf"], accept_multiple_files=True, key="pdf_files")
    force = st.selectbox("Formato", ["Autodetectar", "K9", "ROYAN", "WASH", "ANA_CECILIA"], index=0)
    if st.button("Procesar PDF", type="primary") and files:
        all_rows: List[Dict[str, Any]] = []
        debug_rows: List[Dict[str, Any]] = []
        for f in files:
            data = f.getvalue()
            rows, dbg = parse_pdf_file(f.name, data, force)
            all_rows.extend(rows)
            debug_rows.append(dbg)
        render_result(all_rows, debug_rows, "FACTURAS_PDF_CONSOLIDADO.xlsx")


def page_xml():
    st.title("Lector de XML")
    st.caption("Lee CFDI XML. En WASH, #REPORTE, # DE UNIDAD y fecha de servicio se dejan vacios porque no vienen en el XML.")
    files = st.file_uploader("Sube tus XML", type=["xml"], accept_multiple_files=True, key="xml_files")
    if st.button("Procesar XML", type="primary") and files:
        all_rows: List[Dict[str, Any]] = []
        debug_rows: List[Dict[str, Any]] = []
        for f in files:
            rows, dbg = parse_xml_file(f.name, f.getvalue())
            all_rows.extend(rows)
            debug_rows.append(dbg)
        render_result(all_rows, debug_rows, "FACTURAS_XML_CONSOLIDADO.xlsx")


def main():
    if st is None:
        raise RuntimeError("Streamlit no esta instalado. Instala con: pip install streamlit pdfplumber pandas openpyxl")
    st.set_page_config(page_title="Lector PDF/XML", layout="wide")
    st.markdown("## Consolidador de facturas")
    try:
        pg = st.navigation([
            st.Page(page_pdf, title="Lector PDF", icon="📑"),
            st.Page(page_xml, title="Lector XML", icon="🧾"),
        ], position="top")
        pg.run()
    except Exception:
        tab_pdf, tab_xml = st.tabs(["📑 Lector PDF", "🧾 Lector XML"])
        with tab_pdf:
            page_pdf()
        with tab_xml:
            page_xml()


if __name__ == "__main__":
    main()
