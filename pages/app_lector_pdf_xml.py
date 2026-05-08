import io
import re
import unicodedata
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
import streamlit as st

# PDF
try:
    import pdfplumber
except ModuleNotFoundError:
    pdfplumber = None


# ============================================================
# CONFIGURACION GENERAL
# ============================================================

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

PROVEEDORES_XML = {
    "K9": ["MA. DEL CARMEN BALDERAS ESCAMILLA", "MA DEL CARMEN BALDERAS ESCAMILLA", "BAEM890616HW5"],
    "ROYAN": ["ALLAN ADRIAN NAVARRO MACIAS", "NAMA820330G3A"],
    "WASH N CROSS": ["WASH N CROSS", "WNC070608P43"],
}

SAT_GENERICO = "SAT GENERICO"


# ============================================================
# UTILIDADES
# ============================================================

def strip_accents(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", s or "")
        if unicodedata.category(c) != "Mn"
    )


def norm(text: Optional[str]) -> str:
    return (text or "").upper().strip()


def find_first(pattern: str, text: str, flags=0) -> str:
    m = re.search(pattern, text or "", flags)
    if not m:
        return ""
    return (m.group(1) if m.lastindex else m.group(0)).strip()


def norm_money(s: Any) -> float:
    s = str(s or "").replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def D(value, default="0") -> Decimal:
    if value is None or value == "":
        value = default
    try:
        return Decimal(str(value).replace(",", "")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return Decimal(default).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def q2(value: Decimal) -> float:
    return float(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def make_row(
    empresa="",
    factura="",
    uuid="",
    fecha_factura="",
    fecha_servicio="",
    reporte="",
    unidad="",
    actividad="",
    cantidad=1,
    subtotal=0,
    iva=None,
    iva_rate=0.08,
) -> Dict[str, Any]:
    subtotal_f = norm_money(subtotal)
    if iva is None or iva == "":
        iva_f = round(subtotal_f * iva_rate, 2)
    else:
        iva_f = round(norm_money(iva), 2)
    return {
        "EMPRESA": empresa or "",
        "# FACTURA": factura or "",
        "UUID": uuid or "",
        "FECHA FACTURA": fecha_factura or "",
        "FECHA Y HR SERVICIO REALIZADO": fecha_servicio or "",
        "#REPORTE": reporte or "",
        "# DE UNIDAD": unidad or "",
        "ACTIVIDAD": actividad or "",
        "CANTIDAD": cantidad if cantidad not in [None, ""] else 1,
        "SUBTOTAL": round(subtotal_f, 2),
        "IVA": round(iva_f, 2),
        "TOTAL": round(subtotal_f + iva_f, 2),
    }


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="FACTURAS")
        ws = writer.book["FACTURAS"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)
    return output.getvalue()


def final_df(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    return pd.DataFrame(rows, columns=FINAL_COLUMNS)


# ============================================================
# PDF PARSERS
# ============================================================

def extract_pages_text(pdf_bytes: bytes) -> List[str]:
    if pdfplumber is None:
        raise RuntimeError("Falta instalar pdfplumber. Agrega pdfplumber a requirements.txt.")
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        return [(p.extract_text() or "") for p in pdf.pages]


def autodetect_pdf_format(full_text: str) -> str:
    t = strip_accents(full_text).upper()

    if "WNC070608P43" in t or "WASH N CROSS" in t:
        return "WASH"

    if "NAMA820330G3A" in t or "ROYAN-" in t:
        return "ROYAN"

    if "LOGA8509108NA" in t:
        return "ANA_CECILIA"

    if "BAEM890616HW5" in t or "COMENTARIOS:" in t or "ORDEN K9" in t:
        return "K9"

    return "K9"


def clean_k9_service_dt(raw: str) -> str:
    if not raw:
        return ""
    raw = re.sub(r"\bHORA\b", "", raw, flags=re.I).strip()
    raw = re.sub(r"(\d{1,2})\.(\d{2})", r"\1:\2", raw)
    raw = re.sub(r"\s+", " ", raw).strip()
    raw = raw.replace("AM", "am").replace("PM", "pm")
    return raw


def prettify_receiver_name(s: str) -> str:
    if not s:
        return ""
    u = s.upper().replace(" ", "")
    if u == "LINCOLNFREIGHTCOMPANYLLC":
        return "LINCOLN FREIGHT COMPANY LLC"
    return s


def parse_pdf_k9(pdf_bytes: bytes) -> Tuple[List[Dict[str, Any]], str]:
    pages = extract_pages_text(pdf_bytes)
    full = "\n".join(pages)

    empresa = find_first(r"NOMBRE COMERCIAL:\s*(.+)", full)
    factura_pdf = find_first(r"\bFACTURA:\s*\n?\s*([A-Z0-9\-]+)", full, flags=re.I)
    uuid = find_first(r"\bUUID\s*\n\s*([0-9a-fA-F-]{36})", full, flags=re.I)

    fecha_factura = find_first(
        r"Fecha\s*Expedici[oó]n:?\s*\n?\s*(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}\s*[ap]\.?m\.?)",
        full,
        flags=re.I,
    )
    if not fecha_factura:
        fecha_factura = find_first(
            r"(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}\s*[ap]\.?m\.?)",
            full,
            flags=re.I,
        )

    comentarios = find_first(r"Comentarios:\s*(.+)", full)
    orden_k9 = find_first(r"\bORDEN\s+(K9\s*\d+)\b", comentarios, flags=re.I).upper().replace("  ", " ")

    unidad = ""
    m = re.search(r"^\s*([A-ZÁÉÍÓÚÑ]+)\s+([A-Z0-9\-]+)\b", comentarios.strip(), flags=re.I)
    if m:
        unidad = m.group(2).strip()

    servicio_raw = find_first(r"\bSERVICIO REALIZADO\s+(.+)$", comentarios, flags=re.I)
    servicio = clean_k9_service_dt(servicio_raw)

    factura = orden_k9 or factura_pdf

    items: List[Dict[str, Any]] = []
    pat_full = re.compile(
        r"^(?P<clave>\d{8})\s+(?P<desc>.+?)\s+(?P<unidad>[A-ZÁÉÍÓÚÑ]+)\s+"
        r"(?P<cant>\d+)\s+(?P<precio>[\d,]+\.\d{2})\s+(?P<importe>[\d,]+\.\d{2})$",
        re.I,
    )
    pat_close = re.compile(
        r"^(?P<desc2>.+?)\s+(?P<unidad>[A-ZÁÉÍÓÚÑ]+)\s+(?P<cant>\d+)\s+"
        r"(?P<precio>[\d,]+\.\d{2})\s+(?P<importe>[\d,]+\.\d{2})$",
        re.I,
    )

    pending_desc = ""
    for line in full.splitlines():
        s = re.sub(r"\s+", " ", line.strip())
        if not s:
            continue

        m1 = pat_full.match(s)
        if m1:
            items.append(make_row(
                empresa=empresa, factura=factura, uuid=uuid, fecha_factura=fecha_factura,
                fecha_servicio=servicio, reporte=factura, unidad=unidad,
                actividad=m1.group("desc").strip(), cantidad=int(m1.group("cant")),
                subtotal=m1.group("importe"), iva_rate=0.08
            ))
            pending_desc = ""
            continue

        if re.match(r"^\d{8}\s+", s):
            pending_desc = re.sub(r"^\d{8}\s+", "", s).strip()
            continue

        if pending_desc:
            m2 = pat_close.match(s)
            if m2:
                full_desc = (pending_desc + " " + m2.group("desc2")).strip()
                items.append(make_row(
                    empresa=empresa, factura=factura, uuid=uuid, fecha_factura=fecha_factura,
                    fecha_servicio=servicio, reporte=factura, unidad=unidad,
                    actividad=full_desc, cantidad=int(m2.group("cant")),
                    subtotal=m2.group("importe"), iva_rate=0.08
                ))
                pending_desc = ""
            else:
                pending_desc = (pending_desc + " " + s).strip()

    return items, ""


def parse_pdf_royan(pdf_bytes: bytes) -> Tuple[List[Dict[str, Any]], str]:
    pages = extract_pages_text(pdf_bytes)
    full = "\n".join(pages)

    empresa = find_first(r"\nCliente:\s*\n?([A-Z0-9ÁÉÍÓÚÑ ]+)\n", full).strip()
    factura = find_first(r"\b(ROYAN-\d+)\b", full)
    uuid = find_first(r"\b([0-9a-f]{8}-[0-9a-f\-]{27})\b", full, flags=re.I)
    fecha = find_first(r"\b(\d{2}/\d{2}/\d{4})\b", full)
    unidad = find_first(r"\bCaja:\s*([A-Z0-9\-]+)\b", full, flags=re.I)

    items: List[Dict[str, Any]] = []
    pat_start = re.compile(r"^(?P<importe>[\d,]+\.\d{2})\s+Actividad\s+(?P<desc>.+)$", re.I)
    pat_end_act = re.compile(r".*\bACT\b$", re.I)

    current_importe = None
    current_desc_parts: List[str] = []

    for page_text in pages:
        for raw in (page_text or "").splitlines():
            line = raw.strip()
            if not line:
                continue

            m = pat_start.match(line)
            if m:
                if current_importe is not None and current_desc_parts:
                    desc = " ".join(current_desc_parts).strip()
                    items.append(make_row(
                        empresa=empresa, factura=factura, uuid=uuid, fecha_factura=fecha,
                        unidad=unidad, actividad=desc, cantidad=1, subtotal=current_importe, iva_rate=0.16
                    ))
                    current_desc_parts = []

                current_importe = m.group("importe")
                current_desc_parts = [m.group("desc").strip()]
                continue

            if current_importe is not None:
                if line.upper() == "ACT" or pat_end_act.match(line):
                    cleaned = re.sub(r"\bACT\b", "", line, flags=re.I).strip()
                    if cleaned:
                        current_desc_parts.append(cleaned)

                    desc = " ".join(current_desc_parts).strip()
                    items.append(make_row(
                        empresa=empresa, factura=factura, uuid=uuid, fecha_factura=fecha,
                        unidad=unidad, actividad=desc, cantidad=1, subtotal=current_importe, iva_rate=0.16
                    ))
                    current_importe = None
                    current_desc_parts = []
                else:
                    current_desc_parts.append(line)

    if current_importe is not None and current_desc_parts:
        desc = " ".join(current_desc_parts).strip()
        items.append(make_row(
            empresa=empresa, factura=factura, uuid=uuid, fecha_factura=fecha,
            unidad=unidad, actividad=desc, cantidad=1, subtotal=current_importe, iva_rate=0.16
        ))

    return items, ""


def _wash_fix_glued_text(s: str) -> str:
    s = s or ""
    s = re.sub(r"\b(\d+)\s*(E48)\b", r"\1 \2", s, flags=re.I)
    s = re.sub(r"(servicio)(\d{8})", r"\1 \2", s, flags=re.I)
    s = re.sub(r"(SERVICIOS)([A-ZÁÉÍÓÚÑ])", r"\1 \2", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_pdf_wash(pdf_bytes: bytes) -> Tuple[List[Dict[str, Any]], str]:
    pages = extract_pages_text(pdf_bytes)
    if not pages:
        return [], "PDF sin paginas."

    p1 = strip_accents(pages[0])
    full = "\n".join(strip_accents(p) for p in pages)

    # EMPRESA: receptor que aparece despues de Regimen Fiscal.
    empresa = find_first(r"Regimen\s+Fiscal\s+\d+\s+([A-Z0-9 ,.&'\-]+)", p1, flags=re.I)
    if not empresa:
        empresa = find_first(
            r"FECHA\s+DE\s+EMISION\s+\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}(?::\d{2})?\s+.*?\n([A-Z0-9 ,.&'\-]+)",
            p1,
            flags=re.I | re.S,
        )

    factura = find_first(r"SERIE\s+Y\s+FOLIO\s+([A-Z0-9\-]+)", p1, flags=re.I)
    uuid = find_first(r"FOLIO\s+FISCAL\s*\(UUID\)\s*([0-9A-F\-]{36})", p1, flags=re.I)
    fecha_factura = find_first(
        r"FECHA\s+DE\s+EMISION\s+(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}(?::\d{2})?)",
        p1,
        flags=re.I,
    )

    rows: List[Dict[str, Any]] = []

    line_pat = re.compile(
        r"(?P<cant>\d+)\s*"
        r"E48-?Unidad\s*de\s*servicio\s*"
        r"(?P<prod>\d{8})\s*"
        r"SERVICIOS\s*"
        r"(?P<desc>.+?)\s+"
        r"(?P<traf>\d{6})\s+"
        r"(?P<ref>[A-Z0-9\-]+(?:\s+[A-Z0-9\-]+)?)\s+"
        r"(?P<obs>\d{4}-\d{2}-\d{2})\s+"
        r"(?P<precio>[\d,]+\.\d{2})\s+"
        r"(?P<importe>[\d,]+\.\d{2})",
        flags=re.I,
    )

    for raw in pages:
        t = strip_accents(raw or "")
        t_upper = t.upper()

        # Paginas de observaciones no tienen partidas.
        if "OBSERVACIONES" in t_upper and not re.search(r"\b\d+\s*E48", t, flags=re.I):
            continue

        # En paginas posteriores el texto viene pegado: servicio78181500 / SERVICIOSCAMBIO.
        t = _wash_fix_glued_text(t)

        # Forzar corte antes de cada partida.
        t = re.sub(r"\s+(?=\d+\s*E48-?Unidad\s*de\s*servicio)", "\n", t, flags=re.I)

        for ln in [re.sub(r"\s+", " ", x).strip() for x in t.splitlines() if x.strip()]:
            m = line_pat.search(ln)
            if not m:
                continue

            rows.append(make_row(
                empresa=empresa.strip(),
                factura=factura.strip(),
                uuid=uuid.strip(),
                fecha_factura=fecha_factura.strip(),
                fecha_servicio=m.group("obs").strip(),       # OBS = fecha servicio
                reporte=m.group("traf").strip(),             # TRAFICO = # reporte
                unidad=m.group("ref").strip(),               # REF.PAGO = unidad
                actividad=re.sub(r"\s+", " ", m.group("desc")).strip(),
                cantidad=int(m.group("cant")),
                subtotal=m.group("importe"),
                iva_rate=0.08,
            ))

    msg = ""
    if not rows:
        msg = "No se detectaron partidas WASH. Revisa que el PDF tenga texto seleccionable."
    return rows, msg


def parse_pdf_ana_cecilia(pdf_bytes: bytes) -> Tuple[List[Dict[str, Any]], str]:
    pages = extract_pages_text(pdf_bytes)
    full = "\n".join(pages)

    t = strip_accents(full)
    t1 = re.sub(r"\s+", " ", t).strip()

    empresa_raw = find_first(r"Nombre\s*receptor:\s*([A-Z0-9 ]+)", t, flags=re.I)
    empresa = prettify_receiver_name(empresa_raw)
    factura = find_first(r"Folio:\s*(\d+)", t, flags=re.I)
    uuid = find_first(r"Folio\s*fiscal:\s*([0-9A-F-]{36})", t, flags=re.I)

    mdt = re.search(
        r"Codigo\s*postal,?fechayhorade.*?(\d{5})\s*(\d{4}-\d{2}-\d{2})\s*(\d{2}:\d{2}:\d{2})",
        t1,
        flags=re.I,
    )
    fecha_factura = f"{mdt.group(2)} {mdt.group(3)}" if mdt else ""

    rows: List[Dict[str, Any]] = []
    concept_pat = re.compile(
        r"(?P<clave>\d{8})\s+"
        r"(?P<cant>\d+\.\d+)\s+E48\s+Unidaddeservicio\s+"
        r"(?P<valor_unit>\d+)\s+(?P<imp_concepto>\d+\.\d+)\s+Siobjetodeimpuesto\.\s+"
        r".*?Descripcion\s+(?P<desc>.+?)\s+"
        r"IVA\s+Traslado\s+(?P<base>\d+\.\d+)\s+Tasa\s+(?P<tasa>\d+\.\d+)%\s+(?P<iva>\d+\.\d+)\s+"
        r"Numerodepedimento",
        flags=re.I | re.S,
    )

    for m in concept_pat.finditer(t1):
        desc = m.group("desc")
        desc = re.sub(r"\bFactor\b", " ", desc, flags=re.I)
        desc = re.sub(r"\bCuota\b", " ", desc, flags=re.I)
        desc = re.sub(r"\s+", " ", desc).strip()

        unidad = ""
        mu = re.search(r":\s*([A-Z]{2}\d+)\b", desc.upper())
        if mu:
            unidad = mu.group(1).strip()
        if not unidad:
            mu2 = re.search(r":\s*([A-Z0-9\-]+)\b", desc.upper())
            if mu2:
                unidad = mu2.group(1).strip()

        rows.append(make_row(
            empresa=empresa,
            factura=factura,
            uuid=uuid,
            fecha_factura=fecha_factura,
            unidad=unidad,
            actividad=desc,
            cantidad=int(float(m.group("cant"))),
            subtotal=m.group("base"),
            iva=m.group("iva"),
            iva_rate=0.08,
        ))

    return rows, ""


def parse_pdf_file(file_name: str, pdf_bytes: bytes, autodetect=True, forced_fmt="K9") -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    debug = {"archivo": file_name, "formato": "", "filas": 0, "estatus": "OK", "mensaje": ""}
    try:
        pages = extract_pages_text(pdf_bytes)
        full = "\n".join(pages)
        fmt = autodetect_pdf_format(full) if autodetect else forced_fmt
        debug["formato"] = fmt

        if fmt == "WASH":
            rows, msg = parse_pdf_wash(pdf_bytes)
        elif fmt == "ROYAN":
            rows, msg = parse_pdf_royan(pdf_bytes)
        elif fmt == "ANA_CECILIA":
            rows, msg = parse_pdf_ana_cecilia(pdf_bytes)
        else:
            rows, msg = parse_pdf_k9(pdf_bytes)

        debug["filas"] = len(rows)
        debug["mensaje"] = msg
        if msg:
            debug["estatus"] = "OK CON AVISO" if rows else "ERROR"
        return rows, debug
    except Exception as e:
        debug.update({"estatus": "ERROR", "mensaje": str(e)})
        return [], debug


# ============================================================
# XML PARSERS
# ============================================================

def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def find_xml_first(root: ET.Element, name: str) -> Optional[ET.Element]:
    for elem in root.iter():
        if local_name(elem.tag) == name:
            return elem
    return None


def find_xml_all(root: ET.Element, name: str) -> List[ET.Element]:
    return [elem for elem in root.iter() if local_name(elem.tag) == name]


def attr(elem: Optional[ET.Element], key: str, default: str = "") -> str:
    return elem.attrib.get(key, default) if elem is not None else default


def get_uuid(root: ET.Element) -> str:
    return attr(find_xml_first(root, "TimbreFiscalDigital"), "UUID")


def get_emisor_receptor(root: ET.Element) -> Tuple[Dict[str, str], Dict[str, str]]:
    emisor = find_xml_first(root, "Emisor")
    receptor = find_xml_first(root, "Receptor")
    return (emisor.attrib if emisor is not None else {}, receptor.attrib if receptor is not None else {})


def get_concepts(root: ET.Element) -> List[ET.Element]:
    return find_xml_all(root, "Concepto")


def get_iva_from_concept(concepto: ET.Element) -> Decimal:
    iva = Decimal("0")
    for traslado in find_xml_all(concepto, "Traslado"):
        if traslado.attrib.get("Impuesto") == "002" or "IVA" in norm(traslado.attrib.get("Impuesto")):
            iva += D(traslado.attrib.get("Importe"))
    return iva.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def get_base_from_concept(concepto: ET.Element) -> Decimal:
    for traslado in find_xml_all(concepto, "Traslado"):
        if traslado.attrib.get("Base"):
            return D(traslado.attrib.get("Base"))
    return D(concepto.attrib.get("Importe"))


def detect_xml_format(root: ET.Element) -> str:
    emisor, _ = get_emisor_receptor(root)
    emisor_text = norm(" ".join([emisor.get("Nombre", ""), emisor.get("Rfc", "")]))
    for formato, needles in PROVEEDORES_XML.items():
        if any(norm(n) in emisor_text for n in needles):
            return formato
    if local_name(root.tag) == "Comprobante" and get_concepts(root):
        return SAT_GENERICO
    return "NO DETECTADO"


def serie_folio(root: ET.Element) -> str:
    serie = attr(root, "Serie")
    folio = attr(root, "Folio")
    return f"{serie}-{folio}" if serie and folio else (folio or serie)


def common_xml_header(root: ET.Element) -> Dict[str, str]:
    emisor, receptor = get_emisor_receptor(root)
    return {
        "empresa": receptor.get("Nombre", ""),
        "folio": attr(root, "Folio"),
        "serie_folio": serie_folio(root),
        "uuid": get_uuid(root),
        "fecha": attr(root, "Fecha"),
        "emisor_nombre": emisor.get("Nombre", ""),
    }


def all_xml_text(root: ET.Element) -> str:
    parts = []
    for elem in root.iter():
        parts.extend([str(v) for v in elem.attrib.values()])
        if elem.text and elem.text.strip():
            parts.append(elem.text.strip())
    return " ".join(parts)


def complemento_text(root: ET.Element) -> str:
    texts = []
    for elem in root.iter():
        ln = local_name(elem.tag).upper()
        if ln in {"ADDENDA", "COMPLEMENTO", "OBSERVACIONES", "COMENTARIOS"}:
            texts.extend([str(v) for v in elem.attrib.values()])
            if elem.text:
                texts.append(elem.text)
        for k, v in elem.attrib.items():
            if norm(k) in {"OBSERVACIONES", "OBSERVACION", "COMENTARIOS", "COMENTARIO"}:
                texts.append(v)
    return " ".join(texts)


def extract_order_k9(text: str) -> str:
    m = re.search(r"ORDEN\s+K9\s*[-:]?\s*(\d+)", text, flags=re.I)
    return f"K9 {m.group(1)}" if m else ""


def extract_service_datetime_k9(text: str) -> str:
    m = re.search(r"SERVICIO\s+REALIZADO\s+(.+?)(?:\s+CAJA|\s+TRACTOR|\s+CAMION|$)", text, flags=re.I)
    return m.group(1).strip() if m else ""


def extract_unit_k9(text: str) -> str:
    m = re.search(r"\b(CAJA|TRACTOR|CAMION)\s+([^\s]+)", text, flags=re.I)
    return m.group(2).strip() if m else ""


def unit_from_description_after_colon(descripcion: str) -> str:
    if ":" not in descripcion:
        return ""
    value = descripcion.rsplit(":", 1)[-1].strip()
    return re.split(r"\s+", value)[0].strip(".,;:")


def concept_custom_value(concepto: ET.Element, keys: List[str]) -> str:
    wanted = [norm(k).replace(".", "").replace("_", "").replace(" ", "") for k in keys]
    for elem in [concepto] + list(concepto.iter()):
        for k, v in elem.attrib.items():
            nk = norm(k).replace(".", "").replace("_", "").replace(" ", "")
            if any(w in nk for w in wanted):
                return v
    return ""


def parse_xml_k9(root: ET.Element) -> Tuple[List[Dict[str, Any]], str]:
    h = common_xml_header(root)
    text = complemento_text(root) or all_xml_text(root)
    factura = extract_order_k9(text) or h["folio"] or h["serie_folio"]
    fecha_serv = extract_service_datetime_k9(text)
    unidad = extract_unit_k9(text)

    rows = []
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = get_iva_from_concept(c) or subtotal * Decimal("0.08")
        rows.append(make_row(
            empresa=h["empresa"], factura=factura, uuid=h["uuid"], fecha_factura=h["fecha"],
            fecha_servicio=fecha_serv, reporte=factura, unidad=unidad,
            actividad=c.attrib.get("Descripcion", ""), cantidad=q2(D(c.attrib.get("Cantidad"))),
            subtotal=q2(subtotal), iva=q2(iva), iva_rate=0.08
        ))

    msg = ""
    if not (extract_order_k9(text) and fecha_serv and unidad):
        msg = "XML procesado, pero no trae comentarios K9 completos; esos datos pueden venir solo en PDF."
    return rows, msg


def parse_xml_royan(root: ET.Element) -> Tuple[List[Dict[str, Any]], str]:
    h = common_xml_header(root)
    rows = []
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = get_iva_from_concept(c) or subtotal * Decimal("0.16")
        rows.append(make_row(
            empresa=h["empresa"], factura=h["folio"] or h["serie_folio"], uuid=h["uuid"],
            fecha_factura=h["fecha"], actividad=c.attrib.get("Descripcion", ""),
            cantidad=1, subtotal=q2(subtotal), iva=q2(iva), iva_rate=0.16
        ))
    return rows, "XML ROYAN procesado; el detalle de hoja 2 normalmente viene en PDF, no en XML."


def parse_xml_wash(root: ET.Element) -> Tuple[List[Dict[str, Any]], str]:
    """
    WASH XML: el CFDI normalmente NO trae TRAFICO / REF.PAGO / OBS.
    Por eso #REPORTE, # DE UNIDAD y FECHA SERVICIO se dejan vacios salvo que aparezcan en atributos/addenda.
    """
    h = common_xml_header(root)
    rows = []
    missing_ref_obs = False

    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = get_iva_from_concept(c) or subtotal * Decimal("0.08")

        reporte = concept_custom_value(c, ["TRAFICO", "REPORTE", "REPORTES"])
        ref_pago = concept_custom_value(c, ["REFPAGO", "REF PAGO", "REF.PAGO", "REFERENCIA PAGO"])
        obs = concept_custom_value(c, ["OBS", "OBSERVACION", "OBSERVACIONES", "FECHA SERVICIO"])

        if not reporte or not ref_pago or not obs:
            missing_ref_obs = True

        rows.append(make_row(
            empresa=h["empresa"], factura=h["serie_folio"], uuid=h["uuid"], fecha_factura=h["fecha"],
            fecha_servicio=obs, reporte=reporte, unidad=ref_pago,
            actividad=c.attrib.get("Descripcion", ""), cantidad=q2(D(c.attrib.get("Cantidad"))),
            subtotal=q2(subtotal), iva=q2(iva), iva_rate=0.08
        ))

    msg = ""
    if missing_ref_obs:
        msg = "WASH XML no trae TRAFICO/REF.PAGO/OBS dentro del CFDI; esos datos se obtienen del PDF."
    return rows, msg


def parse_xml_sat_generico(root: ET.Element) -> Tuple[List[Dict[str, Any]], str]:
    h = common_xml_header(root)
    rows = []
    for c in get_concepts(root):
        descripcion = c.attrib.get("Descripcion", "")
        subtotal = get_base_from_concept(c)
        iva = get_iva_from_concept(c)
        cantidad = D(c.attrib.get("Cantidad"))
        cantidad_out = int(cantidad) if cantidad == cantidad.to_integral() else float(cantidad)
        rows.append(make_row(
            empresa=h["empresa"], factura=h["folio"] or h["serie_folio"], uuid=h["uuid"],
            fecha_factura=h["fecha"], unidad=unit_from_description_after_colon(descripcion),
            actividad=descripcion, cantidad=cantidad_out, subtotal=q2(subtotal), iva=q2(iva), iva_rate=0.08
        ))
    return rows, ""


def parse_xml_file(file_name: str, xml_bytes: bytes) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    debug = {"archivo": file_name, "formato": "", "filas": 0, "estatus": "OK", "mensaje": ""}

    if not xml_bytes or len(xml_bytes.strip()) == 0:
        debug.update({"formato": "NO LEIDO", "estatus": "ERROR", "mensaje": "Archivo XML vacio."})
        return [], debug

    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        debug.update({"formato": "NO LEIDO", "estatus": "ERROR", "mensaje": f"XML mal formado: {e}"})
        return [], debug

    formato = detect_xml_format(root)
    debug["formato"] = formato

    try:
        if formato == "K9":
            rows, msg = parse_xml_k9(root)
        elif formato == "ROYAN":
            rows, msg = parse_xml_royan(root)
        elif formato == "WASH N CROSS":
            rows, msg = parse_xml_wash(root)
        elif formato == SAT_GENERICO:
            rows, msg = parse_xml_sat_generico(root)
        else:
            rows, msg = [], "No se detecto como CFDI valido."
            debug["estatus"] = "ERROR"

        debug["filas"] = len(rows)
        debug["mensaje"] = msg
        if msg and debug["estatus"] == "OK":
            debug["estatus"] = "OK CON AVISO"
        return rows, debug
    except Exception as e:
        debug.update({"estatus": "ERROR", "mensaje": str(e)})
        return [], debug


# ============================================================
# UI STREAMLIT - PAGINA CON TABS
# ============================================================

st.title("📑 Lector de Facturas PDF / XML")
st.caption("Consolida facturas en Excel sin modificar el menú principal de Streamlit. Esta página usa tabs internos.")

tab_pdf, tab_xml = st.tabs(["📄 Lector PDF", "🧾 Lector XML"])


with tab_pdf:
    st.subheader("📄 Lector de PDF")
    st.caption("Formatos soportados: K9, ROYAN, WASH N CROSS y Ana Cecilia.")

    files_pdf = st.file_uploader(
        "Sube tus facturas PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key="upload_pdf_facturas",
    )

    c1, c2 = st.columns(2)
    with c1:
        do_autodetect_pdf = st.checkbox("Autodetectar formato PDF", value=True, key="autodetect_pdf")
    with c2:
        forced_pdf_fmt = st.selectbox(
            "Formato manual",
            ["K9", "WASH", "ROYAN", "ANA_CECILIA"],
            index=0,
            disabled=do_autodetect_pdf,
            key="forced_pdf_fmt",
        )

    show_debug_pdf = st.checkbox("Mostrar debug PDF", value=True, key="show_debug_pdf")

    if st.button("Procesar PDF", key="btn_procesar_pdf"):
        if not files_pdf:
            st.warning("Sube uno o varios PDF.")
        else:
            all_rows: List[Dict[str, Any]] = []
            debug_rows: List[Dict[str, Any]] = []

            for f in files_pdf:
                rows, dbg = parse_pdf_file(
                    f.name,
                    f.getvalue(),
                    autodetect=do_autodetect_pdf,
                    forced_fmt=forced_pdf_fmt,
                )
                all_rows.extend(rows)
                debug_rows.append(dbg)

            df = final_df(all_rows)

            st.success(f"Listo: {len(df)} registros PDF generados.")
            st.dataframe(df, use_container_width=True)

            if show_debug_pdf:
                st.subheader("Debug PDF")
                st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)

            if not df.empty:
                st.download_button(
                    "⬇️ Descargar Excel PDF",
                    data=dataframe_to_excel_bytes(df),
                    file_name="FACTURAS_PDF_CONSOLIDADO.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_pdf_excel",
                )


with tab_xml:
    st.subheader("🧾 Lector de XML")
    st.caption("El XML se procesa con la estructura CFDI. En WASH, #REPORTE, unidad y fecha servicio suelen venir solo en PDF.")

    files_xml = st.file_uploader(
        "Sube tus facturas XML",
        type=["xml"],
        accept_multiple_files=True,
        key="upload_xml_facturas",
    )

    show_debug_xml = st.checkbox("Mostrar debug XML", value=True, key="show_debug_xml")

    if st.button("Procesar XML", key="btn_procesar_xml"):
        if not files_xml:
            st.warning("Sube uno o varios XML.")
        else:
            all_rows: List[Dict[str, Any]] = []
            debug_rows: List[Dict[str, Any]] = []

            for f in files_xml:
                rows, dbg = parse_xml_file(f.name, f.getvalue())
                all_rows.extend(rows)
                debug_rows.append(dbg)

            df = final_df(all_rows)

            st.success(f"Listo: {len(df)} registros XML generados.")
            st.dataframe(df, use_container_width=True)

            if show_debug_xml:
                st.subheader("Debug XML")
                st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)

            if not df.empty:
                st.download_button(
                    "⬇️ Descargar Excel XML",
                    data=dataframe_to_excel_bytes(df),
                    file_name="FACTURAS_XML_CONSOLIDADO.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_xml_excel",
                )
