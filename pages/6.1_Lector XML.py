import io
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET

import pandas as pd
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


def D(value, default="0") -> Decimal:
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


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def find_first(root: ET.Element, name: str) -> Optional[ET.Element]:
    for elem in root.iter():
        if local_name(elem.tag) == name:
            return elem
    return None


def find_all(root: ET.Element, name: str) -> List[ET.Element]:
    return [elem for elem in root.iter() if local_name(elem.tag) == name]


def attr(elem: Optional[ET.Element], key: str, default: str = "") -> str:
    return elem.attrib.get(key, default) if elem is not None else default


def get_uuid(root: ET.Element) -> str:
    timbre = find_first(root, "TimbreFiscalDigital")
    return attr(timbre, "UUID")


def get_emisor_receptor(root: ET.Element) -> Tuple[Dict[str, str], Dict[str, str]]:
    emisor = find_first(root, "Emisor")
    receptor = find_first(root, "Receptor")
    return (emisor.attrib if emisor is not None else {}, receptor.attrib if receptor is not None else {})


def get_concepts(root: ET.Element) -> List[ET.Element]:
    return find_all(root, "Concepto")


def get_iva_from_concept(concepto: ET.Element) -> Decimal:
    iva = Decimal("0")
    for traslado in find_all(concepto, "Traslado"):
        if traslado.attrib.get("Impuesto") == "002" or "IVA" in norm(traslado.attrib.get("Impuesto")):
            iva += D(traslado.attrib.get("Importe"))
    return iva.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def get_base_from_concept(concepto: ET.Element) -> Decimal:
    traslados = find_all(concepto, "Traslado")
    for traslado in traslados:
        if traslado.attrib.get("Base"):
            return D(traslado.attrib.get("Base"))
    return D(concepto.attrib.get("Importe"))


def detect_format(root: ET.Element) -> str:
    emisor, _ = get_emisor_receptor(root)
    emisor_text = norm(" ".join([emisor.get("Nombre", ""), emisor.get("Rfc", "")]))
    for formato, needles in PROVEEDORES.items():
        if any(norm(n) in emisor_text for n in needles):
            return formato
    # Todo XML CFDI valido que no sea proveedor propio entra como SAT generico.
    if local_name(root.tag) == "Comprobante" and get_concepts(root):
        return SAT_GENERICO
    return "NO DETECTADO"


def serie_folio(root: ET.Element) -> str:
    serie = attr(root, "Serie")
    folio = attr(root, "Folio")
    return f"{serie}-{folio}" if serie and folio else (folio or serie)


def complemento_text(root: ET.Element) -> str:
    # Algunos proveedores guardan comentarios/observaciones en Addenda o Complemento.
    texts = []
    for elem in root.iter():
        ln = local_name(elem.tag).upper()
        if ln in {"ADDENDA", "COMPLEMENTO", "OBSERVACIONES", "COMENTARIOS"}:
            texts.append(" ".join([str(v) for v in elem.attrib.values()]))
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
    # SAT generico: toma lo que venga despues del ultimo ':'; ej. "CAJA: PI59" -> PI59.
    if ":" not in descripcion:
        return ""
    value = descripcion.rsplit(":", 1)[-1].strip()
    return re.split(r"\s+", value)[0].strip(".,;:")


def common_header(root: ET.Element) -> Dict[str, str]:
    emisor, receptor = get_emisor_receptor(root)
    return {
        "empresa": receptor.get("Nombre", ""),
        "folio": attr(root, "Folio"),
        "serie_folio": serie_folio(root),
        "uuid": get_uuid(root),
        "fecha": attr(root, "Fecha"),
        "emisor_nombre": emisor.get("Nombre", ""),
    }


def make_row(empresa, factura, uuid, fecha, fecha_servicio, unidad, actividad, cantidad, subtotal, iva) -> Dict[str, object]:
    subtotal = D(subtotal)
    iva = D(iva)
    return {
        "EMPRESA": empresa,
        "# FACTURA": factura,
        "UUID": uuid,
        "FECHA FACTURA": fecha,
        "FECHA Y HR SERVICIO REALIZADO": fecha_servicio,
        "# DE UNIDAD": unidad,
        "ACTIVIDAD": actividad,
        "CANTIDAD": cantidad,
        "SUBTOTAL": q2(subtotal),
        "IVA": q2(iva),
        "TOTAL": q2(subtotal + iva),
    }


def parse_k9(root: ET.Element) -> List[Dict[str, object]]:
    h = common_header(root)
    text = complemento_text(root)
    factura = extract_order_k9(text) or h["folio"] or h["serie_folio"]
    fecha_serv = extract_service_datetime_k9(text)
    unidad = extract_unit_k9(text)
    rows = []
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = subtotal * Decimal("0.08")
        rows.append(make_row(h["empresa"], factura, h["uuid"], h["fecha"], fecha_serv, unidad,
                             c.attrib.get("Descripcion", ""), D(c.attrib.get("Cantidad")), subtotal, iva))
    return rows


def parse_royan(root: ET.Element) -> List[Dict[str, object]]:
    h = common_header(root)
    rows = []
    # En XML puro normalmente solo viene el concepto fiscal resumido. Si el XML trae Addenda con partidas, se procesan.
    # Si no hay Addenda, se genera la(s) partida(s) disponibles en Concepto.
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = subtotal * Decimal("0.16")
        rows.append(make_row(h["empresa"], h["folio"], h["uuid"], h["fecha"], "", "",
                             c.attrib.get("Descripcion", ""), Decimal("1"), subtotal, iva))
    return rows


def parse_wash(root: ET.Element) -> List[Dict[str, object]]:
    h = common_header(root)
    rows = []
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = subtotal * Decimal("0.08")
        rows.append(make_row(h["empresa"], h["serie_folio"], h["uuid"], h["fecha"], "", "",
                             c.attrib.get("Descripcion", ""), D(c.attrib.get("Cantidad")), subtotal, iva))
    return rows


def parse_sat_generico(root: ET.Element) -> List[Dict[str, object]]:
    h = common_header(root)
    rows = []
    for c in get_concepts(root):
        descripcion = c.attrib.get("Descripcion", "")
        subtotal = get_base_from_concept(c)
        iva = get_iva_from_concept(c)
        cantidad = D(c.attrib.get("Cantidad"))
        if cantidad == cantidad.to_integral():
            cantidad_out = int(cantidad)
        else:
            cantidad_out = float(cantidad)
        rows.append(make_row(h["empresa"], h["folio"] or h["serie_folio"], h["uuid"], h["fecha"], "",
                             unit_from_description_after_colon(descripcion), descripcion, cantidad_out, subtotal, iva))
    return rows


def parse_xml_bytes(file_name: str, xml_bytes: bytes) -> Tuple[List[Dict[str, object]], Dict[str, object]]:
    debug = {"archivo": file_name, "formato": "", "filas": 0, "estatus": "OK", "mensaje": ""}
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        debug.update({"formato": "NO LEIDO", "estatus": "ERROR", "mensaje": f"XML mal formado: {e}"})
        return [], debug

    formato = detect_format(root)
    debug["formato"] = formato
    try:
        if formato == "K9":
            rows = parse_k9(root)
        elif formato == "ROYAN":
            rows = parse_royan(root)
        elif formato == "WASH N CROSS":
            rows = parse_wash(root)
        elif formato == SAT_GENERICO:
            rows = parse_sat_generico(root)
        else:
            rows = []
            debug.update({"estatus": "ERROR", "mensaje": "No se detecto como CFDI valido."})
        debug["filas"] = len(rows)
        return rows, debug
    except Exception as e:
        debug.update({"estatus": "ERROR", "mensaje": str(e)})
        return [], debug


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Consolidado")
        ws = writer.book["Consolidado"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)
    return output.getvalue()


def main():
    st.set_page_config(page_title="Consolidador XML CFDI", layout="wide")
    st.title("Consolidador de facturas XML CFDI")
    st.caption("Sube varios XML CFDI y descarga un Excel unico con todas las partidas.")

    uploaded = st.file_uploader("Archivos XML", type=["xml"], accept_multiple_files=True)

    if not uploaded:
        st.info("Sube uno o varios XML para iniciar.")
        return

    all_rows: List[Dict[str, object]] = []
    debug_rows: List[Dict[str, object]] = []

    for f in uploaded:
        rows, dbg = parse_xml_bytes(f.name, f.getvalue())
        all_rows.extend(rows)
        debug_rows.append(dbg)

    debug_df = pd.DataFrame(debug_rows)
    st.subheader("Debug de procesamiento")
    st.dataframe(debug_df, use_container_width=True)

    df = pd.DataFrame(all_rows, columns=FINAL_COLUMNS)
    st.subheader("Preview consolidado")
    st.dataframe(df, use_container_width=True)

    if df.empty:
        st.warning("No se generaron filas. Revisa errores en el debug.")
        return

    excel_bytes = dataframe_to_excel_bytes(df)
    st.download_button(
        "Descargar Excel consolidado",
        data=excel_bytes,
        file_name="facturas_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
