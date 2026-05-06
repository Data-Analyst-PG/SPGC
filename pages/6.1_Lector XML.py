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
    "EMPRESA", "# FACTURA", "UUID", "FECHA FACTURA", "FECHA Y HR SERVICIO REALIZADO",
    "# DE UNIDAD", "ACTIVIDAD", "CANTIDAD", "SUBTOTAL", "IVA", "TOTAL",
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
    return attr(find_first(root, "TimbreFiscalDigital"), "UUID")


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
    for traslado in find_all(concepto, "Traslado"):
        if traslado.attrib.get("Base"):
            return D(traslado.attrib.get("Base"))
    return D(concepto.attrib.get("Importe"))


def detect_format(root: ET.Element) -> str:
    emisor, _ = get_emisor_receptor(root)
    emisor_text = norm(" ".join([emisor.get("Nombre", ""), emisor.get("Rfc", "")]))
    for formato, needles in PROVEEDORES.items():
        if any(norm(n) in emisor_text for n in needles):
            return formato
    if local_name(root.tag) == "Comprobante" and get_concepts(root):
        return SAT_GENERICO
    return "NO DETECTADO"


def serie_folio(root: ET.Element) -> str:
    serie = attr(root, "Serie")
    folio = attr(root, "Folio")
    return f"{serie}-{folio}" if serie and folio else (folio or serie)


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


def concept_custom_value(concepto: ET.Element, keys: List[str]) -> str:
    wanted = [norm(k).replace(".", "").replace("_", "").replace(" ", "") for k in keys]
    for elem in [concepto] + list(concepto.iter()):
        for k, v in elem.attrib.items():
            nk = norm(k).replace(".", "").replace("_", "").replace(" ", "")
            if any(w in nk for w in wanted):
                return v
    return ""


def parse_k9(root: ET.Element) -> Tuple[List[Dict[str, object]], str]:
    h = common_header(root)
    text = complemento_text(root) or all_xml_text(root)
    factura = extract_order_k9(text) or h["folio"] or h["serie_folio"]
    fecha_serv = extract_service_datetime_k9(text)
    unidad = extract_unit_k9(text)
    rows = []
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = get_iva_from_concept(c) or subtotal * Decimal("0.08")
        rows.append(make_row(h["empresa"], factura, h["uuid"], h["fecha"], fecha_serv, unidad,
                             c.attrib.get("Descripcion", ""), D(c.attrib.get("Cantidad")), subtotal, iva))
    msg = ""
    if not (extract_order_k9(text) and fecha_serv and unidad):
        msg = "XML procesado, pero no trae comentarios K9 (orden/unidad/servicio); esos datos solo aparecen en el PDF si el proveedor no los incluye en Addenda."
    return rows, msg


def parse_royan(root: ET.Element) -> Tuple[List[Dict[str, object]], str]:
    h = common_header(root)
    rows = []
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = get_iva_from_concept(c) or subtotal * Decimal("0.16")
        rows.append(make_row(h["empresa"], h["folio"] or h["serie_folio"], h["uuid"], h["fecha"], "", "",
                             c.attrib.get("Descripcion", ""), Decimal("1"), subtotal, iva))
    return rows, "XML ROYAN procesado; el detalle de hoja 2 no viene en este XML, solo viene el concepto fiscal resumido."


def parse_wash(root: ET.Element) -> Tuple[List[Dict[str, object]], str]:
    h = common_header(root)
    rows = []
    missing_ref_obs = False
    for c in get_concepts(root):
        subtotal = D(c.attrib.get("Importe"))
        iva = get_iva_from_concept(c) or subtotal * Decimal("0.08")
        ref_pago = concept_custom_value(c, ["REFPAGO", "REF PAGO", "REF.PAGO", "REFERENCIA PAGO"])
        obs = concept_custom_value(c, ["OBS", "OBSERVACION", "OBSERVACIONES", "FECHA SERVICIO"])
        if not ref_pago or not obs:
            missing_ref_obs = True
        rows.append(make_row(h["empresa"], h["serie_folio"], h["uuid"], h["fecha"], obs, ref_pago,
                             c.attrib.get("Descripcion", ""), D(c.attrib.get("Cantidad")), subtotal, iva))
    msg = ""
    if missing_ref_obs:
        msg = "XML procesado, pero REF.PAGO y OBS no vienen dentro del XML CFDI; por eso # DE UNIDAD y fecha servicio quedan vacios. Esos datos aparecen en el PDF/representacion impresa."
    return rows, msg


def parse_sat_generico(root: ET.Element) -> Tuple[List[Dict[str, object]], str]:
    h = common_header(root)
    rows = []
    for c in get_concepts(root):
        descripcion = c.attrib.get("Descripcion", "")
        subtotal = get_base_from_concept(c)
        iva = get_iva_from_concept(c)
        cantidad = D(c.attrib.get("Cantidad"))
        cantidad_out = int(cantidad) if cantidad == cantidad.to_integral() else float(cantidad)
        rows.append(make_row(h["empresa"], h["folio"] or h["serie_folio"], h["uuid"], h["fecha"], "",
                             unit_from_description_after_colon(descripcion), descripcion, cantidad_out, subtotal, iva))
    return rows, ""


def parse_xml_bytes(file_name: str, xml_bytes: bytes) -> Tuple[List[Dict[str, object]], Dict[str, object]]:
    debug = {"archivo": file_name, "formato": "", "filas": 0, "estatus": "OK", "mensaje": ""}
    if not xml_bytes or len(xml_bytes.strip()) == 0:
        debug.update({"formato": "NO LEIDO", "estatus": "ERROR", "mensaje": "Archivo XML vacio (0 bytes)."})
        return [], debug
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        debug.update({"formato": "NO LEIDO", "estatus": "ERROR", "mensaje": f"XML mal formado: {e}"})
        return [], debug

    formato = detect_format(root)
    debug["formato"] = formato
    try:
        if formato == "K9":
            rows, msg = parse_k9(root)
        elif formato == "ROYAN":
            rows, msg = parse_royan(root)
        elif formato == "WASH N CROSS":
            rows, msg = parse_wash(root)
        elif formato == SAT_GENERICO:
            rows, msg = parse_sat_generico(root)
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
    st.caption("Sube varios XML CFDI y descarga un Excel unico con todas las partidas encontradas dentro del XML.")

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

    st.subheader("Debug de procesamiento")
    st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)

    df = pd.DataFrame(all_rows, columns=FINAL_COLUMNS)
    st.subheader("Preview consolidado")
    st.dataframe(df, use_container_width=True)

    if df.empty:
        st.warning("No se generaron filas. Revisa errores en el debug.")
        return

    st.download_button(
        "Descargar Excel consolidado",
        data=dataframe_to_excel_bytes(df),
        file_name="facturas_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
