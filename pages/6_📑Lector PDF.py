# app.py
import re
import io
from dataclasses import asdict, dataclass
from typing import List, Dict, Any, Tuple

import pandas as pd
import streamlit as st

# Recomendado: pip install pdfplumber openpyxl
import pdfplumber


# ----------------------------
# Helpers
# ----------------------------
def norm_num(s: str) -> float:
    """Convierte '3,650.00' -> 3650.0"""
    s = (s or "").strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def find_first(pattern: str, text: str, flags=0) -> str:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""


def extract_pages_text(pdf_bytes: bytes) -> List[str]:
    pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages:
            pages.append(p.extract_text() or "")
    return pages


def to_excel_bytes(header: Dict[str, Any], items: List[Dict[str, Any]]) -> bytes:
    header_df = pd.DataFrame([header])
    items_df = pd.DataFrame(items)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        header_df.to_excel(writer, index=False, sheet_name="Encabezado")
        items_df.to_excel(writer, index=False, sheet_name="Conceptos")
    return output.getvalue()


# ----------------------------
# Parsers
# ----------------------------
def parse_royan(pdf_bytes: bytes) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
    """
    ROYAN (según tu ejemplo): encabezado + hoja 2 con partidas tipo:
    '517.24 Actividad REVISAR MASAS ACT'
    y en hoja 1 los totales: Subtotal / Iva 16 % / Ret 4 % / Total
    """
    pages = extract_pages_text(pdf_bytes)
    full = "\n".join(pages)

    # Encabezado (heurístico basado en el PDF ejemplo)
    folio = find_first(r"\b(ROYAN-\d+)\b", full)
    fecha = find_first(r"\b(\d{2}/\d{2}/\d{4})\b", full)

    cliente = find_first(r"\n([A-ZÁÉÍÓÚÑ ]{5,})\n[A-Z0-9]{12,13}\n", full)  # nombre antes de RFC
    rfc_cliente = find_first(r"\n([A-Z0-9]{12,13})\n", full)

    # Totales (en tu ejemplo salen como líneas con $)
    subtotal = find_first(r"\$(\d[\d,]*\.\d{2})\s*\n\$(\d[\d,]*\.\d{2})\s*\n\$\.*0*\.?0*\s*\n\$(\d[\d,]*\.\d{2})", full)
    # Lo anterior puede fallar si cambia el orden; entonces buscamos por etiquetas:
    subtotal2 = find_first(r"\n\$(\d[\d,]*\.\d{2})\s*\n\$(\d[\d,]*\.\d{2})", full)

    sub = ""
    iva = ""
    ret = ""
    total = ""

    # Método robusto por “bloque” de etiquetas
    # (En el ejemplo: $27,867.06 / $4,458.73 / $.00 / $32,325.79)
    money = re.findall(r"\$(\d[\d,]*\.\d{2}|\.\d{2})", full)
    # Tomamos los últimos 4 importes como subtotal/iva/ret/total si existen
    if len(money) >= 4:
        sub, iva, ret, total = money[-4], money[-3], money[-2], money[-1]

    header = {
        "proveedor_formato": "ROYAN",
        "folio": folio,
        "fecha": fecha,
        "cliente": cliente,
        "rfc_cliente": rfc_cliente,
        "subtotal": norm_num(sub),
        "iva": norm_num(iva),
        "retencion": norm_num(ret),
        "total": norm_num(total),
    }

    # Partidas: normalmente vienen en hoja 2
    items: List[Dict[str, Any]] = []
    for page_text in pages:
        for line in (page_text or "").splitlines():
            line = line.strip()

            # patrón típico hoja 2: "517.24 Actividad REVISAR MASAS ACT"
            m = re.match(r"^(?P<importe>[\d,]+\.\d{2})\s+Actividad\s+(?P<desc>.+?)\s+ACT$", line)
            if m:
                items.append(
                    {
                        "clave": "",
                        "unidad": "ACT",
                        "cantidad": 1,
                        "descripcion": m.group("desc").strip(),
                        "precio_unitario": norm_num(m.group("importe")),
                        "importe": norm_num(m.group("importe")),
                        "pagina_origen": "detalle",
                    }
                )

    # Si no encontró partidas hoja 2, intenta capturar la partida “principal” de hoja 1
    if not items:
        # ejemplo hoja 1: "E48 78181500 ... $27,867.06 ... $27,867.06"
        for line in full.splitlines():
            if "78181500" in line and "$" in line:
                # capturamos algo básico
                items.append(
                    {
                        "clave": "78181500",
                        "unidad": "E48",
                        "cantidad": 1,
                        "descripcion": re.sub(r"\s+", " ", re.sub(r"\$.*", "", line)).strip(),
                        "precio_unitario": float("nan"),
                        "importe": float("nan"),
                        "pagina_origen": "resumen",
                    }
                )

    return header, items


def parse_k9(pdf_bytes: bytes) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
    """
    K9 (según tu ejemplo): encabezado con FACTURA A000..., fecha,
    y tabla con líneas tipo:
    '78181507 REPARACION LLANTA POS 5 SERVICIO 1 450.00 450.00'
    """
    pages = extract_pages_text(pdf_bytes)
    full = "\n".join(pages)

    factura = find_first(r"\b(A\d{10})\b", full)  # A0000006866
    # fecha/hora: "05/02/2026 10:37 a.m." aparece 2 veces en tu ejemplo
    fecha = find_first(r"\b(\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}\s*[ap]\.m\.)\b", full, flags=re.IGNORECASE)
    rfc_emisor = find_first(r"\bR\.F\.C\.\s*([A-Z0-9]{12,13})\b", full)
    vendedor_a = find_first(r"VENDIDO A:\s*(.+?)\n", full)

    # Totales (en el ejemplo: Subtotal 3,650.00 / IVA 292.00 / Total 3,942.00)
    # Ojo: el texto trae "Total Retencion Subtotal IVA( 8 %)" en distinto orden.
    total = find_first(r"\bTotal\s*\n?\s*([\d,]+\.\d{2})", full)
    subtotal = find_first(r"\bSubtotal\s*\n?\s*([\d,]+\.\d{2})", full)
    iva = find_first(r"\bIVA\(\s*8\s*%\)\s*\n?\s*([\d,]+\.\d{2})", full)
    ret = find_first(r"\bRetencion\s*\n?\s*([\d,]+\.\d{2})", full)

    header = {
        "proveedor_formato": "K9",
        "factura": factura,
        "fecha": fecha,
        "vendedor_a": vendedor_a,
        "rfc_emisor": rfc_emisor,
        "subtotal": norm_num(subtotal),
        "iva": norm_num(iva),
        "retencion": norm_num(ret),
        "total": norm_num(total),
    }

    # Parseo de conceptos (líneas con clave de 8 dígitos)
    items: List[Dict[str, Any]] = []

    # Algunas descripciones se parten en varias líneas; hacemos un “acumulador”
    pending_desc = ""
    pending_clave = ""
    pending_unidad = ""

    for line in full.splitlines():
        s = re.sub(r"\s+", " ", line.strip())
        if not s:
            continue

        # Línea completa típica:
        # 78181507 CALIBRADO ... SERVICIO 8 100.00 800.00
        m = re.match(
            r"^(?P<clave>\d{8})\s+(?P<desc>.+?)\s+(?P<unidad>[A-ZÁÉÍÓÚÑ]+)\s+(?P<cant>[\d,]+)\s+(?P<precio>[\d,]+\.\d{2})\s+(?P<importe>[\d,]+\.\d{2})$",
            s,
            flags=re.IGNORECASE,
        )
        if m:
            items.append(
                {
                    "clave": m.group("clave"),
                    "unidad": m.group("unidad").upper(),
                    "cantidad": int(m.group("cant").replace(",", "")),
                    "descripcion": m.group("desc").strip(),
                    "precio_unitario": norm_num(m.group("precio")),
                    "importe": norm_num(m.group("importe")),
                }
            )
            pending_desc = ""
            pending_clave = ""
            pending_unidad = ""
            continue

        # Si inicia con clave pero NO trae montos, acumulamos descripción
        m2 = re.match(r"^(?P<clave>\d{8})\s+(?P<rest>.+)$", s)
        if m2:
            pending_clave = m2.group("clave")
            pending_desc = m2.group("rest")
            pending_unidad = ""
            continue

        # Si estamos acumulando, puede venir línea que termina con "... SERVICIO 1 1200.00 1200.00"
        if pending_clave:
            m3 = re.match(
                r"^(?P<desc2>.+?)\s+(?P<unidad>[A-ZÁÉÍÓÚÑ]+)\s+(?P<cant>[\d,]+)\s+(?P<precio>[\d,]+\.\d{2})\s+(?P<importe>[\d,]+\.\d{2})$",
                s,
                flags=re.IGNORECASE,
            )
            if m3:
                full_desc = (pending_desc + " " + m3.group("desc2")).strip()
                items.append(
                    {
                        "clave": pending_clave,
                        "unidad": m3.group("unidad").upper(),
                        "cantidad": int(m3.group("cant").replace(",", "")),
                        "descripcion": full_desc,
                        "precio_unitario": norm_num(m3.group("precio")),
                        "importe": norm_num(m3.group("importe")),
                    }
                )
                pending_desc = ""
                pending_clave = ""
                pending_unidad = ""
            else:
                # sigue siendo parte de la descripción
                pending_desc = (pending_desc + " " + s).strip()

    return header, items


# ----------------------------
# Streamlit UI
# ----------------------------
st.set_page_config(page_title="Lector PDF → Excel (K9 / ROYAN)", layout="centered")
st.title("Lector de facturas PDF → Excel")
st.caption("Elige el formato, sube el PDF y descarga el Excel.")

formato = st.selectbox("Selecciona el formato de factura", ["ROYAN", "K9"])

pdf_file = st.file_uploader("Sube tu factura en PDF", type=["pdf"])

col1, col2 = st.columns(2)
with col1:
    procesar = st.button("Procesar")
with col2:
    autodetect = st.checkbox("Autodetectar formato (si falla, usa el selector)", value=True)

if procesar:
    if not pdf_file:
        st.error("Primero sube un PDF.")
        st.stop()

    pdf_bytes = pdf_file.read()

    # Autodetección básica por keywords (puedes ampliarla)
    if autodetect:
        pages = extract_pages_text(pdf_bytes)
        full = "\n".join(pages)
        if "ROYAN-" in full:
            formato_final = "ROYAN"
        elif "K9" in full or "A000" in full or "FACTURA" in full:
            formato_final = "K9"
        else:
            formato_final = formato  # fallback
    else:
        formato_final = formato

    try:
        if formato_final == "ROYAN":
            header, items = parse_royan(pdf_bytes)
        else:
            header, items = parse_k9(pdf_bytes)

        if not items:
            st.warning("No se encontraron conceptos/partidas. El PDF podría ser escaneado o cambió el formato.")
        else:
            st.success(f"Listo: {len(items)} partidas encontradas ({formato_final}).")

        st.subheader("Encabezado detectado")
        st.dataframe(pd.DataFrame([header]), use_container_width=True)

        st.subheader("Conceptos / partidas")
        st.dataframe(pd.DataFrame(items), use_container_width=True)

        excel_bytes = to_excel_bytes(header, items)
        out_name = f"{formato_final}_{pdf_file.name.replace('.pdf','')}.xlsx"
        st.download_button(
            "Descargar Excel",
            data=excel_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error procesando el PDF: {e}")
        st.info("Tip: si el PDF es una imagen escaneada, hay que añadir OCR (pytesseract) o usar otro extractor.")
