import os
from datetime import datetime, timezone

import streamlit as st
from supabase import create_client, Client
import smtplib
from email.message import EmailMessage

# ====== SUPABASE CLIENT ======
def get_supabase_client() -> Client:
    # Puedes usar st.secrets o variables de entorno, según tengas
    url = st.secrets.get("SUPABASE_URL", os.getenv("SUPABASE_URL"))
    key = st.secrets.get("SUPABASE_KEY", os.getenv("SUPABASE_KEY"))
    if not url or not key:
        st.error("Faltan SUPABASE_URL o SUPABASE_KEY en secrets/variables de entorno.")
        st.stop()
    return create_client(url, key)

supabase = get_supabase_client()

# ====== CONFIG CORREO (puedes apagarlo con ENABLE_EMAIL) ======
ENABLE_EMAIL = st.secrets.get("ENABLE_EMAIL", os.getenv("ENABLE_EMAIL", "false")).lower() == "true"
SMTP_HOST = st.secrets.get("SMTP_HOST", os.getenv("SMTP_HOST"))
SMTP_PORT = int(st.secrets.get("SMTP_PORT", os.getenv("SMTP_PORT", "587")))
SMTP_USER = st.secrets.get("SMTP_USER", os.getenv("SMTP_USER"))   # cuenta genérica
SMTP_PASSWORD = st.secrets.get("SMTP_PASSWORD", os.getenv("SMTP_PASSWORD"))


def folio_visible_from_id(row_id: int) -> str:
    """Convierte el id numérico en un folio tipo #00001."""
    return f"#{row_id:05d}"


def enviar_correo_solicitud(data: dict, factura_pdf=None):
    """Envía correo de notificación de nueva solicitud (si está habilitado)."""
    if not ENABLE_EMAIL:
        # Simplemente no envía, para que puedas probar sin correo
        return False, "Envío de correo desactivado (ENABLE_EMAIL = false)."

    if not (SMTP_HOST and SMTP_PORT and SMTP_USER and SMTP_PASSWORD):
        return False, "Configuración SMTP incompleta."

    msg = EmailMessage()

    asunto = f"Nueva solicitud complementaria {data.get('folio')} - {data.get('empresa')}"
    msg["Subject"] = asunto
    msg["From"] = SMTP_USER
    msg["To"] = "auditoria.operaciones@palosgarza.com, abel.chontal@palosgarza.com"

    cuerpo = f"""
Se ha registrado una nueva solicitud de complemento.

Folio: {data.get('folio')}
Empresa: {data.get('empresa')}

Número de tráfico: {data.get('numero_trafico')}
Concepto correcto: {data.get('concepto_correcto')}
Monto correcto: {data.get('monto_correcto')}
Proveedor correcto: {data.get('proveedor_correcto')}

Motivo de la modificación:
{data.get('motivo_modificacion')}

Solicitante:
{data.get('nombre_solicitante')} ({data.get('correo_solicitante')})

Link hilo de correo:
{data.get('link_hilo_correo') or 'N/A'}

Links de evidencias:
{data.get('link_evidencias') or 'N/A'}

PDF factura (Supabase):
{data.get('factura_archivo') or 'N/A'}

Estatus inicial: {data.get('estatus')}
Fecha de captura: {data.get('fecha_captura')}
    """.strip()

    msg.set_content(cuerpo)

    # Adjuntar PDF si viene
    if factura_pdf is not None:
        pdf_bytes = factura_pdf.getvalue()
        msg.add_attachment(
            pdf_bytes,
            maintype="application",
            subtype="pdf",
            filename=factura_pdf.name
        )

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)
def modulo_solicitud_complementaria():
    st.header("Solicitud de factura complementaria")

    with st.form("form_solicitud_complementaria"):
        empresa = st.selectbox(
            "Empresa",
            ["Set Freight", "Lincoln", "Set Logis", "Picus", "Igloo"],
            index=None,
            placeholder="Selecciona una empresa"
        )

        numero_trafico = st.text_input("Número de tráfico")
        factura_pdf = st.file_uploader(
            "Factura archivo (PDF)",
            type=["pdf"],
            accept_multiple_files=False
        )

        concepto_correcto = st.text_input("Concepto correcto")
        monto_correcto = st.number_input("Monto correcto", min_value=0.0, step=0.01, format="%.2f")
        proveedor_correcto = st.text_input("Proveedor correcto")
        motivo_modificacion = st.text_area("Motivo de la modificación")

        nombre_solicitante = st.text_input("Nombre del solicitante")
        correo_solicitante = st.text_input("Correo del solicitante")

        link_hilo_correo = st.text_input("Link del hilo de correo (Outlook / OWA)")
        link_evidencias = st.text_area("Links de fotos / evidencias (opcional)")

        evidencias_files = st.file_uploader(
            "Subir fotos / evidencias (opcional)",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True
        )

        submitted = st.form_submit_button("Enviar solicitud")

    if not submitted:
        return

    # ===== Validaciones mínimas =====
    if not empresa:
        st.error("Debes seleccionar una empresa.")
        return
    if not numero_trafico:
        st.error("El campo 'Número de tráfico' es obligatorio.")
        return
    if factura_pdf is None:
        st.error("Debes subir el PDF de la factura.")
        return
    if not motivo_modificacion:
        st.error("Por favor indica el motivo de la modificación.")
        return
    if not nombre_solicitante or not correo_solicitante:
        st.error("Nombre y correo del solicitante son obligatorios.")
        return

    fecha_captura = datetime.now(timezone.utc).isoformat()
    estatus_inicial = "Pendiente"

    # ===== 1) Subir PDF a Storage =====
    factura_url = None
    try:
        bucket_pdf = supabase.storage.from_("complementarias")
        pdf_path = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{factura_pdf.name}"
        upload_res = bucket_pdf.upload(pdf_path, factura_pdf.getvalue())
        # si no lanza excepción, asumimos éxito
        factura_url = bucket_pdf.get_public_url(pdf_path)
    except Exception as e:
        st.warning(f"No se pudo subir el PDF a Storage: {e}")

    # ===== 2) Subir evidencias (imágenes) =====
    evidencias_urls = []
    if evidencias_files:
        bucket_img = supabase.storage.from_("complementarias-evidencias")
        for i, img in enumerate(evidencias_files, start=1):
            try:
                img_path = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_evidencia_{i}_{img.name}"
                bucket_img.upload(img_path, img.getvalue())
                url = bucket_img.get_public_url(img_path)
                evidencias_urls.append(url)
            except Exception as e:
                st.warning(f"No se pudo subir la evidencia {img.name}: {e}")

    # ===== 3) Insertar registro en la tabla =====
    # Ponemos folio temporal, luego lo actualizamos con el id
    data_insert = {
        "folio": "PENDIENTE",
        "empresa": empresa,
        "numero_trafico": numero_trafico,
        "factura_archivo": factura_url,
        "concepto_correcto": concepto_correcto,
        "monto_correcto": float(monto_correcto),
        "proveedor_correcto": proveedor_correcto,
        "motivo_modificacion": motivo_modificacion,
        "nombre_solicitante": nombre_solicitante,
        "correo_solicitante": correo_solicitante,
        "link_hilo_correo": link_hilo_correo,
        "link_evidencias": link_evidencias,
        "evidencias_urls": evidencias_urls or None,
        "estatus": estatus_inicial,
        "fecha_captura": fecha_captura,
        "fecha_resuelto": None,
        "auditor": None,
    }

    try:
        res = supabase.table("solicitudes_complementarias").insert(data_insert).execute()
        if not res.data:
            st.error("No se pudo insertar la solicitud en la base de datos.")
            return
        row = res.data[0]
        row_id = row["id"]
    except Exception as e:
        st.error(f"Error al guardar en la base de datos: {e}")
        return

    # ===== 4) Actualizar folio con formato #00001 =====
    folio_def = folio_visible_from_id(row_id)
    try:
        supabase.table("solicitudes_complementarias").update({"folio": folio_def}).eq("id", row_id).execute()
    except Exception as e:
        st.warning(f"La solicitud se guardó, pero no se pudo actualizar el folio: {e}")

    # Actualizamos el dict para usarlo en el correo
    data_insert["folio"] = folio_def

    # ===== 5) Enviar correo (opcional) =====
    ok_mail, mail_error = enviar_correo_solicitud(data_insert, factura_pdf=factura_pdf)

    if ok_mail:
        st.success(f"Solicitud registrada correctamente. Folio: {folio_def}")
    else:
        st.warning(
            f"Solicitud registrada correctamente. Folio: {folio_def}\n"
            f"Pero hubo un problema al enviar el correo: {mail_error}"
        )
