import os
from datetime import datetime, timezone

import streamlit as st
from supabase import create_client, Client

# =========================
# SUPABASE CLIENT
# =========================
def get_supabase_client() -> Client:
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
    except Exception:
        url = os.getenv("SUPABASE_URL")
        key = os.getenv("SUPABASE_KEY")

    if not url or not key:
        st.error("Faltan SUPABASE_URL o SUPABASE_KEY en secrets/variables de entorno.")
        st.stop()
    return create_client(url, key)


supabase = get_supabase_client()

# =========================
# CATÁLOGOS (UI)
# =========================
EMPRESAS = ["Set Freight", "Lincoln", "Set Logis Plus", "Picus", "Igloo"]
MONEDAS = ["MXN", "USD"]

TIPOS_CONCEPTO = [
    "OTROS",
    "TIPO MOVIMIENTO",
    "FLETE MX",
    "AUTOPISTAS",
    "PAD",
    "FLETE USA",
    "CRUCE",
    "DOMESTICO MX",
    "DOMESTICO USA",
    "ALMACENAJE",
    "MANIOBRAS",
    "SUELDO CARGADO",
    "CARGA",
    "DESCARGA",
    "TRASBORDO",
    "BONO",
    "SUELDO",
    "GRUA",
]

PLATAFORMAS_POR_EMPRESA = {
    "Lincoln": ["STAR USA"],
    "Set Freight": [
        "STAR 2.0 SET FREIGHT",
        "STAR 2.0 PALOS GARZA LOGISMEX",
    ],
    "Set Logis Plus": [
        "STAR USA",
        "STAR 2.0 SET FREIGHT",
        "STAR 2.0 PALOS GARZA LOGISMEX",
    ],
    "Picus": [
        "STAR 2.0 PALOS GARZA LOGISTIC",
        "STAR 2.0 PALOS GARZA LOGISMEX",
    ],
    "Igloo": [
        "STAR 2.0 PALOS GARZA LOGISTIC",
        "STAR 2.0 PALOS GARZA LOGISMEX",
    ],
}

FALLBACK_CONCEPTOS = {
    "OTROS": [
        "EXTRA STOP/CT. PARADA EXTRA",
        "HANDLING CHARGES/CT. MANIOBRAS",
        "LAY OVER/CT. ESTANCIAS",
        "LOADLOCKS/ CT.GATAS/BLOQUEOS",
        "LOGISTICS COORDINATION/ CT. COORDINACION LOGISTICA",
        "LUMPER FEES/ CT. DESCARGA",
        "SALES EXPENSES 1/CT. GASTOS DE VENTA",
        "SALES EXPENSES 2/CT. GASTOS DE VENTA",
        "SALES EXPENSES 3/CT. GASTOS DE VENTA",
        "SCALE / CT. BASCULA RB",
        "STORAGE COSTS/CT. ALMACENAJES",
        "TEAM DRIVER /CT. DOBLE OPERADOR",
        "THERMO RENT/CT. RENTA DE THERMO",
        "TIRES /CT.LLANTAS",
        "TNU - TRUCK NOT USED/CT. MOVIMIENTO EN FALSO",
        "TRAILER PARTS /CT. REFACCIONES",
        "TRAILER REPAIR & OTHER EXPENSES/CT. REP. Y OTROS GASTOS DE VIAJE",
        "TRANSLOAD/ CT. TRANSBORDO",
    ],
    "GRUA": [],
}

@st.cache_data(ttl=300)
def load_conceptos_from_supabase() -> dict[str, list[str]]:
    try:
        res = (
            supabase.table("catalogo_conceptos")
            .select("tipo_concepto, concepto, activo")
            .execute()
        )
        rows = res.data or []
        if not rows:
            return {}

        conceptos: dict[str, list[str]] = {}
        for r in rows:
            if "activo" in r and r["activo"] is False:
                continue
            t = (r.get("tipo_concepto") or "").strip()
            c = (r.get("concepto") or "").strip()
            if not t or not c:
                continue
            conceptos.setdefault(t, []).append(c)

        for t in conceptos:
            conceptos[t] = sorted(set(conceptos[t]))
        return conceptos
    except Exception:
        return {}

CONCEPTOS_DB = load_conceptos_from_supabase()

def get_conceptos(tipo: str) -> list[str]:
    if tipo in CONCEPTOS_DB:
        return CONCEPTOS_DB[tipo]
    return FALLBACK_CONCEPTOS.get(tipo, [])


# =========================
# UI
# =========================
st.header("Registro de complementaria (solo registro)")

tab_captura, tab_auditor = st.tabs(["📝 Captura", "🕵️ Auditor"])


def bloque_concepto(prefix: str, titulo: str):
    st.subheader(titulo)
    col1, col2 = st.columns(2)

    with col1:
        tipo = st.selectbox(
            "Tipo Concepto",
            TIPOS_CONCEPTO,
            key=f"{prefix}_tipo",
            index=None,
            placeholder="Selecciona un tipo",
        )

    conceptos = get_conceptos(tipo) if tipo else []
    concepto_disabled = (not tipo) or (len(conceptos) == 0)

    with col2:
        if concepto_disabled:
            st.selectbox(
                "Concepto",
                ["Sin datos para mostrar"] if tipo else ["Selecciona primero un tipo"],
                key=f"{prefix}_concepto",
                disabled=True,
            )
        else:
            st.selectbox(
                "Concepto",
                conceptos,
                key=f"{prefix}_concepto",
                index=None,
                placeholder="Selecciona un concepto",
            )

    col3, col4 = st.columns(2)
    with col3:
        proveedor = st.text_input("Proveedor", key=f"{prefix}_proveedor")
        moneda = st.selectbox("Moneda", MONEDAS, key=f"{prefix}_moneda", index=None, placeholder="Selecciona moneda")
    with col4:
        importe = st.number_input("Importe", min_value=0.0, step=0.01, format="%.2f", key=f"{prefix}_importe")

    return {
        "tipo": tipo,
        "concepto": st.session_state.get(f"{prefix}_concepto"),
        "proveedor": proveedor,
        "moneda": moneda,
        "importe": float(importe),
    }


# =========================
# TAB CAPTURA
# =========================
with tab_captura:
    st.text_input("Fecha", value=datetime.now().strftime("%d/%m/%Y"), disabled=True)

    c1, c2 = st.columns(2)
    with c1:
        empresa = st.selectbox("Empresa", EMPRESAS, index=None, placeholder="Selecciona una empresa")

    plataformas_opciones = PLATAFORMAS_POR_EMPRESA.get(empresa, [])
    with c2:
        plataforma = st.selectbox(
            "Plataforma",
            plataformas_opciones,
            index=None,
            placeholder="Selecciona una plataforma" if empresa else "Selecciona primero una empresa",
            disabled=(empresa is None),
        )

    solicitante = st.text_input("Solicitante")
    motivo_solicitud = st.text_area("Motivo de la solicitud")
    numero_trafico = st.text_input("Número de tráfico")

    st.divider()
    actual = bloque_concepto("actual", "Datos actuales (como están)")
    st.divider()
    nuevo = bloque_concepto("nuevo", "Datos correctos (como deben quedar)")

    st.divider()
    registrar = st.button("Registrar", type="primary")

    if registrar:
        errores = []
        if not empresa: errores.append("Debes seleccionar una empresa.")
        if not plataforma: errores.append("Debes seleccionar una plataforma.")
        if not solicitante.strip(): errores.append("El campo 'Solicitante' es obligatorio.")
        if not motivo_solicitud.strip(): errores.append("El campo 'Motivo de la solicitud' es obligatorio.")
        if not numero_trafico.strip(): errores.append("El campo 'Número de tráfico' es obligatorio.")

        for label, block in [("actual", actual), ("correcto", nuevo)]:
            if not block["tipo"]:
                errores.append(f"Debes seleccionar 'Tipo Concepto' ({label}).")
            if block["tipo"] and len(get_conceptos(block["tipo"])) > 0 and (not block["concepto"] or "Sin datos" in str(block["concepto"])):
                errores.append(f"Debes seleccionar 'Concepto' ({label}).")
            if not block["proveedor"].strip():
                errores.append(f"Debes capturar 'Proveedor' ({label}).")
            if not block["moneda"]:
                errores.append(f"Debes seleccionar 'Moneda' ({label}).")

        if errores:
            for e in errores:
                st.error(e)
            st.stop()
    
        fecha_captura = datetime.now(timezone.utc).isoformat()

        data_insert = {
            "fecha_captura": fecha_captura,
            "estatus": "Pendiente",

            "empresa": empresa,
            "plataforma": plataforma,
            "solicitante": solicitante.strip(),
            "motivo_solicitud": motivo_solicitud.strip(),
            "numero_trafico": numero_trafico.strip(),

            "tipo_concepto_actual": actual["tipo"],
            "concepto_actual": None if "Sin datos" in str(actual["concepto"]) else actual["concepto"],
            "proveedor_actual": actual["proveedor"].strip(),
            "moneda_actual": actual["moneda"],
            "importe_actual": float(actual["importe"]),

            "tipo_concepto_nuevo": nuevo["tipo"],
            "concepto_nuevo": None if "Sin datos" in str(nuevo["concepto"]) else nuevo["concepto"],
            "proveedor_nuevo": nuevo["proveedor"].strip(),
            "moneda_nuevo": nuevo["moneda"],
            "importe_nuevo": float(nuevo["importe"]),

            "fecha_resuelto": None,
            "auditor": None,
        }

        try:
            res = supabase.table("solicitudes_complementarias").insert(data_insert).execute()
            if not res.data:
                st.error("No se pudo insertar la solicitud en la base de datos.")
                st.stop()

            # ✅ folio identity (NO id)
            folio_num = int(res.data[0]["folio"])

        except Exception as e:
            st.error(f"Error al guardar en la base de datos: {e}")
            st.stop()

        folio_formateado = f"{folio_num:04d}"

        st.success(
            "Tú solicitud se ha capturado correctamente, favor de enviar el siguiente texto "
            "al correo auditoria.operaciones@palosgarza.com"
        )
        st.code(
            f"Mi folio de complementaria es el '#{folio_formateado}', favor de atender mi solicitud",
            language="text",
        )


# =========================
# TAB AUDITOR
# =========================
with tab_auditor:
    st.subheader("Solicitudes registradas")

    auditor_pwd = st.text_input("Contraseña auditor", type="password")

    secret_pwd = st.secrets.get("AUDITOR_PASSWORD")
    if not secret_pwd:
        st.error("No existe AUDITOR_PASSWORD en secrets. Agrégala en Settings > Secrets.")
        st.stop()

    if auditor_pwd == "":
        st.info("Ingresa la contraseña para ver las solicitudes.")
        st.stop()

    if auditor_pwd != secret_pwd:
        st.error("Contraseña incorrecta.")
        st.stop()

    ESTATUS_OPCIONES = ["Pendiente", "En revisión", "Cancelado", "Resuelto"]

    colf1, colf2, colf3 = st.columns(3)
    with colf1:
        f_empresa = st.selectbox("Empresa", ["(Todas)"] + EMPRESAS, index=0)
    with colf2:
        f_estatus = st.selectbox("Estatus", ["(Todos)"] + ESTATUS_OPCIONES, index=0)
    with colf3:
        texto = st.text_input("Buscar (folio / solicitante / tráfico)")

    q = supabase.table("solicitudes_complementarias").select(
        "folio, fecha_captura, estatus, empresa, plataforma, solicitante, numero_trafico, motivo_solicitud, fecha_resuelto"
    )

    if f_empresa != "(Todas)":
        q = q.eq("empresa", f_empresa)
    if f_estatus != "(Todos)":
        q = q.eq("estatus", f_estatus)

    q = q.order("folio", desc=True).limit(500)

    try:
        res = q.execute()
        rows = res.data or []
    except Exception as e:
        st.error(f"No se pudieron cargar solicitudes: {e}")
        st.stop()

    if texto.strip():
        t = texto.strip().lower()
        def match(r):
            return (
                t in f"{int(r.get('folio', 0)):04d}".lower()
                or t in str(r.get("solicitante", "")).lower()
                or t in str(r.get("numero_trafico", "")).lower()
            )
        rows = [r for r in rows if match(r)]

    st.write(f"Total: {len(rows)}")

    for r in rows:
        folio_num = int(r["folio"])
        folio_fmt = f"{folio_num:04d}"
        estatus_actual = r.get("estatus") or "Pendiente"

        with st.expander(f"Folio #{folio_fmt} | {r.get('empresa')} | {estatus_actual}"):
            st.write(f"**Fecha:** {r.get('fecha_captura')}")
            st.write(f"**Plataforma:** {r.get('plataforma')}")
            st.write(f"**Solicitante:** {r.get('solicitante')}")
            st.write(f"**Tráfico:** {r.get('numero_trafico')}")
            st.write(f"**Motivo:** {r.get('motivo_solicitud')}")
            if r.get("fecha_resuelto"):
                st.write(f"**Fecha resuelto/cancelado:** {r.get('fecha_resuelto')}")

            st.divider()
            c1, c2 = st.columns([2, 1])

            with c1:
                nuevo_estatus = st.selectbox(
                    "Cambiar estatus",
                    ESTATUS_OPCIONES,
                    index=ESTATUS_OPCIONES.index(estatus_actual) if estatus_actual in ESTATUS_OPCIONES else 0,
                    key=f"estatus_{folio_num}",
                )

            with c2:
                if st.button("Guardar estatus", key=f"btn_guardar_{folio_num}"):
                    update_payload = {"estatus": nuevo_estatus}

                    # Si se resuelve o cancela, guardamos fecha_resuelto
                    if nuevo_estatus in ["Resuelto", "Cancelado"]:
                        update_payload["fecha_resuelto"] = datetime.now(timezone.utc).isoformat()
                    else:
                        update_payload["fecha_resuelto"] = None

                    try:
                        # ✅ Actualiza por folio (NO por id)
                        supabase.table("solicitudes_complementarias").update(update_payload).eq("folio", folio_num).execute()
                        st.success(f"Estatus actualizado para folio #{folio_fmt}")
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo actualizar el estatus: {e}")
