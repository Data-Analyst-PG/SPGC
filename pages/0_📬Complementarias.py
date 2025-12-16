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

def folio_visible_from_id(row_id: int) -> str:
    return f"#{row_id:05d}"

# =========================
# CATÁLOGOS (UI)
# =========================
EMPRESAS = ["Set Freight", "Lincoln", "Set Logis Plus", "Picus", "Igloo"]
PLATAFORMAS = [
    "Star USA (LINCOLB, SLP)",
    "STAR 2.0 (SET FREIGHT, PALOS GARZA LOGISTICS, PALOS GARZA LOGISMEX)",
]
MONEDAS = ["MXN", "USD"]

# Tipos vistos en tu modal (video)
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

# Fallback mínimo (si aún no tienes catálogo en Supabase).
# Aquí puedes ir agregando más tipos/valores si lo necesitas.
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
    "GRUA": [],  # en tu video aparece "Sin datos para mostrar"
}

@st.cache_data(ttl=300)
def load_conceptos_from_supabase() -> dict[str, list[str]]:
    """
    Intenta cargar el catálogo desde Supabase:
    tabla sugerida: catalogo_conceptos(tipo_concepto, concepto, activo)
    """
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

        # Ordena conceptos para que se vean “limpios”
        for t in conceptos:
            conceptos[t] = sorted(set(conceptos[t]))
        return conceptos
    except Exception:
        return {}

CONCEPTOS_DB = load_conceptos_from_supabase()

def get_conceptos(tipo: str) -> list[str]:
    # Prioridad: Supabase; si no hay, usa fallback
    if tipo in CONCEPTOS_DB:
        return CONCEPTOS_DB[tipo]
    return FALLBACK_CONCEPTOS.get(tipo, [])

# =========================
# UI
# =========================
st.header("Registro de complementaria (solo registro, sin correo)")

# 1) Fecha automática (solo se muestra)
st.text_input("Fecha", value=datetime.now().strftime("%d/%m/%Y"), disabled=True)

# 2) Empresa / Plataforma
c1, c2 = st.columns(2)
with c1:
    empresa = st.selectbox("Empresa", EMPRESAS, index=None, placeholder="Selecciona una empresa")
with c2:
    plataforma = st.selectbox("Plataforma", PLATAFORMAS, index=None, placeholder="Selecciona una plataforma")

# 3) Solicitante / Motivo / Tráfico
solicitante = st.text_input("Solicitante")
motivo_solicitud = st.text_area("Motivo de la solicitud")
numero_trafico = st.text_input("Número de tráfico")

st.divider()

def bloque_concepto(prefix: str, titulo: str):
    st.subheader(titulo)

    col1, col2 = st.columns(2)

    # Tipo Concepto
    with col1:
        tipo = st.selectbox(
            "Tipo Concepto",
            TIPOS_CONCEPTO,
            key=f"{prefix}_tipo",
            index=None,
            placeholder="Selecciona un tipo",
        )

    # Concepto dependiente
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

actual = bloque_concepto("actual", "Datos actuales (como están)")
st.divider()
nuevo = bloque_concepto("nuevo", "Datos correctos (como deben quedar)")

st.divider()
registrar = st.button("Registrar")

if not registrar:
    st.stop()

# =========================
# VALIDACIONES
# =========================
errores = []
if not empresa: errores.append("Debes seleccionar una empresa.")
if not plataforma: errores.append("Debes seleccionar una plataforma.")
if not solicitante.strip(): errores.append("El campo 'Solicitante' es obligatorio.")
if not motivo_solicitud.strip(): errores.append("El campo 'Motivo de la solicitud' es obligatorio.")
if not numero_trafico.strip(): errores.append("El campo 'Número de tráfico' es obligatorio.")

for label, block in [("actual", actual), ("correcto", nuevo)]:
    if not block["tipo"]:
        errores.append(f"Debes seleccionar 'Tipo Concepto' ({label}).")
    # Concepto puede estar deshabilitado si no hay conceptos para ese tipo (como GRUA).
    # Si quieres obligar a concepto SOLO cuando hay catálogo:
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

# =========================
# INSERT EN BD
# =========================
fecha_captura = datetime.now(timezone.utc).isoformat()

data_insert = {
    "folio": "PENDIENTE",
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
    row_id = res.data[0]["id"]
except Exception as e:
    st.error(f"Error al guardar en la base de datos: {e}")
    st.stop()

folio_def = folio_visible_from_id(row_id)
try:
    supabase.table("solicitudes_complementarias").update({"folio": folio_def}).eq("id", row_id).execute()
except Exception as e:
    st.warning(f"Se guardó, pero no se pudo actualizar el folio: {e}")

st.success(f"Registro guardado correctamente. Folio: {folio_def}")
