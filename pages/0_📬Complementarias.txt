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
def _title_case(s: str) -> str:
    # "CONSOLIDADO QUERETARO" -> "Consolidado Queretaro"
    return " ".join(w.capitalize() for w in str(s).strip().split())

SUCURSALES_POR_EMPRESA = {
    "Lincoln": [],  # no tiene
    "Picus": ["Carrier Rl", "Carrier RC", "Plus NLD", "Plus QRO", "Logistica"],
    "Igloo": ["Carrier", "Plus", "Logistica"],
    "Set Freight": [
        _title_case(x) for x in [
            "CARGAR", "CHICAGO", "CONSOLIDADO", "CONSOLIDADO QUERETARO", "DALLAS",
            "GUADALAJARA", "LEON", "LINCOLN LOGISTICS", "LUIS MONCAYO", "MG HAULERS",
            "MONTERREY", "NUEVO LAREDO", "QUERETARO", "RAMOS ARIZPE", "ROLANDO ALFARO",
            "SLP LOGISTICS"
        ]
    ],
    "Set Logis Plus": [
        _title_case(x) for x in [
            "BASICOS GJ", "AG FLEET (ROLANDO)", "AURA TRANSPORT",
            "JOEDAN TRANSPORT", "EFRAIN GARCIA"
        ]
    ],
}
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
st.header("📬 Complementarias")

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

    c1, c2, c3 = st.columns(3)
    with c1:
        empresa = st.selectbox("Empresa", EMPRESAS, index=None, placeholder="Selecciona una empresa")

    sucursales_opciones = SUCURSALES_POR_EMPRESA.get(empresa, []) if empresa else []
    with c2:
        if not empresa:
            sucursal = st.selectbox(
                "Sucursal",
                ["Selecciona primero una empresa"],
                index=0,
                disabled=True,
            )
            sucursal = None
        elif len(sucursales_opciones) == 0:
            # Lincoln u otra sin sucursales
            sucursal = "N/A"
            st.text_input("Sucursal", value="N/A", disabled=True)
        else:
            sucursal = st.selectbox(
                "Sucursal",
                sucursales_opciones,
                index=None,
                placeholder="Selecciona una sucursal",
            )

    plataformas_opciones = PLATAFORMAS_POR_EMPRESA.get(empresa, [])
    with c3:
        plataforma = st.selectbox(
            "Plataforma",
            plataformas_opciones,
            index=None,
            placeholder="Selecciona una plataforma" if empresa else "Selecciona primero una empresa",
            disabled=(empresa is None),
        )

    solicitante = st.text_input("Solicitante")
    motivo_solicitud = st.text_area("Motivo de la solicitud")

    c1, c2 = st.columns(2)
    with c1:
        numero_trafico = st.text_input("Número de tráfico")
    with c2:
        tipo_complementaria = st.radio(
            "Tipo de complementaria",
            ["Modificación", "Agregar Concepto"],
            horizontal=True,
        )

    st.divider()

    # --- Bloques (UI) ---
    if tipo_complementaria == "Modificación":
        actual = bloque_concepto("actual", "📝Datos actuales (como están)")
        st.divider()
    else:
        # Para "Agregar Concepto", no aplica capturar los datos actuales.
        # Se guardarán como N/A
        actual = {
            "tipo": "N/A",
            "concepto": "N/A",
            "proveedor": "N/A",
            "moneda": "N/A",
            "importe": None,  # numérico -> guardamos NULL
        }

    nuevo = bloque_concepto("nuevo", "✅Datos correctos (como deben quedar)")

    st.divider()
    registrar = st.button("Registrar", type="primary")

    if registrar:
        errores = []

        # --- Validaciones generales ---
        if not empresa:
            errores.append("Debes seleccionar una empresa.")
        if not plataforma:
            errores.append("Debes seleccionar una plataforma.")
        if not solicitante.strip():
            errores.append("El campo 'Solicitante' es obligatorio.")
        if not motivo_solicitud.strip():
            errores.append("El campo 'Motivo de la solicitud' es obligatorio.")
        if not numero_trafico.strip():
            errores.append("El campo 'Número de tráfico' es obligatorio.")
        if empresa and len(SUCURSALES_POR_EMPRESA.get(empresa, [])) > 0:
            if not sucursal or str(sucursal).strip() == "":
                errores.append("Debes seleccionar una sucursal.")


        # --- Validación de bloques ---
        blocks_to_validate = [("correcto", nuevo)]
        if tipo_complementaria == "Modificación":
            blocks_to_validate.insert(0, ("actual", actual))

        for label, block in blocks_to_validate:
            if not block["tipo"]:
                errores.append(f"Debes seleccionar 'Tipo Concepto' ({label}).")

            if block["tipo"] and len(get_conceptos(block["tipo"])) > 0 and (
                not block["concepto"] or "Sin datos" in str(block["concepto"])
            ):
                errores.append(f"Debes seleccionar 'Concepto' ({label}).")

            if not str(block["proveedor"]).strip():
                errores.append(f"Debes capturar 'Proveedor' ({label}).")

            if not block["moneda"]:
                errores.append(f"Debes seleccionar 'Moneda' ({label}).")

            # El importe normalmente viene como número, pero validamos por seguridad
            if block.get("importe", None) in [None, ""]:
                errores.append(f"Debes capturar 'Importe' ({label}).")

        if errores:
            for e in errores:
                st.error(e)
            st.stop()

        # =========================
        # ALERTA: Tráfico ya registrado (puede haber varias complementarias)
        # =========================
        numero_trafico_clean = numero_trafico.strip().upper()

        try:
            dup = (
                supabase.table("solicitudes_complementarias")
                .select("folio, fecha_captura, solicitante, numero_trafico, empresa, sucursal, estatus")
                .eq("numero_trafico", numero_trafico_clean)
                .order("folio", desc=True)
                .limit(10)
                .execute()
            )
            dup_rows = dup.data or []
        except Exception as e:
            dup_rows = []
            st.warning(f"No se pudo validar tráfico duplicado: {e}")

        if dup_rows:
            lineas = [
                "⚠️ **Este tráfico ya fue registrado previamente**",
                f"**Esta sería la complementaria #{len(dup_rows)+1} para este tráfico.**",
                f"**Tráfico:** {numero_trafico_clean}",
                f"**Complementarias previas encontradas:** {len(dup_rows)}",
                "",
                "**Detalle (más reciente primero):**",
            ]

            for d in dup_rows[:5]:
                folio_prev = f"{int(d.get('folio', 0)):04d}"
                lineas.append(
                    f"- **#{folio_prev}** | {d.get('empresa','')} | {d.get('sucursal','')} | {d.get('estatus','')} | "
                    f"{d.get('solicitante','')} | {d.get('fecha_captura','')}"
                )

            if len(dup_rows) > 5:
                lineas.append(f"- ... y **{len(dup_rows) - 5}** más")

            st.error("\n".join(lineas))



        # --- Insert ---
        fecha_captura = datetime.now(timezone.utc).isoformat()

        data_insert = {
            "fecha_captura": fecha_captura,
            "estatus": "Pendiente",

            "empresa": empresa,
            "sucursal": sucursal if sucursal else "N/A",
            "plataforma": plataforma,
            "solicitante": solicitante.strip(),
            "motivo_solicitud": motivo_solicitud.strip(),
            "tipo_complementaria": tipo_complementaria,
            "numero_trafico": numero_trafico_clean,

            # ACTUAL
            "tipo_concepto_actual": actual["tipo"],
            "concepto_actual": None if "Sin datos" in str(actual["concepto"]) else actual["concepto"],
            "proveedor_actual": str(actual["proveedor"]).strip(),
            "moneda_actual": actual["moneda"],
            "importe_actual": None if actual["importe"] is None else float(actual["importe"]),

            # NUEVO
            "tipo_concepto_nuevo": nuevo["tipo"],
            "concepto_nuevo": None if "Sin datos" in str(nuevo["concepto"]) else nuevo["concepto"],
            "proveedor_nuevo": str(nuevo["proveedor"]).strip(),
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
    import pandas as pd

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
    ESTATUS_ABIERTOS = ["Pendiente", "En revisión"]
    ESTATUS_CERRADOS = ["Resuelto", "Cancelado"]

    AUDITORES = ["Abel Chontal", "Sasha Raz", "Adrian Texna", "Heidi Rodriguez"]

    # =========================
    # Helpers
    # =========================
    def _is_na(val) -> bool:
        if val is None:
            return True
        s = str(val).strip()
        return s == "" or s.upper() == "N/A" or s.lower() == "nan"

    def show_field(label: str, value):
        """Muestra solo si no es N/A / vacío."""
        if not _is_na(value):
            st.write(f"**{label}:** {value}")

    def show_money(label: str, value):
        if value is None:
            return
        try:
            v = float(value)
        except Exception:
            return
        st.write(f"**{label}:** {v:,.2f}")

    # =========================
    # Filtros generales (aplican a la carga base)
    # =========================
    colf1, colf2, colf3, colf4 = st.columns(4)

    with colf1:
        f_empresa = st.selectbox("Empresa", ["(Todas)"] + EMPRESAS, index=0)

    with colf2:
        # opciones de sucursal dependientes de la empresa elegida
        suc_opts = ["(Todas)"]
        if f_empresa != "(Todas)":
            suc_list = SUCURSALES_POR_EMPRESA.get(f_empresa, [])
            suc_opts += (suc_list if len(suc_list) > 0 else ["N/A"])
        else:
            # si es todas, permitir filtrar por N/A (Lincoln) o dejar en (Todas)
            suc_opts += ["N/A"]

        f_sucursal = st.selectbox("Sucursal", suc_opts, index=0)

    with colf3:
        f_estatus = st.selectbox("Estatus", ["(Todos)"] + ESTATUS_OPCIONES, index=0)

    with colf4:
        texto = st.text_input("Buscar (folio / solicitante / tráfico)")

    # Traemos TODA la data para poder mostrar “todo lo capturado”
    q = supabase.table("solicitudes_complementarias").select("*").order("folio", desc=True).limit(500)

    if f_empresa != "(Todas)":
        q = q.eq("empresa", f_empresa)
    if f_sucursal != "(Todas)":
        q = q.eq("sucursal", f_sucursal)
    if f_estatus != "(Todos)":
        q = q.eq("estatus", f_estatus)

    try:
        res = q.execute()
        rows = res.data or []
    except Exception as e:
        st.error(f"No se pudieron cargar solicitudes: {e}")
        st.stop()

    # Búsqueda local
    if texto.strip():
        t = texto.strip().lower()

        def match(r):
            return (
                t in f"{int(r.get('folio', 0)):04d}".lower()
                or t in str(r.get("solicitante", "")).lower()
                or t in str(r.get("numero_trafico", "")).lower()
            )

        rows = [r for r in rows if match(r)]

    # Separación abiertos / cerrados
    abiertos = [r for r in rows if (r.get("estatus") in ESTATUS_ABIERTOS)]
    cerrados = [r for r in rows if (r.get("estatus") in ESTATUS_CERRADOS)]

    # TOTAL SOLO ABIERTOS
    st.write(f"**Total (Pendiente / En revisión): {len(abiertos)}**")

    # =========================
    # LISTA ABIERTOS (expanders)
    # =========================
    if not abiertos:
        st.info("No hay solicitudes pendientes o en revisión con los filtros actuales.")
    else:
        for r in abiertos:
            folio_num = int(r["folio"])
            folio_fmt = f"{folio_num:04d}"
            estatus_actual = r.get("estatus") or "Pendiente"
            tipo_comp = r.get("tipo_complementaria")

            with st.expander(f"Folio #{folio_fmt} | {r.get('empresa')} | {estatus_actual}"):
                # --- Datos generales (mostrar todo lo que no sea N/A) ---
                show_field("Fecha captura", r.get("fecha_captura"))
                show_field("Última modificación", r.get("fecha_ultima_modificacion"))
                show_field("Empresa", r.get("empresa"))
                show_field("Sucursal", r.get("sucursal"))
                show_field("Plataforma", r.get("plataforma"))
                show_field("Solicitante", r.get("solicitante"))
                show_field("Tráfico", r.get("numero_trafico"))
                show_field("Motivo", r.get("motivo_solicitud"))
                show_field("Tipo complementaria", tipo_comp)

                # --- Datos actuales (solo si NO es agregar concepto o si existen valores reales) ---
                # (si es "Agregar Concepto", vienen N/A y no se mostrarán)
                st.divider()
                st.markdown("### Datos actuales (como están)")
                show_field("Tipo concepto (actual)", r.get("tipo_concepto_actual"))
                show_field("Concepto (actual)", r.get("concepto_actual"))
                show_field("Proveedor (actual)", r.get("proveedor_actual"))
                show_field("Moneda (actual)", r.get("moneda_actual"))
                show_money("Importe (actual)", r.get("importe_actual"))

                st.divider()
                st.markdown("### Datos correctos (como deben quedar)")
                show_field("Tipo concepto (nuevo)", r.get("tipo_concepto_nuevo"))
                show_field("Concepto (nuevo)", r.get("concepto_nuevo"))
                show_field("Proveedor (nuevo)", r.get("proveedor_nuevo"))
                show_field("Moneda (nuevo)", r.get("moneda_nuevo"))
                show_money("Importe (nuevo)", r.get("importe_nuevo"))

                # --- Sección de actualización ---
                st.divider()
                c1, c2 = st.columns([2, 1])

                with c1:
                    # Estatus
                    nuevo_estatus = st.selectbox(
                        "Cambiar estatus",
                        ESTATUS_OPCIONES,
                        index=ESTATUS_OPCIONES.index(estatus_actual) if estatus_actual in ESTATUS_OPCIONES else 0,
                        key=f"estatus_{folio_num}",
                    )

                    # Auditor
                    auditor_actual = r.get("auditor")
                    idx_aud = AUDITORES.index(auditor_actual) if auditor_actual in AUDITORES else 0
                    auditor_sel = st.selectbox(
                        "Auditor que actualiza",
                        AUDITORES,
                        index=idx_aud,
                        key=f"auditor_{folio_num}",
                    )

                    # Comentarios
                    comentarios_prev = r.get("comentarios_auditor") or ""
                    comentarios = st.text_area(
                        "Comentarios del auditor",
                        value=comentarios_prev,
                        height=120,
                        key=f"coment_{folio_num}",
                    )

                with c2:
                    if st.button("Guardar cambios", key=f"btn_guardar_{folio_num}"):
                        now_iso = datetime.now(timezone.utc).isoformat()

                        update_payload = {
                            "estatus": nuevo_estatus,
                            "auditor": auditor_sel,
                            "comentarios_auditor": comentarios.strip(),
                            "fecha_ultima_modificacion": now_iso,
                        }

                        # Si se resuelve o cancela, guardamos fecha_resuelto
                        if nuevo_estatus in ["Resuelto", "Cancelado"]:
                            update_payload["fecha_resuelto"] = now_iso
                        else:
                            update_payload["fecha_resuelto"] = None

                        try:
                            supabase.table("solicitudes_complementarias") \
                                .update(update_payload) \
                                .eq("folio", folio_num) \
                                .execute()

                            st.success(f"Actualizado folio #{folio_fmt}")
                            st.rerun()
                        except Exception as e:
                            st.error(f"No se pudo actualizar: {e}")

    # =========================
    # TABLA CERRADOS + EXPORTS
    # =========================
    st.divider()
    st.subheader("Histórico (Resueltos / Cancelados)")

    # Filtros para cerrados
    colh1, colh2, colh3, colh4, colh5 = st.columns(5)
    with colh1:
        h_empresa = st.selectbox("Empresa (histórico)", ["(Todas)"] + EMPRESAS, index=0, key="h_emp")
    with colh2:
        # opciones de sucursal dependientes de la empresa elegida
        suc_opts = ["(Todas)"]
        if h_empresa != "(Todas)":
            suc_opts += (SUCURSALES_POR_EMPRESA.get(h_empresa, []) or ["N/A"])
        else:
            # si es todas, dejamos (Todas) y filtramos por texto opcional
            suc_opts += ["N/A"]
        h_sucursal = st.selectbox("Sucursal (histórico)", suc_opts, index=0, key="h_suc")
    with colh3:
        h_estatus = st.selectbox("Estatus (histórico)", ["(Todos)"] + ESTATUS_CERRADOS, index=0, key="h_est")
    with colh4:
        h_solic = st.text_input("Solicitante (contiene)", key="h_sol")
    with colh5:
        h_rango = st.date_input(
            "Rango de fecha (captura)",
            value=(datetime.now().date().replace(day=1), datetime.now().date()),
            key="h_fecha",
        )

    cerr = cerrados.copy()

    if h_empresa != "(Todas)":
        cerr = [r for r in cerr if r.get("empresa") == h_empresa]
    if h_sucursal != "(Todas)":
        cerr = [r for r in cerr if (r.get("sucursal") or "N/A") == h_sucursal]
    if h_estatus != "(Todos)":
        cerr = [r for r in cerr if r.get("estatus") == h_estatus]
    if h_solic.strip():
        s = h_solic.strip().lower()
        cerr = [r for r in cerr if s in str(r.get("solicitante", "")).lower()]

    # filtro por rango fecha usando fecha_captura (ISO)
    if isinstance(h_rango, tuple) and len(h_rango) == 2 and h_rango[0] and h_rango[1]:
        d1, d2 = h_rango
        def in_range(r):
            fc = r.get("fecha_captura")
            if not fc:
                return False
            try:
                dt = pd.to_datetime(fc, utc=True).date()
            except Exception:
                return False
            return d1 <= dt <= d2
        cerr = [r for r in cerr if in_range(r)]

    if not cerr:
        st.info("No hay registros en el histórico con esos filtros.")
    else:
        # Tabla resumida (pero exporta todo)
        df_cerr = pd.DataFrame(cerr)

        # Orden de columnas sugerido (si existen)
        cols_pref = [
            "folio", "estatus", "empresa", "sucursal", "plataforma", "solicitante", "numero_trafico",
            "tipo_complementaria", "fecha_captura", "fecha_ultima_modificacion", "fecha_resuelto",
            "auditor", "comentarios_auditor",
            "tipo_concepto_actual", "concepto_actual", "proveedor_actual", "moneda_actual", "importe_actual",
            "tipo_concepto_nuevo", "concepto_nuevo", "proveedor_nuevo", "moneda_nuevo", "importe_nuevo",
            "motivo_solicitud",
        ]
        cols = [c for c in cols_pref if c in df_cerr.columns] + [c for c in df_cerr.columns if c not in cols_pref]
        df_cerr = df_cerr[cols]

        st.dataframe(df_cerr, width="stretch")

        # Exports (cerrados filtrados)
        csv_bytes = df_cerr.to_csv(index=False).encode("utf-8")

        cexp1, cexp2, cexp3 = st.columns(3)
        with cexp1:
            st.download_button(
                "Descargar CSV (filtrado)",
                data=csv_bytes,
                file_name="historico_filtrado.csv",
                mime="text/csv",
            )
        with cexp2:
            try:
                import io
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    df_cerr.to_excel(writer, index=False, sheet_name="Historico")
                st.download_button(
                    "Descargar Excel (filtrado)",
                    data=buf.getvalue(),
                    file_name="historico_filtrado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.warning(f"No se pudo generar Excel: {e}")

        # Descargar TODA la data (según carga base rows, sin filtros de histórico)
        with cexp3:
            df_all = pd.DataFrame(rows) if rows else pd.DataFrame()
            if not df_all.empty:
                all_csv = df_all.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "Descargar TODA la data (CSV)",
                    data=all_csv,
                    file_name="toda_la_data.csv",
                    mime="text/csv",
                )
            else:
                st.download_button(
                    "Descargar TODA la data (CSV)",
                    data="".encode("utf-8"),
                    file_name="toda_la_data.csv",
                    mime="text/csv",
                    disabled=True,
                )
