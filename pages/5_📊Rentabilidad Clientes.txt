import streamlit as st
import pandas as pd
import numpy as np
import json
from io import BytesIO
from supabase import create_client

# ==============================
# CONFIGURACIÓN SUPABASE
# ==============================
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
supabase = create_client(url, key)

st.set_page_config(page_title="Distribución CI Clientes", layout="wide")
st.title("👥 Prorrateo de Costos Indirectos por Cliente")

st.markdown(
    """
Esta app te ayuda a:
1. Calcular, a partir de la **DATA detallada**, la tabla de viajes por cliente/año/mes
   (viajes totales, con remolque, con unidad, millas/kilómetros, tipo de cliente),
   excluyendo ciertos operadores logísticos (solo cuando aplique).
2. Repartir costos **no ligados a la operación** entre viajes con/sin unidad
   (y obtener costos unitarios).
3. Definir un catálogo de distribución para costos **ligados a la operación**.
4. Prorratear esos costos entre clientes.
5. **Asignar los costos indirectos (CI) a nivel viaje**.
"""
)

# ============================================================
# 0️⃣ SELECCIÓN DE EMPRESA + CONFIGURACIÓN
# ============================================================
st.subheader("🏢 Empresa")

EMPRESAS = ["Lincoln Freight", "Set Logis Plus", "Picus Carrier", "Igloo Carrier"]
empresa = st.selectbox("Selecciona la empresa", EMPRESAS, index=0)

def get_empresa_config(nombre_empresa: str):
    """
    Config por empresa:
    - columnas base
    - si aplica filtro/validación de operador logístico
    - prefijo de remolque
    - etiqueta de distancia (Millas/Kilómetros)
    """
    if nombre_empresa == "Lincoln Freight":
        return {
            "candidates_fecha": ["Bill date", "Bill Date", "Fecha", "FECHA", "Date"],
            "candidates_customer": ["Customer", "Cliente", "cliente"],
            "candidates_trip": ["Trip Number", "Trip number", "TripNumber", "Viaje", "No Viaje", "Folio"],
            "candidates_trailer": ["Trailer", "Remolque", "remolque"],
            "candidates_unit": ["Unit", "Unidad", "unidad"],
            "candidates_dist": ["Real Miles", "Real miles", "REAL MILES", "Miles reales", "Real_miles", "Real Mi"],
            "candidates_operador": ["Logistic Operator", "Operador logistico", "Operador logístico"],
            "usa_operador": True,          # ✅ Lincoln sí aplica
            "trailer_prefix": "LF",
            "dist_label": "Millas",
        }

    if nombre_empresa == "Set Logis Plus":
        return {
            "candidates_fecha": ["Bill date", "Bill Date", "Fecha", "FECHA", "Date"],
            "candidates_customer": ["Customer", "Cliente", "cliente"],
            "candidates_trip": ["Trip Number", "Trip number", "TripNumber", "Viaje", "No Viaje", "Folio"],
            "candidates_trailer": ["Trailer", "Remolque", "remolque"],
            "candidates_unit": ["Unit", "Unidad", "unidad"],
            "candidates_dist": ["Real Miles", "Real miles", "REAL MILES", "Miles reales", "Real_miles", "Real Mi"],
            "candidates_operador": ["Logistic Operator", "Operador logistico", "Operador logístico"],
            "usa_operador": False,         # ✅ Set: NO se hace comprobación por ahora
            "trailer_prefix": "STL",
            "dist_label": "Millas",
        }

    if nombre_empresa == "Picus Carrier":
        return {
            "candidates_fecha": ["Fecha", "FECHA", "Bill date", "Bill Date", "Date"],
            "candidates_customer": ["Cliente", "cliente", "Customer"],
            "candidates_trip": ["Trip Number", "Trip number", "TripNumber", "Viaje", "No Viaje", "Folio", "Folio Viaje"],
            "candidates_trailer": ["Remolque", "remolque", "Trailer"],
            "candidates_unit": ["Unidad", "unidad", "Unit"],
            "candidates_dist": ["KMS Ruta", "Kms Ruta", "KMS_Ruta", "Kilometros", "Kilómetros", "KM", "KMS"],
            "candidates_operador": [],
            "usa_operador": False,
            "trailer_prefix": "PI",        # ✅ Picus: PI
            "dist_label": "Kilómetros",
        }

    # Igloo Carrier
    return {
        "candidates_fecha": ["Fecha", "FECHA", "Bill date", "Bill Date", "Date"],
        "candidates_customer": ["Cliente", "cliente", "Customer"],
        "candidates_trip": ["Trip Number", "Trip number", "TripNumber", "Viaje", "No Viaje", "Folio", "Folio Viaje"],
        "candidates_trailer": ["Remolque", "remolque", "Trailer"],
        "candidates_unit": ["Unidad", "unidad", "Unit"],
        "candidates_dist": ["KMS Ruta", "Kms Ruta", "KMS_Ruta", "Kilometros", "Kilómetros", "KM", "KMS"],
        "candidates_operador": [],
        "usa_operador": False,
        "trailer_prefix": "IGT",          # ✅ Igloo: IGT
        "dist_label": "Kilómetros",
    }

CFG = get_empresa_config(empresa)
DIST_COL_NAME = CFG["dist_label"]  # "Millas" o "Kilómetros" (se usa en tablas/UI)

# ============================================================
# Auxiliares
# ============================================================
def normaliza_tipo_distribucion(valor):
    if isinstance(valor, list):
        return valor[0] if valor else None
    if isinstance(valor, str):
        v = valor.strip()
        if not v:
            return None
        try:
            parsed = json.loads(v)
            if isinstance(parsed, list) and parsed:
                return str(parsed[0])
            if isinstance(parsed, str):
                return parsed
        except Exception:
            pass
        v = v.strip("[]{}").strip().strip('"\'')

        return v or None
    return valor

def find_column(df, candidates):
    norm_map = {str(c).lower().replace(" ", "").replace("_", ""): c for c in df.columns}
    for cand in candidates:
        key = cand.lower().replace(" ", "").replace("_", "")
        if key in norm_map:
            return norm_map[key]
    return None

OPERADORES_EXCLUIR = {
    "ERICK LARA",
    "JULIETA REYNA",
    "GLADYS GUTIERREZ",
    "JUAN EDUARDO VILLARREAL VALDEZ",
    "LUIS ALDO VELIZ DE LEON",
    "VICTOR CHAVEZ SILVA",
    "ANETTE ROJO",
    "LUIS EDUARDO GUTIERREZ RAMIREZ",
    "GABRIEL ACOSTA VITAL",
    "GRISELDA JIMENEZ",
}

def build_flag_trailer(series_trailer: pd.Series, prefix: str) -> pd.Series:
    s = series_trailer.astype(str).str.strip().str.upper()
    return s.str.startswith(prefix).astype(int)

# ============================================================
# 1️⃣ CARGA DATA DETALLADA Y TABLA POR CLIENTE
# ============================================================
st.header("1️⃣ Cargar DATA detallada de viajes y generar tabla por cliente")

file_data = st.file_uploader(
    "Sube la DATA detallada de viajes",
    type=["xlsx"],
    key="data_file",
)

if file_data:
    try:
        xls = pd.ExcelFile(file_data)
        hoja_data = st.selectbox("Selecciona la hoja con la DATA detallada", xls.sheet_names)
        df_data = pd.read_excel(xls, sheet_name=hoja_data)
        df_data.columns = df_data.columns.astype(str)

        st.subheader("Vista previa DATA de viajes")
        st.dataframe(df_data.head(), width="stretch")

        col_fecha = find_column(df_data, CFG["candidates_fecha"])
        col_customer = find_column(df_data, CFG["candidates_customer"])
        col_trip = find_column(df_data, CFG["candidates_trip"])
        col_trailer = find_column(df_data, CFG["candidates_trailer"])
        col_unit = find_column(df_data, CFG["candidates_unit"])
        col_dist = find_column(df_data, CFG["candidates_dist"])

        col_logop = None
        if CFG["usa_operador"]:
            col_logop = find_column(df_data, CFG["candidates_operador"])

        faltan_cols = []
        if col_fecha is None: faltan_cols.append("Fecha")
        if col_customer is None: faltan_cols.append("Customer/Cliente")
        if col_trip is None: faltan_cols.append("Trip Number/Viaje")
        if col_trailer is None: faltan_cols.append("Trailer/Remolque")
        if col_unit is None: faltan_cols.append("Unit/Unidad")
        if col_dist is None: faltan_cols.append(DIST_COL_NAME)
        if CFG["usa_operador"] and col_logop is None: faltan_cols.append("Logistic Operator")

        if faltan_cols:
            st.error("No se encontraron las siguientes columnas necesarias en la DATA: " + ", ".join(faltan_cols))
            st.stop()

        # Normalización base
        df_data = df_data.copy()
        df_data[col_fecha] = pd.to_datetime(df_data[col_fecha], errors="coerce")
        df_data["Año"] = df_data[col_fecha].dt.year
        df_data["Mes"] = df_data[col_fecha].dt.month

        df_data["_trip"] = df_data[col_trip].astype(str).str.strip()
        df_data["_flag_trailer"] = build_flag_trailer(df_data[col_trailer], CFG["trailer_prefix"])
        df_data["_flag_unidad"] = (
            df_data[col_unit].notna() & (df_data[col_unit].astype(str).str.strip() != "")
        ).astype(int)

        df_data["_dist"] = pd.to_numeric(df_data[col_dist], errors="coerce").fillna(0.0)
        df_data["_dist_con_unidad"] = df_data["_dist"] * df_data["_flag_unidad"]

        # Guardar para paso 5
        st.session_state["df_data_original"] = df_data.copy()
        st.session_state["empresa_sel"] = empresa
        st.session_state["cfg_empresa"] = CFG
        st.session_state["dist_col_name"] = DIST_COL_NAME

        # ----------------------------------
        # Filtro de operadores SOLO Lincoln (por configuración)
        # ----------------------------------
        if CFG["usa_operador"]:
            logop_upper = df_data[col_logop].astype(str).str.upper().str.strip()
            mask_log_ok = (logop_upper == "") | (~logop_upper.isin(OPERADORES_EXCLUIR))
            df_filt = df_data[mask_log_ok].copy()
            st.write(
                f"Registros usados para drivers (excluyendo operadores lista): "
                f"{df_filt.shape[0]} de {df_data.shape[0]}."
            )
        else:
            df_filt = df_data.copy()
            st.info("No se aplica filtro/comprobación de operadores para esta empresa.")

        # Selección año/mes (sobre la DATA completa)
        col1, col2 = st.columns(2)
        with col1:
            anios = sorted(df_data["Año"].dropna().unique())
            anio_sel = st.selectbox("Año", anios)
        with col2:
            meses = sorted(df_data.loc[df_data["Año"] == anio_sel, "Mes"].dropna().unique())
            mes_sel = st.selectbox("Mes (número)", meses)

        st.session_state["anio_sel"] = int(anio_sel)
        st.session_state["mes_sel"] = int(mes_sel)

        df_data_mes = df_data[(df_data["Año"] == anio_sel) & (df_data["Mes"] == mes_sel)].copy()
        df_mes_filt = df_filt[(df_filt["Año"] == anio_sel) & (df_filt["Mes"] == mes_sel)].copy()

        if df_data_mes.empty:
            st.warning("No hay registros para ese año/mes en la DATA.")
            st.stop()

        # Totales base CI (todos los viajes del mes)
        total_viajes_ci = df_data_mes["_trip"].nunique()
        total_con_unidad_ci = int(df_data_mes["_flag_unidad"].sum())
        total_sin_unidad_ci = total_viajes_ci - total_con_unidad_ci
        total_con_remolque_ci = int(df_data_mes["_flag_trailer"].sum())
        dist_mes_ci = float(df_data_mes.loc[df_data_mes["_flag_unidad"] == 1, "_dist"].sum())

        pct_con_unidad_ci = total_con_unidad_ci / total_viajes_ci if total_viajes_ci else 0.0
        pct_sin_unidad_ci = 1.0 - pct_con_unidad_ci if total_viajes_ci else 0.0

        resumen_ci = pd.DataFrame(
            {
                "Métrica": [
                    "Viajes totales (base CI)",
                    "Viajes con unidad (base CI)",
                    "Viajes sin unidad (base CI)",
                    f"Viajes con remolque ({CFG['trailer_prefix']}) (base CI)",
                    f"{DIST_COL_NAME} con unidad (base CI)",
                ],
                "Valor": [
                    total_viajes_ci,
                    total_con_unidad_ci,
                    total_sin_unidad_ci,
                    total_con_remolque_ci,
                    dist_mes_ci,
                ],
            }
        )

        st.subheader(f"Totales del mes {anio_sel}-{mes_sel:02d} para CI (TODOS los viajes del mes)")
        st.dataframe(resumen_ci, width="stretch")

        st.write(
            f"**% viajes con unidad (base CI):** {pct_con_unidad_ci:.4%}   |   "
            f"**% viajes sin unidad (base CI):** {pct_sin_unidad_ci:.4%}"
        )

        # Guardar para paso 2
        st.session_state["pct_con_unidad"] = pct_con_unidad_ci
        st.session_state["pct_sin_unidad"] = pct_sin_unidad_ci
        st.session_state["dist_mes"] = dist_mes_ci
        st.session_state["total_sin_unidad"] = total_sin_unidad_ci

        # Tabla por cliente (drivers)
        if df_mes_filt.empty:
            st.warning("Después del filtro (si aplica), no quedan registros para tabla por cliente.")
        else:
            tabla_mes = (
                df_mes_filt.groupby(["Año", "Mes", col_customer], as_index=False)
                .agg(
                    Viajes=("_trip", "nunique"),
                    **{
                        "Viajes con remolques": ("_flag_trailer", "sum"),
                        "Viajes con unidad": ("_flag_unidad", "sum"),
                        DIST_COL_NAME: ("_dist_con_unidad", "sum"),
                    },
                )
            )

            tabla_mes = tabla_mes.rename(columns={col_customer: "Customer"})
            tabla_mes["Viajes sin unidad"] = tabla_mes["Viajes"] - tabla_mes["Viajes con unidad"]

            tabla_mes["pct_equipo"] = np.where(
                tabla_mes["Viajes"] > 0,
                tabla_mes["Viajes con remolques"] / tabla_mes["Viajes"],
                0.0,
            )
            tabla_mes["pct_unidad"] = np.where(
                tabla_mes["Viajes"] > 0,
                tabla_mes["Viajes con unidad"] / tabla_mes["Viajes"],
                0.0,
            )

            st.subheader(f"Tabla por cliente ({anio_sel}-{mes_sel:02d})")
            st.dataframe(tabla_mes, width="stretch")

            st.session_state["df_mes_clientes"] = tabla_mes

    except Exception as e:
        st.error(f"Error leyendo la DATA: {e}")
        st.stop()

# ============================================================
# 2️⃣ COSTOS NO LIGADOS A LA OPERACIÓN
# ============================================================
st.header("2️⃣ Costos indirectos **no ligados a la operación**")

if "df_mes_clientes" not in st.session_state:
    st.warning("Primero carga la DATA y selecciona año/mes en el paso 1.")
else:
    file_no_op = st.file_uploader(
        "Sube el archivo de costos NO ligados a la operación (Concepto + meses)",
        type=["xlsx"],
        key="no_op_file",
    )

    if file_no_op:
        try:
            xls_no = pd.ExcelFile(file_no_op)
            hoja_sel = st.selectbox("Hoja de costos no operativos", xls_no.sheet_names)
            df_no = pd.read_excel(xls_no, sheet_name=hoja_sel)
            df_no.columns = df_no.columns.astype(str).str.strip()

            st.subheader("Vista previa costos no operativos")
            st.dataframe(df_no.head(), width="stretch")

            columnas_mes = [
                c for c in df_no.columns
                if c not in ["Concepto", "CONCEPTO"] and df_no[c].dtype != "O"
            ] or [c for c in df_no.columns if c not in ["Concepto", "CONCEPTO"]]

            mes_sel_global = st.session_state.get("mes_sel")
            index_default = 0
            if mes_sel_global is not None and len(columnas_mes) > 0:
                mapa_meses = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
                nombre_mes = mapa_meses.get(int(mes_sel_global))
                if nombre_mes in columnas_mes:
                    index_default = columnas_mes.index(nombre_mes)

            col_mes_sel = st.selectbox("Selecciona la columna del mes (ej. 'Ene')", columnas_mes, index=index_default)

            concepto_col = "Concepto" if "Concepto" in df_no.columns else "CONCEPTO"
            df_no["Concepto"] = df_no[concepto_col].astype(str)
            df_no["Monto_mes"] = pd.to_numeric(df_no[col_mes_sel], errors="coerce").fillna(0)

            costo_total_no_op = df_no["Monto_mes"].sum()
            st.write(f"**Costo total no operativo del mes:** ${costo_total_no_op:,.2f}")

            pct_con_unidad = st.session_state["pct_con_unidad"]
            pct_sin_unidad = st.session_state["pct_sin_unidad"]
            dist_mes = st.session_state["dist_mes"]
            total_sin_unidad = st.session_state["total_sin_unidad"]

            bolsa_con_unidad = costo_total_no_op * pct_con_unidad
            bolsa_sin_unidad = costo_total_no_op * pct_sin_unidad

            st.write(f"**Monto asignado a viajes con unidad:** ${bolsa_con_unidad:,.2f} ({pct_con_unidad:.4%})")
            st.write(f"**Monto asignado a viajes sin unidad:** ${bolsa_sin_unidad:,.2f} ({pct_sin_unidad:.4%})")

            costo_x_dist = bolsa_con_unidad / dist_mes if dist_mes else None
            costo_x_viaje_sin = (bolsa_sin_unidad / total_sin_unidad if total_sin_unidad else None)

            st.subheader("Costos unitarios derivados (informativos)")
            if costo_x_dist is not None:
                st.write(f"**Costo por {DIST_COL_NAME.lower()} (viajes con unidad):** ${costo_x_dist:,.6f}")
            else:
                st.write(f"No se pudo calcular costo por {DIST_COL_NAME.lower()} (falta total).")

            if costo_x_viaje_sin is not None:
                st.write(f"**Costo por viaje sin unidad:** ${costo_x_viaje_sin:,.6f}")
            else:
                st.write("No se pudo calcular costo por viaje sin unidad (no hay viajes sin unidad en el mes).")

            st.session_state["costo_no_op_total"] = costo_total_no_op
            st.session_state["costo_no_op_x_dist"] = costo_x_dist
            st.session_state["costo_no_op_x_viaje_sin"] = costo_x_viaje_sin

        except Exception as e:
            st.error(f"Error procesando los costos no operativos: {e}")

# ============================================================
# 3️⃣ CATÁLOGO PARA COSTOS LIGADOS A LA OPERACIÓN (SUPABASE POR EMPRESA)
# ============================================================
st.header("3️⃣ Catálogo de costos ligados a la operación")

tipos_distribucion = [
    "Volumen Viajes",
    "Viajes con Remolque",
    "Viajes con unidad",
    DIST_COL_NAME,  # ✅ Millas o Kilómetros según empresa
]

# Cargar catálogo existente filtrado por empresa
try:
    data_cat = (
        supabase.table("catalogo_costos_clientes")
        .select("*")
        .eq("empresa", empresa)
        .execute()
        .data
    )
    catalogo_existente = pd.DataFrame(data_cat)
except Exception:
    catalogo_existente = pd.DataFrame()

file_op = st.file_uploader(
    "Sube el archivo de costos ligados a operación (Concepto + meses)",
    type=["xlsx"],
    key="op_file",
)

if file_op:
    try:
        xls_op = pd.ExcelFile(file_op)
        hoja_op_sel = st.selectbox("Hoja de costos ligados a operación", xls_op.sheet_names)
        df_op = pd.read_excel(xls_op, sheet_name=hoja_op_sel)
        df_op.columns = df_op.columns.astype(str).str.strip()

        st.subheader("Vista previa costos ligados a operación")
        st.dataframe(df_op.head(), width="stretch")

        columnas_mes_op = [
            c for c in df_op.columns
            if c not in ["Concepto", "CONCEPTO"] and df_op[c].dtype != "O"
        ] or [c for c in df_op.columns if c not in ["Concepto", "CONCEPTO"]]

        mes_sel_global = st.session_state.get("mes_sel")
        index_default_op = 0
        if mes_sel_global is not None and len(columnas_mes_op) > 0:
            mapa_meses = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
            nombre_mes = mapa_meses.get(int(mes_sel_global))
            if nombre_mes in columnas_mes_op:
                index_default_op = columnas_mes_op.index(nombre_mes)

        col_mes_op = st.selectbox("Selecciona la columna del mes a prorratear (ej. 'Ene')", columnas_mes_op, index=index_default_op)

        concepto_col_op = "Concepto" if "Concepto" in df_op.columns else "CONCEPTO"
        df_op["Concepto"] = df_op[concepto_col_op].astype(str).str.strip()
        df_op["Monto_mes"] = pd.to_numeric(df_op[col_mes_op], errors="coerce").fillna(0)

        st.write(f"**Total costos ligados a operación ({col_mes_op}):** ${df_op['Monto_mes'].sum():,.2f}")

        conceptos = df_op[["Concepto"]].drop_duplicates().reset_index(drop=True)

        if not catalogo_existente.empty:
            catalogo_existente = catalogo_existente.rename(columns={"concepto": "Concepto", "tipo_distribucion": "Tipo distribución"})
            catalogo_existente["Tipo distribución"] = catalogo_existente["Tipo distribución"].apply(normaliza_tipo_distribucion)
            merged_cat = conceptos.merge(catalogo_existente, on="Concepto", how="left")
            # ✅ Mostrar SOLO lo que el usuario debe editar
            keep_cols = ["Concepto", "Tipo distribución"]
            for c in list(merged_cat.columns):
                if c not in keep_cols:
                    merged_cat.drop(columns=[c], inplace=True)

        else:
            merged_cat = conceptos.copy()
            merged_cat["Tipo distribución"] = None

        st.subheader("Catálogo de distribución por concepto (por empresa)")
        merged_cat = merged_cat.sort_values(by=["Tipo distribución", "Concepto"], na_position="first").reset_index(drop=True)

        edited_cat = st.data_editor(
            merged_cat,
            width="stretch",
            column_config={
                "Tipo distribución": st.column_config.SelectboxColumn(
                    label="Tipo de distribución",
                    options=tipos_distribucion,
                    required=True,
                )
            },
            key="cat_editor",
        )

        if st.button("💾 Guardar catálogo en Supabase", key="save_cat"):
            try:
                registros = []
                for _, row in edited_cat.iterrows():
                    if pd.notna(row["Tipo distribución"]):
                        concepto = str(row["Concepto"]).strip()
                        tipo = normaliza_tipo_distribucion(row["Tipo distribución"])

                        registros.append({
                            "empresa": empresa,                      # ✅ automático
                            "concepto": concepto,                    # ✅ de la fila
                            "tipo_distribucion": str(tipo).strip(),  # ✅ limpio
                            "empresa,concepto": f"{empresa},{concepto}",  # ✅ automático (tu columna extra)
                        })

                if registros:
                    supabase.table("catalogo_costos_clientes").upsert(
                        registros,
                        on_conflict="empresa,concepto",
                    ).execute()
                st.success("Catálogo actualizado en Supabase (por empresa).")
            except Exception as e:
                st.error(f"Error al guardar el catálogo: {e}")

        st.session_state["df_costos_op_mes"] = df_op[["Concepto", "Monto_mes"]]

    except Exception as e:
        st.error(f"Error procesando los costos ligados a operación: {e}")

# ============================================================
# 4️⃣ PRORRATEO DE COSTOS LIGADOS A OPERACIÓN ENTRE CLIENTES
# ============================================================
st.header("4️⃣ Prorrateo de costos ligados a operación por cliente")

costo_x_dist_info = st.session_state.get("costo_no_op_x_dist")
costo_x_viaje_sin_info = st.session_state.get("costo_no_op_x_viaje_sin")

if ("df_mes_clientes" not in st.session_state) or ("df_costos_op_mes" not in st.session_state):
    st.info("Necesitas completar los pasos 1 y 3 para poder prorratear.")
else:
    df_mes_clientes = st.session_state["df_mes_clientes"].copy()
    df_op_mes = st.session_state["df_costos_op_mes"].copy()

    data_cat = (
        supabase.table("catalogo_costos_clientes")
        .select("*")
        .eq("empresa", empresa)
        .execute()
        .data
    )
    catalogo = pd.DataFrame(data_cat)

    if catalogo.empty:
        st.error("No hay catálogo de distribución en 'catalogo_costos_clientes' para esta empresa.")
    else:
        catalogo = catalogo.rename(columns={"concepto": "Concepto", "tipo_distribucion": "Tipo distribución"})
        catalogo["Tipo distribución"] = catalogo["Tipo distribución"].apply(normaliza_tipo_distribucion)

        df_op_mes = df_op_mes.merge(catalogo[["Concepto", "Tipo distribución"]], on="Concepto", how="left")

        if df_op_mes["Tipo distribución"].isna().any():
            faltan = df_op_mes.loc[df_op_mes["Tipo distribución"].isna(), "Concepto"].unique()
            st.error("Hay conceptos sin tipo de distribución definido en el catálogo: " + ", ".join(faltan[:10]) + ("..." if len(faltan) > 10 else ""))
        else:
            driver_map = {
                "Volumen Viajes": "Viajes",
                "Viajes con Remolque": "Viajes con remolques",
                "Viajes con unidad": "Viajes con unidad",
                DIST_COL_NAME: DIST_COL_NAME,  # ✅ Millas o Kilómetros
            }

            base_clientes = df_mes_clientes.groupby(["Customer"], as_index=False).agg(
                {"Viajes": "sum", "Viajes con remolques": "sum", "Viajes con unidad": "sum", DIST_COL_NAME: "sum"}
            )

            asignaciones = []
            for _, row in df_op_mes.iterrows():
                concepto = row["Concepto"]
                monto = float(row["Monto_mes"])
                tipo_dist = row["Tipo distribución"]

                col_driver = driver_map.get(tipo_dist)
                if col_driver not in base_clientes.columns:
                    st.warning(f"Tipo '{tipo_dist}' requiere columna '{col_driver}', no existe. Se omite {concepto}.")
                    continue

                df_driver = base_clientes[["Customer", col_driver]].copy()
                total_driver = df_driver[col_driver].sum()

                if total_driver == 0:
                    st.warning(f"Driver '{col_driver}' para {concepto} es 0. Se omite.")
                    continue

                df_driver["%driver"] = df_driver[col_driver] / total_driver
                df_driver["Concepto"] = concepto
                df_driver["Tipo distribución"] = tipo_dist
                df_driver["Costo asignado"] = df_driver["%driver"] * monto
                asignaciones.append(df_driver)

            if not asignaciones:
                st.warning("No se pudo asignar ningún costo (revisa drivers y catálogo).")
            else:
                asignaciones_df = pd.concat(asignaciones, ignore_index=True)

                st.subheader("Detalle de asignación por concepto y cliente")
                st.dataframe(asignaciones_df, width="stretch")

                pivot_clientes = (
                    asignaciones_df.pivot_table(
                        index=["Customer"],
                        columns="Concepto",
                        values="Costo asignado",
                        aggfunc="sum",
                        fill_value=0.0,
                    )
                    .reset_index()
                )

                pivot_clientes["Total costos ligados op"] = pivot_clientes.drop(columns=["Customer"]).sum(axis=1)

                st.subheader("Totales por cliente (solo costos ligados a la operación)")
                st.dataframe(pivot_clientes, width="stretch")

                st.session_state["asignaciones_df"] = asignaciones_df
                st.session_state["conceptos_tipos"] = df_op_mes[["Concepto", "Tipo distribución"]]
                st.session_state["pivot_clientes_ci"] = pivot_clientes

                def to_excel_bytes(df1, df2):
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                        df1.to_excel(writer, index=False, sheet_name="Detalle_asignaciones")
                        df2.to_excel(writer, index=False, sheet_name="Totales_por_cliente")
                    return buffer.getvalue()

                st.download_button(
                    "📥 Descargar resultados (Excel clientes)",
                    data=to_excel_bytes(asignaciones_df, pivot_clientes),
                    file_name="prorrateo_costos_clientes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ============================================================
# 5️⃣ ASIGNACIÓN DE CI A NIVEL VIAJE
# ============================================================
st.header("5️⃣ Asignación de CI a nivel viaje")

faltan_requisitos = []
if st.session_state.get("costo_no_op_x_dist") is None:
    faltan_requisitos.append(f"Costo unitario por {DIST_COL_NAME.lower()} (paso 2).")
if st.session_state.get("costo_no_op_x_viaje_sin") is None:
    faltan_requisitos.append("Costo unitario por viaje sin unidad (paso 2).")
if "asignaciones_df" not in st.session_state or "conceptos_tipos" not in st.session_state:
    faltan_requisitos.append("Prorrateo por cliente (paso 4).")

if faltan_requisitos:
    st.info("Para usar este apartado necesitas haber completado:\n- " + "\n- ".join(faltan_requisitos))
else:
    origen_trips = st.radio(
        "¿Qué base quieres usar para asignar CI?",
        ["Usar la misma DATA del paso 1", "Subir otro archivo"],
        index=0,
        key="origen_trips_radio",
    )

    df_trips = None
    file_trips = None

    if origen_trips == "Usar la misma DATA del paso 1":
        if "df_data_original" not in st.session_state:
            st.error("No se encontró la DATA del paso 1 en memoria. Vuelve a cargarla o sube otro archivo.")
        else:
            df_trips = st.session_state["df_data_original"].copy()
    else:
        file_trips = st.file_uploader(
            "Sube la base de viajes a nivel detalle",
            type=["xlsx"],
            key="file_trips_ci",
        )

    if (df_trips is not None) or file_trips:
        try:
            if df_trips is None and file_trips is not None:
                xls_trips = pd.ExcelFile(file_trips)
                hoja_trips = st.selectbox("Hoja con los viajes detallados", xls_trips.sheet_names, key="hoja_trips_sel")
                df_trips = pd.read_excel(xls_trips, sheet_name=hoja_trips)
                df_trips.columns = df_trips.columns.astype(str)

            col_fecha = find_column(df_trips, CFG["candidates_fecha"])
            col_customer = find_column(df_trips, CFG["candidates_customer"])
            col_trip = find_column(df_trips, CFG["candidates_trip"])
            col_unit = find_column(df_trips, CFG["candidates_unit"])
            col_trailer = find_column(df_trips, CFG["candidates_trailer"])
            col_dist = find_column(df_trips, CFG["candidates_dist"])

            col_operador = None
            if CFG["usa_operador"]:
                col_operador = find_column(df_trips, CFG["candidates_operador"])

            # filtro por mes si aplica
            if col_fecha is not None and "anio_sel" in st.session_state and "mes_sel" in st.session_state:
                df_trips[col_fecha] = pd.to_datetime(df_trips[col_fecha], errors="coerce")
                anio_sel = int(st.session_state["anio_sel"])
                mes_sel = int(st.session_state["mes_sel"])
                mask_mes = (df_trips[col_fecha].dt.year == anio_sel) & (df_trips[col_fecha].dt.month == mes_sel)
                df_trips = df_trips[mask_mes].copy()
                st.write(f"Se usarán {df_trips.shape[0]} viajes del mes {mes_sel:02d}/{anio_sel}.")
            else:
                st.warning("No se encontró columna de fecha o no hay año/mes definido. Se usarán todos los viajes.")

            st.subheader("Vista previa viajes (después de filtro por mes)")
            st.dataframe(df_trips.head(), width="stretch")

            # Validaciones mínimas
            columnas_faltan = []
            if col_customer is None: columnas_faltan.append("Customer/Cliente")
            if col_trip is None: columnas_faltan.append("Trip Number/Viaje")
            if col_unit is None: columnas_faltan.append("Unit/Unidad")
            if col_trailer is None: columnas_faltan.append("Trailer/Remolque")
            if col_dist is None: columnas_faltan.append(DIST_COL_NAME)
            if CFG["usa_operador"] and col_operador is None:
                columnas_faltan.append("Operador logístico")

            if columnas_faltan:
                st.error("No se encontraron columnas necesarias en viajes: " + ", ".join(columnas_faltan))
                st.stop()

            df_trips_work = df_trips.copy()

            # Operadores excluidos: SOLO Lincoln
            if CFG["usa_operador"]:
                op_upper = df_trips_work[col_operador].astype(str).str.upper().str.strip()
                mask_excl = op_upper.isin(OPERADORES_EXCLUIR)
                df_trips_work["Excluido_por_operador"] = mask_excl
                st.write(
                    f"Viajes con operador en lista de 'excluidos': {int(mask_excl.sum())} "
                    f"de {len(df_trips_work)} (sí pueden recibir CI si tienen unidad y distancia)."
                )
            else:
                df_trips_work["Excluido_por_operador"] = False

            # CI NO OPERATIVOS a nivel viaje
            costo_x_dist = float(st.session_state["costo_no_op_x_dist"])
            costo_x_viaje_sin = float(st.session_state["costo_no_op_x_viaje_sin"])

            has_unit = df_trips_work[col_unit].notna() & (df_trips_work[col_unit].astype(str).str.strip() != "")
            dist = pd.to_numeric(df_trips_work[col_dist], errors="coerce").fillna(0.0)

            ci_no_op = np.zeros(len(df_trips_work), dtype=float)

            idx_con_unidad = (has_unit & (dist > 0)).to_numpy()
            ci_no_op[idx_con_unidad] = costo_x_dist * dist.to_numpy()[idx_con_unidad]

            idx_sin_unidad = (~has_unit).to_numpy()
            ci_no_op[idx_sin_unidad] = costo_x_viaje_sin

            df_trips_work["CI_no_operativo"] = ci_no_op

            # CI LIGADOS A OPERACIÓN a nivel viaje
            asignaciones_df = st.session_state["asignaciones_df"].copy()
            conceptos_tipos = st.session_state["conceptos_tipos"].copy()

            tot_client_conc = (
                asignaciones_df.groupby(["Customer", "Concepto"], as_index=False)["Costo asignado"].sum()
            )
            tot_client_conc = tot_client_conc.merge(conceptos_tipos, on="Concepto", how="left")

            df_trips_work["CI_op_ligado_operacion"] = 0.0

            trailer_flag_trip = build_flag_trailer(df_trips_work[col_trailer], CFG["trailer_prefix"]).astype(bool)

            def asignar_op_por_subgrupo(df, mask_aplica, has_unit_s, dist_s, col_dest, monto):
                if monto == 0:
                    return
                mask_aplica = mask_aplica.fillna(False)
                n_total = int(mask_aplica.sum())
                if n_total == 0:
                    return

                n_u = int((mask_aplica & has_unit_s).sum())
                mask_su = mask_aplica & (~has_unit_s)
                n_su = int(mask_su.sum())

                pct_u = n_u / n_total
                pct_su = n_su / n_total

                bolsa_u = monto * pct_u
                bolsa_su = monto * pct_su

                if n_su > 0 and bolsa_su != 0:
                    df.loc[mask_su, col_dest] += (bolsa_su / n_su)

                mask_u_dist = mask_aplica & has_unit_s & (dist_s > 0)
                total_dist_u = float(dist_s.where(mask_u_dist, 0).sum())

                if total_dist_u > 0 and bolsa_u != 0:
                    costo_x = bolsa_u / total_dist_u
                    df.loc[mask_u_dist, col_dest] += costo_x * dist_s.where(mask_u_dist, 0)
                else:
                    mask_u_fallback = mask_aplica & has_unit_s
                    n_u_fallback = int(mask_u_fallback.sum())
                    if n_u_fallback > 0 and bolsa_u != 0:
                        df.loc[mask_u_fallback, col_dest] += (bolsa_u / n_u_fallback)

            for _, row in tot_client_conc.iterrows():
                cliente = str(row["Customer"])
                tipo_dist = row.get("Tipo distribución", "Volumen Viajes")
                monto_cliente = float(row["Costo asignado"]) if pd.notna(row["Costo asignado"]) else 0.0
                if monto_cliente == 0:
                    continue

                mask_base = df_trips_work[col_customer].astype(str) == cliente
                if not mask_base.any():
                    continue

                if tipo_dist == "Volumen Viajes":
                    mask_aplica = mask_base
                elif tipo_dist == "Viajes con unidad":
                    mask_aplica = mask_base & has_unit
                elif tipo_dist == "Viajes con Remolque":
                    mask_aplica = mask_base & trailer_flag_trip
                elif tipo_dist == DIST_COL_NAME:
                    mask_aplica = mask_base & has_unit
                else:
                    mask_aplica = mask_base

                asignar_op_por_subgrupo(
                    df=df_trips_work,
                    mask_aplica=mask_aplica,
                    has_unit_s=has_unit,
                    dist_s=dist,
                    col_dest="CI_op_ligado_operacion",
                    monto=monto_cliente,
                )

            df_trips_work["CI_total"] = df_trips_work["CI_no_operativo"] + df_trips_work["CI_op_ligado_operacion"]

            st.subheader("Vista previa con CI asignado")
            st.dataframe(df_trips_work.head(), width="stretch")

            def trips_to_excel_bytes(df):
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False, sheet_name="CI_por_viaje")
                return buffer.getvalue()

            st.download_button(
                "📥 Descargar viajes con CI asignado (Excel)",
                data=trips_to_excel_bytes(df_trips_work),
                file_name="viajes_con_CI_asignado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Error procesando la base de viajes detallados: {e}")
