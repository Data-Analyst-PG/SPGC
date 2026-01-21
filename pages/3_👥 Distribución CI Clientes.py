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

st.title("👥 Prorrateo de Costos Indirectos por Cliente")

st.markdown(
    """
Esta app te ayuda a:
1. Calcular, a partir de la **DATA detallada**, la tabla de viajes por cliente/año/mes
   (viajes totales, con remolque, con unidad, millas),
   excluyendo ciertos operadores logísticos.
2. Repartir costos **no ligados a la operación** entre viajes con/sin unidad
   (y obtener costos unitarios).
3. Definir un catálogo de distribución para costos **ligados a la operación**.
4. Prorratear esos costos entre clientes.
5. **Asignar los costos indirectos (CI) a nivel viaje**, permitiendo que los
   viajes de operadores “excluidos” también reciban CI si tienen unidad y millas.
"""
)

# ============================================================
# Función auxiliar para normalizar tipo_distribucion
# ============================================================

def normaliza_tipo_distribucion(valor):
    """
    Convierte distintos formatos de almacenamiento a una cadena simple:
    - list -> primer elemento
    - '["Volumen Viajes"]' -> 'Volumen Viajes'
    - '{Volumen Viajes}' -> 'Volumen Viajes'
    """
    if isinstance(valor, list):
        return valor[0] if valor else None

    if isinstance(valor, str):
        v = valor.strip()
        if not v:
            return None

        # Intentar como JSON
        try:
            parsed = json.loads(v)
            if isinstance(parsed, list) and parsed:
                return str(parsed[0])
            if isinstance(parsed, str):
                return parsed
        except Exception:
            pass

        # Limpiar brackets o llaves y comillas
        v = v.strip("[]{}").strip().strip('"\'')

        return v or None

    return valor

# ============================================================
# Función auxiliar para encontrar columnas por candidatos
# ============================================================

def find_column(df, candidates):
    """
    Busca en df.columns una columna cuyo nombre "normalizado"
    coincida con alguno de los candidatos.
    Normalización: minúsculas, sin espacios ni guiones bajos.
    """
    norm_map = {
        str(c).lower().replace(" ", "").replace("_", ""): c for c in df.columns
    }
    for cand in candidates:
        key = cand.lower().replace(" ", "").replace("_", "")
        if key in norm_map:
            return norm_map[key]
    return None

# ============================================================
# Lista de operadores "excluidos" para el cálculo de drivers
# (se ignoran en la tabla por cliente, pero SÍ pueden recibir CI
# en el paso 5 si tienen unidad y millas).
# ============================================================

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

# ============================================================
# 1️⃣ CARGA DE LA DATA DETALLADA Y CÁLCULO DE TABLA POR CLIENTE
# ============================================================
st.header("1️⃣ Cargar DATA detallada de viajes y generar tabla por cliente")

file_data = st.file_uploader(
    "Sube la DATA detallada de viajes (ej. DATA LINCOLN FREIGHT MASTER.xlsx)",
    type=["xlsx"],
    key="data_file",
)

tabla = None

if file_data:
    try:
        xls = pd.ExcelFile(file_data)
        hoja_data = st.selectbox(
            "Selecciona la hoja con la DATA detallada",
            xls.sheet_names,
        )
        df_data = pd.read_excel(xls, sheet_name=hoja_data)
        df_data.columns = df_data.columns.astype(str)

        st.subheader("Vista previa DATA de viajes")
        st.dataframe(df_data.head(), use_container_width=True)

        # --- Detectar columnas clave ---
        col_fecha = find_column(
            df_data,
            ["Bill date", "Bill Date", "Fecha", "FECHA", "Date"],
        )
        col_customer = find_column(df_data, ["Customer"])
        col_trip = find_column(df_data, ["Trip Number", "Trip number", "TripNumber"])
        col_trailer = find_column(df_data, ["Trailer", "Remolque"])
        col_unit = find_column(df_data, ["Unit", "Unidad"])
        col_miles = find_column(
            df_data,
            ["Real Miles", "Real miles", "REAL MILES", "Miles reales", "Real_miles"],
        )
        col_logop = find_column(
            df_data,
            ["Logistic Operator", "Operador logistico", "Operador logístico"],
        )

        faltan_cols = []
        if col_fecha is None:
            faltan_cols.append("Bill date / Fecha")
        if col_customer is None:
            faltan_cols.append("Customer")
        if col_trip is None:
            faltan_cols.append("Trip Number")
        if col_trailer is None:
            faltan_cols.append("Trailer")
        if col_unit is None:
            faltan_cols.append("Unit")
        if col_miles is None:
            faltan_cols.append("Real Miles")
        if col_logop is None:
            faltan_cols.append("Logistic Operator")

        if faltan_cols:
            st.error(
                "No se encontraron las siguientes columnas necesarias en la DATA: "
                + ", ".join(faltan_cols)
            )
        else:
            # ----------------------------------
            # 1) Preparar DATA completa (base CI)
            # ----------------------------------
            df_data[col_fecha] = pd.to_datetime(df_data[col_fecha], errors="coerce")
            df_data["Año"] = df_data[col_fecha].dt.year
            df_data["Mes"] = df_data[col_fecha].dt.month

            df_data["_trip"] = df_data[col_trip].astype(str)
            df_data["_flag_trailer_lf"] = (
                df_data[col_trailer].astype(str).str.upper().str.startswith("LF")
            ).astype(int)
            df_data["_flag_unidad"] = (
                df_data[col_unit].notna()
                & (df_data[col_unit].astype(str).str.strip() != "")
            ).astype(int)
            df_data["_millas"] = pd.to_numeric(
                df_data[col_miles], errors="coerce"
            ).fillna(0.0)
            df_data["_millas_con_unidad"] = df_data["_millas"] * df_data["_flag_unidad"]

            # Guardar copia de la DATA ya preparada para usarla en el Paso 5
            st.session_state["df_data_original"] = df_data.copy()

            # ----------------------------------
            # 2) Filtro de operadores para drivers (como tu SQL)
            #    ESTA DATA SE USA PARA LA TABLA POR CLIENTE Y TIPO_CLIENTE
            # ----------------------------------
            logop_upper = df_data[col_logop].astype(str).str.upper().str.strip()
            mask_log_ok = (logop_upper == "") | (~logop_upper.isin(OPERADORES_EXCLUIR))
            df_filt = df_data[mask_log_ok].copy()

            st.write(
                f"Registros usados para drivers (excluyendo operadores lista): "
                f"{df_filt.shape[0]} de {df_data.shape[0]}."
            )

            # ----------------------------------
            # 3) Selección de año/mes (sobre la DATA COMPLETA)
            #    Para que Paso 1 y Paso 5 compartan la misma base temporal
            # ----------------------------------
            col1, col2 = st.columns(2)
            with col1:
                anios = sorted(df_data["Año"].dropna().unique())
                anio_sel = st.selectbox("Año", anios)
            with col2:
                meses = sorted(
                    df_data.loc[df_data["Año"] == anio_sel, "Mes"].dropna().unique()
                )
                mes_sel = st.selectbox("Mes (número)", meses)

            st.session_state["anio_sel"] = int(anio_sel)
            st.session_state["mes_sel"] = int(mes_sel)

            # DATA del mes (BASE CI) -> incluye TODOS los operadores
            df_data_mes = df_data[
                (df_data["Año"] == anio_sel) & (df_data["Mes"] == mes_sel)
            ].copy()

            # DATA del mes FILTRADA por operador (como tu SQL) -> drivers por cliente
            df_mes_filt = df_filt[
                (df_filt["Año"] == anio_sel) & (df_filt["Mes"] == mes_sel)
            ].copy()

            if df_data_mes.empty:
                st.warning("No hay registros para ese año/mes en la DATA.")
            else:
                # ----------------------------------
                # 4) Totales base CI (TODOS los viajes del mes)
                #    Esto debe cuadrar con lo que uses en el Paso 5
                # ----------------------------------
                total_viajes_ci = df_data_mes["_trip"].nunique()
                total_con_unidad_ci = int(df_data_mes["_flag_unidad"].sum())
                total_sin_unidad_ci = total_viajes_ci - total_con_unidad_ci
                total_con_remolque_ci = int(df_data_mes["_flag_trailer_lf"].sum())
                millas_mes_ci = float(
                    df_data_mes.loc[df_data_mes["_flag_unidad"] == 1, "_millas"].sum()
                )

                pct_con_unidad_ci = (
                    total_con_unidad_ci / total_viajes_ci if total_viajes_ci else 0.0
                )
                pct_sin_unidad_ci = 1.0 - pct_con_unidad_ci if total_viajes_ci else 0.0

                resumen_ci = pd.DataFrame(
                    {
                        "Métrica": [
                            "Viajes totales (base CI)",
                            "Viajes con unidad (base CI)",
                            "Viajes sin unidad (base CI)",
                            "Viajes con remolque (base CI)",
                            "Millas con unidad (base CI)",
                        ],
                        "Valor": [
                            total_viajes_ci,
                            total_con_unidad_ci,
                            total_sin_unidad_ci,
                            total_con_remolque_ci,
                            millas_mes_ci,
                        ],
                    }
                )

                st.subheader(
                    f"Totales del mes {anio_sel}-{mes_sel:02d} para CI "
                    "(TODOS los viajes del mes)"
                )
                st.dataframe(resumen_ci, use_container_width=True)

                st.write(
                    f"**% viajes con unidad (base CI):** {pct_con_unidad_ci:.4%}   |   "
                    f"**% viajes sin unidad (base CI):** {pct_sin_unidad_ci:.4%}"
                )

                # Guardar estos datos para el Paso 2 (costos no operativos)
                st.session_state["pct_con_unidad"] = pct_con_unidad_ci
                st.session_state["pct_sin_unidad"] = pct_sin_unidad_ci
                st.session_state["millas_mes"] = millas_mes_ci
                st.session_state["total_sin_unidad"] = total_sin_unidad_ci

                # ----------------------------------
                # 5) Tabla por cliente / tipo_cliente (FILTRADA como tu SQL)
                #    Esta se usa para pasos 3 y 4 (drivers por cliente).
                # ----------------------------------
                if df_mes_filt.empty:
                    st.warning(
                        "Para ese año/mes, después de filtrar operadores, "
                        "no quedan registros para la tabla por cliente."
                    )
                else:
                    tabla_mes = (
                        df_mes_filt.groupby(["Año", "Mes", col_customer], as_index=False)
                        .agg(
                            Viajes=("_trip", "nunique"),
                            **{
                                "Viajes con remolques": ("_flag_trailer_lf", "sum"),
                                "Viajes con unidad": ("_flag_unidad", "sum"),
                                "Millas": ("_millas_con_unidad", "sum"),
                            },
                        )
                    )

                    tabla_mes = tabla_mes.rename(columns={col_customer: "Customer"})
                    tabla_mes["Viajes sin unidad"] = (
                        tabla_mes["Viajes"] - tabla_mes["Viajes con unidad"]
                    )

                    # Porcentajes y clasificaciones (como tu SQL)
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

                    st.subheader(
                        f"Tabla por cliente"
                        f"(DATA filtrada por operadores, {anio_sel}-{mes_sel:02d})"
                    )
                    st.dataframe(tabla_mes, use_container_width=True)

                    # Guardar para pasos 3 y 4
                    st.session_state["df_mes_clientes"] = tabla_mes

    except Exception as e:
        st.error(f"Error leyendo la DATA: {e}")


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
            hoja_sel = st.selectbox(
                "Hoja de costos no operativos",
                xls_no.sheet_names,
            )
            df_no = pd.read_excel(xls_no, sheet_name=hoja_sel)
            df_no.columns = df_no.columns.astype(str).str.strip()

            st.subheader("Vista previa costos no operativos")
            st.dataframe(df_no.head(), use_container_width=True)

            columnas_mes = [
                c
                for c in df_no.columns
                if c not in ["Concepto", "CONCEPTO"] and df_no[c].dtype != "O"
            ] or [c for c in df_no.columns if c not in ["Concepto", "CONCEPTO"]]

            # Sugerir por default el mes elegido en el Paso 1 (si coincide el nombre)
            mes_sel_global = st.session_state.get("mes_sel")
            index_default = 0
            if mes_sel_global is not None and len(columnas_mes) > 0:
                mapa_meses = {
                    1: "Ene",
                    2: "Feb",
                    3: "Mar",
                    4: "Abr",
                    5: "May",
                    6: "Jun",
                    7: "Jul",
                    8: "Ago",
                    9: "Sep",
                    10: "Oct",
                    11: "Nov",
                    12: "Dic",
                }
                nombre_mes = mapa_meses.get(int(mes_sel_global))
                if nombre_mes in columnas_mes:
                    index_default = columnas_mes.index(nombre_mes)
            
            col_mes_sel = st.selectbox(
                "Selecciona la columna del mes (ej. 'Ene')",
                columnas_mes,
                index=index_default,
            )

            concepto_col = "Concepto" if "Concepto" in df_no.columns else "CONCEPTO"
            df_no["Concepto"] = df_no[concepto_col].astype(str)
            df_no["Monto_mes"] = pd.to_numeric(df_no[col_mes_sel], errors="coerce").fillna(0)

            costo_total_no_op = df_no["Monto_mes"].sum()

            st.write(f"**Costo total no operativo del mes:** ${costo_total_no_op:,.2f}")

            pct_con_unidad = st.session_state["pct_con_unidad"]
            pct_sin_unidad = st.session_state["pct_sin_unidad"]
            millas_mes = st.session_state["millas_mes"]
            total_sin_unidad = st.session_state["total_sin_unidad"]

            bolsa_con_unidad = costo_total_no_op * pct_con_unidad
            bolsa_sin_unidad = costo_total_no_op * pct_sin_unidad

            st.write(
                f"**Monto asignado a viajes con unidad:** ${bolsa_con_unidad:,.2f}  "
                f"({pct_con_unidad:.4%})"
            )
            st.write(
                f"**Monto asignado a viajes sin unidad:** ${bolsa_sin_unidad:,.2f}  "
                f"({pct_sin_unidad:.4%})"
            )

            costo_x_milla = bolsa_con_unidad / millas_mes if millas_mes else None
            costo_x_viaje_sin = (
                bolsa_sin_unidad / total_sin_unidad if total_sin_unidad else None
            )

            st.subheader("Costos unitarios derivados (informativos)")
            if costo_x_milla is not None:
                st.write(f"**Costo por milla (viajes con unidad):** ${costo_x_milla:,.6f}")
            else:
                st.write("No se pudo calcular costo por milla (faltan millas totales).")

            if costo_x_viaje_sin is not None:
                st.write(
                    f"**Costo por viaje sin unidad:** ${costo_x_viaje_sin:,.6f}"
                )
            else:
                st.write(
                    "No se pudo calcular costo por viaje sin unidad "
                    "(no hay viajes sin unidad en el mes)."
                )

            # Guardar en estado
            st.session_state["costo_no_op_total"] = costo_total_no_op
            st.session_state["costo_no_op_x_milla"] = costo_x_milla
            st.session_state["costo_no_op_x_viaje_sin"] = costo_x_viaje_sin

        except Exception as e:
            st.error(f"Error procesando los costos no operativos: {e}")

# ============================================================
# 3️⃣ CATÁLOGO PARA COSTOS LIGADOS A LA OPERACIÓN
# ============================================================
st.header("3️⃣ Catálogo de costos ligados a la operación")

tipos_distribucion = [
    "Volumen Viajes",        # Usa columna 'Viajes'
    "Viajes con Remolque",  # Usa 'Viajes con remolques'
    "Viajes con unidad",    # Usa 'Viajes con unidad'
    "Millas",               # Usa columna 'Millas'
]

# Cargar catálogo existente desde Supabase
try:
    data_cat = supabase.table("catalogo_costos_clientes").select("*").execute().data
    catalogo_existente = pd.DataFrame(data_cat)
except Exception:
    catalogo_existente = pd.DataFrame()

file_op = st.file_uploader(
    "Sube el archivo de costos ligados a operación (Concepto + meses)",
    type=["xlsx"],
    key="op_file",
)

df_op_mes = None

if file_op:
    try:
        xls_op = pd.ExcelFile(file_op)
        hoja_op_sel = st.selectbox(
            "Hoja de costos ligados a operación",
            xls_op.sheet_names,
        )
        df_op = pd.read_excel(xls_op, sheet_name=hoja_op_sel)
        df_op.columns = df_op.columns.astype(str).str.strip()

        st.subheader("Vista previa costos ligados a operación")
        st.dataframe(df_op.head(), use_container_width=True)

        columnas_mes_op = [
            c
            for c in df_op.columns
            if c not in ["Concepto", "CONCEPTO"] and df_op[c].dtype != "O"
        ] or [c for c in df_op.columns if c not in ["Concepto", "CONCEPTO"]]

        # Sugerir por default el mes elegido en el Paso 1
        mes_sel_global = st.session_state.get("mes_sel")
        index_default_op = 0
        if mes_sel_global is not None and len(columnas_mes_op) > 0:
            mapa_meses = {
                1: "Ene",
                2: "Feb",
                3: "Mar",
                4: "Abr",
                5: "May",
                6: "Jun",
                7: "Jul",
                8: "Ago",
                9: "Sep",
                10: "Oct",
                11: "Nov",
                12: "Dic",
            }
            nombre_mes = mapa_meses.get(int(mes_sel_global))
            if nombre_mes in columnas_mes_op:
                index_default_op = columnas_mes_op.index(nombre_mes)

        col_mes_op = st.selectbox(
            "Selecciona la columna del mes a prorratear (ej. 'Ene')",
            columnas_mes_op,
            index=index_default_op,
        )

        concepto_col_op = "Concepto" if "Concepto" in df_op.columns else "CONCEPTO"
        df_op["Concepto"] = df_op[concepto_col_op].astype(str).str.strip()
        df_op["Monto_mes"] = pd.to_numeric(df_op[col_mes_op], errors="coerce").fillna(0)

        st.write(
            f"**Total costos ligados a operación ({col_mes_op}):** "
            f"${df_op['Monto_mes'].sum():,.2f}"
        )

        # ========================
        # Construir catálogo editable
        # ========================
        conceptos = df_op[["Concepto"]].drop_duplicates().reset_index(drop=True)

        if not catalogo_existente.empty:
            catalogo_existente = catalogo_existente.rename(
                columns={
                    "concepto": "Concepto",
                    "tipo_distribucion": "Tipo distribución",
                }
            )

            catalogo_existente["Tipo distribución"] = catalogo_existente[
                "Tipo distribución"
            ].apply(normaliza_tipo_distribucion)

            merged_cat = conceptos.merge(
                catalogo_existente,
                on="Concepto",
                how="left",
            )
        else:
            merged_cat = conceptos.copy()
            merged_cat["Tipo distribución"] = None

        st.subheader("Catálogo de distribución por concepto")
        merged_cat = merged_cat.sort_values(
            by=["Tipo distribución", "Concepto"],
            na_position="first",
        ).reset_index(drop=True)

        edited_cat = st.data_editor(
            merged_cat,
            use_container_width=True,
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
                        registros.append(
                            {
                                "concepto": row["Concepto"],
                                "tipo_distribucion": str(row["Tipo distribución"]),
                            }
                        )
                if registros:
                    supabase.table("catalogo_costos_clientes").upsert(
                        registros,
                        on_conflict="concepto",
                    ).execute()
                st.success("Catálogo actualizado en Supabase.")
            except Exception as e:
                st.error(f"Error al guardar el catálogo: {e}")

        # Guardar en sesión para el prorrateo
        st.session_state["df_costos_op_mes"] = df_op[["Concepto", "Monto_mes"]]
    except Exception as e:
        st.error(f"Error procesando los costos ligados a operación: {e}")

# ============================================================
# 4️⃣ PRORRATEO DE COSTOS LIGADOS A OPERACIÓN ENTRE CLIENTES
# ============================================================
st.header("4️⃣ Prorrateo de costos ligados a operación por cliente")

# Mostrar de nuevo los costos unitarios informativos del paso 2 (si existen)
costo_x_milla_info = st.session_state.get("costo_no_op_x_milla")
costo_x_viaje_sin_info = st.session_state.get("costo_no_op_x_viaje_sin")

if (costo_x_milla_info is not None) or (costo_x_viaje_sin_info is not None):
    st.subheader("Recordatorio costos unitarios no operativos (informativos)")
    if costo_x_milla_info is not None:
        st.write(f"**Costo por milla (viajes con unidad):** ${costo_x_milla_info:,.6f}")
    if costo_x_viaje_sin_info is not None:
        st.write(
            f"**Costo por viaje sin unidad:** ${costo_x_viaje_sin_info:,.6f}"
        )

if (
    "df_mes_clientes" not in st.session_state
    or "df_costos_op_mes" not in st.session_state
):
    st.info("Necesitas completar los pasos 1 y 3 para poder prorratear.")
else:
    df_mes_clientes = st.session_state["df_mes_clientes"].copy()
    df_op_mes = st.session_state["df_costos_op_mes"].copy()

    # Releer catálogo desde Supabase
    data_cat = supabase.table("catalogo_costos_clientes").select("*").execute().data
    catalogo = pd.DataFrame(data_cat)
    if catalogo.empty:
        st.error("No hay catálogo de distribución en 'catalogo_costos_clientes'.")
    else:
        catalogo = catalogo.rename(
            columns={"concepto": "Concepto", "tipo_distribucion": "Tipo distribución"}
        )

        catalogo["Tipo distribución"] = catalogo["Tipo distribución"].apply(
            normaliza_tipo_distribucion
        )

        # Unir catálogo con los costos del mes
        df_op_mes = df_op_mes.merge(catalogo, on="Concepto", how="left")

        if df_op_mes["Tipo distribución"].isna().any():
            faltan = df_op_mes.loc[
                df_op_mes["Tipo distribución"].isna(), "Concepto"
            ].unique()
            st.error(
                "Hay conceptos sin tipo de distribución definido en el catálogo: "
                + ", ".join(faltan[:10])
                + ("..." if len(faltan) > 10 else "")
            )
        else:
            # Mapear tipo de distribución -> columna de driver
            driver_map = {
                "Volumen Viajes": "Viajes",
                "Viajes con Remolque": "Viajes con remolques",
                "Viajes con unidad": "Viajes con unidad",
                "Millas": "Millas",
            }

            # Preparar tabla base por cliente
            base_clientes = df_mes_clientes.groupby(
                ["Customer"], as_index=False
            ).agg(
                {
                    "Viajes": "sum",
                    "Viajes con remolques": "sum",
                    "Viajes con unidad": "sum",
                    "Millas": "sum",
                }
            )

            # Vamos a ir acumulando costos por concepto
            asignaciones = []

            for _, row in df_op_mes.iterrows():
                concepto = row["Concepto"]
                monto = float(row["Monto_mes"])
                tipo_dist = row["Tipo distribución"]

                col_driver = driver_map.get(tipo_dist)
                if col_driver not in base_clientes.columns:
                    st.warning(
                        f"Tipo de distribución '{tipo_dist}' requiere columna '{col_driver}', "
                        f"que no existe. Se omite el concepto {concepto}."
                    )
                    continue

                df_driver = base_clientes[["Customer", col_driver]].copy()
                total_driver = df_driver[col_driver].sum()

                if total_driver == 0:
                    st.warning(
                        f"El driver '{col_driver}' para el concepto {concepto} "
                        f"es 0 en el mes. Se omite."
                    )
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
                st.dataframe(asignaciones_df, use_container_width=True)

                # Pivot para ver totales por cliente
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

                pivot_clientes["Total costos ligados op"] = (
                    pivot_clientes.drop(
                        columns=["Customer"]
                    ).sum(axis=1)
                )

                st.subheader("Totales por cliente (solo costos ligados a la operación)")
                st.dataframe(pivot_clientes, use_container_width=True)

                # Guardar para el paso 5
                st.session_state["asignaciones_df"] = asignaciones_df
                st.session_state["conceptos_tipos"] = df_op_mes[
                    ["Concepto", "Tipo distribución"]
                ]
                st.session_state["pivot_clientes_ci"] = pivot_clientes

                # Exportar resultados
                def to_excel_bytes(df1, df2, df_unitarios=None):
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                        df1.to_excel(
                            writer,
                            index=False,
                            sheet_name="Detalle_asignaciones",
                        )
                        df2.to_excel(
                            writer,
                            index=False,
                            sheet_name="Totales_por_cliente",
                        )
                        if df_unitarios is not None:
                            df_unitarios.to_excel(
                                writer,
                                index=False,
                                sheet_name="Costos_no_operativos_unit",
                            )
                    return buffer.getvalue()

                # Armar hoja con costos unitarios no operativos (si existen)
                df_unitarios = None
                if (costo_x_milla_info is not None) or (
                    costo_x_viaje_sin_info is not None
                ):
                    filas_unit = []
                    if costo_x_milla_info is not None:
                        filas_unit.append(
                            {
                                "Concepto": "Costo por milla (viajes con unidad)",
                                "Valor": costo_x_milla_info,
                            }
                        )
                    if costo_x_viaje_sin_info is not None:
                        filas_unit.append(
                            {
                                "Concepto": "Costo por viaje sin unidad",
                                "Valor": costo_x_viaje_sin_info,
                            }
                        )
                    if filas_unit:
                        df_unitarios = pd.DataFrame(filas_unit)

                st.download_button(
                    "📥 Descargar resultados (Excel clientes)",
                    data=to_excel_bytes(asignaciones_df, pivot_clientes, df_unitarios),
                    file_name="prorrateo_costos_clientes.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                )


# ============================================================
# 5️⃣ ASIGNACIÓN DE CI A NIVEL VIAJE
# ============================================================
st.header("5️⃣ Asignación de CI a nivel viaje")

# Validar que tenemos todo lo necesario
faltan_requisitos = []
if st.session_state.get("costo_no_op_x_milla") is None:
    faltan_requisitos.append("Costo unitario por milla (paso 2).")
if st.session_state.get("costo_no_op_x_viaje_sin") is None:
    faltan_requisitos.append("Costo unitario por viaje sin unidad (paso 2).")
if "asignaciones_df" not in st.session_state or "conceptos_tipos" not in st.session_state:
    faltan_requisitos.append("Prorrateo por cliente (paso 4).")

if faltan_requisitos:
    st.info(
        "Para usar este apartado necesitas haber completado:\n- "
        + "\n- ".join(faltan_requisitos)
    )
else:
    # Elegir origen de la base de viajes
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
            st.error(
                "No se encontró la DATA del paso 1 en memoria. "
                "Vuelve a cargarla en el Paso 1 o elige 'Subir otro archivo'."
            )
        else:
            df_trips = st.session_state["df_data_original"].copy()
    else:
        file_trips = st.file_uploader(
            "Sube la base de viajes a nivel detalle (ej. DATA LINCOLN 2025..xlsx)",
            type=["xlsx"],
            key="file_trips_ci",
        )

    if (df_trips is not None) or file_trips:
        try:
            # Si se eligió subir archivo, leemos el Excel aquí
            if df_trips is None and file_trips is not None:
                xls_trips = pd.ExcelFile(file_trips)
                hoja_trips = st.selectbox(
                    "Hoja con los viajes detallados",
                    xls_trips.sheet_names,
                    key="hoja_trips_sel",
                )
                df_trips = pd.read_excel(xls_trips, sheet_name=hoja_trips)
                df_trips.columns = df_trips.columns.astype(str)

            # --- Filtrar por año/mes usando columna de fecha (si existe) ---
            col_fecha = find_column(
                df_trips,
                ["Fecha", "FECHA", "Bill date", "Bill Date", "Date"],
            )

            if (
                col_fecha is not None
                and "anio_sel" in st.session_state
                and "mes_sel" in st.session_state
            ):
                df_trips[col_fecha] = pd.to_datetime(df_trips[col_fecha], errors="coerce")
                anio_sel = int(st.session_state["anio_sel"])
                mes_sel = int(st.session_state["mes_sel"])

                mask_mes = (
                    (df_trips[col_fecha].dt.year == anio_sel)
                    & (df_trips[col_fecha].dt.month == mes_sel)
                )

                df_trips_mes = df_trips[mask_mes].copy()
                st.write(
                    f"Se usarán {df_trips_mes.shape[0]} viajes del mes "
                    f"{mes_sel:02d}/{anio_sel} de un total de {df_trips.shape[0]} viajes en la base."
                )
                df_trips = df_trips_mes
            else:
                st.warning(
                    "No se encontró una columna de fecha o no está definido el año/mes. "
                    "Se usarán todos los viajes de la base."
                )

            st.subheader("Vista previa viajes (después de filtro por mes)")
            st.dataframe(df_trips.head(), use_container_width=True)

            # Buscar columnas clave
            col_customer = find_column(df_trips, ["Customer"])
            col_trip = find_column(df_trips, ["Trip Number", "Trip number", "TripNumber"])
            col_operador = find_column(
                df_trips,
                [
                    "Operador logistico",
                    "Operador logístico",
                    "OPERADOR LOGISTICO",
                    "Operador_Logistico",
                    "Logistic Operator",
                ],
            )
            col_unit = find_column(df_trips, ["Unit", "Unidad"])
            col_trailer = find_column(df_trips, ["Trailer", "Remolque"])
            col_miles = find_column(
                df_trips,
                ["Real Miles", "Real miles", "REAL MILES", "Miles reales", "Real_miles", "Real Mi"],
            )

            columnas_faltan = []
            if col_customer is None:
                columnas_faltan.append("Customer")
            if col_trip is None:
                columnas_faltan.append("Trip Number")
            if col_operador is None:
                columnas_faltan.append("Operador logistico / Logistic Operator")
            if col_unit is None:
                columnas_faltan.append("Unit / Unidad")
            if col_trailer is None:
                columnas_faltan.append("Trailer / Remolque")
            if col_miles is None:
                columnas_faltan.append("Real Miles")

            if columnas_faltan:
                st.error(
                    "No se encontraron las siguientes columnas necesarias en el archivo de viajes: "
                    + ", ".join(columnas_faltan)
                )
            else:
                df_trips_work = df_trips.copy()

                # Normalizar valores clave
                df_trips_work[col_customer] = df_trips_work[col_customer].astype(str)
                df_trips_work[col_trip] = df_trips_work[col_trip].astype(str)

                # Identificar operadores "excluidos" (solo informativo)
                op_upper = df_trips_work[col_operador].astype(str).str.upper().str.strip()
                mask_excl = op_upper.isin(OPERADORES_EXCLUIR)
                df_trips_work["Excluido_por_operador"] = mask_excl

                st.write(
                    f"Viajes con operador en lista de 'excluidos': {int(mask_excl.sum())} "
                    f"de {len(df_trips_work)} (estos viajes **sí** pueden recibir CI si "
                    "tienen unidad y millas)."
                )

                # ======================================================
                # CI NO OPERATIVOS a nivel viaje
                # ======================================================
                costo_x_milla = st.session_state["costo_no_op_x_milla"]
                costo_x_viaje_sin = st.session_state["costo_no_op_x_viaje_sin"]

                has_unit = df_trips_work[col_unit].notna() & (
                    df_trips_work[col_unit].astype(str).str.strip() != ""
                )
                miles = pd.to_numeric(df_trips_work[col_miles], errors="coerce").fillna(0.0)

                ci_no_op = np.zeros(len(df_trips_work), dtype=float)

                # Con unidad -> por milla (incluye operadores "excluidos")
                idx_con_unidad = (has_unit & (miles > 0)).to_numpy()
                ci_no_op[idx_con_unidad] = float(costo_x_milla) * miles.to_numpy()[idx_con_unidad]

                # Sin unidad -> por viaje (para todos los que NO tienen unidad)
                idx_sin_unidad = (~has_unit).to_numpy()
                ci_no_op[idx_sin_unidad] = float(costo_x_viaje_sin)

                df_trips_work["CI_no_operativo"] = ci_no_op

                st.subheader("Resumen de CI no operativo a nivel viaje")
                st.write(df_trips_work["CI_no_operativo"].describe(percentiles=[0.25, 0.5, 0.75]))

                # ======================================================
                # CI LIGADOS A OPERACIÓN a nivel viaje
                # ======================================================
                st.subheader("Asignación de CI ligados a la operación a nivel viaje")

                asignaciones_df = st.session_state["asignaciones_df"].copy()
                conceptos_tipos = st.session_state["conceptos_tipos"].copy()

                # Totales por cliente y concepto (monto total de CI operativo por cliente y concepto)
                tot_client_conc = (
                    asignaciones_df.groupby(["Customer", "Concepto"], as_index=False)["Costo asignado"]
                    .sum()
                )

                # Agregar el tipo de distribución por concepto (del catálogo)
                tot_client_conc = tot_client_conc.merge(conceptos_tipos, on="Concepto", how="left")

                # Inicializar columna CI operativo por viaje
                df_trips_work["CI_op_ligado_operacion"] = 0.0

                # Pre-calcular banderas por viaje
                has_unit_trip = has_unit
                trailer_str = df_trips_work[col_trailer].astype(str)
                mask_trailer_lf = trailer_str.str.upper().str.startswith("LF")

                def asignar_op_por_subgrupo(df, mask_aplica, has_unit_s, miles_s, col_dest, monto):
                    """
                    Dentro de mask_aplica:
                      - separa viajes con unidad vs sin unidad
                      - parte el monto según % de viajes con/sin unidad (por conteo)
                      - bolsa sin unidad -> igual por viaje
                      - bolsa con unidad -> por millas (costo/milla * millas del trip)
                    """
                    if monto == 0:
                        return

                    mask_aplica = mask_aplica.fillna(False)
                    n_total = int(mask_aplica.sum())
                    if n_total == 0:
                        return

                    # conteos para porcentajes
                    n_u = int((mask_aplica & has_unit_s).sum())   # con unidad (aunque millas=0)
                    mask_su = mask_aplica & (~has_unit_s)         # sin unidad
                    n_su = int(mask_su.sum())

                    pct_u = n_u / n_total
                    pct_su = n_su / n_total

                    bolsa_u = monto * pct_u
                    bolsa_su = monto * pct_su

                    # 1) Sin unidad -> igual por viaje
                    if n_su > 0 and bolsa_su != 0:
                        df.loc[mask_su, col_dest] += (bolsa_su / n_su)

                    # 2) Con unidad -> por millas (solo donde millas > 0)
                    mask_u_millas = mask_aplica & has_unit_s & (miles_s > 0)
                    total_millas_u = float(miles_s.where(mask_u_millas, 0).sum())

                    if total_millas_u > 0 and bolsa_u != 0:
                        costo_x_milla_u = bolsa_u / total_millas_u
                        df.loc[mask_u_millas, col_dest] += costo_x_milla_u * miles_s.where(mask_u_millas, 0)
                    else:
                        # Fallback: si hay viajes con unidad pero sin millas válidas, repartir por viaje con unidad
                        mask_u_fallback = mask_aplica & has_unit_s
                        n_u_fallback = int(mask_u_fallback.sum())
                        if n_u_fallback > 0 and bolsa_u != 0:
                            df.loc[mask_u_fallback, col_dest] += (bolsa_u / n_u_fallback)

                # Loop: asignar por cliente-concepto
                for _, row in tot_client_conc.iterrows():
                    cliente = str(row["Customer"])
                    concepto = row["Concepto"]
                    tipo_dist = row.get("Tipo distribución", "Volumen Viajes")
                    monto_cliente = float(row["Costo asignado"]) if pd.notna(row["Costo asignado"]) else 0.0

                    if monto_cliente == 0:
                        continue

                    mask_base = df_trips_work[col_customer].astype(str) == cliente
                    if not mask_base.any():
                        continue

                    # Definir subconjunto que aplica según el tipo de distribución
                    if tipo_dist == "Volumen Viajes":
                        mask_aplica = mask_base

                    elif tipo_dist == "Viajes con unidad":
                        mask_aplica = mask_base & has_unit_trip

                    elif tipo_dist == "Viajes con Remolque":
                        mask_aplica = mask_base & mask_trailer_lf

                    elif tipo_dist == "Millas":
                        mask_aplica = mask_base & has_unit_trip

                    else:
                        mask_aplica = mask_base

                    # Asignar según tu regla: % unidad/sin unidad y unidad por millas
                    asignar_op_por_subgrupo(
                        df=df_trips_work,
                        mask_aplica=mask_aplica,
                        has_unit_s=has_unit_trip,
                        miles_s=miles,
                        col_dest="CI_op_ligado_operacion",
                        monto=monto_cliente,
                    )

                # ======================================================
                # Total CI por viaje y descarga
                # ======================================================
                df_trips_work["CI_total"] = (
                    df_trips_work["CI_no_operativo"] + df_trips_work["CI_op_ligado_operacion"]
                )

                st.subheader("Vista previa con CI asignado")
                st.dataframe(df_trips_work.head(), use_container_width=True)

                # Descarga a Excel
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
