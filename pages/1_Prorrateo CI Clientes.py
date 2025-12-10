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
1. Calcular porcentajes por tipo de cliente (viajes, remolques, unidades).
2. Repartir costos **no ligados a la operación** entre viajes con/sin unidad (y obtener costos unitarios).
3. Definir un catálogo de distribución para costos **ligados a la operación**.
4. Prorratear esos costos entre clientes.
5. **Asignar los costos indirectos (CI) a nivel viaje**, excluyendo ciertos operadores logísticos.
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
# 1️⃣ CARGA DEL ARCHIVO BASE (TABLA + OPCIONAL DISTRIBUCIÓN)
# ============================================================
st.header("1️⃣ Cargar archivo base de viajes (Tabla)")

file_base = st.file_uploader(
    "Sube el archivo base (ej. Resumen_Clientes_Lincoln.xlsx) con la hoja 'Tabla'",
    type=["xlsx"],
    key="base_file",
)

tabla = None
dist_df = None

if file_base:
    try:
        xls = pd.ExcelFile(file_base)
        if "Tabla" not in xls.sheet_names:
            st.error("El archivo no contiene una hoja llamada 'Tabla'.")
        else:
            tabla = pd.read_excel(xls, sheet_name="Tabla")
            tabla.columns = tabla.columns.str.strip()

            st.subheader("Vista previa de la hoja 'Tabla'")
            st.dataframe(tabla.head(), use_container_width=True)

        # Hoja opcional: Distribución (para millas totales por mes)
        if "Distribución" in xls.sheet_names:
            dist_df = pd.read_excel(xls, sheet_name="Distribución")
            dist_df.columns = dist_df.columns.astype(str).str.strip()
            st.info("Se detectó hoja 'Distribución' (para millas totales por mes).")

    except Exception as e:
        st.error(f"Error leyendo el archivo base: {e}")

if tabla is not None:
    # Validar columnas mínimas
    cols_minimas = [
        "Año",
        "Mes",
        "Customer",
        "Viajes",
        "Viajes con remolques",
        "Viajes con unidad",
    ]
    faltan = [c for c in cols_minimas if c not in tabla.columns]
    if faltan:
        st.error(f"Faltan columnas en 'Tabla': {faltan}")
        tabla = None

if tabla is not None:
    # Selección de año y mes
    col1, col2 = st.columns(2)
    with col1:
        anios = sorted(tabla["Año"].dropna().unique())
        anio_sel = st.selectbox("Año", anios)
    with col2:
        meses = sorted(tabla.loc[tabla["Año"] == anio_sel, "Mes"].dropna().unique())
        mes_sel = st.selectbox("Mes (número)", meses)

    df_mes = tabla[(tabla["Año"] == anio_sel) & (tabla["Mes"] == mes_sel)].copy()

    if df_mes.empty:
        st.warning("No hay registros para ese año/mes.")
    else:
        st.subheader(f"Resumen de viajes - {anio_sel}-{mes_sel:02d}")

        # Agregar columna de viajes sin unidad
        df_mes["Viajes sin unidad"] = df_mes["Viajes"] - df_mes["Viajes con unidad"]

        total_viajes = df_mes["Viajes"].sum()
        total_remolques = df_mes["Viajes con remolques"].sum()
        total_unidad = df_mes["Viajes con unidad"].sum()
        total_sin_unidad = df_mes["Viajes sin unidad"].sum()

        pct_con_unidad = total_unidad / total_viajes if total_viajes else 0
        pct_sin_unidad = total_sin_unidad / total_viajes if total_viajes else 0

        resumen_global = pd.DataFrame(
            {
                "Métrica": [
                    "Viajes totales",
                    "Viajes con remolques",
                    "Viajes con unidad",
                    "Viajes sin unidad",
                ],
                "Valor": [
                    total_viajes,
                    total_remolques,
                    total_unidad,
                    total_sin_unidad,
                ],
            }
        )

        st.write("**Totales del mes (todas los clientes):**")
        st.dataframe(resumen_global, use_container_width=True)

        st.write(
            f"**% viajes con unidad:** {pct_con_unidad:.4%}   |   "
            f"**% viajes sin unidad:** {pct_sin_unidad:.4%}"
        )

        # Millas totales del mes (si vienen de hoja Distribución)
        millas_mes = None
        if dist_df is not None:
            try:
                fila_millas = dist_df.loc[
                    dist_df["Mes"].astype(str).str.upper() == "MILLAS"
                ]
                mapa_mes = {
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
                col_mes = mapa_mes.get(int(mes_sel))
                if col_mes and col_mes in dist_df.columns and not fila_millas.empty:
                    millas_mes = float(fila_millas[col_mes].iloc[0])
                    st.write(f"**Millas totales del mes ({col_mes}):** {millas_mes:,.2f}")
            except Exception:
                st.warning(
                    "No se pudieron leer las millas de la hoja 'Distribución'. "
                    "Si quieres costo por milla, revisa el formato."
                )

        # Porcentajes por Tipo cliente
        if "Tipo cliente" in df_mes.columns:
            st.subheader("Porcentajes por Tipo de cliente")

            agg_cols = [
                "Viajes",
                "Viajes con remolques",
                "Viajes con unidad",
                "Viajes sin unidad",
            ]
            por_tipo = (
                df_mes.groupby("Tipo cliente")[agg_cols]
                .sum()
                .reset_index()
            )

            for col in agg_cols:
                total_col = por_tipo[col].sum()
                por_tipo[f"%{col}"] = por_tipo[col] / total_col if total_col else 0.0

            st.dataframe(por_tipo, use_container_width=True)
        else:
            st.info("La tabla no tiene columna 'Tipo cliente'; se omite el resumen por tipo.")

        # Guardar en session_state para secciones siguientes
        st.session_state["df_mes_clientes"] = df_mes
        st.session_state["pct_con_unidad"] = pct_con_unidad
        st.session_state["pct_sin_unidad"] = pct_sin_unidad
        st.session_state["millas_mes"] = millas_mes
        st.session_state["total_sin_unidad"] = total_sin_unidad

# ============================================================
# 2️⃣ COSTOS NO LIGADOS A LA OPERACIÓN
# ============================================================
st.header("2️⃣ Costos indirectos **no ligados a la operación**")

if "df_mes_clientes" not in st.session_state:
    st.warning("Primero carga el archivo base y selecciona año/mes en el paso 1.")
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
            col_mes_sel = st.selectbox(
                "Selecciona la columna del mes (ej. 'Ene')",
                columnas_mes,
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
    "Millas",               # Usa columna 'Millas' (si existe)
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
        col_mes_op = st.selectbox(
            "Selecciona la columna del mes a prorratear (ej. 'Ene')",
            columnas_mes_op,
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
                ["Customer", "Tipo cliente"], as_index=False
            ).agg(
                {
                    "Viajes": "sum",
                    "Viajes con remolques": "sum",
                    "Viajes con unidad": "sum",
                }
            )

            # Opcional: Millas por cliente si existe
            if "Millas" in df_mes_clientes.columns:
                millas_cliente = (
                    df_mes_clientes.groupby(["Customer", "Tipo cliente"])["Millas"]
                    .sum()
                    .reset_index()
                )
                base_clientes = base_clientes.merge(
                    millas_cliente,
                    on=["Customer", "Tipo cliente"],
                    how="left",
                )
            else:
                base_clientes["Millas"] = 0.0

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

                df_driver = base_clientes[["Customer", "Tipo cliente", col_driver]].copy()
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
                        index=["Customer", "Tipo cliente"],
                        columns="Concepto",
                        values="Costo asignado",
                        aggfunc="sum",
                        fill_value=0.0,
                    )
                    .reset_index()
                )

                pivot_clientes["Total costos ligados op"] = (
                    pivot_clientes.drop(
                        columns=["Customer", "Tipo cliente"]
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

# Lista de operadores a excluir
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
    file_trips = st.file_uploader(
        "Sube la base de viajes a nivel detalle (ej. DATA LINCOLN 2025..xlsx)",
        type=["xlsx"],
        key="file_trips_ci",
    )

    if file_trips:
        try:
            xls_trips = pd.ExcelFile(file_trips)
            hoja_trips = st.selectbox(
                "Hoja con los viajes detallados",
                xls_trips.sheet_names,
            )
            df_trips = pd.read_excel(xls_trips, sheet_name=hoja_trips)
            df_trips.columns = df_trips.columns.astype(str)

            st.subheader("Vista previa viajes")
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
                ],
            )
            col_unit = find_column(df_trips, ["Unit", "Unidad"])
            col_trailer = find_column(df_trips, ["Trailer", "Remolque"])
            col_miles = find_column(
                df_trips,
                ["Real Miles", "Real miles", "REAL MILES", "Miles reales", "Real_miles"],
            )

            columnas_faltan = []
            if col_customer is None:
                columnas_faltan.append("Customer")
            if col_trip is None:
                columnas_faltan.append("Trip Number")
            if col_operador is None:
                columnas_faltan.append("Operador logistico")
            if col_unit is None:
                columnas_faltan.append("Unit")
            if col_trailer is None:
                columnas_faltan.append("Trailer")
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

                # Identificar viajes excluidos por operador
                op_upper = df_trips_work[col_operador].astype(str).str.upper().str.strip()
                mask_excl = op_upper.isin(OPERADORES_EXCLUIR)
                df_trips_work["Excluido_por_operador"] = mask_excl

                st.write(
                    f"Viajes excluidos por operador logístico: {mask_excl.sum()} "
                    f"de {len(df_trips_work)}."
                )

                # =======================
                # CI NO OPERATIVOS a nivel viaje
                # =======================
                costo_x_milla = st.session_state["costo_no_op_x_milla"]
                costo_x_viaje_sin = st.session_state["costo_no_op_x_viaje_sin"]

                has_unit = df_trips_work[col_unit].notna() & (
                    df_trips_work[col_unit].astype(str).str.strip() != ""
                )
                miles = pd.to_numeric(df_trips_work[col_miles], errors="coerce").fillna(0)

                ci_no_op = np.zeros(len(df_trips_work))

                # Solo asignar a viajes NO excluidos
                idx_incluir = (~mask_excl).to_numpy()

                # Con unidad -> por milla
                if costo_x_milla is not None:
                    idx_con_unidad = (has_unit & ~mask_excl).to_numpy()
                    ci_no_op[idx_con_unidad] = costo_x_milla * miles.to_numpy()[idx_con_unidad]

                # Sin unidad -> por viaje
                if costo_x_viaje_sin is not None:
                    idx_sin_unidad = (~has_unit & ~mask_excl).to_numpy()
                    ci_no_op[idx_sin_unidad] = costo_x_viaje_sin

                df_trips_work["CI_no_operativo"] = ci_no_op

                # =======================
                # CI LIGADOS A OPERACIÓN a nivel viaje
                # =======================
                asignaciones_df = st.session_state["asignaciones_df"].copy()
                conceptos_tipos = st.session_state["conceptos_tipos"].copy()

                # Totales por cliente y concepto
                tot_client_conc = (
                    asignaciones_df.groupby(["Customer", "Concepto"], as_index=False)[
                        "Costo asignado"
                    ]
                    .sum()
                )

                tot_client_conc = tot_client_conc.merge(
                    conceptos_tipos, on="Concepto", how="left"
                )

                # Crear columna de CI op en trips
                df_trips_work["CI_op_ligado_operacion"] = 0.0

                # Pre-calcular banderas por viaje
                has_unit_trip = has_unit
                trailer_str = df_trips_work[col_trailer].astype(str)
                mask_trailer_lf = trailer_str.str.upper().str.startswith("LF")

                for _, row in tot_client_conc.iterrows():
                    cliente = str(row["Customer"])
                    concepto = row["Concepto"]
                    tipo_dist = row["Tipo distribución"]
                    monto_cliente = float(row["Costo asignado"])

                    # Filtrar viajes de este cliente (y no excluidos)
                    mask_cliente = df_trips_work[col_customer].astype(str) == cliente
                    mask_base = mask_cliente & (~mask_excl)

                    if not mask_base.any():
                        continue

                    if tipo_dist == "Volumen Viajes":
                        mask_aplica = mask_base
                    elif tipo_dist == "Viajes con unidad":
                        mask_aplica = mask_base & has_unit_trip
                    elif tipo_dist == "Viajes con Remolque":
                        mask_aplica = mask_base & mask_trailer_lf
                    else:
                        # Por si se agrega otro tipo
                        mask_aplica = mask_base

                    idx = df_trips_work.index[mask_aplica]
                    n_traficos = len(idx)

                    if n_traficos == 0:
                        st.warning(
                            f"Para el cliente '{cliente}' y concepto '{concepto}' "
                            f"({tipo_dist}) no hay viajes que cumplan la condición. "
                            "No se asigna ese costo a ningún viaje."
                        )
                        continue

                    ci_por_viaje = monto_cliente / n_traficos
                    df_trips_work.loc[idx, "CI_op_ligado_operacion"] += ci_por_viaje

                # =======================
                # Total CI por viaje y descarga
                # =======================
                df_trips_work["CI_total"] = (
                    df_trips_work["CI_no_operativo"]
                    + df_trips_work["CI_op_ligado_operacion"]
                )

                st.subheader("Vista previa con CI asignado")
                st.dataframe(
                    df_trips_work.head(),
                    use_container_width=True,
                )

                def trips_to_excel_bytes(df):
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                        df.to_excel(
                            writer,
                            index=False,
                            sheet_name="CI_por_viaje",
                        )
                    return buffer.getvalue()

                st.download_button(
                    "📥 Descargar viajes con CI asignado (Excel)",
                    data=trips_to_excel_bytes(df_trips_work),
                    file_name="viajes_con_CI_asignado.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                )

        except Exception as e:
            st.error(f"Error procesando la base de viajes detallados: {e}")
