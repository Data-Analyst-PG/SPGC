import streamlit as st
import pandas as pd
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
2. Repartir costos **no ligados a la operación** entre viajes con/sin unidad y obtener el costo unitario.
3. Prorratear costos **ligados indirectamente a la operación** entre clientes usando un catálogo de tipos de distribución.
"""
)

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
            # Se asume que la fila 'Millas' está en la col 'Mes'
            try:
                fila_millas = dist_df.loc[dist_df["Mes"].astype(str).str.upper() == "MILLAS"]
                # Mapear número de mes a nombre de columna
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
                if total_col:
                    por_tipo[f"%{col}"] = por_tipo[col] / total_col
                else:
                    por_tipo[f"%{col}"] = 0.0

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

            # Seleccionar columna de mes
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

            st.subheader("Costos unitarios derivados")
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

            # Guardar en estado por si luego quieres usarlo para asignar por cliente
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

# Cargar archivo de costos operativos (para nuevos conceptos)
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

            # 🔹 NUEVO: si viene como lista ["Volumen Viajes"], tomar solo el primer valor
            catalogo_existente["Tipo distribución"] = catalogo_existente["Tipo distribución"].apply(
                lambda x: x[0] if isinstance(x, list) and len(x) > 0 else x
            )

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
                                "tipo_distribucion": row["Tipo distribución"],
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

        # 🔹 NUEVO: aplanar listas como ["Volumen Viajes"] -> "Volumen Viajes"
        catalogo["Tipo distribución"] = catalogo["Tipo distribución"].apply(
            lambda x: x[0] if isinstance(x, list) and len(x) > 0 else x
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

                # 🔹 Por si en df_op_mes todavía quedara alguna lista
                if isinstance(tipo_dist, list) and len(tipo_dist) > 0:
                    tipo_dist = tipo_dist[0]

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

                # Exportar resultados
                def to_excel_bytes(df1, df2):
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
                    return buffer.getvalue()

                st.download_button(
                    "📥 Descargar resultados (Excel)",
                    data=to_excel_bytes(asignaciones_df, pivot_clientes),
                    file_name="prorrateo_costos_clientes.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                )
