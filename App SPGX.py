import streamlit as st
import pandas as pd
from io import BytesIO
from supabase import create_client, Client

# Conexión a Supabase
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
supabase: Client = create_client(url, key)

# Tipos de distribución válidos
TIPOS_DISTRIBUCION = [
    "FACTURACION DLLS", "MC", "TRAFICOS",
    "EMPLEADO HUB", "EMPLEADO MV", "XTRALEASE", "USO CAJAS"
]

# Activar navegación superior
st.set_page_config(page_title="Prorrateo de Costos", layout="wide")
st.navigation(label="Menú", links=[
    "Home",
    "Módulo 1: Subir archivo",
    "Módulo 2: Catálogo",
    "Módulo 3: GTS",
    "Módulo 4: Prorrateo",
    "Módulo 5: Generales"
], position="top")

# Ruta activa
pagina = st.get_page_location()

# ---------------------- HOME ----------------------
if pagina == "Home":
    st.title("🏠 Bienvenido al sistema de prorrateo")
    st.markdown("""
    Esta app te permite automatizar el prorrateo de gastos generales por sucursal:

    1. Sube el archivo de gastos.
    2. Define el tipo de distribución por *AREA/GASTO*.
    3. Carga los datos generales mensuales (GTS).
    4. Obtén los prorrateos y reportes.
    """)

# ---------------------- MODULO 1 ----------------------
elif pagina == "Módulo 1: Subir archivo":
    st.title("📄 Módulo 1: Subir archivo y generar resumen")
    archivo = st.file_uploader("Sube el archivo con la hoja 'DATA MAYO-OCTUBRE'", type=["xlsx"])

    if archivo:
        df = pd.read_excel(archivo, sheet_name="DATA MAYO-OCTUBRE")
        df.columns = df.columns.str.strip().str.upper()
        st.session_state["df_original"] = df

        gastos = df[df["SUCURSAL"] == "GASTO GENERAL"]
        resumen = gastos.groupby("AREA/GASTO", as_index=False)["CARGOS"].sum().sort_values(by="CARGOS", ascending=False)
        st.session_state["resumen"] = resumen

        st.success("Resumen generado con éxito.")
        st.dataframe(resumen, use_container_width=True)

# ---------------------- MODULO 2 ----------------------
elif pagina == "Módulo 2: Catálogo":
    st.title("📘 Módulo 2: Catálogo de Distribución por AREA/GASTO")
    if "resumen" in st.session_state:
        resumen = st.session_state["resumen"]["AREA/GASTO"].drop_duplicates().to_frame()

        data_supabase = supabase.table("catalogo_distribucion").select("*").execute().data
        catalogo_existente = pd.DataFrame(data_supabase)

        if not catalogo_existente.empty:
            catalogo_existente = catalogo_existente.rename(columns={"area_gasto": "AREA/GASTO", "tipo_distribucion": "TIPO DISTRIBUCIÓN"})
            resumen = resumen.merge(catalogo_existente, on="AREA/GASTO", how="left")
        else:
            resumen["TIPO DISTRIBUCIÓN"] = None

        resumen = resumen.sort_values(by=["TIPO DISTRIBUCIÓN", "AREA/GASTO"], na_position="first").reset_index(drop=True)
        edited_df = st.data_editor(
            resumen,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "TIPO DISTRIBUCIÓN": st.column_config.SelectboxColumn(
                    label="Tipo de Distribución", options=TIPOS_DISTRIBUCION, required=True
                )
            }
        )

        if st.button("💾 Guardar en Supabase"):
            for _, row in edited_df.iterrows():
                if row["TIPO DISTRIBUCIÓN"]:
                    supabase.table("catalogo_distribucion").upsert({
                        "area_gasto": row["AREA/GASTO"],
                        "tipo_distribucion": row["TIPO DISTRIBUCIÓN"]
                    }).execute()
            st.success("Catálogo actualizado en Supabase.")
    else:
        st.warning("Debes ejecutar primero el Módulo 1 para ver el resumen.")

# ---------------------- MODULO 3 ----------------------
elif pagina == "Módulo 3: GTS":
    st.title("📊 Módulo 3: Datos Generales y Porcentajes")
    archivo_gts = st.file_uploader("Sube el archivo con la hoja 'GTS'", type=["xlsx"])

    if archivo_gts:
        xls = pd.ExcelFile(archivo_gts)
        hoja_gts = next((h for h in xls.sheet_names if 'gts' in h.lower()), None)

        if hoja_gts is None:
            st.error("No se encontró una hoja llamada 'GTS'")
        else:
            df_gts = pd.read_excel(xls, sheet_name=hoja_gts)
            df_gts.columns = df_gts.columns.str.upper()
            st.dataframe(df_gts, use_container_width=True)

            cols = df_gts.columns.drop("SUCURSAL")
            porcentajes = df_gts.copy()
            for col in cols:
                total = df_gts[col].sum()
                porcentajes[col] = df_gts[col] / total if total != 0 else 0

            porcentajes.columns = [col.upper() for col in porcentajes.columns]
            st.session_state["porcentajes"] = porcentajes.set_index("SUCURSAL")
            st.success("Porcentajes calculados y guardados.")
            st.dataframe(st.session_state["porcentajes"], use_container_width=True)

# ---------------------- MODULO 4 ----------------------
elif pagina == "Módulo 4: Prorrateo":
    st.title("🔄 Módulo 4: Prorrateo de Gastos Generales")
    if "resumen" not in st.session_state or "porcentajes" not in st.session_state:
        st.warning("Faltan datos. Ejecuta los módulos 1 y 3.")
    else:
        resumen = st.session_state["resumen"]
        porcentajes = st.session_state["porcentajes"]

        data_supabase = supabase.table("catalogo_distribucion").select("*").execute().data
        catalogo = pd.DataFrame(data_supabase)
        gastos_tipo = resumen.merge(
            catalogo.rename(columns={"area_gasto": "AREA/GASTO", "tipo_distribucion": "TIPO DISTRIBUCIÓN"}),
            on="AREA/GASTO", how="left"
        )

        porcentajes.columns = [col.upper() for col in porcentajes.columns]
        filas = []

        for _, row in gastos_tipo.iterrows():
            tipo = row["TIPO DISTRIBUCIÓN"]
            tipo_col = tipo.upper()
            if tipo_col not in porcentajes.columns:
                st.warning(f"No hay datos de porcentaje para tipo: {tipo}")
                continue
            for suc, pct in porcentajes[tipo_col].items():
                if pct > 0:
                    filas.append({
                        "AREA/GASTO": row["AREA/GASTO"],
                        "TIPO DISTRIBUCIÓN": tipo,
                        "SUCURSAL": suc,
                        "CARGO ASIGNADO": round(row["CARGOS"] * pct, 2)
                    })

        df_prorrateo = pd.DataFrame(filas)
        st.session_state["prorrateo"] = df_prorrateo
        st.dataframe(df_prorrateo, use_container_width=True)

# ---------------------- MODULO 5 ----------------------
elif pagina == "Módulo 5: Generales":
    st.title("📘 Módulo 5: Generales e Indirectos")
    if "df_original" not in st.session_state or "prorrateo" not in st.session_state:
        st.warning("Faltan datos. Ejecuta los módulos 1 y 4.")
    else:
        df_o = st.session_state["df_original"]
        df_o.columns = df_o.columns.str.upper()
        prr = st.session_state["prorrateo"]

        gasto_general = prr.groupby("SUCURSAL", as_index=False)["CARGO ASIGNADO"].sum().rename(columns={"CARGO ASIGNADO": "GASTO GENERAL"})
        indirectos = df_o[df_o["SUCURSAL"] != "GASTO GENERAL"].groupby("SUCURSAL", as_index=False)["CARGOS"].sum().rename(columns={"CARGOS": "INDIRECTOS"})
        total = pd.merge(gasto_general, indirectos, on="SUCURSAL", how="outer").fillna(0)
        total["TOTAL"] = total["GASTO GENERAL"] + total["INDIRECTOS"]

        st.subheader("Consolidado Final")
        st.dataframe(total, use_container_width=True)

        def to_excel():
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                gasto_general.to_excel(writer, sheet_name="Gasto General", index=False)
                indirectos.to_excel(writer, sheet_name="Indirectos", index=False)
                total.to_excel(writer, sheet_name="Consolidado", index=False)
            return buffer.getvalue()

        st.download_button("📥 Descargar Excel Consolidado", to_excel(), file_name="generales_indirectos.xlsx")
