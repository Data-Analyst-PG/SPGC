import streamlit as st
import pandas as pd
from io import BytesIO
from supabase import create_client, Client
from datetime import date

# --- CONFIGURACIÓN SUPABASE ---
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
supabase = create_client(url, key)

st.title("🧾 Módulo 1: Prorrateo de Gastos Generales")

# ================================
# Subir archivo y generar resumen
# ================================

# Subir archivo Excel
uploaded_file = st.file_uploader("Sube el archivo con la hoja 'PASO 1'", type=["xlsx"])

if uploaded_file:
    try:
        # Leer hoja específica
        df = pd.read_excel(uploaded_file, sheet_name="PASO 1")
        df.columns = df.columns.str.strip().str.upper()
        st.session_state["df_original"] = df 

        # Asegurar nombres consistentes
        df.columns = df.columns.str.strip().str.upper()

        # Filtrar "GASTO GENERAL"
        gasto_general = df[df["SUCURSAL"] == "GASTO GENERAL"]

        # Agrupar por AREA/GASTO y sumar los CARGOS
        resumen = (
            gasto_general
            .groupby("AREA/CUENTA", as_index=False)["CARGOS"]
            .sum()
            .sort_values(by="CARGOS", ascending=False)
        )

        # Guardar en session_state para el módulo 2
        st.session_state['resumen'] = resumen

        st.success("Resumen generado con éxito.")
        st.dataframe(resumen, use_container_width=True)

        # Exportar a Excel
        def export_excel(df):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name="Resumen Gastos", index=False)
            return buffer.getvalue()

        st.download_button(
            "📥 Descargar resumen en Excel",
            data=export_excel(resumen),
            file_name="resumen_gastos_generales.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")

# ========================================
# Captura de Tráfico y Fecha por Sucursal
# ========================================

st.subheader("🚛 Captura de Tráfico por Sucursal")

# Lista fija de sucursales (precargadas)
sucursales = [
    "CAR-GAR", "CHICAGO", "CONSOLIDADO", "DALLAS", "GUADALAJARA",
    "LEON", "LINCOLN LOGISTICS", "MG HAULERS", "MONTERREY",
    "NUEVO LAREDO", "QUERETARO", "ROLANDO ALFARO"
]

# Fecha global que se aplicará a todos los registros
fecha_global = st.date_input("📅 Fecha de tráfico", value=date.today())

# Crear un DataFrame editable para capturar los números de tráfico
traficos_df = pd.DataFrame({
    "Sucursal": sucursales,
    "Tráfico": ["" for _ in sucursales]
})

st.markdown("### ✏️ Captura los números de tráfico")
edit_df = st.data_editor(
    traficos_df,
    use_container_width=True,
    num_rows="fixed",
    key="trafico_editor"
)

# Botón para guardar en Supabase
if st.button("💾 Guardar en Supabase"):
    try:
        # Validar que no haya vacíos
        if any(edit_df["Tráfico"] == ""):
            st.warning("Completa todos los números de tráfico antes de guardar.")
        else:
            # Agregar la fecha global
            edit_df["Fecha"] = fecha_global

            # Insertar en Supabase
            data = edit_df.to_dict(orient="records")
            res = supabase.table("viajes_distribucion").insert(data).execute()

            if res.data:
                st.success("✅ Tráficos guardados correctamente en Supabase.")
            else:
                st.warning("No se insertaron datos. Verifica la conexión o duplicados.")

    except Exception as e:
        st.error(f"Error al guardar en Supabase: {e}")

st.divider()
