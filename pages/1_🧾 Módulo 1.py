import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🧾 Módulo 1: Prorrateo de Gastos Generales")

# Subir archivo Excel
uploaded_file = st.file_uploader("Sube el archivo con la hoja 'DATA MAYO-OCTUBRE'", type=["xlsx"])

if uploaded_file:
    try:
        # Leer hoja específica
        df = pd.read_excel(uploaded_file, sheet_name="DATA MAYO-OCTUBRE")
        df.columns = df.columns.str.strip().str.upper()
        st.session_state["df_original"] = df 

        # Asegurar nombres consistentes
        df.columns = df.columns.str.strip().str.upper()

        # Filtrar "GASTO GENERAL"
        gasto_general = df[df["SUCURSAL"] == "GASTO GENERAL"]

        # Agrupar por AREA/GASTO y sumar los CARGOS
        resumen = (
            gasto_general
            .groupby("AREA/GASTO", as_index=False)["CARGOS"]
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
