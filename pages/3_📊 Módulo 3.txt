import streamlit as st
import pandas as pd

st.title("📊 Módulo 3: Datos Generales y Cálculo de Porcentajes")

# Subida del archivo con la hoja GTS
archivo_gts = st.file_uploader("📤 Sube el archivo con la hoja 'GTS'", type=["xlsx"])

if archivo_gts:
    try:
        df_gts = pd.read_excel(archivo_gts, sheet_name="GTS")

        # Limpiar encabezados
        df_gts.columns = df_gts.columns.str.strip().str.upper()

        # Mostrar tabla original
        st.subheader("📄 Datos GTS cargados:")
        st.dataframe(df_gts, use_container_width=True)

        # Seleccionar solo columnas numéricas (los tipos de distribución)
        tipo_distrib_cols = df_gts.columns.drop("SUCURSAL")

        # Calcular totales por tipo
        totales = df_gts[tipo_distrib_cols].sum().rename("TOTAL").to_frame().T

        st.subheader("📌 Totales por Tipo de Distribución")
        st.dataframe(totales, use_container_width=True)

        # Calcular porcentajes por sucursal y tipo
        porcentajes = df_gts.copy()
        for col in tipo_distrib_cols:
            total = df_gts[col].sum()
            if total != 0:
                porcentajes[col] = df_gts[col] / total
            else:
                porcentajes[col] = 0

        st.subheader("📈 Porcentaje de Participación")
        st.dataframe(porcentajes, use_container_width=True)

        # Guardar en memoria para módulo 4
        st.session_state["porcentajes"] = porcentajes.set_index("SUCURSAL")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar la hoja GTS: {e}")
