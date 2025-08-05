import streamlit as st
import pandas as pd

st.title("📘 Módulo 2: Catálogo de Distribución por AREA/GASTO")

# Tipos de distribución válidos
tipos_distribucion = [
    "Facturación Dlls", "MC", "Tráficos",
    "Empleado hub", "Empleados mv", "XTRALEASE", "Uso Cajas"
]

# Si ya existe un resumen de gastos generales generado en el Módulo 1
if 'resumen' in st.session_state:
    resumen = st.session_state['resumen']

    st.subheader("Catálogo de Distribución")

    # Crear tabla editable con columna para asignar tipo de distribución
    if 'catalogo' not in st.session_state:
        catalogo_df = resumen.copy()
        catalogo_df["TIPO DISTRIBUCIÓN"] = ""
        st.session_state['catalogo'] = catalogo_df
    else:
        catalogo_df = st.session_state['catalogo']

    edited_catalogo = st.data_editor(
        catalogo_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "TIPO DISTRIBUCIÓN": st.column_config.SelectboxColumn(
                label="Tipo de Distribución",
                options=tipos_distribucion,
                required=True
            )
        }
    )

    # Guardar cambios temporalmente
    st.session_state['catalogo'] = edited_catalogo

    # Descargar catálogo como Excel
    def to_excel(df):
        from io import BytesIO
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Catálogo")
        return buffer.getvalue()

    st.download_button(
        "📥 Descargar catálogo en Excel",
        data=to_excel(edited_catalogo),
        file_name="catalogo_distribucion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("Primero genera el resumen de GASTO GENERAL en el Módulo 1.")
