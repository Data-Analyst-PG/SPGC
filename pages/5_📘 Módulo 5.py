import streamlit as st
import pandas as pd

st.title("📘 Módulo 5: Generales e Indirectos")

# Validar inputs
if "prorrateo" not in st.session_state or "df_original" not in st.session_state:
    st.warning("Faltan datos. Asegúrate de haber ejecutado los módulos 1 y 4.")
else:
    prorrateo_df = st.session_state["prorrateo"]
    df_original = st.session_state["df_original"]

    # Gasto General (suma de lo prorrateado por sucursal)
    gasto_general = (
        prorrateo_df
        .groupby("SUCURSAL", as_index=False)["CARGO ASIGNADO"]
        .sum()
        .rename(columns={"CARGO ASIGNADO": "GASTO GENERAL"})
    )

    st.subheader("📌 Gasto General por Sucursal (prorrateado)")
    st.dataframe(gasto_general, use_container_width=True)

    # Indirectos (suma directa de cargos que NO son GASTO GENERAL)
    df_original.columns = df_original.columns.str.upper()
    indirectos = (
        df_original[df_original["SUCURSAL"] != "GASTO GENERAL"]
        .groupby("SUCURSAL", as_index=False)["CARGOS"]
        .sum()
        .rename(columns={"CARGOS": "INDIRECTOS"})
    )

    st.subheader("📌 Indirectos por Sucursal (del archivo original)")
    st.dataframe(indirectos, use_container_width=True)

    # Unir ambas tablas para una vista consolidada
    final = pd.merge(gasto_general, indirectos, on="SUCURSAL", how="outer").fillna(0)
    final["TOTAL"] = final["GASTO GENERAL"] + final["INDIRECTOS"]

    st.subheader("📊 Consolidado Final")
    st.dataframe(final, use_container_width=True)

    # Descargar
    def exportar_excel():
        from io import BytesIO
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            gasto_general.to_excel(writer, sheet_name="Gasto General", index=False)
            indirectos.to_excel(writer, sheet_name="Indirectos", index=False)
            final.to_excel(writer, sheet_name="Consolidado", index=False)
        return buffer.getvalue()

    st.download_button(
        "📥 Descargar Excel con resultados",
        data=exportar_excel(),
        file_name="generales_y_indirectos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
