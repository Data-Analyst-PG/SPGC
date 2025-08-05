import streamlit as st
import pandas as pd

st.title("🔄 Módulo 4: Prorrateo de Gastos Generales")

# Validar dependencias
if not all(k in st.session_state for k in ["resumen", "porcentajes"]):
    st.warning("Faltan datos. Asegúrate de haber completado los módulos 1 a 3.")
else:
    # Cargar catálogo desde Supabase
    from supabase import create_client

    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    supabase = create_client(url, key)

    data_supabase = supabase.table("catalogo_distribucion").select("*").execute().data
    catalogo_df = pd.DataFrame(data_supabase)

    resumen = st.session_state["resumen"]
    porcentajes = st.session_state["porcentajes"]  # con índice SUCURSAL

    # Unir resumen con catálogo
    gastos_tipo = resumen.merge(
        catalogo_df.rename(columns={
            "area_gasto": "AREA/GASTO",
            "tipo_distribucion": "TIPO DISTRIBUCIÓN"
        }),
        on="AREA/GASTO",
        how="left"
    )

    if gastos_tipo["TIPO DISTRIBUCIÓN"].isna().any():
        st.error("Hay gastos sin tipo de distribución asignado. Revisa el catálogo.")
    else:
        # Crear tabla final de prorrateo
        distribucion_rows = []

        for _, row in gastos_tipo.iterrows():
            area = row["AREA/GASTO"]
            tipo = row["TIPO DISTRIBUCIÓN"]
            total_cargo = row["CARGOS"]

            if tipo not in porcentajes.columns:
                st.warning(f"No hay datos de porcentaje para tipo: {tipo}")
                continue

            for sucursal, porcentaje in porcentajes[tipo].items():
                if porcentaje > 0:
                    distribucion_rows.append({
                        "AREA/GASTO": area,
                        "TIPO DISTRIBUCIÓN": tipo,
                        "SUCURSAL": sucursal,
                        "CARGO ASIGNADO": round(total_cargo * porcentaje, 2)
                    })

        prorrateo_df = pd.DataFrame(distribucion_rows)

        st.subheader("📊 Resultado del Prorrateo")
        st.dataframe(prorrateo_df, use_container_width=True)

        # Guardar para módulo final o exportación
        st.session_state["prorrateo"] = prorrateo_df

        # Descargar como Excel
        def to_excel(df):
            from io import BytesIO
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Prorrateo")
            return buffer.getvalue()

        st.download_button(
            "📥 Descargar prorrateo en Excel",
            data=to_excel(prorrateo_df),
            file_name="prorrateo_gastos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
