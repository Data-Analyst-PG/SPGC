import streamlit as st
import pandas as pd
from supabase import create_client, Client

# Conexión a Supabase
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
supabase: Client = create_client(url, key)

st.title("📘 Módulo 2: Catálogo de Distribución por AREA/GASTO")

tipos_distribucion = [
    "Facturación Dlls", "MC", "Tráficos",
    "Empleado hub", "Empleados mv", "XTRALEASE", "Uso Cajas"
]

# Validamos que venga del Módulo 1
if 'resumen' in st.session_state:
    resumen = st.session_state['resumen']
    resumen = resumen[['AREA/GASTO']].drop_duplicates().reset_index(drop=True)

    # Cargar catálogo existente desde Supabase
    data_supabase = supabase.table("catalogo_distribucion").select("*").execute().data
    catalogo_existente = pd.DataFrame(data_supabase)

    # Unir resumen con catálogo existente
    if not catalogo_existente.empty:
        catalogo_existente = catalogo_existente.rename(columns={
            "area_gasto": "AREA/GASTO",
            "tipo_distribucion": "TIPO DISTRIBUCIÓN"
        })
        resumen_merged = resumen.merge(catalogo_existente, on="AREA/GASTO", how="left")
    else:
        resumen_merged = resumen.copy()
        resumen_merged["TIPO DISTRIBUCIÓN"] = None

    st.subheader("Catálogo de Distribución")
    resumen_merged = resumen_merged.sort_values(by=["TIPO DISTRIBUCIÓN", "AREA/GASTO"], na_position="first").reset_index(drop=True)
    edited_df = st.data_editor(
        resumen_merged,
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

    # Botón para guardar nuevos registros y actualizaciones en Supabase
    if st.button("💾 Guardar en Supabase"):
        nuevos = edited_df[edited_df["TIPO DISTRIBUCIÓN"].notna()]
        for _, row in nuevos.iterrows():
            supabase.table("catalogo_distribucion").upsert({
                "area_gasto": row["AREA/GASTO"],
                "tipo_distribucion": row["TIPO DISTRIBUCIÓN"]
            }).execute()
        st.success("Catálogo actualizado en Supabase.")

else:
    st.warning("Primero genera el resumen de GASTO GENERAL en el Módulo 1.")
