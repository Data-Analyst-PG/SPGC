import streamlit as st
import pandas as pd

st.title("📘 Módulo 5: Generales (Comunes separados) e Indirectos")

# --- Obtener el prorrateo completo del Módulo 4 ---
key_candidates = ["prorrateo_completo", "prorrateo"]
prorr_key = next((k for k in key_candidates if k in st.session_state), None)

if prorr_key is None or "df_original" not in st.session_state:
    st.warning("Faltan datos. Asegúrate de haber ejecutado el Módulo 4 y tener df_original.")
    st.stop()

prorr = st.session_state[prorr_key].copy()
df_original = st.session_state["df_original"].copy()

# Normalización mínima
for df in (prorr, df_original):
    df.columns = df.columns.str.upper()

col_suc = "SUCURSAL"
col_val = "CARGO ASIGNADO"
col_tipo = "TIPO COSTO"

# 1) Comunes por sucursal (ya vienen con nombres finales)
comunes = prorr[prorr[col_tipo].isin(["COMUN INTERNO", "COMUN EXTERNO"])]

pivot_comunes = (
    comunes.pivot_table(
        index=col_suc,
        columns=col_tipo,
        values=col_val,
        aggfunc="sum",
        fill_value=0.0,
    )
    .reset_index()
)

for expected in ["COMUN INTERNO", "COMUN EXTERNO"]:
    if expected not in pivot_comunes.columns:
        pivot_comunes[expected] = 0.0

st.subheader("📌 Comunes por sucursal")
st.dataframe(pivot_comunes[[col_suc, "COMUN INTERNO", "COMUN EXTERNO"]], use_container_width=True)

# ---------- 2) Indirectos por sucursal ----------
# Suma todo lo etiquetado como COSTO INDIRECTO (incluye directos y Gasto General cuyo concepto no inicia IN/EX)
indirectos = (
    prorr[prorr[col_tipo] == "COSTO INDIRECTO"]
    .groupby(col_suc, as_index=False)[col_val]
    .sum()
    .rename(columns={col_val: "INDIRECTO"})
)

st.subheader("📌 Indirectos por sucursal")
st.dataframe(indirectos, use_container_width=True)

# ---------- 3) Consolidado: comunes separados + indirecto + total ----------
final = (
    pivot_comunes.merge(indirectos, on=col_suc, how="outer")
    .fillna(0.0)
)

final["TOTAL"] = final["COMUN INTERNO"] + final["COMUN EXTERNO"] + final["INDIRECTO"]

st.subheader("📊 Consolidado final")
st.dataframe(final[[col_suc, "COMUN INTERNO", "COMUN EXTERNO", "INDIRECTO", "TOTAL"]],
             use_container_width=True)

# ---------- 4) Exportar a Excel ----------
def exportar_excel():
    from io import BytesIO
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        pivot_comunes[[col_s
