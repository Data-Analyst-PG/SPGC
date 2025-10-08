import streamlit as st
import pandas as pd

st.title("🔄 Módulo 4: Gasto General + Costos por Sucursal (con Tráfico/Fecha)")

# ---------- Comprobación de insumos previos ----------
if not all(k in st.session_state for k in ["df_original", "resumen", "porcentajes"]):
    st.warning("Faltan datos. Completa primero los módulos previos (df_original, resumen, porcentajes).")
    st.stop()

df_original = st.session_state["df_original"].copy()
resumen_gasto_general = st.session_state["resumen"].copy()  # agrupado por AREA/CUENTA
porcentajes = st.session_state["porcentajes"].copy()

# Normalizaciones
df_original.columns = df_original.columns.str.strip().str.upper()
porcentajes.columns = [str(c).upper() for c in porcentajes.columns]
if "CONCEPTO" not in df_original.columns:
    df_original["CONCEPTO"] = ""

# ---------- Supabase: catálogo y viajes  ----------
from supabase import create_client
url = st.secrets["supabase"]["url"]
key = st.secrets["supabase"]["key"]
supabase = create_client(url, key)

# Catálogo (AREA/CUENTA -> TIPO DISTRIBUCIÓN)
cat_resp = supabase.table("catalogo_distribucion").select("*").execute()
catalogo = pd.DataFrame(cat_resp.data)
if catalogo.empty:
    st.error("No hay datos en 'catalogo_distribucion'.")
    st.stop()
catalogo = catalogo.rename(columns={
    "area_cuenta": "AREA/CUENTA",
    "tipo_distribucion": "TIPO DISTRIBUCIÓN"
})

# Viajes/fechas (para seleccionar fecha específica)
viajes_resp = supabase.table("viajes_distribucion").select("*").execute()
viajes = pd.DataFrame(viajes_resp.data)
if viajes.empty:
    st.info("No hay registros en 'viajes_distribucion'. Trafico/Fecha quedarán vacíos.")
    viajes = pd.DataFrame(columns=["Sucursal", "Trafico", "Fecha"])
viajes = viajes.rename(columns={"Sucursal": "SUCURSAL", "Trafico": "TRAFICO", "Fecha": "FECHA"})
viajes["SUCURSAL"] = viajes["SUCURSAL"].astype(str).str.upper().str.strip()
viajes["FECHA"] = pd.to_datetime(viajes["FECHA"], errors="coerce")

# Selector de fecha específica (usa las disponibles en la tabla)
fechas_disponibles = sorted(viajes["FECHA"].dropna().dt.date.unique())
fecha_elegida = st.date_input(
    "📅 Selecciona la FECHA de viajes_distribucion a usar",
    value=(fechas_disponibles[-1] if len(fechas_disponibles) else pd.Timestamp.today().date()),
    min_value=(fechas_disponibles[0] if len(fechas_disponibles) else pd.Timestamp.today().date())
)

viajes_sel = viajes[viajes["FECHA"].dt.date.eq(fecha_elegida)].copy()
viajes_sel["FECHA"] = viajes_sel["FECHA"].dt.strftime("%Y-%m-%d")  # ISO string
viajes_sel = viajes_sel.drop_duplicates(subset=["SUCURSAL"], keep="last")

# Helper para anexar Trafico/Fecha por sucursal
def anexar_trafico_fecha(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.assign(TRAFICO=None, FECHA=None)
    out = df.merge(viajes_sel[["SUCURSAL", "TRAFICO", "FECHA"]], on="SUCURSAL", how="left")
    return out

# ============ BLOQUE A: COSTOS CON SUCURSAL ASIGNADA ============
no_general = df_original[df_original["SUCURSAL"].str.upper().ne("GASTO GENERAL")].copy()

# TIPO COSTO fijo para este bloque
no_general["TIPO COSTO"] = "COSTO INDIRECTO"
# TIPO DISTRIBUCIÓN fijo para este bloque
no_general["TIPO DISTRIBUCIÓN"] = "Costo fijo en sucursal"

# Agrupar por AREA/CUENTA + SUCURSAL + TIPO DISTRIBUCIÓN + TIPO COSTO
directos_agr = (
    no_general
    .assign(SUCURSAL=lambda d: d["SUCURSAL"].astype(str).str.upper().str.strip())
    .groupby(["AREA/CUENTA", "SUCURSAL", "TIPO DISTRIBUCIÓN", "TIPO COSTO"], as_index=False)["CARGOS"]
    .sum()
    .rename(columns={"CARGOS": "CARGO ASIGNADO"})
)

# Anexar Trafico/Fecha por sucursal (fecha seleccionada)
directos_agr = anexar_trafico_fecha(directos_agr)

# ============ BLOQUE B: GASTO GENERAL (prorrateo) ============
gasto_general = df_original[df_original["SUCURSAL"].str.upper().eq("GASTO GENERAL")].copy()

# TIPO COSTO por regla de CONCEPTO (prefijos)
def tipo_costo_por_concepto(concepto: str) -> str:
    c = str(concepto).strip().upper()
    if c.startswith("IN"):
        return "COMUN INDIRECTO"
    if c.startswith("EX"):
        return "COMUN EXTERNO"
    return "COSTO INDIRECTO"

gasto_general["TIPO COSTO"] = gasto_general["CONCEPTO"].map(tipo_costo_por_concepto)

# Agrupar por AREA/CUENTA + TIPO COSTO (para preservar la etiqueta)
gg_agr = (
    gasto_general
    .groupby(["AREA/CUENTA", "TIPO COSTO"], as_index=False)["CARGOS"]
    .sum()
    .rename(columns={"CARGOS": "TOTAL_AREA"})
)

# Unir con catálogo para obtener TIPO DISTRIBUCIÓN
gg_agr = gg_agr.merge(catalogo, on="AREA/CUENTA", how="left")
if gg_agr["TIPO DISTRIBUCIÓN"].isna().any():
    faltantes = gg_agr.loc[gg_agr["TIPO DISTRIBUCIÓN"].isna(), "AREA/CUENTA"].unique().tolist()
    st.error(f"Faltan tipos de distribución en el catálogo para: {faltantes[:10]}{'...' if len(faltantes)>10 else ''}")
    st.stop()

# Expandir por sucursales según porcentajes del tipo
prorr_rows = []
for _, r in gg_agr.iterrows():
    area = r["AREA/CUENTA"]
    tipo_dist = str(r["TIPO DISTRIBUCIÓN"]).upper()
    tipo_costo = r["TIPO COSTO"]
    total = r["TOTAL_AREA"]

    if tipo_dist not in porcentajes.columns:
        st.warning(f"No hay porcentajes para el tipo '{tipo_dist}'. Se omite AREA/CUENTA: {area}")
        continue

    for suc, pct in porcentajes[tipo_dist].items():
        if pct and pct > 0:
            prorr_rows.append({
                "AREA/CUENTA": area,
                "SUCURSAL": str(suc).upper().strip(),
                "TIPO DISTRIBUCIÓN": r["TIPO DISTRIBUCIÓN"],  # conservar como está en catálogo (no upper si no quieres)
                "TIPO COSTO": tipo_costo,
                "CARGO ASIGNADO": round(total * float(pct), 2)
            })

prorr_gg = pd.DataFrame(prorr_rows)
prorr_gg = anexar_trafico_fecha(prorr_gg)

# ============ RESULTADO FINAL ============
resultado = pd.concat([directos_agr, prorr_gg], ignore_index=True)

st.subheader("📊 Resultado final (con Tráfico/Fecha por sucursal y fecha seleccionada)")
st.dataframe(resultado, use_container_width=True)

# Guardar para módulos siguientes
st.session_state["prorrateo_completo"] = resultado

# Descargar
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    from io import BytesIO
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as xw:
        df.to_excel(xw, index=False, sheet_name="Prorrateo")
    return buffer.getvalue()

st.download_button(
    "📥 Descargar prorrateo completo",
    data=to_excel_bytes(resultado),
    file_name="prorrateo_completo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
