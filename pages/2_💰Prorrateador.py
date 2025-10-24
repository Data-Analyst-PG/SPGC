import streamlit as st
import pandas as pd
from io import BytesIO
from supabase import create_client, Client
from datetime import date

# --- CONFIGURACIÓN SUPABASE ---
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
supabase = create_client(url, key)

# ======================================================================================================
st.title("🧾Prorrateo de Gastos Generales")

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

        # Agrupar por AREA/CUENTA y sumar los CARGOS
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
if st.button("💾 Guardar en Supabase", key="save_traficos"):
    try:
        df_to_save = edit_df.copy()

        # Validaciones básicas
        faltan = df_to_save["Tráfico"].astype(str).str.strip().eq("")
        if faltan.any():
            st.warning("Completa todos los números de tráfico antes de guardar.")
        else:
            # Normaliza texto y FECHA -> string ISO (YYYY-MM-DD)
            df_to_save["Sucursal"] = df_to_save["Sucursal"].astype(str).str.strip()
            df_to_save["Tráfico"]  = df_to_save["Tráfico"].astype(str).str.strip()
            df_to_save["Fecha"]    = pd.to_datetime(fecha_global).strftime("%Y-%m-%d")

            # Renombra a las columnas reales de la tabla
            payload = (
                df_to_save[["Sucursal", "Tráfico", "Fecha"]]
                .rename(columns={"Tráfico": "Trafico"})
                .to_dict(orient="records")
            )

            # Inserta (o usa upsert si quieres evitar duplicados Sucursal+Fecha)
            res = supabase.table("viajes_distribucion").insert(payload).execute()
            # Si prefieres actualizar/enlazar:
            # res = supabase.table("viajes_distribucion") \
            #     .upsert(payload, on_conflict="Sucursal,Fecha").execute()

            st.success(f"✅ Tráficos guardados ({len(payload)} filas).")
    except Exception as e:
        st.error(f"Error al guardar en Supabase: {e}")

st.divider()

# ======================================================================================================
st.title("📘Catálogo de Distribución por AREA/CUENTA")

tipos_distribucion = [
    "Facturación Dlls", "MC", "Tráficos",
    "Empleado hub", "Empleados mv", "XTRALEASE", "Uso Cajas"
]

# Validamos que venga del Módulo 1
if 'resumen' in st.session_state:
    resumen = st.session_state['resumen']
    resumen = resumen[['AREA/CUENTA']].drop_duplicates().reset_index(drop=True)

    # Cargar catálogo existente desde Supabase
    data_supabase = supabase.table("catalogo_distribucion").select("*").execute().data
    catalogo_existente = pd.DataFrame(data_supabase)

    # Unir resumen con catálogo existente
    if not catalogo_existente.empty:
        catalogo_existente = catalogo_existente.rename(columns={
            "area_cuenta": "AREA/CUENTA",
            "tipo_distribucion": "TIPO DISTRIBUCIÓN"
        })
        resumen_merged = resumen.merge(catalogo_existente, on="AREA/CUENTA", how="left")
    else:
        resumen_merged = resumen.copy()
        resumen_merged["TIPO DISTRIBUCIÓN"] = None

    st.subheader("Catálogo de Distribución")
    resumen_merged = resumen_merged.sort_values(by=["TIPO DISTRIBUCIÓN", "AREA/CUENTA"], na_position="first").reset_index(drop=True)
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
    if st.button("💾 Guardar en Supabase", key="save_catalogo"):
        nuevos = edited_df[edited_df["TIPO DISTRIBUCIÓN"].notna()]
        for _, row in nuevos.iterrows():
            supabase.table("catalogo_distribucion").upsert({
                "area_cuenta": row["AREA/CUENTA"],
                "tipo_distribucion": row["TIPO DISTRIBUCIÓN"]
            }).execute()
        st.success("Catálogo actualizado en Supabase.")

else:
    st.warning("Primero genera el resumen de GASTO GENERAL en el Módulo 1.")

# ======================================================================================================
st.title("📊Datos Generales y Cálculo de Porcentajes")

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

# ======================================================================================================
st.title("🔄Gasto General + Costos por Sucursal (con Tráfico/Fecha)")

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
        return "COMUN INTERNO"
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

# ======================================================================================================
st.title("📘Generales (Comunes separados) e Indirectos")

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
        pivot_comunes[[col_suc, "COMUN INTERNO", "COMUN EXTERNO"]].to_excel(
            writer, sheet_name="Generales", index=False
        )
        indirectos.to_excel(writer, sheet_name="Indirectos", index=False)
        final[[col_suc, "COMUN INTERNO", "COMUN EXTERNO", "INDIRECTO", "TOTAL"]].to_excel(
            writer, sheet_name="Consolidado", index=False
        )
    return buffer.getvalue()

st.download_button(
    "📥 Descargar Excel (Generales/Indirectos/Consolidado)",
    data=exportar_excel(),
    file_name="generales_indirectos_consolidado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
