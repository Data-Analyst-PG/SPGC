import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import date

st.set_page_config(page_title="Prorrateador", layout="wide")

# =========================
# CACHE SUPABASE
# =========================
@st.cache_resource
def get_supabase():
    from supabase import create_client
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

@st.cache_data(ttl=600)
def fetch_catalogo():
    supabase = get_supabase()
    data = supabase.table("catalogo_distribucion").select("*").execute().data
    return pd.DataFrame(data)

@st.cache_data(ttl=600)
def fetch_viajes():
    supabase = get_supabase()
    data = supabase.table("viajes_distribucion").select("*").execute().data
    return pd.DataFrame(data)

@st.cache_data
def read_excel_sheet(file_bytes: bytes, sheet_name: str):
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)
# =========================
# PASO 1: Upload PASO 1 + Resumen
# =========================
def paso1_upload():
    st.header("🧾 Paso 1 - Subir PASO 1 y generar resumen")

    uploaded_file = st.file_uploader("Sube el archivo con la hoja 'PASO 1'", type=["xlsx"], key="paso1_file")
    if not uploaded_file:
        st.info("Sube el archivo para continuar.")
        return

    df = read_excel_sheet(uploaded_file.getvalue(), "PASO 1")
    df.columns = df.columns.str.strip().str.upper()

    required = {"SUCURSAL", "AREA/CUENTA", "CARGOS"}
    missing = required - set(df.columns)
    if missing:
        st.error(f"❌ Faltan columnas requeridas en PASO 1: {sorted(missing)}")
        return

    df["SUCURSAL"] = df["SUCURSAL"].astype(str).str.strip().str.upper()
    comunes_base = df[df["SUCURSAL"].isin(["GASTO GENERAL", "INTERNO", "EXTERNO"])].copy()

    if comunes_base.empty:
        st.error("❌ No hay filas con SUCURSAL = GASTO GENERAL / INTERNO / EXTERNO.")
        return

    resumen = (
        comunes_base.groupby("AREA/CUENTA", as_index=False)["CARGOS"]
        .sum()
        .sort_values("CARGOS", ascending=False)
    )

    st.session_state["df_original"] = df
    st.session_state["resumen"] = resumen

    st.success("✅ Resumen generado con éxito.")
    st.dataframe(resumen, width="stretch")


# =========================
# PASO 2: Tráfico (con FORM para evitar rerun por cada edición)
# =========================
def paso2_trafico():
    st.header("🚛 Paso 2 - Captura de Tráfico por Sucursal")

    sucursales = [
        "CAR-GAR", "CHICAGO", "CONSOLIDADO", "DALLAS", "GUADALAJARA",
        "LEON", "LINCOLN LOGISTICS", "MG HAULERS", "MONTERREY",
        "NUEVO LAREDO", "QUERETARO", "ROLANDO ALFARO"
    ]

    fecha_global = st.date_input("📅 Fecha de tráfico", value=date.today(), key="traf_fecha")

    traficos_df = pd.DataFrame({"Sucursal": sucursales, "Tráfico": ["" for _ in sucursales]})

    with st.form("form_trafico"):
        edited = st.data_editor(traficos_df, num_rows="fixed", width="stretch")
        submitted = st.form_submit_button("💾 Guardar en Supabase")

    if not submitted:
        return

    df_to_save = edited.copy()
    faltan = df_to_save["Tráfico"].astype(str).str.strip().eq("")
    if faltan.any():
        st.warning("Completa todos los números de tráfico antes de guardar.")
        return

    df_to_save["Sucursal"] = df_to_save["Sucursal"].astype(str).str.strip()
    df_to_save["Tráfico"]  = df_to_save["Tráfico"].astype(str).str.strip()
    df_to_save["Fecha"]    = pd.to_datetime(fecha_global).strftime("%Y-%m-%d")

    payload = (
        df_to_save[["Sucursal", "Tráfico", "Fecha"]]
        .rename(columns={"Tráfico": "Trafico"})
        .to_dict(orient="records")
    )

    supabase = get_supabase()
    supabase.table("viajes_distribucion").insert(payload).execute()

    fetch_viajes.clear()
    st.success(f"✅ Tráficos guardados ({len(payload)} filas).")


# =========================
# PASO 3: Catálogo (con FORM)
# =========================
def paso3_catalogo():
    st.header("📘 Paso 3 - Catálogo de Distribución por AREA/CUENTA")

    if "resumen" not in st.session_state:
        st.warning("Primero genera el resumen en el Paso 1.")
        return

    tipos_distribucion = [
        "Facturación Dlls", "MC", "Tráficos",
        "Empleado hub", "Empleados mv", "XTRALEASE", "Uso Cajas"
    ]

    resumen = st.session_state["resumen"][["AREA/CUENTA"]].drop_duplicates().reset_index(drop=True)

    catalogo_existente = fetch_catalogo()
    if not catalogo_existente.empty:
        catalogo_existente = catalogo_existente.rename(columns={
            "area_cuenta": "AREA/CUENTA",
            "tipo_distribucion": "TIPO DISTRIBUCIÓN"
        })
        merged = resumen.merge(catalogo_existente, on="AREA/CUENTA", how="left")
    else:
        merged = resumen.copy()
        merged["TIPO DISTRIBUCIÓN"] = None

    with st.form("form_catalogo"):
        edited = st.data_editor(
            merged,
            num_rows="dynamic",
            width="stretch",
            column_config={
                "TIPO DISTRIBUCIÓN": st.column_config.SelectboxColumn(
                    "Tipo de Distribución",
                    options=tipos_distribucion,
                    required=True
                )
            }
        )
        submitted = st.form_submit_button("💾 Guardar catálogo")

    if not submitted:
        return

    supabase = get_supabase()
    nuevos = edited[edited["TIPO DISTRIBUCIÓN"].notna()]
    for _, row in nuevos.iterrows():
        supabase.table("catalogo_distribucion").upsert({
            "area_cuenta": row["AREA/CUENTA"],
            "tipo_distribucion": row["TIPO DISTRIBUCIÓN"]
        }).execute()

    fetch_catalogo.clear()
    st.success("✅ Catálogo actualizado.")


# =========================
# PASO 4: GTS + porcentajes
# =========================
def paso4_gts():
    st.header("📊 Paso 4 - Subir GTS y calcular porcentajes")

    archivo_gts = st.file_uploader("📤 Sube el archivo con la hoja 'GTS'", type=["xlsx"], key="gts_file")
    if not archivo_gts:
        return

    df_gts = read_excel_sheet(archivo_gts.getvalue(), "GTS")
    df_gts.columns = df_gts.columns.str.strip().str.upper()

    if "SUCURSAL" not in df_gts.columns:
        st.error("❌ La hoja GTS debe tener la columna SUCURSAL.")
        return

    st.dataframe(df_gts, width="stretch")

    tipo_cols = [c for c in df_gts.columns if c != "SUCURSAL"]
    porcentajes = df_gts.copy()
    for col in tipo_cols:
        total = df_gts[col].sum()
        porcentajes[col] = (df_gts[col] / total) if total != 0 else 0

    st.session_state["df_gts"] = df_gts
    st.session_state["porcentajes"] = porcentajes.assign(
        SUCURSAL=lambda d: d["SUCURSAL"].astype(str).str.strip().str.upper()
    ).set_index("SUCURSAL")

    st.success("✅ Porcentajes calculados.")
    st.dataframe(st.session_state["porcentajes"].reset_index(), width="stretch")

@st.cache_data
def calcular_prorrateo_cached(df_original: pd.DataFrame,
                             porcentajes: pd.DataFrame,
                             catalogo: pd.DataFrame,
                             viajes: pd.DataFrame,
                             fecha_elegida: date) -> pd.DataFrame:
    # Normalizaciones
    df_original = df_original.copy()
    porcentajes = porcentajes.copy()

    df_original.columns = df_original.columns.str.strip().str.upper()
    if "CONCEPTO" not in df_original.columns:
        df_original["CONCEPTO"] = ""

    df_original["SUCURSAL"] = df_original["SUCURSAL"].astype(str).str.strip().str.upper()
    porcentajes.columns = [str(c).upper() for c in porcentajes.columns]

    # Catalogo
    catalogo = catalogo.copy()
    if catalogo.empty:
        raise ValueError("No hay datos en 'catalogo_distribucion'.")

    catalogo = catalogo.rename(columns={
        "area_cuenta": "AREA/CUENTA",
        "tipo_distribucion": "TIPO DISTRIBUCIÓN"
    })
    catalogo["AREA/CUENTA"] = catalogo["AREA/CUENTA"].astype(str).str.strip().str.upper()

    # Viajes
    viajes = viajes.copy()
    if viajes.empty:
        viajes = pd.DataFrame(columns=["Sucursal", "Trafico", "Fecha"])

    viajes = viajes.rename(columns={"Sucursal": "SUCURSAL", "Trafico": "TRAFICO", "Fecha": "FECHA"})
    viajes["SUCURSAL"] = viajes["SUCURSAL"].astype(str).str.upper().str.strip()
    viajes["FECHA"] = pd.to_datetime(viajes["FECHA"], errors="coerce")

    viajes_sel = viajes[viajes["FECHA"].dt.date.eq(fecha_elegida)].copy()
    viajes_sel["FECHA"] = viajes_sel["FECHA"].dt.strftime("%Y-%m-%d")
    viajes_sel = viajes_sel.drop_duplicates(subset=["SUCURSAL"], keep="last")

    def anexar_trafico_fecha(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df.assign(TRAFICO=None, FECHA=None)
        return df.merge(viajes_sel[["SUCURSAL", "TRAFICO", "FECHA"]], on="SUCURSAL", how="left")

    # ===== BLOQUE A: COSTOS CON SUCURSAL ASIGNADA (COSTO INDIRECTO) =====
    def _suc_key(x: str) -> str:
        return re.sub(r"[^A-Z0-9]", "", str(x).strip().upper())

    suc_validas = [s for s in porcentajes.index.tolist()]
    map_suc = {_suc_key(s): s for s in suc_validas}

    df_original["SUCURSAL_ORIG"] = df_original["SUCURSAL"]
    df_original["SUCURSAL"] = df_original["SUCURSAL"].apply(lambda s: map_suc.get(_suc_key(s), s))

    no_general = df_original[~df_original["SUCURSAL"].isin(["GASTO GENERAL", "INTERNO", "EXTERNO"])].copy()
    no_general["TIPO COSTO"] = "COSTO INDIRECTO"
    no_general["TIPO DISTRIBUCIÓN"] = "Costo fijo en sucursal"

    directos_agr = (
        no_general
        .assign(SUCURSAL=lambda d: d["SUCURSAL"].astype(str).str.upper().str.strip())
        .groupby(["AREA/CUENTA", "SUCURSAL", "TIPO DISTRIBUCIÓN", "TIPO COSTO"], as_index=False)["CARGOS"]
        .sum()
        .rename(columns={"CARGOS": "CARGO ASIGNADO"})
    )
    directos_agr = anexar_trafico_fecha(directos_agr)

    # ===== BLOQUE B: COMUNES (GASTO GENERAL / INTERNO / EXTERNO) =====
    comunes = df_original[df_original["SUCURSAL"].isin(["GASTO GENERAL", "INTERNO", "EXTERNO"])].copy()
    if comunes.empty:
        raise ValueError("No hay filas con SUCURSAL = GASTO GENERAL / INTERNO / EXTERNO para prorratear.")

    def tipo_costo_hibrido(row) -> str:
        suc = str(row.get("SUCURSAL", "")).strip().upper()
        if suc == "INTERNO":
            return "COMUN INTERNO"
        if suc == "EXTERNO":
            return "COMUN EXTERNO"
        if suc == "GASTO GENERAL":
            c = str(row.get("CONCEPTO", "")).strip().upper()
            if c.startswith("IN"):
                return "COMUN INTERNO"
            if c.startswith("EX"):
                return "COMUN EXTERNO"
            return "COMUN INTERNO"
        return "COMUN INTERNO"

    comunes["TIPO COSTO"] = comunes.apply(tipo_costo_hibrido, axis=1)

    gg_agr = (
        comunes
        .groupby(["AREA/CUENTA", "TIPO COSTO"], as_index=False)["CARGOS"]
        .sum()
        .rename(columns={"CARGOS": "TOTAL_AREA"})
    )

    gg_agr["AREA/CUENTA"] = gg_agr["AREA/CUENTA"].astype(str).str.strip().str.upper()
    gg_agr = gg_agr.merge(catalogo, on="AREA/CUENTA", how="left")

    if gg_agr["TIPO DISTRIBUCIÓN"].isna().any():
        faltantes = gg_agr.loc[gg_agr["TIPO DISTRIBUCIÓN"].isna(), "AREA/CUENTA"].unique().tolist()
        raise ValueError(f"Faltan tipos de distribución en el catálogo para: {faltantes[:10]}")

    # ===== Vectorización del prorrateo (sin loops) =====

    # 1) porcentajes en formato largo: SUCURSAL, TIPO (columna), PCT
    pct_long = (
        porcentajes.reset_index()
        .melt(id_vars="SUCURSAL", var_name="TIPO_DIST", value_name="PCT")
    )
    pct_long["TIPO_DIST"] = pct_long["TIPO_DIST"].astype(str).str.upper().str.strip()
    pct_long = pct_long[pct_long["PCT"].fillna(0).astype(float) > 0].copy()

    # 2) gg_agr normalizado para hacer match con el tipo
    gg2 = gg_agr.copy()
    gg2["TIPO_DIST"] = gg2["TIPO DISTRIBUCIÓN"].astype(str).str.upper().str.strip()

    # 3) merge y cálculo directo
    prorr_gg = gg2.merge(pct_long, on="TIPO_DIST", how="left")
    prorr_gg = prorr_gg.dropna(subset=["PCT"]).copy()

    prorr_gg["CARGO ASIGNADO"] = (prorr_gg["TOTAL_AREA"].astype(float) * prorr_gg["PCT"].astype(float)).round(2)

    # 4) dejar columnas finales como antes
    prorr_gg = prorr_gg[[
        "AREA/CUENTA", "SUCURSAL", "TIPO DISTRIBUCIÓN", "TIPO COSTO", "CARGO ASIGNADO"
    ]]

    prorr_gg = anexar_trafico_fecha(prorr_gg)

    resultado = pd.concat([directos_agr, prorr_gg], ignore_index=True)
    return resultado


def paso5_prorrateo():
    st.header("🔄 Paso 5 - Prorrateo (con botón)")

    # Validaciones
    if not all(k in st.session_state for k in ["df_original", "porcentajes"]):
        st.warning("Faltan datos. Completa primero Paso 1 y Paso 4 (df_original + porcentajes).")
        return

    df_original = st.session_state["df_original"]
    porcentajes = st.session_state["porcentajes"]

    catalogo = fetch_catalogo()
    viajes = fetch_viajes()

    # Selector de fecha
    if not viajes.empty and "Fecha" in viajes.columns:
        fechas = pd.to_datetime(viajes["Fecha"], errors="coerce").dropna().dt.date.unique()
        fechas = sorted(list(fechas))
    else:
        fechas = []

    fecha_default = fechas[-1] if fechas else date.today()
    fecha_elegida = st.date_input("📅 Selecciona la fecha de tráfico a usar", value=fecha_default)

    # Calcular SOLO con botón
    if st.button("⚙️ Calcular prorrateo", type="primary"):
        try:
            with st.spinner("Calculando prorrateo..."):
                resultado = calcular_prorrateo_cached(df_original, porcentajes, catalogo, viajes, fecha_elegida)
            st.session_state["prorrateo_completo"] = resultado
            st.success(f"✅ Prorrateo listo: {len(resultado):,} filas.")
        except Exception as e:
            st.error(f"Error en prorrateo: {e}")
            return

    if "prorrateo_completo" not in st.session_state:
        st.info("Presiona 'Calcular prorrateo' para generar el resultado.")
        return

    resultado = st.session_state["prorrateo_completo"]
    st.subheader("📊 Resultado final")
    st.dataframe(resultado, width="stretch")

    def to_excel_bytes(df: pd.DataFrame) -> bytes:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as xw:
            df.to_excel(xw, index=False, sheet_name="Prorrateo")
        return buffer.getvalue()

    st.download_button(
        "📥 Descargar prorrateo completo",
        data=to_excel_bytes(resultado),
        file_name="prorrateo_completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def paso6_reportes():
    st.header("📅 Paso 6 - Reportes")

    if "prorrateo_completo" not in st.session_state:
        st.warning("Primero ejecuta el Paso 5 para generar el prorrateo.")
        return

    if "df_gts" not in st.session_state:
        st.warning("Primero carga el archivo GTS en el Paso 4.")
        return

    prorr = st.session_state["prorrateo_completo"].copy()
    df_gts_local = st.session_state["df_gts"].copy()

    # Normalizaciones
    prorr.columns = prorr.columns.astype(str).str.strip().str.upper()
    prorr["SUCURSAL"] = prorr["SUCURSAL"].astype(str).str.strip().str.upper()

    df_gts_local.columns = df_gts_local.columns.astype(str).str.strip().str.upper()
    df_gts_local["SUCURSAL"] = df_gts_local["SUCURSAL"].astype(str).str.strip().str.upper()

    # =========================
    # GENERALES (comunes + indirectos)
    # =========================
    st.subheader("📘 Generales (Comunes separados) e Indirectos")

    col_suc = "SUCURSAL"
    col_val = "CARGO ASIGNADO"
    col_tipo = "TIPO COSTO"

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

    st.dataframe(pivot_comunes[[col_suc, "COMUN INTERNO", "COMUN EXTERNO"]], width="stretch")

    indirectos = (
        prorr[prorr[col_tipo] == "COSTO INDIRECTO"]
        .groupby(col_suc, as_index=False)[col_val]
        .sum()
        .rename(columns={col_val: "INDIRECTO"})
    )
    st.dataframe(indirectos, width="stretch")

    final = pivot_comunes.merge(indirectos, on=col_suc, how="outer").fillna(0.0)
    final["TOTAL"] = final["COMUN INTERNO"] + final["COMUN EXTERNO"] + final["INDIRECTO"]

    st.subheader("📊 Consolidado final")
    st.dataframe(final[[col_suc, "COMUN INTERNO", "COMUN EXTERNO", "INDIRECTO", "TOTAL"]], width="stretch")

    def exportar_excel_generales():
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            pivot_comunes[[col_suc, "COMUN INTERNO", "COMUN EXTERNO"]].to_excel(writer, sheet_name="Generales", index=False)
            indirectos.to_excel(writer, sheet_name="Indirectos", index=False)
            final[[col_suc, "COMUN INTERNO", "COMUN EXTERNO", "INDIRECTO", "TOTAL"]].to_excel(writer, sheet_name="Consolidado", index=False)
        return buffer.getvalue()

    st.download_button(
        "📥 Descargar Excel (Generales/Indirectos/Consolidado)",
        data=exportar_excel_generales(),
        file_name="generales_indirectos_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    # =========================
    # COSTO POR SUCURSAL
    # =========================
    st.subheader("📌 Costo por Sucursal (Tablitas)")

    def _find_col(df, candidates):
        cols = list(df.columns)
        for c in candidates:
            if c in cols:
                return c
        for c in cols:
            for cand in candidates:
                if cand in c:
                    return c
        return None

    COL_FACT = _find_col(df_gts_local, ["FACTURACION DLLS", "FACTURACIÓN DLLS", "FACTURACION", "FACTURACIÓN"])
    COL_MC   = _find_col(df_gts_local, ["MC", "M.C.", "MARGEN", "MARGEN CONTRIBUCION", "MARGEN DE CONTRIBUCION"])

    if not COL_FACT or not COL_MC:
        st.error(f"No pude detectar columnas en GTS. Columnas encontradas: {list(df_gts_local.columns)}")
        return

    # Traer catálogo (cacheado) para mostrar método de distribución
    catalogo = fetch_catalogo()
    if catalogo.empty:
        st.error("No hay catálogo para método de distribución.")
        return

    catalogo = catalogo.rename(columns={
        "area_cuenta": "AREA/CUENTA",
        "tipo_distribucion": "TIPO DISTRIBUCIÓN"
    })
    catalogo["AREA/CUENTA"] = catalogo["AREA/CUENTA"].astype(str).str.strip().str.upper()

    def generar_tablitas_mes_sucursal(sucursal: str):
        suc = str(sucursal).strip().upper()

        row_gts = df_gts_local[df_gts_local["SUCURSAL"].eq(suc)]
        if row_gts.empty:
            return None, None, None, f"No existe la sucursal '{suc}' en el archivo GTS."

        facturacion = float(row_gts.iloc[0][COL_FACT] or 0)
        mc = float(row_gts.iloc[0][COL_MC] or 0)

        costos_directos = facturacion - mc
        utilidad = mc

        gi = prorr[(prorr["SUCURSAL"].eq(suc)) & (prorr["TIPO COSTO"].eq("COSTO INDIRECTO"))].copy()
        tabla_gi = (
            gi.groupby("AREA/CUENTA", as_index=False)["CARGO ASIGNADO"]
              .sum()
              .rename(columns={"AREA/CUENTA": "GASTOS INDIRECTOS", "CARGO ASIGNADO": "IMPORTE"})
              .sort_values("IMPORTE", ascending=False)
              .reset_index(drop=True)
        )
        tabla_gi["%"] = np.where(facturacion != 0, tabla_gi["IMPORTE"] / facturacion, 0.0)
        total_ci = float(tabla_gi["IMPORTE"].sum())
        pct_ci = (total_ci / facturacion) if facturacion != 0 else 0.0

        ge = prorr[(prorr["SUCURSAL"].eq(suc)) & (prorr["TIPO COSTO"].isin(["COMUN EXTERNO", "COMUN INTERNO"]))].copy()
        tabla_ge = (
            ge.groupby("AREA/CUENTA", as_index=False)["CARGO ASIGNADO"]
              .sum()
              .rename(columns={"AREA/CUENTA": "AREA-TIPO GASTO", "CARGO ASIGNADO": "IMPORTE"})
        )
        tabla_ge["AREA_KEY"] = tabla_ge["AREA-TIPO GASTO"].astype(str).str.strip().str.upper()
        tabla_ge = tabla_ge.merge(
            catalogo[["AREA/CUENTA", "TIPO DISTRIBUCIÓN"]].rename(columns={"AREA/CUENTA": "AREA_KEY"}),
            on="AREA_KEY",
            how="left"
        ).drop(columns=["AREA_KEY"])

        tabla_ge["%"] = np.where(facturacion != 0, tabla_ge["IMPORTE"] / facturacion, 0.0)
        tabla_ge = tabla_ge.sort_values("IMPORTE", ascending=False).reset_index(drop=True)

        total_gn = float(tabla_ge["IMPORTE"].sum())
        pct_gn = (total_gn / facturacion) if facturacion != 0 else 0.0

        pct_ut_bruta = (utilidad / facturacion) if facturacion != 0 else 0.0
        ut_per = utilidad - total_ci - total_gn
        pct_ut_per = (ut_per / facturacion) if facturacion != 0 else 0.0

        tabla_top = pd.DataFrame([{
            "Sucursal": suc,
            "Facturación": facturacion,
            "Costos Directos": costos_directos,
            "Utilidad": utilidad,
            "% Ut Bruta": pct_ut_bruta,
            "Costos Indirectos": total_ci,
            "% CI": pct_ci,
            "Gastos Generales": total_gn,
            "% GN": pct_gn,
            "UT/PER": ut_per,
            "%UT/PER": pct_ut_per
        }])

        return tabla_top, tabla_gi, tabla_ge, None

    sucursales_disponibles = sorted(df_gts_local["SUCURSAL"].dropna().unique().tolist())
    sucursal_sel = st.selectbox("Selecciona sucursal", sucursales_disponibles)

    tabla_top, tabla_gi, tabla_ge, err = generar_tablitas_mes_sucursal(sucursal_sel)
    if err:
        st.error(err)
        return

    st.subheader("🧾 Resumen superior")
    st.dataframe(tabla_top, width="stretch")

    st.subheader("📌 GASTOS INDIRECTOS (Costo indirecto)")
    st.dataframe(tabla_gi, width="stretch")

    st.subheader("📌 AREA-TIPO GASTO (Común interno + Común externo)")
    st.dataframe(tabla_ge, width="stretch")

    def exportar_costo_sucursal_excel(tabla_top, tabla_gi, tabla_ge, nombre_sucursal):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            sheet = str(nombre_sucursal)[:31]
            tabla_top.to_excel(writer, sheet_name=sheet, index=False, startrow=0, startcol=0)
            tabla_gi.to_excel(writer, sheet_name=sheet, index=False, startrow=4, startcol=0)
            tabla_ge.to_excel(writer, sheet_name=sheet, index=False, startrow=4, startcol=6)
        return buffer.getvalue()

    st.download_button(
        "📥 Descargar Excel de esta sucursal",
        data=exportar_costo_sucursal_excel(tabla_top, tabla_gi, tabla_ge, sucursal_sel),
        file_name=f"costo_por_sucursal_{sucursal_sel}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.title("💰 Prorrateador")

paso = st.sidebar.radio(
    "Pasos",
    [
        "1) PASO 1 (Excel)",
        "2) Tráfico",
        "3) Catálogo",
        "4) GTS",
        "5) Prorrateo",
        "6) Reportes",
    ]
)

if paso.startswith("1)"):
    paso1_upload()
elif paso.startswith("2)"):
    paso2_trafico()
elif paso.startswith("3)"):
    paso3_catalogo()
elif paso.startswith("4)"):
    paso4_gts()
elif paso.startswith("5)"):
    paso5_prorrateo()
else:
    paso6_reportes()

