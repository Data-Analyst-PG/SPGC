import re
import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Confronta Liquidaciones vs Contabilidad", layout="wide")

# -----------------------------
# Helpers
# -----------------------------
def norm_text(x: object) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip().upper()
    s = re.sub(r"\s+", " ", s)
    return s

def norm_amount(x: object, ndigits: int = 2) -> float:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return float("nan")
        return round(float(x), ndigits)
    except Exception:
        return float("nan")

def build_seq(df: pd.DataFrame, key_cols: list[str], seq_col: str = "_seq") -> pd.DataFrame:
    # consecutivo por repetición dentro de la llave para empatar duplicados 1:1
    df = df.copy()
    df[seq_col] = df.groupby(key_cols).cumcount() + 1
    return df

def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, d in sheets.items():
            d.to_excel(writer, sheet_name=name[:31], index=False)
    bio.seek(0)
    return bio.getvalue()

# -----------------------------
# UI
# -----------------------------
st.title("Confronta STAR vs SAC: Liquidaciones vs Contabilidad")

with st.sidebar:
    st.header("Entradas")
    liq_file = st.file_uploader("Excel Liquidaciones (STAR)", type=["xlsx"])
    cont_file = st.file_uploader("Excel Contabilidad", type=["xlsx"])

    st.divider()
    st.header("Catálogo de operadores")

    catalogo_file = st.file_uploader(
        "Catálogo operadores / owners",
        type=["xlsx"],
        help="Archivo con columnas NOMBRE, Usuario STAR, Usuario SAC y Tipo"
    )

    solo_owner = st.checkbox(
        "Mostrar solo registros de Owners",
        value=True
    )

    st.divider()
    st.header("Reglas de comparación")

    ndigits = st.number_input(
        "Redondeo de importe (decimales)",
        min_value=0, max_value=4, value=2, step=1
    )
    st.divider()
    st.header("Filtros por tipo")

    liq_tipo = st.selectbox(
        "Liquidaciones: Tipo_Concepto a considerar",
        options=["E", "I"],
        index=0,  # por default E
        help="Solo se compararán filas de Liquidaciones con este Tipo_Concepto"
    )

    cont_tipo = st.selectbox(
        "Contabilidad: TipoMovimiento a considerar",
        options=["H", "D"],
        index=0,  # por default H
        help="Solo se compararán filas de Contabilidad con este TipoMovimiento"
    )

    st.caption("Sugerencias: si no encuentra exacto, busca por PR+Unidad+TipoPago+Importe (ignorando Viaje).")
    enable_suggestions = st.checkbox("Generar sugerencias para no-matcheados", value=True)
    suggestions_limit = st.number_input("Máx. sugerencias por renglón", 1, 10, 3)

# -----------------------------
# Run
# -----------------------------
if not liq_file or not cont_file:
    st.info("Carga ambos archivos para iniciar.")
    st.stop()

# Column mapping (según tus archivos)
liq_usecols = ["Liquidacion", "Numero_Viaje", "TipoPago", "Monto", "Unidad", "Operador", "Tipo_Concepto"]
cont_usecols = ["Factura", "Referencia", "TipoPago", "Importe", "Unidad", "NombreCuentaContable", "TipoMovimiento"]

try:
    liq = pd.read_excel(liq_file, sheet_name="LiquidacionesSET_PLUS_datos", usecols=liq_usecols)
    cont = pd.read_excel(cont_file, sheet_name="ContabilidadSET_PLUS_datos", usecols=cont_usecols)
except Exception as e:
    st.error(f"No pude leer los excels. Error: {e}")
    st.stop()

# -----------------------------
# Catálogo de operadores
# -----------------------------
catalogo = None
star_to_nombre = {}
sac_to_nombre = {}

if catalogo_file is not None:
    catalogo = pd.read_excel(catalogo_file)

    # Normaliza nombres de columnas del catálogo
    catalogo.columns = (
        catalogo.columns.astype(str)
        .str.strip()
        .str.upper()
    )

    # Renombra a nombres estándar usando los nombres reales de tu catálogo
    catalogo = catalogo.rename(columns={
        "NOMBRE": "NOMBRE",
        "USUARIO STAR (SUGERIDO)": "USUARIO_STAR",
        "USUARIO SAC (SUGERIDO)": "USUARIO_SAC",
        "TIPO": "TIPO",
    })

    # Validación
    expected_cols = ["NOMBRE", "USUARIO_STAR", "USUARIO_SAC", "TIPO"]
    faltantes = [c for c in expected_cols if c not in catalogo.columns]
    if faltantes:
        st.error(f"Al catálogo le faltan columnas: {faltantes}")
        st.write("Columnas encontradas en catálogo:", catalogo.columns.tolist())
        st.stop()

    # Normaliza contenido
    for col in ["NOMBRE", "USUARIO_STAR", "USUARIO_SAC", "TIPO"]:
        catalogo[col] = catalogo[col].apply(norm_text)

    if solo_owner:
        catalogo = catalogo[catalogo["TIPO"] == "OWNER"].copy()

    star_to_nombre = dict(
        catalogo.loc[catalogo["USUARIO_STAR"] != "", ["USUARIO_STAR", "NOMBRE"]].values
    )

    sac_to_nombre = dict(
        catalogo.loc[catalogo["USUARIO_SAC"] != "", ["USUARIO_SAC", "NOMBRE"]].values
    )
        
# Normalize
liq = liq.rename(columns={
    "Liquidacion": "PR",
    "Numero_Viaje": "VIAJE",
    "TipoPago": "TIPO_PAGO",
    "Monto": "IMPORTE",
    "Unidad": "UNIDAD",
    "Operador": "OWNER_LIQ",
    "Tipo_Concepto": "TIPO_CONCEPTO",
})
cont = cont.rename(columns={
    "Factura": "PR",
    "Referencia": "VIAJE",
    "TipoPago": "TIPO_PAGO",
    "Importe": "IMPORTE",
    "Unidad": "UNIDAD",
    "NombreCuentaContable": "OWNER_CONT",
    "TipoMovimiento": "TIPO_MOV",
})

for c in ["PR", "VIAJE", "TIPO_PAGO", "UNIDAD", "OWNER_LIQ", "TIPO_CONCEPTO"]:
    if c in liq.columns:
        liq[c] = liq[c].apply(norm_text)

for c in ["PR", "VIAJE", "TIPO_PAGO", "UNIDAD", "OWNER_CONT", "TIPO_MOV"]:
    if c in cont.columns:
        cont[c] = cont[c].apply(norm_text)

# Owner estándar según catálogo
# Owner estándar según catálogo
if catalogo_file is not None:
    liq["OWNER_STD_LIQ"] = liq["OWNER_LIQ"].map(star_to_nombre).fillna("")
    cont["OWNER_STD_CONT"] = cont["OWNER_CONT"].map(sac_to_nombre).fillna("")

    st.sidebar.caption(f"Filas catálogo después de filtro OWNER: {len(catalogo)}")
    st.sidebar.caption(f"Matches STAR en catálogo: {(liq['OWNER_STD_LIQ'] != '').sum()}")
    st.sidebar.caption(f"Matches SAC en catálogo: {(cont['OWNER_STD_CONT'] != '').sum()}")
else:
    liq["OWNER_STD_LIQ"] = ""
    cont["OWNER_STD_CONT"] = ""
    
liq["IMPORTE"] = liq["IMPORTE"].apply(lambda x: norm_amount(x, ndigits))
cont["IMPORTE"] = cont["IMPORTE"].apply(lambda x: norm_amount(x, ndigits))

# Regla de negocio (editable desde sidebar)
liq_f = liq[liq["TIPO_CONCEPTO"] == liq_tipo].copy()
cont_f = cont[cont["TIPO_MOV"] == cont_tipo].copy()

# Filtrar solo registros que estén en el catálogo (owners)
if solo_owner and catalogo_file is not None:
    liq_f = liq_f[liq_f["OWNER_STD_LIQ"] != ""].copy()
    cont_f = cont_f[cont_f["OWNER_STD_CONT"] != ""].copy()
    
st.subheader("Resumen de carga")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Liquidaciones (original)", len(liq))
c2.metric("Liquidaciones (filtrado)", len(liq_f))
c3.metric("Contabilidad (original)", len(cont))
c4.metric("Contabilidad (filtrado)", len(cont_f))

# Matching key (SIN owner) + consecutivo por duplicado
key_cols = ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE"]

liq_k = build_seq(liq_f.copy(), key_cols, "_seq")
cont_k = build_seq(cont_f.copy(), key_cols, "_seq")

# Construcción de llave estable
liq_k["_KEY"] = (
    liq_k["PR"].fillna("").astype(str) + "||" +
    liq_k["VIAJE"].fillna("").astype(str) + "||" +
    liq_k["UNIDAD"].fillna("").astype(str) + "||" +
    liq_k["TIPO_PAGO"].fillna("").astype(str) + "||" +
    liq_k["IMPORTE"].fillna("").astype(str)
)

cont_k["_KEY"] = (
    cont_k["PR"].fillna("").astype(str) + "||" +
    cont_k["VIAJE"].fillna("").astype(str) + "||" +
    cont_k["UNIDAD"].fillna("").astype(str) + "||" +
    cont_k["TIPO_PAGO"].fillna("").astype(str) + "||" +
    cont_k["IMPORTE"].fillna("").astype(str)
)

liq_k["_KEYSEQ"] = liq_k["_KEY"] + "||SEQ=" + liq_k["_seq"].astype(str)
cont_k["_KEYSEQ"] = cont_k["_KEY"] + "||SEQ=" + cont_k["_seq"].astype(str)

# Exact merge by KEY+SEQ
m = liq_k.merge(
    cont_k,
    how="outer",
    on="_KEYSEQ",
    suffixes=("_LIQ", "_CONT"),
    indicator=True
)

matched = m[m["_merge"] == "both"].copy()
only_liq = m[m["_merge"] == "left_only"].copy()
only_cont = m[m["_merge"] == "right_only"].copy()

matched["OWNER_MATCH"] = matched["OWNER_LIQ"].fillna("") == matched["OWNER_CONT"].fillna("")
matched["DIF_OWNER"] = ~matched["OWNER_MATCH"]

matches_ok = matched[matched["DIF_OWNER"] == False].copy()
matches_owner_diff = matched[matched["DIF_OWNER"] == True].copy()

def pick_cols(df: pd.DataFrame, side: str) -> list[str]:
    cols = []
    for c in ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE"]:
        cand = f"{c}_{side}"
        if cand in df.columns:
            cols.append(cand)
    owner = f"OWNER_{side}"
    if owner in df.columns:
        cols.append(owner)
    return cols

ok_view = matches_ok[pick_cols(matches_ok, "LIQ") + pick_cols(matches_ok, "CONT")].copy()
diff_view = matches_owner_diff[pick_cols(matches_owner_diff, "LIQ") + pick_cols(matches_owner_diff, "CONT")].copy()
liq_missing_view = only_liq[pick_cols(only_liq, "LIQ")].copy()
cont_missing_view = only_cont[pick_cols(only_cont, "CONT")].copy()

# ----------------------------------------
# Duplicados detectados
# ----------------------------------------
st.divider()
st.subheader("🔁 Registros duplicados detectados")

dup_key_cols = ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE"]

# Filas duplicadas (detalle)
dup_liq = liq_f[liq_f.duplicated(subset=dup_key_cols, keep=False)].copy()
dup_cont = cont_f[cont_f.duplicated(subset=dup_key_cols, keep=False)].copy()

# Resumen por llave
dup_liq_resumen = (
    liq_f.groupby(dup_key_cols, dropna=False)
    .size()
    .reset_index(name="REPETICIONES")
)
dup_liq_resumen = dup_liq_resumen[dup_liq_resumen["REPETICIONES"] > 1].copy()

dup_cont_resumen = (
    cont_f.groupby(dup_key_cols, dropna=False)
    .size()
    .reset_index(name="REPETICIONES")
)
dup_cont_resumen = dup_cont_resumen[dup_cont_resumen["REPETICIONES"] > 1].copy()

c1, c2 = st.columns(2)
c1.metric("Duplicados en Liquidaciones", len(dup_liq))
c2.metric("Duplicados en Contabilidad", len(dup_cont))

tabs_dup = st.tabs([
    f"Detalle duplicados Liquidaciones ({len(dup_liq)})",
    f"Resumen duplicados Liquidaciones ({len(dup_liq_resumen)})",
    f"Detalle duplicados Contabilidad ({len(dup_cont)})",
    f"Resumen duplicados Contabilidad ({len(dup_cont_resumen)})",
])

with tabs_dup[0]:
    if dup_liq.empty:
        st.success("No hay filas duplicadas en Liquidaciones.")
    else:
        cols_liq_dup = [c for c in ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_LIQ"] if c in dup_liq.columns]
        st.dataframe(
            dup_liq[cols_liq_dup].sort_values(dup_key_cols),
            use_container_width=True,
            height=420
        )

with tabs_dup[1]:
    if dup_liq_resumen.empty:
        st.success("No hay combinaciones duplicadas en Liquidaciones.")
    else:
        st.dataframe(
            dup_liq_resumen.sort_values(["REPETICIONES"] + dup_key_cols, ascending=[False, True, True, True, True, True]),
            use_container_width=True,
            height=420
        )

with tabs_dup[2]:
    if dup_cont.empty:
        st.success("No hay filas duplicadas en Contabilidad.")
    else:
        cols_cont_dup = [c for c in ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_CONT"] if c in dup_cont.columns]
        st.dataframe(
            dup_cont[cols_cont_dup].sort_values(dup_key_cols),
            use_container_width=True,
            height=420
        )

with tabs_dup[3]:
    if dup_cont_resumen.empty:
        st.success("No hay combinaciones duplicadas en Contabilidad.")
    else:
        st.dataframe(
            dup_cont_resumen.sort_values(["REPETICIONES"] + dup_key_cols, ascending=[False, True, True, True, True, True]),
            use_container_width=True,
            height=420
        )
        
# ----------------------------------------
# Auditoría de exclusiones (desplegable)
# ----------------------------------------

# 1) Excluidos por tipo
liq_excl_tipo = liq[liq["TIPO_CONCEPTO"] != liq_tipo].copy()
cont_excl_tipo = cont[cont["TIPO_MOV"] != cont_tipo].copy()

# 2) Excluidos por catálogo owner (solo sobre los que sí pasaron el filtro de tipo)
if solo_owner and catalogo_file is not None:
    liq_excl_owner = liq[
        (liq["TIPO_CONCEPTO"] == liq_tipo) &
        (liq["OWNER_STD_LIQ"] == "")
    ].copy()

    cont_excl_owner = cont[
        (cont["TIPO_MOV"] == cont_tipo) &
        (cont["OWNER_STD_CONT"] == "")
    ].copy()
else:
    liq_excl_owner = pd.DataFrame(columns=liq.columns)
    cont_excl_owner = pd.DataFrame(columns=cont.columns)

# 3) Excluidos totales = todo lo que NO quedó en el dataframe final filtrado
liq_excl_total = liq.loc[~liq.index.isin(liq_f.index)].copy()
cont_excl_total = cont.loc[~cont.index.isin(cont_f.index)].copy()

# Orden opcional para que sea más fácil revisar
for df in (
    liq_excl_tipo, liq_excl_owner, liq_excl_total,
    cont_excl_tipo, cont_excl_owner, cont_excl_total
):
    if "PR" in df.columns and "VIAJE" in df.columns:
        df.sort_values(by=["PR", "VIAJE", "UNIDAD", "TIPO_PAGO"], inplace=True, kind="stable")

with st.expander("🔎 Ver filtros aplicados (criterios)", expanded=False):
    st.markdown("### Criterios activos")
    st.write(f"- Liquidaciones: **TIPO_CONCEPTO = '{liq_tipo}'**")
    st.write(f"- Contabilidad: **TIPO_MOV = '{cont_tipo}'**")

    if solo_owner and catalogo_file is not None:
        st.write("- Catálogo: **solo registros con owner encontrado en catálogo**")
    else:
        st.write("- Catálogo: **sin filtro por owner**")

    st.divider()
    st.markdown("### Resumen de exclusiones")

    c1, c2, c3 = st.columns(3)
    c1.metric("Liquidaciones excluidas por tipo", len(liq_excl_tipo))
    c2.metric("Liquidaciones excluidas por owner", len(liq_excl_owner))
    c3.metric("Liquidaciones excluidas total", len(liq_excl_total))

    c4, c5, c6 = st.columns(3)
    c4.metric("Contabilidad excluida por tipo", len(cont_excl_tipo))
    c5.metric("Contabilidad excluida por owner", len(cont_excl_owner))
    c6.metric("Contabilidad excluida total", len(cont_excl_total))

    st.divider()
    st.markdown("### Filas excluidas (detalle)")

    colA, colB, colC = st.columns([1, 1, 2])
    with colA:
        max_rows = st.number_input(
            "Mostrar máximo de filas por tabla",
            min_value=100, max_value=200000, value=2000, step=100
        )
    with colB:
        show_all = st.checkbox("Mostrar TODO (puede ser pesado)", value=False)
    with colC:
        q = st.text_input("Buscar (PR / Viaje / Unidad / TipoPago / Owner)", value="").strip()

    def filter_df(df: pd.DataFrame, query: str) -> pd.DataFrame:
        if df.empty:
            return df
        if not query:
            return df
        cols = [c for c in ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_LIQ", "OWNER_CONT"] if c in df.columns]
        if not cols:
            return df
        mask = df[cols].astype(str).apply(lambda s: s.str.contains(query, case=False, na=False)).any(axis=1)
        return df.loc[mask].copy()

    tabs = st.tabs([
        f"Liq excluidas por tipo ({len(liq_excl_tipo)})",
        f"Liq excluidas por owner ({len(liq_excl_owner)})",
        f"Liq excluidas total ({len(liq_excl_total)})",
        f"Cont excluidas por tipo ({len(cont_excl_tipo)})",
        f"Cont excluidas por owner ({len(cont_excl_owner)})",
        f"Cont excluidas total ({len(cont_excl_total)})",
    ])

    tablas = [
        (tabs[0], liq_excl_tipo, ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_LIQ"]),
        (tabs[1], liq_excl_owner, ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_LIQ"]),
        (tabs[2], liq_excl_total, ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_LIQ"]),
        (tabs[3], cont_excl_tipo, ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_CONT"]),
        (tabs[4], cont_excl_owner, ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_CONT"]),
        (tabs[5], cont_excl_total, ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_CONT"]),
    ]

    for tab, df_src, cols_show in tablas:
        with tab:
            df_view = filter_df(df_src, q)
            if not show_all:
                df_view = df_view.head(int(max_rows))

            if df_src.empty:
                st.success("Sin filas en esta categoría.")
            else:
                cols_ok = [c for c in cols_show if c in df_view.columns]
                st.dataframe(df_view[cols_ok], use_container_width=True, height=420)
    st.info(
        "Tip: si necesitas evidencia completa, también puedes incluir estas filas en el Excel descargable "
        "(pestañas Filtrados_Liquidaciones y Filtrados_Contabilidad)."
    )
    
st.divider()
st.subheader("Resultados")

t1, t2, t3, t4 = st.tabs([
    f"✅ Matcheados OK ({len(ok_view)})",
    f"⚠️ Matcheados con discrepancia ({len(diff_view)})",
    f"❌ Falta en Contabilidad ({len(liq_missing_view)})",
    f"❌ Falta en Liquidaciones ({len(cont_missing_view)})",
])

with t1:
    st.dataframe(ok_view, use_container_width=True, height=420)

with t2:
    st.dataframe(diff_view, use_container_width=True, height=420)

with t3:
    st.dataframe(liq_missing_view, use_container_width=True, height=420)

with t4:
    st.dataframe(cont_missing_view, use_container_width=True, height=420)

# 2) Suggestions for unmatched (optional)
suggestions_df = pd.DataFrame()
if enable_suggestions and (len(liq_missing_view) > 0 or len(cont_missing_view) > 0):
    st.divider()
    st.subheader("Sugerencias (match relajado ignorando VIAJE)")

    # relaxed key
    relaxed_cols = ["PR", "UNIDAD", "TIPO_PAGO", "IMPORTE"]

    liq_u = liq_k[liq_k["_KEYSEQ"].isin(only_liq["_KEYSEQ"])].copy()
    cont_u = cont_k[cont_k["_KEYSEQ"].isin(only_cont["_KEYSEQ"])].copy()

    liq_u["_REL"] = (
        liq_u["PR"].fillna("").astype(str) + "||" +
        liq_u["UNIDAD"].fillna("").astype(str) + "||" +
        liq_u["TIPO_PAGO"].fillna("").astype(str) + "||" +
        liq_u["IMPORTE"].fillna("").astype(str)
    )

    cont_u["_REL"] = (
        cont_u["PR"].fillna("").astype(str) + "||" +
        cont_u["UNIDAD"].fillna("").astype(str) + "||" +
        cont_u["TIPO_PAGO"].fillna("").astype(str) + "||" +
        cont_u["IMPORTE"].fillna("").astype(str)
    )

    # Build candidates: join by relaxed key
    cand = liq_u.merge(cont_u, how="inner", on="_REL", suffixes=("_LIQ", "_CONT"))
    if len(cand) > 0:
        # rank: same viaje exact gets priority even though relaxed key ignores it
        cand["SAME_VIAJE"] = cand["VIAJE_LIQ"].fillna("") == cand["VIAJE_CONT"].fillna("")
        cand["OWNER_MATCH"] = cand["OWNER_LIQ"].fillna("") == cand["OWNER_CONT"].fillna("")
        cand = cand.sort_values(by=["SAME_VIAJE", "OWNER_MATCH"], ascending=[False, False])

        # cap suggestions per each LIQ row
        cand["_rank"] = cand.groupby("_KEYSEQ_LIQ").cumcount() + 1
        cand = cand[cand["_rank"] <= suggestions_limit].copy()

        suggestions_df = cand[[
            "PR_LIQ","VIAJE_LIQ","UNIDAD_LIQ","TIPO_PAGO_LIQ","IMPORTE_LIQ","OWNER_LIQ",
            "PR_CONT","VIAJE_CONT","UNIDAD_CONT","TIPO_PAGO_CONT","IMPORTE_CONT","OWNER_CONT",
            "SAME_VIAJE","OWNER_MATCH","_rank"
        ]].copy()

        st.dataframe(suggestions_df, use_container_width=True, height=420)
    else:
        st.info("No encontré candidatos bajo el criterio relajado (PR+Unidad+TipoPago+Importe).")

# Export
st.divider()
sheets = {
    "Matched_OK": ok_view,
    "Matched_Owner_Diff": diff_view,
    "Missing_in_Contabilidad": liq_missing_view,
    "Missing_in_Liquidaciones": cont_missing_view,
    "Filtrados_Liq_Tipo": liq_excl_tipo,
    "Filtrados_Liq_Owner": liq_excl_owner,
    "Filtrados_Liq_Total": liq_excl_total,
    "Filtrados_Cont_Tipo": cont_excl_tipo,
    "Filtrados_Cont_Owner": cont_excl_owner,
    "Filtrados_Cont_Total": cont_excl_total,
    "Criterios_Filtro": pd.DataFrame({
        "criterio": [
            f"Liquidaciones: TIPO_CONCEPTO = '{liq_tipo}'",
            f"Contabilidad: TIPO_MOV = '{cont_tipo}'",
            f"Filtro owner activo: {solo_owner}",
        ]
    }),
}
if enable_suggestions:
    sheets["Suggestions_Relaxed"] = suggestions_df

xlsx_bytes = to_excel_bytes(sheets)
st.download_button(
    "⬇️ Descargar reporte de confronta (Excel)",
    data=xlsx_bytes,
    file_name="reporte_confronta_star.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

