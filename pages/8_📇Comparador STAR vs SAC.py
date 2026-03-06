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

def apply_exclusions(df: pd.DataFrame, col: str, patterns: list[str]) -> pd.DataFrame:
    if not patterns:
        return df
    rx = re.compile("|".join([f"({p})" for p in patterns]), re.IGNORECASE)
    mask = ~df[col].fillna("").astype(str).apply(lambda s: bool(rx.search(s)))
    return df.loc[mask].copy()

# -----------------------------
# UI
# -----------------------------
st.title("Confronta STAR vs SAC: Liquidaciones vs Contabilidad")

with st.sidebar:
    st.header("Entradas")
    liq_file = st.file_uploader("Excel Liquidaciones (STAR)", type=["xlsx"])
    cont_file = st.file_uploader("Excel Contabilidad", type=["xlsx"])

    st.divider()
    st.header("Reglas de comparación")

    ndigits = st.number_input("Redondeo de importe (decimales)", min_value=0, max_value=4, value=2, step=1)

    st.caption("Exclusiones (regex). Ejemplos típicos:")
    default_exclusions = [
        r"CARGOS ADICIONALES",
        r"CT\.\s*PAGO\s*X\s*MILLAS",
    ]
    exclusions = st.text_area(
        "Patrones a IGNORAR (uno por línea)",
        value="\n".join(default_exclusions),
        height=120
    )
    exclusion_patterns = [p.strip() for p in exclusions.splitlines() if p.strip()]

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
liq_usecols = ["Liquidacion", "Numero_Viaje", "TipoPago", "Monto", "Unidad", "Operador"]
cont_usecols = ["Factura", "Referencia", "TipoPago", "Importe", "Unidad", "NombreCuentaContable"]

try:
    liq = pd.read_excel(liq_file, sheet_name=0, usecols=liq_usecols)
    cont = pd.read_excel(cont_file, sheet_name=0, usecols=cont_usecols)
except Exception as e:
    st.error(f"No pude leer los excels. Error: {e}")
    st.stop()

# Normalize
liq = liq.rename(columns={
    "Liquidacion": "PR",
    "Numero_Viaje": "VIAJE",
    "TipoPago": "TIPO_PAGO",
    "Monto": "IMPORTE",
    "Unidad": "UNIDAD",
    "Operador": "OWNER_LIQ",
})
cont = cont.rename(columns={
    "Factura": "PR",
    "Referencia": "VIAJE",
    "TipoPago": "TIPO_PAGO",
    "Importe": "IMPORTE",
    "Unidad": "UNIDAD",
    "NombreCuentaContable": "OWNER_CONT",
})

for c in ["PR", "VIAJE", "TIPO_PAGO", "UNIDAD", "OWNER_LIQ"]:
    if c in liq.columns:
        liq[c] = liq[c].apply(norm_text)

for c in ["PR", "VIAJE", "TIPO_PAGO", "UNIDAD", "OWNER_CONT"]:
    if c in cont.columns:
        cont[c] = cont[c].apply(norm_text)

liq["IMPORTE"] = liq["IMPORTE"].apply(lambda x: norm_amount(x, ndigits))
cont["IMPORTE"] = cont["IMPORTE"].apply(lambda x: norm_amount(x, ndigits))

# Apply exclusions by TIPO_PAGO (esto resuelve tu caso "cargos adicionales" y "ct pago x millas" si así lo defines)
liq_f = apply_exclusions(liq, "TIPO_PAGO", exclusion_patterns)
cont_f = apply_exclusions(cont, "TIPO_PAGO", exclusion_patterns)

st.subheader("Resumen de carga")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Liquidaciones (original)", len(liq))
c2.metric("Liquidaciones (filtrado)", len(liq_f))
c3.metric("Contabilidad (original)", len(cont))
c4.metric("Contabilidad (filtrado)", len(cont_f))

# Matching key (SIN owner) + consecutivo por duplicado
key_cols = ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE"]

liq_k = build_seq(liq_f, key_cols, "_seq")
cont_k = build_seq(cont_f, key_cols, "_seq")

liq_k["_KEY"] = liq_k[key_cols].astype(str).agg("||".join, axis=1)
cont_k["_KEY"] = cont_k[key_cols].astype(str).agg("||".join, axis=1)

liq_k["_KEYSEQ"] = liq_k["_KEY"] + f"||SEQ=" + liq_k["_seq"].astype(str)
cont_k["_KEYSEQ"] = cont_k["_KEY"] + f"||SEQ=" + cont_k["_seq"].astype(str)

# 1) Exact merge by KEY+SEQ
m = liq_k.merge(
    cont_k,
    how="outer",
    on="_KEYSEQ",
    suffixes=("_LIQ", "_CONT"),
    indicator=True
)

# Split outputs
matched = m[m["_merge"] == "both"].copy()
only_liq = m[m["_merge"] == "left_only"].copy()
only_cont = m[m["_merge"] == "right_only"].copy()

# Owner discrepancy flag on matched
matched["OWNER_MATCH"] = matched["OWNER_LIQ"].fillna("") == matched["OWNER_CONT"].fillna("")
matched["DIF_OWNER"] = ~matched["OWNER_MATCH"]

matches_ok = matched[matched["DIF_OWNER"] == False].copy()
matches_owner_diff = matched[matched["DIF_OWNER"] == True].copy()

# Make them pretty (select columns)
def pick_cols(df: pd.DataFrame, side: str) -> list[str]:
    # side: "LIQ" or "CONT"
    cols = []
    for c in ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE"]:
        cand = f"{c}_{side}"
        if cand in df.columns:
            cols.append(cand)
        elif c in df.columns:
            cols.append(c)
    owner = f"OWNER_{side}"
    if owner in df.columns:
        cols.append(owner)
    return cols

ok_view = matches_ok[pick_cols(matches_ok, "LIQ") + pick_cols(matches_ok, "CONT")].copy()
diff_view = matches_owner_diff[pick_cols(matches_owner_diff, "LIQ") + pick_cols(matches_owner_diff, "CONT")].copy()

liq_missing_view = only_liq[pick_cols(only_liq, "LIQ")].copy()
cont_missing_view = only_cont[pick_cols(only_cont, "CONT")].copy()

# ----------------------------------------
# Auditoría de exclusiones (desplegable)
# ----------------------------------------
liq_excl = liq.loc[~liq.index.isin(liq_f.index)].copy()
cont_excl = cont.loc[~cont.index.isin(cont_f.index)].copy()

# Orden opcional para que sea más fácil revisar
for df in (liq_excl, cont_excl):
    if "PR" in df.columns and "VIAJE" in df.columns:
        df.sort_values(by=["PR", "VIAJE", "UNIDAD", "TIPO_PAGO"], inplace=True, kind="stable")

with st.expander("🔎 Ver filtros aplicados (exclusiones)", expanded=False):
    st.markdown("### Patrones activos (regex)")
    if exclusion_patterns:
        st.code("\n".join(exclusion_patterns))
    else:
        st.info("No hay exclusiones activas.")

    c1, c2 = st.columns(2)
    c1.metric("Filtrados en Liquidaciones", len(liq_excl))
    c2.metric("Filtrados en Contabilidad", len(cont_excl))

    st.divider()
    st.markdown("### Filas excluidas (detalle)")

    # Controles para no saturar la pantalla
    colA, colB, colC = st.columns([1, 1, 2])
    with colA:
        max_rows = st.number_input("Mostrar máximo de filas por tabla", min_value=100, max_value=200000, value=2000, step=100)
    with colB:
        show_all = st.checkbox("Mostrar TODO (puede ser pesado)", value=False)
    with colC:
        q = st.text_input("Buscar (PR / Viaje / Unidad / TipoPago / Owner)", value="").strip()

    def filter_df(df: pd.DataFrame, query: str) -> pd.DataFrame:
        if not query:
            return df
        # busca en columnas relevantes si existen
        cols = [c for c in ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_LIQ", "OWNER_CONT"] if c in df.columns]
        if not cols:
            return df
        mask = df[cols].astype(str).apply(lambda s: s.str.contains(query, case=False, na=False)).any(axis=1)
        return df.loc[mask].copy()

    liq_excl_view = filter_df(liq_excl, q)
    cont_excl_view = filter_df(cont_excl, q)

    if not show_all:
        liq_excl_view = liq_excl_view.head(int(max_rows))
        cont_excl_view = cont_excl_view.head(int(max_rows))

    t1, t2 = st.tabs([
        f"Liquidaciones excluidas ({len(liq_excl)} total)",
        f"Contabilidad excluida ({len(cont_excl)} total)"
    ])

    with t1:
        st.caption("Estas filas se eliminaron ANTES de comparar.")
        if len(liq_excl) == 0:
            st.success("No se excluyó ninguna fila en Liquidaciones.")
        else:
            # Solo columnas clave para auditoría (más legible)
            cols_liq = [c for c in ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_LIQ"] if c in liq_excl_view.columns]
            st.dataframe(liq_excl_view[cols_liq], use_container_width=True, height=420)

    with t2:
        st.caption("Estas filas se eliminaron ANTES de comparar.")
        if len(cont_excl) == 0:
            st.success("No se excluyó ninguna fila en Contabilidad.")
        else:
            cols_cont = [c for c in ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE", "OWNER_CONT"] if c in cont_excl_view.columns]
            st.dataframe(cont_excl_view[cols_cont], use_container_width=True, height=420)

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

    liq_u["_REL"] = liq_u[relaxed_cols].astype(str).agg("||".join, axis=1)
    cont_u["_REL"] = cont_u[relaxed_cols].astype(str).agg("||".join, axis=1)

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
    "Filtrados_Liquidaciones": liq_excl,
    "Filtrados_Contabilidad": cont_excl,
    "Patrones_Exclusion": pd.DataFrame({"regex": exclusion_patterns}),
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
