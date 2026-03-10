import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Comparador STAR vs SAC v2", layout="wide")

# ============================================================
# Helpers
# ============================================================

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
    out = df.copy()
    out[seq_col] = out.groupby(key_cols, dropna=False).cumcount() + 1
    return out


def ensure_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if not out.columns.duplicated().any():
        return out
    seen = {}
    new_cols = []
    for c in out.columns:
        if c not in seen:
            seen[c] = 0
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
    out.columns = new_cols
    return out


def prepare_df_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    out = ensure_unique_columns(df)
    for col in out.columns:
        try:
            dtype = str(out[col].dtype)
            if dtype == "category":
                out[col] = out[col].astype("string").fillna("")
            elif "string" in dtype:
                out[col] = out[col].fillna("")
        except Exception:
            pass
    return out


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(
        bio,
        engine="xlsxwriter",
        engine_kwargs={"options": {"constant_memory": True}},
    ) as writer:
        for name, df in sheets.items():
            prepare_df_for_excel(df).to_excel(writer, sheet_name=name[:31], index=False)
    bio.seek(0)
    return bio.getvalue()


def show_df(df: pd.DataFrame, height: int = 420, max_rows: int = 2000):
    if df.empty:
        st.dataframe(df, use_container_width=True, height=height)
        return
    if len(df) > max_rows:
        st.caption(f"Mostrando {max_rows:,} de {len(df):,} filas. El archivo descargable incluye todo.")
        st.dataframe(df.head(max_rows), use_container_width=True, height=height)
        return
    st.dataframe(df, use_container_width=True, height=height)


def read_table(file_obj, preferred_sheet: str | None = None, usecols: list[str] | None = None) -> pd.DataFrame:
    suffix = Path(file_obj.name).suffix.lower()
    raw = file_obj.getvalue()
    if suffix == ".csv":
        return pd.read_csv(BytesIO(raw), usecols=usecols, low_memory=False)
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        if preferred_sheet:
            try:
                return pd.read_excel(BytesIO(raw), sheet_name=preferred_sheet, usecols=usecols)
            except Exception:
                return pd.read_excel(BytesIO(raw), usecols=usecols)
        return pd.read_excel(BytesIO(raw), usecols=usecols)
    raise ValueError(f"Formato no soportado: {suffix}")


def standardize_catalog_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = out.columns.astype(str).str.strip().str.upper()
    rename_map = {
        "NOMBRE": "NOMBRE",
        "USUARIO STAR (SUGERIDO)": "USUARIO_STAR",
        "USUARIO STAR": "USUARIO_STAR",
        "USUARIO SAC (SUGERIDO)": "USUARIO_SAC",
        "USUARIO SAC": "USUARIO_SAC",
        "TIPO": "TIPO",
    }
    out = out.rename(columns=rename_map)
    required = ["NOMBRE", "USUARIO_STAR", "USUARIO_SAC", "TIPO"]
    missing = [c for c in required if c not in out.columns]
    if missing:
        raise ValueError(f"Al catálogo le faltan columnas: {missing}")
    for col in required:
        out[col] = out[col].apply(norm_text)
    return out[required].copy()


# ============================================================
# UI
# ============================================================
st.title("Comparador STAR vs SAC v2")
st.caption("Clasificación fila por fila de los archivos filtrados. Cada fila recibe un solo estatus final.")

with st.sidebar:
    st.header("Archivos")
    liq_file = st.file_uploader("Liquidaciones (xlsx o csv)", type=["xlsx", "xls", "xlsm", "csv"])
    cont_file = st.file_uploader("Contabilidad (xlsx o csv)", type=["xlsx", "xls", "xlsm", "csv"])
    catalogo_file = st.file_uploader("Catálogo operadores / owners (opcional)", type=["xlsx", "xls", "xlsm", "csv"])

    st.divider()
    st.header("Configuración")
    ndigits = st.number_input("Redondeo de importe", min_value=0, max_value=4, value=2, step=1)
    liq_tipo = st.selectbox("Liquidaciones: Tipo_Concepto", options=["E", "I"], index=0)
    cont_tipo = st.selectbox("Contabilidad: TipoMovimiento", options=["H", "D"], index=0)
    usar_catalogo = st.checkbox("Aplicar filtro por catálogo", value=True)
    enable_relaxed = st.checkbox("Generar sugerencia de match relajado", value=True)
    run_process = st.button("Procesar", type="primary")

if not liq_file or not cont_file:
    st.info("Carga ambos archivos para iniciar.")
    st.stop()

if not run_process:
    st.info("Configura y da clic en Procesar.")
    st.stop()

# ============================================================
# Lectura
# ============================================================
liq_usecols = ["Liquidacion", "Numero_Viaje", "TipoPago", "Monto", "Unidad", "Owner", "Tipo_Concepto"]
cont_usecols = ["Factura", "Referencia", "TipoPago", "Importe", "Unidad", "NombreCuentaContable", "TipoMovimiento"]

try:
    liq = read_table(liq_file, preferred_sheet="LiquidacionesSET_PLUS_datos", usecols=liq_usecols)
    cont = read_table(cont_file, preferred_sheet="ContabilidadSET_PLUS_datos", usecols=cont_usecols)
except Exception as e:
    st.error(f"No pude leer los archivos base. Error: {e}")
    st.stop()

catalogo = None
star_to_nombre = {}
sac_to_nombre = {}
tipos_catalogo_seleccionados = []

if catalogo_file is not None:
    try:
        catalogo_raw = read_table(catalogo_file)
        catalogo = standardize_catalog_columns(catalogo_raw)
        tipos_disponibles = sorted([t for t in catalogo["TIPO"].dropna().unique().tolist() if t != ""])
        with st.sidebar:
            tipos_catalogo_seleccionados = st.multiselect(
                "Tipos del catálogo a considerar",
                options=tipos_disponibles,
                default=["OWNER"] if "OWNER" in tipos_disponibles else tipos_disponibles,
            )
        if tipos_catalogo_seleccionados:
            catalogo = catalogo[catalogo["TIPO"].isin(tipos_catalogo_seleccionados)].copy()
        star_to_nombre = dict(catalogo.loc[catalogo["USUARIO_STAR"] != "", ["USUARIO_STAR", "NOMBRE"]].values)
        sac_to_nombre = dict(catalogo.loc[catalogo["USUARIO_SAC"] != "", ["USUARIO_SAC", "NOMBRE"]].values)
    except Exception as e:
        st.error(f"No pude leer el catálogo. Error: {e}")
        st.stop()

# ============================================================
# Normalización
# ============================================================
liq = liq.rename(columns={
    "Liquidacion": "PR",
    "Numero_Viaje": "VIAJE",
    "TipoPago": "TIPO_PAGO",
    "Monto": "IMPORTE",
    "Unidad": "UNIDAD",
    "Owner": "OWNER_LIQ",
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

liq["IMPORTE"] = pd.to_numeric(liq["IMPORTE"].apply(lambda x: norm_amount(x, ndigits)), errors="coerce")
cont["IMPORTE"] = pd.to_numeric(cont["IMPORTE"].apply(lambda x: norm_amount(x, ndigits)), errors="coerce")

if catalogo is not None:
    liq["OWNER_STD_LIQ"] = liq["OWNER_LIQ"].map(star_to_nombre).fillna("")
    cont["OWNER_STD_CONT"] = cont["OWNER_CONT"].map(sac_to_nombre).fillna("")
else:
    liq["OWNER_STD_LIQ"] = ""
    cont["OWNER_STD_CONT"] = ""

# ============================================================
# Filtros operativos
# ============================================================
liq_f = liq[liq["TIPO_CONCEPTO"] == liq_tipo].copy()
cont_f = cont[cont["TIPO_MOV"] == cont_tipo].copy()

filtro_catalogo_activo = catalogo is not None and usar_catalogo and len(tipos_catalogo_seleccionados) > 0
if filtro_catalogo_activo:
    liq_f = liq_f[liq_f["OWNER_STD_LIQ"] != ""].copy()
    cont_f = cont_f[cont_f["OWNER_STD_CONT"] != ""].copy()

liq_f = liq_f.reset_index(drop=True)
cont_f = cont_f.reset_index(drop=True)
liq_f["ROW_ID_LIQ"] = liq_f.index + 1
cont_f["ROW_ID_CONT"] = cont_f.index + 1

st.subheader("Resumen de carga")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Liquidaciones original", len(liq))
c2.metric("Liquidaciones filtrado", len(liq_f))
c3.metric("Contabilidad original", len(cont))
c4.metric("Contabilidad filtrado", len(cont_f))

# ============================================================
# Match exacto fila a fila
# ============================================================
key_cols = ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE"]
merge_keys = key_cols + ["_seq"]

liq_k = build_seq(liq_f, key_cols)
cont_k = build_seq(cont_f, key_cols)

m = liq_k.merge(
    cont_k,
    how="outer",
    on=merge_keys,
    suffixes=("_LIQ", "_CONT"),
    indicator=True,
)

matched = m[m["_merge"] == "both"].copy()
only_liq = m[m["_merge"] == "left_only"].copy()
only_cont = m[m["_merge"] == "right_only"].copy()

matched["OWNER_MATCH"] = (
    matched["OWNER_LIQ"].astype("string").fillna("") ==
    matched["OWNER_CONT"].astype("string").fillna("")
)
matched["ESTATUS_MATCH"] = matched["OWNER_MATCH"].map({True: "MATCH_OK", False: "MATCH_CON_DISCREPANCIA"})
matched["OBSERVACION"] = matched["OWNER_MATCH"].map({
    True: "Coincide llave exacta y owner.",
    False: "Coincide llave exacta, pero owner es distinto.",
})

only_liq["ESTATUS_MATCH"] = "NO_EXISTE_EN_CONTABILIDAD"
only_liq["OBSERVACION"] = "La fila filtrada de Liquidaciones no encontró contraparte exacta en Contabilidad."
only_cont["ESTATUS_MATCH"] = "NO_EXISTE_EN_LIQUIDACIONES"
only_cont["OBSERVACION"] = "La fila filtrada de Contabilidad no encontró contraparte exacta en Liquidaciones."

# ============================================================
# Traer el estatus de vuelta a cada archivo filtrado completo
# ============================================================
liq_status_from_matched = matched[[
    "ROW_ID_LIQ", "ROW_ID_CONT", "ESTATUS_MATCH", "OBSERVACION",
    "OWNER_CONT", "OWNER_STD_CONT"
]].copy()

liq_status_from_only = only_liq[[
    "ROW_ID_LIQ", "ESTATUS_MATCH", "OBSERVACION"
]].copy()
liq_status_from_only["ROW_ID_CONT"] = pd.NA
liq_status_from_only["OWNER_CONT"] = ""
liq_status_from_only["OWNER_STD_CONT"] = ""

liq_status = pd.concat([liq_status_from_matched, liq_status_from_only], ignore_index=True)
liq_clasificado = liq_f.merge(liq_status, on="ROW_ID_LIQ", how="left")

cont_status_from_matched = matched[[
    "ROW_ID_CONT", "ROW_ID_LIQ", "ESTATUS_MATCH", "OBSERVACION",
    "OWNER_LIQ", "OWNER_STD_LIQ"
]].copy()

cont_status_from_only = only_cont[[
    "ROW_ID_CONT", "ESTATUS_MATCH", "OBSERVACION"
]].copy()
cont_status_from_only["ROW_ID_LIQ"] = pd.NA
cont_status_from_only["OWNER_LIQ"] = ""
cont_status_from_only["OWNER_STD_LIQ"] = ""

cont_status = pd.concat([cont_status_from_matched, cont_status_from_only], ignore_index=True)
cont_clasificado = cont_f.merge(cont_status, on="ROW_ID_CONT", how="left")

# ============================================================
# Banderas útiles por PR y match relajado
# ============================================================
liq_pr_set = set(liq_f["PR"].dropna().astype(str))
cont_pr_set = set(cont_f["PR"].dropna().astype(str))

liq_clasificado["PR_EXISTE_EN_CONT"] = liq_clasificado["PR"].astype(str).isin(cont_pr_set)
cont_clasificado["PR_EXISTE_EN_LIQ"] = cont_clasificado["PR"].astype(str).isin(liq_pr_set)

liq_totales_pr = liq_f.groupby("PR", dropna=False).agg(REG_LIQ=("PR", "size"), IMPORTE_TOTAL_LIQ=("IMPORTE", "sum")).reset_index()
cont_totales_pr = cont_f.groupby("PR", dropna=False).agg(REG_CONT=("PR", "size"), IMPORTE_TOTAL_CONT=("IMPORTE", "sum")).reset_index()

liq_clasificado = liq_clasificado.merge(cont_totales_pr, on="PR", how="left")
cont_clasificado = cont_clasificado.merge(liq_totales_pr, on="PR", how="left")

# Match relajado solo para no encontrados exactos
if enable_relaxed:
    liq_relaxed_keys = only_liq[["ROW_ID_LIQ", "PR", "UNIDAD", "TIPO_PAGO", "IMPORTE", "VIAJE"]].copy()
    cont_relaxed_keys = only_cont[["ROW_ID_CONT", "PR", "UNIDAD", "TIPO_PAGO", "IMPORTE", "VIAJE"]].copy()

    if not liq_relaxed_keys.empty and not cont_relaxed_keys.empty:
        liq_relaxed_keys["REL_KEY"] = (
            liq_relaxed_keys["PR"].astype(str) + "||" +
            liq_relaxed_keys["UNIDAD"].astype(str) + "||" +
            liq_relaxed_keys["TIPO_PAGO"].astype(str) + "||" +
            liq_relaxed_keys["IMPORTE"].astype(str)
        )
        cont_relaxed_keys["REL_KEY"] = (
            cont_relaxed_keys["PR"].astype(str) + "||" +
            cont_relaxed_keys["UNIDAD"].astype(str) + "||" +
            cont_relaxed_keys["TIPO_PAGO"].astype(str) + "||" +
            cont_relaxed_keys["IMPORTE"].astype(str)
        )

        liq_relaxed_keys["MATCH_RELAXED"] = liq_relaxed_keys["REL_KEY"].isin(set(cont_relaxed_keys["REL_KEY"]))
        cont_relaxed_keys["MATCH_RELAXED"] = cont_relaxed_keys["REL_KEY"].isin(set(liq_relaxed_keys["REL_KEY"]))

        liq_clasificado = liq_clasificado.merge(liq_relaxed_keys[["ROW_ID_LIQ", "MATCH_RELAXED"]], on="ROW_ID_LIQ", how="left")
        cont_clasificado = cont_clasificado.merge(cont_relaxed_keys[["ROW_ID_CONT", "MATCH_RELAXED"]], on="ROW_ID_CONT", how="left")
    else:
        liq_clasificado["MATCH_RELAXED"] = False
        cont_clasificado["MATCH_RELAXED"] = False
else:
    liq_clasificado["MATCH_RELAXED"] = False
    cont_clasificado["MATCH_RELAXED"] = False

liq_clasificado["MATCH_RELAXED"] = liq_clasificado["MATCH_RELAXED"].fillna(False)
cont_clasificado["MATCH_RELAXED"] = cont_clasificado["MATCH_RELAXED"].fillna(False)


def motivo_probable_liq(row):
    if row["ESTATUS_MATCH"] == "MATCH_OK":
        return "Match exacto correcto"
    if row["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA":
        return "Owner distinto"
    if bool(row.get("PR_EXISTE_EN_CONT", False)) and bool(row.get("MATCH_RELAXED", False)):
        return "Existe PR en Contabilidad y hay candidato relajado; revisar VIAJE o duplicados"
    if bool(row.get("PR_EXISTE_EN_CONT", False)):
        return "Existe PR en Contabilidad, pero no hubo match exacto"
    return "PR no encontrado en Contabilidad filtrada"


def motivo_probable_cont(row):
    if row["ESTATUS_MATCH"] == "MATCH_OK":
        return "Match exacto correcto"
    if row["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA":
        return "Owner distinto"
    if bool(row.get("PR_EXISTE_EN_LIQ", False)) and bool(row.get("MATCH_RELAXED", False)):
        return "Existe PR en Liquidaciones y hay candidato relajado; revisar VIAJE o duplicados"
    if bool(row.get("PR_EXISTE_EN_LIQ", False)):
        return "Existe PR en Liquidaciones, pero no hubo match exacto"
    return "PR no encontrado en Liquidaciones filtrada"


liq_clasificado["MOTIVO_PROBABLE"] = liq_clasificado.apply(motivo_probable_liq, axis=1)
cont_clasificado["MOTIVO_PROBABLE"] = cont_clasificado.apply(motivo_probable_cont, axis=1)

# ============================================================
# Resúmenes que sí cuadran con el total filtrado de cada lado
# ============================================================
liq_resumen = (
    liq_clasificado["ESTATUS_MATCH"]
    .value_counts(dropna=False)
    .rename_axis("ESTATUS_MATCH")
    .reset_index(name="FILAS")
)
cont_resumen = (
    cont_clasificado["ESTATUS_MATCH"]
    .value_counts(dropna=False)
    .rename_axis("ESTATUS_MATCH")
    .reset_index(name="FILAS")
)

st.divider()
st.subheader("Resumen final que sí cuadra con los filtrados")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Liq MATCH_OK", int((liq_clasificado["ESTATUS_MATCH"] == "MATCH_OK").sum()))
c2.metric("Liq MATCH_CON_DISCREPANCIA", int((liq_clasificado["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA").sum()))
c3.metric("Liq NO_EXISTE_EN_CONTABILIDAD", int((liq_clasificado["ESTATUS_MATCH"] == "NO_EXISTE_EN_CONTABILIDAD").sum()))
c4.metric("Control Liq", f"{len(liq_clasificado):,}")

c5, c6, c7, c8 = st.columns(4)
c5.metric("Cont MATCH_OK", int((cont_clasificado["ESTATUS_MATCH"] == "MATCH_OK").sum()))
c6.metric("Cont MATCH_CON_DISCREPANCIA", int((cont_clasificado["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA").sum()))
c7.metric("Cont NO_EXISTE_EN_LIQUIDACIONES", int((cont_clasificado["ESTATUS_MATCH"] == "NO_EXISTE_EN_LIQUIDACIONES").sum()))
c8.metric("Control Cont", f"{len(cont_clasificado):,}")

col1, col2 = st.columns(2)
with col1:
    st.markdown("**Resumen Liquidaciones**")
    show_df(liq_resumen, height=240, max_rows=100)
    st.caption(f"Validación: {len(liq_clasificado):,} filas clasificadas vs {len(liq_f):,} filas filtradas.")
with col2:
    st.markdown("**Resumen Contabilidad**")
    show_df(cont_resumen, height=240, max_rows=100)
    st.caption(f"Validación: {len(cont_clasificado):,} filas clasificadas vs {len(cont_f):,} filas filtradas.")

# ============================================================
# Vistas de revisión
# ============================================================
st.divider()
st.subheader("Archivos clasificados fila por fila")

t1, t2, t3, t4 = st.tabs([
    f"Liquidaciones clasificadas ({len(liq_clasificado)})",
    f"Contabilidad clasificada ({len(cont_clasificado)})",
    f"Solo discrepancias ({int((liq_clasificado['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum())})",
    "Control por PR",
])

with t1:
    cols = [
        "ROW_ID_LIQ", "PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE",
        "OWNER_LIQ", "OWNER_STD_LIQ", "ESTATUS_MATCH", "OBSERVACION", "MOTIVO_PROBABLE",
        "ROW_ID_CONT", "OWNER_CONT", "OWNER_STD_CONT", "PR_EXISTE_EN_CONT", "MATCH_RELAXED",
        "REG_CONT", "IMPORTE_TOTAL_CONT",
    ]
    show_df(liq_clasificado[[c for c in cols if c in liq_clasificado.columns]], height=500)

with t2:
    cols = [
        "ROW_ID_CONT", "PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE",
        "OWNER_CONT", "OWNER_STD_CONT", "ESTATUS_MATCH", "OBSERVACION", "MOTIVO_PROBABLE",
        "ROW_ID_LIQ", "OWNER_LIQ", "OWNER_STD_LIQ", "PR_EXISTE_EN_LIQ", "MATCH_RELAXED",
        "REG_LIQ", "IMPORTE_TOTAL_LIQ",
    ]
    show_df(cont_clasificado[[c for c in cols if c in cont_clasificado.columns]], height=500)

with t3:
    diff_liq = liq_clasificado[liq_clasificado["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA"].copy()
    cols = [
        "ROW_ID_LIQ", "PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE",
        "OWNER_LIQ", "OWNER_STD_LIQ", "OWNER_CONT", "OWNER_STD_CONT",
        "ESTATUS_MATCH", "OBSERVACION",
    ]
    show_df(diff_liq[[c for c in cols if c in diff_liq.columns]], height=500)

with t4:
    liq_pr = liq_f.groupby("PR", dropna=False).agg(REG_LIQ=("PR", "size"), IMPORTE_LIQ=("IMPORTE", "sum")).reset_index()
    cont_pr = cont_f.groupby("PR", dropna=False).agg(REG_CONT=("PR", "size"), IMPORTE_CONT=("IMPORTE", "sum")).reset_index()
    control_pr = liq_pr.merge(cont_pr, on="PR", how="outer").fillna(0)
    control_pr["DIF_REG"] = control_pr["REG_LIQ"] - control_pr["REG_CONT"]
    control_pr["DIF_IMPORTE"] = control_pr["IMPORTE_LIQ"] - control_pr["IMPORTE_CONT"]
    show_df(control_pr.sort_values(["PR"]), height=500)

# ============================================================
# Exportación
# ============================================================
st.divider()
st.subheader("Descarga")

sheets = {
    "Liq_Clasificadas": liq_clasificado,
    "Cont_Clasificadas": cont_clasificado,
    "Liq_Resumen": liq_resumen,
    "Cont_Resumen": cont_resumen,
    "Match_OK_Liq": liq_clasificado[liq_clasificado["ESTATUS_MATCH"] == "MATCH_OK"],
    "Match_Diff_Liq": liq_clasificado[liq_clasificado["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA"],
    "No_Existe_En_Cont": liq_clasificado[liq_clasificado["ESTATUS_MATCH"] == "NO_EXISTE_EN_CONTABILIDAD"],
    "Match_OK_Cont": cont_clasificado[cont_clasificado["ESTATUS_MATCH"] == "MATCH_OK"],
    "Match_Diff_Cont": cont_clasificado[cont_clasificado["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA"],
    "No_Existe_En_Liq": cont_clasificado[cont_clasificado["ESTATUS_MATCH"] == "NO_EXISTE_EN_LIQUIDACIONES"],
}

excel_bytes = to_excel_bytes(sheets)
st.download_button(
    "⬇️ Descargar Excel clasificado",
    data=excel_bytes,
    file_name="comparador_star_sac_v2_clasificado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
