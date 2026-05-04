import re
import unicodedata
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Saldos Owner - Costos con Vales", layout="wide")

# ============================================================
# Helpers generales
# ============================================================

def norm_text(x: object) -> str:
    if x is None or pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s


def norm_for_key(x: object) -> str:
    s = norm_text(x)
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def norm_amount(x: object, ndigits: int = 2) -> float:
    try:
        if x is None or pd.isna(x):
            return float("nan")
        if isinstance(x, str):
            x = x.replace(",", "").replace("$", "").strip()
        return round(float(x), ndigits)
    except Exception:
        return float("nan")


def strip_concept_suffix(x: object) -> str:
    """Quita sufijos tipo ' - 20170908' sin destruir el concepto base."""
    s = norm_text(x)
    s = re.sub(r"\s+-\s+\d+.*$", "", s)
    s = re.sub(r"\s+-\s+[A-Z0-9]+.*$", "", s)
    return s.strip()


def canonical_concept(x: object, concept_map: dict[str, str] | None = None) -> str:
    """Normaliza conceptos. No requiere catalogo; el catalogo solo mejora equivalencias futuras."""
    s = strip_concept_suffix(x)
    k = norm_for_key(s)
    if concept_map and k in concept_map:
        return concept_map[k]

    # Reglas base muy conservadoras. Se pueden ampliar despues con catalogo.
    rules = [
        (r"\bPERSONAL LOAN\b|\bLOAN\b|\bPRESTAMO\b", "LOAN/PERSONAL LOAN"),
        (r"\bDIESEL\b|\bCONSUMIBLES\b", "CXP DIESEL/CONSUMIBLES"),
        (r"\bANTICIPO\b|\bADVANCE\b", "CXP ANTICIPO"),
    ]
    for pattern, value in rules:
        if re.search(pattern, k):
            return value
    return k


def read_table(file_obj, preferred_sheet: str | None = None, usecols=None) -> pd.DataFrame:
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


def resolve_col(df: pd.DataFrame, candidates: list[str], required: bool = True) -> str | None:
    normalized = {norm_for_key(c): c for c in df.columns}
    for c in candidates:
        key = norm_for_key(c)
        if key in normalized:
            return normalized[key]
    if required:
        raise ValueError(f"No encontre ninguna columna de estas opciones: {candidates}. Columnas disponibles: {list(df.columns)}")
    return None


def resolve_all_cols(df: pd.DataFrame, candidates: list[str]) -> list[str]:
    wanted = {norm_for_key(c) for c in candidates}
    out = []
    for col in df.columns:
        base = re.sub(r"\.\d+$", "", str(col))
        if norm_for_key(base) in wanted:
            out.append(col)
    return out


def choose_cont_import_col(cont_raw: pd.DataFrame) -> str:
    """
    En Contabilidad puede haber dos columnas Importe:
    - una de encabezado/poliza total
    - otra del movimiento individual, que es la correcta para cruzar.
    Pandas suele renombrar duplicados como Importe e Importe.1.
    Preferimos el ultimo Importe disponible.
    """
    importe_cols = resolve_all_cols(cont_raw, ["Importe", "Monto", "Total"])
    if not importe_cols:
        raise ValueError("No encontre columna de importe en Contabilidad.")
    return importe_cols[-1]


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter", engine_kwargs={"options": {"constant_memory": True}}) as writer:
        for name, df in sheets.items():
            out = df.copy()
            out.columns = [str(c)[:250] for c in out.columns]
            out.to_excel(writer, sheet_name=name[:31], index=False)
    bio.seek(0)
    return bio.getvalue()


def show_df(df: pd.DataFrame, height: int = 560):
    st.dataframe(df, use_container_width=True, height=height)


def load_concept_map(file_obj) -> dict[str, str]:
    if file_obj is None:
        return {}
    df = read_table(file_obj)
    src = resolve_col(df, ["concepto_origen", "concepto", "ingles", "english", "source"], required=False)
    dst = resolve_col(df, ["concepto_canonico", "canonico", "espanol", "spanish", "target"], required=False)
    if not src or not dst:
        st.warning("El catalogo de conceptos debe tener columnas tipo concepto_origen y concepto_canonico. Se ignoro el catalogo.")
        return {}
    return {norm_for_key(a): norm_for_key(b) for a, b in zip(df[src], df[dst]) if norm_for_key(a) and norm_for_key(b)}


# ============================================================
# Preparacion de archivos
# ============================================================

def prep_contabilidad(cont_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str], tipo_mov: str | None = "D") -> tuple[pd.DataFrame, dict[str, str]]:
    c_mov = resolve_col(cont_raw, ["TipoMovimiento", "Movimiento", "Tipo Movimiento"])
    c_importe = choose_cont_import_col(cont_raw)
    c_unidad = resolve_col(cont_raw, ["Unidad", "Numero de Unidad", "Numero_Unidad"])
    c_referencia = resolve_col(cont_raw, ["Referencia", "Numero_Viaje", "Numero Viaje", "Viaje"], required=False)
    c_poliza = resolve_col(cont_raw, ["Clave Poliza", "Clave Póliza", "ClavePoliza", "Factura", "Contrarrecibo"])
    c_concepto = resolve_col(cont_raw, ["Concepto detalle", "Concepto Detalle", "Concepto", "NombreCuentaContable"], required=False)
    c_vale = resolve_col(cont_raw, ["Vale", "No Vale", "Numero Vale"], required=False)

    out = cont_raw.copy()
    out["TIPO_MOV"] = out[c_mov].apply(norm_text)
    if tipo_mov is not None:
        out = out[out["TIPO_MOV"] == norm_text(tipo_mov)].copy()
    out["POLIZA_KEY"] = out[c_poliza].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["VIAJE_KEY"] = out[c_referencia].apply(norm_for_key) if c_referencia else ""
    out["VALE_KEY"] = out[c_vale].apply(norm_for_key) if c_vale else ""
    out["CONCEPTO_KEY"] = out[c_concepto].apply(lambda x: canonical_concept(x, concept_map)) if c_concepto else ""
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, ndigits))
    out["ROW_ID_CONT"] = range(1, len(out) + 1)

    colmap = {
        "movimiento": c_mov,
        "importe_movimiento_usado": c_importe,
        "unidad": c_unidad,
        "referencia": c_referencia or "",
        "poliza": c_poliza,
        "concepto": c_concepto or "",
        "vale": c_vale or "",
    }
    return out, colmap


def prep_base_saldos(base_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str]) -> pd.DataFrame:
    c_poliza = resolve_col(base_raw, ["FOLIO_CONTRARRECIBO", "Contrarrecibo", "Clave Poliza", "Clave Póliza", "ClavePoliza", "Factura"])
    c_unidad = resolve_col(base_raw, ["NUMERO_UNIDAD", "Unidad", "Numero de Unidad", "Numero_Unidad"])
    c_viaje = resolve_col(base_raw, ["NUMERO_VIAJE", "Referencia", "Numero_Viaje", "Numero Viaje", "Viaje"], required=False)
    c_concepto = resolve_col(base_raw, ["Tipo concepto", "Concepto contabilidad", "Concepto detalle", "Concepto Detalle", "Concepto", "NombreCuentaContable"], required=False)
    c_importe = resolve_col(base_raw, ["Importe", "Total", "Monto"])

    out = base_raw.copy()
    out["POLIZA_KEY"] = out[c_poliza].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["VIAJE_KEY"] = out[c_viaje].apply(norm_for_key) if c_viaje else ""
    out["CONCEPTO_KEY"] = out[c_concepto].apply(lambda x: canonical_concept(x, concept_map)) if c_concepto else ""
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, ndigits))
    out["ROW_ID_BASE"] = range(1, len(out) + 1)
    return out


def prep_vales(vales_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str]) -> pd.DataFrame:
    """
    Prepara el archivo de Vales usando:
    - Unidad
    - Total (como importe)
    - Contrarecibo (como poliza/contrarrecibo)
    - Concepto
    """
    c_unidad = resolve_col(vales_raw, ["Unidad", "Numero de Unidad", "Numero_Unidad"])
    c_importe = resolve_col(vales_raw, ["Total", "Importe", "TotalVale"])
    c_contrarrecibo = resolve_col(vales_raw, ["Contrarecibo", "Contrarrecibo", "Clave Poliza"])
    c_concepto = resolve_col(vales_raw, ["Concepto", "Concepto detalle"], required=False)
    c_vale = resolve_col(vales_raw, ["Vale", "No Vale", "Numero Vale"], required=False)

    out = vales_raw.copy()
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, ndigits))
    out["POLIZA_KEY"] = out[c_contrarrecibo].apply(norm_for_key)
    out["CONCEPTO_KEY"] = out[c_concepto].apply(lambda x: canonical_concept(x, concept_map)) if c_concepto else ""
    out["VALE_KEY"] = out[c_vale].apply(norm_for_key) if c_vale else ""
    out["VIAJE_KEY"] = ""  # Los vales no tienen viaje
    out["ROW_ID_VALE"] = range(1, len(out) + 1)
    out["ORIGEN"] = "VALES"
    return out


# ============================================================
# Matching
# ============================================================

def score_match(
    cand_poliza_key: str,
    cand_unidad_key: str,
    cand_viaje_key: str,
    cand_concepto_key: str,
    cand_importe_key: float,
    ref_poliza_key: str,
    ref_unidad_key: str,
    ref_viaje_key: str,
    ref_concepto_key: str,
    ref_importe_key: float,
) -> dict[str, object]:
    """
    Evalua 5 criterios: poliza, unidad, viaje, concepto, importe.
    Devuelve dict con booleanos de coincidencia y score total.
    """
    coin_pol = bool(cand_poliza_key and ref_poliza_key and cand_poliza_key == ref_poliza_key)
    coin_uni = bool(cand_unidad_key and ref_unidad_key and cand_unidad_key == ref_unidad_key)
    coin_via = bool(cand_viaje_key and ref_viaje_key and cand_viaje_key == ref_viaje_key)
    coin_con = bool(cand_concepto_key and ref_concepto_key and cand_concepto_key == ref_concepto_key)
    coin_imp = False
    if pd.notna(cand_importe_key) and pd.notna(ref_importe_key):
        coin_imp = round(float(cand_importe_key), 2) == round(float(ref_importe_key), 2)
    score = sum([coin_pol, coin_uni, coin_via, coin_con, coin_imp])
    return {
        "COINCIDE_POLIZA": coin_pol,
        "COINCIDE_UNIDAD": coin_uni,
        "COINCIDE_VIAJE": coin_via,
        "COINCIDE_CONCEPTO": coin_con,
        "COINCIDE_IMPORTE": coin_imp,
        "TOTAL_COINCIDENCIAS": score,
    }


def match_base_vs_cont_mayoria(base: pd.DataFrame, cont_d: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Cruza Base Saldos vs Contabilidad D usando regla de mayoria: 5 criterios.
    MATCH_OK si coinciden 5 de 5, MATCH_CON_DISCREPANCIA si 3 o 4, NO_EXISTE si <3.
    """
    if base.empty or cont_d.empty:
        return base, cont_d, pd.DataFrame(), pd.DataFrame()

    scored = []
    for _, r in base.iterrows():
        candidates = cont_d[
            (cont_d["POLIZA_KEY"] == r["POLIZA_KEY"]) |
            (cont_d["UNIDAD_KEY"] == r["UNIDAD_KEY"]) |
            (cont_d["VIAJE_KEY"] == r["VIAJE_KEY"]) if r.get("VIAJE_KEY") else False |
            (cont_d["CONCEPTO_KEY"] == r["CONCEPTO_KEY"]) if r.get("CONCEPTO_KEY") else False
        ].copy()
        if candidates.empty:
            continue
        for _, c in candidates.iterrows():
            s = score_match(
                c["POLIZA_KEY"], c["UNIDAD_KEY"], c["VIAJE_KEY"], c["CONCEPTO_KEY"], c["IMPORTE_KEY"],
                r["POLIZA_KEY"], r["UNIDAD_KEY"], r.get("VIAJE_KEY", ""), r.get("CONCEPTO_KEY", ""), r["IMPORTE_KEY"],
            )
            if s["TOTAL_COINCIDENCIAS"] >= 3:
                scored.append({"ROW_ID_BASE": r["ROW_ID_BASE"], "ROW_ID_CONT": c["ROW_ID_CONT"], **s})

    if not scored:
        base_clas = base.copy()
        base_clas["ESTATUS_MATCH"] = "NO_EXISTE_EN_CONTABILIDAD_D"
        base_clas["TOTAL_COINCIDENCIAS"] = 0
        return base_clas, cont_d, pd.DataFrame(), pd.DataFrame()

    scored_df = pd.DataFrame(scored)
    best = scored_df.loc[scored_df.groupby("ROW_ID_BASE")["TOTAL_COINCIDENCIAS"].idxmax()].copy()
    best["ESTATUS_MATCH"] = best["TOTAL_COINCIDENCIAS"].apply(lambda x: "MATCH_OK" if x == 5 else "MATCH_CON_DISCREPANCIA")

    base_clas = base.merge(
        best[["ROW_ID_BASE", "ROW_ID_CONT", "ESTATUS_MATCH", "TOTAL_COINCIDENCIAS", "COINCIDE_POLIZA", "COINCIDE_UNIDAD", "COINCIDE_VIAJE", "COINCIDE_CONCEPTO", "COINCIDE_IMPORTE"]],
        on="ROW_ID_BASE",
        how="left",
    )
    base_clas["ESTATUS_MATCH"] = base_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_CONTABILIDAD_D")
    base_clas["TOTAL_COINCIDENCIAS"] = base_clas["TOTAL_COINCIDENCIAS"].fillna(0).astype(int)

    cont_status = best[["ROW_ID_CONT", "ESTATUS_MATCH", "TOTAL_COINCIDENCIAS"]].copy()
    cont_clas = cont_d.merge(cont_status, on="ROW_ID_CONT", how="left")
    cont_clas["ESTATUS_MATCH"] = cont_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_BASE_SALDOS")
    cont_clas["TOTAL_COINCIDENCIAS"] = cont_clas["TOTAL_COINCIDENCIAS"].fillna(0).astype(int)
    return base_clas, cont_clas, scored_df, best


def match_vales_vs_cont_mayoria(vales: pd.DataFrame, cont_d: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Cruza Vales vs Contabilidad D usando regla de mayoria: 5 criterios.
    MATCH_OK si coinciden 5 de 5, MATCH_CON_DISCREPANCIA si 3 o 4, NO_EXISTE si <3.
    """
    if vales.empty or cont_d.empty:
        return vales, cont_d, pd.DataFrame(), pd.DataFrame()

    scored = []
    for _, r in vales.iterrows():
        candidates = cont_d[
            (cont_d["POLIZA_KEY"] == r["POLIZA_KEY"]) |
            (cont_d["UNIDAD_KEY"] == r["UNIDAD_KEY"]) |
            (cont_d["CONCEPTO_KEY"] == r["CONCEPTO_KEY"]) if r.get("CONCEPTO_KEY") else False |
            (cont_d["VALE_KEY"] == r["VALE_KEY"]) if r.get("VALE_KEY") else False
        ].copy()
        if candidates.empty:
            continue
        for _, c in candidates.iterrows():
            s = score_match(
                c["POLIZA_KEY"], c["UNIDAD_KEY"], c.get("VIAJE_KEY", ""), c["CONCEPTO_KEY"], c["IMPORTE_KEY"],
                r["POLIZA_KEY"], r["UNIDAD_KEY"], r.get("VIAJE_KEY", ""), r.get("CONCEPTO_KEY", ""), r["IMPORTE_KEY"],
            )
            if s["TOTAL_COINCIDENCIAS"] >= 3:
                scored.append({"ROW_ID_VALE": r["ROW_ID_VALE"], "ROW_ID_CONT": c["ROW_ID_CONT"], **s})

    if not scored:
        vales_clas = vales.copy()
        vales_clas["ESTATUS_MATCH"] = "NO_EXISTE_EN_CONTABILIDAD_D"
        vales_clas["TOTAL_COINCIDENCIAS"] = 0
        return vales_clas, cont_d, pd.DataFrame(), pd.DataFrame()

    scored_df = pd.DataFrame(scored)
    best = scored_df.loc[scored_df.groupby("ROW_ID_VALE")["TOTAL_COINCIDENCIAS"].idxmax()].copy()
    best["ESTATUS_MATCH"] = best["TOTAL_COINCIDENCIAS"].apply(lambda x: "MATCH_OK" if x == 5 else "MATCH_CON_DISCREPANCIA")

    vales_clas = vales.merge(
        best[["ROW_ID_VALE", "ROW_ID_CONT", "ESTATUS_MATCH", "TOTAL_COINCIDENCIAS", "COINCIDE_POLIZA", "COINCIDE_UNIDAD", "COINCIDE_VIAJE", "COINCIDE_CONCEPTO", "COINCIDE_IMPORTE"]],
        on="ROW_ID_VALE",
        how="left",
    )
    vales_clas["ESTATUS_MATCH"] = vales_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_CONTABILIDAD_D")
    vales_clas["TOTAL_COINCIDENCIAS"] = vales_clas["TOTAL_COINCIDENCIAS"].fillna(0).astype(int)

    cont_status = best[["ROW_ID_CONT", "ESTATUS_MATCH", "TOTAL_COINCIDENCIAS"]].copy()
    cont_clas = cont_d.merge(cont_status, on="ROW_ID_CONT", how="left")
    cont_clas["ESTATUS_MATCH"] = cont_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_VALES")
    cont_clas["TOTAL_COINCIDENCIAS"] = cont_clas["TOTAL_COINCIDENCIAS"].fillna(0).astype(int)
    return vales_clas, cont_clas, scored_df, best


def resumen_dh_contabilidad(cont_all: pd.DataFrame) -> pd.DataFrame:
    if cont_all.empty:
        return pd.DataFrame()
    base = cont_all.copy()
    base["IMPORTE_KEY"] = pd.to_numeric(base["IMPORTE_KEY"], errors="coerce").fillna(0)
    g = base.groupby(["POLIZA_KEY", "UNIDAD_KEY", "CONCEPTO_KEY"], dropna=False).agg(
        TOTAL_D=("IMPORTE_KEY", lambda s: s[base.loc[s.index, "TIPO_MOV"] == "D"].sum()),
        TOTAL_H=("IMPORTE_KEY", lambda s: s[base.loc[s.index, "TIPO_MOV"] == "H"].sum()),
        MOVIMIENTOS=("IMPORTE_KEY", "size"),
    ).reset_index()
    g["SALDO_D_MENOS_H"] = g["TOTAL_D"] - g["TOTAL_H"]
    g["ESTATUS_DH"] = g["SALDO_D_MENOS_H"].apply(lambda x: "SALDADO_D_H" if round(float(x), 2) == 0 else "CON_SALDO")
    return g


# ============================================================
# UI
# ============================================================

st.title("Saldos Owner - Desarrollo de Costos (Versión Vales)")
st.caption("Version con Vales: utiliza el archivo de Vales en lugar de Cheques y Vouchers.")

with st.expander("Como leer esta version", expanded=True):
    st.markdown(
        """
        **Base Saldos vs Contabilidad D** usa 5 criterios: poliza/contrarrecibo, unidad, viaje/referencia, concepto flexible e importe.

        - **MATCH_OK**: coinciden 5 de 5 criterios.
        - **MATCH_CON_DISCREPANCIA**: coinciden 3 o 4 de 5 criterios. Es candidato probable, pero hay algo que revisar.
        - **NO_EXISTE_EN_CONTABILIDAD_D**: no encontro candidato con al menos 3 criterios.

        **Vales vs Contabilidad D** usa las columnas: Unidad, Total (importe), Contrarecibo (póliza) y Concepto.

        La tabla muestra columnas de diagnostico como `COINCIDE_POLIZA`, `COINCIDE_UNIDAD`, `COINCIDE_VIAJE`, `COINCIDE_CONCEPTO`, `COINCIDE_IMPORTE` y `TOTAL_COINCIDENCIAS` para que no tengas que adivinar que fallo.
        """
    )

with st.sidebar:
    st.header("Archivos")
    cont_file = st.file_uploader("Contabilidad", type=["xlsx", "xls", "xlsm", "csv"])
    base_file = st.file_uploader("Base Saldos corregida", type=["xlsx", "xls", "xlsm", "csv"])
    vales_file = st.file_uploader("Vales", type=["xlsx", "xls", "xlsm", "csv"])
    concept_file = st.file_uploader("Catalogo conceptos opcional", type=["xlsx", "xls", "xlsm", "csv"])
    st.divider()
    ndigits = st.number_input("Redondeo de importe", min_value=0, max_value=4, value=2, step=1)
    proceso = st.radio("Proceso", ["Base Saldos vs Contabilidad D", "Vales vs Contabilidad D", "Ambos"], index=0)
    run = st.button("Procesar costos", type="primary")

if not run:
    st.info("Carga los archivos y da clic en Procesar costos.")
    st.stop()

if cont_file is None:
    st.error("Carga el archivo de Contabilidad.")
    st.stop()

concept_map = load_concept_map(concept_file)
try:
    cont_raw = read_table(cont_file, preferred_sheet="ContabilidadSET_PLUS_datos")
    cont_d, cont_colmap = prep_contabilidad(cont_raw, ndigits, concept_map, tipo_mov="D")
    cont_all, _ = prep_contabilidad(cont_raw, ndigits, concept_map, tipo_mov=None)
except Exception as e:
    st.error(f"No pude preparar Contabilidad: {e}")
    st.stop()

st.subheader("Contabilidad filtrada a movimientos D")
c1, c2 = st.columns(2)
c1.metric("Movimientos D", f"{len(cont_d):,}")
c2.metric("Columna de importe usada", cont_colmap["importe_movimiento_usado"])
st.caption("Si Contabilidad trae dos columnas Importe, esta version usa la ultima columna Importe detectada, que normalmente es el importe del movimiento individual.")

result_sheets: dict[str, pd.DataFrame] = {"Contabilidad_D": cont_d, "Contabilidad_todos_movs": cont_all, "Columnas_usadas_cont": pd.DataFrame([cont_colmap])}

if proceso in {"Base Saldos vs Contabilidad D", "Ambos"}:
    st.divider()
    st.header("1) Base Saldos vs Contabilidad D")
    if base_file is None:
        st.warning("Falta Base Saldos corregida.")
    else:
        try:
            base_raw = read_table(base_file)
            base = prep_base_saldos(base_raw, ndigits, concept_map)
            base_clas, cont_base_clas, candidatos_base, mejores_base = match_base_vs_cont_mayoria(base, cont_d)
            result_sheets.update({
                "Base_clasificada": base_clas,
                "Cont_vs_Base": cont_base_clas,
                "Candidatos_Base_Cont": candidatos_base,
                "Mejores_matches_Base": mejores_base,
            })
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Base filas", f"{len(base):,}")
            c2.metric("MATCH_OK", f"{int((base_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum()):,}")
            c3.metric("MATCH_CON_DISCREPANCIA", f"{int((base_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum()):,}")
            c4.metric("No existe en Cont D", f"{int((base_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D').sum()):,}")

            t1, t2, t3, t4 = st.tabs(["Base clasificada", "Contabilidad contra Base", "Candidatos tecnicos", "Mejores matches"])
            with t1:
                show_df(base_clas)
            with t2:
                show_df(cont_base_clas)
            with t3:
                show_df(candidatos_base)
            with t4:
                show_df(mejores_base)
        except Exception as e:
            st.error(f"No pude procesar Base Saldos: {e}")

if proceso in {"Vales vs Contabilidad D", "Ambos"}:
    st.divider()
    st.header("2) Vales vs Contabilidad D")
    if vales_file is None:
        st.warning("Falta archivo de Vales.")
    else:
        try:
            vales_raw = read_table(vales_file)
            vales = prep_vales(vales_raw, ndigits, concept_map)
            vales_clas, cont_vales_clas, candidatos_vales, mejores_vales = match_vales_vs_cont_mayoria(vales, cont_d)
            resumen_dh = resumen_dh_contabilidad(cont_all)
            result_sheets.update({
                "Vales_clasificados": vales_clas,
                "Cont_vs_Vales": cont_vales_clas,
                "Candidatos_Vales_Cont": candidatos_vales,
                "Mejores_matches_Vales": mejores_vales,
                "Resumen_DH_Contabilidad": resumen_dh,
            })
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Vales totales", f"{len(vales):,}")
            c2.metric("MATCH_OK", f"{int((vales_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum()):,}")
            c3.metric("MATCH_CON_DISCREPANCIA", f"{int((vales_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum()):,}")
            c4.metric("No existe en Cont D", f"{int((vales_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D').sum()):,}")

            t1, t2, t3, t4, t5 = st.tabs(["Vales clasificados", "Contabilidad contra Vales", "Candidatos tecnicos", "Mejores matches", "Resumen D/H"])
            with t1:
                show_df(vales_clas)
            with t2:
                show_df(cont_vales_clas)
            with t3:
                show_df(candidatos_vales)
            with t4:
                show_df(mejores_vales)
            with t5:
                show_df(resumen_dh)
        except Exception as e:
            st.error(f"No pude procesar Vales: {e}")

if len(result_sheets) > 1:
    st.divider()
    st.download_button(
        "Descargar resultado costos en Excel",
        data=to_excel_bytes(result_sheets),
        file_name="resultado_costos_owner_vales.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
