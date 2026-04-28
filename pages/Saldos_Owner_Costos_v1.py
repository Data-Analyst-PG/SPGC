import re
import unicodedata
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Saldos Owner - Costos", layout="wide")


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
    s = norm_text(x)
    s = re.sub(r"\s+-\s+\d+.*$", "", s)
    s = re.sub(r"\s+-\s+[A-Z0-9]+.*$", "", s)
    return s.strip()


def canonical_concept(x: object, concept_map: dict[str, str] | None = None) -> str:
    s = strip_concept_suffix(x)
    k = norm_for_key(s)
    if concept_map and k in concept_map:
        return concept_map[k]
    rules = [
        (r"\bPERSONAL LOAN\b|\bLOAN\b|\bPRESTAMO\b", "LOAN/PERSONAL LOAN"),
        (r"\bDIESEL\b|\bCONSUMIBLES\b", "CXP DIESEL/CONSUMIBLES"),
        (r"\bANTICIPO\b|\bADVANCE\b", "CXP ANTICIPO"),
    ]
    for pattern, value in rules:
        if re.search(pattern, k):
            return value
    return k


def build_seq(df: pd.DataFrame, key_cols: list[str], seq_col: str = "_seq") -> pd.DataFrame:
    out = df.copy()
    out[seq_col] = out.groupby(key_cols, dropna=False).cumcount() + 1
    return out


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
        raise ValueError(f"No encontré ninguna columna de estas opciones: {candidates}. Columnas disponibles: {list(df.columns)}")
    return None


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
        st.warning("El catálogo de conceptos debe tener columnas tipo concepto_origen y concepto_canonico. Se ignoró el catálogo.")
        return {}
    return {norm_for_key(a): norm_for_key(b) for a, b in zip(df[src], df[dst]) if norm_for_key(a) and norm_for_key(b)}


def prep_contabilidad(cont_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str]) -> pd.DataFrame:
    c_mov = resolve_col(cont_raw, ["TipoMovimiento", "Movimiento", "Tipo Movimiento"])
    c_importe = resolve_col(cont_raw, ["Importe", "Monto", "Total"])
    c_unidad = resolve_col(cont_raw, ["Unidad", "Numero de Unidad", "Numero_Unidad"])
    c_referencia = resolve_col(cont_raw, ["Referencia", "Numero_Viaje", "Numero Viaje", "Viaje"], required=False)
    c_poliza = resolve_col(cont_raw, ["Clave Poliza", "Clave Póliza", "ClavePoliza", "Factura", "Contrarrecibo"])
    c_concepto = resolve_col(cont_raw, ["Concepto detalle", "Concepto Detalle", "Concepto", "NombreCuentaContable"], required=False)
    c_vale = resolve_col(cont_raw, ["Vale", "No Vale", "Numero Vale"], required=False)

    out = cont_raw.copy()
    out["TIPO_MOV"] = out[c_mov].apply(norm_text)
    out = out[out["TIPO_MOV"] == "D"].copy()
    out["POLIZA_KEY"] = out[c_poliza].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["VIAJE_KEY"] = out[c_referencia].apply(norm_for_key) if c_referencia else ""
    out["VALE_KEY"] = out[c_vale].apply(norm_for_key) if c_vale else ""
    out["CONCEPTO_KEY"] = out[c_concepto].apply(lambda x: canonical_concept(x, concept_map)) if c_concepto else ""
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, ndigits))
    out["ROW_ID_CONT"] = range(1, len(out) + 1)
    return out


def prep_base_saldos(base_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str]) -> pd.DataFrame:
    c_poliza = resolve_col(base_raw, ["folio_contrarecibo", "folio contrarecibo", "contrarecibo", "contrarrecibo"])
    c_unidad = resolve_col(base_raw, ["numero de unidad", "numero_unidad", "unidad"])
    c_viaje = resolve_col(base_raw, ["numero_viaje", "numero viaje", "referencia", "viaje"])
    c_concepto = resolve_col(base_raw, ["concepto_contabilidad", "concepto contabilidad", "concepto"])
    c_importe = resolve_col(base_raw, ["importe", "monto", "total"])

    out = base_raw.copy()
    out["POLIZA_KEY"] = out[c_poliza].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["VIAJE_KEY"] = out[c_viaje].apply(norm_for_key)
    out["CONCEPTO_KEY"] = out[c_concepto].apply(lambda x: canonical_concept(x, concept_map))
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, ndigits))
    out["ROW_ID_BASE"] = range(1, len(out) + 1)
    return out


def match_base_vs_cont(base: pd.DataFrame, cont_d: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    keys = ["POLIZA_KEY", "UNIDAD_KEY", "VIAJE_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"]
    b = build_seq(base, keys)
    c = build_seq(cont_d, keys)
    m = b.merge(c, how="outer", on=keys + ["_seq"], suffixes=("_BASE", "_CONT"), indicator=True)
    m["ESTATUS_MATCH"] = m["_merge"].map({"both": "MATCH_OK", "left_only": "NO_EXISTE_EN_CONTABILIDAD_D", "right_only": "NO_EXISTE_EN_BASE_SALDOS"})
    base_status = m[m["_merge"].isin(["both", "left_only"])][["ROW_ID_BASE", "ROW_ID_CONT", "ESTATUS_MATCH"]].copy()
    cont_status = m[m["_merge"].isin(["both", "right_only"])][["ROW_ID_CONT", "ROW_ID_BASE", "ESTATUS_MATCH"]].copy()
    return base.merge(base_status, on="ROW_ID_BASE", how="left"), cont_d.merge(cont_status, on="ROW_ID_CONT", how="left"), m


def prep_cheques(cheques_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    c_vale = resolve_col(cheques_raw, ["Vale"])
    c_unidad = resolve_col(cheques_raw, ["Unidad"])
    c_concepto = resolve_col(cheques_raw, ["Concepto"])
    c_cheque = resolve_col(cheques_raw, ["Cheque", "No Cheque", "Numero Cheque"], required=False)
    c_obs = resolve_col(cheques_raw, ["Observaciones", "Observacion"], required=False)
    c_contra = resolve_col(cheques_raw, ["Contrarrecibo", "Contrarecibo", "Contrarrecibo"])
    c_importe = resolve_col(cheques_raw, ["Importe", "Monto", "Total"])
    c_cargo = resolve_col(cheques_raw, ["CargoA", "Cargo A", "Cargo_a"], required=False)

    out = cheques_raw.copy()
    out["SOURCE"] = "CHEQUES"
    out["ROW_ID_ORIGEN"] = range(1, len(out) + 1)
    out["CARGOA_KEY"] = out[c_cargo].apply(norm_for_key) if c_cargo else ""
    excluidos_company = out[out["CARGOA_KEY"] == "COMPANY"].copy()
    out = out[out["CARGOA_KEY"] != "COMPANY"].copy()
    out["VALE_KEY"] = out[c_vale].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["CONCEPTO_KEY"] = out[c_concepto].apply(lambda x: canonical_concept(x, concept_map))
    cheque_obs = (out[c_cheque].fillna("").astype(str) if c_cheque else "") + " " + (out[c_obs].fillna("").astype(str) if c_obs else "")
    out["OBS_KEY"] = pd.Series(cheque_obs, index=out.index).apply(norm_for_key)
    out["POLIZA_KEY"] = out[c_contra].apply(norm_for_key)
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, ndigits))
    return out, excluidos_company


def prep_vouchers(vouchers_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str]) -> pd.DataFrame:
    c_vale = resolve_col(vouchers_raw, ["Vale"])
    c_unidad = resolve_col(vouchers_raw, ["Unidad"])
    c_concepto = resolve_col(vouchers_raw, ["Concepto"])
    c_obs = resolve_col(vouchers_raw, ["Observaciones", "Observacion"], required=False)
    c_contra = resolve_col(vouchers_raw, ["Contrarrecibo", "Contrarecibo", "Contrarrecibo"])
    c_total = resolve_col(vouchers_raw, ["Total", "Importe", "Monto"])

    out = vouchers_raw.copy()
    out["SOURCE"] = "VOUCHERS"
    out["ROW_ID_ORIGEN"] = range(1, len(out) + 1)
    out["VALE_KEY"] = out[c_vale].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["CONCEPTO_KEY"] = out[c_concepto].apply(lambda x: canonical_concept(x, concept_map))
    out["OBS_KEY"] = out[c_obs].apply(norm_for_key) if c_obs else ""
    out["POLIZA_KEY"] = out[c_contra].apply(norm_for_key)
    out["IMPORTE_KEY"] = out[c_total].apply(lambda x: norm_amount(x, ndigits))
    return out


def dedupe_costos(cheques: pd.DataFrame, vouchers: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    costos = pd.concat([cheques, vouchers], ignore_index=True, sort=False)
    costos["ROW_ID_COSTO"] = range(1, len(costos) + 1)
    dup_keys = ["VALE_KEY", "UNIDAD_KEY", "CONCEPTO_KEY", "OBS_KEY", "POLIZA_KEY", "IMPORTE_KEY"]
    costos["DUP_COUNT"] = costos.groupby(dup_keys, dropna=False)["ROW_ID_COSTO"].transform("size")
    costos["DUP_SEQ"] = costos.groupby(dup_keys, dropna=False).cumcount() + 1
    duplicados = costos[costos["DUP_COUNT"] > 1].copy()
    depurados = costos.sort_values(["DUP_SEQ", "SOURCE"]).drop_duplicates(dup_keys, keep="first").copy()
    return depurados, duplicados


def match_costos_vs_cont(costos: pd.DataFrame, cont_d: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    keys = ["VALE_KEY", "UNIDAD_KEY", "CONCEPTO_KEY", "POLIZA_KEY", "IMPORTE_KEY"]
    cst = build_seq(costos, keys)
    cont = build_seq(cont_d, keys)
    m = cst.merge(cont, how="outer", on=keys + ["_seq"], suffixes=("_COSTO", "_CONT"), indicator=True)
    m["ESTATUS_MATCH"] = m["_merge"].map({"both": "MATCH_OK", "left_only": "NO_EXISTE_EN_CONTABILIDAD_D", "right_only": "NO_EXISTE_EN_COSTOS_DEPURADOS"})
    costos_status = m[m["_merge"].isin(["both", "left_only"])][["ROW_ID_COSTO", "ROW_ID_CONT", "ESTATUS_MATCH"]].copy()
    cont_status = m[m["_merge"].isin(["both", "right_only"])][["ROW_ID_CONT", "ROW_ID_COSTO", "ESTATUS_MATCH"]].copy()
    return costos.merge(costos_status, on="ROW_ID_COSTO", how="left"), cont_d.merge(cont_status, on="ROW_ID_CONT", how="left"), m


st.title("Saldos Owner - Desarrollo de Costos")
st.caption("Primera versión para validar Base Saldos, Cheques y Vouchers contra movimientos D de Contabilidad.")

with st.sidebar:
    st.header("Archivos")
    cont_file = st.file_uploader("Contabilidad", type=["xlsx", "xls", "xlsm", "csv"])
    base_file = st.file_uploader("Base Saldos corregida", type=["xlsx", "xls", "xlsm", "csv"])
    cheques_file = st.file_uploader("Cheques", type=["xlsx", "xls", "xlsm", "csv"])
    vouchers_file = st.file_uploader("Vouchers", type=["xlsx", "xls", "xlsm", "csv"])
    concept_file = st.file_uploader("Catálogo conceptos inglés/español opcional", type=["xlsx", "xls", "xlsm", "csv"])
    st.divider()
    ndigits = st.number_input("Redondeo de importe", min_value=0, max_value=4, value=2, step=1)
    proceso = st.radio("Proceso", ["Base Saldos vs Contabilidad D", "Cheques/Vouchers vs Contabilidad D", "Ambos"], index=2)
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
    cont_d = prep_contabilidad(cont_raw, ndigits, concept_map)
except Exception as e:
    st.error(f"No pude preparar Contabilidad: {e}")
    st.stop()

st.subheader("Contabilidad filtrada a movimientos D")
st.metric("Movimientos D", f"{len(cont_d):,}")

result_sheets: dict[str, pd.DataFrame] = {"Contabilidad_D": cont_d}

if proceso in {"Base Saldos vs Contabilidad D", "Ambos"}:
    st.divider()
    st.header("1) Base Saldos vs Contabilidad D")
    if base_file is None:
        st.warning("Falta Base Saldos corregida.")
    else:
        try:
            base_raw = read_table(base_file)
            base = prep_base_saldos(base_raw, ndigits, concept_map)
            base_clas, cont_base_clas, base_match = match_base_vs_cont(base, cont_d)
            result_sheets.update({"Base_clasificada": base_clas, "Cont_vs_Base": cont_base_clas, "Match_Base_Cont": base_match})
            c1, c2, c3 = st.columns(3)
            c1.metric("Base filas", f"{len(base):,}")
            c2.metric("Base MATCH_OK", f"{int((base_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum()):,}")
            c3.metric("Base no existe en Cont D", f"{int((base_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D').sum()):,}")
            t1, t2, t3 = st.tabs(["Base clasificada", "Contabilidad contra Base", "Cruce técnico"])
            with t1:
                show_df(base_clas)
            with t2:
                show_df(cont_base_clas)
            with t3:
                show_df(base_match)
        except Exception as e:
            st.error(f"No pude procesar Base Saldos: {e}")

if proceso in {"Cheques/Vouchers vs Contabilidad D", "Ambos"}:
    st.divider()
    st.header("2) Cheques/Vouchers vs Contabilidad D")
    if cheques_file is None or vouchers_file is None:
        st.warning("Faltan Cheques y/o Vouchers.")
    else:
        try:
            cheques_raw = read_table(cheques_file)
            vouchers_raw = read_table(vouchers_file)
            cheques, cheques_company = prep_cheques(cheques_raw, ndigits, concept_map)
            vouchers = prep_vouchers(vouchers_raw, ndigits, concept_map)
            costos_depurados, costos_duplicados = dedupe_costos(cheques, vouchers)
            costos_clas, cont_costos_clas, costos_match = match_costos_vs_cont(costos_depurados, cont_d)
            result_sheets.update({
                "Cheques_excluidos_Company": cheques_company,
                "Costos_depurados": costos_depurados,
                "Costos_duplicados": costos_duplicados,
                "Costos_clasificados": costos_clas,
                "Cont_vs_Costos": cont_costos_clas,
                "Match_Costos_Cont": costos_match,
            })
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Cheques excluidos Company", f"{len(cheques_company):,}")
            c2.metric("Costos depurados", f"{len(costos_depurados):,}")
            c3.metric("Duplicados detectados", f"{len(costos_duplicados):,}")
            c4.metric("Costos MATCH_OK", f"{int((costos_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum()):,}")
            t1, t2, t3, t4, t5 = st.tabs(["Costos clasificados", "Duplicados", "Excluidos Company", "Contabilidad contra Costos", "Cruce técnico"])
            with t1:
                show_df(costos_clas)
            with t2:
                show_df(costos_duplicados)
            with t3:
                show_df(cheques_company)
            with t4:
                show_df(cont_costos_clas)
            with t5:
                show_df(costos_match)
        except Exception as e:
            st.error(f"No pude procesar Cheques/Vouchers: {e}")

if len(result_sheets) > 1:
    st.divider()
    st.download_button(
        "Descargar resultado costos en Excel",
        data=to_excel_bytes(result_sheets),
        file_name="resultado_costos_owner.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
