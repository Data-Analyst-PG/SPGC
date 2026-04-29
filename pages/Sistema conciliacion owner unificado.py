"""
Sistema Unificado de Conciliación de Saldos Owner - VERSIÓN FINAL COMPLETA

Une:
- Script INGRESOS (9___Saldos_Owner__5_.py) → 100% funcional
- Script COSTOS (Saldos_Owner_Costos_v1.py) → 100% funcional  

Genera 16+ hojas Excel para auditoría contable
"""

import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# ======================================================================================
# CONFIGURACIÓN
# ======================================================================================

@dataclass
class Config:
    ndigits: int = 2

# ======================================================================================
# NORMALIZACIÓN (Compatible con ambos scripts)
# ======================================================================================

def norm_text(x: Any) -> str:
    if x is None or pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s

def norm_for_key(x: Any) -> str:
    s = norm_text(x)
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def norm_amount(x: Any, ndigits: int = 2) -> float:
    try:
        if x is None or pd.isna(x):
            return float("nan")
        if isinstance(x, str):
            x = x.replace(",", "").replace("$", "").strip()
        return round(float(x), ndigits)
    except Exception:
        return float("nan")

def strip_concept_suffix(x: Any) -> str:
    s = norm_text(x)
    s = re.sub(r"\s+-\s+\d+.*$", "", s)
    s = re.sub(r"\s+-\s+[A-Z0-9]+.*$", "", s)
    return s.strip()

def canonical_concept(x: Any) -> str:
    s = strip_concept_suffix(x)
    k = norm_for_key(s)
    rules = [
        (r"\bPERSONAL LOAN\b|\bLOAN\b|\bPRESTAMO\b", "LOAN/PERSONAL LOAN"),
        (r"\bDIESEL\b|\bCONSUMIBLES\b", "CXP DIESEL/CONSUMIBLES"),
        (r"\bANTICIPO\b|\bADVANCE\b", "CXP ANTICIPO"),
    ]
    for pattern, value in rules:
        if re.search(pattern, k):
            return value
    return k

def build_seq(df: pd.DataFrame, key_cols: List[str], seq_col: str = "_seq") -> pd.DataFrame:
    out = df.copy()
    out[seq_col] = out.groupby(key_cols, dropna=False).cumcount() + 1
    return out

def resolve_col(df: pd.DataFrame, candidates: List[str], required: bool = True) -> Optional[str]:
    normalized = {norm_for_key(c): c for c in df.columns}
    for c in candidates:
        key = norm_for_key(c)
        if key in normalized:
            return normalized[key]
    if required:
        raise ValueError(f"No encontré columna de: {candidates}")
    return None

def choose_cont_import_col(cont_raw: pd.DataFrame) -> str:
    importe_cols = []
    for col in cont_raw.columns:
        base = re.sub(r'\.\d+$', '', str(col))
        if norm_for_key(base) in [norm_for_key('Importe'), norm_for_key('Monto'), norm_for_key('Total')]:
            importe_cols.append(col)
    if not importe_cols:
        raise ValueError("No encontré columna de importe")
    return importe_cols[-1]

# ======================================================================================
# MANEJO DE ARCHIVOS
# ======================================================================================

def read_table(file_obj, preferred_sheet: Optional[str] = None, usecols=None) -> pd.DataFrame:
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

def exportar_excel(sheets: Dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine='xlsxwriter', engine_kwargs={'options': {'constant_memory': True}}) as writer:
        for nombre, df in sheets.items():
            df_out = df.copy()
            df_out.columns = [str(c)[:250] for c in df_out.columns]
            df_out.to_excel(writer, sheet_name=nombre[:31], index=False)
    bio.seek(0)
    return bio.getvalue()

# ======================================================================================
# PREPARADORES
# ======================================================================================

def prep_contabilidad(cont_raw: pd.DataFrame, config: Config, tipo_mov: Optional[str] = None) -> pd.DataFrame:
    c_mov = resolve_col(cont_raw, ["TipoMovimiento", "Movimiento"])
    c_importe = choose_cont_import_col(cont_raw)
    c_unidad = resolve_col(cont_raw, ["Unidad", "Numero de Unidad"])
    c_referencia = resolve_col(cont_raw, ["Referencia", "Numero_Viaje", "Viaje"], required=False)
    c_poliza = resolve_col(cont_raw, ["Clave Poliza", "ClavePoliza", "Factura", "Contrarrecibo"])
    c_concepto = resolve_col(cont_raw, ["Concepto detalle", "Concepto Detalle", "Concepto", "NombreCuentaContable"], required=False)
    c_vale = resolve_col(cont_raw, ["Vale", "No Vale"], required=False)
    c_tipo_pago = resolve_col(cont_raw, ["TipoPago", "Tipo Pago"], required=False)
    
    out = cont_raw.copy()
    out["TIPO_MOV"] = out[c_mov].apply(norm_text)
    
    if tipo_mov is not None:
        out = out[out["TIPO_MOV"] == norm_text(tipo_mov)].copy()
    
    out["POLIZA_KEY"] = out[c_poliza].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["VIAJE_KEY"] = out[c_referencia].apply(norm_for_key) if c_referencia else ""
    out["VALE_KEY"] = out[c_vale].apply(norm_for_key) if c_vale else ""
    out["CONCEPTO_KEY"] = out[c_concepto].apply(canonical_concept) if c_concepto else ""
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, config.ndigits))
    out["TIPO_PAGO_KEY"] = out[c_tipo_pago].apply(norm_for_key) if c_tipo_pago else ""
    out["ROW_ID_CONT"] = range(1, len(out) + 1)
    
    if c_concepto:
        out["OWNER_CONT"] = out[c_concepto].apply(norm_text)
    
    return out

def prep_liquidaciones(liq_raw: pd.DataFrame, config: Config, tipo_concepto: str = 'E') -> pd.DataFrame:
    out = liq_raw.rename(columns={
        "Liquidacion": "PR",
        "Numero_Viaje": "VIAJE",
        "TipoPago": "TIPO_PAGO",
        "Monto": "IMPORTE",
        "Unidad": "UNIDAD",
        "Owner": "OWNER_LIQ",
        "Tipo_Concepto": "TIPO_CONCEPTO",
    })
    
    for c in ["PR", "VIAJE", "TIPO_PAGO", "UNIDAD", "OWNER_LIQ", "TIPO_CONCEPTO"]:
        if c in out.columns:
            out[c] = out[c].apply(norm_text)
    
    out["IMPORTE"] = pd.to_numeric(out["IMPORTE"].apply(lambda x: norm_amount(x, config.ndigits)), errors="coerce")
    out = out[out["TIPO_CONCEPTO"] == tipo_concepto].copy()
    out = out.reset_index(drop=True)
    out["ROW_ID_LIQ"] = out.index + 1
    
    return out

def prep_base_saldos(base_raw: pd.DataFrame, config: Config) -> pd.DataFrame:
    c_poliza = resolve_col(base_raw, ["folio_contrarecibo", "folio contrarecibo", "contrarecibo", "contrarrecibo"])
    c_unidad = resolve_col(base_raw, ["numero de unidad", "numero_unidad", "unidad"])
    c_viaje = resolve_col(base_raw, ["numero_viaje", "numero viaje", "referencia", "viaje"])
    c_concepto = resolve_col(base_raw, ["concepto_contabilidad", "concepto contabilidad", "concepto"])
    c_importe = resolve_col(base_raw, ["importe", "monto", "total"])
    
    out = base_raw.copy()
    out["POLIZA_KEY"] = out[c_poliza].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["VIAJE_KEY"] = out[c_viaje].apply(norm_for_key)
    out["CONCEPTO_KEY"] = out[c_concepto].apply(canonical_concept)
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, config.ndigits))
    out["ROW_ID_BASE"] = range(1, len(out) + 1)
    
    return out

def prep_cheques(cheques_raw: pd.DataFrame, config: Config) -> Tuple[pd.DataFrame, pd.DataFrame]:
    c_vale = resolve_col(cheques_raw, ["Vale"])
    c_unidad = resolve_col(cheques_raw, ["Unidad"])
    c_concepto = resolve_col(cheques_raw, ["Concepto"])
    c_cheque = resolve_col(cheques_raw, ["Cheque", "No Cheque"], required=False)
    c_obs = resolve_col(cheques_raw, ["Observaciones", "Observacion"], required=False)
    c_contra = resolve_col(cheques_raw, ["Contrarrecibo", "Contrarecibo"])
    c_importe = resolve_col(cheques_raw, ["Importe", "Monto", "Total"])
    c_cargo = resolve_col(cheques_raw, ["CargoA", "Cargo A"], required=False)
    
    out = cheques_raw.copy()
    out["SOURCE"] = "CHEQUES"
    out["ROW_ID_ORIGEN"] = range(1, len(out) + 1)
    out["CARGOA_KEY"] = out[c_cargo].apply(norm_for_key) if c_cargo else ""
    
    excluidos = out[out["CARGOA_KEY"] == "COMPANY"].copy()
    out = out[out["CARGOA_KEY"] != "COMPANY"].copy()
    
    out["VALE_KEY"] = out[c_vale].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["CONCEPTO_KEY"] = out[c_concepto].apply(canonical_concept)
    cheque_obs = (out[c_cheque].fillna("").astype(str) if c_cheque else "") + " " + (out[c_obs].fillna("").astype(str) if c_obs else "")
    out["OBS_KEY"] = pd.Series(cheque_obs, index=out.index).apply(norm_for_key)
    out["POLIZA_KEY"] = out[c_contra].apply(norm_for_key)
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, config.ndigits))
    
    return out, excluidos

def prep_vouchers(vouchers_raw: pd.DataFrame, config: Config) -> Tuple[pd.DataFrame, pd.DataFrame]:
    c_vale = resolve_col(vouchers_raw, ["Vale"])
    c_unidad = resolve_col(vouchers_raw, ["Unidad"])
    c_concepto = resolve_col(vouchers_raw, ["Concepto"])
    c_obs = resolve_col(vouchers_raw, ["Observaciones", "Observacion"], required=False)
    c_contra = resolve_col(vouchers_raw, ["Contrarrecibo", "Contrarecibo"], required=False)
    c_total = resolve_col(vouchers_raw, ["Total", "Importe", "Monto"])
    c_operador = resolve_col(vouchers_raw, ["Operador", "Owner"], required=False)
    
    out = vouchers_raw.copy()
    out["SOURCE"] = "VOUCHERS"
    out["ROW_ID_ORIGEN"] = range(1, len(out) + 1)
    out["OPERADOR_KEY"] = out[c_operador].apply(norm_for_key) if c_operador else ""
    
    excluidos = out[out["OPERADOR_KEY"].str.contains(r"\bFILIAL\b", na=False)].copy() if c_operador else out.iloc[0:0].copy()
    out = out[~out.index.isin(excluidos.index)].copy()
    
    out["VALE_KEY"] = out[c_vale].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["CONCEPTO_KEY"] = out[c_concepto].apply(canonical_concept)
    out["OBS_KEY"] = out[c_obs].apply(norm_for_key) if c_obs else ""
    out["POLIZA_KEY"] = out[c_contra].apply(norm_for_key) if c_contra else ""
    out["IMPORTE_KEY"] = out[c_total].apply(lambda x: norm_amount(x, config.ndigits))
    
    return out, excluidos

# ======================================================================================
# MOTOR INGRESOS (Merge exacto - Script original)
# ======================================================================================

def match_liquidaciones_vs_cont(liq: pd.DataFrame, cont: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    st.info(f"🔍 Matching Liquidaciones ({len(liq):,}) vs Contabilidad ({len(cont):,})...")
    
    # Renombrar contabilidad para que coincidan las llaves
    cont_for_merge = cont.rename(columns={
        'POLIZA_KEY': 'PR',
        'VIAJE_KEY': 'VIAJE',
        'UNIDAD_KEY': 'UNIDAD',
        'IMPORTE_KEY': 'IMPORTE',
        'TIPO_PAGO_KEY': 'TIPO_PAGO'
    })
    
    # Merge exacto por 5 campos + secuencia
    key_cols = ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE"]
    merge_keys = key_cols + ["_seq"]
    
    liq_k = build_seq(liq, key_cols)
    cont_k = build_seq(cont_for_merge, key_cols)
    
    m = liq_k.merge(cont_k, how="outer", on=merge_keys, suffixes=("_LIQ", "_CONT"), indicator=True)
    
    matched = m[m["_merge"] == "both"].copy()
    only_liq = m[m["_merge"] == "left_only"].copy()
    only_cont = m[m["_merge"] == "right_only"].copy()
    
    st.success(f"✅ MATCH: {len(matched):,} | Solo Liq: {len(only_liq):,} | Solo Cont: {len(only_cont):,}")
    
    # Clasificar por owner
    if "OWNER_LIQ" in matched.columns and "OWNER_CONT" in matched.columns:
        matched["OWNER_MATCH"] = (matched["OWNER_LIQ"].astype("string").fillna("") == matched["OWNER_CONT"].astype("string").fillna(""))
        matched["ESTATUS_MATCH"] = matched["OWNER_MATCH"].map({True: "MATCH_OK", False: "MATCH_CON_DISCREPANCIA"})
    else:
        matched["ESTATUS_MATCH"] = "MATCH_OK"
    
    only_liq["ESTATUS_MATCH"] = "NO_EXISTE_EN_CONTABILIDAD"
    only_cont["ESTATUS_MATCH"] = "NO_EXISTE_EN_LIQUIDACIONES"
    
    # Reconstruir clasificados
    liq_status_matched = matched[["ROW_ID_LIQ", "ROW_ID_CONT", "ESTATUS_MATCH"]].copy()
    liq_status_only = only_liq[["ROW_ID_LIQ", "ESTATUS_MATCH"]].copy()
    liq_status_only["ROW_ID_CONT"] = pd.NA
    
    liq_status = pd.concat([liq_status_matched, liq_status_only], ignore_index=True)
    liq_clas = liq.merge(liq_status, on="ROW_ID_LIQ", how="left")
    
    cont_status_matched = matched[["ROW_ID_CONT", "ROW_ID_LIQ", "ESTATUS_MATCH"]].copy()
    cont_status_only = only_cont[["ROW_ID_CONT", "ESTATUS_MATCH"]].copy()
    cont_status_only["ROW_ID_LIQ"] = pd.NA
    
    cont_status = pd.concat([cont_status_matched, cont_status_only], ignore_index=True)
    cont_clas = cont.merge(cont_status, on="ROW_ID_CONT", how="left")
    
    return liq_clas, cont_clas

# ======================================================================================
# MOTOR COSTOS (Bloques + Scoring + Greedy - Script original)
# ======================================================================================

def _pairs_by_block(left: pd.DataFrame, right: pd.DataFrame, block_cols: List[str], left_id: str, right_id: str) -> pd.DataFrame:
    l = left[[left_id] + block_cols].copy()
    r = right[[right_id] + block_cols].copy()
    for c in block_cols:
        l = l[l[c].notna() & (l[c].astype(str) != "")]
        r = r[r[c].notna() & (r[c].astype(str) != "")]
    if l.empty or r.empty:
        return pd.DataFrame(columns=[left_id, right_id])
    p = l.merge(r, on=block_cols, how="inner")[[left_id, right_id]]
    return p.drop_duplicates()

def make_candidate_pairs(left: pd.DataFrame, right: pd.DataFrame, left_id: str, right_id: str, mode: str) -> pd.DataFrame:
    if mode == "base":
        blocks = [
            ["POLIZA_KEY", "IMPORTE_KEY"],
            ["POLIZA_KEY", "UNIDAD_KEY"],
            ["UNIDAD_KEY", "VIAJE_KEY", "IMPORTE_KEY"],
            ["POLIZA_KEY", "VIAJE_KEY"],
            ["UNIDAD_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"],
        ]
    else:  # mode == "costos"
        blocks = [
            ["UNIDAD_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"],
            ["UNIDAD_KEY", "IMPORTE_KEY"],
            ["CONCEPTO_KEY", "IMPORTE_KEY"],
            ["VALE_KEY", "IMPORTE_KEY"],
            ["VALE_KEY", "UNIDAD_KEY"],
            ["POLIZA_KEY", "IMPORTE_KEY"],
            ["POLIZA_KEY", "UNIDAD_KEY"],
        ]
    
    pieces = [_pairs_by_block(left, right, b, left_id, right_id) for b in blocks]
    pieces = [p for p in pieces if not p.empty]
    if not pieces:
        return pd.DataFrame(columns=[left_id, right_id])
    return pd.concat(pieces, ignore_index=True).drop_duplicates()

def score_pairs_base(base: pd.DataFrame, cont: pd.DataFrame, pairs: pd.DataFrame) -> pd.DataFrame:
    if pairs.empty:
        return pairs
    b = base[["ROW_ID_BASE", "POLIZA_KEY", "UNIDAD_KEY", "VIAJE_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"]]
    c = cont[["ROW_ID_CONT", "POLIZA_KEY", "UNIDAD_KEY", "VIAJE_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"]]
    x = pairs.merge(b, on="ROW_ID_BASE").merge(c, on="ROW_ID_CONT", suffixes=("_BASE", "_CONT"))
    
    x["COINCIDE_POLIZA"] = x["POLIZA_KEY_BASE"] == x["POLIZA_KEY_CONT"]
    x["COINCIDE_UNIDAD"] = x["UNIDAD_KEY_BASE"] == x["UNIDAD_KEY_CONT"]
    x["COINCIDE_VIAJE"] = x["VIAJE_KEY_BASE"] == x["VIAJE_KEY_CONT"]
    x["COINCIDE_CONCEPTO"] = x["CONCEPTO_KEY_BASE"] == x["CONCEPTO_KEY_CONT"]
    x["COINCIDE_IMPORTE"] = x["IMPORTE_KEY_BASE"] == x["IMPORTE_KEY_CONT"]
    
    score_cols = ["COINCIDE_POLIZA", "COINCIDE_UNIDAD", "COINCIDE_VIAJE", "COINCIDE_CONCEPTO", "COINCIDE_IMPORTE"]
    x["TOTAL_COINCIDENCIAS"] = x[score_cols].sum(axis=1).astype(int)
    x["ESTATUS_MATCH"] = x["TOTAL_COINCIDENCIAS"].map(lambda n: "MATCH_OK" if n == 5 else ("MATCH_CON_DISCREPANCIA" if n >= 3 else "CANDIDATO_DEBIL"))
    
    return x

def score_pairs_costos(costos: pd.DataFrame, cont: pd.DataFrame, pairs: pd.DataFrame) -> pd.DataFrame:
    if pairs.empty:
        return pairs
    
    lcols = ["ROW_ID_COSTO", "VALE_KEY", "UNIDAD_KEY", "CONCEPTO_KEY", "POLIZA_KEY", "IMPORTE_KEY"]
    rcols = ["ROW_ID_CONT", "VALE_KEY", "UNIDAD_KEY", "CONCEPTO_KEY", "POLIZA_KEY", "IMPORTE_KEY"]
    x = pairs.merge(costos[lcols], on="ROW_ID_COSTO").merge(cont[rcols], on="ROW_ID_CONT", suffixes=("_COSTO", "_CONT"))
    
    def has_value(v):
        return pd.notna(v) and str(v).strip().upper() not in {"", "NULL", "NONE", "NAN"}
    
    criteria = []
    for name, lcol, rcol in [
        ("VALE", "VALE_KEY_COSTO", "VALE_KEY_CONT"),
        ("UNIDAD", "UNIDAD_KEY_COSTO", "UNIDAD_KEY_CONT"),
        ("CONCEPTO", "CONCEPTO_KEY_COSTO", "CONCEPTO_KEY_CONT"),
        ("POLIZA", "POLIZA_KEY_COSTO", "POLIZA_KEY_CONT"),
        ("IMPORTE", "IMPORTE_KEY_COSTO", "IMPORTE_KEY_CONT"),
    ]:
        eval_col = f"EVALUA_{name}"
        ok_col = f"COINCIDE_{name}"
        x[eval_col] = x.apply(lambda r: has_value(r.get(lcol)) and has_value(r.get(rcol)), axis=1)
        x[ok_col] = x.apply(lambda r: bool(r[eval_col]) and r.get(lcol) == r.get(rcol), axis=1)
        criteria.append((eval_col, ok_col))
    
    eval_cols = [a for a, _ in criteria]
    ok_cols = [b for _, b in criteria]
    x["CRITERIOS_EVALUADOS"] = x[eval_cols].sum(axis=1).astype(int)
    x["TOTAL_COINCIDENCIAS"] = x[ok_cols].sum(axis=1).astype(int)
    
    def estatus(row):
        if row["CRITERIOS_EVALUADOS"] >= 3 and row["TOTAL_COINCIDENCIAS"] == row["CRITERIOS_EVALUADOS"]:
            return "MATCH_OK"
        if row["TOTAL_COINCIDENCIAS"] >= 3:
            return "MATCH_CON_DISCREPANCIA"
        return "CANDIDATO_DEBIL"
    
    x["ESTATUS_MATCH"] = x.apply(estatus, axis=1)
    return x

def greedy_best_match(scored: pd.DataFrame, left_id: str, right_id: str) -> pd.DataFrame:
    candidates = scored[scored["TOTAL_COINCIDENCIAS"] >= 3].copy()
    if candidates.empty:
        return candidates
    
    sort_cols = ["TOTAL_COINCIDENCIAS"]
    ascending = [False]
    if "CRITERIOS_EVALUADOS" in candidates.columns:
        sort_cols.append("CRITERIOS_EVALUADOS")
        ascending.append(False)
    sort_cols += [left_id, right_id]
    ascending += [True, True]
    
    candidates = candidates.sort_values(sort_cols, ascending=ascending)
    
    used_left = set()
    used_right = set()
    rows = []
    
    for _, row in candidates.iterrows():
        lid = row[left_id]
        rid = row[right_id]
        if lid not in used_left and rid not in used_right:
            used_left.add(lid)
            used_right.add(rid)
            rows.append(row)
    
    if not rows:
        return candidates.iloc[0:0].copy()
    return pd.DataFrame(rows)

def match_base_vs_cont(base: pd.DataFrame, cont_d: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    st.info(f"🔍 Matching Base Saldos ({len(base):,}) vs Contabilidad D ({len(cont_d):,})...")
    
    pairs = make_candidate_pairs(base, cont_d, "ROW_ID_BASE", "ROW_ID_CONT", mode="base")
    if pairs.empty:
        st.warning("No se generaron pares candidatos")
        base_clas = base.copy()
        base_clas["ESTATUS_MATCH"] = "NO_EXISTE_EN_CONTABILIDAD_D"
        base_clas["TOTAL_COINCIDENCIAS"] = 0
        cont_clas = cont_d.copy()
        cont_clas["ESTATUS_MATCH"] = "NO_EXISTE_EN_BASE"
        return base_clas, cont_clas, pd.DataFrame(), pd.DataFrame()
    
    scored = score_pairs_base(base, cont_d, pairs)
    best = greedy_best_match(scored, "ROW_ID_BASE", "ROW_ID_CONT")
    
    score_cols = ["COINCIDE_POLIZA", "COINCIDE_UNIDAD", "COINCIDE_VIAJE", "COINCIDE_CONCEPTO", "COINCIDE_IMPORTE", "TOTAL_COINCIDENCIAS", "ESTATUS_MATCH"]
    base_status = best[["ROW_ID_BASE", "ROW_ID_CONT"] + score_cols].copy() if not best.empty else pd.DataFrame(columns=["ROW_ID_BASE", "ROW_ID_CONT"] + score_cols)
    
    base_clas = base.merge(base_status, on="ROW_ID_BASE", how="left")
    base_clas["ESTATUS_MATCH"] = base_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_CONTABILIDAD_D")
    base_clas["TOTAL_COINCIDENCIAS"] = base_clas["TOTAL_COINCIDENCIAS"].fillna(0).astype(int)
    
    cont_status = best[["ROW_ID_CONT", "ROW_ID_BASE"] + score_cols].copy() if not best.empty else pd.DataFrame(columns=["ROW_ID_CONT", "ROW_ID_BASE"] + score_cols)
    cont_clas = cont_d.merge(cont_status, on="ROW_ID_CONT", how="left")
    cont_clas["ESTATUS_MATCH"] = cont_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_BASE")
    cont_clas["TOTAL_COINCIDENCIAS"] = cont_clas["TOTAL_COINCIDENCIAS"].fillna(0).astype(int)
    
    st.success(f"✅ Matches: {len(best):,}")
    
    return base_clas, cont_clas, scored, best

def dedupe_costos(cheques: pd.DataFrame, vouchers: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    st.info("🔍 Deduplicando Cheques + Vouchers...")
    
    costos = pd.concat([cheques, vouchers], ignore_index=True, sort=False)
    costos["ROW_ID_COSTO"] = range(1, len(costos) + 1)
    
    dup_keys = ["VALE_KEY", "UNIDAD_KEY", "CONCEPTO_KEY", "OBS_KEY", "POLIZA_KEY", "IMPORTE_KEY"]
    costos["DUP_COUNT"] = costos.groupby(dup_keys, dropna=False)["ROW_ID_COSTO"].transform("size")
    costos["DUP_SEQ"] = costos.groupby(dup_keys, dropna=False).cumcount() + 1
    
    duplicados = costos[costos["DUP_COUNT"] > 1].copy()
    depurados = costos.sort_values(["DUP_SEQ", "SOURCE"]).drop_duplicates(dup_keys, keep="first").copy()
    
    st.success(f"✅ Depurados: {len(depurados):,} | Duplicados: {len(duplicados):,}")
    
    return depurados, duplicados

def match_costos_vs_cont(costos: pd.DataFrame, cont_d: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    st.info(f"🔍 Matching Costos ({len(costos):,}) vs Contabilidad D ({len(cont_d):,})...")
    
    pairs = make_candidate_pairs(costos, cont_d, "ROW_ID_COSTO", "ROW_ID_CONT", mode="costos")
    if pairs.empty:
        st.warning("No se generaron pares candidatos")
        costos_clas = costos.copy()
        costos_clas["ESTATUS_MATCH"] = "NO_EXISTE_EN_CONTABILIDAD_D"
        cont_clas = cont_d.copy()
        cont_clas["ESTATUS_MATCH"] = "NO_EXISTE_EN_COSTOS"
        return costos_clas, cont_clas, pd.DataFrame(), pd.DataFrame()
    
    scored = score_pairs_costos(costos, cont_d, pairs)
    best = greedy_best_match(scored, "ROW_ID_COSTO", "ROW_ID_CONT")
    
    score_cols = [
        "COINCIDE_VALE", "COINCIDE_UNIDAD", "COINCIDE_CONCEPTO", "COINCIDE_POLIZA", "COINCIDE_IMPORTE",
        "EVALUA_VALE", "EVALUA_UNIDAD", "EVALUA_CONCEPTO", "EVALUA_POLIZA", "EVALUA_IMPORTE",
        "CRITERIOS_EVALUADOS", "TOTAL_COINCIDENCIAS", "ESTATUS_MATCH",
    ]
    
    costo_status = best[["ROW_ID_COSTO", "ROW_ID_CONT"] + [c for c in score_cols if c in best.columns]].copy() if not best.empty else pd.DataFrame(columns=["ROW_ID_COSTO", "ROW_ID_CONT"] + score_cols)
    
    costos_clas = costos.merge(costo_status, on="ROW_ID_COSTO", how="left")
    costos_clas["ESTATUS_MATCH"] = costos_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_CONTABILIDAD_D")
    costos_clas["TOTAL_COINCIDENCIAS"] = costos_clas["TOTAL_COINCIDENCIAS"].fillna(0).astype(int)
    
    cont_status = best[["ROW_ID_CONT", "ROW_ID_COSTO"] + [c for c in score_cols if c in best.columns]].copy() if not best.empty else pd.DataFrame(columns=["ROW_ID_CONT", "ROW_ID_COSTO"] + score_cols)
    cont_clas = cont_d.merge(cont_status, on="ROW_ID_CONT", how="left")
    cont_clas["ESTATUS_MATCH"] = cont_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_COSTOS")
    cont_clas["TOTAL_COINCIDENCIAS"] = cont_clas["TOTAL_COINCIDENCIAS"].fillna(0).astype(int)
    
    st.success(f"✅ Matches: {len(best):,}")
    
    return costos_clas, cont_clas, scored, best

# ======================================================================================
# ANÁLISIS D vs H
# ======================================================================================

def analizar_DH(cont_completa: pd.DataFrame) -> pd.DataFrame:
    st.info("⚖️ Analizando saldos D vs H...")
    
    if cont_completa.empty:
        return pd.DataFrame()
    
    base = cont_completa.copy()
    base["IMPORTE_KEY"] = pd.to_numeric(base["IMPORTE_KEY"], errors='coerce').fillna(0)
    
    resumen = base.groupby(['POLIZA_KEY', 'UNIDAD_KEY', 'CONCEPTO_KEY'], dropna=False).agg(
        TOTAL_D=('IMPORTE_KEY', lambda s: s[base.loc[s.index, 'TIPO_MOV'] == 'D'].sum()),
        TOTAL_H=('IMPORTE_KEY', lambda s: s[base.loc[s.index, 'TIPO_MOV'] == 'H'].sum()),
        MOVIMIENTOS_D=('IMPORTE_KEY', lambda s: (base.loc[s.index, 'TIPO_MOV'] == 'D').sum()),
        MOVIMIENTOS_H=('IMPORTE_KEY', lambda s: (base.loc[s.index, 'TIPO_MOV'] == 'H').sum()),
    ).reset_index()
    
    resumen['SALDO_D_MENOS_H'] = resumen['TOTAL_D'] - resumen['TOTAL_H']
    resumen['ESTATUS_DH'] = resumen['SALDO_D_MENOS_H'].apply(
        lambda x: 'SALDADO' if abs(x) < 0.01 else ('PENDIENTE_PAGO' if x > 0 else 'SOBREPAGO')
    )
    
    st.success(f"✅ Analizado: {len(resumen):,} grupos")
    
    return resumen

# ======================================================================================
# INTERFAZ STREAMLIT
# ======================================================================================

def main():
    st.set_page_config(page_title="Sistema Conciliación Owner - FINAL", layout="wide")
    
    st.title("🔄 Sistema Unificado de Conciliación de Saldos Owner")
    st.caption("VERSIÓN FINAL COMPLETA - Integra INGRESOS + COSTOS + Análisis D vs H")
    
    with st.sidebar:
        st.header("⚙️ Configuración")
        config = Config(ndigits=st.number_input("Decimales", 0, 4, 2))
        
        st.divider()
        st.header("📁 Archivos")
        
        cont_file = st.file_uploader("📊 Contabilidad (obligatorio)", type=['xlsx', 'xls', 'csv'])
        
        st.subheader("INGRESOS")
        liq_file = st.file_uploader("📈 Liquidaciones", type=['xlsx', 'xls', 'csv'])
        tipo_concepto = st.selectbox("Tipo_Concepto", ["E", "I"], 0)
        
        st.subheader("COSTOS")
        base_file = st.file_uploader("📋 Base Saldos", type=['xlsx', 'xls', 'csv'])
        cheques_file = st.file_uploader("💵 Cheques", type=['xlsx', 'xls', 'csv'])
        vouchers_file = st.file_uploader("🎫 Vouchers", type=['xlsx', 'xls', 'csv'])
        
        st.divider()
        procesar_btn = st.button("▶️ PROCESAR", type="primary", use_container_width=True)
    
    with st.expander("ℹ️ Sistema Unificado", expanded=False):
        st.markdown("""
        ### ✅ Módulos Integrados
        
        **INGRESOS**: Liquidaciones vs Contabilidad H (merge exacto)
        **COSTOS - Base**: Base Saldos vs Contabilidad D (bloques + scoring)
        **COSTOS - Operativo**: Cheques/Vouchers vs Contabilidad D (con dedup)
        **ANÁLISIS**: D vs H (saldos liquidados)
        
        ### 📊 Salidas (16+ hojas Excel)
        - Todas las clasificaciones con ESTATUS_MATCH
        - Columnas diagnóstico (COINCIDE_*, EVALUA_*)
        - Scoring completo para auditoría
        - Resumen general
        """)
    
    if not procesar_btn:
        st.info("👆 Carga archivos y da clic en PROCESAR")
        return
    
    if not cont_file:
        st.error("❌ Debes cargar Contabilidad")
        return
    
    inicio = datetime.now()
    
    try:
        resultados = {}
        
        # CONTABILIDAD
        st.header("📊 1. Procesando Contabilidad")
        cont_raw = read_table(cont_file, preferred_sheet='ContabilidadSET_PLUS_datos')
        cont_completa = prep_contabilidad(cont_raw, config, tipo_mov=None)
        cont_h = prep_contabilidad(cont_raw, config, tipo_mov='H')
        cont_d = prep_contabilidad(cont_raw, config, tipo_mov='D')
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Movimientos D", f"{len(cont_d):,}")
        col2.metric("Movimientos H", f"{len(cont_h):,}")
        col3.metric("Total", f"{len(cont_completa):,}")
        
        resultados['Contabilidad_D'] = cont_d
        resultados['Contabilidad_H'] = cont_h
        
        # INGRESOS
        if liq_file:
            st.header("📈 2. INGRESOS (Liquidaciones)")
            
            liq_usecols = ["Liquidacion", "Numero_Viaje", "TipoPago", "Monto", "Unidad", "Owner", "Tipo_Concepto"]
            liq_raw = read_table(liq_file, preferred_sheet='LiquidacionesSET_PLUS_datos', usecols=liq_usecols)
            liq = prep_liquidaciones(liq_raw, config, tipo_concepto)
            
            liq_clas, cont_h_clas = match_liquidaciones_vs_cont(liq, cont_h)
            
            col1, col2, col3 = st.columns(3)
            col1.metric("✅ MATCH_OK", f"{(liq_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum():,}")
            col2.metric("⚠️ DISCREPANCIA", f"{(liq_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum():,}")
            col3.metric("❌ NO_MATCH", f"{(liq_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD').sum():,}")
            
            resultados['Liquidaciones_Clasificadas'] = liq_clas
            resultados['Contabilidad_H_Clasificada'] = cont_h_clas
        
        # COSTOS - BASE
        if base_file:
            st.header("📋 3. COSTOS - Base Saldos")
            
            base_raw = read_table(base_file)
            base = prep_base_saldos(base_raw, config)
            
            base_clas, cont_d_clas_base, scoring_base, best_base = match_base_vs_cont(base, cont_d)
            
            col1, col2, col3 = st.columns(3)
            col1.metric("✅ MATCH_OK", f"{(base_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum():,}")
            col2.metric("⚠️ DISCREPANCIA", f"{(base_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum():,}")
            col3.metric("❌ NO_MATCH", f"{(base_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D').sum():,}")
            
            resultados['Base_Clasificada'] = base_clas
            resultados['Contabilidad_D_vs_Base'] = cont_d_clas_base
            if not scoring_base.empty:
                resultados['Scoring_Base'] = scoring_base
            if not best_base.empty:
                resultados['Best_Matches_Base'] = best_base
        
        # COSTOS - CHEQUES/VOUCHERS
        if cheques_file and vouchers_file:
            st.header("💵 4. COSTOS - Cheques/Vouchers")
            
            cheques_raw = read_table(cheques_file)
            vouchers_raw = read_table(vouchers_file)
            
            cheques, cheques_excl = prep_cheques(cheques_raw, config)
            vouchers, vouchers_excl = prep_vouchers(vouchers_raw, config)
            
            st.info(f"Cheques válidos: {len(cheques):,} | Excluidos (Company): {len(cheques_excl):,}")
            st.info(f"Vouchers válidos: {len(vouchers):,} | Excluidos (Filial): {len(vouchers_excl):,}")
            
            costos_depurados, duplicados = dedupe_costos(cheques, vouchers)
            costos_clas, cont_d_clas_costos, scoring_costos, best_costos = match_costos_vs_cont(costos_depurados, cont_d)
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("🧹 Depurados", f"{len(costos_depurados):,}")
            col2.metric("✅ MATCH_OK", f"{(costos_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum():,}")
            col3.metric("⚠️ DISCREPANCIA", f"{(costos_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum():,}")
            col4.metric("❌ NO_MATCH", f"{(costos_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D').sum():,}")
            
            resultados['Cheques_Excluidos_Company'] = cheques_excl
            resultados['Vouchers_Excluidos_Filial'] = vouchers_excl
            resultados['Costos_Depurados'] = costos_depurados
            resultados['Duplicados'] = duplicados
            resultados['Costos_Clasificados'] = costos_clas
            resultados['Contabilidad_D_vs_Costos'] = cont_d_clas_costos
            if not scoring_costos.empty:
                resultados['Scoring_Costos'] = scoring_costos
            if not best_costos.empty:
                resultados['Best_Matches_Costos'] = best_costos
        
        # ANÁLISIS D vs H
        st.header("⚖️ 5. Análisis D vs H")
        saldos_dh = analizar_DH(cont_completa)
        
        if not saldos_dh.empty:
            col1, col2, col3 = st.columns(3)
            col1.metric("✅ Saldados", f"{(saldos_dh['ESTATUS_DH'] == 'SALDADO').sum():,}")
            col2.metric("⏳ Pendientes", f"{(saldos_dh['ESTATUS_DH'] == 'PENDIENTE_PAGO').sum():,}")
            col3.metric("⚠️ Sobrepagos", f"{(saldos_dh['ESTATUS_DH'] == 'SOBREPAGO').sum():,}")
            
            resultados['Analisis_DH'] = saldos_dh
        
        # RESUMEN GENERAL
        resumen_general = pd.DataFrame([{
            'MODULO': 'INGRESOS',
            'REGISTROS_PROCESADOS': len(liq_clas) if liq_file else 0,
            'MATCH_OK': (liq_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum() if liq_file else 0,
            'MATCH_CON_DISCREPANCIA': (liq_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum() if liq_file else 0,
            'NO_MATCH': (liq_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD').sum() if liq_file else 0,
        }, {
            'MODULO': 'COSTOS_BASE',
            'REGISTROS_PROCESADOS': len(base_clas) if base_file else 0,
            'MATCH_OK': (base_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum() if base_file else 0,
            'MATCH_CON_DISCREPANCIA': (base_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum() if base_file else 0,
            'NO_MATCH': (base_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D').sum() if base_file else 0,
        }, {
            'MODULO': 'COSTOS_OPERATIVO',
            'REGISTROS_PROCESADOS': len(costos_clas) if cheques_file and vouchers_file else 0,
            'MATCH_OK': (costos_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum() if cheques_file and vouchers_file else 0,
            'MATCH_CON_DISCREPANCIA': (costos_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum() if cheques_file and vouchers_file else 0,
            'NO_MATCH': (costos_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D').sum() if cheques_file and vouchers_file else 0,
        }])
        resultados['Resumen_General'] = resumen_general
        
        # EXPORTAR
        st.divider()
        tiempo = (datetime.now() - inicio).total_seconds()
        st.success(f"✅ Completado en {tiempo:.1f}s ({tiempo/60:.1f} min)")
        
        if resultados:
            excel = exportar_excel(resultados)
            st.download_button(
                "📥 Descargar Excel Completo",
                excel,
                f"conciliacion_final_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            with st.expander("📋 Hojas incluidas", expanded=True):
                for nombre, df in resultados.items():
                    st.write(f"✅ **{nombre}**: {len(df):,} filas")
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.exception(e)

if __name__ == "__main__":
    main()
