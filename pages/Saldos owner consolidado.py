"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                    SALDOS OWNER - SISTEMA CONSOLIDADO                        ║
║                                                                              ║
║  Flujo unificado que previene duplicados y mantiene trazabilidad completa   ║
║                                                                              ║
║  ETAPA 1: INGRESOS (Liquidaciones vs Contabilidad H)                        ║
║  ETAPA 2: COSTOS   (Vales/Base vs Contabilidad D)                           ║
║  ETAPA 3: CROSSMATCH (CA/PD/H para NO_EXISTE)                               ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import re
import time
import unicodedata
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Saldos Owner Consolidado", layout="wide")

# ============================================================
# HELPERS GENERALES
# ============================================================

def norm_text(x: object) -> str:
    """Normalización de texto conservadora"""
    if x is None or pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s


def norm_for_key(x: object) -> str:
    """Normalización agresiva para keys (elimina caracteres especiales)"""
    s = norm_text(x)
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def norm_amount(x: object, ndigits: int = 2) -> float:
    """Normalización de importes"""
    try:
        if x is None or pd.isna(x):
            return float("nan")
        if isinstance(x, str):
            x = x.replace(",", "").replace("$", "").strip()
        return round(float(x), ndigits)
    except Exception:
        return float("nan")


def normalizar_viaje(serie):
    """Normaliza viajes para crossmatch (elimina / y -)"""
    return serie.fillna('').astype(str).str.replace('/', '', regex=False).str.replace('-', '', regex=False).str.strip().str.upper()


def strip_concept_suffix(x: object) -> str:
    """Quita sufijos tipo ' - 20170908' sin destruir el concepto base"""
    s = norm_text(x)
    s = re.sub(r"\s+-\s+\d+.*$", "", s)
    s = re.sub(r"\s+-\s+[A-Z0-9]+.*$", "", s)
    return s.strip()


def canonical_concept(x: object, concept_map: dict[str, str] | None = None) -> str:
    """Normaliza conceptos con reglas base"""
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


def read_table(file_obj, preferred_sheet: str | None = None, usecols=None) -> pd.DataFrame:
    """Lee CSV o Excel de forma robusta"""
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
    """Encuentra columna por múltiples nombres candidatos"""
    normalized = {norm_for_key(c): c for c in df.columns}
    for c in candidates:
        key = norm_for_key(c)
        if key in normalized:
            return normalized[key]
    if required:
        raise ValueError(f"No encontré columna: {candidates}. Disponibles: {list(df.columns)}")
    return None


def resolve_all_cols(df: pd.DataFrame, candidates: list[str]) -> list[str]:
    """Encuentra todas las columnas que coinciden (para duplicados como Importe.1)"""
    wanted = {norm_for_key(c) for c in candidates}
    out = []
    for col in df.columns:
        base = re.sub(r"\.\d+$", "", str(col))
        if norm_for_key(base) in wanted:
            out.append(col)
    return out


def choose_cont_import_col(cont_raw: pd.DataFrame) -> str:
    """Elige la columna de importe correcta (última Importe si hay duplicados)"""
    importe_cols = resolve_all_cols(cont_raw, ["Importe", "Monto", "Total"])
    if not importe_cols:
        raise ValueError("No encontré columna de importe en Contabilidad")
    return importe_cols[-1]


def build_seq(df: pd.DataFrame, key_cols: list[str], seq_col: str = "_seq") -> pd.DataFrame:
    """Agrega secuencia por grupo para manejar duplicados"""
    out = df.copy()
    out[seq_col] = out.groupby(key_cols, dropna=False).cumcount() + 1
    return out


# ============================================================
# PREPARACIÓN DE ARCHIVOS
# ============================================================

def prep_contabilidad(cont_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str], tipo_mov: str | None = None) -> pd.DataFrame:
    """Prepara archivo de contabilidad completo"""
    c_mov = resolve_col(cont_raw, ["TipoMovimiento", "Movimiento", "Tipo Movimiento"])
    c_importe = choose_cont_import_col(cont_raw)
    c_unidad = resolve_col(cont_raw, ["Unidad", "Numero de Unidad", "Numero_Unidad"])
    c_referencia = resolve_col(cont_raw, ["Referencia", "Numero_Viaje", "Numero Viaje", "Viaje"], required=False)
    c_poliza = resolve_col(cont_raw, ["Clave Poliza", "Clave Póliza", "ClavePoliza", "Factura", "Contrarrecibo"])
    c_concepto = resolve_col(cont_raw, ["Concepto detalle", "Concepto Detalle", "ConceptoDetalle", "Concepto", "NombreCuentaContable"], required=False)
    c_vale = resolve_col(cont_raw, ["Vale", "No Vale", "Numero Vale"], required=False)
    c_owner = resolve_col(cont_raw, ["NombreCuentaContable", "Cuenta Contable", "Owner"], required=False)

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
    out["OWNER_CONT"] = out[c_owner].apply(norm_text) if c_owner else ""
    out["ROW_ID_CONT"] = range(1, len(out) + 1)
    out["MATCHED_IN_ETAPA"] = None  # Marca qué etapa matcheó este registro
    
    # Guardar referencias originales
    out["_POLIZA_ORIG"] = out[c_poliza]
    out["_UNIDAD_ORIG"] = out[c_unidad]
    out["_VIAJE_ORIG"] = out[c_referencia] if c_referencia else ""
    out["_IMPORTE_ORIG"] = out[c_importe]
    out["_CONCEPTO_ORIG"] = out[c_concepto] if c_concepto else ""
    
    return out


def prep_liquidaciones(liq_raw: pd.DataFrame, ndigits: int) -> pd.DataFrame:
    """Prepara archivo de liquidaciones"""
    liq = liq_raw.rename(columns={
        "Liquidacion": "PR",
        "Numero_Viaje": "VIAJE",
        "TipoPago": "TIPO_PAGO",
        "Monto": "IMPORTE",
        "Unidad": "UNIDAD",
        "Owner": "OWNER_LIQ",
        "Tipo_Concepto": "TIPO_CONCEPTO",
    })
    
    # Excluir conceptos que no se reflejan en contabilidad
    if "Concepto" in liq.columns:
        liq["Concepto"] = liq["Concepto"].apply(norm_text)
        conceptos_excluir = ["ADICIONAL CHARGES"]
        liq = liq[~liq["Concepto"].isin(conceptos_excluir)].copy()
    
    for c in ["PR", "VIAJE", "TIPO_PAGO", "UNIDAD", "OWNER_LIQ", "TIPO_CONCEPTO"]:
        if c in liq.columns:
            liq[c] = liq[c].apply(norm_text)
    
    liq["IMPORTE"] = pd.to_numeric(liq["IMPORTE"].apply(lambda x: norm_amount(x, ndigits)), errors="coerce")
    liq["ROW_ID_LIQ"] = range(1, len(liq) + 1)
    liq["MATCHED_IN_ETAPA"] = None
    
    return liq


def prep_base_saldos(base_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str]) -> pd.DataFrame:
    """Prepara Base Saldos (para costos)"""
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
    out["MATCHED_IN_ETAPA"] = None
    
    return out


def prep_vales(vales_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str]) -> pd.DataFrame:
    """Prepara archivo de Vales (para costos)"""
    c_vale = resolve_col(vales_raw, ["Vale", "No Vale", "Numero Vale"])
    c_unidad = resolve_col(vales_raw, ["Unidad", "Numero de Unidad", "Numero_Unidad"])
    c_concepto = resolve_col(vales_raw, ["Concepto", "Concepto detalle"])
    c_contrarrecibo = resolve_col(vales_raw, ["Contrarecibo", "Contrarrecibo", "Clave Poliza"], required=False)
    c_importe = resolve_col(vales_raw, ["Total", "Importe", "TotalVale"])

    out = vales_raw.copy()
    out["ROW_ID_VALE"] = range(1, len(out) + 1)
    out["VALE_KEY"] = out[c_vale].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["CONCEPTO_KEY"] = out[c_concepto].apply(lambda x: canonical_concept(x, concept_map))
    out["POLIZA_KEY"] = out[c_contrarrecibo].apply(norm_for_key) if c_contrarrecibo else ""
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, ndigits))
    out["VIAJE_KEY"] = ""
    out["MATCHED_IN_ETAPA"] = None
    
    return out


# ============================================================
# ETAPA 1: INGRESOS (Liquidaciones vs Contabilidad H)
# ============================================================

def etapa_1_ingresos(liq: pd.DataFrame, cont_h: pd.DataFrame, ndigits: int) -> tuple:
    """
    Procesa Liquidaciones vs Contabilidad H con matching exacto
    Retorna: liq_clasificado, cont_clasificado, resumen
    """
    st.subheader("🔵 ETAPA 1: INGRESOS")
    inicio = time.time()
    
    # Filtrar solo Tipo_Concepto = E/I
    liq_tipo = st.session_state.get('liq_tipo', 'E')
    liq_f = liq[liq["TIPO_CONCEPTO"] == liq_tipo].copy()
    liq_f["ROW_ID_LIQ_F"] = range(1, len(liq_f) + 1)
    
    st.info(f"📊 Liquidaciones filtradas (Tipo={liq_tipo}): **{len(liq_f):,}** | Contabilidad H: **{len(cont_h):,}**")
    
    # Match exacto por PR + VIAJE + UNIDAD + TIPO_PAGO + IMPORTE
    key_cols = ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE"]
    
    # Preparar columnas para merge
    liq_k = liq_f[["ROW_ID_LIQ"] + key_cols + ["OWNER_LIQ"]].copy()
    cont_k = cont_h[["ROW_ID_CONT"] + key_cols + ["OWNER_CONT"]].copy()
    
    # Agregar secuencia para duplicados
    liq_k = build_seq(liq_k, key_cols)
    cont_k = build_seq(cont_k, key_cols)
    
    merge_keys = key_cols + ["_seq"]
    
    with st.spinner("Ejecutando matching de Ingresos..."):
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
        
        # Clasificar matches
        matched["OWNER_MATCH"] = (
            matched["OWNER_LIQ"].fillna("") == matched["OWNER_CONT"].fillna("")
        )
        matched["ESTATUS_MATCH"] = matched["OWNER_MATCH"].map({
            True: "MATCH_OK", 
            False: "MATCH_CON_DISCREPANCIA"
        })
        
        only_liq["ESTATUS_MATCH"] = "NO_EXISTE_EN_CONTABILIDAD"
        only_cont["ESTATUS_MATCH"] = "NO_EXISTE_EN_LIQUIDACIONES"
        
        # Marcar registros matcheados
        matched_liq_ids = matched["ROW_ID_LIQ"].dropna().unique()
        matched_cont_ids = matched["ROW_ID_CONT"].dropna().unique()
    
    # Construir clasificados
    liq_status = pd.concat([
        matched[["ROW_ID_LIQ", "ROW_ID_CONT", "ESTATUS_MATCH", "OWNER_CONT"]],
        only_liq[["ROW_ID_LIQ", "ESTATUS_MATCH"]].assign(ROW_ID_CONT=pd.NA, OWNER_CONT="")
    ], ignore_index=True)
    
    cont_status = pd.concat([
        matched[["ROW_ID_CONT", "ROW_ID_LIQ", "ESTATUS_MATCH", "OWNER_LIQ"]],
        only_cont[["ROW_ID_CONT", "ESTATUS_MATCH"]].assign(ROW_ID_LIQ=pd.NA, OWNER_LIQ="")
    ], ignore_index=True)
    
    liq_clasificado = liq_f.merge(liq_status, on="ROW_ID_LIQ", how="left")
    cont_clasificado = cont_h.merge(cont_status, on="ROW_ID_CONT", how="left")
    
    # Marcar en dataframes originales
    liq.loc[liq["ROW_ID_LIQ"].isin(matched_liq_ids), "MATCHED_IN_ETAPA"] = "ETAPA_1_INGRESOS"
    cont_h.loc[cont_h["ROW_ID_CONT"].isin(matched_cont_ids), "MATCHED_IN_ETAPA"] = "ETAPA_1_INGRESOS"
    
    # Resumen
    resumen = {
        "MATCH_OK": int((liq_clasificado["ESTATUS_MATCH"] == "MATCH_OK").sum()),
        "MATCH_CON_DISCREPANCIA": int((liq_clasificado["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA").sum()),
        "NO_EXISTE": int((liq_clasificado["ESTATUS_MATCH"] == "NO_EXISTE_EN_CONTABILIDAD").sum()),
        "TOTAL": len(liq_clasificado)
    }
    
    tiempo = time.time() - inicio
    st.success(f"✅ Etapa 1 completada en {tiempo:.1f}s | Matches: {resumen['MATCH_OK']:,} OK + {resumen['MATCH_CON_DISCREPANCIA']:,} Discrepancia")
    
    return liq_clasificado, cont_clasificado, resumen


# ============================================================
# ETAPA 2: COSTOS (Base/Vales vs Contabilidad D)
# ============================================================

def score_pairs_base(base: pd.DataFrame, cont: pd.DataFrame, pairs: pd.DataFrame) -> pd.DataFrame:
    """Score de pares Base vs Contabilidad con 5 criterios"""
    if pairs.empty:
        return pairs
    
    b = base[["ROW_ID_BASE", "POLIZA_KEY", "UNIDAD_KEY", "VIAJE_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"]]
    c = cont[["ROW_ID_CONT", "POLIZA_KEY", "UNIDAD_KEY", "VIAJE_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"]]
    
    x = pairs.merge(b, on="ROW_ID_BASE", how="left").merge(
        c, on="ROW_ID_CONT", how="left", suffixes=("_BASE", "_CONT")
    )
    
    x["COINCIDE_POLIZA"] = x["POLIZA_KEY_BASE"] == x["POLIZA_KEY_CONT"]
    x["COINCIDE_UNIDAD"] = x["UNIDAD_KEY_BASE"] == x["UNIDAD_KEY_CONT"]
    x["COINCIDE_VIAJE"] = x["VIAJE_KEY_BASE"] == x["VIAJE_KEY_CONT"]
    x["COINCIDE_CONCEPTO"] = x["CONCEPTO_KEY_BASE"] == x["CONCEPTO_KEY_CONT"]
    x["COINCIDE_IMPORTE"] = x["IMPORTE_KEY_BASE"] == x["IMPORTE_KEY_CONT"]
    
    x["TOTAL_COINCIDENCIAS"] = (
        x["COINCIDE_POLIZA"].astype(int) +
        x["COINCIDE_UNIDAD"].astype(int) +
        x["COINCIDE_VIAJE"].astype(int) +
        x["COINCIDE_CONCEPTO"].astype(int) +
        x["COINCIDE_IMPORTE"].astype(int)
    )
    
    def estatus(row):
        if row["TOTAL_COINCIDENCIAS"] == 5:
            return "MATCH_OK"
        if row["TOTAL_COINCIDENCIAS"] >= 3:
            return "MATCH_CON_DISCREPANCIA"
        return "CANDIDATO_DEBIL"
    
    x["ESTATUS_MATCH"] = x.apply(estatus, axis=1)
    return x


def greedy_best_match(scored: pd.DataFrame, left_id: str, right_id: str, used_right: set) -> pd.DataFrame:
    """
    Matching greedy que EXCLUYE registros ya matcheados en etapas anteriores
    """
    if scored.empty:
        return scored
    
    # Filtrar candidatos >= 3 coincidencias Y que no estén ya matcheados
    candidates = scored[
        (scored["TOTAL_COINCIDENCIAS"] >= 3) &
        (~scored[right_id].isin(used_right))
    ].copy()
    
    if candidates.empty:
        return candidates
    
    # Ordenar por calidad de match
    candidates = candidates.sort_values(
        ["TOTAL_COINCIDENCIAS", left_id, right_id],
        ascending=[False, True, True]
    )
    
    used_left = set()
    new_used_right = set()
    rows = []
    
    for _, row in candidates.iterrows():
        lid = row[left_id]
        rid = row[right_id]
        
        if lid in used_left or rid in new_used_right:
            continue
        
        used_left.add(lid)
        new_used_right.add(rid)
        rows.append(row)
    
    if not rows:
        return candidates.iloc[0:0].copy()
    
    return pd.DataFrame(rows)


def etapa_2_costos(base: pd.DataFrame, vales: pd.DataFrame, cont_d: pd.DataFrame, 
                   used_cont_ids: set, proceso: str) -> tuple:
    """
    Procesa Base Saldos y/o Vales vs Contabilidad D
    Excluye registros ya matcheados en Etapa 1
    """
    st.subheader("🟢 ETAPA 2: COSTOS")
    inicio = time.time()
    
    resultado_sheets = {}
    resumen_final = {}
    
    # IMPORTANTE: Filtrar Contabilidad D para excluir ya matcheados
    cont_d_disponible = cont_d[~cont_d["ROW_ID_CONT"].isin(used_cont_ids)].copy()
    
    st.info(f"📊 Contabilidad D disponible (sin matches Etapa 1): **{len(cont_d_disponible):,}**")
    
    # ========== BASE SALDOS ==========
    if proceso in {"Base Saldos vs Contabilidad D", "Ambos"} and base is not None:
        with st.spinner("Procesando Base Saldos vs Contabilidad D..."):
            # Generar pares candidatos
            blocks = [
                ["POLIZA_KEY", "IMPORTE_KEY"],
                ["POLIZA_KEY", "UNIDAD_KEY"],
                ["UNIDAD_KEY", "VIAJE_KEY", "IMPORTE_KEY"],
                ["POLIZA_KEY", "VIAJE_KEY"],
                ["UNIDAD_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"],
            ]
            
            pairs_list = []
            for block in blocks:
                b = base[["ROW_ID_BASE"] + block].dropna(subset=block)
                c = cont_d_disponible[["ROW_ID_CONT"] + block].dropna(subset=block)
                if not b.empty and not c.empty:
                    p = b.merge(c, on=block, how="inner")[["ROW_ID_BASE", "ROW_ID_CONT"]]
                    pairs_list.append(p)
            
            pairs = pd.concat(pairs_list, ignore_index=True).drop_duplicates() if pairs_list else pd.DataFrame(columns=["ROW_ID_BASE", "ROW_ID_CONT"])
            
            # Scorear
            scored = score_pairs_base(base, cont_d_disponible, pairs)
            
            # Match greedy EXCLUYENDO ya matcheados
            best = greedy_best_match(scored, "ROW_ID_BASE", "ROW_ID_CONT", used_cont_ids)
            
            # Clasificar
            base_status = best[["ROW_ID_BASE", "ROW_ID_CONT", "ESTATUS_MATCH", "TOTAL_COINCIDENCIAS"]].copy() if not best.empty else pd.DataFrame()
            base_clas = base.merge(base_status, on="ROW_ID_BASE", how="left")
            base_clas["ESTATUS_MATCH"] = base_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_CONTABILIDAD_D")
            base_clas["TOTAL_COINCIDENCIAS"] = base_clas["TOTAL_COINCIDENCIAS"].fillna(0).astype(int)
            
            # Marcar matcheados
            if not best.empty:
                matched_base_ids = best["ROW_ID_BASE"].dropna().unique()
                matched_cont_ids = best["ROW_ID_CONT"].dropna().unique()
                base.loc[base["ROW_ID_BASE"].isin(matched_base_ids), "MATCHED_IN_ETAPA"] = "ETAPA_2_COSTOS_BASE"
                cont_d.loc[cont_d["ROW_ID_CONT"].isin(matched_cont_ids), "MATCHED_IN_ETAPA"] = "ETAPA_2_COSTOS_BASE"
                used_cont_ids.update(matched_cont_ids)
            
            resultado_sheets["Base_Clasificada"] = base_clas
            
            resumen_final["Base"] = {
                "MATCH_OK": int((base_clas["ESTATUS_MATCH"] == "MATCH_OK").sum()),
                "MATCH_CON_DISCREPANCIA": int((base_clas["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA").sum()),
                "NO_EXISTE": int((base_clas["ESTATUS_MATCH"] == "NO_EXISTE_EN_CONTABILIDAD_D").sum()),
                "TOTAL": len(base_clas)
            }
            
            st.write(f"✅ Base Saldos: {resumen_final['Base']['MATCH_OK']:,} OK + {resumen_final['Base']['MATCH_CON_DISCREPANCIA']:,} Discrepancia")
    
    # ========== VALES ==========
    if proceso in {"Vales vs Contabilidad D", "Ambos"} and vales is not None:
        # Actualizar disponibles después de Base
        cont_d_disponible = cont_d[~cont_d["ROW_ID_CONT"].isin(used_cont_ids)].copy()
        
        with st.spinner("Procesando Vales vs Contabilidad D..."):
            # Similar lógica pero para Vales
            blocks = [
                ["UNIDAD_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"],
                ["UNIDAD_KEY", "IMPORTE_KEY"],
                ["CONCEPTO_KEY", "IMPORTE_KEY"],
                ["VALE_KEY", "IMPORTE_KEY"],
                ["POLIZA_KEY", "IMPORTE_KEY"],
            ]
            
            pairs_list = []
            for block in blocks:
                v = vales[["ROW_ID_VALE"] + block].dropna(subset=block)
                c = cont_d_disponible[["ROW_ID_CONT"] + block].dropna(subset=block)
                if not v.empty and not c.empty:
                    p = v.merge(c, on=block, how="inner")[["ROW_ID_VALE", "ROW_ID_CONT"]]
                    pairs_list.append(p)
            
            pairs = pd.concat(pairs_list, ignore_index=True).drop_duplicates() if pairs_list else pd.DataFrame(columns=["ROW_ID_VALE", "ROW_ID_CONT"])
            
            # Score simplificado para Vales (4 criterios principales)
            if not pairs.empty:
                v = vales[["ROW_ID_VALE", "VALE_KEY", "UNIDAD_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"]]
                c = cont_d_disponible[["ROW_ID_CONT", "VALE_KEY", "UNIDAD_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"]]
                
                scored = pairs.merge(v, on="ROW_ID_VALE", how="left").merge(
                    c, on="ROW_ID_CONT", how="left", suffixes=("_VALE", "_CONT")
                )
                
                scored["COINCIDE_VALE"] = scored["VALE_KEY_VALE"] == scored["VALE_KEY_CONT"]
                scored["COINCIDE_UNIDAD"] = scored["UNIDAD_KEY_VALE"] == scored["UNIDAD_KEY_CONT"]
                scored["COINCIDE_CONCEPTO"] = scored["CONCEPTO_KEY_VALE"] == scored["CONCEPTO_KEY_CONT"]
                scored["COINCIDE_IMPORTE"] = scored["IMPORTE_KEY_VALE"] == scored["IMPORTE_KEY_CONT"]
                
                scored["TOTAL_COINCIDENCIAS"] = (
                    scored["COINCIDE_VALE"].astype(int) +
                    scored["COINCIDE_UNIDAD"].astype(int) +
                    scored["COINCIDE_CONCEPTO"].astype(int) +
                    scored["COINCIDE_IMPORTE"].astype(int)
                )
                
                scored["ESTATUS_MATCH"] = scored["TOTAL_COINCIDENCIAS"].apply(
                    lambda x: "MATCH_OK" if x == 4 else ("MATCH_CON_DISCREPANCIA" if x >= 2 else "CANDIDATO_DEBIL")
                )
            else:
                scored = pairs
            
            best = greedy_best_match(scored, "ROW_ID_VALE", "ROW_ID_CONT", used_cont_ids)
            
            vales_status = best[["ROW_ID_VALE", "ROW_ID_CONT", "ESTATUS_MATCH", "TOTAL_COINCIDENCIAS"]].copy() if not best.empty else pd.DataFrame()
            vales_clas = vales.merge(vales_status, on="ROW_ID_VALE", how="left")
            vales_clas["ESTATUS_MATCH"] = vales_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_CONTABILIDAD_D")
            vales_clas["TOTAL_COINCIDENCIAS"] = vales_clas["TOTAL_COINCIDENCIAS"].fillna(0).astype(int)
            
            if not best.empty:
                matched_vale_ids = best["ROW_ID_VALE"].dropna().unique()
                matched_cont_ids = best["ROW_ID_CONT"].dropna().unique()
                vales.loc[vales["ROW_ID_VALE"].isin(matched_vale_ids), "MATCHED_IN_ETAPA"] = "ETAPA_2_COSTOS_VALES"
                cont_d.loc[cont_d["ROW_ID_CONT"].isin(matched_cont_ids), "MATCHED_IN_ETAPA"] = "ETAPA_2_COSTOS_VALES"
                used_cont_ids.update(matched_cont_ids)
            
            resultado_sheets["Vales_Clasificados"] = vales_clas
            
            resumen_final["Vales"] = {
                "MATCH_OK": int((vales_clas["ESTATUS_MATCH"] == "MATCH_OK").sum()),
                "MATCH_CON_DISCREPANCIA": int((vales_clas["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA").sum()),
                "NO_EXISTE": int((vales_clas["ESTATUS_MATCH"] == "NO_EXISTE_EN_CONTABILIDAD_D").sum()),
                "TOTAL": len(vales_clas)
            }
            
            st.write(f"✅ Vales: {resumen_final['Vales']['MATCH_OK']:,} OK + {resumen_final['Vales']['MATCH_CON_DISCREPANCIA']:,} Discrepancia")
    
    tiempo = time.time() - inicio
    st.success(f"✅ Etapa 2 completada en {tiempo:.1f}s")
    
    return resultado_sheets, resumen_final, used_cont_ids


# ============================================================
# ETAPA 3: CROSSMATCH (CA/PD/H para NO_EXISTE)
# ============================================================

def etapa_3_crossmatch(base_no_existe: pd.DataFrame, cont_all: pd.DataFrame, 
                       used_cont_ids: set) -> tuple:
    """
    Busca matches en pólizas CA/PD/H para registros NO_EXISTE
    Solo procesa registros NO matcheados en etapas anteriores
    """
    st.subheader("🟣 ETAPA 3: CROSSMATCH")
    inicio = time.time()
    
    if base_no_existe.empty:
        st.warning("No hay registros NO_EXISTE para procesar en Crossmatch")
        return pd.DataFrame(), {}
    
    st.info(f"📊 Registros NO_EXISTE a analizar: **{len(base_no_existe):,}**")
    
    # Filtrar contabilidad excluyendo ya matcheados
    cont_disponible = cont_all[~cont_all["ROW_ID_CONT"].isin(used_cont_ids)].copy()
    
    with st.spinner("Ejecutando análisis crossmatch..."):
        # Preparar datos
        base = base_no_existe.copy()
        base['idx_original'] = range(len(base))
        base['poliza_norm'] = base.get('POLIZA_KEY', base.get('FOLIO_CONTRARECIBO', pd.Series())).fillna('').astype(str).str.strip().str.upper()
        base['viaje_norm'] = normalizar_viaje(base.get('VIAJE_KEY', base.get('NUMERO_VIAJE', pd.Series())))
        base['importe'] = pd.to_numeric(base.get('IMPORTE_KEY', base.get('Importe', pd.Series())), errors='coerce').fillna(0).round(2)
        base['concepto_norm'] = base.get('CONCEPTO_KEY', '').fillna('').astype(str).str.upper()
        base['es_diesel'] = base['concepto_norm'].str.contains('DIESEL|CONSUMIBLES', na=False)
        
        # Preparar contabilidad
        cont = cont_disponible.copy()
        cont['poliza_norm'] = cont['POLIZA_KEY'].fillna('').astype(str).str.strip().str.upper()
        cont['viaje_norm'] = normalizar_viaje(cont['VIAJE_KEY'])
        cont['importe'] = pd.to_numeric(cont['IMPORTE_KEY'], errors='coerce').fillna(0).round(2)
        cont['concepto_norm'] = cont.get('CONCEPTO_KEY', '').fillna('').astype(str).str.upper()
        cont['tipo_poliza'] = cont['POLIZA_KEY'].fillna('').astype(str).str[:2]
        
        # Separar por tipo
        cont_d = cont[cont['TIPO_MOV'] == 'D'].copy()
        cont_h = cont[cont['TIPO_MOV'] == 'H'].copy()
        cont_d_ca = cont_d[cont_d['tipo_poliza'] == 'CA'].copy()
        cont_d_pd = cont_d[cont_d['tipo_poliza'] == 'PD'].copy()
        cont_h_no_ca = cont_h[~cont_h['tipo_poliza'].isin(['CA'])].copy()
        
        # ===== BUSCAR CARGOS CA =====
        base_ca = base[['idx_original', 'poliza_norm', 'importe']].copy()
        cont_ca = cont_d_ca[['poliza_norm', 'importe', '_UNIDAD_ORIG', '_VIAJE_ORIG', 'ROW_ID_CONT']].copy()
        
        matches_ca = base_ca.merge(
            cont_ca,
            on=['poliza_norm', 'importe'],
            how='left',
            suffixes=('', '_ca')
        )
        matches_ca = matches_ca.groupby('idx_original').first().reset_index()
        
        base['tiene_cargo_ca'] = base['idx_original'].isin(matches_ca[matches_ca['_UNIDAD_ORIG'].notna()]['idx_original'])
        base = base.merge(
            matches_ca[['idx_original', '_UNIDAD_ORIG', '_VIAJE_ORIG', 'ROW_ID_CONT']].rename(columns={
                '_UNIDAD_ORIG': 'ca_unidad',
                '_VIAJE_ORIG': 'ca_viaje',
                'ROW_ID_CONT': 'ca_row_id'
            }),
            on='idx_original',
            how='left'
        )
        
        # ===== BUSCAR CARGOS PD =====
        base_pd = base[['idx_original', 'viaje_norm', 'importe', 'es_diesel', 'concepto_norm']].copy()
        cont_pd = cont_d_pd[['viaje_norm', 'importe', '_UNIDAD_ORIG', '_VIAJE_ORIG', '_POLIZA_ORIG', 'concepto_norm', 'ROW_ID_CONT']].copy()
        cont_pd['es_diesel_pd'] = cont_pd['concepto_norm'].str.contains('DIESEL', na=False)
        
        # Merge exacto diesel
        matches_pd_exacto = base_pd.merge(
            cont_pd[cont_pd['es_diesel_pd']],
            on=['viaje_norm', 'importe'],
            how='inner',
            suffixes=('', '_pd')
        )
        matches_pd_exacto = matches_pd_exacto[matches_pd_exacto['es_diesel']].groupby('idx_original').first().reset_index()
        
        base['tiene_pd_exacto'] = base['idx_original'].isin(matches_pd_exacto['idx_original'])
        base = base.merge(
            matches_pd_exacto[['idx_original', '_UNIDAD_ORIG', '_VIAJE_ORIG', '_POLIZA_ORIG', 'ROW_ID_CONT']].rename(columns={
                '_UNIDAD_ORIG': 'pd_unidad',
                '_VIAJE_ORIG': 'pd_viaje',
                '_POLIZA_ORIG': 'pd_poliza',
                'ROW_ID_CONT': 'pd_row_id'
            }),
            on='idx_original',
            how='left'
        )
        
        # ===== BUSCAR ABONOS H =====
        base_h = base[['idx_original', 'viaje_norm', 'importe']].copy()
        cont_h_data = cont_h_no_ca[['viaje_norm', 'importe', '_UNIDAD_ORIG', '_VIAJE_ORIG', '_POLIZA_ORIG', 'OWNER_CONT', 'ROW_ID_CONT']].copy()
        
        matches_h = base_h.merge(
            cont_h_data,
            on=['viaje_norm', 'importe'],
            how='left',
            suffixes=('', '_h')
        )
        matches_h = matches_h.groupby('idx_original').first().reset_index()
        
        base['tiene_abono_h'] = base['idx_original'].isin(matches_h[matches_h['_UNIDAD_ORIG'].notna()]['idx_original'])
        base = base.merge(
            matches_h[['idx_original', '_UNIDAD_ORIG', '_VIAJE_ORIG', '_POLIZA_ORIG', 'OWNER_CONT', 'ROW_ID_CONT']].rename(columns={
                '_UNIDAD_ORIG': 'h_unidad',
                '_VIAJE_ORIG': 'h_viaje',
                '_POLIZA_ORIG': 'h_poliza',
                'OWNER_CONT': 'h_owner',
                'ROW_ID_CONT': 'h_row_id'
            }),
            on='idx_original',
            how='left'
        )
        
        # ===== CLASIFICAR =====
        def clasificar(row):
            if row['tiene_cargo_ca'] and row['tiene_abono_h']:
                return 'COMPLETO_CA_H'
            elif row['tiene_pd_exacto'] and row['tiene_abono_h']:
                return 'COMPLETO_PD_H'
            elif row['tiene_cargo_ca']:
                return 'SOLO_CARGO_CA'
            elif row['tiene_pd_exacto']:
                return 'SOLO_CARGO_PD'
            elif row['tiene_abono_h']:
                return 'SOLO_ABONO_H'
            else:
                return 'NO_ENCONTRADO_CROSSMATCH'
        
        base['TIPO_CASO_CROSSMATCH'] = base.apply(clasificar, axis=1)
        
        # Diagnosticos
        diagnosticos = []
        for _, row in base.iterrows():
            if row['TIPO_CASO_CROSSMATCH'] == 'COMPLETO_CA_H':
                diag = f"✅ CA {row['ca_unidad']}|{row['ca_viaje']} + H {row['h_poliza']}"
            elif row['TIPO_CASO_CROSSMATCH'] == 'COMPLETO_PD_H':
                diag = f"✅ PD {row['pd_poliza']} + H {row['h_poliza']}"
            elif row['TIPO_CASO_CROSSMATCH'] == 'SOLO_CARGO_CA':
                diag = f"⚠️ Solo CA: {row['ca_unidad']}|{row['ca_viaje']}"
            elif row['TIPO_CASO_CROSSMATCH'] == 'SOLO_CARGO_PD':
                diag = f"⚠️ Solo PD: {row['pd_poliza']}"
            elif row['TIPO_CASO_CROSSMATCH'] == 'SOLO_ABONO_H':
                diag = f"🔄 Solo H: {row['h_poliza']}"
            else:
                diag = "❌ No encontrado en crossmatch"
            
            diagnosticos.append(diag)
        
        base['DIAGNOSTICO_CROSSMATCH'] = diagnosticos
    
    # Marcar registros matcheados en crossmatch
    matched_cont_ids = set()
    for col in ['ca_row_id', 'pd_row_id', 'h_row_id']:
        if col in base.columns:
            matched_cont_ids.update(base[col].dropna().unique())
    
    if matched_cont_ids:
        cont_all.loc[cont_all["ROW_ID_CONT"].isin(matched_cont_ids), "MATCHED_IN_ETAPA"] = "ETAPA_3_CROSSMATCH"
    
    resumen = {
        "COMPLETO": int(base['TIPO_CASO_CROSSMATCH'].str.contains('COMPLETO', na=False).sum()),
        "PARCIAL": int(base['TIPO_CASO_CROSSMATCH'].str.contains('SOLO', na=False).sum()),
        "NO_ENCONTRADO": int((base['TIPO_CASO_CROSSMATCH'] == 'NO_ENCONTRADO_CROSSMATCH').sum()),
        "TOTAL": len(base)
    }
    
    tiempo = time.time() - inicio
    st.success(f"✅ Etapa 3 completada en {tiempo:.1f}s | Completos: {resumen['COMPLETO']:,} | Parciales: {resumen['PARCIAL']:,}")
    
    return base, resumen


# ============================================================
# GENERACIÓN DE EXCEL CONSOLIDADO
# ============================================================

def generar_excel_consolidado(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Genera archivo Excel con todas las hojas de resultados"""
    bio = BytesIO()
    
    with pd.ExcelWriter(bio, engine="xlsxwriter", engine_kwargs={"options": {"constant_memory": True}}) as writer:
        for name, df in sheets.items():
            if df is not None and not df.empty:
                # Limpiar nombres de columnas
                out = df.copy()
                out.columns = [str(c)[:250] for c in out.columns]
                
                # Eliminar columnas auxiliares que empiezan con _
                out = out[[c for c in out.columns if not str(c).startswith('_')]]
                
                out.to_excel(writer, sheet_name=name[:31], index=False)
    
    bio.seek(0)
    return bio.getvalue()


# ============================================================
# UI PRINCIPAL
# ============================================================

st.title("🎯 Saldos Owner - Sistema Consolidado")
st.caption("Integración completa: Ingresos → Costos → Crossmatch | Sin duplicados")

with st.expander("ℹ️ Cómo funciona el sistema consolidado", expanded=False):
    st.markdown("""
    ### Arquitectura de 3 Etapas
    
    **ETAPA 1: INGRESOS** (Liquidaciones vs Contabilidad H)
    - Match exacto: PR + VIAJE + UNIDAD + TIPO_PAGO + IMPORTE
    - Marca registros matcheados para excluirlos de siguientes etapas
    
    **ETAPA 2: COSTOS** (Base Saldos/Vales vs Contabilidad D)
    - Scoring por 5 criterios (Base) o 4 criterios (Vales)
    - Solo procesa Contabilidad D NO matcheada en Etapa 1
    - Greedy matching previene duplicados
    
    **ETAPA 3: CROSSMATCH** (Búsqueda en CA/PD/H)
    - Solo procesa registros NO_EXISTE de etapas anteriores
    - Busca en pólizas CA (cargos), PD (diesel/anticipos), H (abonos)
    - Clasifica: Completo (CA+H o PD+H), Parcial (solo uno), No encontrado
    
    ### Ventajas
    ✅ **Sin duplicados**: Un registro solo hace match una vez
    ✅ **Trazabilidad**: Columna MATCHED_IN_ETAPA indica dónde se procesó
    ✅ **Eficiencia**: Cache solo en carga inicial, no en resultados
    ✅ **Preview directo**: Descarga desde las tablas mostradas
    """)

# ============================================================
# SIDEBAR - CARGA DE ARCHIVOS
# ============================================================

with st.sidebar:
    st.header("📁 Archivos Requeridos")
    
    st.subheader("Contabilidad (ambas etapas)")
    cont_file = st.file_uploader("ContabilidadSET_PLUS_datos.xlsx", type=["xlsx", "xls", "xlsm", "csv"], key="cont")
    
    st.divider()
    
    st.subheader("Etapa 1: Ingresos")
    liq_file = st.file_uploader("Liquidaciones.xlsx", type=["xlsx", "xls", "xlsm", "csv"], key="liq")
    liq_tipo = st.selectbox("Tipo_Concepto Liquidaciones", options=["E", "I"], index=0)
    st.session_state['liq_tipo'] = liq_tipo
    
    st.divider()
    
    st.subheader("Etapa 2: Costos")
    base_file = st.file_uploader("Base Saldos (opcional)", type=["xlsx", "xls", "xlsm", "csv"], key="base")
    vales_file = st.file_uploader("Vales (opcional)", type=["xlsx", "xls", "xlsm", "csv"], key="vales")
    
    proceso_costos = st.radio(
        "Procesar en Etapa 2",
        ["Base Saldos vs Contabilidad D", "Vales vs Contabilidad D", "Ambos"],
        index=2
    )
    
    st.divider()
    
    st.subheader("Configuración")
    ndigits = st.number_input("Redondeo de importe", min_value=0, max_value=4, value=2, step=1)
    
    concept_file = st.file_uploader("Catálogo conceptos (opcional)", type=["xlsx", "xls", "xlsm", "csv"], key="concepts")
    
    st.divider()
    
    ejecutar = st.button("🚀 EJECUTAR CONSOLIDADO", type="primary", use_container_width=True)

# ============================================================
# EJECUCIÓN
# ============================================================

if not ejecutar:
    st.info("👈 Configura los archivos en el sidebar y presiona **EJECUTAR CONSOLIDADO**")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Etapa 1", "Ingresos", "Liq vs Cont H")
    with col2:
        st.metric("Etapa 2", "Costos", "Base/Vales vs Cont D")
    with col3:
        st.metric("Etapa 3", "Crossmatch", "CA/PD/H")
    
    st.stop()

# Validar archivos mínimos
if not cont_file:
    st.error("❌ Falta archivo de Contabilidad (requerido)")
    st.stop()

if not liq_file:
    st.error("❌ Falta archivo de Liquidaciones para Etapa 1")
    st.stop()

# Cargar catálogo de conceptos (opcional)
concept_map = {}
if concept_file:
    try:
        df_concepts = read_table(concept_file)
        src = resolve_col(df_concepts, ["concepto_origen", "concepto", "source"], required=False)
        dst = resolve_col(df_concepts, ["concepto_canonico", "canonico", "target"], required=False)
        if src and dst:
            concept_map = {norm_for_key(a): norm_for_key(b) for a, b in zip(df_concepts[src], df_concepts[dst]) if norm_for_key(a) and norm_for_key(b)}
            st.sidebar.success(f"✅ Catálogo cargado: {len(concept_map)} conceptos")
    except Exception as e:
        st.sidebar.warning(f"⚠️ No pude cargar catálogo: {e}")

# ============================================================
# INICIO DEL PROCESO
# ============================================================

inicio_total = time.time()

try:
    # ===== CARGAR CONTABILIDAD COMPLETA =====
    with st.spinner("Cargando Contabilidad completa..."):
        cont_raw = read_table(cont_file, preferred_sheet="ContabilidadSET_PLUS_datos")
        cont_all = prep_contabilidad(cont_raw, ndigits, concept_map, tipo_mov=None)
        cont_h = cont_all[cont_all["TIPO_MOV"] == "H"].copy()
        cont_d = cont_all[cont_all["TIPO_MOV"] == "D"].copy()
        
        st.success(f"✅ Contabilidad cargada: {len(cont_all):,} registros | H: {len(cont_h):,} | D: {len(cont_d):,}")
    
    # ===== CARGAR LIQUIDACIONES =====
    with st.spinner("Cargando Liquidaciones..."):
        liq_usecols = ["Liquidacion", "Numero_Viaje", "TipoPago", "Monto", "Unidad", "Owner", "Tipo_Concepto"]
        liq_raw = read_table(liq_file, preferred_sheet="LiquidacionesSET_PLUS_datos", usecols=liq_usecols)
        
        # Renombrar columnas
        liq_raw = liq_raw.rename(columns={
            "Liquidacion": "PR",
            "Numero_Viaje": "VIAJE",
            "TipoPago": "TIPO_PAGO",
            "Monto": "IMPORTE",
            "Unidad": "UNIDAD",
            "Owner": "OWNER_LIQ",
            "Tipo_Concepto": "TIPO_CONCEPTO",
        })
        
        # Normalizar contabilidad para merge
        cont_h_for_merge = cont_h.copy()
        cont_h_for_merge["PR"] = cont_h_for_merge["_POLIZA_ORIG"].apply(norm_text)
        cont_h_for_merge["VIAJE"] = cont_h_for_merge["_VIAJE_ORIG"].apply(norm_text)
        cont_h_for_merge["TIPO_PAGO"] = ""  # Contabilidad no tiene TipoPago directamente
        cont_h_for_merge["UNIDAD"] = cont_h_for_merge["_UNIDAD_ORIG"].apply(norm_text)
        cont_h_for_merge["IMPORTE"] = cont_h_for_merge["IMPORTE_KEY"]
        
        liq = prep_liquidaciones(liq_raw, ndigits)
        st.success(f"✅ Liquidaciones cargadas: {len(liq):,} registros")
    
    # Control de matches globales
    used_cont_ids = set()
    resultado_final = {}
    
    st.divider()
    
    # ===== ETAPA 1: INGRESOS =====
    liq_clasificado, cont_h_clasificado, resumen_ingresos = etapa_1_ingresos(
        liq, cont_h_for_merge, ndigits
    )
    
    resultado_final["1_Liquidaciones_Clasificadas"] = liq_clasificado
    resultado_final["1_Contabilidad_H_Clasificada"] = cont_h_clasificado
    
    # Actualizar IDs usados
    matched_cont_h = cont_h_clasificado[cont_h_clasificado["ESTATUS_MATCH"].isin(["MATCH_OK", "MATCH_CON_DISCREPANCIA"])]
    used_cont_ids.update(matched_cont_h["ROW_ID_CONT"].dropna().unique())
    
    st.divider()
    
    # ===== ETAPA 2: COSTOS =====
    base = None
    vales = None
    
    if base_file and proceso_costos in {"Base Saldos vs Contabilidad D", "Ambos"}:
        with st.spinner("Cargando Base Saldos..."):
            base_raw = read_table(base_file)
            base = prep_base_saldos(base_raw, ndigits, concept_map)
            st.success(f"✅ Base Saldos cargada: {len(base):,} registros")
    
    if vales_file and proceso_costos in {"Vales vs Contabilidad D", "Ambos"}:
        with st.spinner("Cargando Vales..."):
            vales_raw = read_table(vales_file)
            vales = prep_vales(vales_raw, ndigits, concept_map)
            st.success(f"✅ Vales cargados: {len(vales):,} registros")
    
    resultado_costos, resumen_costos, used_cont_ids = etapa_2_costos(
        base, vales, cont_d, used_cont_ids, proceso_costos
    )
    
    resultado_final.update({f"2_{k}": v for k, v in resultado_costos.items()})
    
    st.divider()
    
    # ===== ETAPA 3: CROSSMATCH =====
    # Recolectar todos los NO_EXISTE de etapas anteriores
    no_existe_registros = []
    
    # De Liquidaciones
    liq_no_existe = liq_clasificado[liq_clasificado["ESTATUS_MATCH"] == "NO_EXISTE_EN_CONTABILIDAD"].copy()
    if not liq_no_existe.empty:
        liq_no_existe["ORIGEN_CROSSMATCH"] = "LIQUIDACIONES"
        no_existe_registros.append(liq_no_existe)
    
    # De Base Saldos
    if "Base_Clasificada" in resultado_costos:
        base_no_existe = resultado_costos["Base_Clasificada"][
            resultado_costos["Base_Clasificada"]["ESTATUS_MATCH"] == "NO_EXISTE_EN_CONTABILIDAD_D"
        ].copy()
        if not base_no_existe.empty:
            base_no_existe["ORIGEN_CROSSMATCH"] = "BASE_SALDOS"
            no_existe_registros.append(base_no_existe)
    
    # De Vales
    if "Vales_Clasificados" in resultado_costos:
        vales_no_existe = resultado_costos["Vales_Clasificados"][
            resultado_costos["Vales_Clasificados"]["ESTATUS_MATCH"] == "NO_EXISTE_EN_CONTABILIDAD_D"
        ].copy()
        if not vales_no_existe.empty:
            vales_no_existe["ORIGEN_CROSSMATCH"] = "VALES"
            no_existe_registros.append(vales_no_existe)
    
    if no_existe_registros:
        todos_no_existe = pd.concat(no_existe_registros, ignore_index=True)
        
        crossmatch_resultado, resumen_crossmatch = etapa_3_crossmatch(
            todos_no_existe, cont_all, used_cont_ids
        )
        
        resultado_final["3_Crossmatch_Analisis"] = crossmatch_resultado
    else:
        st.info("✨ No hay registros NO_EXISTE para analizar en Crossmatch")
        resumen_crossmatch = {}
    
    # ===== RESUMEN FINAL =====
    st.divider()
    st.header("📊 Resumen Consolidado Final")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("Etapa 1: Ingresos")
        st.metric("MATCH_OK", f"{resumen_ingresos['MATCH_OK']:,}")
        st.metric("DISCREPANCIA", f"{resumen_ingresos['MATCH_CON_DISCREPANCIA']:,}")
        st.metric("NO_EXISTE", f"{resumen_ingresos['NO_EXISTE']:,}")
    
    with col2:
        st.subheader("Etapa 2: Costos")
        if "Base" in resumen_costos:
            st.metric("Base MATCH_OK", f"{resumen_costos['Base']['MATCH_OK']:,}")
        if "Vales" in resumen_costos:
            st.metric("Vales MATCH_OK", f"{resumen_costos['Vales']['MATCH_OK']:,}")
    
    with col3:
        st.subheader("Etapa 3: Crossmatch")
        if resumen_crossmatch:
            st.metric("Completos", f"{resumen_crossmatch['COMPLETO']:,}")
            st.metric("Parciales", f"{resumen_crossmatch['PARCIAL']:,}")
            st.metric("No encontrado", f"{resumen_crossmatch['NO_ENCONTRADO']:,}")
    
    tiempo_total = time.time() - inicio_total
    st.success(f"✅ **Proceso consolidado completado en {tiempo_total:.1f} segundos**")
    
    # ===== PREVIEW DE RESULTADOS =====
    st.divider()
    st.header("📋 Resultados Detallados")
    
    tabs = st.tabs([
        "Etapa 1: Ingresos",
        "Etapa 2: Costos",
        "Etapa 3: Crossmatch",
        "Descargar Todo"
    ])
    
    with tabs[0]:
        st.subheader("Liquidaciones Clasificadas")
        st.dataframe(liq_clasificado, use_container_width=True, height=600)
        
        st.subheader("Contabilidad H Clasificada")
        st.dataframe(cont_h_clasificado, use_container_width=True, height=600)
    
    with tabs[1]:
        if "Base_Clasificada" in resultado_costos:
            st.subheader("Base Saldos Clasificada")
            st.dataframe(resultado_costos["Base_Clasificada"], use_container_width=True, height=600)
        
        if "Vales_Clasificados" in resultado_costos:
            st.subheader("Vales Clasificados")
            st.dataframe(resultado_costos["Vales_Clasificados"], use_container_width=True, height=600)
    
    with tabs[2]:
        if "3_Crossmatch_Analisis" in resultado_final:
            st.dataframe(resultado_final["3_Crossmatch_Analisis"], use_container_width=True, height=600)
        else:
            st.info("No se ejecutó crossmatch (no hay registros NO_EXISTE)")
    
    with tabs[3]:
        st.subheader("Descargar Excel Consolidado")
        st.caption("Incluye todas las etapas en hojas separadas")
        
        excel_bytes = generar_excel_consolidado(resultado_final)
        
        st.download_button(
            label="📥 Descargar Resultado Consolidado",
            data=excel_bytes,
            file_name=f"Saldos_Owner_Consolidado_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.info(f"📄 Total de hojas: {len(resultado_final)}")

except Exception as e:
    st.error(f"❌ Error durante el proceso: {str(e)}")
    st.exception(e)
