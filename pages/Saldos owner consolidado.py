"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                SALDOS OWNER - SISTEMA MODULAR POR ETAPAS                     ║
║                                                                              ║
║  Ejecuta cada etapa por separado para evitar límites de memoria             ║
║  Los resultados de una etapa se cargan como input de la siguiente           ║
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

st.set_page_config(page_title="Saldos Owner Modular", layout="wide")

# ============================================================
# HELPERS GENERALES
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


def normalizar_viaje(serie):
    return serie.fillna('').astype(str).str.replace('/', '', regex=False).str.replace('-', '', regex=False).str.strip().str.upper()


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
        raise ValueError(f"No encontré columna: {candidates}")
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
    importe_cols = resolve_all_cols(cont_raw, ["Importe", "Monto", "Total"])
    if not importe_cols:
        raise ValueError("No encontré columna de importe")
    return importe_cols[-1]


def build_seq(df: pd.DataFrame, key_cols: list[str], seq_col: str = "_seq") -> pd.DataFrame:
    out = df.copy()
    out[seq_col] = out.groupby(key_cols, dropna=False).cumcount() + 1
    return out


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter", engine_kwargs={"options": {"constant_memory": True}}) as writer:
        for name, df in sheets.items():
            if df is not None and not df.empty:
                out = df.copy()
                out.columns = [str(c)[:250] for c in out.columns]
                # Eliminar columnas auxiliares
                out = out[[c for c in out.columns if not str(c).startswith('_') and not str(c) in ['idx_original']]]
                out.to_excel(writer, sheet_name=name[:31], index=False)
    bio.seek(0)
    return bio.getvalue()


# ============================================================
# ETAPA 1: INGRESOS
# ============================================================

def ejecutar_etapa_1_ingresos(liq_file, cont_file, ndigits: int, liq_tipo: str):
    """Procesa solo Etapa 1: Liquidaciones vs Contabilidad H"""
    
    inicio = time.time()
    st.subheader("🔵 ETAPA 1: INGRESOS")
    
    # Cargar Liquidaciones
    with st.spinner("Cargando Liquidaciones..."):
        liq_usecols = ["Liquidacion", "Numero_Viaje", "TipoPago", "Monto", "Unidad", "Owner", "Tipo_Concepto"]
        liq_raw = read_table(liq_file, preferred_sheet="LiquidacionesSET_PLUS_datos", usecols=liq_usecols)
        
        liq = liq_raw.rename(columns={
            "Liquidacion": "PR",
            "Numero_Viaje": "VIAJE",
            "TipoPago": "TIPO_PAGO",
            "Monto": "IMPORTE",
            "Unidad": "UNIDAD",
            "Owner": "OWNER_LIQ",
            "Tipo_Concepto": "TIPO_CONCEPTO",
        })
        
        # Normalizar
        for c in ["PR", "VIAJE", "TIPO_PAGO", "UNIDAD", "OWNER_LIQ", "TIPO_CONCEPTO"]:
            if c in liq.columns:
                liq[c] = liq[c].apply(norm_text)
        
        liq["IMPORTE"] = pd.to_numeric(liq["IMPORTE"].apply(lambda x: norm_amount(x, ndigits)), errors="coerce")
        liq["ROW_ID_LIQ"] = range(1, len(liq) + 1)
        
        # Filtrar
        liq_f = liq[liq["TIPO_CONCEPTO"] == liq_tipo].copy()
        liq_f = liq_f.reset_index(drop=True)
        
        st.success(f"✅ Liquidaciones cargadas: {len(liq):,} | Filtradas: {len(liq_f):,}")
    
    # Cargar Contabilidad
    with st.spinner("Cargando Contabilidad..."):
        cont_usecols = ["Factura", "Referencia", "TipoPago", "Importe", "Unidad", "NombreCuentaContable", "TipoMovimiento"]
        cont_raw = read_table(cont_file, preferred_sheet="ContabilidadSET_PLUS_datos", usecols=cont_usecols)
        
        cont = cont_raw.rename(columns={
            "Factura": "PR",
            "Referencia": "VIAJE",
            "TipoPago": "TIPO_PAGO",
            "Importe": "IMPORTE",
            "Unidad": "UNIDAD",
            "NombreCuentaContable": "OWNER_CONT",
            "TipoMovimiento": "TIPO_MOV",
        })
        
        for c in ["PR", "VIAJE", "TIPO_PAGO", "UNIDAD", "OWNER_CONT", "TIPO_MOV"]:
            if c in cont.columns:
                cont[c] = cont[c].apply(norm_text)
        
        cont["IMPORTE"] = pd.to_numeric(cont["IMPORTE"].apply(lambda x: norm_amount(x, ndigits)), errors="coerce")
        cont["ROW_ID_CONT"] = range(1, len(cont) + 1)
        
        # Filtrar solo H
        cont_f = cont[cont["TIPO_MOV"] == "H"].copy()
        cont_f = cont_f.reset_index(drop=True)
        
        st.success(f"✅ Contabilidad cargada: {len(cont):,} | H filtrados: {len(cont_f):,}")
    
    # Matching
    with st.spinner("Ejecutando matching..."):
        key_cols = ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE"]
        
        liq_k = build_seq(liq_f, key_cols)
        cont_k = build_seq(cont_f, key_cols)
        
        merge_keys = key_cols + ["_seq"]
        
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
        
        # Clasificar
        matched["OWNER_MATCH"] = (matched["OWNER_LIQ"].fillna("") == matched["OWNER_CONT"].fillna(""))
        matched["ESTATUS_MATCH"] = matched["OWNER_MATCH"].map({True: "MATCH_OK", False: "MATCH_CON_DISCREPANCIA"})
        only_liq["ESTATUS_MATCH"] = "NO_EXISTE_EN_CONTABILIDAD"
        only_cont["ESTATUS_MATCH"] = "NO_EXISTE_EN_LIQUIDACIONES"
    
    # Construir resultados
    liq_status = pd.concat([
        matched[["ROW_ID_LIQ", "ROW_ID_CONT", "ESTATUS_MATCH", "OWNER_CONT"]],
        only_liq[["ROW_ID_LIQ", "ESTATUS_MATCH"]].assign(ROW_ID_CONT=pd.NA, OWNER_CONT="")
    ], ignore_index=True)
    
    cont_status = pd.concat([
        matched[["ROW_ID_CONT", "ROW_ID_LIQ", "ESTATUS_MATCH", "OWNER_LIQ"]],
        only_cont[["ROW_ID_CONT", "ESTATUS_MATCH"]].assign(ROW_ID_LIQ=pd.NA, OWNER_LIQ="")
    ], ignore_index=True)
    
    liq_clasificado = liq_f.merge(liq_status, on="ROW_ID_LIQ", how="left")
    cont_clasificado = cont_f.merge(cont_status, on="ROW_ID_CONT", how="left")
    
    # Marcar matcheados
    liq_clasificado["MATCHED_IN_ETAPA"] = liq_clasificado["ESTATUS_MATCH"].apply(
        lambda x: "ETAPA_1_INGRESOS" if x in ["MATCH_OK", "MATCH_CON_DISCREPANCIA"] else None
    )
    cont_clasificado["MATCHED_IN_ETAPA"] = cont_clasificado["ESTATUS_MATCH"].apply(
        lambda x: "ETAPA_1_INGRESOS" if x in ["MATCH_OK", "MATCH_CON_DISCREPANCIA"] else None
    )
    
    # Resumen
    resumen = {
        "MATCH_OK": int((liq_clasificado["ESTATUS_MATCH"] == "MATCH_OK").sum()),
        "MATCH_CON_DISCREPANCIA": int((liq_clasificado["ESTATUS_MATCH"] == "MATCH_CON_DISCREPANCIA").sum()),
        "NO_EXISTE": int((liq_clasificado["ESTATUS_MATCH"] == "NO_EXISTE_EN_CONTABILIDAD").sum()),
        "TOTAL": len(liq_clasificado)
    }
    
    tiempo = time.time() - inicio
    st.success(f"✅ Etapa 1 completada en {tiempo:.1f}s")
    
    # Métricas
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total", f"{resumen['TOTAL']:,}")
    col2.metric("MATCH_OK", f"{resumen['MATCH_OK']:,}")
    col3.metric("DISCREPANCIA", f"{resumen['MATCH_CON_DISCREPANCIA']:,}")
    col4.metric("NO_EXISTE", f"{resumen['NO_EXISTE']:,}")
    
    # Preparar Excel para siguiente etapa
    matched_cont_ids = cont_clasificado[
        cont_clasificado["ESTATUS_MATCH"].isin(["MATCH_OK", "MATCH_CON_DISCREPANCIA"])
    ]["ROW_ID_CONT"].dropna().unique().tolist()
    
    # Crear hoja de IDs matcheados para Etapa 2
    ids_matcheados = pd.DataFrame({
        "ROW_ID_CONT_MATCHEADO": matched_cont_ids,
        "ETAPA_ORIGEN": "ETAPA_1_INGRESOS"
    })
    
    resultado = {
        "Liquidaciones_Clasificadas": liq_clasificado,
        "Contabilidad_H_Clasificada": cont_clasificado,
        "IDs_Matcheados_Etapa1": ids_matcheados,
        "Resumen": pd.DataFrame([resumen])
    }
    
    return resultado


# ============================================================
# ETAPA 2: COSTOS
# ============================================================

def ejecutar_etapa_2_costos(cont_file, base_file, vales_file, ids_etapa1_file, 
                            ndigits: int, proceso: str, concept_map: dict):
    """Procesa Etapa 2: Base/Vales vs Contabilidad D"""
    
    inicio = time.time()
    st.subheader("🟢 ETAPA 2: COSTOS")
    
    # Cargar IDs ya matcheados de Etapa 1
    used_cont_ids = set()
    if ids_etapa1_file:
        with st.spinner("Cargando IDs matcheados de Etapa 1..."):
            ids_df = pd.read_excel(BytesIO(ids_etapa1_file.getvalue()), sheet_name="IDs_Matcheados_Etapa1")
            used_cont_ids = set(ids_df["ROW_ID_CONT_MATCHEADO"].dropna().unique())
            st.info(f"📌 Excluyendo {len(used_cont_ids):,} registros ya matcheados en Etapa 1")
    
    # Cargar Contabilidad
    with st.spinner("Cargando Contabilidad..."):
        cont_raw = read_table(cont_file, preferred_sheet="ContabilidadSET_PLUS_datos")
        
        c_mov = resolve_col(cont_raw, ["TipoMovimiento", "Movimiento"])
        c_importe = choose_cont_import_col(cont_raw)
        c_unidad = resolve_col(cont_raw, ["Unidad", "Numero de Unidad", "Numero_Unidad"])
        c_referencia = resolve_col(cont_raw, ["Referencia", "Numero_Viaje", "Viaje"], required=False)
        c_poliza = resolve_col(cont_raw, ["Clave Poliza", "ClavePoliza", "Factura"])
        c_concepto = resolve_col(cont_raw, ["Concepto detalle", "ConceptoDetalle", "Concepto"], required=False)
        c_vale = resolve_col(cont_raw, ["Vale", "No Vale"], required=False)
        
        cont = cont_raw.copy()
        cont["TIPO_MOV"] = cont[c_mov].apply(norm_text)
        cont = cont[cont["TIPO_MOV"] == "D"].copy()  # Solo D
        
        cont["POLIZA_KEY"] = cont[c_poliza].apply(norm_for_key)
        cont["UNIDAD_KEY"] = cont[c_unidad].apply(norm_for_key)
        cont["VIAJE_KEY"] = cont[c_referencia].apply(norm_for_key) if c_referencia else ""
        cont["VALE_KEY"] = cont[c_vale].apply(norm_for_key) if c_vale else ""
        cont["CONCEPTO_KEY"] = cont[c_concepto].apply(lambda x: canonical_concept(x, concept_map)) if c_concepto else ""
        cont["IMPORTE_KEY"] = cont[c_importe].apply(lambda x: norm_amount(x, ndigits))
        cont["ROW_ID_CONT"] = range(1, len(cont) + 1)
        
        # Filtrar ya matcheados
        cont_disponible = cont[~cont["ROW_ID_CONT"].isin(used_cont_ids)].copy()
        
        st.success(f"✅ Contabilidad D: {len(cont):,} | Disponible: {len(cont_disponible):,}")
    
    resultado = {}
    
    # ========== PROCESAR BASE SALDOS ==========
    if base_file and proceso in {"Base Saldos vs Contabilidad D", "Ambos"}:
        with st.spinner("Procesando Base Saldos..."):
            base_raw = read_table(base_file)
            
            c_poliza = resolve_col(base_raw, ["folio_contrarecibo", "contrarecibo", "FOLIO_CONTRARECIBO"])
            c_unidad = resolve_col(base_raw, ["numero de unidad", "numero_unidad", "unidad"])
            c_viaje = resolve_col(base_raw, ["numero_viaje", "numero viaje", "viaje", "NUMERO_VIAJE"])
            c_concepto = resolve_col(base_raw, ["concepto_contabilidad", "concepto contabilidad", "concepto", "Concepto contabilidad"])
            c_importe = resolve_col(base_raw, ["importe", "monto", "total", "Importe"])
            
            base = base_raw.copy()
            base["POLIZA_KEY"] = base[c_poliza].apply(norm_for_key)
            base["UNIDAD_KEY"] = base[c_unidad].apply(norm_for_key)
            base["VIAJE_KEY"] = base[c_viaje].apply(norm_for_key)
            base["CONCEPTO_KEY"] = base[c_concepto].apply(lambda x: canonical_concept(x, concept_map))
            base["IMPORTE_KEY"] = base[c_importe].apply(lambda x: norm_amount(x, ndigits))
            base["ROW_ID_BASE"] = range(1, len(base) + 1)
            
            # Matching simplificado
            matches = []
            
            # Bloque 1: Poliza + Importe
            m1 = base[["ROW_ID_BASE", "POLIZA_KEY", "IMPORTE_KEY"]].dropna().merge(
                cont_disponible[["ROW_ID_CONT", "POLIZA_KEY", "IMPORTE_KEY"]].dropna(),
                on=["POLIZA_KEY", "IMPORTE_KEY"],
                how="inner"
            )
            if not m1.empty:
                matches.append(m1)
            
            # Bloque 2: Unidad + Viaje + Importe
            m2 = base[["ROW_ID_BASE", "UNIDAD_KEY", "VIAJE_KEY", "IMPORTE_KEY"]].dropna().merge(
                cont_disponible[["ROW_ID_CONT", "UNIDAD_KEY", "VIAJE_KEY", "IMPORTE_KEY"]].dropna(),
                on=["UNIDAD_KEY", "VIAJE_KEY", "IMPORTE_KEY"],
                how="inner"
            )
            if not m2.empty:
                matches.append(m2)
            
            if matches:
                all_matches = pd.concat(matches, ignore_index=True).drop_duplicates()
                
                # Greedy: tomar solo primer match por cada lado
                all_matches = all_matches.sort_values(["ROW_ID_BASE", "ROW_ID_CONT"])
                all_matches = all_matches.drop_duplicates(subset=["ROW_ID_BASE"], keep="first")
                all_matches = all_matches.drop_duplicates(subset=["ROW_ID_CONT"], keep="first")
                
                all_matches["ESTATUS_MATCH"] = "MATCH_OK"
            else:
                all_matches = pd.DataFrame(columns=["ROW_ID_BASE", "ROW_ID_CONT", "ESTATUS_MATCH"])
            
            # Clasificar
            base_clas = base.merge(all_matches[["ROW_ID_BASE", "ROW_ID_CONT", "ESTATUS_MATCH"]], on="ROW_ID_BASE", how="left")
            base_clas["ESTATUS_MATCH"] = base_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_CONTABILIDAD_D")
            base_clas["MATCHED_IN_ETAPA"] = base_clas["ESTATUS_MATCH"].apply(
                lambda x: "ETAPA_2_COSTOS_BASE" if x == "MATCH_OK" else None
            )
            
            resultado["Base_Clasificada"] = base_clas
            
            # Actualizar IDs usados
            if not all_matches.empty:
                used_cont_ids.update(all_matches["ROW_ID_CONT"].dropna().unique())
            
            match_ok = int((base_clas["ESTATUS_MATCH"] == "MATCH_OK").sum())
            st.write(f"✅ Base Saldos: {match_ok:,} matches de {len(base):,}")
    
    # ========== PROCESAR VALES ==========
    if vales_file and proceso in {"Vales vs Contabilidad D", "Ambos"}:
        # Actualizar disponibles
        cont_disponible = cont[~cont["ROW_ID_CONT"].isin(used_cont_ids)].copy()
        
        with st.spinner("Procesando Vales..."):
            vales_raw = read_table(vales_file)
            
            c_vale = resolve_col(vales_raw, ["Vale", "No Vale"])
            c_unidad = resolve_col(vales_raw, ["Unidad", "Numero de Unidad"])
            c_concepto = resolve_col(vales_raw, ["Concepto", "Concepto detalle"])
            c_importe = resolve_col(vales_raw, ["Total", "Importe", "TotalVale"])
            
            vales = vales_raw.copy()
            vales["VALE_KEY"] = vales[c_vale].apply(norm_for_key)
            vales["UNIDAD_KEY"] = vales[c_unidad].apply(norm_for_key)
            vales["CONCEPTO_KEY"] = vales[c_concepto].apply(lambda x: canonical_concept(x, concept_map))
            vales["IMPORTE_KEY"] = vales[c_importe].apply(lambda x: norm_amount(x, ndigits))
            vales["ROW_ID_VALE"] = range(1, len(vales) + 1)
            
            # Matching
            matches = []
            
            # Unidad + Concepto + Importe
            m1 = vales[["ROW_ID_VALE", "UNIDAD_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"]].dropna().merge(
                cont_disponible[["ROW_ID_CONT", "UNIDAD_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"]].dropna(),
                on=["UNIDAD_KEY", "CONCEPTO_KEY", "IMPORTE_KEY"],
                how="inner"
            )
            if not m1.empty:
                matches.append(m1)
            
            # Vale + Importe
            m2 = vales[["ROW_ID_VALE", "VALE_KEY", "IMPORTE_KEY"]].dropna().merge(
                cont_disponible[["ROW_ID_CONT", "VALE_KEY", "IMPORTE_KEY"]].dropna(),
                on=["VALE_KEY", "IMPORTE_KEY"],
                how="inner"
            )
            if not m2.empty:
                matches.append(m2)
            
            if matches:
                all_matches = pd.concat(matches, ignore_index=True).drop_duplicates()
                all_matches = all_matches.sort_values(["ROW_ID_VALE", "ROW_ID_CONT"])
                all_matches = all_matches.drop_duplicates(subset=["ROW_ID_VALE"], keep="first")
                all_matches = all_matches.drop_duplicates(subset=["ROW_ID_CONT"], keep="first")
                all_matches["ESTATUS_MATCH"] = "MATCH_OK"
            else:
                all_matches = pd.DataFrame(columns=["ROW_ID_VALE", "ROW_ID_CONT", "ESTATUS_MATCH"])
            
            vales_clas = vales.merge(all_matches[["ROW_ID_VALE", "ROW_ID_CONT", "ESTATUS_MATCH"]], on="ROW_ID_VALE", how="left")
            vales_clas["ESTATUS_MATCH"] = vales_clas["ESTATUS_MATCH"].fillna("NO_EXISTE_EN_CONTABILIDAD_D")
            vales_clas["MATCHED_IN_ETAPA"] = vales_clas["ESTATUS_MATCH"].apply(
                lambda x: "ETAPA_2_COSTOS_VALES" if x == "MATCH_OK" else None
            )
            
            resultado["Vales_Clasificados"] = vales_clas
            
            if not all_matches.empty:
                used_cont_ids.update(all_matches["ROW_ID_CONT"].dropna().unique())
            
            match_ok = int((vales_clas["ESTATUS_MATCH"] == "MATCH_OK").sum())
            st.write(f"✅ Vales: {match_ok:,} matches de {len(vales):,}")
    
    # Guardar IDs para Etapa 3
    ids_nuevos = list(used_cont_ids - (set(ids_df["ROW_ID_CONT_MATCHEADO"].dropna().unique()) if ids_etapa1_file else set()))
    ids_matcheados = pd.DataFrame({
        "ROW_ID_CONT_MATCHEADO": ids_nuevos,
        "ETAPA_ORIGEN": "ETAPA_2_COSTOS"
    })
    
    resultado["IDs_Matcheados_Etapa2"] = ids_matcheados
    
    tiempo = time.time() - inicio
    st.success(f"✅ Etapa 2 completada en {tiempo:.1f}s")
    
    return resultado


# ============================================================
# ETAPA 3: CROSSMATCH
# ============================================================

def ejecutar_etapa_3_crossmatch(base_no_existe_file, cont_file, ids_previos_file, ndigits: int):
    """Procesa Etapa 3: Crossmatch de NO_EXISTE"""
    
    inicio = time.time()
    st.subheader("🟣 ETAPA 3: CROSSMATCH")
    
    # Cargar IDs ya matcheados
    used_cont_ids = set()
    if ids_previos_file:
        with st.spinner("Cargando IDs matcheados previos..."):
            # Leer todas las hojas de IDs
            excel_file = pd.ExcelFile(BytesIO(ids_previos_file.getvalue()))
            for sheet in excel_file.sheet_names:
                if "IDs_Matcheados" in sheet:
                    ids_df = pd.read_excel(BytesIO(ids_previos_file.getvalue()), sheet_name=sheet)
                    if "ROW_ID_CONT_MATCHEADO" in ids_df.columns:
                        used_cont_ids.update(ids_df["ROW_ID_CONT_MATCHEADO"].dropna().unique())
            st.info(f"📌 Excluyendo {len(used_cont_ids):,} registros ya matcheados")
    
    # Cargar registros NO_EXISTE
    with st.spinner("Cargando registros NO_EXISTE..."):
        base_raw = pd.read_excel(BytesIO(base_no_existe_file.getvalue()))
        st.success(f"✅ Registros NO_EXISTE: {len(base_raw):,}")
    
    # Cargar Contabilidad completa
    with st.spinner("Cargando Contabilidad..."):
        cont_raw = read_table(cont_file, preferred_sheet="ContabilidadSET_PLUS_datos")
        
        c_mov = resolve_col(cont_raw, ["TipoMovimiento"])
        c_importe = choose_cont_import_col(cont_raw)
        c_unidad = resolve_col(cont_raw, ["Unidad"])
        c_referencia = resolve_col(cont_raw, ["Referencia"], required=False)
        c_poliza = resolve_col(cont_raw, ["ClavePoliza", "Clave Poliza", "Factura"])
        c_concepto = resolve_col(cont_raw, ["ConceptoDetalle", "Concepto detalle"], required=False)
        c_owner = resolve_col(cont_raw, ["NombreCuentaContable"], required=False)
        
        cont = cont_raw.copy()
        cont["TIPO_MOV"] = cont[c_mov].apply(norm_text)
        cont["POLIZA_KEY"] = cont[c_poliza].apply(norm_for_key)
        cont["VIAJE_KEY"] = cont[c_referencia].apply(norm_for_key) if c_referencia else ""
        cont["IMPORTE_KEY"] = cont[c_importe].apply(lambda x: norm_amount(x, ndigits))
        cont["CONCEPTO_KEY"] = cont[c_concepto].apply(norm_text) if c_concepto else ""
        cont["OWNER_CONT"] = cont[c_owner].apply(norm_text) if c_owner else ""
        cont["ROW_ID_CONT"] = range(1, len(cont) + 1)
        cont["_UNIDAD_ORIG"] = cont[c_unidad]
        cont["_VIAJE_ORIG"] = cont[c_referencia] if c_referencia else ""
        cont["_POLIZA_ORIG"] = cont[c_poliza]
        
        # Filtrar ya matcheados
        cont = cont[~cont["ROW_ID_CONT"].isin(used_cont_ids)].copy()
        
        st.success(f"✅ Contabilidad disponible: {len(cont):,}")
    
    # Preparar base
    with st.spinner("Ejecutando crossmatch..."):
        base = base_raw.copy()
        base['idx_original'] = range(len(base))
        
        # Intentar diferentes nombres de columnas
        if 'POLIZA_KEY' in base.columns:
            base['poliza_norm'] = base['POLIZA_KEY'].fillna('').astype(str).str.strip().str.upper()
        elif 'FOLIO_CONTRARECIBO' in base.columns:
            base['poliza_norm'] = base['FOLIO_CONTRARECIBO'].fillna('').astype(str).str.strip().str.upper()
        else:
            base['poliza_norm'] = ''
        
        if 'VIAJE_KEY' in base.columns:
            base['viaje_norm'] = normalizar_viaje(base['VIAJE_KEY'])
        elif 'NUMERO_VIAJE' in base.columns:
            base['viaje_norm'] = normalizar_viaje(base['NUMERO_VIAJE'])
        else:
            base['viaje_norm'] = ''
        
        if 'IMPORTE_KEY' in base.columns:
            base['importe'] = pd.to_numeric(base['IMPORTE_KEY'], errors='coerce').fillna(0).round(2)
        elif 'Importe' in base.columns:
            base['importe'] = pd.to_numeric(base['Importe'], errors='coerce').fillna(0).round(2)
        elif 'IMPORTE' in base.columns:
            base['importe'] = pd.to_numeric(base['IMPORTE'], errors='coerce').fillna(0).round(2)
        else:
            base['importe'] = 0
        
        base['concepto_norm'] = base.get('CONCEPTO_KEY', base.get('Concepto contabilidad', '')).fillna('').astype(str).str.upper()
        base['es_diesel'] = base['concepto_norm'].str.contains('DIESEL|CONSUMIBLES', na=False)
        
        # Preparar contabilidad
        cont['poliza_norm'] = cont['POLIZA_KEY'].fillna('').astype(str).str.strip().str.upper()
        cont['viaje_norm'] = normalizar_viaje(cont['VIAJE_KEY'])
        cont['importe'] = pd.to_numeric(cont['IMPORTE_KEY'], errors='coerce').fillna(0).round(2)
        cont['concepto_norm'] = cont['CONCEPTO_KEY'].fillna('').astype(str).str.upper()
        cont['tipo_poliza'] = cont['POLIZA_KEY'].fillna('').astype(str).str[:2]
        
        # Separar tipos
        cont_d = cont[cont['TIPO_MOV'] == 'D'].copy()
        cont_h = cont[cont['TIPO_MOV'] == 'H'].copy()
        cont_d_ca = cont_d[cont_d['tipo_poliza'] == 'CA'].copy()
        cont_d_pd = cont_d[cont_d['tipo_poliza'] == 'PD'].copy()
        cont_h_no_ca = cont_h[~cont_h['tipo_poliza'].isin(['CA'])].copy()
        
        # Buscar CA
        if not cont_d_ca.empty:
            matches_ca = base[['idx_original', 'poliza_norm', 'importe']].merge(
                cont_d_ca[['poliza_norm', 'importe', '_UNIDAD_ORIG', '_VIAJE_ORIG']],
                on=['poliza_norm', 'importe'],
                how='left'
            ).groupby('idx_original').first().reset_index()
            
            base = base.merge(
                matches_ca[['idx_original', '_UNIDAD_ORIG', '_VIAJE_ORIG']].rename(columns={
                    '_UNIDAD_ORIG': 'ca_unidad',
                    '_VIAJE_ORIG': 'ca_viaje'
                }),
                on='idx_original',
                how='left'
            )
            base['tiene_cargo_ca'] = base['ca_unidad'].notna()
        else:
            base['tiene_cargo_ca'] = False
        
        # Buscar PD
        if not cont_d_pd.empty:
            cont_pd_diesel = cont_d_pd[cont_d_pd['concepto_norm'].str.contains('DIESEL', na=False)].copy()
            if not cont_pd_diesel.empty and base['es_diesel'].any():
                matches_pd = base[base['es_diesel']][['idx_original', 'viaje_norm', 'importe']].merge(
                    cont_pd_diesel[['viaje_norm', 'importe', '_POLIZA_ORIG']],
                    on=['viaje_norm', 'importe'],
                    how='left'
                ).groupby('idx_original').first().reset_index()
                
                base = base.merge(
                    matches_pd[['idx_original', '_POLIZA_ORIG']].rename(columns={'_POLIZA_ORIG': 'pd_poliza'}),
                    on='idx_original',
                    how='left'
                )
                base['tiene_pd_exacto'] = base['pd_poliza'].notna()
            else:
                base['tiene_pd_exacto'] = False
        else:
            base['tiene_pd_exacto'] = False
        
        # Buscar H
        if not cont_h_no_ca.empty:
            matches_h = base[['idx_original', 'viaje_norm', 'importe']].merge(
                cont_h_no_ca[['viaje_norm', 'importe', '_POLIZA_ORIG', 'OWNER_CONT']],
                on=['viaje_norm', 'importe'],
                how='left'
            ).groupby('idx_original').first().reset_index()
            
            base = base.merge(
                matches_h[['idx_original', '_POLIZA_ORIG', 'OWNER_CONT']].rename(columns={
                    '_POLIZA_ORIG': 'h_poliza',
                    'OWNER_CONT': 'h_owner'
                }),
                on='idx_original',
                how='left'
            )
            base['tiene_abono_h'] = base['h_poliza'].notna()
        else:
            base['tiene_abono_h'] = False
        
        # Clasificar
        def clasificar(row):
            ca = row.get('tiene_cargo_ca', False)
            pd = row.get('tiene_pd_exacto', False)
            h = row.get('tiene_abono_h', False)
            
            if ca and h:
                return 'COMPLETO_CA_H'
            elif pd and h:
                return 'COMPLETO_PD_H'
            elif ca:
                return 'SOLO_CARGO_CA'
            elif pd:
                return 'SOLO_CARGO_PD'
            elif h:
                return 'SOLO_ABONO_H'
            else:
                return 'NO_ENCONTRADO_CROSSMATCH'
        
        base['TIPO_CASO_CROSSMATCH'] = base.apply(clasificar, axis=1)
        
        # Diagnosticos
        diagnosticos = []
        for _, row in base.iterrows():
            tipo = row['TIPO_CASO_CROSSMATCH']
            if tipo == 'COMPLETO_CA_H':
                diag = f"✅ CA + H {row.get('h_poliza', '')}"
            elif tipo == 'COMPLETO_PD_H':
                diag = f"✅ PD {row.get('pd_poliza', '')} + H {row.get('h_poliza', '')}"
            elif tipo == 'SOLO_CARGO_CA':
                diag = f"⚠️ Solo CA"
            elif tipo == 'SOLO_CARGO_PD':
                diag = f"⚠️ Solo PD {row.get('pd_poliza', '')}"
            elif tipo == 'SOLO_ABONO_H':
                diag = f"🔄 Solo H {row.get('h_poliza', '')}"
            else:
                diag = "❌ No encontrado"
            diagnosticos.append(diag)
        
        base['DIAGNOSTICO_CROSSMATCH'] = diagnosticos
    
    resumen = {
        "COMPLETO": int(base['TIPO_CASO_CROSSMATCH'].str.contains('COMPLETO', na=False).sum()),
        "PARCIAL": int(base['TIPO_CASO_CROSSMATCH'].str.contains('SOLO', na=False).sum()),
        "NO_ENCONTRADO": int((base['TIPO_CASO_CROSSMATCH'] == 'NO_ENCONTRADO_CROSSMATCH').sum()),
        "TOTAL": len(base)
    }
    
    tiempo = time.time() - inicio
    st.success(f"✅ Etapa 3 completada en {tiempo:.1f}s")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total", f"{resumen['TOTAL']:,}")
    col2.metric("Completos", f"{resumen['COMPLETO']:,}")
    col3.metric("Parciales", f"{resumen['PARCIAL']:,}")
    col4.metric("No encontrado", f"{resumen['NO_ENCONTRADO']:,}")
    
    resultado = {
        "Crossmatch_Analisis": base,
        "Resumen": pd.DataFrame([resumen])
    }
    
    return resultado


# ============================================================
# UI PRINCIPAL
# ============================================================

st.title("🎯 Saldos Owner - Sistema Modular")
st.caption("Ejecuta cada etapa por separado para evitar límites de memoria")

# Selector de etapa
etapa = st.radio(
    "Selecciona la etapa a ejecutar",
    ["ETAPA 1: Ingresos", "ETAPA 2: Costos", "ETAPA 3: Crossmatch"],
    horizontal=True
)

st.divider()

# ============================================================
# ETAPA 1: INGRESOS
# ============================================================

if etapa == "ETAPA 1: Ingresos":
    st.markdown("""
    ### 🔵 ETAPA 1: INGRESOS
    Procesa **Liquidaciones vs Contabilidad H**
    
    **Resultado:** Excel con Liquidaciones y Contabilidad clasificadas + IDs matcheados para Etapa 2
    """)
    
    with st.sidebar:
        st.header("Archivos Etapa 1")
        liq_file = st.file_uploader("Liquidaciones", type=["xlsx", "csv"], key="liq1")
        cont_file = st.file_uploader("Contabilidad", type=["xlsx", "csv"], key="cont1")
        
        st.divider()
        liq_tipo = st.selectbox("Tipo Concepto", ["E", "I"])
        ndigits = st.number_input("Redondeo", 0, 4, 2)
        
        ejecutar = st.button("▶️ EJECUTAR ETAPA 1", type="primary")
    
    if ejecutar and liq_file and cont_file:
        try:
            resultado = ejecutar_etapa_1_ingresos(liq_file, cont_file, ndigits, liq_tipo)
            
            st.divider()
            st.subheader("📋 Resultados")
            
            tab1, tab2, tab3 = st.tabs(["Liquidaciones", "Contabilidad H", "Descargar"])
            
            with tab1:
                st.dataframe(resultado["Liquidaciones_Clasificadas"], use_container_width=True, height=600)
            
            with tab2:
                st.dataframe(resultado["Contabilidad_H_Clasificada"], use_container_width=True, height=600)
            
            with tab3:
                excel = to_excel_bytes(resultado)
                st.download_button(
                    "📥 Descargar Resultado Etapa 1",
                    excel,
                    f"Etapa1_Ingresos_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.warning("⚠️ **IMPORTANTE:** Guarda este archivo para usar en Etapa 2")
        
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
    
    elif ejecutar:
        st.error("Falta cargar archivos")


# ============================================================
# ETAPA 2: COSTOS
# ============================================================

elif etapa == "ETAPA 2: Costos":
    st.markdown("""
    ### 🟢 ETAPA 2: COSTOS
    Procesa **Base Saldos/Vales vs Contabilidad D**
    
    **Requiere:** Archivo resultado de Etapa 1 (para excluir matcheados)
    
    **Resultado:** Excel con Base/Vales clasificados + IDs matcheados para Etapa 3
    """)
    
    with st.sidebar:
        st.header("Archivos Etapa 2")
        cont_file = st.file_uploader("Contabilidad", type=["xlsx", "csv"], key="cont2")
        ids_etapa1 = st.file_uploader("Resultado Etapa 1 (.xlsx)", type=["xlsx"], key="ids1")
        
        st.divider()
        
        base_file = st.file_uploader("Base Saldos (opcional)", type=["xlsx", "csv"], key="base2")
        vales_file = st.file_uploader("Vales (opcional)", type=["xlsx", "csv"], key="vales2")
        
        proceso = st.radio("Procesar", ["Base Saldos vs Contabilidad D", "Vales vs Contabilidad D", "Ambos"])
        
        st.divider()
        ndigits = st.number_input("Redondeo", 0, 4, 2)
        
        ejecutar = st.button("▶️ EJECUTAR ETAPA 2", type="primary")
    
    if ejecutar and cont_file:
        try:
            concept_map = {}
            
            resultado = ejecutar_etapa_2_costos(
                cont_file, base_file, vales_file, ids_etapa1,
                ndigits, proceso, concept_map
            )
            
            st.divider()
            st.subheader("📋 Resultados")
            
            tabs = []
            if "Base_Clasificada" in resultado:
                tabs.append("Base Saldos")
            if "Vales_Clasificados" in resultado:
                tabs.append("Vales")
            tabs.append("Descargar")
            
            tab_objs = st.tabs(tabs)
            
            idx = 0
            if "Base_Clasificada" in resultado:
                with tab_objs[idx]:
                    st.dataframe(resultado["Base_Clasificada"], use_container_width=True, height=600)
                idx += 1
            
            if "Vales_Clasificados" in resultado:
                with tab_objs[idx]:
                    st.dataframe(resultado["Vales_Clasificados"], use_container_width=True, height=600)
                idx += 1
            
            with tab_objs[idx]:
                excel = to_excel_bytes(resultado)
                st.download_button(
                    "📥 Descargar Resultado Etapa 2",
                    excel,
                    f"Etapa2_Costos_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.warning("⚠️ **IMPORTANTE:** Guarda este archivo para usar en Etapa 3")
        
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
    
    elif ejecutar:
        st.error("Falta cargar archivo de Contabilidad")


# ============================================================
# ETAPA 3: CROSSMATCH
# ============================================================

else:  # ETAPA 3
    st.markdown("""
    ### 🟣 ETAPA 3: CROSSMATCH
    Analiza registros **NO_EXISTE** en pólizas CA/PD/H
    
    **Requiere:** 
    - Archivo con registros NO_EXISTE (de Etapa 1 o 2)
    - Resultados de etapas previas (para excluir matcheados)
    
    **Resultado:** Excel con análisis crossmatch
    """)
    
    with st.sidebar:
        st.header("Archivos Etapa 3")
        cont_file = st.file_uploader("Contabilidad", type=["xlsx", "csv"], key="cont3")
        no_existe_file = st.file_uploader("Registros NO_EXISTE (.xlsx)", type=["xlsx"], key="noexiste")
        ids_previos = st.file_uploader("Resultados Etapas Previas (.xlsx)", type=["xlsx"], key="idsprev")
        
        st.divider()
        ndigits = st.number_input("Redondeo", 0, 4, 2)
        
        ejecutar = st.button("▶️ EJECUTAR ETAPA 3", type="primary")
    
    if ejecutar and cont_file and no_existe_file:
        try:
            resultado = ejecutar_etapa_3_crossmatch(no_existe_file, cont_file, ids_previos, ndigits)
            
            st.divider()
            st.subheader("📋 Resultados")
            
            tab1, tab2 = st.tabs(["Análisis Crossmatch", "Descargar"])
            
            with tab1:
                st.dataframe(resultado["Crossmatch_Analisis"], use_container_width=True, height=600)
            
            with tab2:
                excel = to_excel_bytes(resultado)
                st.download_button(
                    "📥 Descargar Resultado Etapa 3",
                    excel,
                    f"Etapa3_Crossmatch_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
    
    elif ejecutar:
        st.error("Faltan archivos requeridos")
