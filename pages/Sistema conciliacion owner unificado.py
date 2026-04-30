"""
Saldos Owner - Desarrollo de Costos v2.2
Mejora sobre v2.1 con detección automática de casos cross-match por póliza

NUEVAS FUNCIONALIDADES:
1. Detecta casos donde el costo está en un tráfico diferente (misma póliza, diferente unidad/viaje)
2. Identifica matches por importe exacto o similar
3. Genera análisis de movimientos D/H cruzados
4. Clasifica automáticamente los tipos de discrepancia
"""

import re
import unicodedata
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Saldos Owner - Costos v2.2", layout="wide")

# ============================================================
# Helpers generales (mantienen compatibilidad con v2.1)
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
        raise ValueError(f"No encontre ninguna columna de estas opciones: {candidates}")
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
# NUEVA FUNCIONALIDAD: Cross-Match por Póliza
# ============================================================

def analizar_crossmatch_por_poliza(base_no_existe: pd.DataFrame, cont_d: pd.DataFrame, cont_h: pd.DataFrame, 
                                   ndigits: int = 2) -> pd.DataFrame:
    """
    Analiza casos NO_EXISTE_EN_CONTABILIDAD_D buscando la misma póliza en contabilidad
    con diferente unidad/viaje (cross-match por póliza)
    """
    
    # Crear índices por póliza
    cont_d_por_poliza = {}
    for _, row in cont_d.iterrows():
        poliza = row['POLIZA_KEY']
        if poliza not in cont_d_por_poliza:
            cont_d_por_poliza[poliza] = []
        cont_d_por_poliza[poliza].append(row)
    
    cont_h_por_poliza = {}
    for _, row in cont_h.iterrows():
        poliza = row['POLIZA_KEY']
        if poliza not in cont_h_por_poliza:
            cont_h_por_poliza[poliza] = []
        cont_h_por_poliza[poliza].append(row)
    
    # Analizar cada registro
    resultados = []
    
    for _, row in base_no_existe.iterrows():
        poliza_key = row['POLIZA_KEY']
        importe_base = row['IMPORTE_KEY']
        
        resultado = {
            'CROSSMATCH_POLIZA_EN_D': False,
            'CROSSMATCH_POLIZA_EN_H': False,
            'CROSSMATCH_IMPORTE_EXACTO_D': False,
            'CROSSMATCH_IMPORTE_SIMILAR_D': False,
            'CROSSMATCH_IMPORTE_EXACTO_H': False,
            'CROSSMATCH_UNIDADES_D': '',
            'CROSSMATCH_VIAJES_D': '',
            'CROSSMATCH_IMPORTES_D': '',
            'CROSSMATCH_OWNERS_D': '',
            'CROSSMATCH_UNIDADES_H': '',
            'CROSSMATCH_VIAJES_H': '',
            'CROSSMATCH_TIPO': 'NO_CROSSMATCH',
            'CROSSMATCH_DIAGNOSTICO': ''
        }
        
        # Buscar en movimientos D
        if poliza_key in cont_d_por_poliza:
            resultado['CROSSMATCH_POLIZA_EN_D'] = True
            movs_d = cont_d_por_poliza[poliza_key]
            
            unidades = set()
            viajes = set()
            importes = []
            owners = set()
            
            for mov in movs_d:
                unidades.add(mov['UNIDAD_KEY'])
                viajes.add(mov['VIAJE_KEY'])
                importes.append(mov['IMPORTE_KEY'])
                # Obtener el owner del row original de contabilidad si existe
                if 'NombreCuentaContable' in mov:
                    owners.add(str(mov.get('NombreCuentaContable', '')))
            
            resultado['CROSSMATCH_UNIDADES_D'] = ', '.join([str(u) for u in unidades if u])
            resultado['CROSSMATCH_VIAJES_D'] = ', '.join([str(v) for v in viajes if v])
            resultado['CROSSMATCH_IMPORTES_D'] = ', '.join([str(round(i, ndigits)) for i in importes if not pd.isna(i)])
            resultado['CROSSMATCH_OWNERS_D'] = ', '.join([str(o) for o in owners if o])
            
            # Verificar match de importe
            for imp in importes:
                if not pd.isna(imp) and not pd.isna(importe_base):
                    if abs(imp - importe_base) < 0.01:  # Match exacto
                        resultado['CROSSMATCH_IMPORTE_EXACTO_D'] = True
                        break
                    elif imp > 0 and abs(imp - importe_base) / imp < 0.01:  # Match similar (< 1% diff)
                        resultado['CROSSMATCH_IMPORTE_SIMILAR_D'] = True
        
        # Buscar en movimientos H
        if poliza_key in cont_h_por_poliza:
            resultado['CROSSMATCH_POLIZA_EN_H'] = True
            movs_h = cont_h_por_poliza[poliza_key]
            
            unidades_h = set()
            viajes_h = set()
            importes_h = []
            
            for mov in movs_h:
                unidades_h.add(mov['UNIDAD_KEY'])
                viajes_h.add(mov['VIAJE_KEY'])
                importes_h.append(mov['IMPORTE_KEY'])
            
            resultado['CROSSMATCH_UNIDADES_H'] = ', '.join([str(u) for u in unidades_h if u])
            resultado['CROSSMATCH_VIAJES_H'] = ', '.join([str(v) for v in viajes_h if v])
            
            # Verificar match de importe en H
            for imp in importes_h:
                if not pd.isna(imp) and not pd.isna(importe_base):
                    if abs(imp - importe_base) < 0.01:
                        resultado['CROSSMATCH_IMPORTE_EXACTO_H'] = True
                        break
        
        # Clasificar tipo de cross-match
        if resultado['CROSSMATCH_IMPORTE_EXACTO_D']:
            resultado['CROSSMATCH_TIPO'] = 'TRAFICO_DIFERENTE_IMPORTE_EXACTO'
            resultado['CROSSMATCH_DIAGNOSTICO'] = 'Mismo importe en Cont D pero diferente unidad/viaje'
        elif resultado['CROSSMATCH_IMPORTE_SIMILAR_D']:
            resultado['CROSSMATCH_TIPO'] = 'TRAFICO_DIFERENTE_IMPORTE_SIMILAR'
            resultado['CROSSMATCH_DIAGNOSTICO'] = 'Importe similar (~1%) en Cont D pero diferente unidad/viaje'
        elif resultado['CROSSMATCH_IMPORTE_EXACTO_H']:
            resultado['CROSSMATCH_TIPO'] = 'COSTO_EN_MOVIMIENTO_H'
            resultado['CROSSMATCH_DIAGNOSTICO'] = 'El costo está en movimiento H (abono) en vez de D (cargo)'
        elif resultado['CROSSMATCH_POLIZA_EN_D']:
            resultado['CROSSMATCH_TIPO'] = 'POLIZA_SIN_MATCH_IMPORTE'
            resultado['CROSSMATCH_DIAGNOSTICO'] = 'Póliza existe en Cont D pero sin match de importe'
        elif resultado['CROSSMATCH_POLIZA_EN_H']:
            resultado['CROSSMATCH_TIPO'] = 'POLIZA_SOLO_EN_H'
            resultado['CROSSMATCH_DIAGNOSTICO'] = 'Póliza solo tiene movimientos H'
        
        resultados.append(resultado)
    
    return pd.DataFrame(resultados)


def prep_contabilidad_completa(cont_raw: pd.DataFrame, ndigits: int, concept_map: dict[str, str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Prepara contabilidad separando movimientos D y H"""
    c_mov = resolve_col(cont_raw, ["TipoMovimiento", "Movimiento", "Tipo Movimiento"])
    c_importe = choose_cont_import_col(cont_raw)
    c_unidad = resolve_col(cont_raw, ["Unidad", "Numero de Unidad", "Numero_Unidad"])
    c_referencia = resolve_col(cont_raw, ["Referencia", "Numero_Viaje", "Numero Viaje", "Viaje"], required=False)
    c_poliza = resolve_col(cont_raw, ["Clave Poliza", "Clave Póliza", "ClavePoliza", "Factura", "Contrarrecibo"])
    c_concepto = resolve_col(cont_raw, ["Concepto detalle", "Concepto Detalle", "Concepto", "NombreCuentaContable"], required=False)
    c_vale = resolve_col(cont_raw, ["Vale", "No Vale", "Numero Vale"], required=False)
    c_owner = resolve_col(cont_raw, ["NombreCuentaContable", "Nombre Cuenta", "Owner"], required=False)

    out = cont_raw.copy()
    out["TIPO_MOV"] = out[c_mov].apply(norm_text)
    out["POLIZA_KEY"] = out[c_poliza].apply(norm_for_key)
    out["UNIDAD_KEY"] = out[c_unidad].apply(norm_for_key)
    out["VIAJE_KEY"] = out[c_referencia].apply(norm_for_key) if c_referencia else ""
    out["VALE_KEY"] = out[c_vale].apply(norm_for_key) if c_vale else ""
    out["CONCEPTO_KEY"] = out[c_concepto].apply(lambda x: canonical_concept(x, concept_map)) if c_concepto else ""
    out["IMPORTE_KEY"] = out[c_importe].apply(lambda x: norm_amount(x, ndigits))
    out["ROW_ID_CONT"] = range(1, len(out) + 1)
    
    if c_owner:
        out["NombreCuentaContable"] = out[c_owner]
    
    cont_d = out[out["TIPO_MOV"] == "D"].copy()
    cont_h = out[out["TIPO_MOV"] == "H"].copy()
    
    return cont_d, cont_h


# Importar las funciones originales necesarias
# (Se mantienen las funciones prep_base_saldos, match_base_vs_cont_mayoria, etc. del script original)
# Por brevedad, aquí solo muestro las nuevas funcionalidades

st.title("Saldos Owner - Desarrollo de Costos v2.2")
st.caption("🆕 Nueva funcionalidad: Detección automática de cross-match por póliza (costo en tráfico diferente)")

with st.expander("Novedades en v2.2", expanded=True):
    st.markdown("""
    **Nueva funcionalidad: Cross-Match por Póliza**
    
    Ahora el sistema detecta automáticamente casos donde:
    - El costo está en un **tráfico diferente** (misma póliza, diferente unidad/viaje)
    - El importe coincide de forma **exacta** o **similar** (< 1% de diferencia)
    - El costo está en movimiento **H** (abono) en vez de D (cargo)
    
    Estos casos se clasifican en:
    - `TRAFICO_DIFERENTE_IMPORTE_EXACTO`: Mismo importe, diferente tráfico
    - `TRAFICO_DIFERENTE_IMPORTE_SIMILAR`: Importe similar (~1%), diferente tráfico
    - `COSTO_EN_MOVIMIENTO_H`: El costo está en H en vez de D
    - `POLIZA_SIN_MATCH_IMPORTE`: Póliza existe pero sin match de importe
    - `POLIZA_SOLO_EN_H`: Póliza solo tiene movimientos H
    
    El reporte incluye columnas adicionales con detalles del cross-match:
    - Unidades/viajes donde está registrado en Contabilidad
    - Importes encontrados
    - Owners asociados
    - Diagnóstico detallado
    """)

# El resto del código de la UI se mantiene igual que en v2.1,
# pero agregando el análisis de cross-match para los casos NO_EXISTE

st.info("⚠️ NOTA: Este script es una versión conceptual que muestra las mejoras propuestas.")
st.info("Para implementación completa, se debe integrar con el código completo de v2.1")
