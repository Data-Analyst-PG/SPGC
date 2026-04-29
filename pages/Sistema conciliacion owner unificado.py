"""
Sistema Unificado de Conciliación de Saldos Owner
Version: 3.3 - CORREGIDA para replicar lógica exacta del script anterior

CAMBIOS CRÍTICOS vs v3.2:
1. Usa las columnas CORRECTAS: Factura (no ClavePoliza) en Contabilidad
2. Replica lógica de merge exacto del script anterior
3. Matching por 5 campos: PR, VIAJE, UNIDAD, TIPO_PAGO, IMPORTE
"""

import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# ============================================================
# CONFIGURACIÓN
# ============================================================

@dataclass
class ConfigConciliacion:
    """Configuración centralizada del sistema"""
    ndigits: int = 2
    umbral_match_ok: int = 5
    umbral_match_con_discrepancia: int = 3


# ============================================================
# NORMALIZACIÓN
# ============================================================

class Normalizador:
    """Normalización compatible con script anterior"""
    
    @staticmethod
    def texto(x: Any) -> str:
        """Normaliza texto: mayúsculas, espacios limpios"""
        if x is None or pd.isna(x):
            return ""
        s = str(x).strip().upper()
        s = re.sub(r"\s+", " ", s)
        return s
    
    @staticmethod
    def monto(x: Any, ndigits: int = 2) -> float:
        """Normaliza montos numéricos"""
        try:
            if x is None or pd.isna(x):
                return float("nan")
            if isinstance(x, str):
                x = x.replace(",", "").replace("$", "").strip()
            return round(float(x), ndigits)
        except Exception:
            return float("nan")


# ============================================================
# UTILIDADES
# ============================================================

def build_seq(df: pd.DataFrame, key_cols: List[str], seq_col: str = "_seq") -> pd.DataFrame:
    """Agrega columna de secuencia por grupo (para duplicados)"""
    out = df.copy()
    out[seq_col] = out.groupby(key_cols, dropna=False).cumcount() + 1
    return out


class ManejadorArchivos:
    """Maneja lectura y escritura de archivos"""
    
    @staticmethod
    def leer_tabla(file_obj, sheet_name: Optional[str] = None, usecols: Optional[List[str]] = None) -> pd.DataFrame:
        """Lee archivo Excel o CSV"""
        suffix = Path(file_obj.name).suffix.lower()
        raw = file_obj.getvalue()
        
        if suffix == '.csv':
            return pd.read_csv(BytesIO(raw), usecols=usecols, low_memory=False)
        elif suffix in {'.xlsx', '.xlsm', '.xls'}:
            if sheet_name:
                try:
                    return pd.read_excel(BytesIO(raw), sheet_name=sheet_name, usecols=usecols)
                except Exception:
                    return pd.read_excel(BytesIO(raw), usecols=usecols)
            return pd.read_excel(BytesIO(raw), usecols=usecols)
        else:
            raise ValueError(f"Formato no soportado: {suffix}")
    
    @staticmethod
    def exportar_excel(sheets: Dict[str, pd.DataFrame]) -> bytes:
        """Exporta múltiples hojas a Excel"""
        bio = BytesIO()
        with pd.ExcelWriter(bio, engine='xlsxwriter', engine_kwargs={'options': {'constant_memory': True}}) as writer:
            for nombre, df in sheets.items():
                df_out = df.copy()
                df_out.columns = [str(c)[:250] for c in df_out.columns]
                df_out.to_excel(writer, sheet_name=nombre[:31], index=False)
        bio.seek(0)
        return bio.getvalue()


# ============================================================
# MOTOR DE MATCHING (Lógica del script anterior)
# ============================================================

class MotorMatchingLiquidaciones:
    """Motor que replica EXACTAMENTE la lógica del script anterior"""
    
    def __init__(self, config: ConfigConciliacion):
        self.config = config
        self.norm = Normalizador()
    
    def match_liquidaciones_vs_contabilidad(self,
                                           liq: pd.DataFrame,
                                           cont: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Matching EXACTO como el script anterior:
        - Merge por 5 campos: PR, VIAJE, UNIDAD, TIPO_PAGO, IMPORTE
        - Con secuencia para manejar duplicados
        """
        
        st.info(f"🔍 Matching Liquidaciones ({len(liq):,}) vs Contabilidad ({len(cont):,})...")
        
        # Columnas clave para el merge (IGUAL que script anterior)
        key_cols = ["PR", "VIAJE", "UNIDAD", "TIPO_PAGO", "IMPORTE"]
        merge_keys = key_cols + ["_seq"]
        
        # Agregar secuencia (para manejar duplicados)
        st.write("📊 Agregando secuencias para duplicados...")
        liq_k = build_seq(liq, key_cols)
        cont_k = build_seq(cont, key_cols)
        
        # Merge exacto
        st.write("🔄 Ejecutando merge exacto...")
        m = liq_k.merge(
            cont_k,
            how="outer",
            on=merge_keys,
            suffixes=("_LIQ", "_CONT"),
            indicator=True,
        )
        
        # Clasificar resultados
        matched = m[m["_merge"] == "both"].copy()
        only_liq = m[m["_merge"] == "left_only"].copy()
        only_cont = m[m["_merge"] == "right_only"].copy()
        
        st.success(f"✅ MATCH: {len(matched):,} | Solo Liq: {len(only_liq):,} | Solo Cont: {len(only_cont):,}")
        
        # Clasificar matched por owner
        if "OWNER_LIQ" in matched.columns and "OWNER_CONT" in matched.columns:
            matched["OWNER_MATCH"] = (
                matched["OWNER_LIQ"].astype("string").fillna("") ==
                matched["OWNER_CONT"].astype("string").fillna("")
            )
            matched["ESTATUS_MATCH"] = matched["OWNER_MATCH"].map({
                True: "MATCH_OK", 
                False: "MATCH_CON_DISCREPANCIA"
            })
            matched["OBSERVACION"] = matched["OWNER_MATCH"].map({
                True: "Coincide llave exacta y owner.",
                False: "Coincide llave exacta, pero owner es distinto.",
            })
        else:
            matched["ESTATUS_MATCH"] = "MATCH_OK"
            matched["OBSERVACION"] = "Coincide llave exacta."
        
        only_liq["ESTATUS_MATCH"] = "NO_EXISTE_EN_CONTABILIDAD"
        only_liq["OBSERVACION"] = "La fila de Liquidaciones no encontró contraparte exacta en Contabilidad."
        
        only_cont["ESTATUS_MATCH"] = "NO_EXISTE_EN_LIQUIDACIONES"
        only_cont["OBSERVACION"] = "La fila de Contabilidad no encontró contraparte exacta en Liquidaciones."
        
        # Reconstruir clasificados
        liq_clasificado = self._reconstruir_liq_clasificado(liq, matched, only_liq)
        cont_clasificado = self._reconstruir_cont_clasificado(cont, matched, only_cont)
        
        return liq_clasificado, cont_clasificado
    
    def _reconstruir_liq_clasificado(self, liq_original: pd.DataFrame, matched: pd.DataFrame, only_liq: pd.DataFrame) -> pd.DataFrame:
        """Reconstruye liquidaciones clasificadas"""
        
        # De matched
        liq_status_from_matched = matched[[
            "ROW_ID_LIQ", "ROW_ID_CONT", "ESTATUS_MATCH", "OBSERVACION"
        ]].copy()
        
        if "OWNER_CONT" in matched.columns:
            liq_status_from_matched["OWNER_CONT"] = matched["OWNER_CONT"]
        
        # De only_liq
        liq_status_from_only = only_liq[["ROW_ID_LIQ", "ESTATUS_MATCH", "OBSERVACION"]].copy()
        liq_status_from_only["ROW_ID_CONT"] = pd.NA
        if "OWNER_CONT" in matched.columns:
            liq_status_from_only["OWNER_CONT"] = ""
        
        # Combinar
        liq_status = pd.concat([liq_status_from_matched, liq_status_from_only], ignore_index=True)
        liq_clasificado = liq_original.merge(liq_status, on="ROW_ID_LIQ", how="left")
        
        return liq_clasificado
    
    def _reconstruir_cont_clasificado(self, cont_original: pd.DataFrame, matched: pd.DataFrame, only_cont: pd.DataFrame) -> pd.DataFrame:
        """Reconstruye contabilidad clasificada"""
        
        # De matched
        cont_status_from_matched = matched[[
            "ROW_ID_CONT", "ROW_ID_LIQ", "ESTATUS_MATCH", "OBSERVACION"
        ]].copy()
        
        if "OWNER_LIQ" in matched.columns:
            cont_status_from_matched["OWNER_LIQ"] = matched["OWNER_LIQ"]
        
        # De only_cont
        cont_status_from_only = only_cont[["ROW_ID_CONT", "ESTATUS_MATCH", "OBSERVACION"]].copy()
        cont_status_from_only["ROW_ID_LIQ"] = pd.NA
        if "OWNER_LIQ" in matched.columns:
            cont_status_from_only["OWNER_LIQ"] = ""
        
        # Combinar
        cont_status = pd.concat([cont_status_from_matched, cont_status_from_only], ignore_index=True)
        cont_clasificado = cont_original.merge(cont_status, on="ROW_ID_CONT", how="left")
        
        return cont_clasificado


# ============================================================
# PREPARADORES ESPECÍFICOS
# ============================================================

class PreparadorLiquidaciones:
    """Prepara liquidaciones EXACTAMENTE como script anterior"""
    
    def __init__(self, config: ConfigConciliacion):
        self.config = config
        self.norm = Normalizador()
    
    def preparar(self, df: pd.DataFrame, tipo_concepto: str = 'E') -> pd.DataFrame:
        """Prepara liquidaciones"""
        
        # Renombrar columnas (IGUAL que script anterior)
        out = df.rename(columns={
            "Liquidacion": "PR",
            "Numero_Viaje": "VIAJE",
            "TipoPago": "TIPO_PAGO",
            "Monto": "IMPORTE",
            "Unidad": "UNIDAD",
            "Owner": "OWNER_LIQ",
            "Tipo_Concepto": "TIPO_CONCEPTO",
        })
        
        # Normalizar campos de texto
        for c in ["PR", "VIAJE", "TIPO_PAGO", "UNIDAD", "OWNER_LIQ", "TIPO_CONCEPTO"]:
            if c in out.columns:
                out[c] = out[c].apply(self.norm.texto)
        
        # Normalizar importe
        out["IMPORTE"] = pd.to_numeric(
            out["IMPORTE"].apply(lambda x: self.norm.monto(x, self.config.ndigits)), 
            errors="coerce"
        )
        
        # Filtrar por tipo de concepto
        out = out[out["TIPO_CONCEPTO"] == tipo_concepto].copy()
        
        # Resetear índice y agregar ROW_ID
        out = out.reset_index(drop=True)
        out["ROW_ID_LIQ"] = out.index + 1
        
        return out


class PreparadorContabilidad:
    """Prepara contabilidad EXACTAMENTE como script anterior"""
    
    def __init__(self, config: ConfigConciliacion):
        self.config = config
        self.norm = Normalizador()
    
    def preparar(self, df: pd.DataFrame, tipo_mov: str = 'H') -> pd.DataFrame:
        """Prepara contabilidad"""
        
        # Renombrar columnas (CLAVE: Usar "Factura", NO "ClavePoliza")
        out = df.rename(columns={
            "Factura": "PR",  # ← ESTO ES CRÍTICO
            "Referencia": "VIAJE",
            "TipoPago": "TIPO_PAGO",
            "Importe": "IMPORTE",
            "Unidad": "UNIDAD",
            "NombreCuentaContable": "OWNER_CONT",
            "TipoMovimiento": "TIPO_MOV",
        })
        
        # Normalizar campos de texto
        for c in ["PR", "VIAJE", "TIPO_PAGO", "UNIDAD", "OWNER_CONT", "TIPO_MOV"]:
            if c in out.columns:
                out[c] = out[c].apply(self.norm.texto)
        
        # Normalizar importe
        out["IMPORTE"] = pd.to_numeric(
            out["IMPORTE"].apply(lambda x: self.norm.monto(x, self.config.ndigits)), 
            errors="coerce"
        )
        
        # Filtrar por tipo de movimiento
        out = out[out["TIPO_MOV"] == tipo_mov].copy()
        
        # Resetear índice y agregar ROW_ID
        out = out.reset_index(drop=True)
        out["ROW_ID_CONT"] = out.index + 1
        
        return out


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

def main():
    st.set_page_config(page_title="Sistema de Conciliación Owner v3.3", layout="wide")
    
    st.title("🔄 Sistema Unificado de Conciliación de Saldos Owner")
    st.caption("Versión 3.3 - CORREGIDA - Replica lógica exacta del script anterior")
    
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        config = ConfigConciliacion(
            ndigits=st.number_input("Decimales para importes", 0, 4, 2),
        )
        
        st.divider()
        st.header("📁 Archivos de Entrada")
        
        cont_file = st.file_uploader("📊 Contabilidad (obligatorio)", type=['xlsx', 'xls', 'csv'])
        liq_file = st.file_uploader("📈 Liquidaciones", type=['xlsx', 'xls', 'csv'])
        
        st.divider()
        
        tipo_concepto = st.selectbox("Liquidaciones: Tipo_Concepto", ["E", "I"], index=0)
        tipo_mov = st.selectbox("Contabilidad: TipoMovimiento", ["H", "D"], index=0)
        
        st.divider()
        
        procesar_btn = st.button("▶️ PROCESAR", type="primary", use_container_width=True)
    
    with st.expander("ℹ️ Versión 3.3 - CAMBIOS", expanded=True):
        st.markdown("""
        ### 🔧 Correcciones Aplicadas
        
        **Problema identificado**:
        - ❌ v3.2 buscaba columna `ClavePoliza` en Contabilidad
        - ✅ La columna correcta es `Factura`
        
        **Solución**:
        - ✅ Usa las MISMAS columnas que el script anterior
        - ✅ Replica la lógica de merge exacto por 5 campos
        - ✅ Matching idéntico al script que funciona
        
        **Columnas usadas**:
        - Liquidaciones: `Liquidacion, Numero_Viaje, TipoPago, Monto, Unidad`
        - Contabilidad: `Factura, Referencia, TipoPago, Importe, Unidad`
        """)
    
    if not procesar_btn:
        st.info("👆 Carga los archivos y da clic en **PROCESAR** para iniciar.")
        return
    
    if cont_file is None:
        st.error("❌ Debes cargar el archivo de Contabilidad.")
        return
    
    if liq_file is None:
        st.warning("⚠️ No cargaste Liquidaciones. Solo se procesará Contabilidad.")
    
    inicio = datetime.now()
    
    try:
        archivos = ManejadorArchivos()
        motor = MotorMatchingLiquidaciones(config)
        prep_liq = PreparadorLiquidaciones(config)
        prep_cont = PreparadorContabilidad(config)
        
        resultados = {}
        
        # PASO 1: Contabilidad
        st.header("📊 1. Procesando Contabilidad")
        
        # Leer con columnas específicas (IGUAL que script anterior)
        cont_usecols = ["Factura", "Referencia", "TipoPago", "Importe", "Unidad", "NombreCuentaContable", "TipoMovimiento"]
        
        cont_raw = archivos.leer_tabla(
            cont_file, 
            sheet_name='ContabilidadSET_PLUS_datos',
            usecols=cont_usecols
        )
        
        st.info(f"Contabilidad cargada: {len(cont_raw):,} registros")
        
        cont_h = prep_cont.preparar(cont_raw, tipo_mov='H')
        cont_d = prep_cont.preparar(cont_raw, tipo_mov='D')
        
        col1, col2 = st.columns(2)
        col1.metric("Movimientos H", f"{len(cont_h):,}")
        col2.metric("Movimientos D", f"{len(cont_d):,}")
        
        resultados['Contabilidad_H'] = cont_h
        resultados['Contabilidad_D'] = cont_d
        
        # PASO 2: LIQUIDACIONES (si se cargó)
        if liq_file:
            st.header("📈 2. Procesando Liquidaciones (Ingresos)")
            
            # Leer con columnas específicas
            liq_usecols = ["Liquidacion", "Numero_Viaje", "TipoPago", "Monto", "Unidad", "Owner", "Tipo_Concepto"]
            
            liq_raw = archivos.leer_tabla(
                liq_file,
                sheet_name='LiquidacionesSET_PLUS_datos',
                usecols=liq_usecols
            )
            
            st.info(f"Liquidaciones cargadas: {len(liq_raw):,} registros")
            
            liq_filtrado = prep_liq.preparar(liq_raw, tipo_concepto=tipo_concepto)
            
            st.info(f"Liquidaciones filtradas (Tipo={tipo_concepto}): {len(liq_filtrado):,} registros")
            
            # MATCHING
            liq_clasificado, cont_clasificado = motor.match_liquidaciones_vs_contabilidad(
                liq_filtrado, 
                cont_h
            )
            
            # Métricas
            col1, col2, col3, col4 = st.columns(4)
            
            match_ok = (liq_clasificado['ESTATUS_MATCH'] == 'MATCH_OK').sum()
            match_disc = (liq_clasificado['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum()
            no_match = (liq_clasificado['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD').sum()
            
            col1.metric("Total Liquidaciones", f"{len(liq_clasificado):,}")
            col2.metric("✅ MATCH_OK", f"{match_ok:,}")
            col3.metric("⚠️ DISCREPANCIA", f"{match_disc:,}")
            col4.metric("❌ NO MATCH", f"{no_match:,}")
            
            # Comparación con script anterior
            st.divider()
            st.subheader("📊 Comparación con Script Anterior")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Script Anterior** (9___Saldos_Owner__5_.py)")
                st.write(f"- MATCH_OK: 74,705")
                st.write(f"- MATCH_CON_DISCREPANCIA: 62,748")
                st.write(f"- NO_EXISTE: 221,692")
                st.write(f"- **Total: 359,145**")
            
            with col2:
                st.markdown("**Este Script** (v3.3)")
                st.write(f"- MATCH_OK: {match_ok:,}")
                st.write(f"- MATCH_CON_DISCREPANCIA: {match_disc:,}")
                st.write(f"- NO_EXISTE: {no_match:,}")
                st.write(f"- **Total: {len(liq_clasificado):,}**")
            
            # Validación
            if abs(len(liq_clasificado) - 359145) < 10:
                if abs(match_ok - 74705) < 100:
                    st.success("✅ Los resultados coinciden con el script anterior!")
                else:
                    st.warning(f"⚠️ Hay diferencia en MATCH_OK: {abs(match_ok - 74705):,} registros")
            else:
                st.error(f"❌ Total de registros no coincide. Diferencia: {abs(len(liq_clasificado) - 359145):,}")
            
            resultados['Liquidaciones_Filtradas'] = liq_filtrado
            resultados['Liquidaciones_Clasificadas'] = liq_clasificado
            resultados['Contabilidad_H_Clasificada'] = cont_clasificado
        
        # EXPORTACIÓN
        st.divider()
        
        tiempo_total = (datetime.now() - inicio).total_seconds()
        st.success(f"✅ Procesamiento completado en {tiempo_total:.1f} segundos")
        
        if resultados:
            excel_bytes = archivos.exportar_excel(resultados)
            
            st.download_button(
                label="📥 Descargar Resultado Completo (Excel)",
                data=excel_bytes,
                file_name=f"conciliacion_owner_v33_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            with st.expander("📋 Hojas incluidas en el archivo"):
                for nombre in resultados.keys():
                    st.write(f"- {nombre} ({len(resultados[nombre]):,} filas)")
    
    except Exception as e:
        st.error(f"❌ Error durante el procesamiento: {str(e)}")
        st.exception(e)


if __name__ == "__main__":
    main()
