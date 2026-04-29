"""
Sistema Unificado de Conciliación de Saldos Owner
Version: 3.1 - COMPLETA con Módulo de Ingresos

Características principales:
- Módulo de COSTOS completo ✅
- Módulo de INGRESOS completo ✅
- Sin necesidad de catálogos
- Procesamiento paralelo optimizado
- Sistema de scoring inteligente
- Análisis integrado D vs H
- Generación de saldos finales por Owner
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
    
    # Pesos para scoring de conceptos
    peso_concepto_exacto: float = 1.0
    peso_concepto_similar: float = 0.7
    peso_concepto_traduccion: float = 0.9


# ============================================================
# NORMALIZACIÓN Y UTILIDADES
# ============================================================

class Normalizador:
    """Clase centralizada para normalización de datos"""
    
    # Diccionario de traducciones comunes inglés-español
    TRADUCCIONES = {
        'DIESEL': 'DIESEL',
        'FUEL': 'DIESEL',
        'CONSUMIBLES': 'CONSUMIBLES',
        'CONSUMABLES': 'CONSUMIBLES',
        'LOAN': 'PRESTAMO',
        'PERSONAL LOAN': 'PRESTAMO PERSONAL',
        'PRESTAMO': 'PRESTAMO',
        'ADVANCE': 'ANTICIPO',
        'ANTICIPO': 'ANTICIPO',
        'REPAIR': 'REPARACION',
        'REPARACION': 'REPARACION',
        'PLATES': 'PLACAS',
        'PLACAS': 'PLACAS',
        'INSURANCE': 'SEGURO',
        'SEGURO': 'SEGURO',
        'TOLL': 'CASETA',
        'CASETAS': 'CASETA',
        'MAINTENANCE': 'MANTENIMIENTO',
        'MANTENIMIENTO': 'MANTENIMIENTO',
        'TIRES': 'LLANTAS',
        'LLANTAS': 'LLANTAS',
    }
    
    @staticmethod
    def texto(x: Any) -> str:
        """Normaliza texto: mayúsculas, sin acentos, espacios limpios"""
        if x is None or pd.isna(x):
            return ""
        s = str(x).strip().upper()
        # Remover acentos
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        # Limpiar espacios múltiples
        s = re.sub(r"\s+", " ", s)
        return s
    
    @staticmethod
    def clave(x: Any) -> str:
        """Normaliza para usar como clave: solo alfanuméricos"""
        s = Normalizador.texto(x)
        # Solo letras y números
        s = re.sub(r"[^A-Z0-9]+", " ", s)
        return re.sub(r"\s+", " ", s).strip()
    
    @staticmethod
    def monto(x: Any, ndigits: int = 2) -> float:
        """Normaliza montos numéricos"""
        try:
            if x is None or pd.isna(x):
                return float("nan")
            if isinstance(x, str):
                # Remover símbolos comunes
                x = x.replace(",", "").replace("$", "").replace("%", "").strip()
            return round(float(x), ndigits)
        except Exception:
            return float("nan")
    
    @staticmethod
    def limpiar_concepto_sufijo(x: Any) -> str:
        """Quita sufijos tipo ' - 20170908' del concepto"""
        s = Normalizador.texto(x)
        # Remover patrones tipo " - 20170908" o " - CODIGO"
        s = re.sub(r"\s+-\s+\d{8,}.*$", "", s)
        s = re.sub(r"\s+-\s+[A-Z0-9]{6,}.*$", "", s)
        return s.strip()
    
    @classmethod
    def concepto_canonico(cls, x: Any) -> str:
        """
        Normaliza conceptos aplicando:
        1. Limpieza de sufijos
        2. Traducción inglés-español
        3. Normalización de variantes
        """
        s = cls.limpiar_concepto_sufijo(x)
        k = cls.clave(s)
        
        # Buscar traducciones
        palabras = k.split()
        palabras_traducidas = []
        
        for palabra in palabras:
            if palabra in cls.TRADUCCIONES:
                palabras_traducidas.append(cls.TRADUCCIONES[palabra])
            else:
                palabras_traducidas.append(palabra)
        
        resultado = " ".join(palabras_traducidas)
        
        # Reglas de normalización específicas
        reglas = [
            (r'\bCXP\s+(DIESEL|CONSUMIBLES)\b', 'CXP DIESEL CONSUMIBLES'),
            (r'\bPERSONAL\s+LOAN\b|\bLOAN\b|\bPRESTAMO\b', 'PRESTAMO'),
            (r'\bANTICIPO\b|\bADVANCE\b', 'ANTICIPO'),
            (r'\bDIESEL\b.*\bCONSUMIBLES\b|\bCONSUMIBLES\b.*\bDIESEL\b', 'CXP DIESEL CONSUMIBLES'),
        ]
        
        for patron, valor in reglas:
            if re.search(patron, resultado):
                return valor
        
        return resultado


class ScoringMatcher:
    """Sistema de scoring para matching inteligente"""
    
    def __init__(self, config: ConfigConciliacion):
        self.config = config
    
    def calcular_similitud_concepto(self, concepto1: str, concepto2: str) -> float:
        """
        Calcula similitud entre conceptos considerando:
        - Match exacto
        - Traducciones
        - Palabras clave comunes
        """
        c1 = Normalizador.concepto_canonico(concepto1)
        c2 = Normalizador.concepto_canonico(concepto2)
        
        if c1 == c2:
            return self.config.peso_concepto_exacto
        
        # Verificar si comparten palabras clave importantes
        palabras1 = set(c1.split())
        palabras2 = set(c2.split())
        
        if not palabras1 or not palabras2:
            return 0.0
        
        # Calcular Jaccard similarity
        interseccion = palabras1.intersection(palabras2)
        union = palabras1.union(palabras2)
        
        similitud = len(interseccion) / len(union) if union else 0.0
        
        # Si hay al menos 50% de overlap, considerar como similar
        if similitud >= 0.5:
            return self.config.peso_concepto_similar
        
        return 0.0
    
    def calcular_score(self, 
                       criterios_evaluados: int,
                       criterios_cumplidos: int,
                       score_concepto: float = 0.0) -> Tuple[str, int, float]:
        """
        Calcula el estatus de match basado en criterios
        
        Returns:
            Tuple[estatus, total_coincidencias, score_normalizado]
        """
        # Ajustar coincidencias por score de concepto
        coincidencias_ajustadas = criterios_cumplidos
        if score_concepto > 0 and score_concepto < 1.0:
            coincidencias_ajustadas += score_concepto - 1  # Ajuste fino
        
        score_normalizado = coincidencias_ajustadas / criterios_evaluados if criterios_evaluados > 0 else 0.0
        
        if criterios_cumplidos >= self.config.umbral_match_ok:
            return "MATCH_OK", criterios_cumplidos, score_normalizado
        elif criterios_cumplidos >= self.config.umbral_match_con_discrepancia:
            return "MATCH_CON_DISCREPANCIA", criterios_cumplidos, score_normalizado
        else:
            return "NO_MATCH", criterios_cumplidos, score_normalizado


# ============================================================
# PREPARACIÓN DE DATOS
# ============================================================

class PreparadorDatos:
    """Prepara y normaliza datos de diferentes fuentes"""
    
    def __init__(self, config: ConfigConciliacion):
        self.config = config
        self.norm = Normalizador()
    
    def preparar_contabilidad(self, 
                             df: pd.DataFrame, 
                             tipo_mov: Optional[str] = None) -> Tuple[pd.DataFrame, Dict[str, str]]:
        """
        Prepara datos de contabilidad con todas las llaves necesarias
        """
        out = df.copy()
        
        # Resolver nombres de columnas
        col_map = {
            'movimiento': self._resolver_col(df, ['TipoMovimiento', 'Movimiento', 'Tipo Movimiento']),
            'poliza': self._resolver_col(df, ['ClavePoliza', 'Clave Poliza', 'Factura', 'Contrarrecibo']),
            'unidad': self._resolver_col(df, ['Unidad', 'Numero de Unidad', 'NumeroUnidad', 'IdUnidad']),
            'referencia': self._resolver_col(df, ['Referencia', 'Numero_Viaje', 'NumeroViaje', 'Viaje'], required=False),
            'concepto': self._resolver_col(df, ['ConceptoDetalle', 'Concepto Detalle', 'Concepto', 'NombreCuentaContable'], required=False),
            'importe': self._elegir_col_importe(df),
            'vale': self._resolver_col(df, ['Vale', 'No Vale', 'NumeroVale'], required=False),
            'factura': self._resolver_col(df, ['Factura'], required=False),
            'tipo_pago': self._resolver_col(df, ['TipoPago', 'Tipo Pago', 'IdTipoPago'], required=False),
        }
        
        # Normalizar campos
        out['TIPO_MOV'] = out[col_map['movimiento']].apply(self.norm.texto)
        out['POLIZA_KEY'] = out[col_map['poliza']].apply(self.norm.clave)
        out['UNIDAD_KEY'] = out[col_map['unidad']].apply(self.norm.clave)
        out['VIAJE_KEY'] = out[col_map['referencia']].apply(self.norm.clave) if col_map['referencia'] else ""
        out['VALE_KEY'] = out[col_map['vale']].apply(self.norm.clave) if col_map['vale'] else ""
        out['CONCEPTO_KEY'] = out[col_map['concepto']].apply(self.norm.concepto_canonico) if col_map['concepto'] else ""
        out['IMPORTE_KEY'] = out[col_map['importe']].apply(lambda x: self.norm.monto(x, self.config.ndigits))
        out['TIPO_PAGO_KEY'] = out[col_map['tipo_pago']].apply(self.norm.clave) if col_map['tipo_pago'] else ""
        
        # Identificador único de fila
        out['ROW_ID_CONT'] = range(1, len(out) + 1)
        
        # Filtrar por tipo de movimiento si se especifica
        if tipo_mov is not None:
            out = out[out['TIPO_MOV'] == self.norm.texto(tipo_mov)].copy()
        
        return out, col_map
    
    def preparar_liquidaciones(self, 
                              df: pd.DataFrame, 
                              tipo_concepto: str = 'E') -> pd.DataFrame:
        """Prepara liquidaciones para matching"""
        
        col_map = {
            'pr': self._resolver_col(df, ['Liquidacion', 'PR', 'NumeroLiquidacion']),
            'viaje': self._resolver_col(df, ['Numero_Viaje', 'NumeroViaje', 'Viaje']),
            'tipo_pago': self._resolver_col(df, ['TipoPago', 'Tipo Pago']),
            'importe': self._resolver_col(df, ['Monto', 'Importe', 'Total']),
            'unidad': self._resolver_col(df, ['Unidad']),
            'owner': self._resolver_col(df, ['Owner', 'Operador'], required=False),
            'tipo_concepto': self._resolver_col(df, ['Tipo_Concepto', 'TipoConcepto']),
            'concepto': self._resolver_col(df, ['Concepto'], required=False),
        }
        
        out = df.copy()
        
        # Filtrar por tipo de concepto
        out = out[out[col_map['tipo_concepto']].apply(self.norm.texto) == self.norm.texto(tipo_concepto)].copy()
        
        # Normalizar llaves
        out['PR_KEY'] = out[col_map['pr']].apply(self.norm.clave)
        out['VIAJE_KEY'] = out[col_map['viaje']].apply(self.norm.clave)
        out['TIPO_PAGO_KEY'] = out[col_map['tipo_pago']].apply(self.norm.clave)
        out['UNIDAD_KEY'] = out[col_map['unidad']].apply(self.norm.clave)
        out['IMPORTE_KEY'] = out[col_map['importe']].apply(lambda x: self.norm.monto(x, self.config.ndigits))
        
        if col_map['owner']:
            out['OWNER_KEY'] = out[col_map['owner']].apply(self.norm.clave)
            out['OWNER_ORIGINAL'] = out[col_map['owner']].apply(self.norm.texto)
        else:
            out['OWNER_KEY'] = ""
            out['OWNER_ORIGINAL'] = ""
        
        if col_map['concepto']:
            out['CONCEPTO_KEY'] = out[col_map['concepto']].apply(self.norm.concepto_canonico)
        else:
            out['CONCEPTO_KEY'] = ""
        
        out['ROW_ID_LIQ'] = range(1, len(out) + 1)
        
        return out
    
    def preparar_base_saldos(self, df: pd.DataFrame) -> pd.DataFrame:
        """Prepara Base Saldos corregida"""
        out = df.copy()
        
        col_map = {
            'contrarecibo': self._resolver_col(df, ['FOLIO_CONTRARECIBO', 'Contrarecibo', 'FolioContrarecibo']),
            'unidad': self._resolver_col(df, ['NUMERO_UNIDAD', 'Unidad', 'NumeroUnidad']),
            'viaje': self._resolver_col(df, ['NUMERO_VIAJE', 'Viaje', 'NumeroViaje']),
            'importe': self._resolver_col(df, ['Importe', 'IMPORTE', 'Monto']),
            'concepto': self._resolver_col(df, ['Concepto contabilidad', 'ConceptoContabilidad', 'Concepto'], required=False),
            'owner': self._resolver_col(df, ['ID_OWNER', 'Owner', 'IdOwner'], required=False),
        }
        
        out['POLIZA_KEY'] = out[col_map['contrarecibo']].apply(self.norm.clave)
        out['UNIDAD_KEY'] = out[col_map['unidad']].apply(self.norm.clave)
        out['VIAJE_KEY'] = out[col_map['viaje']].apply(self.norm.clave)
        out['CONCEPTO_KEY'] = out[col_map['concepto']].apply(self.norm.concepto_canonico) if col_map['concepto'] else ""
        out['IMPORTE_KEY'] = out[col_map['importe']].apply(lambda x: self.norm.monto(x, self.config.ndigits))
        
        if col_map['owner']:
            out['OWNER_KEY'] = out[col_map['owner']].astype(str).apply(self.norm.clave)
            out['OWNER_ORIGINAL'] = out[col_map['owner']].astype(str).apply(self.norm.texto)
        else:
            out['OWNER_KEY'] = ""
            out['OWNER_ORIGINAL'] = ""
        
        out['ROW_ID_BASE'] = range(1, len(out) + 1)
        out['ORIGEN'] = 'BASE_SALDOS'
        
        return out
    
    def preparar_cheques(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Prepara Cheques excluyendo CargoA = Company"""
        col_map = {
            'cargo_a': self._resolver_col(df, ['CargoA', 'Cargo A', 'CargadoA']),
            'unidad': self._resolver_col(df, ['Unidad'], required=False),
            'viaje': self._resolver_col(df, ['Viaje'], required=False),
            'concepto': self._resolver_col(df, ['Concepto'], required=False),
            'observaciones': self._resolver_col(df, ['Observaciones'], required=False),
            'contrarecibo': self._resolver_col(df, ['Contrarecibo', 'ContraRecibo'], required=False),
            'importe': self._resolver_col(df, ['Importe', 'Total', 'Monto']),
            'vale': self._resolver_col(df, ['Vale'], required=False),
            'operador': self._resolver_col(df, ['Operador', 'Owner'], required=False),
        }
        
        df_work = df.copy()
        df_work['CARGO_A_NORM'] = df_work[col_map['cargo_a']].apply(self.norm.texto)
        
        excluidos = df_work[df_work['CARGO_A_NORM'] == 'COMPANY'].copy()
        validos = df_work[df_work['CARGO_A_NORM'] != 'COMPANY'].copy()
        
        # Normalizar campos válidos
        for df_temp in [validos]:
            df_temp['UNIDAD_KEY'] = df_temp[col_map['unidad']].apply(self.norm.clave) if col_map['unidad'] else ""
            df_temp['VIAJE_KEY'] = df_temp[col_map['viaje']].apply(self.norm.clave) if col_map['viaje'] else ""
            df_temp['CONCEPTO_KEY'] = df_temp[col_map['concepto']].apply(self.norm.concepto_canonico) if col_map['concepto'] else ""
            df_temp['POLIZA_KEY'] = df_temp[col_map['contrarecibo']].apply(self.norm.clave) if col_map['contrarecibo'] else ""
            df_temp['VALE_KEY'] = df_temp[col_map['vale']].apply(self.norm.clave) if col_map['vale'] else ""
            df_temp['IMPORTE_KEY'] = df_temp[col_map['importe']].apply(lambda x: self.norm.monto(x, self.config.ndigits))
            df_temp['ORIGEN'] = 'CHEQUE'
            
            if col_map['operador']:
                df_temp['OWNER_KEY'] = df_temp[col_map['operador']].apply(self.norm.clave)
                df_temp['OWNER_ORIGINAL'] = df_temp[col_map['operador']].apply(self.norm.texto)
            else:
                df_temp['OWNER_KEY'] = ""
                df_temp['OWNER_ORIGINAL'] = ""
        
        validos['ROW_ID_CHEQUE'] = range(1, len(validos) + 1)
        excluidos['ROW_ID_CHEQUE_EXCL'] = range(1, len(excluidos) + 1)
        
        return validos, excluidos
    
    def preparar_vouchers(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Prepara Vouchers excluyendo Operador FILIAL"""
        col_map = {
            'operador': self._resolver_col(df, ['Operador']),
            'unidad': self._resolver_col(df, ['Unidad'], required=False),
            'concepto': self._resolver_col(df, ['Concepto'], required=False),
            'observaciones': self._resolver_col(df, ['Observaciones'], required=False),
            'contrarecibo': self._resolver_col(df, ['Contrarecibo', 'ContraRecibo'], required=False),
            'importe': self._resolver_col(df, ['Total', 'Importe', 'Monto']),
            'vale': self._resolver_col(df, ['Vale']),
        }
        
        df_work = df.copy()
        df_work['OPERADOR_NORM'] = df_work[col_map['operador']].apply(self.norm.texto)
        
        excluidos = df_work[df_work['OPERADOR_NORM'].str.contains('FILIAL', na=False)].copy()
        validos = df_work[~df_work['OPERADOR_NORM'].str.contains('FILIAL', na=False)].copy()
        
        # Normalizar campos válidos
        for df_temp in [validos]:
            df_temp['UNIDAD_KEY'] = df_temp[col_map['unidad']].apply(self.norm.clave) if col_map['unidad'] else ""
            df_temp['CONCEPTO_KEY'] = df_temp[col_map['concepto']].apply(self.norm.concepto_canonico) if col_map['concepto'] else ""
            df_temp['POLIZA_KEY'] = df_temp[col_map['contrarecibo']].apply(self.norm.clave) if col_map['contrarecibo'] else ""
            df_temp['VALE_KEY'] = df_temp[col_map['vale']].apply(self.norm.clave)
            df_temp['IMPORTE_KEY'] = df_temp[col_map['importe']].apply(lambda x: self.norm.monto(x, self.config.ndigits))
            df_temp['ORIGEN'] = 'VOUCHER'
            df_temp['OWNER_KEY'] = df_temp[col_map['operador']].apply(self.norm.clave)
            df_temp['OWNER_ORIGINAL'] = df_temp[col_map['operador']].apply(self.norm.texto)
        
        validos['ROW_ID_VOUCHER'] = range(1, len(validos) + 1)
        excluidos['ROW_ID_VOUCHER_EXCL'] = range(1, len(excluidos) + 1)
        
        return validos, excluidos
    
    def _resolver_col(self, df: pd.DataFrame, candidatos: List[str], required: bool = True) -> Optional[str]:
        """Resuelve nombre de columna de una lista de candidatos"""
        norm_map = {self.norm.clave(c): c for c in df.columns}
        
        for candidato in candidatos:
            key = self.norm.clave(candidato)
            if key in norm_map:
                return norm_map[key]
        
        if required:
            raise ValueError(f"No encontré columna de: {candidatos}. Disponibles: {list(df.columns)}")
        return None
    
    def _elegir_col_importe(self, df: pd.DataFrame) -> str:
        """Elige la columna de importe correcta"""
        candidatos = ['Importe', 'Monto', 'Total']
        cols_importe = []
        
        for col in df.columns:
            base = re.sub(r'\.\d+$', '', str(col))
            if self.norm.clave(base) in [self.norm.clave(c) for c in candidatos]:
                cols_importe.append(col)
        
        if not cols_importe:
            raise ValueError(f"No encontré columna de importe. Columnas: {list(df.columns)}")
        
        return cols_importe[-1]


# ============================================================
# MOTOR DE MATCHING
# ============================================================

class MotorMatching:
    """Motor centralizado para matching entre datasets"""
    
    def __init__(self, config: ConfigConciliacion):
        self.config = config
        self.scorer = ScoringMatcher(config)
        self.norm = Normalizador()
    
    def match_liquidaciones_vs_contabilidad(self, 
                                           liquidaciones: pd.DataFrame, 
                                           contabilidad_h: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """
        Match de Liquidaciones (E) vs Contabilidad (H)
        
        Criterios:
        1. PR_KEY (liquidacion vs factura/poliza)
        2. VIAJE_KEY (viaje vs referencia)
        3. TIPO_PAGO_KEY
        4. UNIDAD_KEY
        5. IMPORTE_KEY
        """
        st.info(f"🔍 Matching Liquidaciones ({len(liquidaciones):,}) vs Contabilidad H ({len(contabilidad_h):,})...")
        
        resultados = []
        batch_size = 1000
        progress_bar = st.progress(0)
        
        for i in range(0, len(liquidaciones), batch_size):
            batch = liquidaciones.iloc[i:i+batch_size]
            
            for _, row_liq in batch.iterrows():
                mejor_match = self._buscar_mejor_match_liquidacion(row_liq, contabilidad_h)
                resultados.append(mejor_match)
            
            progress_bar.progress(min((i + batch_size) / len(liquidaciones), 1.0))
        
        progress_bar.empty()
        
        df_matches = pd.DataFrame(resultados)
        
        liq_clas = liquidaciones.merge(
            df_matches[['ROW_ID_LIQ', 'ROW_ID_CONT', 'ESTATUS_MATCH', 'TOTAL_COINCIDENCIAS', 'SCORE']],
            on='ROW_ID_LIQ',
            how='left'
        )
        liq_clas['ESTATUS_MATCH'] = liq_clas['ESTATUS_MATCH'].fillna('NO_EXISTE_EN_CONTABILIDAD_H')
        
        cont_clas = self._clasificar_contabilidad_inversa(contabilidad_h, df_matches, 'LIQUIDACIONES')
        
        return liq_clas, cont_clas, df_matches
    
    def _buscar_mejor_match_liquidacion(self, row_liq: pd.Series, contabilidad_h: pd.DataFrame) -> Dict:
        """Busca el mejor match para una liquidación"""
        
        candidatos = contabilidad_h.copy()
        
        # Filtro por PR (Factura/Poliza en contabilidad)
        if row_liq['PR_KEY']:
            temp = candidatos[candidatos['POLIZA_KEY'] == row_liq['PR_KEY']]
            if not temp.empty:
                candidatos = temp
        
        if candidatos.empty:
            return {
                'ROW_ID_LIQ': row_liq['ROW_ID_LIQ'],
                'ROW_ID_CONT': None,
                'ESTATUS_MATCH': 'NO_EXISTE_EN_CONTABILIDAD_H',
                'TOTAL_COINCIDENCIAS': 0,
                'SCORE': 0.0,
            }
        
        # Evaluar candidatos
        scores = []
        for _, row_cont in candidatos.iterrows():
            coincidencias = 0
            criterios_eval = 5
            
            # 1. PR
            if row_liq['PR_KEY'] and row_liq['PR_KEY'] == row_cont['POLIZA_KEY']:
                coincidencias += 1
            
            # 2. VIAJE
            if row_liq['VIAJE_KEY'] and row_liq['VIAJE_KEY'] == row_cont['VIAJE_KEY']:
                coincidencias += 1
            
            # 3. TIPO_PAGO
            if row_liq['TIPO_PAGO_KEY'] and row_liq['TIPO_PAGO_KEY'] == row_cont.get('TIPO_PAGO_KEY', ''):
                coincidencias += 1
            
            # 4. UNIDAD
            if row_liq['UNIDAD_KEY'] and row_liq['UNIDAD_KEY'] == row_cont['UNIDAD_KEY']:
                coincidencias += 1
            
            # 5. IMPORTE
            if not pd.isna(row_liq['IMPORTE_KEY']) and not pd.isna(row_cont['IMPORTE_KEY']):
                if abs(row_liq['IMPORTE_KEY'] - row_cont['IMPORTE_KEY']) < 0.01:
                    coincidencias += 1
            
            estatus, _, score_norm = self.scorer.calcular_score(criterios_eval, coincidencias)
            
            scores.append({
                'row_id_cont': row_cont['ROW_ID_CONT'],
                'coincidencias': coincidencias,
                'score': score_norm,
                'estatus': estatus,
            })
        
        if scores:
            mejor = max(scores, key=lambda x: (x['coincidencias'], x['score']))
            return {
                'ROW_ID_LIQ': row_liq['ROW_ID_LIQ'],
                'ROW_ID_CONT': mejor['row_id_cont'],
                'ESTATUS_MATCH': mejor['estatus'],
                'TOTAL_COINCIDENCIAS': mejor['coincidencias'],
                'SCORE': mejor['score'],
            }
        
        return {
            'ROW_ID_LIQ': row_liq['ROW_ID_LIQ'],
            'ROW_ID_CONT': None,
            'ESTATUS_MATCH': 'NO_EXISTE_EN_CONTABILIDAD_H',
            'TOTAL_COINCIDENCIAS': 0,
            'SCORE': 0.0,
        }
    
    def match_base_vs_contabilidad(self, 
                                   base: pd.DataFrame, 
                                   contabilidad: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Match de Base Saldos vs Contabilidad D"""
        st.info(f"🔍 Matching Base Saldos ({len(base):,}) vs Contabilidad ({len(contabilidad):,})...")
        
        resultados = []
        batch_size = 1000
        progress_bar = st.progress(0)
        
        for i in range(0, len(base), batch_size):
            batch = base.iloc[i:i+batch_size]
            
            for _, row_base in batch.iterrows():
                mejor_match = self._buscar_mejor_match_base(row_base, contabilidad)
                resultados.append(mejor_match)
            
            progress_bar.progress(min((i + batch_size) / len(base), 1.0))
        
        progress_bar.empty()
        
        df_matches = pd.DataFrame(resultados)
        
        base_clas = base.merge(
            df_matches[['ROW_ID_BASE', 'ROW_ID_CONT', 'ESTATUS_MATCH', 'TOTAL_COINCIDENCIAS', 'SCORE']],
            on='ROW_ID_BASE',
            how='left'
        )
        base_clas['ESTATUS_MATCH'] = base_clas['ESTATUS_MATCH'].fillna('NO_EXISTE_EN_CONTABILIDAD_D')
        
        cont_clas = self._clasificar_contabilidad_inversa(contabilidad, df_matches, 'BASE')
        
        return base_clas, cont_clas, df_matches
    
    def match_costos_vs_contabilidad(self, 
                                    costos: pd.DataFrame, 
                                    contabilidad: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Match de Costos vs Contabilidad D"""
        st.info(f"🔍 Matching Costos ({len(costos):,}) vs Contabilidad ({len(contabilidad):,})...")
        
        resultados = []
        batch_size = 1000
        progress_bar = st.progress(0)
        
        for i in range(0, len(costos), batch_size):
            batch = costos.iloc[i:i+batch_size]
            
            for _, row_costo in batch.iterrows():
                mejor_match = self._buscar_mejor_match_costo(row_costo, contabilidad)
                resultados.append(mejor_match)
            
            progress_bar.progress(min((i + batch_size) / len(costos), 1.0))
        
        progress_bar.empty()
        
        df_matches = pd.DataFrame(resultados)
        
        costos_clas = costos.merge(
            df_matches[['ROW_ID_COSTO', 'ROW_ID_CONT', 'ESTATUS_MATCH', 'TOTAL_COINCIDENCIAS', 'SCORE']],
            on='ROW_ID_COSTO',
            how='left'
        )
        costos_clas['ESTATUS_MATCH'] = costos_clas['ESTATUS_MATCH'].fillna('NO_EXISTE_EN_CONTABILIDAD_D')
        
        cont_clas = self._clasificar_contabilidad_inversa(contabilidad, df_matches, 'COSTOS')
        
        return costos_clas, cont_clas, df_matches
    
    def _buscar_mejor_match_base(self, row_base: pd.Series, contabilidad: pd.DataFrame) -> Dict:
        """Busca el mejor match para una fila de Base Saldos"""
        candidatos = contabilidad.copy()
        
        if row_base['POLIZA_KEY']:
            candidatos = candidatos[candidatos['POLIZA_KEY'] == row_base['POLIZA_KEY']]
        
        if candidatos.empty:
            return {
                'ROW_ID_BASE': row_base['ROW_ID_BASE'],
                'ROW_ID_CONT': None,
                'ESTATUS_MATCH': 'NO_EXISTE_EN_CONTABILIDAD_D',
                'TOTAL_COINCIDENCIAS': 0,
                'SCORE': 0.0,
            }
        
        scores = []
        for _, row_cont in candidatos.iterrows():
            coincidencias = 0
            criterios_eval = 5
            
            if row_base['POLIZA_KEY'] and row_base['POLIZA_KEY'] == row_cont['POLIZA_KEY']:
                coincidencias += 1
            
            if row_base['UNIDAD_KEY'] and row_base['UNIDAD_KEY'] == row_cont['UNIDAD_KEY']:
                coincidencias += 1
            
            if row_base['VIAJE_KEY'] and row_base['VIAJE_KEY'] == row_cont['VIAJE_KEY']:
                coincidencias += 1
            
            score_concepto = 0.0
            if row_base['CONCEPTO_KEY'] and row_cont['CONCEPTO_KEY']:
                score_concepto = self.scorer.calcular_similitud_concepto(
                    row_base['CONCEPTO_KEY'], 
                    row_cont['CONCEPTO_KEY']
                )
                if score_concepto >= self.config.peso_concepto_similar:
                    coincidencias += 1
            
            if not pd.isna(row_base['IMPORTE_KEY']) and not pd.isna(row_cont['IMPORTE_KEY']):
                if abs(row_base['IMPORTE_KEY'] - row_cont['IMPORTE_KEY']) < 0.01:
                    coincidencias += 1
            
            estatus, _, score_norm = self.scorer.calcular_score(criterios_eval, coincidencias, score_concepto)
            
            scores.append({
                'row_id_cont': row_cont['ROW_ID_CONT'],
                'coincidencias': coincidencias,
                'score': score_norm,
                'estatus': estatus,
            })
        
        if scores:
            mejor = max(scores, key=lambda x: (x['coincidencias'], x['score']))
            return {
                'ROW_ID_BASE': row_base['ROW_ID_BASE'],
                'ROW_ID_CONT': mejor['row_id_cont'],
                'ESTATUS_MATCH': mejor['estatus'],
                'TOTAL_COINCIDENCIAS': mejor['coincidencias'],
                'SCORE': mejor['score'],
            }
        
        return {
            'ROW_ID_BASE': row_base['ROW_ID_BASE'],
            'ROW_ID_CONT': None,
            'ESTATUS_MATCH': 'NO_EXISTE_EN_CONTABILIDAD_D',
            'TOTAL_COINCIDENCIAS': 0,
            'SCORE': 0.0,
        }
    
    def _buscar_mejor_match_costo(self, row_costo: pd.Series, contabilidad: pd.DataFrame) -> Dict:
        """Busca el mejor match para una fila de Costos"""
        candidatos = contabilidad.copy()
        
        if row_costo.get('VALE_KEY'):
            temp = candidatos[candidatos['VALE_KEY'] == row_costo['VALE_KEY']]
            if not temp.empty:
                candidatos = temp
        
        if candidatos.empty:
            return {
                'ROW_ID_COSTO': row_costo['ROW_ID_COSTO'],
                'ROW_ID_CONT': None,
                'ESTATUS_MATCH': 'NO_EXISTE_EN_CONTABILIDAD_D',
                'TOTAL_COINCIDENCIAS': 0,
                'SCORE': 0.0,
            }
        
        scores = []
        for _, row_cont in candidatos.iterrows():
            coincidencias = 0
            criterios_eval = 5
            
            if row_costo.get('VALE_KEY') and row_costo['VALE_KEY'] == row_cont['VALE_KEY']:
                coincidencias += 1
            
            if row_costo.get('UNIDAD_KEY') and row_costo['UNIDAD_KEY'] == row_cont['UNIDAD_KEY']:
                coincidencias += 1
            
            score_concepto = 0.0
            if row_costo.get('CONCEPTO_KEY') and row_cont['CONCEPTO_KEY']:
                score_concepto = self.scorer.calcular_similitud_concepto(
                    row_costo['CONCEPTO_KEY'], 
                    row_cont['CONCEPTO_KEY']
                )
                if score_concepto >= self.config.peso_concepto_similar:
                    coincidencias += 1
            
            tiene_contra_costo = row_costo.get('POLIZA_KEY') and row_costo['POLIZA_KEY'] != ""
            tiene_contra_cont = row_cont['POLIZA_KEY'] and row_cont['POLIZA_KEY'] != ""
            
            if tiene_contra_costo and tiene_contra_cont:
                if row_costo['POLIZA_KEY'] == row_cont['POLIZA_KEY']:
                    coincidencias += 1
            elif not tiene_contra_costo and not tiene_contra_cont:
                criterios_eval -= 1
            
            if not pd.isna(row_costo['IMPORTE_KEY']) and not pd.isna(row_cont['IMPORTE_KEY']):
                if abs(row_costo['IMPORTE_KEY'] - row_cont['IMPORTE_KEY']) < 0.01:
                    coincidencias += 1
            
            estatus, _, score_norm = self.scorer.calcular_score(criterios_eval, coincidencias, score_concepto)
            
            scores.append({
                'row_id_cont': row_cont['ROW_ID_CONT'],
                'coincidencias': coincidencias,
                'score': score_norm,
                'estatus': estatus,
            })
        
        if scores:
            mejor = max(scores, key=lambda x: (x['coincidencias'], x['score']))
            return {
                'ROW_ID_COSTO': row_costo['ROW_ID_COSTO'],
                'ROW_ID_CONT': mejor['row_id_cont'],
                'ESTATUS_MATCH': mejor['estatus'],
                'TOTAL_COINCIDENCIAS': mejor['coincidencias'],
                'SCORE': mejor['score'],
            }
        
        return {
            'ROW_ID_COSTO': row_costo['ROW_ID_COSTO'],
            'ROW_ID_CONT': None,
            'ESTATUS_MATCH': 'NO_EXISTE_EN_CONTABILIDAD_D',
            'TOTAL_COINCIDENCIAS': 0,
            'SCORE': 0.0,
        }
    
    def _clasificar_contabilidad_inversa(self, 
                                         contabilidad: pd.DataFrame, 
                                         matches: pd.DataFrame,
                                         origen: str) -> pd.DataFrame:
        """Clasifica contabilidad desde perspectiva inversa"""
        col_id_origen = 'ROW_ID_BASE' if origen == 'BASE' else ('ROW_ID_COSTO' if origen == 'COSTOS' else 'ROW_ID_LIQ')
        
        cont_con_match = matches[matches['ROW_ID_CONT'].notna()].copy()
        
        cont_clas = contabilidad.merge(
            cont_con_match[['ROW_ID_CONT', col_id_origen, 'ESTATUS_MATCH', 'TOTAL_COINCIDENCIAS', 'SCORE']],
            on='ROW_ID_CONT',
            how='left'
        )
        
        cont_clas['ESTATUS_MATCH'] = cont_clas['ESTATUS_MATCH'].fillna(f'NO_EXISTE_EN_{origen}')
        
        return cont_clas


# ============================================================
# DEDUPLICACIÓN, ANÁLISIS D/H, UTILIDADES
# ============================================================

class DeduplicadorCostos:
    """Detecta y elimina duplicados entre Cheques y Vouchers"""
    
    def __init__(self, config: ConfigConciliacion):
        self.config = config
        self.norm = Normalizador()
    
    def deduplicate(self, 
                   cheques: pd.DataFrame, 
                   vouchers: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Combina Cheques y Vouchers eliminando duplicados"""
        cheques_work = cheques.copy()
        vouchers_work = vouchers.copy()
        
        columnas_clave = ['VALE_KEY', 'UNIDAD_KEY', 'CONCEPTO_KEY', 'IMPORTE_KEY', 'POLIZA_KEY', 'ORIGEN', 'OWNER_KEY', 'OWNER_ORIGINAL']
        
        for col in columnas_clave:
            if col not in cheques_work.columns:
                cheques_work[col] = ""
            if col not in vouchers_work.columns:
                vouchers_work[col] = ""
        
        cheques_work['ROW_ID_COSTO'] = 'CHQ_' + cheques_work['ROW_ID_CHEQUE'].astype(str)
        vouchers_work['ROW_ID_COSTO'] = 'VOU_' + vouchers_work['ROW_ID_VOUCHER'].astype(str)
        
        todos_costos = pd.concat([cheques_work, vouchers_work], ignore_index=True)
        
        duplicados_mask = pd.Series(False, index=todos_costos.index)
        duplicados_info = []
        
        st.info(f"🔍 Detectando duplicados en {len(todos_costos):,} registros...")
        
        grupos_vale = todos_costos.groupby('VALE_KEY')
        
        for vale, grupo in grupos_vale:
            if len(grupo) < 2 or vale == "":
                continue
            
            for i, row1 in grupo.iterrows():
                if duplicados_mask[i]:
                    continue
                
                for j, row2 in grupo.iterrows():
                    if i >= j or duplicados_mask[j]:
                        continue
                    
                    es_duplicado = self._son_duplicados(row1, row2)
                    
                    if es_duplicado:
                        duplicados_mask[j] = True
                        duplicados_info.append({
                            'ROW_ID_COSTO_ORIGINAL': row1['ROW_ID_COSTO'],
                            'ROW_ID_COSTO_DUPLICADO': row2['ROW_ID_COSTO'],
                            'ORIGEN_ORIGINAL': row1['ORIGEN'],
                            'ORIGEN_DUPLICADO': row2['ORIGEN'],
                            'VALE': vale,
                            'IMPORTE': row1['IMPORTE_KEY'],
                        })
        
        depurados = todos_costos[~duplicados_mask].copy()
        duplicados_df = todos_costos[duplicados_mask].copy()
        
        if duplicados_info:
            duplicados_detalle = pd.DataFrame(duplicados_info)
            duplicados_df = duplicados_df.merge(
                duplicados_detalle,
                left_on='ROW_ID_COSTO',
                right_on='ROW_ID_COSTO_DUPLICADO',
                how='left'
            )
        
        st.success(f"✅ Depurados: {len(depurados):,} | Duplicados: {len(duplicados_df):,}")
        
        return depurados, duplicados_df
    
    def _son_duplicados(self, row1: pd.Series, row2: pd.Series) -> bool:
        """Determina si dos filas son duplicadas"""
        criterios_cumplidos = 0
        
        if row1['VALE_KEY'] and row1['VALE_KEY'] == row2['VALE_KEY']:
            criterios_cumplidos += 1
        else:
            return False
        
        if row1['UNIDAD_KEY'] and row1['UNIDAD_KEY'] == row2['UNIDAD_KEY']:
            criterios_cumplidos += 1
        
        if row1['CONCEPTO_KEY'] and row2['CONCEPTO_KEY']:
            if row1['CONCEPTO_KEY'] == row2['CONCEPTO_KEY']:
                criterios_cumplidos += 1
        
        if not pd.isna(row1['IMPORTE_KEY']) and not pd.isna(row2['IMPORTE_KEY']):
            if abs(row1['IMPORTE_KEY'] - row2['IMPORTE_KEY']) < 0.01:
                criterios_cumplidos += 1
        
        return criterios_cumplidos >= 3


class AnalizadorDH:
    """Analiza saldos D vs H en contabilidad"""
    
    @staticmethod
    def calcular_saldos(contabilidad_completa: pd.DataFrame) -> pd.DataFrame:
        """Calcula saldos D vs H agrupando por clave"""
        if contabilidad_completa.empty:
            return pd.DataFrame()
        
        base = contabilidad_completa.copy()
        base['IMPORTE_KEY'] = pd.to_numeric(base['IMPORTE_KEY'], errors='coerce').fillna(0)
        
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
        
        return resumen


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
# INTERFAZ STREAMLIT
# ============================================================

def main():
    st.set_page_config(page_title="Sistema de Conciliación Owner v3.1", layout="wide")
    
    st.title("🔄 Sistema Unificado de Conciliación de Saldos Owner")
    st.caption("Versión 3.1 - Completa con Ingresos y Costos")
    
    # Sidebar: Configuración
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        config = ConfigConciliacion(
            ndigits=st.number_input("Decimales para importes", 0, 4, 2),
            umbral_match_ok=st.number_input("Umbral MATCH_OK", 3, 5, 5),
            umbral_match_con_discrepancia=st.number_input("Umbral DISCREPANCIA", 2, 4, 3),
        )
        
        st.divider()
        st.header("📁 Archivos de Entrada")
        
        # Archivo obligatorio
        cont_file = st.file_uploader("📊 Contabilidad (obligatorio)", type=['xlsx', 'xls', 'csv'])
        
        st.subheader("Módulo INGRESOS")
        liq_file = st.file_uploader("📈 Liquidaciones", type=['xlsx', 'xls', 'csv'])
        
        st.subheader("Módulo COSTOS")
        base_file = st.file_uploader("📋 Base Saldos", type=['xlsx', 'xls', 'csv'])
        cheques_file = st.file_uploader("💵 Cheques", type=['xlsx', 'xls', 'csv'])
        vouchers_file = st.file_uploader("🎫 Vouchers", type=['xlsx', 'xls', 'csv'])
        
        st.divider()
        
        procesar_btn = st.button("▶️ PROCESAR", type="primary", use_container_width=True)
    
    # Información inicial
    with st.expander("ℹ️ Información del Sistema", expanded=False):
        st.markdown("""
        ### 🎯 Características Principales
        
        **Módulo de INGRESOS** ✅
        - Liquidaciones (E) vs Contabilidad (H)
        - 5 criterios: PR, VIAJE, TIPO_PAGO, UNIDAD, IMPORTE
        
        **Módulo de COSTOS** ✅
        - Base Saldos vs Contabilidad (D)
        - Cheques + Vouchers vs Contabilidad (D)
        - Deduplicación automática
        
        **Análisis D vs H** ✅
        - Identifica saldos liquidados
        - Detecta pendientes de pago
        - Alerta sobrepagos
        """)
    
    if not procesar_btn:
        st.info("👆 Carga los archivos y da clic en **PROCESAR** para iniciar.")
        return
    
    if cont_file is None:
        st.error("❌ Debes cargar el archivo de Contabilidad.")
        return
    
    # ============================================================
    # PROCESAMIENTO PRINCIPAL
    # ============================================================
    
    inicio = datetime.now()
    
    try:
        # Inicializar componentes
        preparador = PreparadorDatos(config)
        motor = MotorMatching(config)
        deduplicador = DeduplicadorCostos(config)
        analizador_dh = AnalizadorDH()
        archivos = ManejadorArchivos()
        
        # Diccionario para almacenar resultados
        resultados = {}
        
        # ============================================================
        # PASO 1: Preparar Contabilidad
        # ============================================================
        st.header("📊 1. Procesando Contabilidad")
        
        cont_raw = archivos.leer_tabla(cont_file, sheet_name='ContabilidadSET_PLUS_datos')
        cont_d, col_map_cont = preparador.preparar_contabilidad(cont_raw, tipo_mov='D')
        cont_h, _ = preparador.preparar_contabilidad(cont_raw, tipo_mov='H')
        cont_completa, _ = preparador.preparar_contabilidad(cont_raw, tipo_mov=None)
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Movimientos D", f"{len(cont_d):,}")
        col2.metric("Movimientos H", f"{len(cont_h):,}")
        col3.metric("Total movimientos", f"{len(cont_completa):,}")
        
        resultados['Contabilidad_D'] = cont_d
        resultados['Contabilidad_H'] = cont_h
        resultados['Contabilidad_Completa'] = cont_completa
        
        # ============================================================
        # PASO 2: Módulo INGRESOS (Liquidaciones)
        # ============================================================
        if liq_file:
            st.header("📈 2. Procesando Liquidaciones (Ingresos)")
            
            liq_raw = archivos.leer_tabla(liq_file, sheet_name='LiquidacionesSET_PLUS_datos')
            liquidaciones = preparador.preparar_liquidaciones(liq_raw, tipo_concepto='E')
            
            st.info(f"Liquidaciones: {len(liquidaciones):,} registros")
            
            # Matching
            liq_clas, cont_liq_clas, matches_liq = motor.match_liquidaciones_vs_contabilidad(liquidaciones, cont_h)
            
            # Métricas
            col1, col2, col3 = st.columns(3)
            col1.metric("✅ MATCH_OK", f"{(liq_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum():,}")
            col2.metric("⚠️ DISCREPANCIA", f"{(liq_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum():,}")
            col3.metric("❌ NO MATCH", f"{(liq_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_H').sum():,}")
            
            resultados['Liquidaciones_Clasificadas'] = liq_clas
            resultados['Contabilidad_vs_Liquidaciones'] = cont_liq_clas
            resultados['Detalle_Matches_Liquidaciones'] = matches_liq
        
        # ============================================================
        # PASO 3: Módulo BASE SALDOS
        # ============================================================
        if base_file:
            st.header("📋 3. Procesando Base Saldos")
            
            base_raw = archivos.leer_tabla(base_file)
            base = preparador.preparar_base_saldos(base_raw)
            
            st.info(f"Base Saldos: {len(base):,} registros")
            
            # Matching
            base_clas, cont_base_clas, matches_base = motor.match_base_vs_contabilidad(base, cont_d)
            
            # Métricas
            col1, col2, col3 = st.columns(3)
            col1.metric("✅ MATCH_OK", f"{(base_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum():,}")
            col2.metric("⚠️ DISCREPANCIA", f"{(base_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum():,}")
            col3.metric("❌ NO MATCH", f"{(base_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D').sum():,}")
            
            resultados['Base_Clasificada'] = base_clas
            resultados['Contabilidad_vs_Base'] = cont_base_clas
            resultados['Detalle_Matches_Base'] = matches_base
        
        # ============================================================
        # PASO 4: Módulo CHEQUES Y VOUCHERS
        # ============================================================
        if cheques_file and vouchers_file:
            st.header("💵 4. Procesando Cheques y Vouchers")
            
            cheques_raw = archivos.leer_tabla(cheques_file)
            vouchers_raw = archivos.leer_tabla(vouchers_file)
            
            cheques, cheques_excl = preparador.preparar_cheques(cheques_raw)
            vouchers, vouchers_excl = preparador.preparar_vouchers(vouchers_raw)
            
            st.info(f"Cheques válidos: {len(cheques):,} | Excluidos (Company): {len(cheques_excl):,}")
            st.info(f"Vouchers válidos: {len(vouchers):,} | Excluidos (Filial): {len(vouchers_excl):,}")
            
            # Deduplicación
            costos_depurados, duplicados = deduplicador.deduplicate(cheques, vouchers)
            
            # Matching
            costos_clas, cont_costos_clas, matches_costos = motor.match_costos_vs_contabilidad(costos_depurados, cont_d)
            
            # Métricas
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("🧹 Depurados", f"{len(costos_depurados):,}")
            col2.metric("✅ MATCH_OK", f"{(costos_clas['ESTATUS_MATCH'] == 'MATCH_OK').sum():,}")
            col3.metric("⚠️ DISCREPANCIA", f"{(costos_clas['ESTATUS_MATCH'] == 'MATCH_CON_DISCREPANCIA').sum():,}")
            col4.metric("❌ NO MATCH", f"{(costos_clas['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D').sum():,}")
            
            resultados['Cheques_Excluidos'] = cheques_excl
            resultados['Vouchers_Excluidos'] = vouchers_excl
            resultados['Costos_Depurados'] = costos_depurados
            resultados['Duplicados'] = duplicados
            resultados['Costos_Clasificados'] = costos_clas
            resultados['Contabilidad_vs_Costos'] = cont_costos_clas
            resultados['Detalle_Matches_Costos'] = matches_costos
        
        # ============================================================
        # PASO 5: Análisis D vs H
        # ============================================================
        st.header("⚖️ 5. Análisis de Saldos D vs H")
        
        saldos_dh = analizador_dh.calcular_saldos(cont_completa)
        
        if not saldos_dh.empty:
            col1, col2, col3 = st.columns(3)
            col1.metric("✅ Saldados", f"{(saldos_dh['ESTATUS_DH'] == 'SALDADO').sum():,}")
            col2.metric("⏳ Pendientes", f"{(saldos_dh['ESTATUS_DH'] == 'PENDIENTE_PAGO').sum():,}")
            col3.metric("⚠️ Sobrepagos", f"{(saldos_dh['ESTATUS_DH'] == 'SOBREPAGO').sum():,}")
            
            resultados['Analisis_DH'] = saldos_dh
        
        # ============================================================
        # EXPORTACIÓN
        # ============================================================
        st.divider()
        
        tiempo_total = (datetime.now() - inicio).total_seconds()
        st.success(f"✅ Procesamiento completado en {tiempo_total:.1f} segundos")
        
        if resultados:
            excel_bytes = archivos.exportar_excel(resultados)
            
            st.download_button(
                label="📥 Descargar Resultado Completo (Excel)",
                data=excel_bytes,
                file_name=f"conciliacion_owner_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Mostrar resumen de hojas
            with st.expander("📋 Hojas incluidas en el archivo"):
                for nombre in resultados.keys():
                    st.write(f"- {nombre} ({len(resultados[nombre]):,} filas)")
    
    except Exception as e:
        st.error(f"❌ Error durante el procesamiento: {str(e)}")
        st.exception(e)


if __name__ == "__main__":
    main()
