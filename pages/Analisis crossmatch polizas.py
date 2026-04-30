"""
Script para detectar casos donde el costo está registrado en un tráfico/unidad diferente
entre Base Saldos y Contabilidad, pero con la misma póliza.

Este script complementa el análisis de Saldos_Owner_Costos_v1 identificando:
1. Casos NO_EXISTE_EN_CONTABILIDAD_D que tienen la misma póliza en Contabilidad
2. Diferencias en Unidad y/o Viaje (tráfico)
3. Posibles matches por importe exacto o similar
4. Análisis de movimientos D/H cruzados
"""

import pandas as pd
import numpy as np
from pathlib import Path


def normalizar_texto(texto):
    """Normaliza texto para comparaciones"""
    if pd.isna(texto):
        return ""
    return str(texto).strip().upper()


def normalizar_importe(importe, decimales=2):
    """Normaliza importes para comparación"""
    try:
        return round(float(importe), decimales)
    except:
        return 0.0


def analizar_crossmatch_polizas(archivo_base_saldos, archivo_contabilidad, archivo_reporte_actual=None):
    """
    Analiza casos donde el costo está en un tráfico diferente pero con la misma póliza
    
    Args:
        archivo_base_saldos: Archivo Excel con Base Saldos
        archivo_contabilidad: Archivo Excel con Contabilidad
        archivo_reporte_actual: Archivo Excel del reporte actual (opcional)
    
    Returns:
        DataFrame con análisis de cross-matches
    """
    
    print("Leyendo archivos...")
    
    # Leer Base Saldos
    df_base = pd.read_excel(archivo_base_saldos)
    
    # Leer Contabilidad
    df_cont = pd.read_excel(archivo_contabilidad, sheet_name='ContabilidadSET_PLUS_datos')
    
    # Si existe reporte actual, cargar solo los NO_EXISTE
    if archivo_reporte_actual:
        df_reporte = pd.read_excel(archivo_reporte_actual)
        base_no_existe = df_reporte[df_reporte['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D'].copy()
    else:
        base_no_existe = df_base.copy()
    
    print(f"Total de registros a analizar: {len(base_no_existe)}")
    
    # Preparar datos de contabilidad
    df_cont_d = df_cont[df_cont['TipoMovimiento'].str.upper() == 'D'].copy()
    df_cont_h = df_cont[df_cont['TipoMovimiento'].str.upper() == 'H'].copy()
    
    # Crear índices para búsqueda rápida
    print("Creando índices de búsqueda...")
    
    # Agrupar contabilidad D por póliza
    cont_d_por_poliza = df_cont_d.groupby('ClavePoliza').agg({
        'Unidad': lambda x: list(x),
        'Referencia': lambda x: list(x),
        'Importe': lambda x: list(x),
        'NombreCuentaContable': lambda x: list(x),
        'ConceptoDetalle': lambda x: list(x)
    }).to_dict('index')
    
    # Agrupar contabilidad H por póliza
    cont_h_por_poliza = df_cont_h.groupby('ClavePoliza').agg({
        'Unidad': lambda x: list(x),
        'Referencia': lambda x: list(x),
        'Importe': lambda x: list(x),
        'NombreCuentaContable': lambda x: list(x),
        'ConceptoDetalle': lambda x: list(x)
    }).to_dict('index')
    
    # Análisis de cross-matches
    print("Analizando cross-matches...")
    resultados = []
    
    for idx, row in base_no_existe.iterrows():
        if idx % 1000 == 0:
            print(f"Procesando registro {idx}/{len(base_no_existe)}...")
        
        poliza = row.get('FOLIO_CONTRARECIBO', '')
        unidad_base = normalizar_texto(row.get('NUMERO_UNIDAD', ''))
        viaje_base = normalizar_texto(row.get('NUMERO_VIAJE', ''))
        importe_base = normalizar_importe(row.get('Importe', 0))
        concepto_base = normalizar_texto(row.get('Concepto contabilidad', ''))
        
        resultado = {
            'POLIZA': poliza,
            'BASE_UNIDAD': unidad_base,
            'BASE_VIAJE': viaje_base,
            'BASE_IMPORTE': importe_base,
            'BASE_CONCEPTO': concepto_base,
            'TIENE_POLIZA_EN_CONT_D': False,
            'TIENE_POLIZA_EN_CONT_H': False,
            'MATCH_IMPORTE_EXACTO_D': False,
            'MATCH_IMPORTE_SIMILAR_D': False,
            'MATCH_IMPORTE_EXACTO_H': False,
            'MATCH_IMPORTE_SIMILAR_H': False,
            'CONT_D_UNIDADES': '',
            'CONT_D_VIAJES': '',
            'CONT_D_IMPORTES': '',
            'CONT_D_OWNERS': '',
            'CONT_H_UNIDADES': '',
            'CONT_H_VIAJES': '',
            'CONT_H_IMPORTES': '',
            'CONT_H_OWNERS': '',
            'TIPO_CASO': '',
            'OBSERVACIONES': ''
        }
        
        # Buscar en movimientos D
        if poliza in cont_d_por_poliza:
            resultado['TIENE_POLIZA_EN_CONT_D'] = True
            datos_d = cont_d_por_poliza[poliza]
            
            resultado['CONT_D_UNIDADES'] = ', '.join([str(u) for u in set(datos_d['Unidad']) if pd.notna(u)])
            resultado['CONT_D_VIAJES'] = ', '.join([str(v) for v in set(datos_d['Referencia']) if pd.notna(v)])
            resultado['CONT_D_IMPORTES'] = ', '.join([str(round(i, 2)) for i in datos_d['Importe'] if pd.notna(i)])
            resultado['CONT_D_OWNERS'] = ', '.join([str(o) for o in set(datos_d['NombreCuentaContable']) if pd.notna(o)])
            
            # Verificar match de importe
            importes_d = [normalizar_importe(i) for i in datos_d['Importe'] if pd.notna(i)]
            
            if importe_base in importes_d:
                resultado['MATCH_IMPORTE_EXACTO_D'] = True
            
            # Match similar (diferencia < 1%)
            for imp_d in importes_d:
                if imp_d > 0 and abs(importe_base - imp_d) / imp_d < 0.01:
                    resultado['MATCH_IMPORTE_SIMILAR_D'] = True
                    break
        
        # Buscar en movimientos H
        if poliza in cont_h_por_poliza:
            resultado['TIENE_POLIZA_EN_CONT_H'] = True
            datos_h = cont_h_por_poliza[poliza]
            
            resultado['CONT_H_UNIDADES'] = ', '.join([str(u) for u in set(datos_h['Unidad']) if pd.notna(u)])
            resultado['CONT_H_VIAJES'] = ', '.join([str(v) for v in set(datos_h['Referencia']) if pd.notna(v)])
            resultado['CONT_H_IMPORTES'] = ', '.join([str(round(i, 2)) for i in datos_h['Importe'] if pd.notna(i)])
            resultado['CONT_H_OWNERS'] = ', '.join([str(o) for o in set(datos_h['NombreCuentaContable']) if pd.notna(o)])
            
            # Verificar match de importe en H
            importes_h = [normalizar_importe(i) for i in datos_h['Importe'] if pd.notna(i)]
            
            if importe_base in importes_h:
                resultado['MATCH_IMPORTE_EXACTO_H'] = True
            
            # Match similar
            for imp_h in importes_h:
                if imp_h > 0 and abs(importe_base - imp_h) / imp_h < 0.01:
                    resultado['MATCH_IMPORTE_SIMILAR_H'] = True
                    break
        
        # Clasificar tipo de caso
        if resultado['TIENE_POLIZA_EN_CONT_D'] and resultado['MATCH_IMPORTE_EXACTO_D']:
            resultado['TIPO_CASO'] = 'COSTO_EN_TRAFICO_DIFERENTE_D'
            resultado['OBSERVACIONES'] = 'Mismo importe en Cont D pero diferente unidad/viaje'
        elif resultado['TIENE_POLIZA_EN_CONT_D'] and resultado['MATCH_IMPORTE_SIMILAR_D']:
            resultado['TIPO_CASO'] = 'COSTO_EN_TRAFICO_DIFERENTE_D_SIMILAR'
            resultado['OBSERVACIONES'] = 'Importe similar en Cont D pero diferente unidad/viaje'
        elif resultado['TIENE_POLIZA_EN_CONT_H'] and resultado['MATCH_IMPORTE_EXACTO_H']:
            resultado['TIPO_CASO'] = 'COSTO_EN_MOVIMIENTO_H'
            resultado['OBSERVACIONES'] = 'El costo está en movimiento H (abono) en vez de D (cargo)'
        elif resultado['TIENE_POLIZA_EN_CONT_H'] and resultado['MATCH_IMPORTE_SIMILAR_H']:
            resultado['TIPO_CASO'] = 'COSTO_EN_MOVIMIENTO_H_SIMILAR'
            resultado['OBSERVACIONES'] = 'Importe similar en movimiento H en vez de D'
        elif resultado['TIENE_POLIZA_EN_CONT_D']:
            resultado['TIPO_CASO'] = 'POLIZA_EXISTE_SIN_MATCH_IMPORTE_D'
            resultado['OBSERVACIONES'] = 'Póliza existe en Cont D pero sin match de importe'
        elif resultado['TIENE_POLIZA_EN_CONT_H']:
            resultado['TIPO_CASO'] = 'POLIZA_SOLO_EN_H'
            resultado['OBSERVACIONES'] = 'Póliza solo tiene movimientos H'
        else:
            resultado['TIPO_CASO'] = 'NO_ENCONTRADO'
            resultado['OBSERVACIONES'] = 'Póliza no encontrada en Contabilidad'
        
        resultados.append(resultado)
    
    # Crear DataFrame con resultados
    df_resultado = pd.DataFrame(resultados)
    
    # Unir con datos originales de base_no_existe
    df_resultado = pd.concat([
        base_no_existe.reset_index(drop=True),
        df_resultado.reset_index(drop=True)
    ], axis=1)
    
    # Estadísticas
    print("\n" + "="*80)
    print("RESUMEN DE ANÁLISIS")
    print("="*80)
    print(f"\nTotal de registros analizados: {len(df_resultado)}")
    print(f"\nDistribución por tipo de caso:")
    print(df_resultado['TIPO_CASO'].value_counts())
    
    return df_resultado


def generar_reporte_excel(df_analisis, archivo_salida):
    """Genera un reporte en Excel con múltiples hojas"""
    
    print(f"\nGenerando reporte en: {archivo_salida}")
    
    with pd.ExcelWriter(archivo_salida, engine='xlsxwriter') as writer:
        # Hoja 1: Análisis completo
        df_analisis.to_excel(writer, sheet_name='Analisis_Completo', index=False)
        
        # Hoja 2: Casos con costo en tráfico diferente
        casos_trafico_diff = df_analisis[
            df_analisis['TIPO_CASO'].isin([
                'COSTO_EN_TRAFICO_DIFERENTE_D',
                'COSTO_EN_TRAFICO_DIFERENTE_D_SIMILAR'
            ])
        ]
        casos_trafico_diff.to_excel(writer, sheet_name='Costo_Trafico_Diferente', index=False)
        
        # Hoja 3: Casos con costo en movimiento H
        casos_mov_h = df_analisis[
            df_analisis['TIPO_CASO'].isin([
                'COSTO_EN_MOVIMIENTO_H',
                'COSTO_EN_MOVIMIENTO_H_SIMILAR'
            ])
        ]
        casos_mov_h.to_excel(writer, sheet_name='Costo_en_Mov_H', index=False)
        
        # Hoja 4: Casos sin match de importe
        casos_sin_match = df_analisis[
            df_analisis['TIPO_CASO'] == 'POLIZA_EXISTE_SIN_MATCH_IMPORTE_D'
        ]
        casos_sin_match.to_excel(writer, sheet_name='Sin_Match_Importe', index=False)
        
        # Hoja 5: Resumen estadístico
        resumen = pd.DataFrame({
            'Tipo_Caso': df_analisis['TIPO_CASO'].value_counts().index,
            'Cantidad': df_analisis['TIPO_CASO'].value_counts().values,
            'Porcentaje': (df_analisis['TIPO_CASO'].value_counts().values / len(df_analisis) * 100).round(2)
        })
        resumen.to_excel(writer, sheet_name='Resumen', index=False)
        
        # Formatear
        workbook = writer.book
        
        # Formato para encabezados
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Formato para casos de tráfico diferente
        format_trafico = workbook.add_format({
            'bg_color': '#FFEB9C',
            'border': 1
        })
        
        # Formato para casos en movimiento H
        format_mov_h = workbook.add_format({
            'bg_color': '#C6E0B4',
            'border': 1
        })
    
    print("Reporte generado exitosamente!")
    print(f"\nCasos encontrados:")
    print(f"  - Costo en tráfico diferente (D): {len(casos_trafico_diff)}")
    print(f"  - Costo en movimiento H: {len(casos_mov_h)}")
    print(f"  - Sin match de importe: {len(casos_sin_match)}")


if __name__ == "__main__":
    # Rutas de archivos
    archivo_base = "/mnt/user-data/uploads/Base_Saldos_corregida_v3.xlsx"
    archivo_cont = "/mnt/user-data/uploads/ContabilidadSET_PLUS_datos.xlsx"
    archivo_reporte = "/mnt/user-data/uploads/2026-04-29T20-39_BASESALDOSVSCONTA.xlsx"
    archivo_salida = "/home/claude/Analisis_CrossMatch_Polizas.xlsx"
    
    # Ejecutar análisis
    df_resultado = analizar_crossmatch_polizas(
        archivo_base_saldos=archivo_base,
        archivo_contabilidad=archivo_cont,
        archivo_reporte_actual=archivo_reporte
    )
    
    # Generar reporte
    generar_reporte_excel(df_resultado, archivo_salida)
    
    print("\n¡Análisis completado!")
