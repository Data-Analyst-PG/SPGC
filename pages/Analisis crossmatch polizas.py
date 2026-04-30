"""
Análisis de Cross-Match por Póliza - Versión Streamlit
Detecta casos donde el costo está en un tráfico diferente
"""

import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Análisis Cross-Match Pólizas", layout="wide")

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


def analizar_crossmatch_polizas(df_base_no_existe, df_cont, decimales=2):
    """
    Analiza casos donde el costo está en un tráfico diferente pero con la misma póliza
    """
    
    st.info(f"Analizando {len(df_base_no_existe):,} registros...")
    
    # Preparar datos de contabilidad
    df_cont_d = df_cont[df_cont['TipoMovimiento'].str.upper() == 'D'].copy()
    df_cont_h = df_cont[df_cont['TipoMovimiento'].str.upper() == 'H'].copy()
    
    # Crear índices para búsqueda rápida
    progress_bar = st.progress(0)
    st.text("Creando índices de búsqueda...")
    
    # Agrupar contabilidad D por póliza
    cont_d_por_poliza = {}
    for _, row in df_cont_d.iterrows():
        poliza = normalizar_texto(row.get('ClavePoliza', ''))
        if poliza not in cont_d_por_poliza:
            cont_d_por_poliza[poliza] = []
        cont_d_por_poliza[poliza].append(row)
    
    progress_bar.progress(0.3)
    
    # Agrupar contabilidad H por póliza
    cont_h_por_poliza = {}
    for _, row in df_cont_h.iterrows():
        poliza = normalizar_texto(row.get('ClavePoliza', ''))
        if poliza not in cont_h_por_poliza:
            cont_h_por_poliza[poliza] = []
        cont_h_por_poliza[poliza].append(row)
    
    progress_bar.progress(0.5)
    
    # Análisis de cross-matches
    st.text("Analizando cross-matches...")
    resultados = []
    
    total = len(df_base_no_existe)
    for idx, row in df_base_no_existe.iterrows():
        if len(resultados) % 1000 == 0:
            progress_bar.progress(0.5 + (len(resultados) / total * 0.5))
        
        poliza = normalizar_texto(row.get('FOLIO_CONTRARECIBO', ''))
        unidad_base = normalizar_texto(row.get('NUMERO_UNIDAD', ''))
        viaje_base = normalizar_texto(row.get('NUMERO_VIAJE', ''))
        importe_base = normalizar_importe(row.get('Importe', 0), decimales)
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
            'CONT_D_UNIDADES': '',
            'CONT_D_VIAJES': '',
            'CONT_D_IMPORTES': '',
            'CONT_D_OWNERS': '',
            'CONT_H_UNIDADES': '',
            'CONT_H_VIAJES': '',
            'CONT_H_IMPORTES': '',
            'TIPO_CASO': '',
            'OBSERVACIONES': ''
        }
        
        # Buscar en movimientos D
        if poliza in cont_d_por_poliza:
            resultado['TIENE_POLIZA_EN_CONT_D'] = True
            movs_d = cont_d_por_poliza[poliza]
            
            unidades = set()
            viajes = set()
            importes = []
            owners = set()
            
            for mov in movs_d:
                unidades.add(normalizar_texto(mov.get('Unidad', '')))
                viajes.add(normalizar_texto(mov.get('Referencia', '')))
                imp = normalizar_importe(mov.get('Importe', 0), decimales)
                if imp > 0:
                    importes.append(imp)
                owners.add(normalizar_texto(mov.get('NombreCuentaContable', '')))
            
            resultado['CONT_D_UNIDADES'] = ', '.join([u for u in unidades if u])[:100]
            resultado['CONT_D_VIAJES'] = ', '.join([v for v in viajes if v])[:100]
            resultado['CONT_D_IMPORTES'] = ', '.join([str(i) for i in importes[:5]])
            resultado['CONT_D_OWNERS'] = ', '.join([o for o in owners if o])[:100]
            
            # Verificar match de importe
            for imp_d in importes:
                if abs(importe_base - imp_d) < 0.01:  # Match exacto
                    resultado['MATCH_IMPORTE_EXACTO_D'] = True
                    break
                elif imp_d > 0 and abs(importe_base - imp_d) / imp_d < 0.01:  # Match similar
                    resultado['MATCH_IMPORTE_SIMILAR_D'] = True
        
        # Buscar en movimientos H
        if poliza in cont_h_por_poliza:
            resultado['TIENE_POLIZA_EN_CONT_H'] = True
            movs_h = cont_h_por_poliza[poliza]
            
            unidades_h = set()
            viajes_h = set()
            importes_h = []
            
            for mov in movs_h:
                unidades_h.add(normalizar_texto(mov.get('Unidad', '')))
                viajes_h.add(normalizar_texto(mov.get('Referencia', '')))
                imp = normalizar_importe(mov.get('Importe', 0), decimales)
                if imp > 0:
                    importes_h.append(imp)
            
            resultado['CONT_H_UNIDADES'] = ', '.join([u for u in unidades_h if u])[:100]
            resultado['CONT_H_VIAJES'] = ', '.join([v for v in viajes_h if v])[:100]
            resultado['CONT_H_IMPORTES'] = ', '.join([str(i) for i in importes_h[:5]])
            
            # Verificar match de importe en H
            for imp_h in importes_h:
                if abs(importe_base - imp_h) < 0.01:
                    resultado['MATCH_IMPORTE_EXACTO_H'] = True
                    break
        
        # Clasificar tipo de caso
        if resultado['MATCH_IMPORTE_EXACTO_D']:
            resultado['TIPO_CASO'] = 'TRAFICO_DIFERENTE_MATCH_EXACTO'
            resultado['OBSERVACIONES'] = 'Mismo importe en Cont D, diferente unidad/viaje'
        elif resultado['MATCH_IMPORTE_SIMILAR_D']:
            resultado['TIPO_CASO'] = 'TRAFICO_DIFERENTE_MATCH_SIMILAR'
            resultado['OBSERVACIONES'] = 'Importe similar en Cont D, diferente unidad/viaje'
        elif resultado['MATCH_IMPORTE_EXACTO_H']:
            resultado['TIPO_CASO'] = 'COSTO_EN_MOVIMIENTO_H'
            resultado['OBSERVACIONES'] = 'Costo en mov H en vez de D'
        elif resultado['TIENE_POLIZA_EN_CONT_D']:
            resultado['TIPO_CASO'] = 'POLIZA_SIN_MATCH_IMPORTE'
            resultado['OBSERVACIONES'] = 'Póliza existe sin match de importe'
        elif resultado['TIENE_POLIZA_EN_CONT_H']:
            resultado['TIPO_CASO'] = 'POLIZA_SOLO_EN_H'
            resultado['OBSERVACIONES'] = 'Póliza solo tiene movimientos H'
        else:
            resultado['TIPO_CASO'] = 'NO_ENCONTRADO'
            resultado['OBSERVACIONES'] = 'Póliza no encontrada'
        
        resultados.append(resultado)
    
    progress_bar.progress(1.0)
    
    # Crear DataFrame con resultados
    df_resultado = pd.DataFrame(resultados)
    
    # Unir con datos originales
    df_final = pd.concat([
        df_base_no_existe.reset_index(drop=True),
        df_resultado.reset_index(drop=True)
    ], axis=1)
    
    return df_final


def generar_excel(df_analisis):
    """Genera archivo Excel con múltiples hojas"""
    
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Hoja 1: Análisis completo
        df_analisis.to_excel(writer, sheet_name='Analisis_Completo', index=False)
        
        # Hoja 2: Casos con costo en tráfico diferente
        casos_trafico = df_analisis[
            df_analisis['TIPO_CASO'].isin([
                'TRAFICO_DIFERENTE_MATCH_EXACTO',
                'TRAFICO_DIFERENTE_MATCH_SIMILAR'
            ])
        ]
        casos_trafico.to_excel(writer, sheet_name='Trafico_Diferente', index=False)
        
        # Hoja 3: Casos con costo en H
        casos_h = df_analisis[df_analisis['TIPO_CASO'] == 'COSTO_EN_MOVIMIENTO_H']
        casos_h.to_excel(writer, sheet_name='Costo_en_H', index=False)
        
        # Hoja 4: Sin match de importe
        casos_sin_match = df_analisis[df_analisis['TIPO_CASO'] == 'POLIZA_SIN_MATCH_IMPORTE']
        casos_sin_match.to_excel(writer, sheet_name='Sin_Match_Importe', index=False)
        
        # Hoja 5: Resumen
        resumen = pd.DataFrame({
            'Tipo_Caso': df_analisis['TIPO_CASO'].value_counts().index,
            'Cantidad': df_analisis['TIPO_CASO'].value_counts().values,
            'Porcentaje': (df_analisis['TIPO_CASO'].value_counts().values / len(df_analisis) * 100).round(2)
        })
        resumen.to_excel(writer, sheet_name='Resumen', index=False)
    
    output.seek(0)
    return output


# ============================================================
# UI de Streamlit
# ============================================================

st.title("🔍 Análisis Cross-Match por Póliza")
st.caption("Detecta costos registrados en tráficos diferentes")

st.markdown("""
Este análisis identifica casos donde:
- El costo está en **Contabilidad** pero en una **unidad/viaje diferente**
- El **importe coincide** (exacto o similar)
- La **póliza es la misma**
""")

with st.sidebar:
    st.header("📁 Cargar Archivos")
    
    # Opción 1: Cargar reporte existente
    st.subheader("Opción 1: Desde Reporte Existente")
    reporte_file = st.file_uploader(
        "Reporte BASESALDOSVSCONTA.xlsx",
        type=['xlsx', 'xls'],
        key='reporte',
        help="El reporte generado por Saldos_Owner_Costos"
    )
    cont_file = st.file_uploader(
        "ContabilidadSET_PLUS_datos.xlsx",
        type=['xlsx', 'xls'],
        key='cont',
        help="Archivo de contabilidad completo"
    )
    
    st.divider()
    
    # Opción 2: Cargar archivos base
    st.subheader("Opción 2: Desde Archivos Base")
    base_file = st.file_uploader(
        "Base Saldos corregida.xlsx",
        type=['xlsx', 'xls'],
        key='base'
    )
    cont_file2 = st.file_uploader(
        "Contabilidad.xlsx",
        type=['xlsx', 'xls'],
        key='cont2'
    )
    
    st.divider()
    decimales = st.number_input("Decimales para comparación", 0, 4, 2)
    
    analizar_btn = st.button("🚀 Analizar", type="primary", use_container_width=True)

if not analizar_btn:
    st.info("👈 Carga los archivos y haz clic en 'Analizar'")
    st.stop()

# Determinar qué opción usar
usar_reporte = reporte_file is not None and cont_file is not None
usar_base = base_file is not None and cont_file2 is not None

if not usar_reporte and not usar_base:
    st.error("Debes cargar archivos usando la Opción 1 o la Opción 2")
    st.stop()

try:
    # Leer archivos según la opción elegida
    if usar_reporte:
        st.info("Usando Opción 1: Reporte existente")
        df_reporte = pd.read_excel(reporte_file)
        df_cont = pd.read_excel(cont_file, sheet_name='ContabilidadSET_PLUS_datos')
        
        # Filtrar solo NO_EXISTE
        if 'ESTATUS_MATCH' not in df_reporte.columns:
            st.error("El reporte no tiene la columna ESTATUS_MATCH")
            st.stop()
        
        df_no_existe = df_reporte[df_reporte['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D'].copy()
        
    else:  # usar_base
        st.info("Usando Opción 2: Archivos base")
        df_no_existe = pd.read_excel(base_file)
        df_cont = pd.read_excel(cont_file2, sheet_name='ContabilidadSET_PLUS_datos')
    
    st.success(f"✅ Archivos cargados correctamente")
    st.write(f"- Registros a analizar: **{len(df_no_existe):,}**")
    st.write(f"- Registros en Contabilidad: **{len(df_cont):,}**")
    
    # Ejecutar análisis
    with st.spinner("Analizando..."):
        df_resultado = analizar_crossmatch_polizas(df_no_existe, df_cont, decimales)
    
    # Mostrar resultados
    st.divider()
    st.header("📊 Resultados del Análisis")
    
    # Métricas
    col1, col2, col3, col4 = st.columns(4)
    
    total = len(df_resultado)
    trafico_diff = len(df_resultado[df_resultado['TIPO_CASO'].isin([
        'TRAFICO_DIFERENTE_MATCH_EXACTO',
        'TRAFICO_DIFERENTE_MATCH_SIMILAR'
    ])])
    sin_match = len(df_resultado[df_resultado['TIPO_CASO'] == 'POLIZA_SIN_MATCH_IMPORTE'])
    no_encontrado = len(df_resultado[df_resultado['TIPO_CASO'] == 'NO_ENCONTRADO'])
    
    col1.metric("Total Analizado", f"{total:,}")
    col2.metric("Tráfico Diferente", f"{trafico_diff:,}", 
                f"{trafico_diff/total*100:.1f}%")
    col3.metric("Sin Match Importe", f"{sin_match:,}",
                f"{sin_match/total*100:.1f}%")
    col4.metric("No Encontrado", f"{no_encontrado:,}",
                f"{no_encontrado/total*100:.1f}%")
    
    # Resumen por tipo
    st.subheader("Distribución por Tipo de Caso")
    resumen = df_resultado['TIPO_CASO'].value_counts().reset_index()
    resumen.columns = ['Tipo de Caso', 'Cantidad']
    resumen['Porcentaje'] = (resumen['Cantidad'] / total * 100).round(2)
    st.dataframe(resumen, use_container_width=True, hide_index=True)
    
    # Tabs con resultados detallados
    tab1, tab2, tab3, tab4 = st.tabs([
        "🎯 Tráfico Diferente",
        "📋 Análisis Completo",
        "❌ Sin Match Importe",
        "💾 Descargar Excel"
    ])
    
    with tab1:
        st.write("Casos donde el costo SÍ existe pero en diferente unidad/viaje")
        casos_trafico = df_resultado[df_resultado['TIPO_CASO'].isin([
            'TRAFICO_DIFERENTE_MATCH_EXACTO',
            'TRAFICO_DIFERENTE_MATCH_SIMILAR'
        ])]
        st.dataframe(casos_trafico, use_container_width=True, height=500)
    
    with tab2:
        st.write("Todos los registros analizados con detalles de cross-match")
        st.dataframe(df_resultado, use_container_width=True, height=500)
    
    with tab3:
        st.write("Casos donde la póliza existe pero sin match de importe")
        casos_sin = df_resultado[df_resultado['TIPO_CASO'] == 'POLIZA_SIN_MATCH_IMPORTE']
        st.dataframe(casos_sin, use_container_width=True, height=500)
    
    with tab4:
        st.write("Descarga el análisis completo en Excel con múltiples hojas")
        
        excel_data = generar_excel(df_resultado)
        
        st.download_button(
            label="📥 Descargar Análisis Completo (Excel)",
            data=excel_data,
            file_name="Analisis_CrossMatch_Polizas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.info("""
        El archivo Excel contiene:
        - **Analisis_Completo**: Todos los registros con análisis
        - **Trafico_Diferente**: Casos con match de importe
        - **Costo_en_H**: Casos en movimiento H
        - **Sin_Match_Importe**: Sin match de importe
        - **Resumen**: Estadísticas del análisis
        """)

except Exception as e:
    st.error(f"Error al procesar: {str(e)}")
    st.exception(e)
