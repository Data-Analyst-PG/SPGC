"""
Análisis de Cross-Match por Póliza - Versión Mejorada
Rastrea movimientos D y H correctamente:
- D: Mismo contrarecibo (CA/PD)
- H: Mismo número de viaje en contrarecibos PR/CH
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


def extraer_movimientos_detallados(movimientos, max_mostrar=10):
    """
    Extrae detalles de movimientos en formato legible
    Retorna string con formato: Unidad|Viaje|Importe|Owner
    """
    detalles = []
    for i, mov in enumerate(movimientos[:max_mostrar]):
        unidad = normalizar_texto(mov.get('Unidad', 'N/A'))
        viaje = normalizar_texto(mov.get('Referencia', 'N/A'))
        importe = normalizar_importe(mov.get('Importe', 0))
        owner = normalizar_texto(mov.get('NombreCuentaContable', 'N/A'))[:30]
        
        detalle = f"{unidad}|{viaje}|${importe:,.2f}|{owner}"
        detalles.append(detalle)
    
    if len(movimientos) > max_mostrar:
        detalles.append(f"... y {len(movimientos) - max_mostrar} más")
    
    return " || ".join(detalles)


def buscar_movimientos_h_por_viaje(viaje_base, df_cont_h, importe_base, decimales=2):
    """
    Busca movimientos H que tengan:
    1. El mismo número de viaje
    2. El mismo importe (exacto o similar)
    3. Contrarecibo tipo PR o CH
    """
    viaje_normalizado = normalizar_texto(viaje_base)
    
    # Filtrar por viaje
    movs_h_viaje = df_cont_h[
        df_cont_h['Referencia'].apply(normalizar_texto) == viaje_normalizado
    ].copy()
    
    if len(movs_h_viaje) == 0:
        return [], []
    
    # Filtrar contrarecibos PR o CH
    movs_h_viaje = movs_h_viaje[
        movs_h_viaje['ClavePoliza'].str.upper().str.startswith(('PR', 'CH'), na=False)
    ].copy()
    
    # Separar por match de importe
    match_exacto = []
    match_similar = []
    
    for _, mov in movs_h_viaje.iterrows():
        imp_h = normalizar_importe(mov.get('Importe', 0), decimales)
        
        if abs(importe_base - imp_h) < 0.01:  # Match exacto
            match_exacto.append(mov)
        elif imp_h > 0 and abs(importe_base - imp_h) / imp_h < 0.01:  # Similar ~1%
            match_similar.append(mov)
    
    return match_exacto, match_similar


def analizar_crossmatch_polizas(df_base_no_existe, df_cont, decimales=2):
    """
    Analiza casos donde el costo está en un tráfico diferente
    LÓGICA MEJORADA:
    - Movimientos D: Busca en mismo contrarecibo (CA/PD) pero diferente tráfico
    - Movimientos H: Busca por número de viaje en contrarecibos PR/CH
    """
    
    st.info(f"Analizando {len(df_base_no_existe):,} registros...")
    
    # Preparar datos de contabilidad
    df_cont_d = df_cont[df_cont['TipoMovimiento'].str.upper() == 'D'].copy()
    df_cont_h = df_cont[df_cont['TipoMovimiento'].str.upper() == 'H'].copy()
    
    progress_bar = st.progress(0)
    st.text("Creando índices de búsqueda...")
    
    # Índice de movimientos D por póliza
    cont_d_por_poliza = {}
    for _, row in df_cont_d.iterrows():
        poliza = normalizar_texto(row.get('ClavePoliza', ''))
        if poliza not in cont_d_por_poliza:
            cont_d_por_poliza[poliza] = []
        cont_d_por_poliza[poliza].append(row)
    
    progress_bar.progress(0.3)
    
    # Análisis de cross-matches
    st.text("Analizando cross-matches...")
    resultados = []
    
    total = len(df_base_no_existe)
    for idx, row in df_base_no_existe.iterrows():
        if len(resultados) % 500 == 0:
            progress_bar.progress(0.3 + (len(resultados) / total * 0.7))
        
        poliza = normalizar_texto(row.get('FOLIO_CONTRARECIBO', ''))
        unidad_base = normalizar_texto(row.get('NUMERO_UNIDAD', ''))
        viaje_base = normalizar_texto(row.get('NUMERO_VIAJE', ''))
        importe_base = normalizar_importe(row.get('Importe', 0), decimales)
        concepto_base = normalizar_texto(row.get('Concepto contabilidad', ''))
        
        # Inicializar resultado
        resultado = {
            'POLIZA_BASE': poliza,
            'UNIDAD_BASE': unidad_base,
            'VIAJE_BASE': viaje_base,
            'IMPORTE_BASE': importe_base,
            'CONCEPTO_BASE': concepto_base,
            
            # Movimientos D (mismo contrarecibo)
            'TIENE_MOVS_D_MISMA_POLIZA': False,
            'CANT_MOVS_D': 0,
            'MATCH_EXACTO_D': False,
            'MATCH_SIMILAR_D': False,
            'DETALLE_MOVS_D': '',
            'TRAFICO_CARGO_D': '',  # Unidad|Viaje donde está el cargo
            'IMPORTE_CARGO_D': '',
            'OWNER_CARGO_D': '',
            'CONTRARECIBO_CARGO_D': '',
            
            # Movimientos H (por número de viaje en PR/CH)
            'TIENE_MOVS_H_MISMO_VIAJE': False,
            'CANT_MOVS_H': 0,
            'MATCH_EXACTO_H': False,
            'MATCH_SIMILAR_H': False,
            'DETALLE_MOVS_H': '',
            'TRAFICO_ABONO_H': '',  # Unidad|Viaje donde está el abono
            'IMPORTE_ABONO_H': '',
            'OWNER_ABONO_H': '',
            'CONTRARECIBO_ABONO_H': '',
            
            # Clasificación
            'TIPO_CASO': '',
            'DIAGNOSTICO': ''
        }
        
        # ============================================================
        # ANÁLISIS DE MOVIMIENTOS D (mismo contrarecibo CA/PD)
        # ============================================================
        if poliza in cont_d_por_poliza:
            movs_d = cont_d_por_poliza[poliza]
            resultado['TIENE_MOVS_D_MISMA_POLIZA'] = True
            resultado['CANT_MOVS_D'] = len(movs_d)
            
            # Buscar matches de importe
            movs_d_match_exacto = []
            movs_d_match_similar = []
            
            for mov in movs_d:
                imp_d = normalizar_importe(mov.get('Importe', 0), decimales)
                
                if abs(importe_base - imp_d) < 0.01:  # Match exacto
                    movs_d_match_exacto.append(mov)
                elif imp_d > 0 and abs(importe_base - imp_d) / imp_d < 0.01:  # Similar
                    movs_d_match_similar.append(mov)
            
            if movs_d_match_exacto:
                resultado['MATCH_EXACTO_D'] = True
                # Tomar el primer match para detalles
                mov = movs_d_match_exacto[0]
                resultado['TRAFICO_CARGO_D'] = f"{normalizar_texto(mov.get('Unidad', ''))}|{normalizar_texto(mov.get('Referencia', ''))}"
                resultado['IMPORTE_CARGO_D'] = f"${normalizar_importe(mov.get('Importe', 0)):,.2f}"
                resultado['OWNER_CARGO_D'] = normalizar_texto(mov.get('NombreCuentaContable', ''))[:50]
                resultado['CONTRARECIBO_CARGO_D'] = normalizar_texto(mov.get('ClavePoliza', ''))
                resultado['DETALLE_MOVS_D'] = extraer_movimientos_detallados(movs_d_match_exacto, 5)
            elif movs_d_match_similar:
                resultado['MATCH_SIMILAR_D'] = True
                mov = movs_d_match_similar[0]
                resultado['TRAFICO_CARGO_D'] = f"{normalizar_texto(mov.get('Unidad', ''))}|{normalizar_texto(mov.get('Referencia', ''))}"
                resultado['IMPORTE_CARGO_D'] = f"${normalizar_importe(mov.get('Importe', 0)):,.2f}"
                resultado['OWNER_CARGO_D'] = normalizar_texto(mov.get('NombreCuentaContable', ''))[:50]
                resultado['CONTRARECIBO_CARGO_D'] = normalizar_texto(mov.get('ClavePoliza', ''))
                resultado['DETALLE_MOVS_D'] = extraer_movimientos_detallados(movs_d_match_similar, 5)
            else:
                # Póliza existe pero sin match de importe - mostrar todos los importes
                importes_d = [normalizar_importe(m.get('Importe', 0)) for m in movs_d]
                resultado['DETALLE_MOVS_D'] = f"Importes en Cont D: {', '.join([f'${i:,.2f}' for i in importes_d[:10]])}"
        
        # ============================================================
        # ANÁLISIS DE MOVIMIENTOS H (por número de viaje en PR/CH)
        # ============================================================
        if viaje_base:  # Solo si hay número de viaje
            movs_h_exacto, movs_h_similar = buscar_movimientos_h_por_viaje(
                viaje_base, df_cont_h, importe_base, decimales
            )
            
            if movs_h_exacto or movs_h_similar:
                resultado['TIENE_MOVS_H_MISMO_VIAJE'] = True
                resultado['CANT_MOVS_H'] = len(movs_h_exacto) + len(movs_h_similar)
            
            if movs_h_exacto:
                resultado['MATCH_EXACTO_H'] = True
                mov = movs_h_exacto[0]
                resultado['TRAFICO_ABONO_H'] = f"{normalizar_texto(mov.get('Unidad', ''))}|{normalizar_texto(mov.get('Referencia', ''))}"
                resultado['IMPORTE_ABONO_H'] = f"${normalizar_importe(mov.get('Importe', 0)):,.2f}"
                resultado['OWNER_ABONO_H'] = normalizar_texto(mov.get('NombreCuentaContable', ''))[:50]
                resultado['CONTRARECIBO_ABONO_H'] = normalizar_texto(mov.get('ClavePoliza', ''))
                resultado['DETALLE_MOVS_H'] = extraer_movimientos_detallados(movs_h_exacto, 5)
            elif movs_h_similar:
                resultado['MATCH_SIMILAR_H'] = True
                mov = movs_h_similar[0]
                resultado['TRAFICO_ABONO_H'] = f"{normalizar_texto(mov.get('Unidad', ''))}|{normalizar_texto(mov.get('Referencia', ''))}"
                resultado['IMPORTE_ABONO_H'] = f"${normalizar_importe(mov.get('Importe', 0)):,.2f}"
                resultado['OWNER_ABONO_H'] = normalizar_texto(mov.get('NombreCuentaContable', ''))[:50]
                resultado['CONTRARECIBO_ABONO_H'] = normalizar_texto(mov.get('ClavePoliza', ''))
                resultado['DETALLE_MOVS_H'] = extraer_movimientos_detallados(movs_h_similar, 5)
        
        # ============================================================
        # CLASIFICACIÓN DEL CASO
        # ============================================================
        if resultado['MATCH_EXACTO_D'] and resultado['MATCH_EXACTO_H']:
            resultado['TIPO_CASO'] = 'COMPLETO_D_Y_H'
            resultado['DIAGNOSTICO'] = f"✅ Cargo en {resultado['TRAFICO_CARGO_D']} ({resultado['CONTRARECIBO_CARGO_D']}) | Abono en {resultado['TRAFICO_ABONO_H']} ({resultado['CONTRARECIBO_ABONO_H']})"
        
        elif resultado['MATCH_EXACTO_D']:
            resultado['TIPO_CASO'] = 'TRAFICO_DIFERENTE_SOLO_D'
            resultado['DIAGNOSTICO'] = f"⚠️ Cargo encontrado en {resultado['TRAFICO_CARGO_D']} ({resultado['CONTRARECIBO_CARGO_D']}) | Owner: {resultado['OWNER_CARGO_D']} | Sin abono H encontrado"
        
        elif resultado['MATCH_SIMILAR_D']:
            resultado['TIPO_CASO'] = 'TRAFICO_DIFERENTE_D_SIMILAR'
            resultado['DIAGNOSTICO'] = f"⚠️ Cargo similar en {resultado['TRAFICO_CARGO_D']} ({resultado['CONTRARECIBO_CARGO_D']}) | Importe: {resultado['IMPORTE_CARGO_D']} vs Base: ${importe_base:,.2f}"
        
        elif resultado['MATCH_EXACTO_H']:
            resultado['TIPO_CASO'] = 'SOLO_ABONO_H'
            resultado['DIAGNOSTICO'] = f"🔄 Solo abono H en {resultado['TRAFICO_ABONO_H']} ({resultado['CONTRARECIBO_ABONO_H']}) | Sin cargo D encontrado"
        
        elif resultado['TIENE_MOVS_D_MISMA_POLIZA']:
            resultado['TIPO_CASO'] = 'POLIZA_SIN_MATCH_IMPORTE'
            resultado['DIAGNOSTICO'] = f"❌ Póliza existe con {resultado['CANT_MOVS_D']} movimientos D pero sin match de importe | {resultado['DETALLE_MOVS_D']}"
        
        else:
            resultado['TIPO_CASO'] = 'NO_ENCONTRADO'
            resultado['DIAGNOSTICO'] = "❓ Póliza no encontrada en Contabilidad"
        
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
        workbook = writer.book
        
        # Formatos
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True
        })
        
        verde_format = workbook.add_format({'bg_color': '#C6E0B4'})
        amarillo_format = workbook.add_format({'bg_color': '#FFE699'})
        rojo_format = workbook.add_format({'bg_color': '#FFC7CE'})
        
        # Hoja 1: Análisis completo
        df_analisis.to_excel(writer, sheet_name='Analisis_Completo', index=False)
        worksheet = writer.sheets['Analisis_Completo']
        worksheet.set_column('A:Z', 15)
        
        # Hoja 2: Casos completos (D y H)
        casos_completos = df_analisis[df_analisis['TIPO_CASO'] == 'COMPLETO_D_Y_H']
        casos_completos.to_excel(writer, sheet_name='Completos_D_y_H', index=False)
        
        # Hoja 3: Solo cargo D
        casos_solo_d = df_analisis[df_analisis['TIPO_CASO'].isin([
            'TRAFICO_DIFERENTE_SOLO_D',
            'TRAFICO_DIFERENTE_D_SIMILAR'
        ])]
        casos_solo_d.to_excel(writer, sheet_name='Solo_Cargo_D', index=False)
        
        # Hoja 4: Solo abono H
        casos_solo_h = df_analisis[df_analisis['TIPO_CASO'] == 'SOLO_ABONO_H']
        casos_solo_h.to_excel(writer, sheet_name='Solo_Abono_H', index=False)
        
        # Hoja 5: Sin match de importe
        casos_sin_match = df_analisis[df_analisis['TIPO_CASO'] == 'POLIZA_SIN_MATCH_IMPORTE']
        casos_sin_match.to_excel(writer, sheet_name='Sin_Match_Importe', index=False)
        
        # Hoja 6: Resumen ejecutivo
        resumen_data = []
        total = len(df_analisis)
        
        for tipo in df_analisis['TIPO_CASO'].unique():
            count = len(df_analisis[df_analisis['TIPO_CASO'] == tipo])
            resumen_data.append({
                'Tipo_Caso': tipo,
                'Cantidad': count,
                'Porcentaje': f"{count/total*100:.2f}%",
                'Descripción': df_analisis[df_analisis['TIPO_CASO'] == tipo]['DIAGNOSTICO'].iloc[0] if count > 0 else ''
            })
        
        resumen = pd.DataFrame(resumen_data)
        resumen.to_excel(writer, sheet_name='Resumen', index=False)
    
    output.seek(0)
    return output


# ============================================================
# UI de Streamlit
# ============================================================

st.title("🔍 Análisis Cross-Match por Póliza v2")
st.caption("Rastrea cargos (D) y abonos (H) en tráficos diferentes")

st.markdown("""
**Lógica de rastreo:**
- **Movimientos D (Cargo)**: Busca en mismo contrarecibo (CA/PD) pero diferente unidad/viaje
- **Movimientos H (Abono)**: Busca por número de viaje en contrarecibos PR/CH
""")

with st.sidebar:
    st.header("📁 Cargar Archivos")
    
    st.subheader("Opción 1: Desde Reporte Existente")
    reporte_file = st.file_uploader(
        "Reporte BASESALDOSVSCONTA.xlsx",
        type=['xlsx', 'xls'],
        key='reporte'
    )
    cont_file = st.file_uploader(
        "ContabilidadSET_PLUS_datos.xlsx",
        type=['xlsx', 'xls'],
        key='cont'
    )
    
    st.divider()
    
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
    # Leer archivos
    if usar_reporte:
        st.info("Usando Opción 1: Reporte existente")
        df_reporte = pd.read_excel(reporte_file)
        df_cont = pd.read_excel(cont_file, sheet_name='ContabilidadSET_PLUS_datos')
        
        if 'ESTATUS_MATCH' not in df_reporte.columns:
            st.error("El reporte no tiene la columna ESTATUS_MATCH")
            st.stop()
        
        df_no_existe = df_reporte[df_reporte['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D'].copy()
        
    else:
        st.info("Usando Opción 2: Archivos base")
        df_no_existe = pd.read_excel(base_file)
        df_cont = pd.read_excel(cont_file2, sheet_name='ContabilidadSET_PLUS_datos')
    
    st.success(f"✅ Archivos cargados")
    st.write(f"- Registros a analizar: **{len(df_no_existe):,}**")
    st.write(f"- Registros en Contabilidad: **{len(df_cont):,}**")
    
    # Ejecutar análisis
    with st.spinner("Analizando..."):
        df_resultado = analizar_crossmatch_polizas(df_no_existe, df_cont, decimales)
    
    # Mostrar resultados
    st.divider()
    st.header("📊 Resultados del Análisis")
    
    # Métricas
    col1, col2, col3, col4, col5 = st.columns(5)
    
    total = len(df_resultado)
    completos = len(df_resultado[df_resultado['TIPO_CASO'] == 'COMPLETO_D_Y_H'])
    solo_d = len(df_resultado[df_resultado['TIPO_CASO'].str.contains('TRAFICO_DIFERENTE', na=False)])
    solo_h = len(df_resultado[df_resultado['TIPO_CASO'] == 'SOLO_ABONO_H'])
    sin_match = len(df_resultado[df_resultado['TIPO_CASO'] == 'POLIZA_SIN_MATCH_IMPORTE'])
    
    col1.metric("Total", f"{total:,}")
    col2.metric("D + H Completos", f"{completos:,}", f"{completos/total*100:.1f}%")
    col3.metric("Solo Cargo D", f"{solo_d:,}", f"{solo_d/total*100:.1f}%")
    col4.metric("Solo Abono H", f"{solo_h:,}", f"{solo_h/total*100:.1f}%")
    col5.metric("Sin Match", f"{sin_match:,}", f"{sin_match/total*100:.1f}%")
    
    # Resumen
    st.subheader("Distribución por Tipo de Caso")
    resumen = df_resultado['TIPO_CASO'].value_counts().reset_index()
    resumen.columns = ['Tipo', 'Cantidad']
    resumen['%'] = (resumen['Cantidad'] / total * 100).round(2)
    st.dataframe(resumen, use_container_width=True, hide_index=True)
    
    # Tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "✅ Completos (D+H)",
        "⚠️ Solo Cargo D",
        "🔄 Solo Abono H",
        "📋 Análisis Completo",
        "💾 Descargar"
    ])
    
    with tab1:
        st.write("Casos donde se encontraron AMBOS: cargo D y abono H")
        casos = df_resultado[df_resultado['TIPO_CASO'] == 'COMPLETO_D_Y_H']
        st.dataframe(casos, use_container_width=True, height=500)
    
    with tab2:
        st.write("Casos donde solo se encontró el cargo D (sin abono H)")
        casos = df_resultado[df_resultado['TIPO_CASO'].str.contains('TRAFICO_DIFERENTE', na=False)]
        st.dataframe(casos, use_container_width=True, height=500)
    
    with tab3:
        st.write("Casos donde solo se encontró el abono H (sin cargo D)")
        casos = df_resultado[df_resultado['TIPO_CASO'] == 'SOLO_ABONO_H']
        st.dataframe(casos, use_container_width=True, height=500)
    
    with tab4:
        st.write("Todos los registros con detalles completos")
        st.dataframe(df_resultado, use_container_width=True, height=500)
    
    with tab5:
        excel_data = generar_excel(df_resultado)
        
        st.download_button(
            label="📥 Descargar Análisis Completo (Excel)",
            data=excel_data,
            file_name="Analisis_CrossMatch_Detallado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.info("""
        **Hojas del Excel:**
        - Analisis_Completo: Todos los registros
        - Completos_D_y_H: Casos con cargo Y abono
        - Solo_Cargo_D: Solo encontró cargo
        - Solo_Abono_H: Solo encontró abono
        - Sin_Match_Importe: Póliza sin match
        - Resumen: Estadísticas
        """)

except Exception as e:
    st.error(f"Error: {str(e)}")
    st.exception(e)
