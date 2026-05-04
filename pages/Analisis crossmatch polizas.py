"""
Análisis Cross-Match por Póliza - VERSIÓN FINAL
Detecta:
1. Bonificación Diesel (diferencia ~$13-14 entre Base y PD)
2. Tráficos Diferentes (cargo en un viaje, abono en otro)
3. Conceptos Diferentes (ej: Diesel vs Aditivos/Mantto Preventivo)
"""

import pandas as pd
import streamlit as st
from io import BytesIO
import re

st.set_page_config(page_title="Análisis Cross-Match Final", layout="wide")

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    return str(texto).strip().upper()

def normalizar_importe(importe, decimales=2):
    try:
        return round(float(importe), decimales)
    except:
        return 0.0

def detectar_bonificacion_diesel(importe_base, importe_cont, concepto):
    """Detecta si la diferencia es por bonificación diesel (~$13-14)"""
    diferencia = abs(importe_cont - importe_base)
    es_diesel = 'DIESEL' in normalizar_texto(concepto) or 'CONSUMIBLES' in normalizar_texto(concepto)
    
    # Bonificación típica es 13-14 pesos
    if es_diesel and 10 < diferencia < 20:
        return True, diferencia
    return False, 0

def buscar_en_pds(viaje_base, importe_base, concepto_base, df_cont):
    """
    Busca el mismo viaje y concepto en pólizas PD (Póliza Diario)
    para detectar si existe con bonificación diesel
    """
    # Normalizar viaje (quitar /)
    viaje_norm = viaje_base.replace('/', '')
    
    # Buscar en PDs
    pds = df_cont[
        (df_cont['ClavePoliza'].str.startswith('PD-', na=False)) &
        (df_cont['TipoMovimiento'] == 'D') &
        (df_cont['Referencia'].str.contains(viaje_norm, na=False, case=False))
    ]
    
    resultados = []
    for _, row in pds.iterrows():
        imp_pd = normalizar_importe(row['Importe'])
        concepto_pd = normalizar_texto(row.get('ConceptoDetalle', ''))
        
        # Verificar si es el mismo concepto general
        es_mismo_concepto = False
        if 'DIESEL' in concepto_base and 'DIESEL' in concepto_pd:
            es_mismo_concepto = True
        elif 'ANTICIPO' in concepto_base and 'ANTICIPO' in concepto_pd:
            es_mismo_concepto = True
        
        if es_mismo_concepto:
            # Verificar si hay match exacto o con bonificación
            if abs(imp_pd - importe_base) < 0.01:
                resultados.append({
                    'poliza': row['ClavePoliza'],
                    'unidad': row['Unidad'],
                    'viaje': row['Referencia'],
                    'importe': imp_pd,
                    'concepto': concepto_pd,
                    'tipo_match': 'EXACTO',
                    'diferencia': 0
                })
            else:
                tiene_bonif, diff = detectar_bonificacion_diesel(importe_base, imp_pd, concepto_base)
                if tiene_bonif:
                    resultados.append({
                        'poliza': row['ClavePoliza'],
                        'unidad': row['Unidad'],
                        'viaje': row['Referencia'],
                        'importe': imp_pd,
                        'concepto': concepto_pd,
                        'tipo_match': 'BONIFICACION_DIESEL',
                        'diferencia': diff
                    })
    
    return resultados

def buscar_abonos_h_por_viaje(viaje_base, importe_base, df_cont, poliza_base):
    """
    Busca movimientos H (abonos) en contrarecibos PR/CH/RE
    que tengan el mismo viaje e importe
    """
    viaje_norm = viaje_base.replace('/', '')
    
    # Buscar en movimientos H (excluyendo la misma póliza base)
    movs_h = df_cont[
        (df_cont['TipoMovimiento'] == 'H') &
        (df_cont['ClavePoliza'] != poliza_base) &
        (df_cont['Referencia'].str.contains(viaje_norm, na=False, case=False))
    ]
    
    resultados = []
    for _, row in movs_h.iterrows():
        imp_h = normalizar_importe(row['Importe'])
        
        # Match exacto
        if abs(imp_h - importe_base) < 0.01:
            resultados.append({
                'poliza': row['ClavePoliza'],
                'unidad': row['Unidad'],
                'viaje': row['Referencia'],
                'importe': imp_h,
                'concepto': normalizar_texto(row.get('ConceptoDetalle', '')),
                'owner': normalizar_texto(row.get('NombreCuentaContable', '')),
                'tipo_match': 'EXACTO',
                'diferencia': 0
            })
        # Match similar (< 1%)
        elif imp_h > 0 and abs(imp_h - importe_base) / imp_h < 0.01:
            resultados.append({
                'poliza': row['ClavePoliza'],
                'unidad': row['Unidad'],
                'viaje': row['Referencia'],
                'importe': imp_h,
                'concepto': normalizar_texto(row.get('ConceptoDetalle', '')),
                'owner': normalizar_texto(row.get('NombreCuentaContable', '')),
                'tipo_match': 'SIMILAR',
                'diferencia': abs(imp_h - importe_base)
            })
    
    return resultados

def analizar_crossmatch_completo(df_base_no_existe, df_cont, decimales=2):
    """
    Análisis completo que incluye:
    1. Búsqueda en mismo contrarecibo CA (cargo D diferente tráfico)
    2. Búsqueda en PDs (con bonificación diesel)
    3. Búsqueda de abonos H por viaje (en PR/CH/RE)
    """
    
    st.info(f"Analizando {len(df_base_no_existe):,} registros...")
    
    progress_bar = st.progress(0)
    
    # Crear índices
    st.text("Creando índices...")
    cont_d_por_poliza = {}
    for _, row in df_cont[df_cont['TipoMovimiento'] == 'D'].iterrows():
        poliza = normalizar_texto(row.get('ClavePoliza', ''))
        if poliza not in cont_d_por_poliza:
            cont_d_por_poliza[poliza] = []
        cont_d_por_poliza[poliza].append(row)
    
    progress_bar.progress(0.2)
    
    # Análisis
    st.text("Analizando casos...")
    resultados = []
    total = len(df_base_no_existe)
    
    for idx, row in df_base_no_existe.iterrows():
        if len(resultados) % 500 == 0:
            progress_bar.progress(0.2 + (len(resultados) / total * 0.8))
        
        poliza = normalizar_texto(row.get('FOLIO_CONTRARECIBO', ''))
        unidad_base = normalizar_texto(row.get('NUMERO_UNIDAD', ''))
        viaje_base = normalizar_texto(row.get('NUMERO_VIAJE', ''))
        importe_base = normalizar_importe(row.get('Importe', 0), decimales)
        concepto_base = normalizar_texto(row.get('Concepto contabilidad', ''))
        
        resultado = {
            'POLIZA_BASE': poliza,
            'UNIDAD_BASE': unidad_base,
            'VIAJE_BASE': viaje_base,
            'IMPORTE_BASE': importe_base,
            'CONCEPTO_BASE': concepto_base,
            
            # Cargo D en CA
            'TIENE_CARGO_CA_D': False,
            'TRAFICO_CARGO_CA': '',
            'IMPORTE_CARGO_CA': '',
            'CONCEPTO_CARGO_CA': '',
            
            # Cargo D en PD (bonificación)
            'TIENE_CARGO_PD_D': False,
            'TRAFICO_CARGO_PD': '',
            'IMPORTE_CARGO_PD': '',
            'CONCEPTO_CARGO_PD': '',
            'POLIZA_PD': '',
            'ES_BONIFICACION_DIESEL': False,
            'DIFERENCIA_BONIFICACION': 0,
            
            # Abono H
            'TIENE_ABONO_H': False,
            'TRAFICO_ABONO_H': '',
            'IMPORTE_ABONO_H': '',
            'CONCEPTO_ABONO_H': '',
            'POLIZA_ABONO_H': '',
            'OWNER_ABONO_H': '',
            
            # Clasificación
            'TIPO_CASO': '',
            'DIAGNOSTICO': ''
        }
        
        # 1. Buscar en mismo contrarecibo CA (D)
        if poliza in cont_d_por_poliza:
            movs_d = cont_d_por_poliza[poliza]
            for mov in movs_d:
                imp_d = normalizar_importe(mov.get('Importe', 0), decimales)
                if abs(imp_d - importe_base) < 0.01:
                    resultado['TIENE_CARGO_CA_D'] = True
                    resultado['TRAFICO_CARGO_CA'] = f"{normalizar_texto(mov.get('Unidad', ''))}|{normalizar_texto(mov.get('Referencia', ''))}"
                    resultado['IMPORTE_CARGO_CA'] = f"${imp_d:.2f}"
                    resultado['CONCEPTO_CARGO_CA'] = normalizar_texto(mov.get('ConceptoDetalle', ''))[:50]
                    break
        
        # 2. Buscar en PDs (bonificación diesel)
        if viaje_base:
            pds_encontradas = buscar_en_pds(viaje_base, importe_base, concepto_base, df_cont)
            if pds_encontradas:
                pd_match = pds_encontradas[0]  # Tomar el primero
                resultado['TIENE_CARGO_PD_D'] = True
                resultado['TRAFICO_CARGO_PD'] = f"{pd_match['unidad']}|{pd_match['viaje']}"
                resultado['IMPORTE_CARGO_PD'] = f"${pd_match['importe']:.2f}"
                resultado['CONCEPTO_CARGO_PD'] = pd_match['concepto'][:50]
                resultado['POLIZA_PD'] = pd_match['poliza']
                
                if pd_match['tipo_match'] == 'BONIFICACION_DIESEL':
                    resultado['ES_BONIFICACION_DIESEL'] = True
                    resultado['DIFERENCIA_BONIFICACION'] = pd_match['diferencia']
        
        # 3. Buscar abonos H por viaje
        if viaje_base:
            abonos_h = buscar_abonos_h_por_viaje(viaje_base, importe_base, df_cont, poliza)
            if abonos_h:
                abono_match = abonos_h[0]  # Tomar el primero
                resultado['TIENE_ABONO_H'] = True
                resultado['TRAFICO_ABONO_H'] = f"{abono_match['unidad']}|{abono_match['viaje']}"
                resultado['IMPORTE_ABONO_H'] = f"${abono_match['importe']:.2f}"
                resultado['CONCEPTO_ABONO_H'] = abono_match['concepto'][:50]
                resultado['POLIZA_ABONO_H'] = abono_match['poliza']
                resultado['OWNER_ABONO_H'] = abono_match['owner'][:50]
        
        # Clasificar
        if resultado['ES_BONIFICACION_DIESEL']:
            resultado['TIPO_CASO'] = 'BONIFICACION_DIESEL'
            resultado['DIAGNOSTICO'] = f"✅ Existe en PD {resultado['POLIZA_PD']} con bonificación diesel de ${resultado['DIFERENCIA_BONIFICACION']:.2f} | Tráfico: {resultado['TRAFICO_CARGO_PD']}"
        
        elif resultado['TIENE_CARGO_CA_D'] and resultado['TIENE_ABONO_H']:
            resultado['TIPO_CASO'] = 'COMPLETO_CARGO_Y_ABONO'
            resultado['DIAGNOSTICO'] = f"✅ Cargo en {resultado['TRAFICO_CARGO_CA']} ({poliza}) | Abono en {resultado['TRAFICO_ABONO_H']} ({resultado['POLIZA_ABONO_H']})"
        
        elif resultado['TIENE_CARGO_PD_D'] and resultado['TIENE_ABONO_H']:
            resultado['TIPO_CASO'] = 'CARGO_PD_Y_ABONO'
            resultado['DIAGNOSTICO'] = f"✅ Cargo en PD {resultado['TRAFICO_CARGO_PD']} ({resultado['POLIZA_PD']}) | Abono en {resultado['TRAFICO_ABONO_H']} ({resultado['POLIZA_ABONO_H']})"
        
        elif resultado['TIENE_CARGO_CA_D']:
            resultado['TIPO_CASO'] = 'SOLO_CARGO_CA'
            resultado['DIAGNOSTICO'] = f"⚠️ Solo cargo en {resultado['TRAFICO_CARGO_CA']} ({poliza}) | Sin abono H"
        
        elif resultado['TIENE_CARGO_PD_D']:
            resultado['TIPO_CASO'] = 'SOLO_CARGO_PD'
            resultado['DIAGNOSTICO'] = f"⚠️ Solo cargo en PD {resultado['TRAFICO_CARGO_PD']} ({resultado['POLIZA_PD']}) | Sin abono H"
        
        elif resultado['TIENE_ABONO_H']:
            resultado['TIPO_CASO'] = 'SOLO_ABONO_H'
            resultado['DIAGNOSTICO'] = f"🔄 Solo abono en {resultado['TRAFICO_ABONO_H']} ({resultado['POLIZA_ABONO_H']}) | Sin cargo D"
        
        else:
            resultado['TIPO_CASO'] = 'NO_ENCONTRADO'
            resultado['DIAGNOSTICO'] = "❌ No se encontró ni en CA, ni en PD, ni abono H"
        
        resultados.append(resultado)
    
    progress_bar.progress(1.0)
    
    df_resultado = pd.DataFrame(resultados)
    df_final = pd.concat([
        df_base_no_existe.reset_index(drop=True),
        df_resultado.reset_index(drop=True)
    ], axis=1)
    
    return df_final


def generar_excel(df_analisis):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Análisis completo
        df_analisis.to_excel(writer, sheet_name='Analisis_Completo', index=False)
        
        # Bonificación Diesel
        bonif = df_analisis[df_analisis['TIPO_CASO'] == 'BONIFICACION_DIESEL']
        bonif.to_excel(writer, sheet_name='Bonificacion_Diesel', index=False)
        
        # Completos (Cargo + Abono)
        completos = df_analisis[df_analisis['TIPO_CASO'].str.contains('COMPLETO|CARGO.*Y_ABONO', na=False, regex=True)]
        completos.to_excel(writer, sheet_name='Completos_Cargo_Abono', index=False)
        
        # Solo Cargo
        solo_cargo = df_analisis[df_analisis['TIPO_CASO'].str.contains('SOLO_CARGO', na=False)]
        solo_cargo.to_excel(writer, sheet_name='Solo_Cargo', index=False)
        
        # Solo Abono
        solo_abono = df_analisis[df_analisis['TIPO_CASO'] == 'SOLO_ABONO_H']
        solo_abono.to_excel(writer, sheet_name='Solo_Abono', index=False)
        
        # No encontrado
        no_encontrado = df_analisis[df_analisis['TIPO_CASO'] == 'NO_ENCONTRADO']
        no_encontrado.to_excel(writer, sheet_name='No_Encontrado', index=False)
        
        # Resumen
        total = len(df_analisis)
        resumen = pd.DataFrame({
            'Tipo_Caso': df_analisis['TIPO_CASO'].value_counts().index,
            'Cantidad': df_analisis['TIPO_CASO'].value_counts().values,
            'Porcentaje': (df_analisis['TIPO_CASO'].value_counts().values / total * 100).round(2)
        })
        resumen.to_excel(writer, sheet_name='Resumen', index=False)
    
    output.seek(0)
    return output


# UI
st.title("🔍 Análisis Cross-Match - Versión Final")
st.caption("Detecta: Bonificación Diesel | Tráficos Diferentes | Cargos/Abonos Cruzados")

with st.sidebar:
    st.header("📁 Cargar Archivos")
    
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
    decimales = st.number_input("Decimales", 0, 4, 2)
    analizar_btn = st.button("🚀 Analizar", type="primary", use_container_width=True)

if not analizar_btn:
    st.info("👈 Carga los archivos y haz clic en 'Analizar'")
    
    with st.expander("ℹ️ ¿Qué detecta este análisis?"):
        st.markdown("""
        **1. Bonificación Diesel**
        - Casos donde el importe en Base Saldos es ~$13-14 menor que en PD
        - Ejemplo: Base=$402.28, PD=$415.88 (diferencia de $13.60)
        
        **2. Tráficos Diferentes**
        - Cargos (D) en un viaje, abonos (H) en otro viaje
        - Ejemplo: Cargo en SEP00812/17, Abono en SEP00878/17
        
        **3. Conceptos Diferentes**
        - Mismo importe pero concepto diferente (ej: Diesel vs Aditivos)
        """)
    
    st.stop()

if not reporte_file or not cont_file:
    st.error("Debes cargar ambos archivos")
    st.stop()

try:
    df_reporte = pd.read_excel(reporte_file)
    df_cont = pd.read_excel(cont_file, sheet_name='ContabilidadSET_PLUS_datos')
    
    if 'ESTATUS_MATCH' not in df_reporte.columns:
        st.error("El reporte no tiene la columna ESTATUS_MATCH")
        st.stop()
    
    df_no_existe = df_reporte[df_reporte['ESTATUS_MATCH'] == 'NO_EXISTE_EN_CONTABILIDAD_D'].copy()
    
    st.success(f"✅ Archivos cargados")
    st.write(f"- Registros NO_EXISTE: **{len(df_no_existe):,}**")
    st.write(f"- Registros Contabilidad: **{len(df_cont):,}**")
    
    with st.spinner("Analizando..."):
        df_resultado = analizar_crossmatch_completo(df_no_existe, df_cont, decimales)
    
    st.divider()
    st.header("📊 Resultados")
    
    # Métricas
    col1, col2, col3, col4, col5 = st.columns(5)
    
    total = len(df_resultado)
    bonif = len(df_resultado[df_resultado['TIPO_CASO'] == 'BONIFICACION_DIESEL'])
    completos = len(df_resultado[df_resultado['TIPO_CASO'].str.contains('COMPLETO|Y_ABONO', na=False)])
    solo_cargo = len(df_resultado[df_resultado['TIPO_CASO'].str.contains('SOLO_CARGO', na=False)])
    no_encontrado = len(df_resultado[df_resultado['TIPO_CASO'] == 'NO_ENCONTRADO'])
    
    col1.metric("Total", f"{total:,}")
    col2.metric("Bonif. Diesel", f"{bonif:,}", f"{bonif/total*100:.1f}%")
    col3.metric("Completos", f"{completos:,}", f"{completos/total*100:.1f}%")
    col4.metric("Solo Cargo", f"{solo_cargo:,}", f"{solo_cargo/total*100:.1f}%")
    col5.metric("No Encontrado", f"{no_encontrado:,}", f"{no_encontrado/total*100:.1f}%")
    
    # Resumen
    st.subheader("Distribución")
    resumen = df_resultado['TIPO_CASO'].value_counts().reset_index()
    resumen.columns = ['Tipo', 'Cantidad']
    resumen['%'] = (resumen['Cantidad'] / total * 100).round(2)
    st.dataframe(resumen, use_container_width=True, hide_index=True)
    
    # Tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "💰 Bonificación Diesel",
        "✅ Completos",
        "⚠️ Solo Cargo",
        "📋 Todos",
        "💾 Descargar"
    ])
    
    with tab1:
        st.write("Casos con bonificación diesel detectada")
        casos = df_resultado[df_resultado['TIPO_CASO'] == 'BONIFICACION_DIESEL']
        st.dataframe(casos, use_container_width=True, height=500)
    
    with tab2:
        st.write("Casos con cargo Y abono encontrados")
        casos = df_resultado[df_resultado['TIPO_CASO'].str.contains('COMPLETO|Y_ABONO', na=False)]
        st.dataframe(casos, use_container_width=True, height=500)
    
    with tab3:
        st.write("Casos donde solo se encontró el cargo")
        casos = df_resultado[df_resultado['TIPO_CASO'].str.contains('SOLO_CARGO', na=False)]
        st.dataframe(casos, use_container_width=True, height=500)
    
    with tab4:
        st.dataframe(df_resultado, use_container_width=True, height=500)
    
    with tab5:
        excel_data = generar_excel(df_resultado)
        
        st.download_button(
            label="📥 Descargar Análisis (Excel)",
            data=excel_data,
            file_name="Analisis_CrossMatch_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.info("""
        **Hojas:**
        - Bonificacion_Diesel: Casos con bonif diesel
        - Completos_Cargo_Abono: Cargo Y abono
        - Solo_Cargo: Solo cargo encontrado
        - Solo_Abono: Solo abono encontrado
        - No_Encontrado: No se encontró nada
        - Resumen: Estadísticas
        """)

except Exception as e:
    st.error(f"Error: {str(e)}")
    st.exception(e)
