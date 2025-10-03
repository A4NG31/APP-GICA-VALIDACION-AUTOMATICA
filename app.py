import os
import sys

# ===== CONFIGURACIÓN CRÍTICA PARA STREAMLIT CLOUD - MEJORADA =====
os.environ['STREAMLIT_SERVER_FILE_WATCHER_TYPE'] = 'none'
os.environ['STREAMLIT_CI'] = 'true'
os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_STATIC_SERVING'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION'] = 'false'

# Monkey patch para evitar problemas de watcher
import streamlit.web.bootstrap
import streamlit.watcher

def no_op_watch(*args, **kwargs):
    return lambda: None

def no_op_watch_file(*args, **kwargs):
    return

streamlit.watcher.path_watcher.watch_file = no_op_watch_file
streamlit.watcher.path_watcher._watch_path = no_op_watch
streamlit.watcher.event_based_path_watcher.EventBasedPathWatcher.__init__ = lambda *args, **kwargs: None
streamlit.web.bootstrap._install_config_watchers = lambda *args, **kwargs: None

# ===== IMPORTS NORMALES =====
import streamlit as st
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
import tempfile

# Configuración adicional para Streamlit
st.set_page_config(
    page_title="Validador Power BI - APP GICA",
    page_icon="💰",
    layout="wide"
)

# Funciones de extracción de Excel
def extract_excel_values(uploaded_file):
    """Extraer valores de las hojas del Excel"""
    try:
        valores = {
            'CHICORAL': 0,
            'GUALANDAY': 0, 
            'COCORA': 0
        }
        
        # Leer el archivo Excel
        xls = pd.ExcelFile(uploaded_file)
        
        # Procesar cada hoja
        for hoja in ['CHICORAL', 'GUALANDAY', 'COCORA']:
            if hoja in xls.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=hoja)
                
                # Estrategia 1: Buscar por patrones de texto
                for col in df.columns:
                    # Buscar celdas que contengan "total"
                    mask = df[col].astype(str).str.contains('total', case=False, na=False)
                    if mask.any():
                        # Buscar en la misma fila un valor numérico
                        row_idx = mask.idxmax()
                        for num_col in df.select_dtypes(include=[np.number]).columns:
                            cell_value = df.iloc[row_idx][num_col]
                            if pd.notna(cell_value) and cell_value != 0:
                                valores[hoja] = float(cell_value)
                                st.info(f"✅ Encontrado en {hoja}: ${cell_value:,.0f}")
                                break
                        break
                
                # Estrategia 2: Si no se encontró, buscar la última fila numérica
                if valores[hoja] == 0:
                    numeric_cols = df.select_dtypes(include=[np.number]).columns
                    if len(numeric_cols) > 0:
                        last_row = df[numeric_cols].iloc[-1]
                        for val in last_row:
                            if pd.notna(val) and val != 0:
                                valores[hoja] = float(val)
                                st.info(f"✅ Usando último valor de {hoja}: ${val:,.0f}")
                                break
        
        total_general = sum(valores.values())
        return valores, total_general
        
    except Exception as e:
        st.error(f"❌ Error al procesar Excel: {str(e)}")
        return {'CHICORAL': 0, 'GUALANDAY': 0, 'COCORA': 0}, 0

def setup_driver():
    """Configurar ChromeDriver para Selenium - SOLUCIÓN PARA VERSIÓN 141"""
    try:
        chrome_options = Options()
        
        # Opciones CRÍTICAS para Streamlit Cloud
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--no-default-browser-check")
        chrome_options.add_argument("--no-first-run")
        chrome_options.add_argument("--disable-default-apps")
        chrome_options.add_argument("--disable-translate")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-popup-blocking")
        
        # User agent real
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36")
        
        # Configuraciones experimentales
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # SOLUCIÓN: Forzar la versión correcta de ChromeDriver
        try:
            # Método 1: Usar webdriver-manager con versión específica
            service = Service(ChromeDriverManager(version="141.0.7390.54").install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            st.success("✅ ChromeDriver 141.0.7390.54 configurado correctamente")
            
        except Exception as e:
            st.warning(f"⚠️ Método 1 falló: {e}")
            
            # Método 2: Usar ChromeDriver del sistema (instalado via packages.txt)
            try:
                driver = webdriver.Chrome(options=chrome_options)
                st.success("✅ ChromeDriver del sistema configurado correctamente")
                
            except Exception as e2:
                st.error(f"❌ Método 2 también falló: {e2}")
                
                # Método 3: Usar chromium-browser directamente
                chrome_options.binary_location = "/usr/bin/chromium"
                driver = webdriver.Chrome(options=chrome_options)
                st.success("✅ Chromium browser configurado correctamente")
        
        # Configuraciones adicionales del driver
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        # Configurar timeouts
        driver.set_page_load_timeout(60)
        driver.set_script_timeout(30)
        
        return driver
        
    except Exception as e:
        st.error(f"❌ Error crítico al configurar ChromeDriver: {e}")
        return None

def extract_powerbi_data(fecha_objetivo):
    """Extraer datos de Power BI - VERSIÓN SIMULADA MEJORADA"""
    try:
        # SIMULACIÓN MEJORADA - Con formato de moneda realista
        st.info("🔧 Modo simulación activado")
        
        # Simular un retraso de conexión
        time.sleep(2)
        
        # Valor simulado más realista
        base_value = 1500000
        variation = (hash(fecha_objetivo) % 500000) - 250000  # ±250,000
        valor_simulado = base_value + variation
        
        # Formato de moneda más realista (como lo mostraría Power BI)
        valor_formateado = f"${valor_simulado:,.0f}".replace(",", ".")
        
        return {
            'valor_texto': valor_formateado,
            'valor_numerico': valor_simulado,
            'fecha': fecha_objetivo,
            'estado': 'simulado'
        }
        
    except Exception as e:
        st.error(f"❌ Error en extracción Power BI: {e}")
        return None

def convert_currency_to_float(currency_string):
    """Convierte string de moneda a float - CORREGIDO"""
    try:
        if isinstance(currency_string, (int, float)):
            return float(currency_string)
            
        if isinstance(currency_string, str):
            # Remover símbolos de moneda y espacios
            cleaned = currency_string.replace('$', '').replace(' ', '')
            
            # Manejar formato con puntos como separadores de miles
            if '.' in cleaned and ',' in cleaned:
                # Formato: 1.000.000,00 -> reemplazar . por nada y , por .
                cleaned = cleaned.replace('.', '').replace(',', '.')
            elif '.' in cleaned:
                # Formato: 1.000.000 -> asumir que . son separadores de miles
                parts = cleaned.split('.')
                if len(parts) > 1 and len(parts[-1]) == 3:
                    # Probablemente separadores de miles (1.000.000)
                    cleaned = cleaned.replace('.', '')
                else:
                    # Podría ser decimal (1000.50)
                    pass
            elif ',' in cleaned:
                # Formato: 1,000,000 o 1,000,000.00
                cleaned = cleaned.replace(',', '')
            
            # Convertir a float
            return float(cleaned) if cleaned else 0.0
            
        return float(currency_string)
        
    except Exception as e:
        st.error(f"❌ Error convirtiendo moneda: '{currency_string}' - {e}")
        return 0.0

def compare_values(valor_powerbi, valor_excel):
    """Comparar valores de Power BI y Excel - CORREGIDO"""
    try:
        # Si es un diccionario (resultado simulado)
        if isinstance(valor_powerbi, dict):
            powerbi_numero = valor_powerbi.get('valor_numerico', 0)
            valor_formateado = valor_powerbi.get('valor_texto', '$0')
        else:
            # Convertir texto a número usando la nueva función
            powerbi_numero = convert_currency_to_float(valor_powerbi)
            valor_formateado = valor_powerbi
            
        excel_numero = float(valor_excel)
        
        # Verificar coincidencia (con tolerancia pequeña por redondeos)
        tolerancia = 0.01  # 1 centavo
        coinciden = abs(powerbi_numero - excel_numero) <= tolerancia
        
        return powerbi_numero, excel_numero, valor_formateado, coinciden
        
    except Exception as e:
        st.error(f"❌ Error comparando valores: {e}")
        return None, None, "", False

def main():
    st.title("💰 Validador Power BI - Conciliaciones APP GICA")
    st.markdown("---")
    
    # Información del reporte
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Cargar archivo Excel con 3 hojas
    - Extraer valores de CHICORAL, GUALANDAY, COCORA
    - Calcular total automáticamente
    - Comparar con Power BI
    
    **Estado:** ✅ ChromeDriver 141 Compatible
    **Modo:** 🔧 Incluye simulación para pruebas
    """)
    
    # Estado del sistema
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}")
    st.sidebar.info(f"✅ Pandas {pd.__version__}")
    
    # Cargar archivo Excel
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel con hojas CHICORAL, GUALANDAY, COCORA", 
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        # Mostrar información del archivo
        file_details = {
            "Nombre": uploaded_file.name,
            "Tipo": uploaded_file.type,
            "Tamaño": f"{uploaded_file.size / 1024:.1f} KB"
        }
        st.json(file_details)
        
        # Extraer valores del Excel
        with st.spinner("🔍 Analizando archivo Excel..."):
            valores, total_general = extract_excel_values(uploaded_file)
        
        if total_general > 0:
            st.success("✅ Valores extraídos correctamente del Excel!")
            
            # Mostrar resumen
            st.subheader("📊 Resumen de Valores Encontrados")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                valor_chicoral = f"${valores['CHICORAL']:,.0f}".replace(",", ".")
                st.metric("TOTAL CHICORAL", valor_chicoral)
            
            with col2:
                valor_gualanday = f"${valores['GUALANDAY']:,.0f}".replace(",", ".")
                st.metric("TOTAL GUALANDAY", valor_gualanday)
            
            with col3:
                valor_cocora = f"${valores['COCORA']:,.0f}".replace(",", ".")
                st.metric("TOTAL COCORA", valor_cocora)
            
            with col4:
                total_formateado = f"${total_general:,.0f}".replace(",", ".")
                st.metric("TOTAL GENERAL", total_formateado, delta="Valor de Referencia")
            
            # Parámetros de búsqueda en Power BI
            st.subheader("📅 Parámetros de Búsqueda")
            col1, col2 = st.columns(2)
            
            with col1:
                fecha_conciliacion = st.date_input(
                    "Fecha de Conciliación",
                    value=pd.to_datetime("2025-09-04")
                )
            
            with col2:
                modo_ejecucion = st.selectbox(
                    "Modo de Ejecución",
                    ["Simulación", "Power BI Real"],
                    help="Simulación: Datos de prueba. Power BI Real: Conexión real al reporte"
                )
            
            fecha_objetivo = fecha_conciliacion.strftime("%Y-%m-%d")
            
            # Botón de extracción
            st.markdown("---")
            st.subheader("🚀 Extracción y Validación")
            
            if st.button("🎯 Extraer y Comparar Valores", type="primary", use_container_width=True):
                with st.spinner("🔄 Procesando..."):
                    if modo_ejecucion == "Simulación":
                        resultados = extract_powerbi_data(fecha_objetivo)
                    else:
                        # Para implementación real
                        st.warning("⚠️ Modo Power BI Real - Implementa tu lógica específica")
                        resultados = extract_powerbi_data(fecha_objetivo)
                    
                    if resultados:
                        valor_powerbi_texto = resultados.get('valor_texto', 'No encontrado')
                        
                        st.success("✅ Extracción completada!")
                        
                        if resultados.get('estado') == 'simulado':
                            st.info("🔧 **MODO SIMULACIÓN** - Usando datos de prueba")
                        
                        st.success(f"**Valor en Power BI:** {valor_powerbi_texto}")
                        
                        # Comparar valores
                        powerbi_numero, excel_numero, valor_formateado, coinciden = compare_values(
                            resultados, 
                            total_general
                        )
                        
                        if powerbi_numero is not None and excel_numero is not None:
                            # Mostrar comparación
                            st.subheader("🔍 Resultado de la Validación")
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Power BI", valor_formateado)
                            with col2:
                                st.metric("Excel", total_formateado)
                            with col3:
                                if coinciden:
                                    st.success("✅ COINCIDEN")
                                    st.balloons()
                                else:
                                    diferencia = abs(powerbi_numero - excel_numero)
                                    diferencia_formateada = f"${diferencia:,.0f}".replace(",", ".")
                                    st.error("❌ NO COINCIDEN")
                                    st.metric("Diferencia", diferencia_formateada)
                            
                            # Mostrar detalles
                            with st.expander("📊 Detalles de la Comparación"):
                                st.write(f"**Power BI (numérico):** {powerbi_numero:,.0f}".replace(",", "."))
                                st.write(f"**Excel (numérico):** {excel_numero:,.0f}".replace(",", "."))
                                st.write(f"**Diferencia absoluta:** {abs(powerbi_numero - excel_numero):,.0f}".replace(",", "."))
                                if excel_numero > 0:
                                    st.write(f"**Diferencia relativa:** {abs(powerbi_numero - excel_numero)/excel_numero*100:.2f}%")
                                
                    else:
                        st.error("❌ No se pudieron extraer datos del reporte Power BI")
        else:
            st.error("❌ No se pudieron extraer valores del archivo Excel.")
            st.info("💡 **Sugerencias:**")
            st.info("- Verifica que las hojas se llamen CHICORAL, GUALANDAY, COCORA")
            st.info("- Asegúrate de que haya valores numéricos en las celdas de total")
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar la validación")

    # Información de ayuda
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso:**
        1. **Cargar Excel**: Archivo con hojas CHICORAL, GUALANDAY, COCORA
        2. **Extracción automática**: Búsqueda inteligente de valores totales
        3. **Seleccionar fecha** y modo de ejecución
        4. **Comparar**: Extrae valor y compara con Excel
        
        **Novedades:**
        - ✅ **ChromeDriver 141** - Compatible con la versión actual
        - ✅ **Conversión de moneda mejorada** - Maneja formatos con puntos
        - 🔧 **Múltiples métodos** de configuración de ChromeDriver
        
        **Para implementación real de Power BI:**
        - Reemplaza la función `extract_powerbi_data()` con tu lógica específica
        - Configura los selectores correctos para tu reporte
        """)

if __name__ == "__main__":
    main()
