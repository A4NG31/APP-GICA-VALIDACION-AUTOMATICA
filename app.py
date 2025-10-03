import os
import sys

# ===== CONFIGURACIÓN CRÍTICA PARA STREAMLIT CLOUD - MEJORADA =====
# Desactivar COMPLETAMENTE el file watcher ANTES de cualquier import
os.environ['STREAMLIT_SERVER_FILE_WATCHER_TYPE'] = 'none'
os.environ['STREAMLIT_CI'] = 'true'
os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_STATIC_SERVING'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION'] = 'false'

# Monkey patch AGGRESIVO para evitar problemas de watcher
import streamlit.web.bootstrap
import streamlit.watcher

# Deshabilitar TODOS los watchers
def no_op_watch(*args, **kwargs):
    return lambda: None

def no_op_watch_file(*args, **kwargs):
    return

# Reemplazar todas las funciones de watcher
streamlit.watcher.path_watcher.watch_file = no_op_watch_file
streamlit.watcher.path_watcher._watch_path = no_op_watch
streamlit.watcher.event_based_path_watcher.EventBasedPathWatcher.__init__ = lambda *args, **kwargs: None
streamlit.web.bootstrap._install_config_watchers = lambda *args, **kwargs: None

# ===== IMPORTS NORMALES =====
import streamlit as st
import pandas as pd
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
st._config.set_option('server.fileWatcherType', 'none')

# Resto de tu código permanece igual...
def setup_driver():
    """Configurar ChromeDriver para Selenium - OPTIMIZADO PARA STREAMLIT CLOUD"""
    try:
        chrome_options = Options()
        
        # Opciones CRÍTICAS para Streamlit Cloud
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-software-rasterizer")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        chrome_options.add_argument("--disable-background-timer-throttling")
        chrome_options.add_argument("--disable-backgrounding-occluded-windows")
        chrome_options.add_argument("--disable-renderer-backgrounding")
        chrome_options.add_argument("--disable-ipc-flooding-protection")
        chrome_options.add_argument("--no-default-browser-check")
        chrome_options.add_argument("--no-first-run")
        chrome_options.add_argument("--disable-default-apps")
        chrome_options.add_argument("--disable-translate")
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--allow-running-insecure-content")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-popup-blocking")
        
        # User agent real
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        # Configuraciones experimentales
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_experimental_option("prefs", {
            "profile.default_content_setting_values.notifications": 2,
            "profile.default_content_settings.popups": 0,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        
        # Método OPTIMIZADO para Streamlit Cloud
        try:
            # Usar ChromeDriverManager con configuración específica
            service = Service(
                ChromeDriverManager().install(),
                service_args=['--verbose']
            )
            driver = webdriver.Chrome(service=service, options=chrome_options)
        except Exception as e:
            st.warning(f"⚠️ Método alternativo de Chrome: {e}")
            # Fallback: usar Chrome directamente
            driver = webdriver.Chrome(options=chrome_options)
        
        # Configuraciones adicionales del driver
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        driver.execute_cdp_cmd('Network.setUserAgentOverride', {
            "userAgent": 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        })
        
        # Configurar timeouts
        driver.set_page_load_timeout(60)
        driver.set_script_timeout(30)
        
        return driver
        
    except Exception as e:
        st.error(f"❌ Error crítico al configurar ChromeDriver: {e}")
        return None

# ... (el resto de tus funciones permanecen igual)

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
    
    **Optimizado para:** Streamlit Cloud
    **Estado:** ✅ Configuración mejorada para límites del sistema
    """)
    
    # Estado del sistema mejorado
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success("✅ Configuración mejorada aplicada")
    st.sidebar.info("🚀 Watchers deshabilitados - Sin límites de inotify")
    
    # Cargar archivo Excel
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader("Selecciona el archivo Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Extraer valores del Excel
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
            st.subheader("📅 Parámetros de Búsqueda en Power BI")
            fecha_conciliacion = st.date_input(
                "Fecha de Conciliación",
                value=pd.to_datetime("2025-09-04")
            )
            
            fecha_objetivo = fecha_conciliacion.strftime("%Y-%m-%d")
            
            # Botón de extracción
            st.markdown("---")
            st.subheader("🚀 Extracción y Validación")
            
            if st.button("🎯 Extraer Valor de Power BI y Comparar", type="primary", use_container_width=True):
                with st.spinner("🔄 Iniciando extracción de Power BI... Esto puede tomar 1-2 minutos"):
                    resultados = extract_powerbi_data(fecha_objetivo)
                    
                    if resultados and resultados.get('valor_texto'):
                        valor_powerbi_texto = resultados['valor_texto']
                        
                        st.success("✅ Extracción completada!")
                        st.success(f"**Valor en Power BI:** {valor_powerbi_texto}")
                        
                        # Comparar valores
                        powerbi_numero, excel_numero, valor_formateado, coinciden = compare_values(
                            valor_powerbi_texto, 
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
                                st.write(f"**Diferencia relativa:** {abs(powerbi_numero - excel_numero)/excel_numero*100:.2f}%")
                                
                    elif resultados:
                        st.error("❌ Se accedió al reporte pero no se encontró el valor específico")
                    else:
                        st.error("❌ No se pudieron extraer datos del reporte Power BI")
        else:
            st.error("❌ No se pudieron extraer valores del archivo Excel. Verifica el formato.")
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar la validación")

    # Información de ayuda mejorada
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso - Streamlit Cloud (Configuración Mejorada)"):
        st.markdown("""
        **Configuración Mejorada para Cloud:**
        - ✅ **Watchers completamente deshabilitados** - Sin límites de inotify
        - ✅ **Monkey patch agresivo** para evitar conflictos
        - ✅ **Selenium optimizado** para entorno serverless
        - ✅ **Variables de entorno críticas** configuradas
        
        **Proceso:**
        1. **Cargar Excel**: Archivo con hojas CHICORAL, GUALANDAY, COCORA
        2. **Extracción automática**: Búsqueda inteligente de "Total" en cada hoja
        3. **Seleccionar fecha** de conciliación en Power BI  
        4. **Comparar**: Extrae valor de Power BI y compara con Excel
        
        **Problemas Solucionados:**
        - ❌ `OSError: [Errno 24] inotify instance limit reached`
        - ❌ Límites de file watchers en sistemas compartidos
        - ❌ Conflictos con watchdog en entornos cloud
        """)

if __name__ == "__main__":
    main()
