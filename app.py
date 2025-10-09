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

# ===== CSS Sidebar =====
st.markdown("""
<style>
/* ===== Sidebar ===== */
[data-testid="stSidebar"] {
    background-color: #1E1E2F !important;
    color: white !important;
    width: 300px !important;
    padding: 20px 10px 20px 10px !important;
    border-right: 1px solid #333 !important;
}

/* Texto general en blanco */
[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stCheckbox label {
    color: white !important; 
}

/* SOLO el label del file_uploader en blanco */
[data-testid="stSidebar"] .stFileUploader > label {
    color: white !important;
    font-weight: bold;
}

/* Mantener en negro el resto del uploader */
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-title,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-subtitle,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-list button,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-Item-name,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-Item-status,
[data-testid="stSidebar"] .stFileUploader span,
[data-testid="stSidebar"] .stFileUploader div {
    color: black !important;
}

/* ===== Botón de expandir/cerrar sidebar ===== */
[data-testid="stSidebarNav"] button {
    background: #2E2E3E !important;
    color: white !important;
    border-radius: 6px !important;
}

/* ===== Encabezados del sidebar ===== */
[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3 {
    color: #00CFFF !important;
}

/* ===== Inputs de texto en el sidebar ===== */
[data-testid="stSidebar"] input[type="text"],
[data-testid="stSidebar"] input[type="password"] {
    color: black !important;
    background-color: white !important;
    border-radius: 6px !important;
    padding: 5px !important;
}

/* ===== BOTÓN "BROWSE FILES" ===== */
[data-testid="stSidebar"] .uppy-Dashboard-AddFiles-list button {
    color: black !important;
    background-color: #f0f0f0 !important;
    border: 1px solid #ccc !important;
}
[data-testid="stSidebar"] .uppy-Dashboard-AddFiles-list button:hover {
    background-color: #e0e0e0 !important;
}

/* ===== Texto en multiselect ===== */
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="select"] {
    color: white !important;
}
[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="tag"] {
    color: black !important;
    background-color: #e0e0e0 !important;
}

/* ===== ICONOS DE AYUDA (?) EN EL SIDEBAR ===== */
[data-testid="stSidebar"] svg.icon {
    stroke: white !important;
    color: white !important;
    fill: none !important;
    opacity: 1 !important;
}

/* ===== MEJORAS PARA STREAMLIT CLOUD ===== */
.stSpinner > div > div {
    border-color: #00CFFF !important;
}

.stProgress > div > div > div > div {
    background-color: #00CFFF !important;
}
</style>
""", unsafe_allow_html=True)

# Logo de GoPass con HTML
st.markdown("""
<div style="display: flex; justify-content: center; margin-bottom: 30px;">
    <img src="https://i.imgur.com/z9xt46F.jpeg"
         style="width: 50%; border-radius: 10px; display: block; margin: 0 auto;" 
         alt="Logo Gopass">
</div>
""", unsafe_allow_html=True)

# ===== FUNCIONES DE EXTRACCIÓN DE POWER BI =====

def setup_driver():
    """Configurar ChromeDriver para Selenium - VERSIÓN PARA STREAMLIT CLOUD"""
    try:
        chrome_options = Options()
        
        # Opciones para Streamlit Cloud
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-software-rasterizer")
        
        # User agent real
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        # Intentar múltiples métodos
        try:
            # Método 1: Chrome del sistema (Streamlit Cloud)
            chrome_options.binary_location = "/usr/bin/chromium"
            service = Service("/usr/bin/chromedriver")
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            st.success("✅ ChromeDriver configurado correctamente")
            return driver
        except Exception as e1:
            try:
                # Método 2: Sin especificar ubicaciones
                driver = webdriver.Chrome(options=chrome_options)
                driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
                st.success("✅ ChromeDriver configurado correctamente")
                return driver
            except Exception as e2:
                try:
                    # Método 3: Usando webdriver-manager
                    from webdriver_manager.chrome import ChromeDriverManager
                    from webdriver_manager.core.os_manager import ChromeType
                    
                    service = Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install())
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
                    st.success("✅ ChromeDriver configurado correctamente")
                    return driver
                except Exception as e3:
                    st.error(f"❌ Error método 1: {e1}")
                    st.error(f"❌ Error método 2: {e2}")
                    st.error(f"❌ Error método 3: {e3}")
                    return None
            
    except Exception as e:
        st.error(f"❌ Error crítico al configurar ChromeDriver: {e}")
        return None

def click_conciliacion_date(driver, fecha_objetivo):
    """Hacer clic en la conciliación específica por fecha"""
    try:
        selectors = [
            f"//*[contains(text(), 'Conciliación APP GICA del {fecha_objetivo}')]",
            f"//*[contains(text(), 'CONCILIACIÓN APP GICA DEL {fecha_objetivo}')]",
            f"//*[contains(text(), '{fecha_objetivo} 00:00 al {fecha_objetivo} 11:59')]",
            f"//div[contains(text(), '{fecha_objetivo}')]",
            f"//span[contains(text(), '{fecha_objetivo}')]",
        ]
        
        elemento_conciliacion = None
        for selector in selectors:
            try:
                elemento = driver.find_element(By.XPATH, selector)
                if elemento.is_displayed():
                    elemento_conciliacion = elemento
                    break
            except:
                continue
        
        if elemento_conciliacion:
            driver.execute_script("arguments[0].scrollIntoView(true);", elemento_conciliacion)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", elemento_conciliacion)
            time.sleep(3)
            return True
        else:
            st.error("❌ No se encontró la conciliación para la fecha especificada")
            return False
            
    except Exception as e:
        st.error(f"❌ Error al hacer clic en conciliación: {str(e)}")
        return False

def find_cantidad_pasos_card(driver):
    """Buscar la tarjeta 'CANTIDAD PASOS' a la derecha de 'VALOR A PAGAR A COMERCIO'"""
    try:
        titulo_selectors = [
            "//*[contains(text(), 'CANTIDAD PASOS')]",
            "//*[contains(text(), 'Cantidad Pasos')]",
            "//*[contains(text(), 'CANTIDAD DE PASOS')]",
            "//*[contains(text(), 'Cantidad de Pasos')]",
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        if "CANTIDAD" in texto.upper() and "PASOS" in texto.upper():
                            titulo_element = elemento
                            break
                if titulo_element:
                    break
            except:
                continue
        
        if not titulo_element:
            return None
        
        # Buscar el valor numérico
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            numeric_elements = container.find_elements(By.XPATH, ".//*")
            
            for elem in numeric_elements:
                texto = elem.text.strip()
                if texto and any(char.isdigit() for char in texto) and texto != titulo_element.text:
                    if len(texto) < 20 and "$" not in texto:
                        return texto
        except:
            pass
        
        # Estrategia 2
        try:
            parent = titulo_element.find_element(By.XPATH, "./..")
            siblings = parent.find_elements(By.XPATH, "./*")
            
            for sibling in siblings:
                if sibling != titulo_element:
                    texto = sibling.text.strip()
                    if texto and any(char.isdigit() for char in texto):
                        if len(texto) < 20 and "$" not in texto:
                            return texto
        except:
            pass
        
        return None
        
    except Exception as e:
        return None

def find_valor_a_pagar_comercio_card(driver):
    """Buscar la tarjeta 'VALOR A PAGAR A COMERCIO'"""
    try:
        titulo_selectors = [
            "//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]",
            "//*[contains(text(), 'Valor a pagar a comercio')]",
            "//*[contains(text(), 'VALOR A PAGAR') and contains(text(), 'COMERCIO')]",
            "//*[contains(text(), 'Valor A Pagar') and contains(text(), 'Comercio')]",
            "//*[contains(text(), 'PAGAR A COMERCIO')]",
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        if "PAGAR" in texto.upper() and "COMERCIO" in texto.upper():
                            titulo_element = elemento
                            break
                if titulo_element:
                    break
            except:
                continue
        
        if not titulo_element:
            st.error("❌ No se encontró 'VALOR A PAGAR A COMERCIO' en el reporte")
            return None
        
        # Estrategia 1
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            numeric_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '$') or contains(text(), ',') or contains(text(), '.')]")
            
            for elem in numeric_elements:
                texto = elem.text.strip()
                if texto and any(char.isdigit() for char in texto) and texto != titulo_element.text:
                    return texto
        except:
            pass
        
        # Estrategia 2
        try:
            parent = titulo_element.find_element(By.XPATH, "./..")
            siblings = parent.find_elements(By.XPATH, "./*")
            
            for sibling in siblings:
                if sibling != titulo_element:
                    texto = sibling.text.strip()
                    if texto and any(char.isdigit() for char in texto):
                        return texto
        except:
            pass
        
        # Estrategia 3
        try:
            following_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]/following::*")
            
            for elem in following_elements[:10]:
                texto = elem.text.strip()
                if texto and any(char.isdigit() for char in texto) and len(texto) < 50:
                    return texto
        except:
            pass
        
        st.error("❌ No se pudo encontrar el valor numérico")
        return None
        
    except Exception as e:
        st.error(f"❌ Error buscando valor: {str(e)}")
        return None

def find_peaje_values(driver):
    """Buscar valores individuales de cada peaje en el Power BI"""
    peajes = {}
    nombres_peajes = ['CHICORAL', 'COCORA', 'GUALANDAY']
    
    for nombre_peaje in nombres_peajes:
        try:
            titulo_selectors = [
                f"//*[contains(text(), '{nombre_peaje}')]",
                f"//*[contains(text(), '{nombre_peaje.title()}')]",
                f"//*[contains(text(), '{nombre_peaje.lower()}')]",
            ]
            
            titulo_element = None
            for selector in titulo_selectors:
                try:
                    elementos = driver.find_elements(By.XPATH, selector)
                    for elemento in elementos:
                        if elemento.is_displayed():
                            texto = elemento.text.strip().upper()
                            if nombre_peaje in texto and 'VALOR' not in texto:
                                titulo_element = elemento
                                break
                    if titulo_element:
                        break
                except:
                    continue
            
            if not titulo_element:
                peajes[nombre_peaje] = None
                continue
            
            # Estrategia 1
            try:
                container = titulo_element.find_element(By.XPATH, "./ancestor::*[position()<=3]")
                valor_pagar_elements = container.find_elements(By.XPATH, ".//*[contains(text(), 'VALOR A PAGAR') or contains(text(), 'Valor a pagar')]")
                
                if valor_pagar_elements:
                    numeric_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '$') or contains(text(), ',') or contains(text(), '.')]")
                    
                    for elem in numeric_elements:
                        texto = elem.text.strip()
                        if texto and any(char.isdigit() for char in texto):
                            if 'VALOR A PAGAR' not in texto.upper() and 'COMERCIO' not in texto.upper():
                                peajes[nombre_peaje] = texto
                                break
            except:
                pass
            
            # Estrategia 2
            if nombre_peaje not in peajes or peajes[nombre_peaje] is None:
                try:
                    following_elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{nombre_peaje}')]/following::*")
                    
                    for elem in following_elements[:15]:
                        texto = elem.text.strip()
                        if texto and any(char.isdigit() for char in texto):
                            if len(texto) > 5 and len(texto) < 50:
                                peajes[nombre_peaje] = texto
                                break
                except:
                    pass
            
            # Estrategia 3
            if nombre_peaje not in peajes or peajes[nombre_peaje] is None:
                try:
                    parent = titulo_element.find_element(By.XPATH, "./..")
                    siblings = parent.find_elements(By.XPATH, "./*")
                    
                    for sibling in siblings:
                        if sibling != titulo_element:
                            texto = sibling.text.strip()
                            if texto and any(char.isdigit() for char in texto):
                                if len(texto) > 3 and len(texto) < 30:
                                    peajes[nombre_peaje] = texto
                                    break
                except:
                    pass
                    
        except Exception as e:
            peajes[nombre_peaje] = None
    
    return peajes

def find_resumen_comercios_table(driver):
    """Buscar tabla 'RESUMEN COMERCIOS' y extraer Cant Pasos por peaje"""
    try:
        st.info("🔍 Buscando tabla 'RESUMEN COMERCIOS'...")
        
        driver.save_screenshot("resumen_comercios_busqueda.png")
        
        page_text = driver.find_element(By.TAG_NAME, "body").text
        lines = page_text.split('\n')
        
        pasos_data = {}
        
        st.info("📋 Analizando contenido de la página...")
        
        for i, line in enumerate(lines):
            line_upper = line.strip().upper()
            
            if 'CHICORAL' in line_upper and 'CHICORAL' not in pasos_data:
                for offset in range(4):
                    if i + offset < len(lines):
                        search_line = lines[i + offset]
                        matches = re.findall(r'\b\d{1,2}\.?\d{3}\b|\b\d{3,4}\b', search_line)
                        if matches:
                            for match in matches:
                                clean_num = match.replace('.', '')
                                if clean_num.isdigit() and 100 <= int(clean_num) <= 100000:
                                    pasos_data['CHICORAL'] = match
                                    st.success(f"✅ CHICORAL Pasos: {match}")
                                    break
                            if 'CHICORAL' in pasos_data:
                                break
            
            elif 'COCORA' in line_upper and 'COCORA' not in pasos_data:
                for offset in range(4):
                    if i + offset < len(lines):
                        search_line = lines[i + offset]
                        matches = re.findall(r'\b\d{1,2}\.?\d{3}\b|\b\d{3,4}\b', search_line)
                        if matches:
                            for match in matches:
                                clean_num = match.replace('.', '')
                                if clean_num.isdigit() and 100 <= int(clean_num) <= 100000:
                                    pasos_data['COCORA'] = match
                                    st.success(f"✅ COCORA Pasos: {match}")
                                    break
                            if 'COCORA' in pasos_data:
                                break
            
            elif 'GUALANDAY' in line_upper and 'GUALANDAY' not in pasos_data:
                for offset in range(4):
                    if i + offset < len(lines):
                        search_line = lines[i + offset]
                        matches = re.findall(r'\b\d{1,2}\.?\d{3}\b|\b\d{3,4}\b', search_line)
                        if matches:
                            for match in matches:
                                clean_num = match.replace('.', '')
                                if clean_num.isdigit() and 100 <= int(clean_num) <= 100000:
                                    pasos_data['GUALANDAY'] = match
                                    st.success(f"✅ GUALANDAY Pasos: {match}")
                                    break
                            if 'GUALANDAY' in pasos_data:
                                break
        
        # Estrategia 2: XPath
        if len(pasos_data) < 3:
            st.info("🔄 Intentando búsqueda con XPath...")
            
            peajes_to_find = [p for p in ['CHICORAL', 'COCORA', 'GUALANDAY'] if p not in pasos_data]
            
            for peaje in peajes_to_find:
                try:
                    xpath_peaje = f"//*[contains(text(), '{peaje}')]"
                    elementos = driver.find_elements(By.XPATH, xpath_peaje)
                    
                    for elem in elementos:
                        if elem.is_displayed():
                            try:
                                parent = elem.find_element(By.XPATH, "./..")
                                parent_text = parent.text
                                
                                matches = re.findall(r'\b\d{1,2}\.?\d{3}\b|\b\d{3,4}\b', parent_text)
                                
                                for match in matches:
                                    clean_num = match.replace('.', '')
                                    if clean_num.isdigit() and 100 <= int(clean_num) <= 100000:
                                        pasos_data[peaje] = match
                                        st.success(f"✅ {peaje} Pasos (XPath): {match}")
                                        break
                                
                                if peaje in pasos_data:
                                    break
                            except:
                                continue
                except:
                    continue
        
        # Estrategia 3: Tablas
        if len(pasos_data) < 3:
            st.info("🔄 Buscando en elementos de tabla...")
            
            try:
                tables = driver.find_elements(By.XPATH, "//table | //*[@role='grid'] | //*[contains(@class, 'pivotTable')]")
                
                for table in tables:
                    if not table.is_displayed():
                        continue
                    
                    table_text = table.text
                    
                    if any(word in table_text.upper() for word in ['RESUMEN', 'COMERCIOS', 'CHICORAL', 'COCORA', 'GUALANDAY']):
                        st.info("📊 Tabla candidata encontrada")
                        
                        rows = table.find_elements(By.XPATH, ".//tr | .//*[@role='row']")
                        
                        for row in rows:
                            row_text = row.text.strip()
                            row_upper = row_text.upper()
                            
                            peajes_to_check = [p for p in ['CHICORAL', 'COCORA', 'GUALANDAY'] if p not in pasos_data]
                            
                            for peaje in peajes_to_check:
                                if peaje in row_upper:
                                    matches = re.findall(r'\b\d{1,2}\.?\d{3}\b|\b\d{3,4}\b', row_text)
                                    
                                    for match in matches:
                                        clean_num = match.replace('.', '')
                                        if clean_num.isdigit() and 100 <= int(clean_num) <= 100000:
                                            pasos_data[peaje] = match
                                            st.success(f"✅ {peaje} Pasos (tabla): {match}")
                                            break
            except Exception as e:
                st.warning(f"⚠️ Error en búsqueda de tabla: {e}")
        
        # Calcular total
        if pasos_data:
            try:
                total = 0
                for peaje, valor in pasos_data.items():
                    if peaje != 'TOTAL':
                        clean_valor = str(valor).replace('.', '').replace(',', '')
                        if clean_valor.isdigit():
                            total += int(clean_valor)
                
                if total > 0:
                    pasos_data['TOTAL'] = f"{total:,}".replace(",", ".")
                    st.success(f"✅ Total Pasos calculado: {pasos_data['TOTAL']}")
            except Exception as e:
                st.warning(f"⚠️ Error calculando total: {e}")
        
        if pasos_data:
            st.success(f"✅ Extracción exitosa: {len(pasos_data)} valores encontrados")
            return pasos_data
        else:
            st.warning("⚠️ No se encontraron datos de pasos")
            return None
            
    except Exception as e:
        st.error(f"❌ Error en find_resumen_comercios_table: {str(e)}")
        return None

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiYTFmOWZkMDAtY2IwYi00OTg4LWIxZDctNGZmYmU0NTMxNGI1IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None
    
    try:
        with st.spinner("🌐 Conectando con Power BI..."):
            driver.get(REPORT_URL)
            time.sleep(10)
        
        driver.save_screenshot("powerbi_inicial.png")
        
        if not click_conciliacion_date(driver, fecha_objetivo):
            return None
        
        time.sleep(3)
        driver.save_screenshot("powerbi_despues_seleccion.png")
        
        valor_texto = find_valor_a_pagar_comercio_card(driver)
        
        cantidad_pasos_texto = find_cantidad_pasos_card(driver)
        
        pasos_por_peaje = find_resumen_comercios_table(driver)
        
        valores_peajes = find_peaje_values(driver)
        
        driver.save_screenshot("powerbi_final.png")
        
        return {
            'valor_texto': valor_texto,
            'cantidad_pasos_texto': cantidad_pasos_texto or 'No encontrado',
            'pasos_por_peaje': pasos_por_peaje or {},
            'valores_peajes': valores_peajes,
            'screenshots': {
                'inicial': 'powerbi_inicial.png',
                'seleccion': 'powerbi_despues_seleccion.png',
                'final': 'powerbi_final.png'
            }
        }
        
    except Exception as e:
        st.error(f"❌ Error durante la extracción: {str(e)}")
        return None
    finally:
        driver.quit()

# ===== FUNCIONES DE EXTRACCIÓN DE EXCEL =====

def extract_excel_values(uploaded_file):
    """Extraer valores de las 3 hojas del Excel"""
    try:
        hojas = ['CHICORAL', 'GUALANDAY', 'COCORA']
        valores = {}
        total_general = 0
        
        for hoja in hojas:
            try:
                df = pd.read_excel(uploaded_file, sheet_name=hoja, header=None)
                
                valor_encontrado = None
                mejor_candidato = None
                mejor_puntaje = -1
                
                for i in range(len(df)-1, -1, -1):
                    fila = df.iloc[i]
                    
                    for j, celda in enumerate(fila):
                        if pd.notna(celda) and isinstance(celda, str) and 'TOTAL' in celda.upper().strip():
                            
                            for k in range(len(fila)):
                                posible_valor = fila.iloc[k]
                                if pd.notna(posible_valor):
                                    valor_str = str(posible_valor)
                                    
                                    puntaje = 0
                                    if ' in valor_str:
                                        puntaje += 10
                                    if any(c.isdigit() for c in valor_str):
                                        puntaje += 5
                                    if '.' in valor_str and len(valor_str.split('.')[-1]) == 3:
                                        puntaje += 3
                                    if len(valor_str) > 6:
                                        puntaje += 2
                                    
                                    if puntaje > 0 and len(valor_str) < 4:
                                        puntaje = 0
                                    if 'pag' in valor_str.lower():
                                        puntaje = 0
                                    
                                    if puntaje > mejor_puntaje:
                                        mejor_puntaje = puntaje
                                        mejor_candidato = posible_valor
                
                if mejor_candidato is not None and mejor_puntaje >= 5:
                    valor_encontrado = mejor_candidato
                else:
                    for i in range(len(df)-1, max(len(df)-11, -1), -1):
                        fila = df.iloc[i]
                        
                        for j, celda in enumerate(fila):
                            if pd.notna(celda) and isinstance(celda, str) and 'TOTAL' in celda.upper().strip():
                                for offset in [18, 17, 16, 19, 15]:
                                    if len(fila) > offset:
                                        valor_col = fila.iloc[offset]
                                        if pd.notna(valor_col):
                                            valor_str = str(valor_col)
                                            if (any(c.isdigit() for c in valor_str) and 
                                                len(valor_str) > 4 and 
                                                (' in valor_str or '.' in valor_str)):
                                                valor_encontrado = valor_col
                                                break
                                if valor_encontrado is not None:
                                    break
                        if valor_encontrado is not None:
                            break
                
                if valor_encontrado is not None:
                    valor_original = str(valor_encontrado)
                    valor_limpio = re.sub(r'[^\d.,]', '', valor_original)
                    
                    try:
                        if '.' in valor_limpio:
                            valor_limpio = valor_limpio.replace('.', '')
                        if ',' in valor_limpio:
                            partes = valor_limpio.split(',')
                            if len(partes) == 2 and len(partes[1]) == 2:
                                valor_limpio = partes[0] + '.' + partes[1]
                            else:
                                valor_limpio = valor_limpio.replace(',', '')
                        
                        valor_numerico = float(valor_limpio)
                        
                        if valor_numerico >= 1000:
                            valores[hoja] = valor_numerico
                            total_general += valor_numerico
                        else:
                            valores[hoja] = 0
                            
                    except:
                        valores[hoja] = 0
                else:
                    valores[hoja] = 0
                    
            except:
                valores[hoja] = 0
        
        return valores, total_general
        
    except Exception as e:
        st.error(f"❌ Error procesando archivo Excel: {str(e)}")
        return {}, 0

# ===== FUNCIONES DE COMPARACIÓN =====

def convert_currency_to_float(currency_string):
    """Convierte string de moneda a float"""
    try:
        if isinstance(currency_string, (int, float)):
            return float(currency_string)
            
        if isinstance(currency_string, str):
            cleaned = currency_string.strip()
            cleaned = cleaned.replace(', '').replace(' ', '')
            
            if '.' in cleaned and ',' in cleaned:
                cleaned = cleaned.replace('.', '').replace(',', '.')
            elif '.' in cleaned and cleaned.count('.') > 1:
                cleaned = cleaned.replace('.', '')
            elif ',' in cleaned:
                if cleaned.count(',') == 2 and '.' in cleaned:
                    cleaned = cleaned.replace(',', '')
                elif cleaned.count(',') == 1:
                    cleaned = cleaned.replace(',', '.')
                else:
                    cleaned = cleaned.replace(',', '')
            
            return float(cleaned) if cleaned else 0.0
            
        return float(currency_string)
        
    except Exception as e:
        st.error(f"❌ Error convirtiendo moneda: '{currency_string}' - {e}")
        return 0.0

def compare_values(valor_powerbi, valor_excel):
    """Comparar valores de Power BI y Excel"""
    try:
        if isinstance(valor_powerbi, dict):
            valor_powerbi_texto = valor_powerbi.get('valor_texto', '')
            powerbi_numero = convert_currency_to_float(valor_powerbi_texto)
        else:
            valor_powerbi_texto = str(valor_powerbi)
            powerbi_numero = convert_currency_to_float(valor_powerbi)
            
        excel_numero = float(valor_excel)
        
        tolerancia = 0.01
        coinciden = abs(powerbi_numero - excel_numero) <= tolerancia
        
        return powerbi_numero, excel_numero, valor_powerbi_texto, coinciden
        
    except Exception as e:
        st.error(f"❌ Error comparando valores: {e}")
        return None, None, str(valor_powerbi), False

def compare_peajes(valores_powerbi_peajes, valores_excel):
    """Comparar valores individuales por peaje"""
    comparaciones = {}
    
    for peaje in ['CHICORAL', 'COCORA', 'GUALANDAY']:
        try:
            valor_powerbi_texto = valores_powerbi_peajes.get(peaje)
            
            if valor_powerbi_texto is None:
                comparaciones[peaje] = {
                    'powerbi_texto': 'No encontrado',
                    'powerbi_numero': 0,
                    'excel_numero': valores_excel.get(peaje, 0),
                    'coinciden': False,
                    'diferencia': valores_excel.get(peaje, 0)
                }
                continue
            
            powerbi_numero = convert_currency_to_float(valor_powerbi_texto)
            excel_numero = valores_excel.get(peaje, 0)
            
            tolerancia = 0.01
            coinciden = abs(powerbi_numero - excel_numero) <= tolerancia
            diferencia = abs(powerbi_numero - excel_numero)
            
            comparaciones[peaje] = {
                'powerbi_texto': valor_powerbi_texto,
                'powerbi_numero': powerbi_numero,
                'excel_numero': excel_numero,
                'coinciden': coinciden,
                'diferencia': diferencia
            }
            
        except Exception as e:
            st.error(f"❌ Error comparando {peaje}: {e}")
            comparaciones[peaje] = {
                'powerbi_texto': 'Error',
                'powerbi_numero': 0,
                'excel_numero': valores_excel.get(peaje, 0),
                'coinciden': False,
                'diferencia': 0
            }
    
    return comparaciones

# ===== INTERFAZ PRINCIPAL =====

def main():
    st.title("💰 Validador Power BI - Conciliaciones APP GICA")
    st.markdown("---")
    
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Cargar archivo Excel con 3 hojas
    - Extraer valores de CHICORAL, GUALANDAY, COCORA
    - Calcular total automáticamente
    - Comparar con Power BI (Total y por Peaje)
    - Extraer cantidad de pasos por peaje
    
    **Estado:** ✅ ChromeDriver Compatible
    **Versión:** v2.3 - Código Corregido
    """)
    
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}")
    st.sidebar.info(f"✅ Pandas {pd.__version__}")
    st.sidebar.info(f"✅ Streamlit {st.__version__}")
    
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel con hojas CHICORAL, GUALANDAY, COCORA", 
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        fecha_desde_archivo = None
        try:
            patron_fecha = r'(\d{4})-(\d{2})-(\d{2})'
            match = re.search(patron_fecha, uploaded_file.name)
            
            if match:
                year, month, day = match.groups()
                fecha_desde_archivo = pd.to_datetime(f"{year}-{month}-{day}")
        except:
            pass
        
        with st.spinner("📊 Procesando archivo Excel..."):
            valores, total_general = extract_excel_values(uploaded_file)
        
        if total_general > 0:
            st.markdown("### 📊 Valores Extraídos del Excel")
            
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
                st.metric("TOTAL GENERAL", total_formateado, delta="Excel")
            
            st.markdown("---")
            
            if fecha_desde_archivo:
                st.info(f"🤖 **Extracción Automática Activada** | Fecha: {fecha_desde_archivo.strftime('%Y-%m-%d')}")
                fecha_objetivo = fecha_desde_archivo.strftime("%Y-%m-%d")
                ejecutar_extraccion = True
            else:
                st.subheader("📅 Parámetros de Búsqueda")
                fecha_conciliacion = st.date_input(
                    "Fecha de Conciliación",
                    value=pd.to_datetime("2025-09-04"),
                    help="No se pudo detectar la fecha del archivo. Ingresa manualmente."
                )
                fecha_objetivo = fecha_conciliacion.strftime("%Y-%m-%d")
                
                if st.button("🎯 Extraer Valores de Power BI y Comparar", type="primary", use_container_width=True):
                    ejecutar_extraccion = True
                else:
                    ejecutar_extraccion = False
            
            if ejecutar_extraccion:
                with st.spinner("🌐 Extrayendo datos de Power BI... Esto puede tomar 1-2 minutos"):
                    resultados = extract_powerbi_data(fecha_objetivo)
                    
                    if resultados and resultados.get('valor_texto'):
                        valor_powerbi_texto = resultados['valor_texto']
                        cantidad_pasos_texto = resultados.get('cantidad_pasos_texto', 'No encontrado')
                        pasos_por_peaje = resultados.get('pasos_por_peaje', {})
                        valores_peajes_powerbi = resultados.get('valores_peajes', {})
                        
                        st.markdown("---")
                        
                        st.markdown("### 📊 Valores Extraídos de Power BI")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.metric("💰 VALOR A PAGAR A COMERCIO", valor_powerbi_texto)
                        
                        with col2:
                            st.metric("👣 CANTIDAD DE PASOS BI", cantidad_pasos_texto)
                        
                        if pasos_por_peaje:
                            st.markdown("#### 👣 Cantidad de Pasos por Peaje (RESUMEN COMERCIOS)")
                            
                            tabla_pasos_data = []
                            for peaje in ['CHICORAL', 'COCORA', 'GUALANDAY']:
                                if peaje in pasos_por_peaje:
                                    tabla_pasos_data.append({
                                        'Peaje': peaje,
                                        'Cantidad de Pasos': pasos_por_peaje[peaje]
                                    })
                            
                            if 'TOTAL' in pasos_por_peaje:
                                tabla_pasos_data.append({
                                    'Peaje': 'TOTAL',
                                    'Cantidad de Pasos': pasos_por_peaje['TOTAL']
                                })
                            
                            if tabla_pasos_data:
                                df_pasos = pd.DataFrame(tabla_pasos_data)
                                st.dataframe(df_pasos, use_container_width=True, hide_index=True)
                        else:
                            st.warning("⚠️ No se pudieron extraer los pasos por peaje de la tabla 'RESUMEN COMERCIOS'")
                        
                        st.markdown("---")
                        
                        st.markdown("### 💰 Validación: Total General")
                        
                        powerbi_numero, excel_numero, valor_formateado, coinciden = compare_values(
                            resultados, 
                            total_general
                        )
                        
                        if powerbi_numero is not None and excel_numero is not None:
                            col1, col2, col3 = st.columns([2, 2, 1])
                            
                            with col1:
                                st.metric("📊 Power BI", valor_formateado)
                            with col2:
                                st.metric("📁 Excel", total_formateado)
                            with col3:
                                if coinciden:
                                    st.markdown("#### ✅")
                                    st.success("COINCIDE")
                                else:
                                    diferencia = abs(powerbi_numero - excel_numero)
                                    st.markdown("#### ❌")
                                    st.error("DIFERENCIA")
                                    st.caption(f"${diferencia:,.0f}".replace(",", "."))
                        
                        st.markdown("---")
                        
                        st.markdown("### 🏢 Validación: Por Peaje")
                        
                        comparaciones_peajes = compare_peajes(valores_peajes_powerbi, valores)
                        
                        tabla_data = []
                        todos_coinciden = True
                        
                        for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                            comp = comparaciones_peajes[peaje]
                            
                            estado_icono = "✅" if comp['coinciden'] else "❌"
                            diferencia_texto = "$0" if comp['coinciden'] else f"${comp['diferencia']:,.0f}".replace(",", ".")
                            
                            tabla_data.append({
                                '': estado_icono,
                                'Peaje': peaje,
                                'Power BI': comp['powerbi_texto'],
                                'Excel': f"${comp['excel_numero']:,.0f}".replace(",", "."),
                                'Dif.': diferencia_texto
                            })
                            
                            if not comp['coinciden']:
                                todos_coinciden = False
                        
                        df_comparacion = pd.DataFrame(tabla_data)
                        st.dataframe(df_comparacion, use_container_width=True, hide_index=True)
                        
                        st.markdown("---")
                        
                        st.markdown("### 📋 Resultado Final")
                        
                        if coinciden and todos_coinciden:
                            st.success("🎉 **VALIDACIÓN EXITOSA** - Todos los valores coinciden")
                            st.balloons()
                        elif coinciden and not todos_coinciden:
                            st.warning("⚠️ **VALIDACIÓN PARCIAL** - El total coincide, pero hay diferencias por peaje")
                        elif not coinciden and todos_coinciden:
                            st.warning("⚠️ **VALIDACIÓN PARCIAL** - Los peajes coinciden, pero el total tiene diferencias")
                        else:
                            st.error("❌ **VALIDACIÓN FALLIDA** - Existen diferencias en total y peajes")
                        
                        with st.expander("🔍 Ver Detalles Completos y Capturas"):
                            st.markdown("#### 📊 Tabla Detallada")
                            resumen_data = []
                            
                            resumen_data.append({
                                'Concepto': 'TOTAL GENERAL',
                                'Power BI': f"${powerbi_numero:,.0f}".replace(",", "."),
                                'Excel': f"${excel_numero:,.0f}".replace(",", "."),
                                'Estado': '✅ Coincide' if coinciden else '❌ No coincide',
                                'Diferencia': f"${abs(powerbi_numero - excel_numero):,.0f}".replace(",", "."),
                                'Dif. %': f"{abs(powerbi_numero - excel_numero)/excel_numero*100:.2f}%" if excel_numero > 0 else "N/A"
                            })
                            
                            resumen_data.append({
                                'Concepto': 'CANTIDAD DE PASOS',
                                'Power BI': cantidad_pasos_texto,
                                'Excel': 'N/A',
                                'Estado': 'ℹ️ Solo Power BI',
                                'Diferencia': 'N/A',
                                'Dif. %': 'N/A'
                            })
                            
                            if pasos_por_peaje:
                                for peaje in ['CHICORAL', 'COCORA', 'GUALANDAY']:
                                    if peaje in pasos_por_peaje:
                                        resumen_data.append({
                                            'Concepto': f'PASOS {peaje}',
                                            'Power BI': pasos_por_peaje[peaje],
                                            'Excel': 'N/A',
                                            'Estado': 'ℹ️ Solo Power BI',
                                            'Diferencia': 'N/A',
                                            'Dif. %': 'N/A'
                                        })
                                
                                if 'TOTAL' in pasos_por_peaje:
                                    resumen_data.append({
                                        'Concepto': 'TOTAL PASOS',
                                        'Power BI': pasos_por_peaje['TOTAL'],
                                        'Excel': 'N/A',
                                        'Estado': 'ℹ️ Solo Power BI',
                                        'Diferencia': 'N/A',
                                        'Dif. %': 'N/A'
                                    })
                            
                            for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                                comp = comparaciones_peajes[peaje]
                                excel_val = comp['excel_numero']
                                resumen_data.append({
                                    'Concepto': peaje,
                                    'Power BI': comp['powerbi_texto'],
                                    'Excel': f"${comp['excel_numero']:,.0f}".replace(",", "."),
                                    'Estado': '✅ Coincide' if comp['coinciden'] else '❌ No coincide',
                                    'Diferencia': f"${comp['diferencia']:,.0f}".replace(",", "."),
                                    'Dif. %': f"{comp['diferencia']/excel_val*100:.2f}%" if excel_val > 0 else "N/A"
                                })
                            
                            df_resumen = pd.DataFrame(resumen_data)
                            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
                            
                            st.markdown("#### 📸 Capturas del Proceso")
                            col1, col2, col3 = st.columns(3)
                            screenshots = resultados.get('screenshots', {})
                            
                            if 'inicial' in screenshots and os.path.exists(screenshots['inicial']):
                                with col1:
                                    st.image(screenshots['inicial'], caption="Vista Inicial", use_column_width=True)
                            
                            if 'seleccion' in screenshots and os.path.exists(screenshots['seleccion']):
                                with col2:
                                    st.image(screenshots['seleccion'], caption="Tras Selección", use_column_width=True)
                            
                            if 'final' in screenshots and os.path.exists(screenshots['final']):
                                with col3:
                                    st.image(screenshots['final'], caption="Vista Final", use_column_width=True)
                                
                    elif resultados:
                        st.error("❌ Se accedió al reporte pero no se encontró el valor específico")
                    else:
                        st.error("❌ No se pudieron extraer datos del reporte Power BI")
        else:
            st.error("❌ No se pudieron extraer valores del archivo Excel")
            with st.expander("💡 Sugerencias para solucionar el problema"):
                st.markdown("""
                - Verifica que las hojas se llamen **CHICORAL**, **GUALANDAY**, **COCORA**
                - Asegúrate de que haya valores numéricos en las celdas de total
                - Revisa que los totales estén claramente identificados con **'TOTAL'**
                """)
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar la validación")

    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso:**
        1. **Cargar Excel**: Archivo con hojas CHICORAL, GUALANDAY, COCORA
        2. **Extracción automática**: Búsqueda inteligente de "Total" en cada hoja
        3. **Seleccionar fecha** de conciliación en Power BI  
        4. **Comparar**: Extrae valores de Power BI y compara con Excel
        
        **Características (v2.3):**
        - ✅ **Comparación Total**: Valida el "VALOR A PAGAR A COMERCIO" total
        - ✅ **Cantidad de Pasos**: Extrae y muestra "CANTIDAD PASOS" del Power BI
        - ✅ **Pasos por Peaje**: Extrae cantidad de pasos de tabla "RESUMEN COMERCIOS"
        - ✅ **Comparación por Peaje**: Valida valores individuales
        - ✅ **Resumen Detallado**: Tabla completa con todas las comparaciones
        - 📸 **Capturas del proceso**: Para verificación y debugging
        """)

if __name__ == "__main__":
    main()
    
    st.markdown("---")
    st.markdown('<div class="footer">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v2.3</div>', unsafe_allow_html=True)
