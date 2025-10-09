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

/* ===== ESTILOS PARA TABLAS ===== */
.dataframe {
    font-size: 14px !important;
}
.table-result {
    margin: 10px 0;
    border-radius: 8px;
    overflow: hidden;
}
.table-header {
    background-color: #1E1E2F;
    color: white;
    padding: 12px;
    font-weight: bold;
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

# ===== FUNCIONES DE EXTRACCIÓN DE POWER BI (OPTIMIZADAS) =====

def setup_driver():
    """Configurar ChromeDriver para Selenium"""
    try:
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            return driver
        except Exception as e:
            st.error(f"❌ Error al configurar ChromeDriver: {e}")
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
    """Buscar la tarjeta 'CANTIDAD PASOS'"""
    try:
        titulo_selectors = [
            "//*[contains(text(), 'CANTIDAD PASOS')]",
            "//*[contains(text(), 'Cantidad Pasos')]",
            "//*[contains(text(), 'CANTIDAD DE PASOS')]",
            "//*[contains(text(), 'Cantidad de Pasos')]",
            "//*[contains(text(), 'CANTIDAD') and contains(text(), 'PASOS')]",
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        if any(palabra in texto.upper() for palabra in ['CANTIDAD', 'PASOS']):
                            titulo_element = elemento
                            break
                if titulo_element:
                    break
            except Exception as e:
                continue
        
        if not titulo_element:
            return None
        
        # Buscar valor numérico
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            all_elements = container.find_elements(By.XPATH, ".//*")
            
            for elem in all_elements:
                texto = elem.text.strip()
                if (texto and 
                    any(char.isdigit() for char in texto) and 
                    len(texto) < 20 and 
                    texto != titulo_element.text and
                    not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
                    
                    digit_count = sum(char.isdigit() for char in texto)
                    if digit_count >= 1:
                        return texto
                        
        except Exception as e:
            pass
        
        # Estrategias alternativas...
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
            return None
        
        # Buscar valor numérico
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            numeric_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '$') or contains(text(), ',') or contains(text(), '.')]")
            
            for elem in numeric_elements:
                texto = elem.text.strip()
                if texto and any(char.isdigit() for char in texto) and texto != titulo_element.text:
                    return texto
        except:
            pass
        
        return None
        
    except Exception as e:
        return None

def find_peaje_values(driver):
    """Buscar valores individuales de cada peaje"""
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
            
            # Buscar valor
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
                    
        except Exception as e:
            peajes[nombre_peaje] = None
    
    return peajes

def find_pasos_por_peaje_bi(driver):
    """Buscar cantidad de pasos por peaje en la tabla RESUMEN COMERCIOS"""
    try:
        pasos_peajes = {}
        nombres_peajes = ['CHICORAL', 'COCORA', 'GUALANDAY']
        total_pasos_bi = 0
        
        # Buscar la tabla
        tabla_selectors = [
            "//*[contains(text(), 'RESUMEN COMERCIOS')]",
            "//*[contains(text(), 'Resumen Comercios')]",
            "//*[contains(text(), 'RESUMEN') and contains(text(), 'COMERCIOS')]",
        ]
        
        tabla_element = None
        for selector in tabla_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        tabla_element = elemento
                        break
                if tabla_element:
                    break
            except:
                continue
        
        if not tabla_element:
            return {}, 0
        
        # Buscar en el contenedor
        try:
            container = tabla_element.find_element(By.XPATH, "./ancestor::*[position()<=5]")
            
            for nombre_peaje in nombres_peajes:
                peaje_selectors = [
                    f".//*[contains(text(), '{nombre_peaje}')]",
                    f".//*[contains(text(), '{nombre_peaje.upper()}')]",
                    f".//*[contains(text(), '{nombre_peaje.lower()}')]",
                ]
                
                peaje_element = None
                for selector in peaje_selectors:
                    try:
                        elementos = container.find_elements(By.XPATH, selector)
                        for elemento in elementos:
                            if elemento.is_displayed():
                                texto = elemento.text.strip().upper()
                                if nombre_peaje in texto:
                                    peaje_element = elemento
                                    break
                        if peaje_element:
                            break
                    except:
                        continue
                
                if peaje_element:
                    try:
                        fila_element = peaje_element.find_element(By.XPATH, "./ancestor::*[position()<=3]")
                        numeric_elements = fila_element.find_elements(By.XPATH, ".//*[text()]")
                        
                        for elem in numeric_elements:
                            texto = elem.text.strip()
                            if (texto and 
                                any(char.isdigit() for char in texto) and
                                1 <= len(texto) <= 6 and
                                texto != peaje_element.text and
                                not any(word in texto.upper() for word in ['CHICORAL', 'COCORA', 'GUALANDAY', 'TOTAL', 'RESUMEN'])):
                                
                                pasos_limpio = re.sub(r'[^\d]', '', texto)
                                if pasos_limpio and pasos_limpio.isdigit():
                                    num_pasos = int(pasos_limpio)
                                    if 1 <= num_pasos <= 999999:
                                        pasos_peajes[nombre_peaje] = num_pasos
                                        total_pasos_bi += num_pasos
                                        break
                        
                    except Exception as e:
                        pasos_peajes[nombre_peaje] = 0
                else:
                    pasos_peajes[nombre_peaje] = 0
            
            return pasos_peajes, total_pasos_bi
            
        except Exception as e:
            return {}, 0
            
    except Exception as e:
        return {}, 0

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiYTFmOWZkMDAtY2IwYi00OTg4LWIxZDctNGZmYmU0NTMxNGI1IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None
    
    try:
        # Navegar al reporte
        with st.spinner("🌐 Conectando con Power BI..."):
            driver.get(REPORT_URL)
            time.sleep(10)
        
        # Hacer clic en la conciliación específica
        if not click_conciliacion_date(driver, fecha_objetivo):
            return None
        
        # Esperar a que cargue
        time.sleep(3)
        
        # Buscar valores
        valor_texto = find_valor_a_pagar_comercio_card(driver)
        cantidad_pasos_texto = find_cantidad_pasos_card(driver)
        valores_peajes = find_peaje_values(driver)
        pasos_peajes_bi, total_pasos_bi = find_pasos_por_peaje_bi(driver)
        
        return {
            'valor_texto': valor_texto,
            'cantidad_pasos_texto': cantidad_pasos_texto or 'No encontrado',
            'valores_peajes': valores_peajes,
            'pasos_peajes_bi': pasos_peajes_bi,
            'total_pasos_bi': total_pasos_bi,
        }
        
    except Exception as e:
        st.error(f"❌ Error durante la extracción: {str(e)}")
        return None
    finally:
        driver.quit()

# ===== FUNCIONES DE EXTRACCIÓN DE EXCEL (OPTIMIZADAS) =====

def extract_excel_values(uploaded_file):
    """Extraer valores monetarios Y cantidad de pasos del Excel"""
    try:
        hojas = ['CHICORAL', 'GUALANDAY', 'COCORA']
        valores = {}
        pasos = {}
        total_general = 0
        total_pasos = 0
        
        for hoja in hojas:
            try:
                df = pd.read_excel(uploaded_file, sheet_name=hoja, header=None)
                
                valor_encontrado = None
                pasos_encontrados = None
                
                # Buscar en las últimas filas
                for i in range(len(df)-1, max(len(df)-15, -1), -1):
                    fila = df.iloc[i]
                    
                    for j, celda in enumerate(fila):
                        if pd.notna(celda):
                            texto = str(celda).strip()
                            
                            # Buscar valores monetarios
                            if (any(char.isdigit() for char in texto) and 
                                len(texto) > 4 and
                                ('$' in texto or ',' in texto or '.' in texto)):
                                
                                try:
                                    valor_limpio = re.sub(r'[^\d.,]', '', texto)
                                    if '.' in valor_limpio and ',' in valor_limpio:
                                        valor_limpio = valor_limpio.replace('.', '').replace(',', '.')
                                    valor_numerico = float(valor_limpio)
                                    if valor_numerico >= 1000:
                                        valor_encontrado = valor_numerico
                                        break
                                except:
                                    continue
                            
                            # Buscar pasos
                            elif (any(char.isdigit() for char in texto) and 
                                  3 <= len(texto) <= 6 and
                                  not any(word in texto.upper() for word in ['TOTAL', 'VALOR', '$', 'PAGAR', 'COMERCIO'])):
                                
                                digit_count = sum(char.isdigit() for char in texto)
                                if digit_count >= len(texto) * 0.7:
                                    pasos_limpio = re.sub(r'[^\d]', '', texto)
                                    if pasos_limpio and pasos_limpio.isdigit():
                                        num_pasos = int(pasos_limpio)
                                        if 100 <= num_pasos <= 99999:
                                            pasos_encontrados = num_pasos
                    
                    if valor_encontrado is not None and pasos_encontrados is not None:
                        break
                
                # Asignar valores
                if valor_encontrado is not None:
                    valores[hoja] = valor_encontrado
                    total_general += valor_encontrado
                else:
                    valores[hoja] = 0
                
                if pasos_encontrados is not None:
                    pasos[hoja] = pasos_encontrados
                    total_pasos += pasos_encontrados
                else:
                    pasos[hoja] = 0
                    
            except Exception as e:
                valores[hoja] = 0
                pasos[hoja] = 0
        
        return valores, total_general, pasos, total_pasos
        
    except Exception as e:
        st.error(f"❌ Error procesando archivo Excel: {str(e)}")
        return {}, 0, {}, 0

# ===== FUNCIONES DE COMPARACIÓN (OPTIMIZADAS) =====

def convert_currency_to_float(currency_string):
    """Convierte string de moneda a float"""
    try:
        if isinstance(currency_string, (int, float)):
            return float(currency_string)
            
        if isinstance(currency_string, str):
            cleaned = currency_string.strip()
            cleaned = cleaned.replace('$', '').replace(' ', '')
            
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
            comparaciones[peaje] = {
                'powerbi_texto': 'Error',
                'powerbi_numero': 0,
                'excel_numero': valores_excel.get(peaje, 0),
                'coinciden': False,
                'diferencia': 0
            }
    
    return comparaciones

def compare_pasos_peajes(pasos_peajes_bi, pasos_excel):
    """Comparar cantidad de pasos por peaje"""
    comparaciones = {}
    
    for peaje in ['CHICORAL', 'COCORA', 'GUALANDAY']:
        try:
            pasos_bi = pasos_peajes_bi.get(peaje, 0)
            pasos_excel_val = pasos_excel.get(peaje, 0)
            
            coinciden = pasos_bi == pasos_excel_val
            diferencia = abs(pasos_bi - pasos_excel_val)
            
            comparaciones[peaje] = {
                'pasos_bi': pasos_bi,
                'pasos_excel': pasos_excel_val,
                'coinciden': coinciden,
                'diferencia': diferencia
            }
            
        except Exception as e:
            comparaciones[peaje] = {
                'pasos_bi': 0,
                'pasos_excel': pasos_excel.get(peaje, 0),
                'coinciden': False,
                'diferencia': 0
            }
    
    return comparaciones

# ===== INTERFAZ PRINCIPAL OPTIMIZADA =====

def main():
    st.title("💰 Validador Power BI - Conciliaciones APP GICA")
    st.markdown("---")
    
    # Sidebar
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Cargar archivo Excel con 3 hojas
    - Extraer valores Y pasos por peaje
    - Comparar con Power BI
    
    **Versión:** v2.3 - Con Pasos por Peaje
    """)
    
    # Cargar archivo Excel
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel con hojas CHICORAL, GUALANDAY, COCORA", 
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        # Extraer fecha del nombre del archivo
        fecha_desde_archivo = None
        try:
            patron_fecha = r'(\d{4})-(\d{2})-(\d{2})'
            match = re.search(patron_fecha, uploaded_file.name)
            if match:
                year, month, day = match.groups()
                fecha_desde_archivo = pd.to_datetime(f"{year}-{month}-{day}")
        except:
            pass
        
        # Extraer valores del Excel
        with st.spinner("📊 Procesando archivo Excel..."):
            valores, total_general, pasos, total_pasos = extract_excel_values(uploaded_file)
        
        if total_general > 0 or total_pasos > 0:
            # ========== TABLA RESUMEN EXCEL ==========
            st.markdown("### 📊 Valores Extraídos del Excel")
            
            # Crear tabla resumen Excel
            tabla_excel = []
            for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                tabla_excel.append({
                    'Peaje': peaje,
                    'Valor': f"${valores[peaje]:,.0f}".replace(",", "."),
                    'Pasos': f"{pasos[peaje]:,}".replace(",", ".")
                })
            
            # Agregar total
            tabla_excel.append({
                'Peaje': '**TOTAL**',
                'Valor': f"**${total_general:,.0f}**".replace(",", "."),
                'Pasos': f"**{total_pasos:,}**".replace(",", ".")
            })
            
            df_excel = pd.DataFrame(tabla_excel)
            st.dataframe(df_excel, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            
            # ========== EXTRACCIÓN POWER BI ==========
            if fecha_desde_archivo:
                st.info(f"🤖 **Extracción Automática Activada** | Fecha: {fecha_desde_archivo.strftime('%Y-%m-%d')}")
                fecha_objetivo = fecha_desde_archivo.strftime("%Y-%m-%d")
                ejecutar_extraccion = True
            else:
                st.subheader("📅 Parámetros de Búsqueda")
                fecha_conciliacion = st.date_input(
                    "Fecha de Conciliación",
                    value=pd.to_datetime("2025-09-04")
                )
                fecha_objetivo = fecha_conciliacion.strftime("%Y-%m-%d")
                
                if st.button("🎯 Extraer Valores de Power BI y Comparar", type="primary", use_container_width=True):
                    ejecutar_extraccion = True
                else:
                    ejecutar_extraccion = False
            
            if ejecutar_extraccion:
                with st.spinner("🌐 Extrayendo datos de Power BI..."):
                    resultados = extract_powerbi_data(fecha_objetivo)
                    
                    if resultados and resultados.get('valor_texto'):
                        valor_powerbi_texto = resultados['valor_texto']
                        cantidad_pasos_texto = resultados.get('cantidad_pasos_texto', 'No encontrado')
                        valores_peajes_powerbi = resultados.get('valores_peajes', {})
                        pasos_peajes_bi = resultados.get('pasos_peajes_bi', {})
                        total_pasos_bi = resultados.get('total_pasos_bi', 0)
                        
                        st.markdown("---")
                        
                        # ========== TABLA RESUMEN POWER BI ==========
                        st.markdown("### 📊 Valores Extraídos de Power BI")
                        
                        # Crear tabla resumen Power BI
                        tabla_powerbi = []
                        for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                            tabla_powerbi.append({
                                'Peaje': peaje,
                                'Valor': valores_peajes_powerbi.get(peaje, 'No encontrado'),
                                'Pasos': f"{pasos_peajes_bi.get(peaje, 0):,}".replace(",", ".")
                            })
                        
                        # Agregar totales
                        powerbi_numero, _, _, _ = compare_values(resultados, total_general)
                        tabla_powerbi.append({
                            'Peaje': '**TOTAL**',
                            'Valor': f"**{valor_powerbi_texto}**",
                            'Pasos': f"**{total_pasos_bi:,}**".replace(",", ".")
                        })
                        
                        df_powerbi = pd.DataFrame(tabla_powerbi)
                        st.dataframe(df_powerbi, use_container_width=True, hide_index=True)
                        
                        st.markdown("---")
                        
                        # ========== TABLA COMPARACIÓN COMPLETA ==========
                        st.markdown("### 📋 Tabla de Comparación Completa")
                        
                        # Comparar valores
                        powerbi_numero, excel_numero, _, coinciden_valor = compare_values(resultados, total_general)
                        
                        # Comparar pasos totales
                        cantidad_pasos_bi = total_pasos_bi
                        if cantidad_pasos_texto and cantidad_pasos_texto != 'No encontrado':
                            try:
                                pasos_limpio = re.sub(r'[^\d]', '', str(cantidad_pasos_texto))
                                if pasos_limpio:
                                    cantidad_pasos_bi = int(pasos_limpio)
                            except:
                                pass
                        
                        coinciden_pasos = cantidad_pasos_bi == total_pasos
                        
                        # Comparar por peaje
                        comparaciones_peajes = compare_peajes(valores_peajes_powerbi, valores)
                        comparaciones_pasos_peajes = compare_pasos_peajes(pasos_peajes_bi, pasos)
                        
                        # Crear tabla comparativa completa
                        tabla_comparativa = []
                        
                        # Totales
                        tabla_comparativa.append({
                            'Concepto': 'TOTAL GENERAL',
                            'Power BI': f"${powerbi_numero:,.0f}".replace(",", ".") if powerbi_numero else 'No encontrado',
                            'Excel': f"${total_general:,.0f}".replace(",", "."),
                            'Estado': '✅' if coinciden_valor else '❌',
                            'Diferencia': f"${abs(powerbi_numero - total_general):,.0f}".replace(",", ".") if powerbi_numero else 'N/A'
                        })
                        
                        tabla_comparativa.append({
                            'Concepto': 'TOTAL PASOS',
                            'Power BI': f"{cantidad_pasos_bi:,}".replace(",", "."),
                            'Excel': f"{total_pasos:,}".replace(",", "."),
                            'Estado': '✅' if coinciden_pasos else '❌',
                            'Diferencia': f"{abs(cantidad_pasos_bi - total_pasos):,}".replace(",", ".")
                        })
                        
                        # Valores por peaje
                        for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                            comp_valor = comparaciones_peajes[peaje]
                            tabla_comparativa.append({
                                'Concepto': f'VALOR {peaje}',
                                'Power BI': comp_valor['powerbi_texto'],
                                'Excel': f"${comp_valor['excel_numero']:,.0f}".replace(",", "."),
                                'Estado': '✅' if comp_valor['coinciden'] else '❌',
                                'Diferencia': f"${comp_valor['diferencia']:,.0f}".replace(",", ".")
                            })
                        
                        # Pasos por peaje
                        for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                            comp_pasos = comparaciones_pasos_peajes[peaje]
                            tabla_comparativa.append({
                                'Concepto': f'PASOS {peaje}',
                                'Power BI': f"{comp_pasos['pasos_bi']:,}".replace(",", "."),
                                'Excel': f"{comp_pasos['pasos_excel']:,}".replace(",", "."),
                                'Estado': '✅' if comp_pasos['coinciden'] else '❌',
                                'Diferencia': f"{comp_pasos['diferencia']:,}".replace(",", ".")
                            })
                        
                        df_comparativa = pd.DataFrame(tabla_comparativa)
                        st.dataframe(df_comparativa, use_container_width=True, hide_index=True)
                        
                        # ========== RESUMEN FINAL ==========
                        st.markdown("---")
                        st.markdown("### 📈 Resultado Final")
                        
                        # Calcular estado general
                        todos_valores_coinciden = all(comp['coinciden'] for comp in comparaciones_peajes.values())
                        todos_pasos_coinciden = all(comp['coinciden'] for comp in comparaciones_pasos_peajes.values())
                        
                        validacion_completa = (coinciden_valor and coinciden_pasos and 
                                             todos_valores_coinciden and todos_pasos_coinciden)
                        
                        if validacion_completa:
                            st.success("🎉 **VALIDACIÓN EXITOSA** - Todos los valores coinciden")
                            st.balloons()
                        else:
                            st.error("❌ **VALIDACIÓN CON DIFERENCIAS** - Revisar la tabla de comparación")
                            
                    else:
                        st.error("❌ No se pudieron extraer datos del reporte Power BI")
        else:
            st.error("❌ No se pudieron extraer valores del archivo Excel")
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar la validación")

if __name__ == "__main__":
    main()

    # Footer
    st.markdown("---")
    st.markdown('<div class="footer">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v2.3</div>', unsafe_allow_html=True)
