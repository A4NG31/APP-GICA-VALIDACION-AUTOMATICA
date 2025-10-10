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
    """Configurar ChromeDriver para Selenium - VERSIÓN COMPATIBLE"""
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
            st.error(f"Error al configurar ChromeDriver: {e}")
            return None
            
    except Exception as e:
        st.error(f"Error crítico al configurar ChromeDriver: {e}")
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
            st.error("No se encontró la conciliación para la fecha especificada")
            return False
            
    except Exception as e:
        st.error(f"Error al hacer clic en conciliación: {str(e)}")
        return False

def find_cantidad_pasos_card(driver):
    """Buscar la tarjeta 'CANTIDAD PASOS'"""
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
                        titulo_element = elemento
                        break
                if titulo_element:
                    break
            except:
                continue
        
        if not titulo_element:
            return None
        
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            all_elements = container.find_elements(By.XPATH, ".//*")
            
            for elem in all_elements:
                texto = elem.text.strip()
                if (texto and any(char.isdigit() for char in texto) and len(texto) < 20 and 
                    texto != titulo_element.text and
                    not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
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
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        titulo_element = elemento
                        break
                if titulo_element:
                    break
            except:
                continue
        
        if not titulo_element:
            return None
        
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            numeric_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '$') or contains(text(), ',')]")
            
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
    """Buscar valores individuales de cada peaje en Power BI"""
    peajes = {}
    nombres_peajes = ['CHICORAL', 'COCORA', 'GUALANDAY']
    
    for nombre_peaje in nombres_peajes:
        try:
            titulo_selectors = [
                f"//*[contains(text(), '{nombre_peaje}')]",
                f"//*[contains(text(), '{nombre_peaje.title()}')]",
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
            
            try:
                container = titulo_element.find_element(By.XPATH, "./ancestor::*[position()<=3]")
                numeric_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '$')]")
                
                for elem in numeric_elements:
                    texto = elem.text.strip()
                    if texto and any(char.isdigit() for char in texto):
                        peajes[nombre_peaje] = texto
                        break
            except:
                pass
                    
        except Exception as e:
            peajes[nombre_peaje] = None
    
    return peajes

def extract_pasos_por_peaje(container_text):
    """Extraer pasos por peaje del texto"""
    try:
        datos_pasos = {}
        
        chicoral_match = re.search(r'CHICORAL[^\d]*(\d{1,3},\d{3})', container_text, re.IGNORECASE)
        if chicoral_match:
            datos_pasos['CHICORAL'] = chicoral_match.group(1)
        
        cocora_match = re.search(r'COCORA[^\d]*(\d{1,3},\d{3}|\d+)', container_text, re.IGNORECASE)
        if cocora_match:
            datos_pasos['COCORA'] = cocora_match.group(1)
        
        gualanday_match = re.search(r'GUALANDAY[^\d]*(\d{1,3},\d{3})', container_text, re.IGNORECASE)
        if gualanday_match:
            datos_pasos['GUALANDAY'] = gualanday_match.group(1)
        
        total_match = re.search(r'Total[^\d]*(\d{1,3},\d{3})', container_text, re.IGNORECASE)
        if total_match:
            datos_pasos['TOTAL'] = total_match.group(1)
        
        return datos_pasos if datos_pasos else {}
            
    except Exception as e:
        return {}

def find_resumen_comercios_pasos(driver):
    """Buscar tabla 'RESUMEN COMERCIOS' para pasos"""
    try:
        titulo_selectors = [
            "//*[contains(text(), 'RESUMEN COMERCIOS')]",
            "//*[contains(text(), 'Resumen Comercios')]",
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        titulo_element = elemento
                        break
                if titulo_element:
                    break
            except:
                continue
        
        if not titulo_element:
            return None
        
        try:
            container = titulo_element.find_element(By.XPATH, "./ancestor::div[position()<=5]")
            container_text = container.text
            
            datos_pasos = extract_pasos_por_peaje(container_text)
            
            return datos_pasos if datos_pasos else None
                
        except Exception as e:
            return None
        
    except Exception as e:
        return None

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiYTFmOWZkMDAtY2IwYi00OTg4LWIxZDctNGZmYmU0NTMxNGI1IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None
    
    try:
        with st.spinner("Conectando con Power BI..."):
            driver.get(REPORT_URL)
            time.sleep(10)
        
        driver.save_screenshot("powerbi_inicial.png")
        
        if not click_conciliacion_date(driver, fecha_objetivo):
            return None
        
        time.sleep(3)
        driver.save_screenshot("powerbi_despues_seleccion.png")
        
        valor_texto = find_valor_a_pagar_comercio_card(driver)
        cantidad_pasos_texto = find_cantidad_pasos_card(driver)
        resumen_pasos = find_resumen_comercios_pasos(driver)
        valores_peajes = find_peaje_values(driver)
        
        driver.save_screenshot("powerbi_final.png")
        
        return {
            'valor_texto': valor_texto,
            'cantidad_pasos_texto': cantidad_pasos_texto or 'No encontrado',
            'resumen_pasos': resumen_pasos or {},
            'valores_peajes': valores_peajes,
            'screenshots': {
                'inicial': 'powerbi_inicial.png',
                'seleccion': 'powerbi_despues_seleccion.png',
                'final': 'powerbi_final.png'
            }
        }
        
    except Exception as e:
        st.error(f"Error durante la extracción: {str(e)}")
        return None
    finally:
        driver.quit()

# ===== FUNCIONES DE EXTRACCIÓN DE EXCEL =====

def extract_excel_values_with_steps(uploaded_file):
    """Extraer valores Y PASOS de las 3 hojas del Excel"""
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
                pasos_encontrado = None
                mejor_candidato_valor = None
                mejor_candidato_pasos = None
                mejor_puntaje_valor = -1
                mejor_puntaje_pasos = -1
                
                # Búsqueda de ABAJO hacia ARRIBA
                for i in range(len(df)-1, -1, -1):
                    fila = df.iloc[i]
                    
                    # Buscar "Total" en esta fila
                    for j, celda in enumerate(fila):
                        if pd.notna(celda) and isinstance(celda, str) and 'TOTAL' in celda.upper().strip():
                            
                            # ===== BÚSQUEDA DE VALOR A PAGAR =====
                            for k in range(len(fila)):
                                posible_valor = fila.iloc[k]
                                if pd.notna(posible_valor):
                                    valor_str = str(posible_valor)
                                    
                                    puntaje = 0
                                    if '$' in valor_str:
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
                                    
                                    if puntaje > mejor_puntaje_valor:
                                        mejor_puntaje_valor = puntaje
                                        mejor_candidato_valor = posible_valor
                            
                            # ===== BÚSQUEDA DE PASOS =====
                            for k in range(len(fila)):
                                posible_paso = fila.iloc[k]
                                if pd.notna(posible_paso):
                                    paso_str = str(posible_paso)
                                    
                                    if '$' in paso_str or ',' in paso_str:
                                        continue
                                    
                                    puntaje_paso = 0
                                    
                                    if any(c.isdigit() for c in paso_str):
                                        puntaje_paso += 5
                                    
                                    if paso_str.replace(',', '').isdigit():
                                        try:
                                            num_paso = int(paso_str.replace(',', ''))
                                            if 100 <= num_paso <= 100000:
                                                puntaje_paso += 10
                                                if 500 <= num_paso <= 50000:
                                                    puntaje_paso += 5
                                        except:
                                            puntaje_paso = 0
                                    
                                    if any(word in paso_str.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', '$']):
                                        puntaje_paso = 0
                                    
                                    if puntaje_paso > mejor_puntaje_pasos:
                                        mejor_puntaje_pasos = puntaje_paso
                                        mejor_candidato_pasos = posible_paso
                
                # Usar los mejores candidatos encontrados
                if mejor_candidato_valor is not None and mejor_puntaje_valor >= 5:
                    valor_encontrado = mejor_candidato_valor
                
                if mejor_candidato_pasos is not None and mejor_puntaje_pasos >= 5:
                    pasos_encontrado = mejor_candidato_pasos
                
                # Procesar el valor encontrado
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
                
                # Procesar los pasos encontrados
                if pasos_encontrado is not None:
                    pasos_original = str(pasos_encontrado)
                    pasos_limpio = re.sub(r'[^\d.,]', '', pasos_original)
                    
                    try:
                        if ',' in pasos_limpio and '.' in pasos_limpio:
                            pasos_limpio = pasos_limpio.replace('.', '')
                        if ',' in pasos_limpio:
                            pasos_limpio = pasos_limpio.replace(',', '')
                        
                        pasos_numerico = int(pasos_limpio)
                        
                        if 100 <= pasos_numerico <= 100000:
                            pasos[hoja] = pasos_numerico
                            total_pasos += pasos_numerico
                        else:
                            pasos[hoja] = 0
                            
                    except:
                        pasos[hoja] = 0
                else:
                    pasos[hoja] = 0
                    
            except Exception as e:
                valores[hoja] = 0
                pasos[hoja] = 0
        
        return valores, pasos, total_general, total_pasos
        
    except Exception as e:
        st.error(f"Error procesando archivo Excel: {str(e)}")
        return {}, {}, 0, 0

# ===== FUNCIONES DE COMPARACIÓN =====

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

def compare_peajes(valores_powerbi_peajes, valores, pasos_excel):
    """Comparar valores Y PASOS individuales por peaje"""
    comparaciones = {}
    
    for peaje in ['CHICORAL', 'COCORA', 'GUALANDAY']:
        try:
            valor_powerbi_texto = valores_powerbi_peajes.get(peaje)
            
            if valor_powerbi_texto is None:
                comparaciones[peaje] = {
                    'powerbi_texto': 'No encontrado',
                    'powerbi_numero': 0,
                    'excel_numero': valores.get(peaje, 0),
                    'coinciden': False,
                    'diferencia': valores.get(peaje, 0),
                    'pasos_excel': pasos_excel.get(peaje, 0),
                }
                continue
            
            powerbi_numero = convert_currency_to_float(valor_powerbi_texto)
            excel_numero = valores.get(peaje, 0)
            pasos_excel_valor = pasos_excel.get(peaje, 0)
            
            tolerancia = 0.01
            coinciden = abs(powerbi_numero - excel_numero) <= tolerancia
            diferencia = abs(powerbi_numero - excel_numero)
            
            comparaciones[peaje] = {
                'powerbi_texto': valor_powerbi_texto,
                'powerbi_numero': powerbi_numero,
                'excel_numero': excel_numero,
                'coinciden': coinciden,
                'diferencia': diferencia,
                'pasos_excel': pasos_excel_valor,
            }
            
        except Exception as e:
            comparaciones[peaje] = {
                'powerbi_texto': 'Error',
                'powerbi_numero': 0,
                'excel_numero': valores.get(peaje, 0),
                'coinciden': False,
                'diferencia': 0,
                'pasos_excel': pasos_excel.get(peaje, 0),
            }
    
    return comparaciones

# ===== INTERFAZ PRINCIPAL =====

def main():
    st.title("Validador Power BI - Conciliaciones APP GICA")
    st.markdown("---")
    
    # Información del reporte
    st.sidebar.header("Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Cargar archivo Excel con 3 hojas
    - Extraer valores Y PASOS de CHICORAL, GUALANDAY, COCORA
    - Calcular totales automáticamente
    - Comparar con Power BI (Total, Pasos y por Peaje)
    
    **Estado:** ChromeDriver Compatible
    **Versión:** v2.4 - Con Pasos por Peaje
    """)
    
    # Estado del sistema
    st.sidebar.header("Estado del Sistema")
    st.sidebar.success(f"Python {sys.version_info.major}.{sys.version_info.minor}")
    st.sidebar.info(f"Pandas {pd.__version__}")
    st.sidebar.info(f"Streamlit {st.__version__}")
    
    # Cargar archivo Excel
    st.subheader("Cargar Archivo Excel")
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
        
        # Extraer valores del Excel CON SPINNER
        with st.spinner("Procesando archivo Excel..."):
            valores, pasos, total_general, total_pasos = extract_excel_values_with_steps(uploaded_file)
        
        if total_general > 0:
            # ========== MOSTRAR RESUMEN DE VALORES Y PASOS ==========
            st.markdown("### Valores Extraídos del Excel")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("TOTAL CHICORAL", f"${valores['CHICORAL']:,.0f}".replace(",", "."))
                st.metric("PASOS CHICORAL", f"{pasos['CHICORAL']:,}".replace(",", "."))
            
            with col2:
                st.metric("TOTAL GUALANDAY", f"${valores['GUALANDAY']:,.0f}".replace(",", "."))
                st.metric("PASOS GUALANDAY", f"{pasos['GUALANDAY']:,}".replace(",", "."))
            
            with col3:
                st.metric("TOTAL COCORA", f"${valores['COCORA']:,.0f}".replace(",", "."))
                st.metric("PASOS COCORA", f"{pasos['COCORA']:,}".replace(",", "."))
            
            # Mostrar totales generales
            col_total1, col_total2 = st.columns(2)
            with col_total1:
                total_formateado = f"${total_general:,.0f}".replace(",", ".")
                st.metric("TOTAL GENERAL (Valores)", total_formateado, delta="Excel")
            
            with col_total2:
                st.metric("TOTAL GENERAL (Pasos)", f"{total_pasos:,}".replace(",", "."))
            
            st.markdown("---")
            
            # ========== PARÁMETROS Y EJECUCIÓN ==========
            if fecha_desde_archivo:
                st.info(f"Extracción Automática Activada | Fecha: {fecha_desde_archivo.strftime('%Y-%m-%d')}")
                fecha_objetivo = fecha_desde_archivo.strftime("%Y-%m-%d")
                ejecutar_extraccion = True
            else:
                st.subheader("Parámetros de Búsqueda")
                fecha_conciliacion = st.date_input(
                    "Fecha de Conciliación",
                    value=pd.to_datetime("2025-09-04"),
                    help="No se pudo detectar la fecha del archivo. Ingresa manualmente."
                )
                fecha_objetivo = fecha_conciliacion.strftime("%Y-%m-%d")
                
                if st.button("Extraer Valores de Power BI y Comparar", type="primary", use_container_width=True):
                    ejecutar_extraccion = True
                else:
                    ejecutar_extraccion = False
            
            # Ejecutar extracción si corresponde
            if ejecutar_extraccion:
                with st.spinner("Extrayendo datos de Power BI... Esto puede tomar 1-2 minutos"):
                    resultados = extract_powerbi_data(fecha_objetivo)
                    
                    if resultados and resultados.get('valor_texto'):
                        valor_powerbi_texto = resultados['valor_texto']
                        cantidad_pasos_texto = resultados.get('cantidad_pasos_texto', 'No encontrado')
                        resumen_pasos = resultados.get('resumen_pasos', {})
                        valores_peajes_powerbi = resultados.get('valores_peajes', {})
                        
                        st.markdown("---")
                        
                        # ========== RESULTADOS - VALORES POWER BI ==========
                        st.markdown("### Valores Extraídos de Power BI")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.metric("VALOR A PAGAR A COMERCIO", valor_powerbi_texto)
                        
                        with col2:
                            st.metric("CANTIDAD DE PASOS BI", cantidad_pasos_texto)
                        
                        # ========== CANTIDAD DE PASOS POR PEAJE ==========
                        if resumen_pasos:
                            st.markdown("### Cantidad de Pasos por Peaje (RESUMEN COMERCIOS)")
                            
                            col1, col2, col3, col4 = st.columns(4)
                            
                            with col1:
                                pasos_chicoral = resumen_pasos.get('CHICORAL', 'N/A')
                                st.metric("CHICORAL - Pasos", pasos_chicoral)
                            
                            with col2:
                                pasos_cocora = resumen_pasos.get('COCORA', 'N/A')
                                st.metric("COCORA - Pasos", pasos_cocora)
                            
                            with col3:
                                pasos_gualanday = resumen_pasos.get('GUALANDAY', 'N/A')
                                st.metric("GUALANDAY - Pasos", pasos_gualanday)
                            
                            with col4:
                                pasos_total = resumen_pasos.get('TOTAL', 'N/A')
                                st.metric("TOTAL - Pasos", pasos_total)
                        
                        st.markdown("---")
                        
                        # ========== VALIDACIÓN: TOTAL GENERAL ==========
                        st.markdown("### Validación: Total General")
                        
                        powerbi_numero, excel_numero, valor_formateado, coinciden = compare_values(
                            resultados, 
                            total_general
                        )
                        
                        if powerbi_numero is not None and excel_numero is not None:
                            col1, col2, col3 = st.columns([2, 2, 1])
                            
                            with col1:
                                st.metric("Power BI", valor_formateado)
                            with col2:
                                st.metric("Excel", f"${excel_numero:,.0f}".replace(",", "."))
                            with col3:
                                if coinciden:
                                    st.success("✓ COINCIDE")
                                else:
                                    diferencia = abs(powerbi_numero - excel_numero)
                                    st.error("✗ DIFERENCIA")
                                    st.caption(f"${diferencia:,.0f}".replace(",", "."))
                        
                        st.markdown("---")
                        
                        # ========== VALIDACIÓN: POR PEAJE ==========
                        st.markdown("### Validación: Por Peaje")
                        
                        comparaciones_peajes = compare_peajes(valores_peajes_powerbi, valores, pasos)
                        
                        # Crear tabla resumen
                        tabla_data = []
                        todos_coinciden = True
                        
                        for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                            comp = comparaciones_peajes[peaje]
                            
                            estado_icono = "✓" if comp['coinciden'] else "✗"
                            diferencia_texto = "$0" if comp['coinciden'] else f"${comp['diferencia']:,.0f}".replace(",", ".")
                            
                            tabla_data.append({
                                '': estado_icono,
                                'Peaje': peaje,
                                'Power BI': comp['powerbi_texto'],
                                'Excel (Valor)': f"${comp['excel_numero']:,.0f}".replace(",", "."),
                                'Excel (Pasos)': f"{comp['pasos_excel']:,}".replace(",", "."),
                                'Dif.': diferencia_texto
                            })
                            
                            if not comp['coinciden']:
                                todos_coinciden = False
                        
                        df_comparacion = pd.DataFrame(tabla_data)
                        st.dataframe(df_comparacion, use_container_width=True, hide_index=True)
                        
                        st.markdown("---")
                        
                        # ========== RESUMEN FINAL ==========
                        st.markdown("### Resultado Final")
                        
                        if coinciden and todos_coinciden:
                            st.success("✓ VALIDACIÓN EXITOSA - Todos los valores coinciden")
                            st.balloons()
                        elif coinciden and not todos_coinciden:
                            st.warning("⚠ VALIDACIÓN PARCIAL - El total coincide, pero hay diferencias por peaje")
                        elif not coinciden and todos_coinciden:
                            st.warning("⚠ VALIDACIÓN PARCIAL - Los peajes coinciden, pero el total tiene diferencias")
                        else:
                            st.error("✗ VALIDACIÓN FALLIDA - Existen diferencias en total y peajes")
                        
                        # Detalles adicionales
                        with st.expander("Ver Detalles Completos y Capturas"):
                            st.markdown("#### Tabla Detallada")
                            resumen_data = []
                            
                            resumen_data.append({
                                'Concepto': 'TOTAL GENERAL',
                                'Power BI': f"${powerbi_numero:,.0f}".replace(",", "."),
                                'Excel': f"${excel_numero:,.0f}".replace(",", "."),
                                'Estado': '✓ Coincide' if coinciden else '✗ No coincide',
                                'Diferencia': f"${abs(powerbi_numero - excel_numero):,.0f}".replace(",", "."),
                            })
                            
                            resumen_data.append({
                                'Concepto': 'CANTIDAD DE PASOS',
                                'Power BI': cantidad_pasos_texto,
                                'Excel': f"{total_pasos:,}".replace(",", "."),
                                'Estado': 'ℹ Información',
                                'Diferencia': 'N/A',
                            })
                            
                            if resumen_pasos:
                                for peaje in ['CHICORAL', 'COCORA', 'GUALANDAY']:
                                    pasos = resumen_pasos.get(peaje, 'N/A')
                                    resumen_data.append({
                                        'Concepto': f'{peaje} - PASOS',
                                        'Power BI': pasos,
                                        'Excel': f"{pasos_excel.get(peaje, 0):,}".replace(",", "."),
                                        'Estado': 'ℹ Información',
                                        'Diferencia': 'N/A',
                                    })
                                
                                if 'TOTAL' in resumen_pasos:
                                    resumen_data.append({
                                        'Concepto': 'TOTAL - PASOS',
                                        'Power BI': resumen_pasos['TOTAL'],
                                        'Excel': f"{total_pasos:,}".replace(",", "."),
                                        'Estado': 'ℹ Información',
                                        'Diferencia': 'N/A',
                                    })
                            
                            for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                                comp = comparaciones_peajes[peaje]
                                excel_val = comp['excel_numero']
                                pct_diff = f"{comp['diferencia']/excel_val*100:.2f}%" if excel_val > 0 else "N/A"
                                
                                resumen_data.append({
                                    'Concepto': peaje,
                                    'Power BI': comp['powerbi_texto'],
                                    'Excel': f"${comp['excel_numero']:,.0f}".replace(",", "."),
                                    'Estado': '✓ Coincide' if comp['coinciden'] else '✗ No coincide',
                                    'Diferencia': f"${comp['diferencia']:,.0f}".replace(",", "."),
                                })
                            
                            df_resumen = pd.DataFrame(resumen_data)
                            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
                            
                            # Screenshots
                            st.markdown("#### Capturas del Proceso")
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
                        st.error("Se accedió al reporte pero no se encontró el valor específico")
                    else:
                        st.error("No se pudieron extraer datos del reporte Power BI")
        else:
            st.error("No se pudieron extraer valores del archivo Excel")
            with st.expander("Sugerencias para solucionar el problema"):
                st.markdown("""
                - Verifica que las hojas se llamen **CHICORAL**, **GUALANDAY**, **COCORA**
                - Asegúrate de que haya valores numéricos en las celdas de total
                - Revisa que los totales estén claramente identificados con **'TOTAL'**
                - Los pasos deben estar en las mismas filas que los valores totales
                """)
    
    else:
        st.info("Por favor, carga un archivo Excel para comenzar la validación")

    # Información de ayuda
    st.markdown("---")
    with st.expander("Instrucciones de Uso"):
        st.markdown("""
        **Proceso:**
        1. **Cargar Excel**: Archivo con hojas CHICORAL, GUALANDAY, COCORA
        2. **Extracción automática**: Búsqueda inteligente de "Total" y "Pasos" en cada hoja
        3. **Seleccionar fecha** de conciliación en Power BI  
        4. **Comparar**: Extrae valores de Power BI y compara con Excel
        
        **Características (v2.4):**
        - Comparación de Total General
        - Extracción de Cantidad de Pasos total
        - Extracción de Pasos por Peaje desde Excel
        - Pasos por Peaje desde "RESUMEN COMERCIOS" Power BI
        - Comparación detallada por Peaje (Valores y Pasos)
        - Resumen completo con validación dual
        - Capturas de pantalla del proceso
        
        **Notas:**
        - La extracción busca el total, pasos y los valores individuales por peaje
        - Los valores deben estar claramente identificados en el Power BI
        - Las fechas deben coincidir exactamente con las del reporte Power BI
        """)

if __name__ == "__main__":
    main()

    # Footer
    st.markdown("---")
    st.markdown('<div class="footer">Desarrollado por Angel Torres | Powered by Streamlit | v2.4</div>', unsafe_allow_html=True)
