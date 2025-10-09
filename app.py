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

# ===== CSS MEJORADO =====
st.markdown("""
<style>
    .main-header {
        text-align: center;
        color: #1E1E2F;
        margin-bottom: 30px;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .comparison-table {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        margin: 15px 0;
    }
</style>
""", unsafe_allow_html=True)

# Logo de GoPass
st.markdown("""
<div style="display: flex; justify-content: center; margin-bottom: 30px;">
    <img src="https://i.imgur.com/z9xt46F.jpeg"
         style="width: 50%; border-radius: 10px; display: block; margin: 0 auto;" 
         alt="Logo Gopass">
</div>
""", unsafe_allow_html=True)

# ===== FUNCIONES DE EXTRACCIÓN MEJORADAS =====

def setup_driver():
    """Configurar ChromeDriver de forma robusta"""
    try:
        chrome_options = Options()
        
        # Configuración esencial para entornos cloud
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # User agent realista
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        # Configuración adicional para mejorar la estabilidad
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-plugins")
        chrome_options.add_argument("--disable-images")
        chrome_options.add_argument("--disable-javascript")
        
        try:
            # Intentar usar ChromeDriver del sistema
            driver = webdriver.Chrome(options=chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            return driver
        except Exception as e:
            st.error(f"❌ Error al iniciar ChromeDriver: {e}")
            return None
            
    except Exception as e:
        st.error(f"❌ Error crítico en setup_driver: {e}")
        return None

def wait_for_element(driver, xpath, timeout=30):
    """Esperar a que un elemento esté presente y visible"""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        return element
    except:
        return None

def click_conciliacion_date(driver, fecha_objetivo):
    """Hacer clic en la conciliación específica por fecha - VERSIÓN MEJORADA"""
    try:
        st.info("🔍 Buscando conciliación en Power BI...")
        
        # Diferentes formatos de búsqueda
        formatos_fecha = [
            fecha_objetivo,  # 2024-01-15
            fecha_objetivo.replace("-", "/"),  # 2024/01/15
            pd.to_datetime(fecha_objetivo).strftime("%d/%m/%Y"),  # 15/01/2024
            pd.to_datetime(fecha_objetivo).strftime("%d-%m-%Y"),  # 15-01-2024
        ]
        
        for formato_fecha in formatos_fecha:
            selectors = [
                f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'conciliación app gica del {formato_fecha.lower()}')]",
                f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'conciliación del {formato_fecha.lower()}')]",
                f"//*[contains(text(), '{formato_fecha}')]",
                f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{formato_fecha.lower()}')]",
                f"//div[contains(text(), '{formato_fecha}')]",
                f"//span[contains(text(), '{formato_fecha}')]",
                f"//button[contains(text(), '{formato_fecha}')]",
                f"//*[contains(@title, '{formato_fecha}')]",
            ]
            
            for selector in selectors:
                try:
                    elementos = driver.find_elements(By.XPATH, selector)
                    for elemento in elementos:
                        if elemento.is_displayed() and elemento.is_enabled():
                            texto = elemento.text.strip()
                            if formato_fecha in texto or any(palabra in texto.lower() for palabra in ['conciliación', 'app gica']):
                                # Hacer clic usando JavaScript para mayor confiabilidad
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
                                time.sleep(1)
                                driver.execute_script("arguments[0].click();", elemento)
                                time.sleep(5)  # Esperar más tiempo después del clic
                                st.success(f"✅ Conciliación encontrada y seleccionada: {texto}")
                                return True
                except Exception as e:
                    continue
        
        st.error("❌ No se encontró la conciliación para la fecha especificada")
        return False
        
    except Exception as e:
        st.error(f"❌ Error al hacer clic en conciliación: {str(e)}")
        return False

def find_valor_a_pagar_comercio_card(driver):
    """Buscar la tarjeta 'VALOR A PAGAR A COMERCIO' - VERSIÓN ROBUSTA"""
    try:
        # Patrones de búsqueda más flexibles
        titulo_selectors = [
            "//*[contains(translate(text(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'VALOR A PAGAR A COMERCIO')]",
            "//*[contains(translate(text(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'VALOR A PAGAR')]",
            "//*[contains(translate(text(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'PAGAR A COMERCIO')]",
            "//*[contains(text(), 'Valor a pagar a comercio')]",
            "//*[contains(text(), 'Valor A Pagar A Comercio')]",
        ]
        
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        if any(palabra in texto.upper() for palabra in ['PAGAR', 'COMERCIO', 'VALOR']):
                            st.info(f"🔍 Título encontrado: {texto}")
                            
                            # Buscar el valor numérico en elementos cercanos
                            parent = elemento.find_element(By.XPATH, "./..")
                            
                            # Buscar en elementos hermanos
                            siblings = parent.find_elements(By.XPATH, "./*")
                            for sibling in siblings:
                                if sibling != elemento:
                                    sibling_text = sibling.text.strip()
                                    if sibling_text and any(c.isdigit() for c in sibling_text):
                                        # Verificar si es un valor monetario
                                        if '$' in sibling_text or any(c in sibling_text for c in [',', '.']):
                                            st.success(f"✅ Valor encontrado: {sibling_text}")
                                            return sibling_text
                            
                            # Buscar en elementos siguientes
                            following_elements = driver.find_elements(By.XPATH, f"{selector}/following::*")
                            for i, elem in enumerate(following_elements[:10]):
                                elem_text = elem.text.strip()
                                if elem_text and any(c.isdigit() for c in elem_text):
                                    if '$' in elem_text or any(c in elem_text for c in [',', '.']):
                                        st.success(f"✅ Valor encontrado (siguiente {i}): {elem_text}")
                                        return elem_text
            except Exception as e:
                continue
        
        st.warning("⚠️ No se encontró 'VALOR A PAGAR A COMERCIO' en el formato esperado")
        return None
        
    except Exception as e:
        st.error(f"❌ Error buscando valor: {str(e)}")
        return None

def find_cantidad_pasos_card(driver):
    """Buscar la tarjeta 'CANTIDAD PASOS' - VERSIÓN MEJORADA"""
    try:
        titulo_selectors = [
            "//*[contains(translate(text(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'CANTIDAD PASOS')]",
            "//*[contains(translate(text(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'CANTIDAD DE PASOS')]",
            "//*[contains(text(), 'Cantidad Pasos')]",
            "//*[contains(text(), 'Cantidad de Pasos')]",
        ]
        
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        if any(palabra in texto.upper() for palabra in ['CANTIDAD', 'PASOS']):
                            st.info(f"🔍 Título CANTIDAD PASOS encontrado: {texto}")
                            
                            # Buscar valor numérico
                            parent = elemento.find_element(By.XPATH, "./..")
                            siblings = parent.find_elements(By.XPATH, "./*")
                            
                            for sibling in siblings:
                                if sibling != elemento:
                                    sibling_text = sibling.text.strip()
                                    if sibling_text and any(c.isdigit() for c in sibling_text):
                                        # Verificar que sea un número de pasos (sin símbolos monetarios)
                                        if '$' not in sibling_text and len(sibling_text) < 10:
                                            st.success(f"✅ Cantidad de pasos encontrada: {sibling_text}")
                                            return sibling_text
                            
                            # Buscar en elementos siguientes
                            following_elements = driver.find_elements(By.XPATH, f"{selector}/following::*")
                            for i, elem in enumerate(following_elements[:10]):
                                elem_text = elem.text.strip()
                                if elem_text and any(c.isdigit() for c in elem_text):
                                    if '$' not in elem_text and len(elem_text) < 10:
                                        st.success(f"✅ Cantidad de pasos encontrada (siguiente {i}): {elem_text}")
                                        return elem_text
            except Exception as e:
                continue
        
        st.warning("⚠️ No se encontró 'CANTIDAD PASOS' en el formato esperado")
        return None
        
    except Exception as e:
        st.error(f"❌ Error buscando cantidad de pasos: {str(e)}")
        return None

def find_peaje_values(driver):
    """Buscar valores individuales de cada peaje - VERSIÓN SIMPLIFICADA"""
    peajes = {}
    nombres_peajes = ['CHICORAL', 'COCORA', 'GUALANDAY']
    
    for nombre_peaje in nombres_peajes:
        try:
            # Buscar el nombre del peaje
            selectors = [
                f"//*[contains(translate(text(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '{nombre_peaje}')]",
                f"//*[contains(text(), '{nombre_peaje.title()}')]",
            ]
            
            for selector in selectors:
                try:
                    elementos = driver.find_elements(By.XPATH, selector)
                    for elemento in elementos:
                        if elemento.is_displayed():
                            texto = elemento.text.strip().upper()
                            if nombre_peaje in texto:
                                # Buscar valores cerca del elemento
                                parent = elemento.find_element(By.XPATH, "./ancestor::div[position()<=3]")
                                numeric_elements = parent.find_elements(By.XPATH, ".//*[contains(text(), '$')]")
                                
                                for elem in numeric_elements:
                                    elem_text = elem.text.strip()
                                    if elem_text and any(c.isdigit() for c in elem_text):
                                        peajes[nombre_peaje] = elem_text
                                        st.success(f"✅ Valor {nombre_peaje}: {elem_text}")
                                        break
                                
                                if nombre_peaje in peajes:
                                    break
                except:
                    continue
                    
            if nombre_peaje not in peajes:
                peajes[nombre_peaje] = None
                st.warning(f"⚠️ No se encontró valor para {nombre_peaje}")
                
        except Exception as e:
            peajes[nombre_peaje] = None
            st.error(f"❌ Error buscando {nombre_peaje}: {e}")
    
    return peajes

def find_pasos_por_peaje_bi(driver):
    """Buscar cantidad de pasos por peaje - VERSIÓN SIMPLIFICADA"""
    try:
        pasos_peajes = {}
        nombres_peajes = ['CHICORAL', 'COCORA', 'GUALANDAY']
        total_pasos_bi = 0
        
        # Buscar en toda la página números que parezcan cantidades de pasos
        all_elements = driver.find_elements(By.XPATH, "//*[text()]")
        
        for elemento in all_elements:
            texto = elemento.text.strip()
            # Buscar números que parezcan cantidades (sin decimales, sin símbolos monetarios)
            if (texto and 
                texto.isdigit() and 
                3 <= len(texto) <= 6 and
                100 <= int(texto) <= 999999):
                
                # Verificar si está cerca de un nombre de peaje
                parent = elemento.find_element(By.XPATH, "./ancestor::div[position()<=3]")
                parent_text = parent.text.upper()
                
                for peaje in nombres_peajes:
                    if peaje in parent_text and peaje not in pasos_peajes:
                        pasos_peajes[peaje] = int(texto)
                        total_pasos_bi += int(texto)
                        st.success(f"✅ Pasos {peaje}: {texto}")
                        break
        
        # Completar peajes no encontrados
        for peaje in nombres_peajes:
            if peaje not in pasos_peajes:
                pasos_peajes[peaje] = 0
                st.warning(f"⚠️ No se encontraron pasos para {peaje}")
        
        return pasos_peajes, total_pasos_bi
        
    except Exception as e:
        st.error(f"❌ Error buscando pasos por peaje: {e}")
        return {peaje: 0 for peaje in ['CHICORAL', 'COCORA', 'GUALANDAY']}, 0

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI - VERSIÓN ROBUSTA"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiYTFmOWZkMDAtY2IwYi00OTg4LWIxZDctNGZmYmU0NTMxNGI1IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    st.info("🚀 Iniciando extracción de Power BI...")
    
    driver = setup_driver()
    if not driver:
        st.error("❌ No se pudo inicializar el navegador")
        return None
    
    try:
        # 1. Navegar al reporte
        with st.spinner("🌐 Conectando con Power BI..."):
            driver.get(REPORT_URL)
            time.sleep(15)  # Más tiempo para carga inicial
            
            # Verificar que cargó correctamente
            if "Power BI" not in driver.title and "powerbi" not in driver.current_url:
                st.error("❌ No se pudo cargar la página de Power BI")
                return None
        
        # 2. Esperar a que cargue el contenido
        time.sleep(10)
        
        # 3. Hacer clic en la conciliación específica
        st.info("📅 Buscando conciliación por fecha...")
        if not click_conciliacion_date(driver, fecha_objetivo):
            st.error("❌ No se pudo seleccionar la conciliación")
            return None
        
        # 4. Esperar a que cargue la selección
        time.sleep(10)
        
        # 5. Buscar valores principales
        st.info("💰 Buscando valores principales...")
        valor_texto = find_valor_a_pagar_comercio_card(driver)
        cantidad_pasos_texto = find_cantidad_pasos_card(driver)
        
        # 6. Buscar valores por peaje
        st.info("🏢 Buscando valores por peaje...")
        valores_peajes = find_peaje_values(driver)
        
        # 7. Buscar pasos por peaje
        st.info("👣 Buscando pasos por peaje...")
        pasos_peajes_bi, total_pasos_bi = find_pasos_por_peaje_bi(driver)
        
        # Verificar que al menos se encontró algo
        if not valor_texto and not any(valores_peajes.values()):
            st.error("❌ No se encontraron valores en el reporte")
            return None
        
        st.success("✅ Extracción de Power BI completada")
        
        return {
            'valor_texto': valor_texto,
            'cantidad_pasos_texto': cantidad_pasos_texto,
            'valores_peajes': valores_peajes,
            'pasos_peajes_bi': pasos_peajes_bi,
            'total_pasos_bi': total_pasos_bi,
        }
        
    except Exception as e:
        st.error(f"❌ Error durante la extracción: {str(e)}")
        import traceback
        st.error(f"🔍 Detalles del error: {traceback.format_exc()}")
        return None
    finally:
        try:
            driver.quit()
        except:
            pass

# ===== FUNCIONES DE EXTRACCIÓN DE EXCEL (MANTENIDAS) =====

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

# ===== FUNCIONES DE COMPARACIÓN (MANTENIDAS) =====

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

# ===== INTERFAZ PRINCIPAL =====

def main():
    st.markdown('<h1 class="main-header">💰 Validador Power BI - Conciliaciones APP GICA</h1>', unsafe_allow_html=True)
    st.markdown("---")
    
    # Sidebar
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Cargar archivo Excel con 3 hojas
    - Extraer valores Y pasos por peaje
    - Comparar con Power BI
    
    **Versión:** v2.4 - Extracción Mejorada
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
                st.success(f"📅 Fecha detectada del archivo: {fecha_desde_archivo.strftime('%Y-%m-%d')}")
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
                fecha_objetivo = fecha_desde_archivo.strftime("%Y-%m-%d")
                ejecutar_extraccion = True
            else:
                st.subheader("📅 Parámetros de Búsqueda")
                fecha_conciliacion = st.date_input(
                    "Fecha de Conciliación",
                    value=pd.to_datetime("2024-01-15")
                )
                fecha_objetivo = fecha_conciliacion.strftime("%Y-%m-%d")
                
                if st.button("🎯 Extraer Valores de Power BI y Comparar", type="primary", use_container_width=True):
                    ejecutar_extraccion = True
                else:
                    ejecutar_extraccion = False
            
            if ejecutar_extraccion:
                with st.spinner("🌐 Extrayendo datos de Power BI... Esto puede tomar 1-2 minutos"):
                    resultados = extract_powerbi_data(fecha_objetivo)
                    
                    if resultados:
                        # ========== TABLA COMPARACIÓN COMPLETA ==========
                        st.markdown("### 📋 Tabla de Comparación Completa")
                        
                        # Crear tabla comparativa
                        tabla_comparativa = []
                        
                        # Totales generales
                        powerbi_numero, excel_numero, _, coinciden_valor = compare_values(resultados, total_general)
                        
                        if powerbi_numero:
                            tabla_comparativa.append({
                                'Concepto': 'TOTAL GENERAL',
                                'Power BI': f"${powerbi_numero:,.0f}".replace(",", "."),
                                'Excel': f"${total_general:,.0f}".replace(",", "."),
                                'Estado': '✅' if coinciden_valor else '❌',
                                'Diferencia': f"${abs(powerbi_numero - total_general):,.0f}".replace(",", ".")
                            })
                        
                        # Pasos totales
                        cantidad_pasos_bi = resultados.get('total_pasos_bi', 0)
                        coinciden_pasos = cantidad_pasos_bi == total_pasos
                        
                        tabla_comparativa.append({
                            'Concepto': 'TOTAL PASOS',
                            'Power BI': f"{cantidad_pasos_bi:,}".replace(",", "."),
                            'Excel': f"{total_pasos:,}".replace(",", "."),
                            'Estado': '✅' if coinciden_pasos else '❌',
                            'Diferencia': f"{abs(cantidad_pasos_bi - total_pasos):,}".replace(",", ".")
                        })
                        
                        # Valores por peaje
                        comparaciones_peajes = compare_peajes(resultados.get('valores_peajes', {}), valores)
                        for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                            comp = comparaciones_peajes[peaje]
                            tabla_comparativa.append({
                                'Concepto': f'VALOR {peaje}',
                                'Power BI': comp['powerbi_texto'],
                                'Excel': f"${comp['excel_numero']:,.0f}".replace(",", "."),
                                'Estado': '✅' if comp['coinciden'] else '❌',
                                'Diferencia': f"${comp['diferencia']:,.0f}".replace(",", ".")
                            })
                        
                        # Pasos por peaje
                        comparaciones_pasos = compare_pasos_peajes(resultados.get('pasos_peajes_bi', {}), pasos)
                        for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                            comp = comparaciones_pasos[peaje]
                            tabla_comparativa.append({
                                'Concepto': f'PASOS {peaje}',
                                'Power BI': f"{comp['pasos_bi']:,}".replace(",", "."),
                                'Excel': f"{comp['pasos_excel']:,}".replace(",", "."),
                                'Estado': '✅' if comp['coinciden'] else '❌',
                                'Diferencia': f"{comp['diferencia']:,}".replace(",", ".")
                            })
                        
                        df_comparativa = pd.DataFrame(tabla_comparativa)
                        st.dataframe(df_comparativa, use_container_width=True, hide_index=True)
                        
                        # ========== RESUMEN FINAL ==========
                        st.markdown("---")
                        st.markdown("### 📈 Resultado Final")
                        
                        # Calcular estado general
                        todos_valores_coinciden = all(comp['coinciden'] for comp in comparaciones_peajes.values())
                        todos_pasos_coinciden = all(comp['coinciden'] for comp in comparaciones_pasos.values())
                        
                        validacion_completa = (coinciden_valor and coinciden_pasos and 
                                             todos_valores_coinciden and todos_pasos_coinciden)
                        
                        if validacion_completa:
                            st.success("🎉 **VALIDACIÓN EXITOSA** - Todos los valores coinciden")
                            st.balloons()
                        else:
                            st.warning("⚠️ **VALIDACIÓN CON DIFERENCIAS** - Revisar la tabla de comparación")
                            
                    else:
                        st.error("❌ No se pudieron extraer datos del reporte Power BI")
                        st.info("""
                        **Posibles soluciones:**
                        1. Verifica que la fecha sea correcta
                        2. Asegúrate de que el reporte Power BI esté accesible
                        3. Revisa que la conciliación exista para esa fecha
                        4. Intenta con otra fecha
                        """)
        else:
            st.error("❌ No se pudieron extraer valores del archivo Excel")
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar la validación")

if __name__ == "__main__":
    main()

    # Footer
    st.markdown("---")
    st.markdown('<div class="footer">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v2.4</div>', unsafe_allow_html=True)
