import os
import sys

# ===== CONFIGURACIÓN CRÍTICA PARA STREAMLIT CLOUD =====
os.environ['STREAMLIT_SERVER_FILE_WATCHER_TYPE'] = 'none'
os.environ['STREAMLIT_CI'] = 'true'
os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_STATIC_SERVING'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION'] = 'false'

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
from selenium.webdriver.chrome.options import Options
import time
import re

# Configuración de Streamlit
st.set_page_config(
    page_title="Validador Power BI - APP GICA",
    page_icon="💰",
    layout="wide"
)

# ===== CSS =====
st.markdown("""
<style>
[data-testid="stSidebar"] {
    background-color: #1E1E2F !important;
    color: white !important;
}
[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stCheckbox label {
    color: white !important; 
}
[data-testid="stSidebar"] .stFileUploader > label {
    color: white !important;
    font-weight: bold;
}
.stSpinner > div > div {
    border-color: #00CFFF !important;
}
.stProgress > div > div > div > div {
    background-color: #00CFFF !important;
}
</style>
""", unsafe_allow_html=True)

# Logo
st.markdown("""
<div style="display: flex; justify-content: center; margin-bottom: 30px;">
    <img src="https://i.imgur.com/z9xt46F.jpeg"
         style="width: 50%; border-radius: 10px; display: block; margin: 0 auto;" 
         alt="Logo Gopass">
</div>
""", unsafe_allow_html=True)

# ===== FUNCIONES DE EXTRACCIÓN DE POWER BI =====

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
        
        driver = webdriver.Chrome(options=chrome_options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver
    except Exception as e:
        st.error(f"❌ Error al configurar ChromeDriver: {e}")
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
        
        for selector in selectors:
            try:
                elemento = driver.find_element(By.XPATH, selector)
                if elemento.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", elemento)
                    time.sleep(3)
                    return True
            except:
                continue
        return False
    except Exception as e:
        return False

def find_cantidad_pasos_card(driver):
    """Buscar la tarjeta 'CANTIDAD PASOS'"""
    try:
        titulo_selectors = [
            "//*[contains(text(), 'CANTIDAD PASOS')]",
            "//*[contains(text(), 'Cantidad Pasos')]",
            "//*[contains(text(), 'CANTIDAD DE PASOS')]",
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
        
        # Buscar en el contenedor padre
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
        except:
            pass
        
        return None
    except Exception:
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
    except Exception:
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
            
            # Buscar valor cerca del título
            try:
                container = titulo_element.find_element(By.XPATH, "./ancestor::*[position()<=3]")
                numeric_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '$') or contains(text(), ',') or contains(text(), '.')]")
                
                for elem in numeric_elements:
                    texto = elem.text.strip()
                    if texto and any(char.isdigit() for char in texto):
                        if 'VALOR A PAGAR' not in texto.upper() and 'COMERCIO' not in texto.upper():
                            peajes[nombre_peaje] = texto
                            break
            except:
                peajes[nombre_peaje] = None
                    
        except Exception:
            peajes[nombre_peaje] = None
    
    return peajes

def extract_pasos_por_peaje(container_text):
    """Extraer los pasos por peaje del texto"""
    try:
        datos_pasos = {}
        
        # ESTRATEGIA 1: Buscar patrones específicos
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
        
        # ESTRATEGIA 2: Si faltan datos, buscar números en contexto
        if len(datos_pasos) < 4:
            all_comma_numbers = re.findall(r'\b\d{1,3},\d{3}\b', container_text)
            all_simple_numbers = re.findall(r'\b\d{3,4}\b', container_text)
            all_numbers = all_comma_numbers + all_simple_numbers
            
            valid_numbers = []
            for num_str in all_numbers:
                num_clean = num_str.replace(',', '')
                if num_clean.isdigit():
                    num_val = int(num_clean)
                    if 100 <= num_val <= 10000:
                        valid_numbers.append(num_str)
            
            if len(valid_numbers) >= 4:
                number_positions = []
                for num in valid_numbers:
                    pos = container_text.find(num)
                    if pos != -1:
                        number_positions.append((pos, num))
                
                number_positions.sort()
                
                if len(number_positions) >= 4:
                    datos_pasos['CHICORAL'] = number_positions[0][1]
                    datos_pasos['COCORA'] = number_positions[1][1]
                    datos_pasos['GUALANDAY'] = number_positions[2][1]
                    datos_pasos['TOTAL'] = number_positions[3][1]
        
        return datos_pasos if datos_pasos else {}
            
    except Exception:
        return {}

def find_resumen_comercios_pasos(driver):
    """Buscar la tabla 'RESUMEN COMERCIOS' y extraer pasos"""
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
                
        except Exception:
            return None
        
    except Exception:
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
        
        if not cantidad_pasos_texto:
            cantidad_pasos_texto = buscar_cantidad_pasos_alternativo(driver)
        
        with st.spinner("🔍 Extrayendo datos de pasos por peaje..."):
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
        st.error(f"❌ Error durante la extracción: {str(e)}")
        return None
    finally:
        driver.quit()
        
def buscar_cantidad_pasos_alternativo(driver):
    """Búsqueda alternativa para CANTIDAD PASOS"""
    try:
        all_elements = driver.find_elements(By.XPATH, "//*[text()]")
        
        for elem in all_elements:
            texto = elem.text.strip()
            if (texto and 
                any(char.isdigit() for char in texto) and
                3 <= len(texto) <= 10 and
                not any(word in texto.upper() for word in ['$', 'TOTAL', 'VALOR', 'PAGAR', 'COMERCIO'])):
                
                clean_text = texto.replace(',', '').replace('.', '')
                if clean_text.isdigit():
                    num_value = int(clean_text)
                    if 100 <= num_value <= 999999:
                        return texto
        
        return None
    except Exception:
        return None

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
                                                ('$' in valor_str or '.' in valor_str)):
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
        
    except Exception:
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
        
    except Exception:
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
            
        except Exception:
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
    
    # Sidebar
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Cargar archivo Excel con 3 hojas
    - Extraer valores de CHICORAL, GUALANDAY, COCORA
    - Calcular total automáticamente
    - Comparar con Power BI (Total, Pasos y por Peaje)
    
    **Versión:** v2.3 - Con Cantidad de Pasos por Peaje
    """)
    
    # Cargar archivo Excel
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
            valores, total_general = extract_excel_values(uploaded_file)
        
        if total_general > 0:
            # Mostrar valores del Excel
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
            
            # Parámetros de ejecución
            if fecha_desde_archivo:
                fecha_objetivo = fecha_desde_archivo.strftime("%Y-%m-%d")
                ejecutar_extraccion = True
            else:
                fecha_conciliacion = st.date_input(
                    "Fecha de Conciliación",
                    value=pd.to_datetime("2025-09-04")
                )
                fecha_objetivo = fecha_conciliacion.strftime("%Y-%m-%d")
                
                if st.button("🎯 Extraer Valores de Power BI y Comparar", type="primary", use_container_width=True):
                    ejecutar_extraccion = True
                else:
                    ejecutar_extraccion = False
            
            # Ejecutar extracción
            if ejecutar_extraccion:
                with st.spinner("🌐 Extrayendo datos de Power BI..."):
                    resultados = extract_powerbi_data(fecha_objetivo)
                    
                    if resultados and resultados.get('valor_texto'):
                        valor_powerbi_texto = resultados['valor_texto']
                        cantidad_pasos_texto = resultados.get('cantidad_pasos_texto', 'No encontrado')
                        resumen_pasos = resultados.get('resumen_pasos', {})
                        valores_peajes_powerbi = resultados.get('valores_peajes', {})
                        
                        st.markdown("---")
                        
                        # Valores de Power BI
                        st.markdown("### 📊 Valores Extraídos de Power BI")
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("💰 VALOR A PAGAR A COMERCIO", valor_powerbi_texto)
                        with col2:
                            st.metric("👣 CANTIDAD DE PASOS BI", cantidad_pasos_texto)
                        
                        # Pasos por peaje
                        if resumen_pasos:
                            st.markdown("### 👣 Cantidad de Pasos por Peaje")
                            
                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                st.metric("CHICORAL - Pasos", resumen_pasos.get('CHICORAL', 'N/A'))
                            with col2:
                                st.metric("COCORA - Pasos", resumen_pasos.get('COCORA', 'N/A'))
                            with col3:
                                st.metric("GUALANDAY - Pasos", resumen_pasos.get('GUALANDAY', 'N/A'))
                            with col4:
                                st.metric("TOTAL - Pasos", resumen_pasos.get('TOTAL', 'N/A'))
                        
                        st.markdown("---")
                        
                        # Comparación total
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
                                    st.success("✅ COINCIDE")
                                else:
                                    diferencia = abs(powerbi_numero - excel_numero)
                                    st.error(f"❌ DIFERENCIA\n${diferencia:,.0f}".replace(",", "."))
                        
                        st.markdown("---")
                        
                        # Comparación por peaje
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
                        
                        st.dataframe(pd.DataFrame(tabla_data), use_container_width=True, hide_index=True)
                        
                        st.markdown("---")
                        
                        # Resultado final
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
                        
                        # Detalles adicionales
                        with st.expander("🔍 Ver Detalles Completos"):
                            resumen_data = []
                            
                            resumen_data.append({
                                'Concepto': 'TOTAL GENERAL',
                                'Power BI': f"${powerbi_numero:,.0f}".replace(",", "."),
                                'Excel': f"${excel_numero:,.0f}".replace(",", "."),
                                'Estado': '✅ Coincide' if coinciden else '❌ No coincide',
                                'Diferencia': f"${abs(powerbi_numero - excel_numero):,.0f}".replace(",", "."),
                            })
                            
                            resumen_data.append({
                                'Concepto': 'CANTIDAD DE PASOS',
                                'Power BI': cantidad_pasos_texto,
                                'Excel': 'N/A',
                                'Estado': 'ℹ️ Solo Power BI',
                                'Diferencia': 'N/A',
                            })
                            
                            if resumen_pasos:
                                for peaje in ['CHICORAL', 'COCORA', 'GUALANDAY']:
                                    resumen_data.append({
                                        'Concepto': f'{peaje} - PASOS',
                                        'Power BI': resumen_pasos.get(peaje, 'N/A'),
                                        'Excel': 'N/A',
                                        'Estado': 'ℹ️ Solo Power BI',
                                        'Diferencia': 'N/A',
                                    })
                            
                            for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                                comp = comparaciones_peajes[peaje]
                                resumen_data.append({
                                    'Concepto': peaje,
                                    'Power BI': comp['powerbi_texto'],
                                    'Excel': f"${comp['excel_numero']:,.0f}".replace(",", "."),
                                    'Estado': '✅ Coincide' if comp['coinciden'] else '❌ No coincide',
                                    'Diferencia': f"${comp['diferencia']:,.0f}".replace(",", "."),
                                })
                            
                            st.dataframe(pd.DataFrame(resumen_data), use_container_width=True, hide_index=True)
                                
                    else:
                        st.error("❌ No se pudieron extraer datos del reporte Power BI")
        else:
            st.error("❌ No se pudieron extraer valores del archivo Excel")
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar la validación")

if __name__ == "__main__":
    main()
    st.markdown("---")
    st.markdown('<div style="text-align: center">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v2.3</div>', unsafe_allow_html=True)
