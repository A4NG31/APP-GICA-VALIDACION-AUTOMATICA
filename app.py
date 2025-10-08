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

# ===== FUNCIONES DE EXTRACCIÓN DE POWER BI (ACTUALIZADAS) =====

def setup_driver():
    """Configurar ChromeDriver para Selenium - VERSIÓN COMPATIBLE"""
    try:
        chrome_options = Options()
        
        # Opciones para mejor compatibilidad
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # User agent real
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        # SOLUCIÓN: Usar ChromeDriver del sistema instalado via packages.txt
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
        # Buscar el elemento que contiene la fecha exacta
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
            # Hacer clic en el elemento
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

def find_valor_a_pagar_comercio_card(driver):
    """Buscar la tarjeta/table 'VALOR A PAGAR A COMERCIO' en la parte superior derecha"""
    try:
        # Buscar por diferentes patrones del título
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
        
        # Buscar el valor numérico debajo del título
        # Estrategia 1: Buscar en el mismo contenedor
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            numeric_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '$') or contains(text(), ',') or contains(text(), '.')]")
            
            for elem in numeric_elements:
                texto = elem.text.strip()
                if texto and any(char.isdigit() for char in texto) and texto != titulo_element.text:
                    return texto
        except:
            pass
        
        # Estrategia 2: Buscar en elementos hermanos
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
        
        # Estrategia 3: Buscar debajo del título
        try:
            following_elements = driver.find_elements(By.XPATH, f"//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]/following::*")
            
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
    """
    NUEVA FUNCIÓN: Buscar valores individuales de cada peaje en el Power BI
    Busca las tarjetas/tablas de CHICORAL, COCORA y GUALANDAY
    """
    peajes = {}
    nombres_peajes = ['CHICORAL', 'COCORA', 'GUALANDAY']
    
    for nombre_peaje in nombres_peajes:
        try:
            # Buscar el título del peaje
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
                st.warning(f"⚠️ No se encontró el título '{nombre_peaje}' en el reporte")
                peajes[nombre_peaje] = None
                continue
            
            # Estrategia 1: Buscar "VALOR A PAGAR" cerca del título del peaje
            try:
                # Buscar en el contenedor padre
                container = titulo_element.find_element(By.XPATH, "./ancestor::*[position()<=3]")
                
                # Buscar "VALOR A PAGAR" dentro del contenedor
                valor_pagar_elements = container.find_elements(By.XPATH, ".//*[contains(text(), 'VALOR A PAGAR') or contains(text(), 'Valor a pagar')]")
                
                if valor_pagar_elements:
                    # Buscar el valor numérico cerca de "VALOR A PAGAR"
                    valor_element = valor_pagar_elements[0]
                    
                    # Buscar valores numéricos en el mismo contenedor
                    numeric_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '$')]")
                    
                    for elem in numeric_elements:
                        texto = elem.text.strip()
                        if texto and any(char.isdigit() for char in texto) and '$' in texto:
                            # Verificar que no sea el título
                            if 'VALOR A PAGAR' not in texto.upper() and 'COMERCIO' not in texto.upper():
                                peajes[nombre_peaje] = texto
                                break
            except:
                pass
            
            # Estrategia 2: Buscar valores numéricos después del título
            if nombre_peaje not in peajes or peajes[nombre_peaje] is None:
                try:
                    following_elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{nombre_peaje}')]/following::*")
                    
                    for elem in following_elements[:15]:
                        texto = elem.text.strip()
                        if texto and any(char.isdigit() for char in texto) and '$' in texto:
                            # Verificar que sea un valor monetario válido
                            if len(texto) > 5 and len(texto) < 50:
                                peajes[nombre_peaje] = texto
                                break
                except:
                    pass
            
            # Estrategia 3: Buscar en elementos hermanos
            if nombre_peaje not in peajes or peajes[nombre_peaje] is None:
                try:
                    parent = titulo_element.find_element(By.XPATH, "./..")
                    siblings = parent.find_elements(By.XPATH, "./*")
                    
                    for sibling in siblings:
                        if sibling != titulo_element:
                            texto = sibling.text.strip()
                            if texto and any(char.isdigit() for char in texto) and '$' in texto:
                                if len(texto) > 5 and len(texto) < 50:
                                    peajes[nombre_peaje] = texto
                                    break
                except:
                    pass
            
            if nombre_peaje not in peajes or peajes[nombre_peaje] is None:
                st.warning(f"⚠️ No se pudo encontrar el valor para {nombre_peaje}")
                peajes[nombre_peaje] = None
            else:
                st.info(f"✅ {nombre_peaje}: {peajes[nombre_peaje]}")
                
        except Exception as e:
            st.error(f"❌ Error buscando valor de {nombre_peaje}: {str(e)}")
            peajes[nombre_peaje] = None
    
    return peajes

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI - VERSIÓN EXTENDIDA CON PEAJES"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiYTFmOWZkMDAtY2IwYi00OTg4LWIxZDctNGZmYmU0NTMxNGI1IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None
    
    try:
        # 1. Navegar al reporte
        with st.spinner("🌐 Conectando con Power BI..."):
            driver.get(REPORT_URL)
            time.sleep(10)
        
        # 2. Tomar screenshot inicial
        driver.save_screenshot("powerbi_inicial.png")
        
        # 3. Hacer clic en la conciliación específica
        if not click_conciliacion_date(driver, fecha_objetivo):
            return None
        
        # 4. Esperar a que cargue la selección
        time.sleep(3)
        driver.save_screenshot("powerbi_despues_seleccion.png")
        
        # 5. Buscar tarjeta "VALOR A PAGAR A COMERCIO" y extraer valor
        valor_texto = find_valor_a_pagar_comercio_card(driver)
        
        # 6. NUEVA FUNCIONALIDAD: Extraer valores por peaje
        st.info("🔍 Buscando valores individuales por peaje...")
        valores_peajes = find_peaje_values(driver)
        
        # 7. Tomar screenshot final
        driver.save_screenshot("powerbi_final.png")
        
        return {
            'valor_texto': valor_texto,
            'valores_peajes': valores_peajes,  # NUEVO: Valores por peaje
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

# ===== FUNCIONES DE EXTRACCIÓN DE EXCEL (MANTENIDAS) =====

def extract_excel_values(uploaded_file):
    """Extraer valores de las 3 hojas del Excel - VERSIÓN MEJORADA"""
    try:
        st.info("📊 Procesando archivo Excel...")
        
        hojas = ['CHICORAL', 'GUALANDAY', 'COCORA']
        valores = {}
        total_general = 0
        
        for hoja in hojas:
            try:
                df = pd.read_excel(uploaded_file, sheet_name=hoja, header=None)
                
                # Buscar el ÚLTIMO "Total" en la hoja
                valor_encontrado = None
                mejor_candidato = None
                mejor_puntaje = -1
                
                # Buscar de ABAJO hacia ARRIBA
                for i in range(len(df)-1, -1, -1):
                    fila = df.iloc[i]
                    
                    # Buscar "Total" en esta fila
                    for j, celda in enumerate(fila):
                        if pd.notna(celda) and isinstance(celda, str) and 'TOTAL' in celda.upper().strip():
                            
                            # Buscar valores monetarios en la MISMA fila
                            for k in range(len(fila)):
                                posible_valor = fila.iloc[k]
                                if pd.notna(posible_valor):
                                    valor_str = str(posible_valor)
                                    
                                    # Calcular puntaje
                                    puntaje = 0
                                    if '$' in valor_str:
                                        puntaje += 10
                                    if any(c.isdigit() for c in valor_str):
                                        puntaje += 5
                                    if '.' in valor_str and len(valor_str.split('.')[-1]) == 3:
                                        puntaje += 3
                                    if len(valor_str) > 6:
                                        puntaje += 2
                                    
                                    # Excluir valores incorrectos
                                    if puntaje > 0 and len(valor_str) < 4:
                                        puntaje = 0
                                    if 'pag' in valor_str.lower():
                                        puntaje = 0
                                    
                                    if puntaje > mejor_puntaje:
                                        mejor_puntaje = puntaje
                                        mejor_candidato = posible_valor
                
                # Usar el mejor candidato
                if mejor_candidato is not None and mejor_puntaje >= 5:
                    valor_encontrado = mejor_candidato
                else:
                    # Búsqueda alternativa en últimas filas
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
                
                # Procesar el valor encontrado
                if valor_encontrado is not None:
                    valor_original = str(valor_encontrado)
                    valor_limpio = re.sub(r'[^\d.,]', '', valor_original)
                    
                    try:
                        # Para formato colombiano
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
                            st.success(f"✅ {hoja}: ${valor_numerico:,.0f}".replace(",", "."))
                        else:
                            st.warning(f"⚠️ Valor muy pequeño en {hoja}, usando 0")
                            valores[hoja] = 0
                            
                    except Exception as conv_error:
                        st.error(f"❌ Error convirtiendo valor de {hoja}: {valor_original}")
                        valores[hoja] = 0
                else:
                    st.error(f"❌ No se encontró valor válido en {hoja}")
                    valores[hoja] = 0
                    
            except Exception as e:
                st.error(f"❌ Error procesando hoja {hoja}: {str(e)}")
                valores[hoja] = 0
        
        return valores, total_general
        
    except Exception as e:
        st.error(f"❌ Error procesando archivo Excel: {str(e)}")
        return {}, 0

# ===== FUNCIONES DE COMPARACIÓN (ACTUALIZADAS) =====

def convert_currency_to_float(currency_string):
    """Convierte string de moneda a float - OPTIMIZADO"""
    try:
        if isinstance(currency_string, (int, float)):
            return float(currency_string)
            
        if isinstance(currency_string, str):
            # Limpiar el string
            cleaned = currency_string.strip()
            
            # Remover símbolos de moneda y espacios
            cleaned = cleaned.replace('$', '').replace(' ', '')
            
            # Manejar formato colombiano (puntos para miles, coma para decimales)
            if '.' in cleaned and ',' in cleaned:
                # Formato: 1.000.000,00 -> quitar puntos, cambiar coma por punto
                cleaned = cleaned.replace('.', '').replace(',', '.')
            elif '.' in cleaned and cleaned.count('.') > 1:
                # Formato: 1.000.000 -> quitar todos los puntos
                cleaned = cleaned.replace('.', '')
            elif ',' in cleaned:
                # Formato: 1,000,000 o 1,000,000.00
                if cleaned.count(',') == 2 and '.' in cleaned:
                    # Formato internacional: 1,000,000.00
                    cleaned = cleaned.replace(',', '')
                elif cleaned.count(',') == 1:
                    # Podría ser decimal: 1000,50
                    cleaned = cleaned.replace(',', '.')
                else:
                    # Múltiples comas como separadores de miles
                    cleaned = cleaned.replace(',', '')
            
            # Convertir a float
            return float(cleaned) if cleaned else 0.0
            
        return float(currency_string)
        
    except Exception as e:
        st.error(f"❌ Error convirtiendo moneda: '{currency_string}' - {e}")
        return 0.0

def compare_values(valor_powerbi, valor_excel):
    """Comparar valores de Power BI y Excel - VERSIÓN MEJORADA"""
    try:
        # Si es un diccionario (resultado de extracción)
        if isinstance(valor_powerbi, dict):
            valor_powerbi_texto = valor_powerbi.get('valor_texto', '')
            powerbi_numero = convert_currency_to_float(valor_powerbi_texto)
        else:
            # Convertir texto a número
            valor_powerbi_texto = str(valor_powerbi)
            powerbi_numero = convert_currency_to_float(valor_powerbi)
            
        excel_numero = float(valor_excel)
        
        # Verificar coincidencia (con tolerancia pequeña por redondeos)
        tolerancia = 0.01  # 1 centavo
        coinciden = abs(powerbi_numero - excel_numero) <= tolerancia
        
        return powerbi_numero, excel_numero, valor_powerbi_texto, coinciden
        
    except Exception as e:
        st.error(f"❌ Error comparando valores: {e}")
        return None, None, str(valor_powerbi), False

def compare_peajes(valores_powerbi_peajes, valores_excel):
    """
    NUEVA FUNCIÓN: Comparar valores individuales por peaje
    """
    comparaciones = {}
    
    for peaje in ['CHICORAL', 'COCORA', 'GUALANDAY']:
        try:
            # Valor de Power BI
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
            
            # Convertir valores
            powerbi_numero = convert_currency_to_float(valor_powerbi_texto)
            excel_numero = valores_excel.get(peaje, 0)
            
            # Comparar con tolerancia
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
    
    # Información del reporte
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Cargar archivo Excel con 3 hojas
    - Extraer valores de CHICORAL, GUALANDAY, COCORA
    - Calcular total automáticamente
    - Comparar con Power BI (Total y por Peaje)
    
    **Estado:** ✅ ChromeDriver Compatible
    **Versión:** v2.0 - Con Comparación por Peaje
    """)
    
    # Estado del sistema
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}")
    st.sidebar.info(f"✅ Pandas {pd.__version__}")
    st.sidebar.info(f"✅ Streamlit {st.__version__}")
    
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
            fecha_conciliacion = st.date_input(
                "Fecha de Conciliación",
                value=pd.to_datetime("2025-09-04")
            )
            
            fecha_objetivo = fecha_conciliacion.strftime("%Y-%m-%d")
            
            # Botón de extracción
            st.markdown("---")
            st.subheader("🚀 Extracción y Validación")
            
            if st.button("🎯 Extraer Valores de Power BI y Comparar", type="primary", use_container_width=True):
                with st.spinner("🌐 Extrayendo datos de Power BI... Esto puede tomar 1-2 minutos"):
                    resultados = extract_powerbi_data(fecha_objetivo)
                    
                    if resultados and resultados.get('valor_texto'):
                        valor_powerbi_texto = resultados['valor_texto']
                        valores_peajes_powerbi = resultados.get('valores_peajes', {})
                        
                        st.success("✅ Extracción completada!")
                        
                        # ========== COMPARACIÓN DEL TOTAL GENERAL ==========
                        st.markdown("---")
                        st.subheader("💰 COMPARACIÓN: VALOR TOTAL A PAGAR A COMERCIO")
                        
                        st.info(f"**Valor en Power BI:** {valor_powerbi_texto}")
                        
                        # Comparar valores totales
                        powerbi_numero, excel_numero, valor_formateado, coinciden = compare_values(
                            resultados, 
                            total_general
                        )
                        
                        if powerbi_numero is not None and excel_numero is not None:
                            # Mostrar comparación total
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Power BI (Total)", valor_formateado)
                            with col2:
                                st.metric("Excel (Total)", total_formateado)
                            with col3:
                                if coinciden:
                                    st.success("✅ COINCIDEN")
                                else:
                                    diferencia = abs(powerbi_numero - excel_numero)
                                    diferencia_formateada = f"${diferencia:,.0f}".replace(",", ".")
                                    st.error("❌ NO COINCIDEN")
                                    st.metric("Diferencia", diferencia_formateada)
                            
                            # Mostrar detalles del total
                            with st.expander("📊 Detalles de la Comparación Total"):
                                st.write(f"**Power BI (numérico):** {powerbi_numero:,.0f}".replace(",", "."))
                                st.write(f"**Excel (numérico):** {excel_numero:,.0f}".replace(",", "."))
                                st.write(f"**Diferencia absoluta:** {abs(powerbi_numero - excel_numero):,.0f}".replace(",", "."))
                                if excel_numero > 0:
                                    st.write(f"**Diferencia relativa:** {abs(powerbi_numero - excel_numero)/excel_numero*100:.2f}%")
                        
                        # ========== NUEVA SECCIÓN: COMPARACIÓN POR PEAJE ==========
                        st.markdown("---")
                        st.subheader("🏢 COMPARACIÓN: VALORES INDIVIDUALES POR PEAJE")
                        
                        # Comparar valores por peaje
                        comparaciones_peajes = compare_peajes(valores_peajes_powerbi, valores)
                        
                        # Mostrar comparaciones en columnas
                        col1, col2, col3 = st.columns(3)
                        
                        peajes_columnas = {
                            'CHICORAL': col1,
                            'GUALANDAY': col2,
                            'COCORA': col3
                        }
                        
                        todos_coinciden = True
                        
                        for peaje, col in peajes_columnas.items():
                            comparacion = comparaciones_peajes[peaje]
                            
                            with col:
                                st.markdown(f"### 📍 {peaje}")
                                
                                # Power BI
                                st.write("**Power BI:**")
                                st.write(f"{comparacion['powerbi_texto']}")
                                
                                # Excel
                                st.write("**Excel:**")
                                excel_formateado = f"${comparacion['excel_numero']:,.0f}".replace(",", ".")
                                st.write(excel_formateado)
                                
                                # Estado
                                if comparacion['coinciden']:
                                    st.success("✅ COINCIDEN")
                                else:
                                    st.error("❌ NO COINCIDEN")
                                    diferencia_formateada = f"${comparacion['diferencia']:,.0f}".replace(",", ".")
                                    st.write(f"**Diferencia:** {diferencia_formateada}")
                                    todos_coinciden = False
                        
                        # Resumen final
                        st.markdown("---")
                        st.subheader("📋 Resumen General de Validación")
                        
                        if coinciden and todos_coinciden:
                            st.success("🎉 ¡VALIDACIÓN EXITOSA! Todos los valores coinciden (Total y por Peaje)")
                            st.balloons()
                        elif coinciden and not todos_coinciden:
                            st.warning("⚠️ El TOTAL GENERAL coincide, pero hay diferencias en los valores individuales por peaje")
                        elif not coinciden and todos_coinciden:
                            st.warning("⚠️ Los valores por PEAJE coinciden, pero el TOTAL GENERAL tiene diferencias")
                        else:
                            st.error("❌ Se encontraron diferencias tanto en el TOTAL como en los valores por PEAJE")
                        
                        # Tabla resumen detallada
                        with st.expander("📊 Ver Tabla Resumen Detallada"):
                            resumen_data = []
                            
                            # Agregar totales
                            resumen_data.append({
                                'Concepto': 'TOTAL GENERAL',
                                'Power BI': f"${powerbi_numero:,.0f}".replace(",", "."),
                                'Excel': f"${excel_numero:,.0f}".replace(",", "."),
                                'Estado': '✅ Coincide' if coinciden else '❌ No coincide',
                                'Diferencia': f"${abs(powerbi_numero - excel_numero):,.0f}".replace(",", ".")
                            })
                            
                            # Agregar peajes
                            for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                                comp = comparaciones_peajes[peaje]
                                resumen_data.append({
                                    'Concepto': peaje,
                                    'Power BI': comp['powerbi_texto'],
                                    'Excel': f"${comp['excel_numero']:,.0f}".replace(",", "."),
                                    'Estado': '✅ Coincide' if comp['coinciden'] else '❌ No coincide',
                                    'Diferencia': f"${comp['diferencia']:,.0f}".replace(",", ".")
                                })
                            
                            df_resumen = pd.DataFrame(resumen_data)
                            st.dataframe(df_resumen, use_container_width=True)
                        
                        # Mostrar screenshots
                        with st.expander("📸 Ver capturas del proceso"):
                            col1, col2, col3 = st.columns(3)
                            screenshots = resultados.get('screenshots', {})
                            
                            if 'inicial' in screenshots and os.path.exists(screenshots['inicial']):
                                with col1:
                                    st.image(screenshots['inicial'], caption="Reporte Inicial", use_column_width=True)
                            
                            if 'seleccion' in screenshots and os.path.exists(screenshots['seleccion']):
                                with col2:
                                    st.image(screenshots['seleccion'], caption="Después de Selección", use_column_width=True)
                            
                            if 'final' in screenshots and os.path.exists(screenshots['final']):
                                with col3:
                                    st.image(screenshots['final'], caption="Vista Final", use_column_width=True)
                                
                    elif resultados:
                        st.error("❌ Se accedió al reporte pero no se encontró el valor específico")
                    else:
                        st.error("❌ No se pudieron extraer datos del reporte Power BI")
        else:
            st.error("❌ No se pudieron extraer valores del archivo Excel.")
            st.info("💡 **Sugerencias:**")
            st.info("- Verifica que las hojas se llamen CHICORAL, GUALANDAY, COCORA")
            st.info("- Asegúrate de que haya valores numéricos en las celdas de total")
            st.info("- Revisa que los totales estén claramente identificados con 'TOTAL'")
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar la validación")

    # Información de ayuda
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso:**
        1. **Cargar Excel**: Archivo con hojas CHICORAL, GUALANDAY, COCORA
        2. **Extracción automática**: Búsqueda inteligente de "Total" en cada hoja
        3. **Seleccionar fecha** de conciliación en Power BI  
        4. **Comparar**: Extrae valores de Power BI y compara con Excel
        
        **Características NUEVAS (v2.0):**
        - ✅ **Comparación Total**: Valida el "VALOR A PAGAR A COMERCIO" total
        - ✅ **Comparación por Peaje**: Valida valores individuales de CHICORAL, COCORA y GUALANDAY
        - ✅ **Resumen Detallado**: Tabla completa con todas las comparaciones
        - ✅ **Validación Dual**: Verifica coincidencias tanto en total como por peaje
        
        **Características Mantenidas:**
        - ✅ **Power BI Funcional**: Usa la extracción probada que funciona
        - ✅ **Búsqueda inteligente**: Múltiples estrategias para encontrar valores
        - ✅ **Conversión de moneda**: Maneja formatos colombianos e internacionales
        - 📸 **Capturas del proceso**: Para verificación y debugging
        
        **Notas:**
        - La extracción busca el total y los valores individuales por peaje
        - Los valores deben estar claramente identificados en el Power BI
        - Las fechas deben coincidir exactamente con las del reporte Power BI
        """)

if __name__ == "__main__":
    main()

    
    # Footer
    st.markdown("---")
    st.markdown('<div class="footer">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v2.0</div>', unsafe_allow_html=True)
