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

def find_cantidad_pasos_card(driver):
    """Buscar la tarjeta/table 'CANTIDAD PASOS' a la derecha de 'VALOR A PAGAR A COMERCIO'"""
    try:
        st.info("🔍 Buscando 'CANTIDAD PASOS' en el reporte...")
        
        # Buscar por diferentes patrones del título - MÁS ESPECÍFICO
        titulo_selectors = [
            "//*[contains(text(), 'CANTIDAD PASOS')]",
            "//*[contains(text(), 'Cantidad Pasos')]",
            "//*[contains(text(), 'CANTIDAD DE PASOS')]",
            "//*[contains(text(), 'Cantidad de Pasos')]",
            "//*[contains(text(), 'CANTIDAD') and contains(text(), 'PASOS')]",
            "//*[text()='CANTIDAD PASOS']",
            "//*[text()='Cantidad Pasos']",
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
                            st.success(f"✅ Título encontrado: {texto}")
                            break
                if titulo_element:
                    break
            except Exception as e:
                continue
        
        if not titulo_element:
            st.warning("❌ No se encontró el título 'CANTIDAD PASOS'")
            return None
        
        # ESTRATEGIA MEJORADA: Buscar en el mismo contenedor o contenedores cercanos
        try:
            # Buscar en el contenedor padre
            container = titulo_element.find_element(By.XPATH, "./..")
            
            # Buscar TODOS los elementos numéricos en el contenedor
            all_elements = container.find_elements(By.XPATH, ".//*")
            
            for elem in all_elements:
                texto = elem.text.strip()
                # Verificar si es un número (contiene dígitos pero no texto largo)
                if (texto and 
                    any(char.isdigit() for char in texto) and 
                    len(texto) < 20 and 
                    texto != titulo_element.text and
                    not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
                    
                    # Verificar formato numérico (puede tener comas, puntos, pero ser principalmente números)
                    digit_count = sum(char.isdigit() for char in texto)
                    if digit_count >= 1:  # Al menos un dígito
                        st.success(f"✅ Valor numérico encontrado: {texto}")
                        return texto
                        
        except Exception as e:
            st.warning(f"⚠️ Estrategia 1 falló: {e}")
        
        # ESTRATEGIA 2: Buscar elementos hermanos específicamente
        try:
            parent = titulo_element.find_element(By.XPATH, "./..")
            siblings = parent.find_elements(By.XPATH, "./*")
            
            for sibling in siblings:
                if sibling != titulo_element:
                    texto = sibling.text.strip()
                    if (texto and 
                        any(char.isdigit() for char in texto) and 
                        len(texto) < 20 and
                        not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
                        
                        digit_count = sum(char.isdigit() for char in texto)
                        if digit_count >= 1:
                            st.success(f"✅ Valor encontrado en hermano: {texto}")
                            return texto
        except Exception as e:
            st.warning(f"⚠️ Estrategia 2 falló: {e}")
        
        # ESTRATEGIA 3: Buscar elementos que siguen al título
        try:
            # Buscar elementos que están después del título
            following_elements = driver.find_elements(By.XPATH, f"//*[contains(text(), 'CANTIDAD PASOS')]/following::*")
            
            for i, elem in enumerate(following_elements[:20]):  # Buscar en los primeros 20 elementos siguientes
                texto = elem.text.strip()
                if (texto and 
                    any(char.isdigit() for char in texto) and 
                    len(texto) < 20 and
                    not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
                    
                    digit_count = sum(char.isdigit() for char in texto)
                    if digit_count >= 1:
                        st.success(f"✅ Valor encontrado en elemento siguiente {i}: {texto}")
                        return texto
        except Exception as e:
            st.warning(f"⚠️ Estrategia 3 falló: {e}")
        
        # ESTRATEGIA 4: Buscar cerca de "VALOR A PAGAR A COMERCIO"
        try:
            # Encontrar "VALOR A PAGAR A COMERCIO" primero
            valor_element = driver.find_element(By.XPATH, "//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]")
            if valor_element:
                # Buscar elementos a la derecha o cerca
                container_valor = valor_element.find_element(By.XPATH, "./..")
                # Buscar en el mismo nivel jerárquico
                all_nearby = container_valor.find_elements(By.XPATH, ".//*")
                
                for elem in all_nearby:
                    texto = elem.text.strip()
                    if (texto and 
                        any(char.isdigit() for char in texto) and 
                        len(texto) < 20 and
                        'CANTIDAD' in texto.upper() and 'PASOS' in texto.upper()):
                        # Este es el título, buscar el siguiente elemento numérico
                        continue
                    
                    if (texto and 
                        any(char.isdigit() for char in texto) and 
                        len(texto) < 20 and
                        not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO'])):
                        
                        digit_count = sum(char.isdigit() for char in texto)
                        if digit_count >= 1:
                            st.success(f"✅ Valor encontrado cerca de VALOR A PAGAR: {texto}")
                            return texto
        except Exception as e:
            st.warning(f"⚠️ Estrategia 4 falló: {e}")
        
        st.error("❌ No se pudo encontrar el valor numérico de CANTIDAD PASOS")
        return None
        
    except Exception as e:
        st.error(f"❌ Error buscando cantidad de pasos: {str(e)}")
        return None

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
    """
    NUEVA FUNCIÓN: Buscar valores individuales de cada peaje en el Power BI (VERSIÓN SILENCIOSA)
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
                    numeric_elements = container.find_elements(By.XPATH, ".//*[contains(text(), '$') or contains(text(), ',') or contains(text(), '.')]")
                    
                    for elem in numeric_elements:
                        texto = elem.text.strip()
                        if texto and any(char.isdigit() for char in texto):
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
                        if texto and any(char.isdigit() for char in texto):
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
                            if texto and any(char.isdigit() for char in texto):
                                if len(texto) > 3 and len(texto) < 30:
                                    peajes[nombre_peaje] = texto
                                    break
                except:
                    pass
                    
        except Exception as e:
            peajes[nombre_peaje] = None
    
    return peajes

def find_pasos_por_peaje_bi(driver):
    """
    NUEVA FUNCIÓN: Buscar cantidad de pasos por peaje en la tabla "RESUMEN COMERCIOS"
    """
    try:
        st.info("🔍 Buscando tabla 'RESUMEN COMERCIOS' para pasos por peaje...")
        
        pasos_peajes = {}
        nombres_peajes = ['CHICORAL', 'COCORA', 'GUALANDAY']
        total_pasos_bi = 0
        
        # Buscar la tabla "RESUMEN COMERCIOS"
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
                        st.success("✅ Tabla 'RESUMEN COMERCIOS' encontrada")
                        break
                if tabla_element:
                    break
            except:
                continue
        
        if not tabla_element:
            st.warning("❌ No se encontró la tabla 'RESUMEN COMERCIOS'")
            return {}, 0
        
        # Buscar en el contenedor de la tabla
        try:
            container = tabla_element.find_element(By.XPATH, "./ancestor::*[position()<=5]")
            
            for nombre_peaje in nombres_peajes:
                # Buscar el nombre del peaje en la tabla
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
                    # Buscar valores numéricos cerca del nombre del peaje (para Cant Pasos)
                    try:
                        # Buscar en la misma fila o contenedor
                        fila_element = peaje_element.find_element(By.XPATH, "./ancestor::*[position()<=3]")
                        
                        # Buscar todos los elementos numéricos en la misma fila
                        numeric_elements = fila_element.find_elements(By.XPATH, ".//*[text()]")
                        
                        for elem in numeric_elements:
                            texto = elem.text.strip()
                            # Verificar si es un número válido para pasos
                            if (texto and 
                                any(char.isdigit() for char in texto) and
                                1 <= len(texto) <= 6 and  # Los pasos son números de 1-6 dígitos
                                texto != peaje_element.text and
                                not any(word in texto.upper() for word in ['CHICORAL', 'COCORA', 'GUALANDAY', 'TOTAL', 'RESUMEN'])):
                                
                                # Limpiar y convertir
                                pasos_limpio = re.sub(r'[^\d]', '', texto)
                                if pasos_limpio and pasos_limpio.isdigit():
                                    num_pasos = int(pasos_limpio)
                                    if 1 <= num_pasos <= 999999:  # Rango razonable para pasos
                                        pasos_peajes[nombre_peaje] = num_pasos
                                        total_pasos_bi += num_pasos
                                        st.success(f"✅ Pasos BI {nombre_peaje}: {num_pasos}")
                                        break
                        
                        # Si no se encontró en la misma fila, buscar elementos cercanos
                        if nombre_peaje not in pasos_peajes:
                            following_elements = peaje_element.find_elements(By.XPATH, "./following::*")
                            for elem in following_elements[:10]:
                                texto = elem.text.strip()
                                if (texto and 
                                    any(char.isdigit() for char in texto) and
                                    1 <= len(texto) <= 6 and
                                    not any(word in texto.upper() for word in ['CHICORAL', 'COCORA', 'GUALANDAY', 'TOTAL', 'RESUMEN'])):
                                    
                                    pasos_limpio = re.sub(r'[^\d]', '', texto)
                                    if pasos_limpio and pasos_limpio.isdigit():
                                        num_pasos = int(pasos_limpio)
                                        if 1 <= num_pasos <= 999999:
                                            pasos_peajes[nombre_peaje] = num_pasos
                                            total_pasos_bi += num_pasos
                                            st.success(f"✅ Pasos BI {nombre_peaje} (cercano): {num_pasos}")
                                            break
                    
                    except Exception as e:
                        st.warning(f"⚠️ Error buscando pasos para {nombre_peaje}: {e}")
                        pasos_peajes[nombre_peaje] = 0
                else:
                    st.warning(f"⚠️ No se encontró el peaje {nombre_peaje} en la tabla")
                    pasos_peajes[nombre_peaje] = 0
            
            st.success(f"✅ Total pasos BI encontrados: {total_pasos_bi}")
            return pasos_peajes, total_pasos_bi
            
        except Exception as e:
            st.error(f"❌ Error procesando tabla RESUMEN COMERCIOS: {e}")
            return {}, 0
            
    except Exception as e:
        st.error(f"❌ Error buscando pasos por peaje BI: {e}")
        return {}, 0

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
        
        # 6. NUEVA FUNCIONALIDAD: Extraer "CANTIDAD PASOS" - CON MÁS DETALLE
        st.info("🔍 Buscando tabla 'CANTIDAD PASOS'...")
        cantidad_pasos_texto = find_cantidad_pasos_card(driver)
        
        # Si no se encuentra, intentar una búsqueda más agresiva
        if not cantidad_pasos_texto or cantidad_pasos_texto == 'No encontrado':
            st.warning("🔄 Intentando búsqueda alternativa para CANTIDAD PASOS...")
            cantidad_pasos_texto = buscar_cantidad_pasos_alternativo(driver)
        
        # 7. NUEVA FUNCIONALIDAD: Extraer valores por peaje (SIN MENSAJES)
        valores_peajes = find_peaje_values(driver)
        
        # 8. NUEVA FUNCIONALIDAD: Extraer pasos por peaje del BI
        pasos_peajes_bi, total_pasos_bi = find_pasos_por_peaje_bi(driver)
        
        # 9. Tomar screenshot final
        driver.save_screenshot("powerbi_final.png")
        
        return {
            'valor_texto': valor_texto,
            'cantidad_pasos_texto': cantidad_pasos_texto or 'No encontrado',
            'valores_peajes': valores_peajes,
            'pasos_peajes_bi': pasos_peajes_bi,  # NUEVO: Pasos por peaje del BI
            'total_pasos_bi': total_pasos_bi,    # NUEVO: Total pasos del BI
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

# Función alternativa de búsqueda
def buscar_cantidad_pasos_alternativo(driver):
    """Búsqueda alternativa y más agresiva para CANTIDAD PASOS"""
    try:
        # Buscar todos los elementos que contengan números
        all_elements = driver.find_elements(By.XPATH, "//*[text()]")
        
        for elem in all_elements:
            texto = elem.text.strip()
            # Buscar patrones numéricos que parezcan cantidades (4,452, 4452, etc.)
            if (texto and 
                any(char.isdigit() for char in texto) and
                3 <= len(texto) <= 10 and
                not any(word in texto.upper() for word in ['$', 'TOTAL', 'VALOR', 'PAGAR', 'COMERCIO'])):
                
                # Verificar si es un número con formato de cantidad (puede tener comas)
                clean_text = texto.replace(',', '').replace('.', '')
                if clean_text.isdigit():
                    num_value = int(clean_text)
                    # Verificar si está en un rango razonable para cantidad de pasos
                    if 100 <= num_value <= 999999:
                        st.success(f"✅ Valor alternativo encontrado: {texto}")
                        return texto
        
        return None
    except Exception as e:
        st.warning(f"⚠️ Búsqueda alternativa falló: {e}")
        return None

# ===== FUNCIONES DE EXTRACCIÓN DE EXCEL (MEJORADAS) =====

def extract_excel_values(uploaded_file):
    """Extraer valores monetarios Y cantidad de pasos de las 3 hojas del Excel - VERSIÓN ROBUSTA"""
    try:
        hojas = ['CHICORAL', 'GUALANDAY', 'COCORA']
        valores = {}
        pasos = {}
        total_general = 0
        total_pasos = 0
        
        for hoja in hojas:
            try:
                # Leer el archivo Excel
                df = pd.read_excel(uploaded_file, sheet_name=hoja, header=None)
                
                st.info(f"🔍 Procesando hoja: {hoja} ({len(df)} filas, {len(df.columns)} columnas)")
                
                # Variables para almacenar resultados
                valor_encontrado = None
                pasos_encontrados = None
                
                # ESTRATEGIA 1: Buscar en las últimas filas (más común)
                st.write(f"📋 **Estrategia 1**: Buscando en últimas filas de {hoja}")
                for i in range(len(df)-1, max(len(df)-20, -1), -1):
                    fila = df.iloc[i]
                    
                    for j, celda in enumerate(fila):
                        if pd.notna(celda):
                            texto = str(celda).strip()
                            
                            # BUSCAR VALOR MONETARIO
                            if valor_encontrado is None:
                                if (any(char.isdigit() for char in texto) and 
                                    len(texto) > 3 and  # Valores monetarios son más largos
                                    ('$' in texto or ',' in texto or '.' in texto or 'TOTAL' in texto.upper())):
                                    
                                    valor_limpio = clean_currency_value(texto)
                                    if valor_limpio and valor_limpio >= 1000:
                                        valor_encontrado = valor_limpio
                                        st.success(f"💰 Valor {hoja}: {texto} -> ${valor_limpio:,.0f}")
                            
                            # BUSCAR CANTIDAD DE PASOS
                            if pasos_encontrados is None:
                                if (any(char.isdigit() for char in texto) and 
                                    1 <= len(texto) <= 10 and  # Pasos son números más cortos
                                    not any(word in texto.upper() for word in ['$', 'TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'PASOS'])):
                                    
                                    pasos_limpio = clean_step_value(texto)
                                    if pasos_limpio and 1 <= pasos_limpio <= 999999:
                                        pasos_encontrados = pasos_limpio
                                        st.success(f"👣 Pasos {hoja}: {texto} -> {pasos_limpio:,}")
                    
                    # Si encontramos ambos, salir del loop
                    if valor_encontrado is not None and pasos_encontrados is not None:
                        break
                
                # ESTRATEGIA 2: Buscar por patrones específicos si no se encontró
                if valor_encontrado is None:
                    st.write(f"🔄 **Estrategia 2**: Búsqueda por patrones en {hoja}")
                    valor_encontrado = find_value_by_patterns(df, hoja, 'valor')
                
                if pasos_encontrados is None:
                    pasos_encontrados = find_value_by_patterns(df, hoja, 'pasos')
                
                # ESTRATEGIA 3: Buscar en toda la hoja si aún no se encuentra
                if valor_encontrado is None:
                    st.write(f"🔎 **Estrategia 3**: Búsqueda exhaustiva en {hoja}")
                    valor_encontrado = exhaustive_value_search(df, hoja, 'valor')
                
                if pasos_encontrados is None:
                    pasos_encontrados = exhaustive_value_search(df, hoja, 'pasos')
                
                # ESTRATEGIA 4: Buscar en celdas con formato de moneda
                if valor_encontrado is None:
                    valor_encontrado = find_currency_formatted_cells(df, hoja)
                
                # ASIGNAR VALORES FINALES
                if valor_encontrado is not None:
                    valores[hoja] = valor_encontrado
                    total_general += valor_encontrado
                else:
                    # Último recurso: buscar cualquier número grande
                    backup_valor = find_backup_value(df, hoja)
                    if backup_valor:
                        valores[hoja] = backup_valor
                        total_general += backup_valor
                        st.warning(f"⚠️ Valor {hoja} (backup): ${backup_valor:,.0f}")
                    else:
                        valores[hoja] = 0
                        st.error(f"❌ No se encontró valor monetario para {hoja}")
                
                if pasos_encontrados is not None:
                    pasos[hoja] = pasos_encontrados
                    total_pasos += pasos_encontrados
                else:
                    # Último recurso: buscar cualquier número de pasos razonable
                    backup_pasos = find_backup_steps(df, hoja)
                    if backup_pasos:
                        pasos[hoja] = backup_pasos
                        total_pasos += backup_pasos
                        st.warning(f"⚠️ Pasos {hoja} (backup): {backup_pasos:,}")
                    else:
                        pasos[hoja] = 0
                        st.error(f"❌ No se encontró cantidad de pasos para {hoja}")
                        
            except Exception as e:
                st.error(f"❌ Error procesando hoja {hoja}: {str(e)}")
                valores[hoja] = 0
                pasos[hoja] = 0
        
        # Mostrar resumen final
        st.success(f"📊 Resumen extracción - Valores: ${total_general:,.0f}, Pasos: {total_pasos:,}")
        
        return valores, total_general, pasos, total_pasos
        
    except Exception as e:
        st.error(f"❌ Error procesando archivo Excel: {str(e)}")
        return {}, 0, {}, 0

# ===== FUNCIONES AUXILIARES MEJORADAS =====

def clean_currency_value(texto):
    """Limpia y convierte valores monetarios - VERSIÓN ROBUSTA"""
    try:
        if pd.isna(texto) or texto == '':
            return None
            
        texto_str = str(texto).strip()
        
        # Si ya es numérico
        if isinstance(texto, (int, float)) and texto > 0:
            return float(texto)
        
        # Remover texto no numérico pero preservar puntos y comas
        cleaned = re.sub(r'[^\d.,]', '', texto_str)
        
        if not cleaned:
            return None
        
        # Analizar formato
        if '.' in cleaned and ',' in cleaned:
            # Formato: 1.000.000,00 -> 1000000.00
            if cleaned.rfind('.') < cleaned.rfind(','):
                # El punto es separador de miles, coma es decimal
                cleaned = cleaned.replace('.', '').replace(',', '.')
            else:
                # La coma es separador de miles, punto es decimal
                cleaned = cleaned.replace(',', '')
        elif '.' in cleaned and cleaned.count('.') == 1:
            # Podría ser decimal simple: 1000.50
            pass
        elif ',' in cleaned and cleaned.count(',') == 1:
            # Podría ser decimal con coma: 1000,50
            cleaned = cleaned.replace(',', '.')
        elif '.' in cleaned and cleaned.count('.') > 1:
            # Múltiples puntos como separadores de miles: 1.000.000
            cleaned = cleaned.replace('.', '')
        elif ',' in cleaned and cleaned.count(',') > 1:
            # Múltiples comas como separadores de miles: 1,000,000
            cleaned = cleaned.replace(',', '')
        
        # Convertir a float
        result = float(cleaned)
        return result if result >= 0 else None
        
    except Exception as e:
        return None

def clean_step_value(texto):
    """Limpia y convierte valores de pasos"""
    try:
        if pd.isna(texto) or texto == '':
            return None
            
        texto_str = str(texto).strip()
        
        # Si ya es numérico
        if isinstance(texto, (int, float)) and texto > 0:
            return int(texto)
        
        # Remover todo excepto dígitos
        cleaned = re.sub(r'[^\d]', '', texto_str)
        
        if not cleaned:
            return None
        
        result = int(cleaned)
        return result if 1 <= result <= 999999 else None
        
    except Exception:
        return None

def find_value_by_patterns(df, hoja, tipo='valor'):
    """Busca valores por patrones específicos"""
    try:
        # Patrones de búsqueda para valores
        valor_patterns = [
            r'.*TOTAL.*', r'.*TOTAL GENERAL.*', r'.*VALOR.*', 
            r'.*PAGAR.*', r'.*MONTO.*', r'.*IMPORTE.*'
        ]
        
        # Patrones de búsqueda para pasos
        paso_patterns = [
            r'.*PASOS.*', r'.*CANTIDAD.*', r'.*CANT.*', 
            r'.*TOTAL.*', r'.*CONTEO.*', r'.*COUNT.*'
        ]
        
        patterns = valor_patterns if tipo == 'valor' else paso_patterns
        
        # Buscar en todas las celdas
        for i in range(len(df)):
            for j in range(len(df.columns)):
                celda = df.iloc[i, j]
                if pd.notna(celda):
                    texto = str(celda).strip().upper()
                    
                    # Verificar si coincide con algún patrón
                    for pattern in patterns:
                        if re.search(pattern, texto, re.IGNORECASE):
                            # Buscar valores numéricos en la misma fila o columna
                            if tipo == 'valor':
                                # Buscar en la misma fila (derecha)
                                for k in range(j+1, min(j+6, len(df.columns))):
                                    valor_celda = df.iloc[i, k]
                                    if pd.notna(valor_celda):
                                        valor = clean_currency_value(valor_celda)
                                        if valor and valor >= 1000:
                                            st.success(f"💰 Valor {hoja} (patrón): {valor_celda} -> ${valor:,.0f}")
                                            return valor
                            else:
                                # Buscar para pasos
                                for k in range(j+1, min(j+6, len(df.columns))):
                                    valor_celda = df.iloc[i, k]
                                    if pd.notna(valor_celda):
                                        pasos = clean_step_value(valor_celda)
                                        if pasos and pasos >= 1:
                                            st.success(f"👣 Pasos {hoja} (patrón): {valor_celda} -> {pasos:,}")
                                            return pasos
                                            
        return None
    except Exception as e:
        return None

def exhaustive_value_search(df, hoja, tipo='valor'):
    """Búsqueda exhaustiva en toda la hoja"""
    try:
        best_candidate = None
        best_value = 0
        
        for i in range(len(df)):
            for j in range(len(df.columns)):
                celda = df.iloc[i, j]
                if pd.notna(celda):
                    if tipo == 'valor':
                        valor = clean_currency_value(celda)
                        if valor and valor > best_value and valor >= 1000:
                            best_value = valor
                            best_candidate = valor
                    else:
                        pasos = clean_step_value(celda)
                        if pasos and pasos > best_value and 100 <= pasos <= 99999:
                            best_value = pasos
                            best_candidate = pasos
        
        if best_candidate:
            if tipo == 'valor':
                st.info(f"💰 Valor {hoja} (exhaustivo): ${best_candidate:,.0f}")
            else:
                st.info(f"👣 Pasos {hoja} (exhaustivo): {best_candidate:,}")
        
        return best_candidate
    except Exception:
        return None

def find_currency_formatted_cells(df, hoja):
    """Busca específicamente celdas con formato de moneda"""
    try:
        # Buscar celdas que empiecen con $
        for i in range(len(df)):
            for j in range(len(df.columns)):
                celda = df.iloc[i, j]
                if pd.notna(celda):
                    texto = str(celda).strip()
                    if texto.startswith('$'):
                        valor = clean_currency_value(texto)
                        if valor and valor >= 1000:
                            st.success(f"💰 Valor {hoja} (formato $): {texto} -> ${valor:,.0f}")
                            return valor
        return None
    except Exception:
        return None

def find_backup_value(df, hoja):
    """Último recurso para encontrar valores"""
    try:
        # Buscar los 5 números más grandes en la hoja
        all_values = []
        for i in range(len(df)):
            for j in range(len(df.columns)):
                celda = df.iloc[i, j]
                if pd.notna(celda):
                    valor = clean_currency_value(celda)
                    if valor and valor >= 1000:
                        all_values.append(valor)
        
        if all_values:
            # Tomar el más grande (probablemente el total)
            return max(all_values)
        return None
    except Exception:
        return None

def find_backup_steps(df, hoja):
    """Último recurso para encontrar pasos"""
    try:
        # Buscar números en un rango razonable para pasos
        all_steps = []
        for i in range(len(df)):
            for j in range(len(df.columns)):
                celda = df.iloc[i, j]
                if pd.notna(celda):
                    pasos = clean_step_value(celda)
                    if pasos and 100 <= pasos <= 99999:
                        all_steps.append(pasos)
        
        if all_steps:
            # Tomar el más común o el más grande
            from collections import Counter
            if all_steps:
                counter = Counter(all_steps)
                return counter.most_common(1)[0][0]
        return None
    except Exception:
        return None

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

def compare_pasos_peajes(pasos_peajes_bi, pasos_excel):
    """
    NUEVA FUNCIÓN: Comparar cantidad de pasos por peaje entre Power BI y Excel
    """
    comparaciones = {}
    
    for peaje in ['CHICORAL', 'COCORA', 'GUALANDAY']:
        try:
            # Valor de Power BI
            pasos_bi = pasos_peajes_bi.get(peaje, 0)
            pasos_excel_val = pasos_excel.get(peaje, 0)
            
            # Comparar
            coinciden = pasos_bi == pasos_excel_val
            diferencia = abs(pasos_bi - pasos_excel_val)
            
            comparaciones[peaje] = {
                'pasos_bi': pasos_bi,
                'pasos_excel': pasos_excel_val,
                'coinciden': coinciden,
                'diferencia': diferencia
            }
            
        except Exception as e:
            st.error(f"❌ Error comparando pasos {peaje}: {e}")
            comparaciones[peaje] = {
                'pasos_bi': 0,
                'pasos_excel': pasos_excel.get(peaje, 0),
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
    - Extraer valores Y pasos de CHICORAL, GUALANDAY, COCORA
    - Calcular totales automáticamente
    - Comparar con Power BI (Total, Pasos y por Peaje)
    
    **Estado:** ✅ ChromeDriver Compatible
    **Versión:** v2.3 - Con Pasos por Peaje BI
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
        # Extraer fecha del nombre del archivo (sin mostrar nada)
        fecha_desde_archivo = None
        try:
            import re
            patron_fecha = r'(\d{4})-(\d{2})-(\d{2})'
            match = re.search(patron_fecha, uploaded_file.name)
            
            if match:
                year, month, day = match.groups()
                fecha_desde_archivo = pd.to_datetime(f"{year}-{month}-{day}")
        except:
            pass
        
        # Extraer valores del Excel CON SPINNER
        with st.spinner("📊 Procesando archivo Excel..."):
            valores, total_general, pasos, total_pasos = extract_excel_values(uploaded_file)
        
        if total_general > 0 or total_pasos > 0:
            # ========== MOSTRAR SOLO RESUMEN DE VALORES ==========
            st.markdown("### 📊 Valores Extraídos del Excel")
            
            # Primera fila: Valores monetarios
            st.markdown("#### 💰 Valores Monetarios")
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
            
            # Segunda fila: Cantidad de Pasos
            st.markdown("#### 👣 Cantidad de Pasos")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("PASOS CHICORAL", f"{pasos['CHICORAL']:,}".replace(",", "."))
            
            with col2:
                st.metric("PASOS GUALANDAY", f"{pasos['GUALANDAY']:,}".replace(",", "."))
            
            with col3:
                st.metric("PASOS COCORA", f"{pasos['COCORA']:,}".replace(",", "."))
            
            with col4:
                st.metric("TOTAL PASOS", f"{total_pasos:,}".replace(",", "."), delta="Excel")
            
            st.markdown("---")
            
            # ========== SECCIÓN 3: PARÁMETROS Y EJECUCIÓN ==========
            # Usar la fecha del archivo si está disponible, sino usar fecha por defecto
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
            
            # Ejecutar extracción si corresponde
            if ejecutar_extraccion:
                with st.spinner("🌐 Extrayendo datos de Power BI... Esto puede tomar 1-2 minutos"):
                    resultados = extract_powerbi_data(fecha_objetivo)
                    
                    if resultados and resultados.get('valor_texto'):
                        valor_powerbi_texto = resultados['valor_texto']
                        cantidad_pasos_texto = resultados.get('cantidad_pasos_texto', 'No encontrado')
                        valores_peajes_powerbi = resultados.get('valores_peajes', {})
                        pasos_peajes_bi = resultados.get('pasos_peajes_bi', {})
                        total_pasos_bi = resultados.get('total_pasos_bi', 0)
                        
                        st.markdown("---")
                        
                        # ========== SECCIÓN 4: RESULTADOS - VALORES POWER BI ==========
                        st.markdown("### 📊 Valores Extraídos de Power BI")
                        
                        # Primera fila: Valores principales
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.metric("💰 VALOR A PAGAR A COMERCIO", valor_powerbi_texto)
                        
                        with col2:
                            st.metric("👣 CANTIDAD DE PASOS BI", cantidad_pasos_texto)
                        
                        # Segunda fila: Pasos por peaje del BI
                        st.markdown("#### 👣 Pasos por Peaje - Power BI")
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            st.metric("PASOS CHICORAL BI", f"{pasos_peajes_bi.get('CHICORAL', 0):,}".replace(",", "."))
                        
                        with col2:
                            st.metric("PASOS GUALANDAY BI", f"{pasos_peajes_bi.get('GUALANDAY', 0):,}".replace(",", "."))
                        
                        with col3:
                            st.metric("PASOS COCORA BI", f"{pasos_peajes_bi.get('COCORA', 0):,}".replace(",", "."))
                        
                        with col4:
                            st.metric("TOTAL PASOS BI", f"{total_pasos_bi:,}".replace(",", "."), delta="Power BI")
                        
                        st.markdown("---")
                        
                        # ========== SECCIÓN 5: RESULTADOS - COMPARACIÓN TOTAL ==========
                        st.markdown("### 💰 Validación: Total General")
                        
                        # Comparar valores totales
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
                        
                        # ========== SECCIÓN 6: RESULTADOS - COMPARACIÓN DE PASOS TOTALES ==========
                        st.markdown("### 👣 Validación: Cantidad de Pasos Totales")
                        
                        # Convertir cantidad de pasos de Power BI a número
                        cantidad_pasos_bi = 0
                        if cantidad_pasos_texto and cantidad_pasos_texto != 'No encontrado':
                            try:
                                # Limpiar el texto (remover comas, puntos, etc.)
                                pasos_limpio = re.sub(r'[^\d]', '', str(cantidad_pasos_texto))
                                if pasos_limpio:
                                    cantidad_pasos_bi = int(pasos_limpio)
                            except:
                                cantidad_pasos_bi = 0
                        
                        # Usar el total de pasos del BI si está disponible
                        if total_pasos_bi > 0:
                            cantidad_pasos_bi = total_pasos_bi
                        
                        # Comparar cantidad de pasos
                        coinciden_pasos = cantidad_pasos_bi == total_pasos
                        
                        col1, col2, col3 = st.columns([2, 2, 1])
                        
                        with col1:
                            st.metric("📊 Power BI", f"{cantidad_pasos_bi:,}".replace(",", "."))
                        with col2:
                            st.metric("📁 Excel", f"{total_pasos:,}".replace(",", "."))
                        with col3:
                            if coinciden_pasos:
                                st.markdown("#### ✅")
                                st.success("COINCIDE")
                            else:
                                diferencia_pasos = abs(cantidad_pasos_bi - total_pasos)
                                st.markdown("#### ❌")
                                st.error("DIFERENCIA")
                                st.caption(f"{diferencia_pasos:,}".replace(",", "."))
                        
                        st.markdown("---")
                        
                        # ========== SECCIÓN 7: RESULTADOS - COMPARACIÓN DE PASOS POR PEAJE ==========
                        st.markdown("### 🏢 Validación: Pasos por Peaje")
                        
                        # Comparar pasos por peaje
                        comparaciones_pasos_peajes = compare_pasos_peajes(pasos_peajes_bi, pasos)
                        
                        # Crear tabla resumen compacta
                        tabla_data = []
                        todos_pasos_coinciden = True
                        
                        for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                            comp = comparaciones_pasos_peajes[peaje]
                            
                            estado_icono = "✅" if comp['coinciden'] else "❌"
                            diferencia_texto = "0" if comp['coinciden'] else f"{comp['diferencia']:,}".replace(",", ".")
                            
                            tabla_data.append({
                                '': estado_icono,
                                'Peaje': peaje,
                                'Power BI': f"{comp['pasos_bi']:,}".replace(",", "."),
                                'Excel': f"{comp['pasos_excel']:,}".replace(",", "."),
                                'Dif.': diferencia_texto
                            })
                            
                            if not comp['coinciden']:
                                todos_pasos_coinciden = False
                        
                        df_comparacion_pasos = pd.DataFrame(tabla_data)
                        st.dataframe(df_comparacion_pasos, use_container_width=True, hide_index=True)
                        
                        st.markdown("---")
                        
                        # ========== SECCIÓN 8: RESULTADOS - COMPARACIÓN POR PEAJE (VALORES) ==========
                        st.markdown("### 💰 Validación: Valores por Peaje")
                        
                        # Comparar valores por peaje
                        comparaciones_peajes = compare_peajes(valores_peajes_powerbi, valores)
                        
                        # Crear tabla resumen compacta
                        tabla_data_valores = []
                        todos_valores_coinciden = True
                        
                        for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                            comp = comparaciones_peajes[peaje]
                            
                            estado_icono = "✅" if comp['coinciden'] else "❌"
                            diferencia_texto = "$0" if comp['coinciden'] else f"${comp['diferencia']:,.0f}".replace(",", ".")
                            
                            tabla_data_valores.append({
                                '': estado_icono,
                                'Peaje': peaje,
                                'Power BI': comp['powerbi_texto'],
                                'Excel': f"${comp['excel_numero']:,.0f}".replace(",", "."),
                                'Dif.': diferencia_texto
                            })
                            
                            if not comp['coinciden']:
                                todos_valores_coinciden = False
                        
                        df_comparacion_valores = pd.DataFrame(tabla_data_valores)
                        st.dataframe(df_comparacion_valores, use_container_width=True, hide_index=True)
                        
                        st.markdown("---")
                        
                        # ========== SECCIÓN 9: RESUMEN FINAL ==========
                        st.markdown("### 📋 Resultado Final")
                        
                        validacion_completa = (coinciden and coinciden_pasos and 
                                             todos_valores_coinciden and todos_pasos_coinciden)
                        
                        validacion_parcial = (coinciden or coinciden_pasos or 
                                            todos_valores_coinciden or todos_pasos_coinciden)
                        
                        if validacion_completa:
                            st.success("🎉 **VALIDACIÓN EXITOSA** - Todos los valores coinciden")
                            st.balloons()
                        elif validacion_parcial:
                            st.warning("⚠️ **VALIDACIÓN PARCIAL** - Algunos valores coinciden, otros no")
                        else:
                            st.error("❌ **VALIDACIÓN FALLIDA** - Existen diferencias en todos los valores")
                        
                        # Mostrar detalles específicos
                        st.markdown("#### 📈 Detalles de Validación:")
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            if coinciden:
                                st.success("✅ Valores monetarios totales: COINCIDEN")
                            else:
                                st.error("❌ Valores monetarios totales: NO COINCIDEN")
                        
                        with col2:
                            if coinciden_pasos:
                                st.success("✅ Cantidad de pasos totales: COINCIDEN")
                            else:
                                st.error("❌ Cantidad de pasos totales: NO COINCIDEN")
                        
                        with col3:
                            if todos_valores_coinciden:
                                st.success("✅ Valores por peaje: COINCIDEN")
                            else:
                                st.error("❌ Valores por peaje: NO COINCIDEN")
                        
                        with col4:
                            if todos_pasos_coinciden:
                                st.success("✅ Pasos por peaje: COINCIDEN")
                            else:
                                st.error("❌ Pasos por peaje: NO COINCIDEN")
                        
                        # Botón para ver detalles adicionales
                        with st.expander("🔍 Ver Detalles Completos y Capturas"):
                            # Tabla detallada
                            st.markdown("#### 📊 Tabla Detallada")
                            resumen_data = []
                            
                            # Valores monetarios
                            resumen_data.append({
                                'Concepto': 'TOTAL GENERAL',
                                'Power BI': f"${powerbi_numero:,.0f}".replace(",", "."),
                                'Excel': f"${excel_numero:,.0f}".replace(",", "."),
                                'Estado': '✅ Coincide' if coinciden else '❌ No coincide',
                                'Diferencia': f"${abs(powerbi_numero - excel_numero):,.0f}".replace(",", "."),
                                'Dif. %': f"{abs(powerbi_numero - excel_numero)/excel_numero*100:.2f}%" if excel_numero > 0 else "N/A"
                            })
                            
                            # Cantidad de pasos totales
                            resumen_data.append({
                                'Concepto': 'CANTIDAD DE PASOS TOTAL',
                                'Power BI': f"{cantidad_pasos_bi:,}".replace(",", "."),
                                'Excel': f"{total_pasos:,}".replace(",", "."),
                                'Estado': '✅ Coincide' if coinciden_pasos else '❌ No coincide',
                                'Diferencia': f"{abs(cantidad_pasos_bi - total_pasos):,}".replace(",", "."),
                                'Dif. %': f"{abs(cantidad_pasos_bi - total_pasos)/total_pasos*100:.2f}%" if total_pasos > 0 else "N/A"
                            })
                            
                            # Pasos por peaje
                            for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                                comp_pasos = comparaciones_pasos_peajes[peaje]
                                excel_pasos = comp_pasos['pasos_excel']
                                resumen_data.append({
                                    'Concepto': f'PASOS {peaje}',
                                    'Power BI': f"{comp_pasos['pasos_bi']:,}".replace(",", "."),
                                    'Excel': f"{comp_pasos['pasos_excel']:,}".replace(",", "."),
                                    'Estado': '✅ Coincide' if comp_pasos['coinciden'] else '❌ No coincide',
                                    'Diferencia': f"{comp_pasos['diferencia']:,}".replace(",", "."),
                                    'Dif. %': f"{comp_pasos['diferencia']/excel_pasos*100:.2f}%" if excel_pasos > 0 else "N/A"
                                })
                            
                            # Valores por peaje
                            for peaje in ['CHICORAL', 'GUALANDAY', 'COCORA']:
                                comp = comparaciones_peajes[peaje]
                                excel_val = comp['excel_numero']
                                resumen_data.append({
                                    'Concepto': f'VALOR {peaje}',
                                    'Power BI': comp['powerbi_texto'],
                                    'Excel': f"${comp['excel_numero']:,.0f}".replace(",", "."),
                                    'Estado': '✅ Coincide' if comp['coinciden'] else '❌ No coincide',
                                    'Diferencia': f"${comp['diferencia']:,.0f}".replace(",", "."),
                                    'Dif. %': f"{comp['diferencia']/excel_val*100:.2f}%" if excel_val > 0 else "N/A"
                                })
                            
                            df_resumen = pd.DataFrame(resumen_data)
                            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
                            
                            # Screenshots
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
                - Revisa que los totales estén claramente identificados
                - Verifica que existan tanto valores monetarios como cantidad de pasos
                """)
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar la validación")

    # Información de ayuda
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso:**
        1. **Cargar Excel**: Archivo con hojas CHICORAL, GUALANDAY, COCORA
        2. **Extracción automática**: Búsqueda inteligente de valores y pasos en cada hoja
        3. **Seleccionar fecha** de conciliación en Power BI  
        4. **Comparar**: Extrae valores de Power BI y compara con Excel
        
        **Características NUEVAS (v2.3):**
        - ✅ **Comparación Total**: Valida el "VALOR A PAGAR A COMERCIO" total
        - ✅ **Cantidad de Pasos**: Extrae y compara "CANTIDAD PASOS" entre Power BI y Excel
        - ✅ **Pasos por Peaje BI**: Extrae pasos individuales de la tabla "RESUMEN COMERCIOS"
        - ✅ **Pasos por Peaje Excel**: Muestra cantidad de pasos individual por cada peaje
        - ✅ **Comparación Completa**: Valida valores y pasos tanto totales como por peaje
        
        **Estructura esperada del Excel:**
        - Cada hoja debe contener valores monetarios y cantidad de pasos
        - Los datos suelen estar en las últimas filas
        - Formato colombiano para valores monetarios (puntos para miles, coma para decimales)
        
        **Notas:**
        - La extracción busca automáticamente en las últimas filas
        - Los valores deben estar claramente identificados en el Power BI
        - Las fechas deben coincidir exactamente con las del reporte Power BI
        """)

if __name__ == "__main__":
    main()

    # Footer
    st.markdown("---")
    st.markdown('<div class="footer">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v2.3</div>', unsafe_allow_html=True)
