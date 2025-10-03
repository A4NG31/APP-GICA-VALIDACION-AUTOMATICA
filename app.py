import os
import sys

# ===== CONFIGURACIÓN CRÍTICA PARA STREAMLIT CLOUD =====
# Desactivar completamente el watcher ANTES de importar streamlit
os.environ['STREAMLIT_SERVER_FILE_WATCHER_TYPE'] = 'none'
os.environ['STREAMLIT_CI'] = 'true'
os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'

# Monkey patch para evitar problemas de watcher
import streamlit.web.bootstrap

def patched_install_config_watchers(*args, **kwargs):
    """No-op replacement para evitar límites de inotify en Streamlit Cloud"""
    return

streamlit.web.bootstrap._install_config_watchers = patched_install_config_watchers

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

# Configuración de la página
st.set_page_config(page_title="Validador Power BI GICA", layout="wide")

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

def click_conciliacion_date(driver, fecha_objetivo):
    """Hacer clic en la conciliación específica por fecha - OPTIMIZADO"""
    try:
        # Esperar a que cargue la página principal
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        
        time.sleep(5)  # Espera adicional para Power BI
        
        # Buscar el elemento que contiene la fecha exacta
        selectors = [
            f"//*[contains(text(), 'Conciliación APP GICA del {fecha_objetivo}')]",
            f"//*[contains(text(), 'CONCILIACIÓN APP GICA DEL {fecha_objetivo}')]",
            f"//*[contains(text(), '{fecha_objetivo} 00:00 al {fecha_objetivo} 11:59')]",
            f"//div[contains(text(), '{fecha_objetivo}')]",
            f"//span[contains(text(), '{fecha_objetivo}')]",
            f"//*[contains(@title, '{fecha_objetivo}')]",
            f"//*[contains(@aria-label, '{fecha_objetivo}')]",
            f"//button[contains(text(), '{fecha_objetivo}')]",
            f"//a[contains(text(), '{fecha_objetivo}')]",
        ]
        
        elemento_conciliacion = None
        for selector in selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed() and elemento.is_enabled():
                        elemento_conciliacion = elemento
                        st.success(f"✅ Encontrado elemento con selector: {selector[:50]}...")
                        break
                if elemento_conciliacion:
                    break
            except Exception as e:
                continue
        
        if elemento_conciliacion:
            # Scroll y clic seguro
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", elemento_conciliacion)
            time.sleep(2)
            
            # Intentar diferentes métodos de clic
            try:
                elemento_conciliacion.click()
            except:
                try:
                    driver.execute_script("arguments[0].click();", elemento_conciliacion)
                except:
                    # Último intento con ActionChains
                    from selenium.webdriver.common.action_chains import ActionChains
                    actions = ActionChains(driver)
                    actions.move_to_element(elemento_conciliacion).click().perform()
            
            st.info("🔄 Esperando a que cargue la conciliación...")
            time.sleep(8)  # Tiempo generoso para carga
            return True
        else:
            st.error("❌ No se encontró la conciliación para la fecha especificada")
            
            # Debug: mostrar elementos visibles
            try:
                elementos_visibles = driver.find_elements(By.XPATH, "//*[text()]")
                textos_importantes = []
                for elem in elementos_visibles[:30]:
                    if elem.is_displayed():
                        texto = elem.text.strip()
                        if texto and len(texto) > 5 and len(texto) < 100:
                            textos_importantes.append(texto)
                
                if textos_importantes:
                    st.info("🔍 Elementos visibles encontrados (primeros 10):")
                    for texto in textos_importantes[:10]:
                        st.write(f"- {texto}")
            except:
                pass
                
            return False
            
    except Exception as e:
        st.error(f"❌ Error al hacer clic en conciliación: {str(e)}")
        return False

def find_valor_a_pagar_comercio_card(driver):
    """Buscar la tarjeta 'VALOR A PAGAR A COMERCIO' - MEJORADO"""
    try:
        # Esperar a que cargue el contenido después del clic
        time.sleep(5)
        
        # Buscar por diferentes patrones del título
        titulo_selectors = [
            "//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]",
            "//*[contains(text(), 'Valor a pagar a comercio')]",
            "//*[contains(text(), 'VALOR A PAGAR') and contains(text(), 'COMERCIO')]",
            "//*[contains(text(), 'Valor A Pagar') and contains(text(), 'Comercio')]",
            "//*[contains(text(), 'PAGAR A COMERCIO')]",
            "//*[contains(text(), 'Valor a Pagar a Comercio')]",
            "//*[@data-title='VALOR A PAGAR A COMERCIO']",
            "//*[contains(@class, 'card') and contains(text(), 'PAGAR')]",
            "//*[contains(@class, 'valuecard') and contains(text(), 'COMERCIO')]",
            "//*[contains(@class, 'visual') and contains(text(), 'PAGAR')]",
            "//div[contains(@class, 'title') and contains(text(), 'COMERCIO')]",
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
                            st.success(f"✅ Encontrado título: {texto}")
                            break
                if titulo_element:
                    break
            except:
                continue
        
        if not titulo_element:
            st.error("❌ No se encontró 'VALOR A PAGAR A COMERCIO' en el reporte")
            return None
        
        # Estrategias múltiples para encontrar el valor
        valor_encontrado = None
        
        # Estrategia 1: Buscar en el mismo contenedor
        try:
            container = titulo_element.find_element(By.XPATH, "./ancestor::*[contains(@class, 'card') or contains(@class, 'visual') or contains(@class, 'valueCard')][1]")
            all_elements = container.find_elements(By.XPATH, ".//*")
            
            for elem in all_elements:
                texto = elem.text.strip()
                if (texto and any(char.isdigit() for char in texto) and 
                    texto != titulo_element.text and
                    len(texto) > 3 and len(texto) < 50):
                    # Verificar si es un valor monetario
                    if '$' in texto or (',' in texto and '.' in texto):
                        valor_encontrado = texto
                        st.success(f"✅ Valor encontrado (estrategia 1): {valor_encontrado}")
                        break
        except:
            pass
        
        # Estrategia 2: Buscar siguiente elemento con número grande
        if not valor_encontrado:
            try:
                siguiente_hermano = titulo_element.find_element(By.XPATH, "./following-sibling::*[1]")
                texto = siguiente_hermano.text.strip()
                if texto and any(char.isdigit() for char in texto) and len(texto) > 3:
                    valor_encontrado = texto
                    st.success(f"✅ Valor encontrado (estrategia 2): {valor_encontrado}")
            except:
                pass
        
        # Estrategia 3: Buscar en elementos cercanos
        if not valor_encontrado:
            try:
                elementos_cercanos = driver.find_elements(By.XPATH, 
                    f"//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]/following::*[position()<=10]")
                
                for elem in elementos_cercanos:
                    texto = elem.text.strip()
                    if texto and any(char.isdigit() for char in texto) and len(texto) < 50:
                        # Filtrar valores que parecen ser monetarios
                        if ('$' in texto or 
                            (',' in texto and '.' in texto and 
                             len(texto.split('.')[-1]) in [2, 3])):
                            valor_encontrado = texto
                            st.success(f"✅ Valor encontrado (estrategia 3): {valor_encontrado}")
                            break
            except:
                pass
        
        # Estrategia 4: Buscar cualquier número grande visible
        if not valor_encontrado:
            try:
                all_elements = driver.find_elements(By.XPATH, "//*[text()]")
                candidatos = []
                
                for elem in all_elements:
                    if elem.is_displayed():
                        texto = elem.text.strip()
                        if (texto and any(char.isdigit() for char in texto) and 
                            len(texto) > 4 and len(texto) < 30 and
                            ('$' in texto or ',' in texto)):
                            # Calcular puntaje
                            puntaje = 0
                            if '$' in texto:
                                puntaje += 10
                            if texto.count(',') >= 1:
                                puntaje += 5
                            if texto.count('.') == 1:
                                puntaje += 3
                            if len(texto) > 8:
                                puntaje += 2
                            
                            if puntaje >= 10:
                                candidatos.append((puntaje, texto))
                
                if candidatos:
                    mejor_candidato = max(candidatos, key=lambda x: x[0])
                    valor_encontrado = mejor_candidato[1]
                    st.success(f"✅ Valor encontrado (estrategia 4): {valor_encontrado}")
            except:
                pass
        
        if valor_encontrado:
            return valor_encontrado
        else:
            st.error("❌ No se pudo encontrar el valor numérico asociado")
            return None
        
    except Exception as e:
        st.error(f"❌ Error buscando valor: {str(e)}")
        return None

def extract_excel_values(uploaded_file):
    """Extraer valores de las 3 hojas del Excel - OPTIMIZADO"""
    try:
        st.info("📊 Procesando archivo Excel...")
        
        # Guardar archivo temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name
        
        hojas = ['CHICORAL', 'GUALANDAY', 'COCORA']
        valores = {}
        total_general = 0
        
        for hoja in hojas:
            try:
                df = pd.read_excel(tmp_path, sheet_name=hoja, header=None, engine='openpyxl')
                
                # Buscar el ÚLTIMO "Total" en la hoja
                valor_encontrado = None
                mejor_candidato = None
                mejor_puntaje = -1
                
                # Buscar de ABAJO hacia ARRIBA (últimas 50 filas)
                start_row = max(0, len(df) - 50)
                for i in range(len(df)-1, start_row-1, -1):
                    fila = df.iloc[i]
                    
                    # Buscar "Total" en esta fila
                    for j, celda in enumerate(fila):
                        if pd.notna(celda) and isinstance(celda, str) and 'TOTAL' in celda.upper().strip():
                            
                            # Buscar valores monetarios en la MISMA fila
                            for k in range(len(fila)):
                                posible_valor = fila.iloc[k]
                                if pd.notna(posible_valor):
                                    valor_str = str(posible_valor)
                                    
                                    # Calcular puntaje de confianza
                                    puntaje = 0
                                    if '$' in valor_str:
                                        puntaje += 10
                                    if any(c.isdigit() for c in valor_str):
                                        puntaje += 5
                                    if '.' in valor_str and len(valor_str.split('.')[-1]) in [2, 3]:
                                        puntaje += 3
                                    if len(valor_str) > 6:
                                        puntaje += 2
                                    if ',' in valor_str:
                                        puntaje += 2
                                    
                                    # Excluir valores incorrectos
                                    if puntaje > 0 and len(valor_str) < 4:
                                        puntaje = 0
                                    if 'pag' in valor_str.lower() or 'total' in valor_str.lower():
                                        puntaje = 0
                                    
                                    if puntaje > mejor_puntaje:
                                        mejor_puntaje = puntaje
                                        mejor_candidato = posible_valor
                
                # Usar el mejor candidato
                if mejor_candidato is not None and mejor_puntaje >= 8:
                    valor_encontrado = mejor_candidato
                    st.success(f"✅ {hoja}: {valor_encontrado} (confianza: {mejor_puntaje}/20)")
                else:
                    st.warning(f"⚠️ {hoja}: No se encontró valor claro, usando 0")
                    valor_encontrado = 0
                
                # Procesar el valor encontrado
                if valor_encontrado != 0:
                    valor_original = str(valor_encontrado)
                    valor_limpio = re.sub(r'[^\d.,]', '', valor_original)
                    
                    try:
                        # Para formato colombiano (puntos para miles, coma para decimal)
                        if '.' in valor_limpio and ',' in valor_limpio:
                            # Formato: 1.000,00 → 1000.00
                            valor_limpio = valor_limpio.replace('.', '').replace(',', '.')
                        elif '.' in valor_limpio:
                            # Formato: 1000.00
                            if len(valor_limpio.split('.')[-1]) == 2:  # Probable decimal
                                valor_limpio = valor_limpio.replace('.', '')  # Eliminar punto de miles
                            else:
                                valor_limpio = valor_limpio.replace('.', '')  # Eliminar todos los puntos
                        elif ',' in valor_limpio:
                            # Formato: 1000,00
                            partes = valor_limpio.split(',')
                            if len(partes) == 2 and len(partes[1]) == 2:
                                valor_limpio = partes[0] + '.' + partes[1]  # Coma como decimal
                            else:
                                valor_limpio = valor_limpio.replace(',', '')  # Coma como separador de miles
                        
                        valor_numerico = float(valor_limpio)
                        
                        if valor_numerico >= 1000:
                            valores[hoja] = valor_numerico
                            total_general += valor_numerico
                        else:
                            valores[hoja] = 0
                            
                    except Exception as e:
                        st.error(f"❌ Error procesando valor de {hoja}: {e}")
                        valores[hoja] = 0
                else:
                    valores[hoja] = 0
                    
            except Exception as e:
                st.error(f"❌ Error leyendo hoja {hoja}: {e}")
                valores[hoja] = 0
        
        # Limpiar archivo temporal
        try:
            os.unlink(tmp_path)
        except:
            pass
        
        return valores, total_general
        
    except Exception as e:
        st.error(f"❌ Error procesando archivo Excel: {str(e)}")
        return {}, 0

def compare_values(valor_powerbi_texto, valor_esperado):
    """Comparar valores con tolerancia"""
    try:
        # Limpiar el valor de Power BI
        powerbi_limpio = str(valor_powerbi_texto)
        
        # Eliminar texto no numérico excepto puntos, comas y $
        powerbi_limpio = re.sub(r'[^\d.,$]', '', powerbi_limpio)
        
        # Remover símbolo de dólar
        powerbi_limpio = powerbi_limpio.replace('$', '')
        
        # Determinar formato y convertir
        if '.' in powerbi_limpio and ',' in powerbi_limpio:
            # Formato: 1.000,00 o 1,000.00
            if powerbi_limpio.rfind('.') > powerbi_limpio.rfind(','):
                # Formato: 1,000.00 → eliminar comas
                powerbi_limpio = powerbi_limpio.replace(',', '')
            else:
                # Formato: 1.000,00 → eliminar puntos, coma a punto
                powerbi_limpio = powerbi_limpio.replace('.', '').replace(',', '.')
        elif '.' in powerbi_limpio:
            # Solo puntos - verificar si es decimal o separador de miles
            partes = powerbi_limpio.split('.')
            if len(partes) > 1 and len(partes[-1]) == 2:
                # Probable decimal: 1000.00 → eliminar puntos
                powerbi_limpio = powerbi_limpio.replace('.', '')
            else:
                # Separador de miles: 1.000 → eliminar puntos
                powerbi_limpio = powerbi_limpio.replace('.', '')
        elif ',' in powerbi_limpio:
            # Solo comas - verificar si es decimal o separador de miles
            partes = powerbi_limpio.split(',')
            if len(partes) > 1 and len(partes[-1]) == 2:
                # Probable decimal: 1000,00 → coma a punto
                powerbi_limpio = powerbi_limpio.replace(',', '.')
            else:
                # Separador de miles: 1,000 → eliminar comas
                powerbi_limpio = powerbi_limpio.replace(',', '')
        
        # Convertir a números
        powerbi_numero = float(powerbi_limpio)
        excel_numero = float(valor_esperado)
        
        # Comparar (con tolerancia para decimales y redondeo)
        tolerancia = max(1.0, excel_numero * 0.001)  # 0.1% o mínimo 1.0
        coinciden = abs(powerbi_numero - excel_numero) <= tolerancia
        
        return powerbi_numero, excel_numero, valor_powerbi_texto, coinciden
        
    except Exception as e:
        st.error(f"❌ Error en comparación: {str(e)}")
        return None, None, valor_powerbi_texto, False

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI - OPTIMIZADA PARA CLOUD"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiYTFmOWZkMDAtY2IwYi00OTg4LWIxZDctNGZmYmU0NTMxNGI1IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        st.error("❌ No se pudo inicializar el navegador en Streamlit Cloud")
        return None
    
    try:
        # 1. Navegar al reporte con manejo de timeouts
        with st.spinner("🔄 Conectando con Power BI... Esto puede tomar hasta 60 segundos"):
            try:
                driver.get(REPORT_URL)
                # Espera progresiva para carga de Power BI
                for i in range(1, 7):
                    time.sleep(5)
                    st.spinner(f"🔄 Cargando Power BI... ({i*5}/30 segundos)")
            except Exception as e:
                st.warning(f"⚠️ Timeout en carga inicial: {e}")
        
        # Verificar si la página cargó correctamente
        current_url = driver.current_url
        if "powerbi.com" not in current_url:
            st.warning(f"⚠️ Posible redirección. URL actual: {current_url}")
        
        # 2. Hacer clic en la conciliación específica
        with st.spinner("🔍 Buscando conciliación..."):
            if not click_conciliacion_date(driver, fecha_objetivo):
                return None
        
        # 3. Buscar tarjeta "VALOR A PAGAR A COMERCIO" y extraer valor
        with st.spinner("💰 Extrayendo valor de Power BI..."):
            valor_texto = find_valor_a_pagar_comercio_card(driver)
        
        if valor_texto:
            st.success("✅ Valor extraído correctamente de Power BI")
            return {
                'valor_texto': valor_texto,
                'status': 'success'
            }
        else:
            st.error("❌ No se pudo encontrar el valor en el reporte de Power BI")
            return {
                'valor_texto': None,
                'status': 'value_not_found'
            }
        
    except Exception as e:
        st.error(f"❌ Error durante la extracción: {str(e)}")
        import traceback
        st.error(f"🔍 Detalles del error: {traceback.format_exc()}")
        return None
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass

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
    """)
    
    # Estado del sistema
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success("✅ Configurado para Streamlit Cloud")
    st.sidebar.info("🚀 Selenium optimizado para entorno cloud")
    
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

    # Información de ayuda
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso - Streamlit Cloud"):
        st.markdown("""
        **Proceso Optimizado para Cloud:**
        1. **Cargar Excel**: Archivo con hojas CHICORAL, GUALANDAY, COCORA
        2. **Extracción automática**: Búsqueda inteligente de "Total" en cada hoja
        3. **Seleccionar fecha** de conciliación en Power BI  
        4. **Comparar**: Extrae valor de Power BI y compara con Excel
        
        **Características Cloud:**
        - ✅ Configurado para evitar límites de inotify
        - ✅ Selenium optimizado para entorno serverless
        - ✅ Manejo robusto de timeouts
        - ✅ Tolerancia para diferencias decimales
        
        **Formato Excel Esperado:**
        - Hojas: CHICORAL, GUALANDAY, COCORA
        - Busca automáticamente la palabra "TOTAL"
        - Soporta formatos colombianos (puntos para miles)
        """)

if __name__ == "__main__":
    main()
