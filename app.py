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
import os
import sys

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
    """Configurar ChromeDriver para Selenium - VERSIÓN COMPATIBLE LOCAL Y CLOUD"""
    try:
        chrome_options = Options()
        
        # Opciones esenciales para compatibilidad
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # User agent real
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        # Intentar diferentes métodos de inicialización
        try:
            # Método 1: Usar webdriver-manager (ideal para local)
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
        except:
            # Método 2: Usar Chrome directamente (para Streamlit Cloud)
            driver = webdriver.Chrome(options=chrome_options)
        
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver
        
    except Exception as e:
        st.error(f"❌ Error al configurar ChromeDriver: {e}")
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
            f"//*[contains(@title, '{fecha_objetivo}')]",
            f"//*[contains(@aria-label, '{fecha_objetivo}')]",
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
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", elemento_conciliacion)
            time.sleep(2)
            
            # Intentar diferentes métodos de clic
            try:
                elemento_conciliacion.click()
            except:
                driver.execute_script("arguments[0].click();", elemento_conciliacion)
            
            time.sleep(5)  # Más tiempo para que cargue
            return True
        else:
            st.error("❌ No se encontró la conciliación para la fecha especificada")
            st.info("💡 Sugerencia: Verifica que la fecha esté en formato YYYY-MM-DD")
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
            "//*[contains(text(), 'Valor a Pagar a Comercio')]",
            "//*[@data-title='VALOR A PAGAR A COMERCIO']",
            "//*[contains(@class, 'card') and contains(text(), 'PAGAR')]",
            "//*[contains(@class, 'valuecard') and contains(text(), 'COMERCIO')]",
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
            # Mostrar información de depuración
            st.info("🔍 Buscando elementos visibles en la página...")
            try:
                elementos_visibles = driver.find_elements(By.XPATH, "//*[text()]")
                textos_encontrados = []
                for elem in elementos_visibles[:20]:  # Primeros 20 elementos
                    if elem.is_displayed():
                        texto = elem.text.strip()
                        if texto and len(texto) > 0:
                            textos_encontrados.append(texto)
                if textos_encontrados:
                    st.write("Elementos visibles encontrados:", textos_encontrados[:10])
            except:
                pass
            return None
        
        # Buscar el valor numérico - Estrategias múltiples
        valor_encontrado = None
        
        # Estrategia 1: Buscar en el mismo contenedor
        try:
            container = titulo_element.find_element(By.XPATH, "./ancestor::*[contains(@class, 'card') or contains(@class, 'visual')][1]")
            numeric_elements = container.find_elements(By.XPATH, ".//*[text()]")
            
            for elem in numeric_elements:
                texto = elem.text.strip()
                if (texto and any(char.isdigit() for char in texto) and 
                    texto != titulo_element.text and
                    len(texto) > 3):
                    # Verificar si es un valor monetario
                    if '$' in texto or (',' in texto and '.' in texto):
                        valor_encontrado = texto
                        break
        except:
            pass
        
        # Estrategia 2: Buscar elementos hermanos
        if not valor_encontrado:
            try:
                parent = titulo_element.find_element(By.XPATH, "./..")
                siblings = parent.find_elements(By.XPATH, "./*")
                
                for sibling in siblings:
                    if sibling != titulo_element:
                        texto = sibling.text.strip()
                        if texto and any(char.isdigit() for char in texto) and len(texto) > 3:
                            valor_encontrado = texto
                            break
            except:
                pass
        
        # Estrategia 3: Buscar siguiente elemento hermano
        if not valor_encontrado:
            try:
                siguiente_hermano = titulo_element.find_element(By.XPATH, "./following-sibling::*[1]")
                texto = siguiente_hermano.text.strip()
                if texto and any(char.isdigit() for char in texto):
                    valor_encontrado = texto
            except:
                pass
        
        # Estrategia 4: Buscar en elementos cercanos
        if not valor_encontrado:
            try:
                elementos_cercanos = driver.find_elements(By.XPATH, 
                    f"//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]/following::*[position()<=5]")
                
                for elem in elementos_cercanos:
                    texto = elem.text.strip()
                    if texto and any(char.isdigit() for char in texto) and len(texto) < 50:
                        valor_encontrado = texto
                        break
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
    """Extraer valores de las 3 hojas del Excel"""
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
                        else:
                            valores[hoja] = 0
                            
                    except Exception:
                        valores[hoja] = 0
                else:
                    valores[hoja] = 0
                    
            except Exception:
                valores[hoja] = 0
        
        return valores, total_general
        
    except Exception as e:
        st.error(f"❌ Error procesando archivo Excel: {str(e)}")
        return {}, 0

def compare_values(valor_powerbi_texto, valor_esperado):
    """Comparar valores"""
    try:
        # Limpiar el valor de Power BI
        powerbi_limpio = str(valor_powerbi_texto)
        powerbi_limpio = powerbi_limpio.replace('.', '')  # Eliminar puntos de miles
        powerbi_limpio = re.sub(r'[^\d,]', '', powerbi_limpio)  # Mantener solo números y comas
        powerbi_limpio = powerbi_limpio.replace(',', '.')  # Convertir coma a punto decimal
        
        # Convertir a números
        powerbi_numero = float(powerbi_limpio)
        excel_numero = float(valor_esperado)
        
        # Comparar (con tolerancia para decimales)
        tolerancia = 0.01
        coinciden = abs(powerbi_numero - excel_numero) <= tolerancia
        
        return powerbi_numero, excel_numero, valor_powerbi_texto, coinciden
        
    except Exception as e:
        st.error(f"❌ Error en comparación: {str(e)}")
        return None, None, valor_powerbi_texto, False

def extract_powerbi_data(fecha_objetivo):
    """Función principal para extraer datos de Power BI - VERSIÓN MEJORADA"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiYTFmOWZkMDAtY2IwYi00OTg4LWIxZDctNGZmYmU0NTMxNGI1IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        st.error("❌ No se pudo inicializar el navegador")
        return None
    
    try:
        # 1. Navegar al reporte con más tiempo de espera
        with st.spinner("🔄 Conectando con Power BI... Esto puede tomar hasta 30 segundos"):
            driver.get(REPORT_URL)
            time.sleep(20)  # Tiempo generoso para carga inicial de Power BI
        
        # Verificar si la página cargó correctamente
        current_url = driver.current_url
        if "powerbi.com" not in current_url:
            st.warning(f"⚠️ Posible redirección. URL actual: {current_url}")
        
        # 2. Hacer clic en la conciliación específica
        with st.spinner("🔍 Buscando conciliación..."):
            if not click_conciliacion_date(driver, fecha_objetivo):
                return None
        
        # 3. Esperar a que cargue la selección
        time.sleep(10)
        
        # 4. Buscar tarjeta "VALOR A PAGAR A COMERCIO" y extraer valor
        with st.spinner("💰 Extrayendo valor de Power BI..."):
            valor_texto = find_valor_a_pagar_comercio_card(driver)
        
        if valor_texto:
            st.success("✅ Valor extraído correctamente")
            return {
                'valor_texto': valor_texto,
                'status': 'success'
            }
        else:
            st.error("❌ No se pudo encontrar el valor en el reporte")
            return {
                'valor_texto': None,
                'status': 'value_not_found'
            }
        
    except Exception as e:
        st.error(f"❌ Error durante la extracción: {str(e)}")
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
    """)
    
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
            if st.button("🚀 Extraer Valor de Power BI y Comparar", type="primary"):
                with st.spinner("Extrayendo datos de Power BI... Esto puede tomar 1-2 minutos"):
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
                                    st.metric("Estado", "✅ COINCIDEN", delta=None)
                                else:
                                    diferencia = abs(powerbi_numero - excel_numero)
                                    diferencia_formateada = f"${diferencia:,.0f}".replace(",", ".")
                                    st.metric("Estado", "❌ NO COINCIDEN", delta=diferencia_formateada)
                            
                            # Mostrar detalles
                            with st.expander("📊 Detalles de la Comparación"):
                                st.write(f"**Power BI (numérico):** {powerbi_numero:,.0f}".replace(",", "."))
                                st.write(f"**Excel (numérico):** {excel_numero:,.0f}".replace(",", "."))
                                st.write(f"**Diferencia:** {abs(powerbi_numero - excel_numero):,.0f}".replace(",", "."))
                                
                    elif resultados:
                        st.error("❌ Se accedió al reporte pero no se encontró el valor")
                    else:
                        st.error("❌ No se pudieron extraer datos del reporte")
        else:
            st.error("❌ No se pudieron extraer valores del archivo Excel")
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar")

    # Información de ayuda
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso:**
        1. **Cargar Excel**: Archivo con hojas CHICORAL, GUALANDAY, COCORA
        2. **Extracción automática**: Busca "Total" en cada hoja y calcula suma
        3. **Seleccionar fecha** de conciliación en Power BI  
        4. **Comparar**: Extrae valor de Power BI y compara con Excel
        
        **Características:**
        - Búsqueda inteligente de valores monetarios
        - Manejo de formato colombiano (puntos para miles)
        - Comparación automática
        - Tolerancia para pequeñas diferencias decimales
        """)

if __name__ == "__main__":
    main()
