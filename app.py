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

# ===== IMPORTS =====
import streamlit as st
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import re

st.set_page_config(page_title="Validador Power BI - APP GICA", page_icon="💰", layout="wide")

# ===== CSS Sidebar =====
st.markdown("""
<style>
[data-testid="stSidebar"] {
    background-color: #1E1E2F !important;
    color: white !important;
    width: 300px !important;
}
[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stCheckbox label {
    color: white !important; 
}
</style>
""", unsafe_allow_html=True)

# ===== Logo =====
st.markdown("""
<div style="display: flex; justify-content: center; margin-bottom: 30px;">
    <img src="https://i.imgur.com/z9xt46F.jpeg"
         style="width: 50%; border-radius: 10px;" 
         alt="Logo Gopass">
</div>
""", unsafe_allow_html=True)

# ===== FUNCIONES =====
def setup_driver():
    """Configura ChromeDriver"""
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

        driver = webdriver.Chrome(options=chrome_options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver
    except Exception as e:
        st.error(f"Error al configurar ChromeDriver: {e}")
        return None

def click_conciliacion_date(driver, fecha_objetivo):
    """Clic en la conciliación por fecha"""
    try:
        selectors = [
            f"//*[contains(text(), 'Conciliación APP GICA del {fecha_objetivo}')]",
            f"//*[contains(text(), 'CONCILIACIÓN APP GICA DEL {fecha_objetivo}')]",
        ]
        for selector in selectors:
            try:
                elemento = driver.find_element(By.XPATH, selector)
                driver.execute_script("arguments[0].click();", elemento)
                time.sleep(2)
                return True
            except:
                continue
        st.error("No se encontró la conciliación para la fecha especificada")
        return False
    except Exception as e:
        st.error(f"Error al hacer clic: {e}")
        return False

def find_valor_a_pagar_comercio_card(driver):
    """Busca la tarjeta de VALOR A PAGAR A COMERCIO"""
    try:
        posibles = driver.find_elements(By.XPATH, "//*[contains(text(), 'PAGAR A COMERCIO')]")
        for e in posibles:
            if e.is_displayed():
                parent = e.find_element(By.XPATH, "./..")
                nums = parent.find_elements(By.XPATH, ".//*[contains(text(), '$') or contains(text(), ',')]")
                for n in nums:
                    t = n.text.strip()
                    if any(c.isdigit() for c in t):
                        return t
        return None
    except Exception as e:
        st.error(f"Error buscando valor: {e}")
        return None

def extract_powerbi_data(fecha_objetivo):
    """Extrae datos de Power BI"""
    url = "https://app.powerbi.com/view?r=eyJrIjoiYTFmOWZkMDAtY2IwYi00OTg4LWIxZDctNGZmYmU0NTMxNGI1IiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    driver = setup_driver()
    if not driver:
        return None
    try:
        driver.get(url)
        time.sleep(8)
        click_conciliacion_date(driver, fecha_objetivo)
        time.sleep(3)
        valor = find_valor_a_pagar_comercio_card(driver)
        driver.quit()
        return {"valor_texto": valor}
    except Exception as e:
        st.error(f"Error extrayendo datos: {e}")
        driver.quit()
        return None

def extract_excel_values(uploaded_file):
    """Extrae valores de Excel"""
    hojas = ['CHICORAL', 'GUALANDAY', 'COCORA']
    valores = {}
    total = 0
    for hoja in hojas:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=hoja, header=None)
            valor = None
            for i in range(len(df)-1, -1, -1):
                fila = df.iloc[i].astype(str)
                if fila.str.contains("TOTAL", case=False).any():
                    nums = fila[fila.str.contains(r"\d", na=False)]
                    if not nums.empty:
                        valor = nums.iloc[-1]
                        break
            if valor:
                limpio = re.sub(r"[^\d.,]", "", valor)
                limpio = limpio.replace('.', '').replace(',', '.')
                valor_num = float(limpio)
                valores[hoja] = valor_num
                total += valor_num
            else:
                valores[hoja] = 0
        except:
            valores[hoja] = 0
    return valores, total

def convert_currency_to_float(texto):
    try:
        if not texto:
            return 0.0
        limpio = str(texto).replace('$','').replace('.','').replace(',','.')
        return float(re.sub(r"[^\d.]", "", limpio))
    except:
        return 0.0

def compare_values(valor_powerbi, valor_excel):
    valor_txt = valor_powerbi.get('valor_texto', '')
    pb_num = convert_currency_to_float(valor_txt)
    exc_num = float(valor_excel)
    return pb_num, exc_num, valor_txt, abs(pb_num - exc_num) < 0.01

# ===== INTERFAZ PRINCIPAL =====
def main():
    st.title("💰 Validador Power BI - Conciliaciones APP GICA")
    st.markdown("---")

    st.sidebar.header("📋 Información")
    uploaded_file = st.file_uploader("Selecciona el archivo Excel", type=['xlsx', 'xls'])
    if uploaded_file:
        valores, total = extract_excel_values(uploaded_file)
        st.success(f"✅ Total general: ${total:,.0f}".replace(",", "."))
        fecha = st.date_input("Fecha de Conciliación")
        fecha_objetivo = fecha.strftime("%Y-%m-%d")

        if st.button("🎯 Extraer y Comparar", type="primary"):
            resultado = extract_powerbi_data(fecha_objetivo)
            if resultado:
                pb, ex, txt, ok = compare_values(resultado, total)
                st.metric("Power BI", txt)
                st.metric("Excel", f"${ex:,.0f}".replace(",", "."))
                if ok:
                    st.success("✅ Coinciden")
                    st.balloons()
                else:
                    st.error(f"❌ No coinciden (Dif: ${abs(pb-ex):,.0f})".replace(",", "."))

if __name__ == "__main__":
    main()

    
    # Footer
    st.markdown("---")
    st.markdown('<div class="footer">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit</div>', unsafe_allow_html=True)
