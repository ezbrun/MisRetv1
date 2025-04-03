import os
import time
import random
import json
import pandas as pd
import shutil
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

def setup_driver(download_path):
    """Configurar el driver de Chrome con las opciones necesarias"""
    chrome_options = Options()
    
    # Configurar la carpeta de descargas (versión optimizada)
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True,
        "download.open_pdf_in_system_reader": False,
        "download.directory_upgrade": True,
        # Asegurarse de que los archivos Excel se descarguen correctamente
        "browser.helperApps.neverAsk.saveToDisk": "application/vnd.ms-excel;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/csv;text/csv"
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Eliminar el mensaje "Un software automatizado está controlando Chrome"
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    
    # Agregar argumento para eliminar el banner de automatización
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    # Inicializar el driver
    driver = webdriver.Chrome(options=chrome_options)
    
    # Ejecutar JavaScript para eliminar el banner de automatización
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined
        })
        """
    })
    
    return driver

def read_credentials(excel_path):
    """Leer credenciales desde el archivo Excel"""
    try:
        df = pd.read_excel(excel_path)
        # Verificar que existan las columnas necesarias
        if 'CUIT' not in df.columns or 'Clave' not in df.columns:
            print(f"Error: El archivo Excel debe contener las columnas 'CUIT' y 'Clave'")
            return []
        
        # Convertir CUIT a string para asegurar formato correcto
        df['CUIT'] = df['CUIT'].astype(str)
        
        return df[['CUIT', 'Clave']].values.tolist()
    except Exception as e:
        print(f"Error al leer el archivo Excel: {str(e)}")
        return []

def type_like_human(element, text):
    """Simular escritura humana tecla por tecla con pausas aleatorias"""
    for char in text:
        # Pausa aleatoria entre 0.1 y 0.3 segundos entre cada tecla
        time.sleep(random.uniform(0.1, 0.3))
        element.send_keys(char)
        # Pausa adicional aleatoria ocasional para simular pensamiento
        if random.random() < 0.2:  # 20% de probabilidad
            time.sleep(random.uniform(0.2, 0.5))

def login_afip(driver, cuit, clave, wait):
    """Realizar el login en AFIP simulando comportamiento humano"""
    try:
        # Ingresar CUIT
        cuit_input = wait.until(EC.element_to_be_clickable((By.ID, "F1:username")))
        cuit_input.clear()
        # Escribir CUIT tecla por tecla
        type_like_human(cuit_input, cuit)
        
        # Pequeña pausa antes de hacer clic en el botón
        time.sleep(random.uniform(0.5, 1.0))
        
        # Click en botón siguiente
        next_button = wait.until(EC.element_to_be_clickable((By.ID, "F1:btnSiguiente")))
        next_button.click()
        
        # Esperar a que aparezca el campo de contraseña
        time.sleep(random.uniform(1.5, 2.5))
        
        # Ingresar Clave
        clave_input = wait.until(EC.element_to_be_clickable((By.ID, "F1:password")))
        clave_input.clear()
        # Escribir clave tecla por tecla
        type_like_human(clave_input, clave)
        
        # Pequeña pausa antes de hacer clic en el botón de login
        time.sleep(random.uniform(0.7, 1.2))
        
        # Click en botón de login
        login_button = wait.until(EC.element_to_be_clickable((By.ID, "F1:btnIngresar")))
        login_button.click()
        
        # Esperar a que cargue la página después del login
        time.sleep(random.uniform(4.0, 6.0))
        return True
    except Exception as e:
        print(f"Error en el login: {str(e)}")
        return False

def check_authentication_error(driver):
    """Verificar si hay un error de autenticación en la página"""
    try:
        # Verificar si hay un mensaje de error de autenticación
        error_text = driver.find_element(By.TAG_NAME, "body").text
        if "HTTP Status 401" in error_text and "AUTHENTICATION_ALREADY_PRESENT" in error_text:
            print("Detectado error de autenticación: HTTP Status 401 - AUTHENTICATION_ALREADY_PRESENT")
            return True
        return False
    except:
        return False

def navigate_to_mis_retenciones(driver, wait, cuit, max_attempts=3):
    """Navegar a Mis Retenciones usando el buscador con reintentos"""
    for attempt in range(1, max_attempts + 1):
        try:
            print(f"Navegando a Mis Retenciones para CUIT: {cuit} (Intento {attempt}/{max_attempts})")
            
            # Esperar a que la página principal cargue completamente
            time.sleep(random.uniform(3.0, 5.0))
            
            # Buscar el campo de búsqueda
            print("Buscando el campo de búsqueda...")
            search_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='buscadorInput']")))
            
            # Mover el mouse al elemento antes de hacer clic
            actions = ActionChains(driver)
            actions.move_to_element(search_input).pause(random.uniform(0.5, 1.0)).perform()
            search_input.click()
            time.sleep(random.uniform(1.0, 2.0))
            
            # Limpiar el campo de búsqueda
            search_input.clear()
            
            # Escribir "MIS RETENCIONES" tecla por tecla
            print("Escribiendo 'MIS RETENCIONES' en el buscador...")
            type_like_human(search_input, "MIS RETENCIONES")
            time.sleep(random.uniform(2.0, 3.0))
            
            # Esperar a que aparezcan los resultados de búsqueda
            print("Esperando resultados de búsqueda...")
            
            # Buscar y hacer clic en el resultado de "MIS RETENCIONES"
            try:
                # Intentar encontrar el resultado específico
                result_item = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//*[@id='rbt-menu-item-0']/a/div/div/div[1]/div/p")))
                
                # Verificar que el texto del resultado sea "Mis Retenciones"
                if "Mis Retenciones" in result_item.text:
                    print(f"Resultado encontrado: {result_item.text}")
                    
                    # Mover el mouse al elemento antes de hacer clic
                    actions = ActionChains(driver)
                    actions.move_to_element(result_item).pause(random.uniform(0.5, 1.0)).perform()
                    result_item.click()
                    time.sleep(random.uniform(4.0, 6.0))
                    
                    # Cambiar a la nueva pestaña que se abre
                    print("Cambiando a la nueva pestaña...")
                    if len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        time.sleep(random.uniform(2.0, 3.0))
                        
                        # Verificar si hay error de autenticación
                        if check_authentication_error(driver):
                            print("Cerrando pestaña con error y reintentando...")
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                            time.sleep(random.uniform(2.0, 3.0))
                            continue  # Reintentar
                        
                        return True
                    else:
                        print("No se abrió una nueva pestaña para Mis Retenciones")
                        return False
                else:
                    print(f"El resultado no coincide con 'Mis Retenciones': {result_item.text}")
                    return False
                    
            except TimeoutException:
                print("No se encontró el resultado específico. Buscando alternativas...")
                
                # Intentar encontrar cualquier resultado que contenga "Mis Retenciones"
                try:
                    # Buscar todos los resultados
                    results = driver.find_elements(By.XPATH, "//p[contains(text(), 'Mis Retenciones')]")
                    
                    if results:
                        # Hacer clic en el primer resultado que contenga "Mis Retenciones"
                        for result in results:
                            if "Mis Retenciones" in result.text and "Reclamos" not in result.text:
                                print(f"Resultado alternativo encontrado: {result.text}")
                                
                                # Mover el mouse al elemento antes de hacer clic
                                actions = ActionChains(driver)
                                actions.move_to_element(result).pause(random.uniform(0.5, 1.0)).perform()
                                result.click()
                                time.sleep(random.uniform(4.0, 6.0))
                                
                                # Cambiar a la nueva pestaña que se abre
                                print("Cambiando a la nueva pestaña...")
                                if len(driver.window_handles) > 1:
                                    driver.switch_to.window(driver.window_handles[-1])
                                    time.sleep(random.uniform(2.0, 3.0))
                                    
                                    # Verificar si hay error de autenticación
                                    if check_authentication_error(driver):
                                        print("Cerrando pestaña con error y reintentando...")
                                        driver.close()
                                        driver.switch_to.window(driver.window_handles[0])
                                        time.sleep(random.uniform(2.0, 3.0))
                                        continue  # Reintentar
                                    
                                    return True
                                else:
                                    print("No se abrió una nueva pestaña para Mis Retenciones")
                                    return False
                        
                        print("No se encontró ningún resultado que coincida exactamente con 'Mis Retenciones'")
                        return False
                    else:
                        print("No se encontraron resultados para 'Mis Retenciones'")
                        return False
                        
                except Exception as e:
                    print(f"Error al buscar resultados alternativos: {str(e)}")
                    return False
            
        except Exception as e:
            print(f"Error al navegar a Mis Retenciones (Intento {attempt}): {str(e)}")
            
            # Si hay pestañas abiertas, cerrarlas y volver a la principal
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                time.sleep(random.uniform(2.0, 3.0))
    
    print(f"No se pudo navegar a Mis Retenciones después de {max_attempts} intentos")
    return False

def get_previous_month_dates():
    """Obtener el rango de fechas del mes anterior en formato ddmmaaaa"""
    today = datetime.now()
    first_day_current_month = datetime(today.year, today.month, 1)
    last_day_previous_month = first_day_current_month - timedelta(days=1)
    first_day_previous_month = datetime(last_day_previous_month.year, last_day_previous_month.month, 1)
    
    # Formatear fechas como ddmmaaaa (sin barras, año con 4 dígitos)
    start_date = first_day_previous_month.strftime("%d%m%Y")
    end_date = last_day_previous_month.strftime("%d%m%Y")
    
    return start_date, end_date

def consultar_retenciones(driver, wait, cuit, codigo_retencion, download_path):
    """Consultar retenciones para un código específico"""
    try:
        print(f"Consultando retenciones para CUIT: {cuit}, Código: {codigo_retencion}")
        
        # 1. Seleccionar CUIT del retenido
        print("Seleccionando CUIT del retenido...")
        cuit_retenido_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='cuitRetenido']")))
        
        # Mover el mouse al elemento antes de hacer clic
        actions = ActionChains(driver)
        actions.move_to_element(cuit_retenido_select).pause(random.uniform(0.5, 1.0)).perform()
        cuit_retenido_select.click()
        time.sleep(random.uniform(1.0, 2.0))
        
        # Seleccionar el CUIT del dropdown (debe coincidir con el CUIT de la base de datos)
        select = Select(cuit_retenido_select)
        
        # Buscar la opción que contiene el CUIT
        found = False
        for option in select.options:
            if cuit in option.text:
                select.select_by_visible_text(option.text)
                found = True
                break
        
        if not found:
            print(f"No se encontró el CUIT {cuit} en las opciones disponibles")
            return False
        
        time.sleep(random.uniform(1.0, 2.0))
        
        # 2. Seleccionar impuesto retenido
        print(f"Seleccionando impuesto retenido: {codigo_retencion}...")
        impuesto_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='impuestos']")))
        
        # Mover el mouse al elemento antes de hacer clic
        actions = ActionChains(driver)
        actions.move_to_element(impuesto_select).pause(random.uniform(0.5, 1.0)).perform()
        impuesto_select.click()
        time.sleep(random.uniform(1.0, 2.0))
        
        # Seleccionar el código de impuesto
        select = Select(impuesto_select)
        
        # Buscar la opción que contiene el código de impuesto
        found = False
        for option in select.options:
            if f"{codigo_retencion} -" in option.text:
                select.select_by_visible_text(option.text)
                found = True
                break
        
        if not found:
            print(f"No se encontró el código de impuesto {codigo_retencion} en las opciones disponibles")
            return False
        
        time.sleep(random.uniform(1.0, 2.0))
        
        # 3. Completar fechas (mes anterior) en formato ddmmaaaa
        start_date, end_date = get_previous_month_dates()
        
        # Fecha desde
        print(f"Completando fecha desde: {start_date}...")
        fecha_desde = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td/form/table[2]/tbody/tr[8]/td[2]/input[1]")))
        fecha_desde.clear()
        type_like_human(fecha_desde, start_date)
        time.sleep(random.uniform(1.0, 2.0))
        
        # Fecha hasta
        print(f"Completando fecha hasta: {end_date}...")
        fecha_hasta = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td/form/table[2]/tbody/tr[8]/td[2]/input[2]")))
        fecha_hasta.clear()
        type_like_human(fecha_hasta, end_date)
        time.sleep(random.uniform(1.0, 2.0))
        
        # 4. Hacer clic en consultar
        print("Haciendo clic en 'Consultar'...")
        consultar_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td/form/table[2]/tbody/tr[13]/td/input")))
        
        # Mover el mouse al elemento antes de hacer clic
        actions = ActionChains(driver)
        actions.move_to_element(consultar_button).pause(random.uniform(0.5, 1.0)).perform()
        consultar_button.click()
        time.sleep(random.uniform(4.0, 6.0))
        
        # 5. Verificar si hay un mensaje de "No se han encontrado datos"
        try:
            # Intentar encontrar el botón "VOLVER" que aparece cuando no hay resultados
            volver_button = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td/table[2]/tbody/tr[2]/td/input")))
            
            print(f"No se encontraron retenciones para el código {codigo_retencion}. Haciendo clic en 'VOLVER'...")
            
            # Mover el mouse al botón VOLVER antes de hacer clic
            actions = ActionChains(driver)
            actions.move_to_element(volver_button).pause(random.uniform(0.5, 1.0)).perform()
            volver_button.click()
            time.sleep(random.uniform(2.0, 3.0))
            
            return False
        except TimeoutException:
            # No se encontró el botón VOLVER, lo que significa que probablemente hay resultados
            pass
        
        # 6. Verificar si hay resultados (botón de exportar a Excel)
        try:
            # Intentar encontrar el botón de exportar a Excel
            exportar_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td/table[3]/tbody/tr/td[2]/table/tbody/tr/td[8]/a")))
            
            # Si llegamos aquí, hay resultados
            print("Resultados encontrados. Exportando a Excel...")
            
            # Obtener la lista de archivos antes de la descarga
            files_before = set(os.listdir(download_path))
            
            # Mover el mouse al elemento antes de hacer clic
            actions = ActionChains(driver)
            actions.move_to_element(exportar_button).pause(random.uniform(0.5, 1.0)).perform()
            exportar_button.click()
            
            # Esperar a que se descargue el archivo (tiempo más largo)
            print("Esperando a que se complete la descarga...")
            time.sleep(random.uniform(8.0, 12.0))  # Aumentar el tiempo de espera
            
            # Verificar que se haya descargado el archivo
            files_after = set(os.listdir(download_path))
            new_files = list(files_after - files_before)
            
            if new_files:
                # Renombrar el archivo descargado
                for file in new_files:
                    file_path = os.path.join(download_path, file)
                    
                    # Verificar que el archivo sea un Excel válido antes de renombrarlo
                    if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.csv'):
                        # Esperar a que el archivo esté completamente descargado
                        time.sleep(2.0)
                        
                        # Comprobar que el archivo no esté bloqueado y sea accesible
                        try:
                            # Intentar abrir el archivo para verificar que esté completo
                            with open(file_path, 'rb') as f:
                                # Solo leer algunos bytes para verificar acceso
                                f.read(10)
                                
                            # Crear el nuevo nombre de archivo
                            new_filename = f"{cuit}_MisRetenciones_{codigo_retencion}{os.path.splitext(file)[1]}"
                            new_path = os.path.join(download_path, new_filename)
                            
                            # Si ya existe un archivo con ese nombre, eliminarlo
                            if os.path.exists(new_path):
                                os.remove(new_path)
                            
                            # Copiar el archivo en lugar de renombrarlo
                            shutil.copy2(file_path, new_path)
                            print(f"Archivo copiado como: {new_filename}")
                            
                            # Eliminar el archivo original después de copiarlo
                            try:
                                os.remove(file_path)
                                print(f"Archivo original eliminado: {file}")
                            except Exception as e:
                                print(f"No se pudo eliminar el archivo original: {str(e)}")
                            
                            # Usar el botón "Atrás" del navegador para volver a la página anterior
                            print("Volviendo a la página anterior...")
                            driver.back()
                            time.sleep(random.uniform(3.0, 5.0))
                            
                            return True
                        except Exception as e:
                            print(f"Error al procesar el archivo descargado: {str(e)}")
                            # Continuar con el siguiente archivo si hay más de uno
                    
                print("No se pudo procesar correctamente el archivo Excel descargado")
                
                # Usar el botón "Atrás" del navegador para volver a la página anterior
                print("Volviendo a la página anterior...")
                driver.back()
                time.sleep(random.uniform(3.0, 5.0))
                
                return False
            else:
                print("No se detectó ningún archivo descargado")
                
                # Usar el botón "Atrás" del navegador para volver a la página anterior
                print("Volviendo a la página anterior...")
                driver.back()
                time.sleep(random.uniform(3.0, 5.0))
                
                return False
            
        except TimeoutException:
            # No hay resultados o no se encontró el botón de exportar
            print(f"No se encontraron resultados para el código {codigo_retencion} o no se pudo encontrar el botón de exportar")
            
            # Intentar hacer clic en el botón VOLVER si está presente
            try:
                volver_button = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
                    (By.XPATH, "/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td/table[2]/tbody/tr[2]/td/input")))
                
                print("Haciendo clic en 'VOLVER'...")
                actions = ActionChains(driver)
                actions.move_to_element(volver_button).pause(random.uniform(0.5, 1.0)).perform()
                volver_button.click()
                time.sleep(random.uniform(2.0, 3.0))
            except:
                # Si no hay botón VOLVER, usar el botón "Atrás" del navegador
                print("Volviendo a la página anterior...")
                driver.back()
                time.sleep(random.uniform(3.0, 5.0))
            
            return False
        
    except Exception as e:
        print(f"Error al consultar retenciones: {str(e)}")
        
        # Intentar volver a la página anterior en caso de error
        try:
            driver.back()
            time.sleep(random.uniform(3.0, 5.0))
        except:
            pass
            
        return False

def close_mis_retenciones_tab(driver):
    """Cerrar la pestaña de Mis Retenciones y volver a la pestaña principal"""
    try:
        # Cerrar la pestaña actual (Mis Retenciones)
        print("Cerrando la pestaña de Mis Retenciones...")
        driver.close()
        
        # Cambiar a la pestaña original (ARCA)
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(random.uniform(2.0, 3.0))
        
        return True
    except Exception as e:
        print(f"Error al cerrar la pestaña de Mis Retenciones: {str(e)}")
        return False

def logout_afip(driver, wait):
    """Cerrar sesión en AFIP/ARCA"""
    try:
        # Cerrar sesión en ARCA
        print("Cerrando sesión en ARCA...")
        
        # Hacer clic en el icono de usuario
        user_icon = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='userIconoChico']")))
        
        # Mover el mouse al icono de usuario antes de hacer clic
        actions = ActionChains(driver)
        actions.move_to_element(user_icon).pause(random.uniform(0.5, 1.0)).perform()
        user_icon.click()
        time.sleep(random.uniform(1.0, 2.0))
        
        # Hacer clic en el botón de cerrar sesión
        logout_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='contBtnContribuyente']/div[6]/button/div/div[2]")))
        
        # Mover el mouse al botón de cerrar sesión antes de hacer clic
        actions = ActionChains(driver)
        actions.move_to_element(logout_button).pause(random.uniform(0.5, 1.0)).perform()
        logout_button.click()
        time.sleep(random.uniform(2.0, 3.0))
        
        print("Sesión cerrada correctamente")
        return True
    except Exception as e:
        print(f"Error al cerrar sesión: {str(e)}")
        return False

def main():
    """Función principal"""
    # Configuración de rutas
    excel_path = r"C:\Users\eze\Downloads\CREDENCIALES.xlsx"
    download_path = r"C:\Users\eze\Downloads"
    
    # Códigos de retención a consultar
    codigos_retencion = ["216", "767"]
    
    # Verificar que el archivo Excel existe
    if not os.path.exists(excel_path):
        print(f"Error: No se encontró el archivo Excel en {excel_path}")
        return
    
    # Leer credenciales
    credentials = read_credentials(excel_path)
    if not credentials:
        print("No se pudieron obtener credenciales válidas. Verifique el archivo Excel.")
        return
    
    print(f"Se encontraron {len(credentials)} registros para procesar.")
    
    # Inicializar el driver una sola vez para todos los CUIT
    driver = None
    
    try:
        # Configurar el driver
        driver = setup_driver(download_path)
        
        # Configurar espera explícita
        wait = WebDriverWait(driver, 20)
        
        # Navegar a la página de AFIP
        driver.get("https://auth.afip.gob.ar/contribuyente_/login.xhtml")
        driver.maximize_window()
        
        # Procesar cada CUIT
        for i, (cuit, clave) in enumerate(credentials):
            print(f"\nProcesando CUIT: {cuit} ({i+1}/{len(credentials)})")
            
            try:
                # Si no es el primer CUIT y ya estamos logueados, cerrar sesión primero
                if i > 0:
                    # Cerrar sesión
                    if not logout_afip(driver, wait):
                        print(f"No se pudo cerrar la sesión anterior. Refrescando la página...")
                        driver.get("https://auth.afip.gob.ar/contribuyente_/login.xhtml")
                        time.sleep(random.uniform(3.0, 5.0))
                
                # Login en AFIP
                if not login_afip(driver, cuit, clave, wait):
                    print(f"No se pudo completar el login para el CUIT {cuit}. Continuando con el siguiente.")
                    continue
                
                # Navegar a Mis Retenciones con manejo de errores de autenticación
                if not navigate_to_mis_retenciones(driver, wait, cuit, max_attempts=3):
                    print(f"No se pudo navegar a Mis Retenciones para el CUIT {cuit}. Continuando con el siguiente.")
                    continue
                
                # Consultar cada código de retención
                for codigo in codigos_retencion:
                    if not consultar_retenciones(driver, wait, cuit, codigo, download_path):
                        print(f"No se pudieron consultar las retenciones para el código {codigo}. Continuando con el siguiente código.")
                    else:
                        print(f"Retenciones para el código {codigo} consultadas exitosamente.")
                
                # Cerrar la pestaña de Mis Retenciones y volver a la pestaña principal
                if not close_mis_retenciones_tab(driver):
                    print(f"No se pudo cerrar la pestaña de Mis Retenciones. Continuando con el siguiente CUIT.")
                    # Si hay un problema, intentar cerrar todas las pestañas excepto la primera
                    while len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                        time.sleep(1)
                    driver.switch_to.window(driver.window_handles[0])
                
            except Exception as e:
                print(f"Error procesando CUIT {cuit}: {str(e)}")
                
                # Intentar recuperarse para el siguiente CUIT
                try:
                    # Cerrar todas las pestañas excepto la primera
                    while len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                        time.sleep(1)
                    driver.switch_to.window(driver.window_handles[0])
                    
                    # Volver a la página de inicio de AFIP
                    driver.get("https://auth.afip.gob.ar/contribuyente_/login.xhtml")
                    time.sleep(random.uniform(3.0, 5.0))
                except Exception as e:
                    print(f"Error al intentar recuperarse: {str(e)}")
    except Exception as e:
        print(f"Error general: {str(e)}")
    finally:
        # Cerrar el navegador al finalizar todos los CUIT
        if driver:
            driver.quit()
    
    print("\nProceso completado.")

if __name__ == "__main__":
    main()