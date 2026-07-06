import os
import sys
import time
import shutil
import asyncio
from concurrent.futures import ThreadPoolExecutor
import threading
import uuid
from pathlib import Path
import traceback
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options


try:
    from webdriver_manager.chrome import ChromeDriverManager
except Exception:
    ChromeDriverManager = None

# Modificado para ser una función async principal
async def crear_descargar(ruta_principal, anio, mes):
    print("PROGRESS::Iniciando módulo de descarga...", flush=True)
    print(f"DEBUG::ruta_principal={ruta_principal}", flush=True)
    print(f"DEBUG::anio={anio}", flush=True)
    print(f"DEBUG::mes={mes}", flush=True)

    fecha_hoy = datetime.now().strftime("%d.%m.%Y")

    carpetas = {
        "cierre": "CIERRE",
        "marco": "MARCO",
        "ejecucion": "EJECUCION",
        "validacion": "VALIDACION GASTO CAPITAL"
    }

    def crear_estructura():
        for nombre in carpetas.values():
            ruta = os.path.join(ruta_principal, nombre)
            os.makedirs(ruta, exist_ok=True)
            print(f"Carpeta creada/verificada: {nombre}")

    def setup_logger(ruta_principal):
        logs_dir = os.path.join(ruta_principal, 'logs')
        os.makedirs(logs_dir, exist_ok=True)
        logger = logging.getLogger('crear_descargar')
        logger.setLevel(logging.DEBUG)
        if not logger.handlers:
            fh = logging.FileHandler(os.path.join(logs_dir, f'crear_descargar_{int(time.time())}.log'), encoding='utf-8')
            fh.setLevel(logging.DEBUG)
            fmt = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
            fh.setFormatter(fmt)
            logger.addHandler(fh)
        return logger

    logger = setup_logger(ruta_principal)
    crear_estructura()
    WAIT_TIME = int(os.environ.get('FONAFE_WAIT', '35')) # Subido ligeramente para dar margen a la otra laptop

    def limpiar_carpeta(ruta):
        try:
            for f in os.listdir(ruta):
                if f.endswith(('.xlsx', '.xls')):
                    try: os.remove(os.path.join(ruta, f))
                    except Exception: pass
        except Exception: pass

    # Optimizamos la función interna para aceptar las opciones custom de cada hilo
    def iniciar_instancia_driver(options):
        try:
            print("PROGRESS::Preparando navegador Chrome...", flush=True)

            options.add_argument("--disable-gpu")
            options.add_argument("--no-first-run")
            options.add_argument("--no-default-browser-check")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-background-networking")

            if ChromeDriverManager:
                print("PROGRESS::Obteniendo ChromeDriver...", flush=True)

                driver_path = ChromeDriverManager().install()

                print(f"DEBUG::ChromeDriver path={driver_path}", flush=True)

                service = Service(driver_path)
                driver = webdriver.Chrome(service=service, options=options)
            else:
                print("PROGRESS::Iniciando Chrome sin webdriver-manager...", flush=True)
                driver = webdriver.Chrome(options=options)

            print("PROGRESS::Navegador Chrome iniciado correctamente.", flush=True)
            return driver

        except WebDriverException as e:
            print(f"ERROR::Error al crear el WebDriver de Chrome: {e}", flush=True)
            print("ERROR::Verifica que Google Chrome esté instalado en esta computadora.", flush=True)
            traceback.print_exc()
            logger.exception("Error creando WebDriver")
            raise

        except Exception as e:
            print(f"ERROR::Error inesperado iniciando ChromeDriver: {e}", flush=True)
            traceback.print_exc()
            logger.exception("Error inesperado creando WebDriver")
            raise

    # Subimos el timeout por defecto a 120 segundos por si la otra laptop es más lenta procesando
    def esperar_descarga(carpeta, timeout=120):
        inicio = time.time()
        while True:
            archivos = [
                os.path.join(carpeta, f)
                for f in os.listdir(carpeta)
                if f.endswith(('.xlsx', '.xls'))
            ]
            if archivos and not any(f.endswith('.crdownload') for f in os.listdir(carpeta)):
                return max(archivos, key=os.path.getctime)
            if time.time() - inicio > timeout:
                raise TimeoutError("Descarga no detectada en el tiempo límite")
            time.sleep(0.5) # Un respiro más amplio para el procesador

    def mover_archivo(archivo, destino, nombre):
        os.makedirs(destino, exist_ok=True)
        ruta_final = os.path.join(destino, nombre)
        if os.path.exists(ruta_final):
            os.remove(ruta_final)
        shutil.move(archivo, ruta_final)
        print(f" ✓ {nombre}")

    def login(driver, wait):
        driver.get("https://app.fonafe.gob.pe/empresaAdmin/")
        wait.until(EC.frame_to_be_available_and_switch_to_it("main"))
        wait.until(EC.element_to_be_clickable((By.ID, "usuarioApp"))).send_keys(os.environ.get("FONAFE_USER", "fona_admin"))
        driver.find_element(By.NAME, "clave").send_keys(os.environ.get("FONAFE_PASS", "20172017....."))
        driver.find_element(By.ID, "loginButton").click()
        print("Login exitoso")

    # ================= GRUPO EJECUCION =================
    def grupo_ejecucion(carpeta_descarga, task_name):
        # Perfil separado de la carpeta de descarga para evitar conflictos de lockfile
        profile_dir = os.path.join(base_temp, 'profiles', 'profile_' + task_name)
        if os.path.exists(profile_dir):
            shutil.rmtree(profile_dir, ignore_errors=True)
        os.makedirs(profile_dir, exist_ok=True)
        os.makedirs(carpeta_descarga, exist_ok=True)
        
        options = Options()
        options.add_experimental_option("prefs", {
            "download.default_directory": carpeta_descarga,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True
        })
        options.add_argument(f"--user-data-dir={profile_dir}")
        
        # CORREGIDO: Ahora sí se le pasan las opciones personalizadas de este grupo
        driver = iniciar_instancia_driver(options)
        wait = WebDriverWait(driver, WAIT_TIME)
        ejec_path = os.path.join(ruta_principal, carpetas["ejecucion"])

        def flujo(radio, nombre, extra=None):
            limpiar_carpeta(carpeta_descarga)
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("menu"))
            script = "abrir('EjecFinanciera')" if "finan" in radio else "abrir('EjecPresupuestal')"
            driver.execute_script(script)
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("right"))
            wait.until(EC.element_to_be_clickable((By.XPATH, f"//input[@value='{radio}']"))).click()
            driver.find_element(By.ID, "siguienteButton").click()
            driver.switch_to.default_content()
            driver.switch_to.frame("right")
            Select(wait.until(EC.presence_of_element_located((By.ID, "ANNO")))).select_by_value(anio)
            Select(driver.find_element(By.ID, "MENSUAL")).select_by_value(mes)
            Select(driver.find_element(By.ID, "ENTIDAD")).select_by_value("064")
            Select(driver.find_element(By.ID, "TRIMESTRE")).select_by_value("13")
            if extra:
                for k, v in extra.items():
                    Select(driver.find_element(By.ID, k)).select_by_value(v)
            wait.until(EC.element_to_be_clickable((By.ID, "exportarButton"))).click()
            try: WebDriverWait(driver, 3).until(EC.alert_is_present()).accept()
            except Exception: pass
            archivo = esperar_descarga(carpeta_descarga)
            mover_archivo(archivo, ejec_path, nombre)

        try:
            login(driver, wait)
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("menu"))
            driver.execute_script("abrir('MODULOS_FONAFE')")
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("right"))
            driver.execute_script("openSelectedMenu('EJECUCION')")

            print("PROGRESS::Ejecución: descargando reportes financieros...", flush=True)
            flujo("form_finan_ejec_balan_gen",   "Balance_General_Ejecucion.xlsx")
            flujo("form_finan_ejec_est_gan_per", "Estado_Ganancias_Perdidas_Ejecucion.xlsx")

            print("PROGRESS::Ejecución: descargando reportes presupuestarios...", flush=True)
            flujo("form_pres_ejec_pres_ing_egr", "Presu_Ingresos_Egresos_Ejecucion.xlsx")
            flujo("form_pres_ejec_flujo_caja",   "Flujo_de_Caja_Ejecucion.xlsx")
            flujo("form_pres_ejec_gas_capital",  "Gastos_Capital_Ejecucion_Presupuesto.xlsx", {"TIPOGASTOCAPITAL": "1"})
            flujo("form_pres_ejec_gas_capital",  "Gastos_Capital_Ejecucion_Flujo_Caja.xlsx",  {"TIPOGASTOCAPITAL": "2"})

            print("PROGRESS::Ejecución: descargando depósitos y colocaciones...", flush=True)
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("menu"))
            driver.execute_script("inicio()")
            driver.execute_script("abrir('MODULOS_FONAFE')")
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("right"))
            driver.execute_script("openSelectedMenu('DEPOSITOS_Y_COLOCACIONES')")
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("menu"))
            driver.execute_script("abrir('depositos_colocaciones')")
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("right"))
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='form_depos_y_coloca']"))).click()
            driver.find_element(By.ID, "siguienteButton").click()
            Select(wait.until(EC.presence_of_element_located((By.ID, "ANNO")))).select_by_value(anio)
            Select(driver.find_element(By.ID, "ENTIDAD")).select_by_value("064")
            Select(driver.find_element(By.ID, "TRIMESTRE")).select_by_value(mes)
            wait.until(EC.element_to_be_clickable((By.ID, "exportarButton"))).click()
            try: WebDriverWait(driver, 3).until(EC.alert_is_present()).accept()
            except Exception: pass
            archivo = esperar_descarga(carpeta_descarga)
            mover_archivo(archivo, ejec_path, "Depositos_Colocaciones_Ejecucion.xlsx")
            print("PROGRESS::Grupo Ejecución completado.", flush=True)
        except Exception:
            logs_dir = os.path.join(ruta_principal, 'logs')
            ts = int(time.time())
            try: driver.save_screenshot(os.path.join(logs_dir, f'{task_name}_ejec_error_{ts}.png'))
            except: logger.exception('No se pudo guardar screenshot')
            raise
        finally:
            driver.quit()

    # ================= GRUPO CIERRE =================
    def grupo_cierre(carpeta_descarga, task_name):
        # Perfil separado de la carpeta de descarga para evitar conflictos de lockfile
        profile_dir = os.path.join(base_temp, 'profiles', 'profile_' + task_name)
        if os.path.exists(profile_dir):
            shutil.rmtree(profile_dir, ignore_errors=True)
        os.makedirs(profile_dir, exist_ok=True)
        os.makedirs(carpeta_descarga, exist_ok=True)
        options = Options()
        options.add_experimental_option("prefs", {
            "download.default_directory": carpeta_descarga,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True
        })
        options.add_argument(f"--user-data-dir={profile_dir}")
        
        # CORREGIDO: Se pasan las opciones correspondientes
        driver = iniciar_instancia_driver(options)
        wait = WebDriverWait(driver, WAIT_TIME)
        cierre_path = os.path.join(ruta_principal, carpetas["cierre"])

        def descargar(modulo, mes_param, nombre):
            limpiar_carpeta(carpeta_descarga)
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("menu"))
            driver.execute_script("abrir('CONSULTAS_GERENCIALES')")
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("right"))
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='form_consulta_gerencial_3']"))).click()
            driver.find_element(By.ID, "siguienteButton").click()
            driver.switch_to.default_content()
            driver.switch_to.frame("right")
            Select(wait.until(EC.presence_of_element_located((By.ID, "ANNO")))).select_by_value(anio)
            Select(driver.find_element(By.ID, "MODULO")).select_by_value(modulo)
            Select(driver.find_element(By.ID, "mes")).select_by_value(mes_param)
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("right"))
            wait.until(EC.element_to_be_clickable((By.ID, "BuscarButton"))).click()
            wait.until(EC.element_to_be_clickable((By.ID, "imprimirButton"))).click()
            iframe_popup = wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[contains(@name,'ventanaPopupMain')]")))
            driver.switch_to.frame(iframe_popup)
            Select(wait.until(EC.visibility_of_element_located((By.ID, "REPORTE")))).select_by_value("4foepradm0004")
            driver.find_element(By.ID, "imprimirButtonXLS").click()
            try: WebDriverWait(driver, 3).until(EC.alert_is_present()).accept()
            except Exception: pass
            archivo = esperar_descarga(carpeta_descarga)
            mover_archivo(archivo, cierre_path, nombre)

        try:
            login(driver, wait)
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("menu"))
            driver.execute_script("abrir('ADMINISTRADOR')")
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("right"))
            driver.execute_script("openSelectedMenu('CONSULTAS_ADMIN')")

            print("PROGRESS::Cierre: descargando reportes de cierre...", flush=True)
            descargar("000", mes,  "Estado_de_Cierre_del_Periodo_Ejecucion.xlsx")
            descargar("002", "16", "Estado_de_Cierre_del_Periodo_Formulacion.xlsx")
            print("PROGRESS::Grupo Cierre completado.", flush=True)
        except Exception:
            logs_dir = os.path.join(ruta_principal, 'logs')
            ts = int(time.time())
            try: driver.save_screenshot(os.path.join(logs_dir, f'{task_name}_cierre_error_{ts}.png'))
            except: logger.exception('No se pudo guardar screenshot')
            raise
        finally:
            driver.quit()

    # ================= GRUPO FORMULACION =================
    def grupo_formulacion(carpeta_descarga, task_name):
        # Perfil separado de la carpeta de descarga para evitar conflictos de lockfile
        profile_dir = os.path.join(base_temp, 'profiles', 'profile_' + task_name)
        if os.path.exists(profile_dir):
            shutil.rmtree(profile_dir, ignore_errors=True)
        os.makedirs(profile_dir, exist_ok=True)
        os.makedirs(carpeta_descarga, exist_ok=True)
        options = Options()
        options.add_experimental_option("prefs", {
            "download.default_directory": carpeta_descarga,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True
        })
        options.add_argument(f"--user-data-dir={profile_dir}")
        
        # CORREGIDO: Se pasan las opciones correspondientes
        driver = iniciar_instancia_driver(options)
        wait = WebDriverWait(driver, WAIT_TIME)
        form_path = os.path.join(ruta_principal, carpetas["marco"])

        def flujo(radio, nombre, extra=None):
            limpiar_carpeta(carpeta_descarga)
            driver.switch_to.default_content()
            driver.switch_to.frame("menu")
            script = "abrir('FormFinancieros')" if "finan" in radio else "abrir('FormPresupuestarios')"
            driver.execute_script(script)
            driver.switch_to.default_content()
            driver.switch_to.frame("right")
            wait.until(EC.element_to_be_clickable((By.XPATH, f"//input[@value='{radio}']"))).click()
            driver.find_element(By.ID, "siguienteButton").click()
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("right"))
            Select(wait.until(EC.element_to_be_clickable((By.ID, "ANNO")))).select_by_value(anio)
            wait.until(EC.presence_of_element_located((By.ID, "TRIMESTRE")))
            Select(driver.find_element(By.ID, "TRIMESTRE")).select_by_value("13")
            if extra:
                for k, v in extra.items():
                    Select(driver.find_element(By.ID, k)).select_by_value(v)
            wait.until(EC.element_to_be_clickable((By.ID, "exportarButton"))).click()
            try: WebDriverWait(driver, 3).until(EC.alert_is_present()).accept()
            except Exception: pass
            archivo = esperar_descarga(carpeta_descarga)
            mover_archivo(archivo, form_path, nombre)

        try:
            login(driver, wait)
            driver.switch_to.default_content()
            driver.switch_to.frame("menu")
            driver.execute_script("abrir('MODULOS_FONAFE')")
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("right"))
            driver.execute_script("openSelectedMenu('FORMULACION')")
            time.sleep(0.5)

            print("PROGRESS::Formulación: descargando reportes financieros...", flush=True)
            flujo("form_finan_form_balan_gen",   "Estado_de_Situacion_Financiera_Formulacion.xlsx")
            flujo("form_finan_form_est_gan_per", "Estado_de_Resultados_Integrales_Formulacion.xlsx")

            print("PROGRESS::Formulación: descargando reportes presupuestarios...", flush=True)
            flujo("form_pres_form_pres_ing_egr", "Presu_Ingresos_Egresos_Formulacion.xlsx")
            flujo("form_pres_form_flujo_caja",   "Flujo_de_Caja_Formulacion.xlsx")
            flujo("form_pres_form_gas_capital",  "Gastos_Capital_Formulacion_Presupuesto.xlsx", {"TIPOGASTOCAPITAL": "1"})
            flujo("form_pres_form_gas_capital",  "Gastos_Capital_Formulacion_Flujo_Caja.xlsx",  {"TIPOGASTOCAPITAL": "2"})
            print("PROGRESS::Grupo Formulación completado.", flush=True)
        except Exception:
            logs_dir = os.path.join(ruta_principal, 'logs')
            ts = int(time.time())
            try: driver.save_screenshot(os.path.join(logs_dir, f'{task_name}_form_error_{ts}.png'))
            except: logger.exception('No se pudo guardar screenshot')
            raise
        finally:
            driver.quit()


    # ================= MAIN ASYNC EXECUTION =================
    print("="*60)
    print("INICIANDO DESCARGA DE REPORTES FONAFE (ASYNC)")
    print(f"Computadora: {os.environ.get('COMPUTERNAME', 'desconocido')}")
    print(f"Usuario: {os.environ.get('USERNAME', 'desconocido')}")
    print("="*60)

    base_temp = os.path.join(os.path.expanduser("~"), "Downloads", "_temp_fonafe")
    
    # Limpiar carpetas temporales de ejecuciones anteriores
    if os.path.exists(base_temp):
        try:
            shutil.rmtree(base_temp, ignore_errors=True)
            logger.info('Carpetas temporales anteriores limpiadas')
        except Exception:
            logger.warning('No se pudieron limpiar todas las carpetas temporales')
    os.makedirs(base_temp, exist_ok=True)
    
    failures = {}
    failures_lock = threading.Lock()  # Protege acceso concurrente al dict
    MAX_CONCURRENCY = int(os.environ.get('FONAFE_MAX_CONCURRENCY', '2'))
    
    loop = asyncio.get_running_loop()
    
    def run_with_retries(func, args, name, retries=1, backoff=5):
        tb = ""
        attempt = 0
        while attempt <= retries:
            try:
                logger.info(f'Iniciando {name} (intento {attempt+1})')
                func(*args)
                logger.info(f'{name} completado correctamente')
                return
            except Exception:
                tb = traceback.format_exc()
                logger.error(f'Error en {name} (intento {attempt+1}):\n{tb}')
                attempt += 1
                if attempt <= retries:
                    wait_time = backoff * attempt  # Backoff exponencial
                    logger.info(f'Esperando {wait_time}s antes de reintentar {name}')
                    time.sleep(wait_time)
        with failures_lock:
            failures[name] = tb

    thread_defs = [
        (grupo_ejecucion,   (os.path.join(base_temp, "g1"), "Grupo-Ejecucion"), "Grupo-Ejecución"),
        (grupo_cierre,      (os.path.join(base_temp, "g2"), "Grupo-Cierre"), "Grupo-Cierre"),
        (grupo_formulacion, (os.path.join(base_temp, "g3"), "Grupo-Formulacion"), "Grupo-Formulación"),
    ]

    print(f"\nIniciando {len(thread_defs)} tareas en paralelo (max {MAX_CONCURRENCY} simultáneas)...")
    
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENCY) as executor:
        tasks = []
        for func, args, name in thread_defs:
            task = loop.run_in_executor(
                executor, 
                run_with_retries, 
                func, 
                args, 
                name, 
                int(os.environ.get('FONAFE_RETRIES', '1'))
            )
            tasks.append(task)
        
        await asyncio.gather(*tasks)

    # Re-ejecución secuencial si fallan
    if failures:
        logger.warning('Se detectaron fallos. Intentando re-ejecución secuencial...')
        print('\n⚠ Algunos grupos fallaron. Reintentando de forma secuencial...')
        for func, args, name in thread_defs:
            if name in failures:
                try:
                    logger.info(f'Re-ejecutando secuencialmente {name}')
                    func(*args)
                    logger.info(f'{name} re-ejecutado correctamente (secuencial)')
                    with failures_lock:
                        del failures[name]
                except Exception:
                    logger.error(f'Fallo al re-ejecutar {name} secuencialmente')

    # Limpieza de carpetas temporales
    try:
        shutil.rmtree(base_temp, ignore_errors=True)
        logger.info('Carpetas temporales limpiadas al finalizar')
    except Exception:
        logger.warning('No se pudieron limpiar las carpetas temporales al finalizar')

    if failures:
        print("ERROR::Algunos grupos de descarga fallaron.", flush=True)
        print(f"ERROR::Revisa los logs en: {os.path.join(ruta_principal, 'logs')}", flush=True)

        for name, tb in failures.items():
            print(f"ERROR::{name} falló.", flush=True)
            print(tb, flush=True)

        return False

    else:
        print("PROGRESS::Proceso de descarga completado exitosamente.", flush=True)
        print("DEBUG::Crear_Descargar_0 finalizó con True", flush=True)
        return True