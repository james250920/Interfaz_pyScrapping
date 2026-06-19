import os
import time
import shutil
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

# webdriver-manager will download a matching chromedriver automatically
try:
    from webdriver_manager.chrome import ChromeDriverManager
except Exception:
    ChromeDriverManager = None

def crear_descargar(ruta_principal, anio, mes):
    
    ruta_principal = ruta_principal
    anio = anio
    mes = mes

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

    # Configuración de logging por proyecto
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
    # Tiempo de espera configurable (segundos) para WebDriverWait
    WAIT_TIME = int(os.environ.get('FONAFE_WAIT', '25'))


    def limpiar_carpeta(ruta):
        try:
            for f in os.listdir(ruta):
                if f.endswith(('.xlsx', '.xls')):
                    try:
                        os.remove(os.path.join(ruta, f))
                    except:
                        pass
        except:
            pass

    def crear_driver(carpeta_descarga):
        os.makedirs(carpeta_descarga, exist_ok=True)
        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {
            "download.default_directory": carpeta_descarga,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True
        })
        try:
            # agregar algunos flags útiles
            options.add_argument("--disable-gpu")
            options.add_argument("--no-first-run")
            options.add_argument("--no-default-browser-check")
            # el profile y carpeta de descargas se asignan fuera y provistos en carpeta_descarga
            if ChromeDriverManager:
                service = Service(ChromeDriverManager().install())
                return webdriver.Chrome(service=service, options=options)
            else:
                # Fallback: rely on chromedriver in PATH
                return webdriver.Chrome(options=options)
        except WebDriverException as e:
            print("✗ Error al crear el WebDriver de Chrome:", e)
            print("Verifica que Chrome esté instalado y que chromedriver sea compatible o esté en el PATH.")
            print("Si quieres que lo gestione automáticamente, instala la dependencia: webdriver-manager")
            traceback.print_exc()
            logger.exception('Error creando WebDriver')
            raise

    def esperar_descarga(carpeta, timeout=60):
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
            time.sleep(0.2)

    def mover_archivo(archivo, destino, nombre):
        os.makedirs(destino, exist_ok=True)
        ruta_final = os.path.join(destino, nombre)
        if os.path.exists(ruta_final):
            os.remove(ruta_final)
        shutil.move(archivo, ruta_final)
        print(f" ✓ {nombre}")

    # ================= LOGIN =================

    def login(driver, wait):
        driver.get("https://app.fonafe.gob.pe/empresaAdmin/")
        wait.until(EC.frame_to_be_available_and_switch_to_it("main"))
        wait.until(EC.element_to_be_clickable((By.ID, "usuarioApp"))).send_keys("fona_admin")
        driver.find_element(By.NAME, "clave").send_keys("20172017.....")
        driver.find_element(By.ID, "loginButton").click()
        print("Login exitoso")

    # ================= GRUPO EJECUCION =================

    def grupo_ejecucion(carpeta_descarga):
        thread_name = threading.current_thread().name
        # crear perfil y carpeta de descarga por hilo
        profile_dir = os.path.join(carpeta_descarga, 'profile_' + thread_name)
        os.makedirs(profile_dir, exist_ok=True)
        # pasar profile y download via opciones recreando driver
        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {
            "download.default_directory": carpeta_descarga,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True
        })
        options.add_argument(f"--user-data-dir={profile_dir}")
        driver = crear_driver(carpeta_descarga)
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
            Select(driver.find_element(By.ID, "ENTIDAD")).select_by_value("064")  # fijo
            Select(driver.find_element(By.ID, "TRIMESTRE")).select_by_value("13")  # fijo

            if extra:
                for k, v in extra.items():
                    Select(driver.find_element(By.ID, k)).select_by_value(v)

            wait.until(EC.element_to_be_clickable((By.ID, "exportarButton"))).click()
            try:
                WebDriverWait(driver, 3).until(EC.alert_is_present()).accept()
            except:
                pass

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

            print("Descargando reportes financieros...")
            flujo("form_finan_ejec_balan_gen",   "Balance_General_Ejecucion.xlsx")
            flujo("form_finan_ejec_est_gan_per", "Estado_Ganancias_Perdidas_Ejecucion.xlsx")

            print("Descargando reportes presupuestarios...")
            flujo("form_pres_ejec_pres_ing_egr", "Presu_Ingresos_Egresos_Ejecucion.xlsx")
            flujo("form_pres_ejec_flujo_caja",   "Flujo_de_Caja_Ejecucion.xlsx")
            flujo("form_pres_ejec_gas_capital",  "Gastos_Capital_Ejecucion_Presupuesto.xlsx", {"TIPOGASTOCAPITAL": "1"})
            flujo("form_pres_ejec_gas_capital",  "Gastos_Capital_Ejecucion_Flujo_Caja.xlsx",  {"TIPOGASTOCAPITAL": "2"})

            # DEPÓSITOS Y COLOCACIONES
            print("Descargando depósitos y colocaciones...")
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
            Select(driver.find_element(By.ID, "ENTIDAD")).select_by_value("064")  # fijo
            Select(driver.find_element(By.ID, "TRIMESTRE")).select_by_value(mes)  # ID=TRIMESTRE pero valor=mes

            wait.until(EC.element_to_be_clickable((By.ID, "exportarButton"))).click()
            try:
                WebDriverWait(driver, 3).until(EC.alert_is_present()).accept()
            except:
                pass

            archivo = esperar_descarga(carpeta_descarga)
            mover_archivo(archivo, ejec_path, "Depositos_Colocaciones_Ejecucion.xlsx")

            print("✓ Grupo Ejecución completado")

        except Exception:
            # guardar estado para diagnóstico
            logs_dir = os.path.join(ruta_principal, 'logs')
            os.makedirs(logs_dir, exist_ok=True)
            ts = int(time.time())
            try:
                shot = os.path.join(logs_dir, f'{thread_name}_ejec_error_{ts}.png')
                driver.save_screenshot(shot)
            except Exception:
                logger.exception('No se pudo guardar screenshot en grupo_ejecucion')
            try:
                src = os.path.join(logs_dir, f'{thread_name}_ejec_page_{ts}.html')
                with open(src, 'w', encoding='utf-8') as f:
                    f.write(driver.page_source)
            except Exception:
                logger.exception('No se pudo guardar page_source en grupo_ejecucion')
            raise
        finally:
            driver.quit()

    # ================= GRUPO CIERRE =================

    def grupo_cierre(carpeta_descarga):
        thread_name = threading.current_thread().name
        profile_dir = os.path.join(carpeta_descarga, 'profile_' + thread_name)
        os.makedirs(profile_dir, exist_ok=True)
        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {
            "download.default_directory": carpeta_descarga,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True
        })
        options.add_argument(f"--user-data-dir={profile_dir}")
        driver = crear_driver(carpeta_descarga)
        wait = WebDriverWait(driver, WAIT_TIME)
        cierre_path = os.path.join(ruta_principal, carpetas["cierre"])

        def descargar(modulo, mes, nombre):
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
            Select(driver.find_element(By.ID, "mes")).select_by_value(mes)
            #Select(driver.find_element(By.ID, "ENTIDAD")).select_by_value("064")  # fijo
            #Select(driver.find_element(By.ID, "ESTADO")).select_by_value("000")   # fijo

            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("right"))
            wait.until(EC.element_to_be_clickable((By.ID, "BuscarButton"))).click()
            wait.until(EC.element_to_be_clickable((By.ID, "imprimirButton"))).click()

            #wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "ventanaPopupMain_97")))

            iframe_popup = wait.until(
        EC.presence_of_element_located(
        (By.XPATH, "//iframe[contains(@name,'ventanaPopupMain')]")
    )
)

            driver.switch_to.frame(iframe_popup)

            Select(wait.until(EC.visibility_of_element_located((By.ID, "REPORTE")))).select_by_value("4foepradm0004")

            driver.find_element(By.ID, "imprimirButtonXLS").click()
            try:
                WebDriverWait(driver, 3).until(EC.alert_is_present()).accept()
            except:
                pass

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

            print("Descargando reportes de cierre...")
            descargar("000", mes,  "Estado_de_Cierre_del_Periodo_Ejecucion.xlsx")    # mes = variable
            descargar("002", "16", "Estado_de_Cierre_del_Periodo_Formulacion.xlsx")  # "16" fijo

            print("✓ Grupo Cierre completado")

        except Exception:
            logs_dir = os.path.join(ruta_principal, 'logs')
            os.makedirs(logs_dir, exist_ok=True)
            ts = int(time.time())
            try:
                shot = os.path.join(logs_dir, f'{thread_name}_cierre_error_{ts}.png')
                driver.save_screenshot(shot)
            except Exception:
                logger.exception('No se pudo guardar screenshot en grupo_cierre')
            try:
                src = os.path.join(logs_dir, f'{thread_name}_cierre_page_{ts}.html')
                with open(src, 'w', encoding='utf-8') as f:
                    f.write(driver.page_source)
            except Exception:
                logger.exception('No se pudo guardar page_source en grupo_cierre')
            raise
        finally:
            driver.quit()

    # ================= GRUPO FORMULACION =================

    def grupo_formulacion(carpeta_descarga):
        thread_name = threading.current_thread().name
        profile_dir = os.path.join(carpeta_descarga, 'profile_' + thread_name)
        os.makedirs(profile_dir, exist_ok=True)
        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {
            "download.default_directory": carpeta_descarga,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True
        })
        options.add_argument(f"--user-data-dir={profile_dir}")
        driver = crear_driver(carpeta_descarga)
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
            Select(driver.find_element(By.ID, "TRIMESTRE")).select_by_value("13")  # fijo

            if extra:
                for k, v in extra.items():
                    Select(driver.find_element(By.ID, k)).select_by_value(v)

            wait.until(EC.element_to_be_clickable((By.ID, "exportarButton"))).click()
            try:
                WebDriverWait(driver, 3).until(EC.alert_is_present()).accept()
            except:
                pass

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

            print("Descargando reportes financieros...")
            flujo("form_finan_form_balan_gen",   "Estado_de_Situacion_Financiera_Formulacion.xlsx")
            flujo("form_finan_form_est_gan_per", "Estado_de_Resultados_Integrales_Formulacion.xlsx")

            print("Descargando reportes presupuestarios...")
            flujo("form_pres_form_pres_ing_egr", "Presu_Ingresos_Egresos_Formulacion.xlsx")
            flujo("form_pres_form_flujo_caja",   "Flujo_de_Caja_Formulacion.xlsx")
            flujo("form_pres_form_gas_capital",  "Gastos_Capital_Formulacion_Presupuesto.xlsx", {"TIPOGASTOCAPITAL": "1"})
            flujo("form_pres_form_gas_capital",  "Gastos_Capital_Formulacion_Flujo_Caja.xlsx",  {"TIPOGASTOCAPITAL": "2"})

            print("✓ Grupo Formulación completado")

        except Exception:
            logs_dir = os.path.join(ruta_principal, 'logs')
            os.makedirs(logs_dir, exist_ok=True)
            ts = int(time.time())
            try:
                shot = os.path.join(logs_dir, f'{thread_name}_form_error_{ts}.png')
                driver.save_screenshot(shot)
            except Exception:
                logger.exception('No se pudo guardar screenshot en grupo_formulacion')
            try:
                src = os.path.join(logs_dir, f'{thread_name}_form_page_{ts}.html')
                with open(src, 'w', encoding='utf-8') as f:
                    f.write(driver.page_source)
            except Exception:
                logger.exception('No se pudo guardar page_source en grupo_formulacion')
            raise
        finally:
            driver.quit()


    # ================= MAIN =================


    print("="*60)
    print("INICIANDO DESCARGA DE REPORTES FONAFE")
    print("="*60)

    crear_estructura()

    base_temp = os.path.join(os.path.expanduser("~"), "Downloads", "_temp_fonafe")

    # Ejecutar grupos en hilos con reintentos y re-ejecución secuencial si fallan
    failures = {}

    # Semáforo para limitar concurrencia simultánea
    MAX_CONCURRENCY = int(os.environ.get('FONAFE_MAX_CONCURRENCY', '3'))
    semaphore = threading.Semaphore(MAX_CONCURRENCY)

    def run_with_retries(func, args, name, retries=1, backoff=2):
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
                    time.sleep(backoff)
        failures[name] = tb

    thread_defs = [
        (grupo_ejecucion,   (os.path.join(base_temp, "g1"),), "Grupo-Ejecución"),
        (grupo_cierre,      (os.path.join(base_temp, "g2"),), "Grupo-Cierre"),
        (grupo_formulacion, (os.path.join(base_temp, "g3"),), "Grupo-Formulación"),
    ]

    threads = []
    print("\nIniciando threads paralelos...")
    for func, args, name in thread_defs:
        def thread_target(f=func, a=args, n=name):
            semaphore.acquire()
            try:
                run_with_retries(f, a, n, retries=int(os.environ.get('FONAFE_RETRIES', '1')))
            finally:
                semaphore.release()

        t = threading.Thread(target=thread_target, name=name)
        threads.append(t)
        t.start()
        time.sleep(1)

    for t in threads:
        t.join()
        time.sleep(1.5)

    # Si hubo fallos, intentar re-ejecutar secuencialmente los grupos fallidos (más tolerante)
    if failures:
        logger.warning('Se detectaron fallos en threads. Intentando re-ejecución secuencial de los grupos fallidos...')
        for func, args, name in thread_defs:
            if name in failures:
                try:
                    logger.info(f'Re-ejecutando secuencialmente {name}')
                    func(*args)
                    logger.info(f'{name} re-ejecutado correctamente (secuencial)')
                    del failures[name]
                except Exception:
                    tb = traceback.format_exc()
                    logger.error(f'Fallo al re-ejecutar {name} secuencialmente:\n{tb}')

    if failures:
        logger.error('Al finalizar hay grupos que continúan fallando: %s', ', '.join(failures.keys()))
        print('\nAlgunos grupos fallaron. Revisa los logs en la carpeta logs para más detalles.')

    print("\n" + "="*60)
    print("✓ PROCESO COMPLETADO EXITOSAMENTE")
    print("="*60)