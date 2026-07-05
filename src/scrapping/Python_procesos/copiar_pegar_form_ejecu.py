import os
import time
import gc
import shutil
import tempfile
import logging
from pathlib import Path

import pythoncom
import win32com.client as win32
import pywintypes
import warnings


warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

XL_UP = -4162  # Constante nativa xlUp

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846

logger = logging.getLogger(__name__)


def com_call(fn, reintentos=12, pausa=1.0):
    for intento in range(1, reintentos + 1):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER):
                if intento == reintentos:
                    raise
                time.sleep(pausa)
            else:
                raise


def copiar_pegar_form_ejecu(ruta_archivo):
    """
    Copia datos desde archivos de MARCO y EJECUCION hacia
    'Validación Data Marco y Ejecución al mes.xlsm'.

    Ejecución estrictamente secuencial:
    - Sin async.
    - Sin ThreadPoolExecutor.
    - Sin procesamiento paralelo.
    - Una sola instancia COM de Excel.
    """

    ruta_base = Path(ruta_archivo).expanduser().resolve()
    destino_file = ruta_base / "Validación Data Marco y Ejecución al mes.xlsm"

    CONFIG = [
        {
            "nombre": "MARCO",
            "carpeta": ruta_base / "MARCO",
            "hoja_origen": "ULTIMA_FORMULACION",
            "columnas": "AB",
            "excluir": ("gastos_capital",),
            "mapa": {
                "Presu_Ingresos_Egresos_Formulacion": "Sistema-Form. PRE",
                "Flujo_de_Caja_Formulacion": "Sistema-Form. FC",
                "Estado_de_Situacion_Financiera_Formulacion": "Sistema-Form. ESF",
                "Estado_de_Resultados_Integrales_Formulacion": "Sistema-Form. ERI",
            },
        },
        {
            "nombre": "EJECUCION",
            "carpeta": ruta_base / "EJECUCION",
            "hoja_origen": 1,
            "columnas": "U",
            "excluir": ("gastos_capital", "depositos_colocaciones"),
            "mapa": {
                "Presu_Ingresos_Egresos_Ejecucion": "Sistema-Ejec. PRE",
                "Flujo_de_Caja_Ejecucion": "Sistema-Ejec. FC",
                "Balance_General_Ejecucion": "Sistema-Ejec. ESF",
                "Estado_Ganancias_Perdidas_Ejecucion": "Sistema-Ejec. ERI",
            },
        },
    ]

    inicio = time.time()

    com_inicializado = False
    proceso_ok = False
    conservar_temporal = False

    excel = None
    wb_destino = None
    wb_origen = None

    ws_origen = None
    ws_destino = None

    cache_destino = {}
    destino_temporal = None

    def cerrar_workbook_seguro(wb, nombre, save_changes=False):
        if wb is None:
            return

        try:
            com_call(lambda: wb.Close(SaveChanges=save_changes))
            logger.info("Workbook cerrado: %s", nombre)
        except Exception as e:
            logger.warning("No se pudo cerrar workbook %s: %s", nombre, e)

    def obtener_nombres_hojas(wb):
        nombres = []
        sheets = None
        hoja = None

        try:
            sheets = com_call(lambda: wb.Worksheets)
            total = com_call(lambda: sheets.Count)

            for i in range(1, total + 1):
                hoja = None
                try:
                    hoja = com_call(lambda i=i: sheets(i))
                    nombres.append(com_call(lambda hoja=hoja: hoja.Name))
                finally:
                    hoja = None

            return nombres

        finally:
            hoja = None
            sheets = None
            gc.collect()

    def obtener_ultima_hoja_formulacion(wb):
        nombres_hojas = obtener_nombres_hojas(wb)

        hojas_form = []

        for nombre in nombres_hojas:
            if not nombre.startswith("Formulacion_"):
                continue

            try:
                numero = int(nombre.split("_", 1)[1])
            except (IndexError, ValueError):
                continue

            hojas_form.append((numero, nombre))

        if hojas_form:
            hoja_final = sorted(hojas_form, key=lambda item: item[0])[-1][1]
            return com_call(lambda: wb.Worksheets(hoja_final))

        return com_call(lambda: wb.Worksheets(1))

    def configurar_excel_no_interactivo(app):
        com_call(lambda: setattr(app, "Visible", False))
        com_call(lambda: setattr(app, "DisplayAlerts", False))
        com_call(lambda: setattr(app, "ScreenUpdating", False))
        com_call(lambda: setattr(app, "EnableEvents", False))
        com_call(lambda: setattr(app, "DisplayStatusBar", False))
        com_call(lambda: setattr(app, "AskToUpdateLinks", False))

        try:
            com_call(lambda: setattr(app, "AutomationSecurity", 3))
        except Exception as e:
            logger.warning("No se pudo establecer AutomationSecurity=3: %s", e)

    try:
        if not ruta_base.exists():
            logger.error("No existe la ruta base: %s", ruta_base)
            return False

        if not ruta_base.is_dir():
            logger.error("La ruta base no es una carpeta válida: %s", ruta_base)
            return False

        if not destino_file.exists():
            logger.error("No existe el archivo destino: %s", destino_file)
            return False

        if not destino_file.is_file():
            logger.error("La ruta destino no es un archivo válido: %s", destino_file)
            return False

        if destino_file.suffix.lower() != ".xlsm":
            logger.error("El archivo destino debe ser .xlsm: %s", destino_file)
            return False

        fd, temp_name = tempfile.mkstemp(
            prefix=f".{destino_file.stem}_",
            suffix=".xlsm",
            dir=str(destino_file.parent),
        )
        os.close(fd)

        destino_temporal = Path(temp_name)

        shutil.copy2(str(destino_file), str(destino_temporal))

        logger.info("Inicializando COM...")
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        com_inicializado = True

        logger.info("Iniciando instancia aislada de Excel...")
        excel = com_call(lambda: win32.DispatchEx("Excel.Application"))

        configurar_excel_no_interactivo(excel)

        logger.info("Abriendo copia temporal del archivo destino...")
        wb_destino = com_call(lambda: excel.Workbooks.Open(
            str(destino_temporal),
            UpdateLinks=False,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        ))

        time.sleep(0.5)
        pythoncom.PumpWaitingMessages()

        for config in CONFIG:
            for hoja_destino_nombre in config["mapa"].values():
                if hoja_destino_nombre in cache_destino:
                    continue

                try:
                    cache_destino[hoja_destino_nombre] = com_call(
                        lambda hoja_destino_nombre=hoja_destino_nombre:
                        wb_destino.Worksheets(hoja_destino_nombre)
                    )
                    logger.info("Hoja destino cacheada: %s", hoja_destino_nombre)
                except Exception as e:
                    logger.error(
                        "Error accediendo a hoja destino %s: %s",
                        hoja_destino_nombre,
                        e,
                    )

        for config in CONFIG:
            logger.info("=" * 60)
            logger.info("PROCESANDO: %s", config["nombre"])
            logger.info("=" * 60)

            carpeta = config["carpeta"]
            mapa_destino = config["mapa"]
            hoja_origen_cfg = config["hoja_origen"]
            columnas = config["columnas"]
            excluir = tuple(x.lower() for x in config["excluir"])

            if not carpeta.exists():
                logger.warning("Carpeta no encontrada: %s", carpeta)
                continue

            if not carpeta.is_dir():
                logger.warning("La ruta no es una carpeta válida: %s", carpeta)
                continue

            archivos = sorted(
                archivo.resolve()
                for archivo in carpeta.iterdir()
                if archivo.is_file()
                and archivo.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
                and not archivo.name.lower().startswith(excluir)
                and not archivo.name.startswith("~$")
            )

            if not archivos:
                logger.info("No se encontraron archivos válidos en: %s", carpeta)
                continue

            for archivo in archivos:
                wb_origen = None
                ws_origen = None
                ws_destino = None
                rango_origen_com = None
                rango_destino_com = None
                rango_limpieza_com = None
                data = None

                nombre_archivo = archivo.stem

                try:
                    logger.info("Procesando: %s", nombre_archivo)

                    hoja_destino_nombre = mapa_destino.get(nombre_archivo)

                    if not hoja_destino_nombre:
                        logger.info(
                            "Saltando archivo sin mapa destino: %s",
                            nombre_archivo,
                        )
                        continue

                    if hoja_destino_nombre not in cache_destino:
                        logger.error(
                            "No existe hoja destino cacheada para %s -> %s",
                            nombre_archivo,
                            hoja_destino_nombre,
                        )
                        continue

                    wb_origen = com_call(lambda archivo=archivo: excel.Workbooks.Open(
                        str(archivo),
                        UpdateLinks=False,
                        ReadOnly=True,
                        IgnoreReadOnlyRecommended=True,
                        AddToMru=False,
                    ))

                    pythoncom.PumpWaitingMessages()

                    if hoja_origen_cfg == "ULTIMA_FORMULACION":
                        ws_origen = obtener_ultima_hoja_formulacion(wb_origen)
                    else:
                        ws_origen = com_call(
                            lambda hoja_origen_cfg=hoja_origen_cfg:
                            wb_origen.Worksheets(hoja_origen_cfg)
                        )

                    logger.info(
                        "Hoja origen detectada: '%s'",
                        com_call(lambda: ws_origen.Name),
                    )

                    ws_destino = cache_destino[hoja_destino_nombre]

                    def _get_last_row():
                        rows = ws_origen.Rows
                        rcount = rows.Count
                        cell = ws_origen.Cells(rcount, 1)
                        last_cell = cell.End(XL_UP)
                        return last_cell.Row

                    ultima_fila = com_call(_get_last_row)

                    if ultima_fila < 2:
                        logger.info(
                            "No existen datos para copiar en %s, fila final: %s",
                            nombre_archivo,
                            ultima_fila,
                        )
                        continue

                    rango_origen = f"A2:{columnas}{ultima_fila}"
                    rango_destino = f"A2:{columnas}{ultima_fila}"
                    rango_limpieza = f"A2:{columnas}50000"

                    rango_origen_com = com_call(
                        lambda rango_origen=rango_origen:
                        ws_origen.Range(rango_origen)
                    )
                    data = com_call(lambda: rango_origen_com.Value)

                    rango_limpieza_com = com_call(
                        lambda rango_limpieza=rango_limpieza:
                        ws_destino.Range(rango_limpieza)
                    )
                    com_call(lambda: rango_limpieza_com.ClearContents())

                    rango_destino_com = com_call(
                        lambda rango_destino=rango_destino:
                        ws_destino.Range(rango_destino)
                    )

                    def _set_value():
                        rango_destino_com.Value = data

                    com_call(_set_value)

                    logger.info(
                        "OK -> %s | Filas transferidas: %s",
                        hoja_destino_nombre,
                        ultima_fila - 1,
                    )

                    pythoncom.PumpWaitingMessages()

                except Exception as e:
                    logger.error(
                        "ERROR procesando %s: %s - %s",
                        nombre_archivo,
                        type(e).__name__,
                        e,
                    )

                    # Importante:
                    # El error queda registrado y se continúa con el siguiente archivo.
                    continue

                finally:
                    data = None
                    rango_origen_com = None
                    rango_destino_com = None
                    rango_limpieza_com = None
                    ws_origen = None
                    ws_destino = None

                    if wb_origen is not None:
                        cerrar_workbook_seguro(
                            wb_origen,
                            nombre=str(archivo),
                            save_changes=False,
                        )
                        wb_origen = None

                    gc.collect()

        logger.info("Guardando copia temporal del archivo destino...")
        com_call(lambda: wb_destino.Save())

        cache_destino.clear()
        ws_origen = None
        ws_destino = None
        gc.collect()

        logger.info("Cerrando copia temporal del archivo destino...")
        com_call(lambda: wb_destino.Close(SaveChanges=False))
        wb_destino = None

        if not destino_temporal.exists():
            raise FileNotFoundError(
                f"No existe el archivo temporal guardado: {destino_temporal}"
            )

        if destino_temporal.stat().st_size <= 0:
            raise ValueError(
                f"El archivo temporal guardado está vacío: {destino_temporal}"
            )

        logger.info("Reemplazando archivo destino original...")
        try:
            os.replace(str(destino_temporal), str(destino_file))
            destino_temporal = None
        except Exception:
            conservar_temporal = True
            raise

        proceso_ok = True
        logger.info("Guardado general exitoso: %s", destino_file)

        return True

    except Exception as e:
        logger.exception(
            "ERROR CRÍTICO GENERAL: %s - %s",
            type(e).__name__,
            e,
        )
        return False

    finally:
        logger.info("=== LIMPIANDO RECURSOS EN COPIADO ===")

        try:
            cache_destino.clear()
        except Exception as e:
            logger.warning("No se pudo limpiar cache_destino: %s", e)

        ws_origen = None
        ws_destino = None
        gc.collect()

        if wb_origen is not None:
            cerrar_workbook_seguro(
                wb_origen,
                nombre="origen pendiente",
                save_changes=False,
            )
            wb_origen = None

        if wb_destino is not None:
            cerrar_workbook_seguro(
                wb_destino,
                nombre="destino temporal",
                save_changes=False,
            )
            wb_destino = None

        if excel is not None:
            try:
                com_call(lambda: excel.Quit())
                logger.info("Proceso Excel cerrado con Quit()")
            except Exception as e:
                logger.warning("No se pudo cerrar Excel correctamente con Quit(): %s", e)
            finally:
                excel = None

        gc.collect()
        time.sleep(0.5)

        if com_inicializado:
            try:
                pythoncom.CoUninitialize()
                logger.info("Sistema COM liberado")
            except Exception as e:
                logger.warning("No se pudo liberar COM correctamente: %s", e)

        if destino_temporal is not None:
            try:
                if destino_temporal.exists():
                    if conservar_temporal:
                        logger.warning(
                            "Se conserva el archivo temporal porque falló el reemplazo final: %s",
                            destino_temporal,
                        )
                    else:
                        destino_temporal.unlink()
                        logger.info("Archivo temporal eliminado: %s", destino_temporal)
            except Exception as e:
                logger.warning(
                    "No se pudo gestionar el archivo temporal %s: %s",
                    destino_temporal,
                    e,
                )

        logger.info(
            "Tiempo total de copiado: %.2f segundos | Resultado: %s",
            time.time() - inicio,
            "OK" if proceso_ok else "ERROR",
        )