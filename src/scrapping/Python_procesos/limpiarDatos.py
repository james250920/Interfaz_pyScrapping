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


def limpiar_datos(ruta_archivo, mes):
    inicio = time.time()

    ruta_base = Path(ruta_archivo).expanduser().resolve()
    destino_file = ruta_base / "Validación Data Marco y Ejecución al mes.xlsm"

    excel = None
    wb = None
    ws = None
    ws_calidad = None
    rango = None

    com_inicializado = False
    proceso_ok = False
    destino_temporal = None
    conservar_temporal = False

    configuracion = [
        ("Sistema-Form. PRE", "A2:AB5000"),
        ("Sistema-Ejec. PRE", "A2:U5000"),
        ("Sistema-Form. FC", "A2:AB5000"),
        ("Sistema-Ejec. FC", "A2:U5000"),
        ("Sistema-Form. ESF", "A2:AB5000"),
        ("Sistema-Ejec. ESF", "A2:U5000"),
        ("Sistema-Form. ERI", "A2:AB5000"),
        ("Sistema-Ejec. ERI", "A2:U5000"),
    ]

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

        logger.info("Iniciando Excel mediante instancia COM aislada...")
        excel = com_call(lambda: win32.DispatchEx("Excel.Application"))

        com_call(lambda: setattr(excel, "Visible", False))
        com_call(lambda: setattr(excel, "DisplayAlerts", False))
        com_call(lambda: setattr(excel, "ScreenUpdating", False))
        com_call(lambda: setattr(excel, "EnableEvents", False))
        com_call(lambda: setattr(excel, "AskToUpdateLinks", False))
        com_call(lambda: setattr(excel, "DisplayStatusBar", False))

        try:
            com_call(lambda: setattr(excel, "AutomationSecurity", 3))
        except Exception as e:
            logger.warning("No se pudo establecer AutomationSecurity=3: %s", e)

        logger.info("Abriendo copia temporal del archivo destino...")
        wb = com_call(lambda: excel.Workbooks.Open(
            str(destino_temporal),
            UpdateLinks=False,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        ))

        pythoncom.PumpWaitingMessages()
        time.sleep(0.5)

        try:
            ws_calidad = com_call(lambda: wb.Worksheets("Calidad de Data"))
            rango = com_call(lambda: ws_calidad.Range("E2"))
            com_call(lambda: setattr(rango, "Value", mes))

            logger.info("Mes %s colocado en 'Calidad de Data'!E2", mes)

        except Exception as e:
            logger.error("Error al actualizar hoja Calidad de Data: %s", e)

        finally:
            rango = None
            ws_calidad = None
            gc.collect()

        for hoja, direccion_rango in configuracion:
            ws = None
            rango = None

            try:
                ws = com_call(lambda hoja=hoja: wb.Worksheets(hoja))
                rango = com_call(lambda direccion_rango=direccion_rango: ws.Range(direccion_rango))

                logger.info("Limpiando %s | Rango %s", hoja, direccion_rango)

                com_call(lambda: rango.ClearContents())

            except Exception as e:
                logger.error("Error en hoja %s: %s", hoja, e)

            finally:
                rango = None
                ws = None
                gc.collect()

        logger.info("Guardando copia temporal...")
        com_call(lambda: wb.Save())

        logger.info("Cerrando libro temporal...")
        com_call(lambda: wb.Close(SaveChanges=False))
        wb = None

        if not destino_temporal.exists():
            raise FileNotFoundError(
                f"No existe el archivo temporal guardado: {destino_temporal}"
            )

        if destino_temporal.stat().st_size <= 0:
            raise ValueError(
                f"El archivo temporal guardado está vacío: {destino_temporal}"
            )

        logger.info("Reemplazando archivo original...")
        try:
            os.replace(str(destino_temporal), str(destino_file))
            destino_temporal = None
        except Exception:
            conservar_temporal = True
            raise

        proceso_ok = True

        logger.info("Proceso terminado correctamente")
        return True

    except Exception as e:
        logger.exception("Error durante limpiar_datos: %s - %s", type(e).__name__, e)
        return False

    finally:
        logger.info("=== INICIANDO LIMPIEZA DE RECURSOS ===")

        rango = None
        ws = None
        ws_calidad = None
        gc.collect()

        if wb is not None:
            try:
                com_call(lambda: wb.Close(SaveChanges=False))
                logger.info("Libro cerrado")
            except Exception as e:
                logger.warning("No se pudo cerrar el libro: %s", e)
            finally:
                wb = None

        if excel is not None:
            try:
                com_call(lambda: excel.Quit())
                logger.info("Excel cerrado con Quit()")
            except Exception as e:
                logger.warning("Error al cerrar Excel con Quit(): %s", e)
            finally:
                excel = None

        gc.collect()
        time.sleep(0.5)

        if com_inicializado:
            try:
                pythoncom.CoUninitialize()
                logger.info("COM liberado")
            except Exception as e:
                logger.warning("Error liberando COM: %s", e)

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
            "Tiempo total: %.2f segundos | Resultado: %s",
            time.time() - inicio,
            "OK" if proceso_ok else "ERROR",
        )