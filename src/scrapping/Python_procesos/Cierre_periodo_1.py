import time
import gc
import os
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


def cierre_periodo(ruta_principal):
    ruta_base = Path(ruta_principal).expanduser().resolve()

    ruta_destino = ruta_base / "Estado de Cierre al mes.xlsm"
    ruta_ejec = ruta_base / "CIERRE" / "Estado_de_Cierre_del_Periodo_Ejecucion.xlsx"
    ruta_marco = ruta_base / "CIERRE" / "Estado_de_Cierre_del_Periodo_Formulacion.xlsx"

    inicio = time.time()

    com_inicializado = False
    proceso_ok = False

    excel = None
    wb_ejec = None
    wb_marco = None
    wb_destino = None

    ws_origen_ejec = None
    ws_origen_marco = None
    ws_destino_ejec = None
    ws_destino_marco = None

    ruta_destino_temporal = None

    try:
        missing = []
        for p, label in (
            (ruta_ejec, "CIERRE Ejecución"),
            (ruta_marco, "CIERRE Marco"),
            (ruta_destino, "Destino"),
        ):
            if not p.exists():
                missing.append((label, p))

        if missing:
            logger.error("Archivos faltantes para cierre de periodo:")
            for label, p in missing:
                logger.error("  - %s: %s", label, p)
            return

        if not ruta_ejec.is_file():
            logger.error("La ruta de ejecución no es un archivo válido: %s", ruta_ejec)
            return

        if not ruta_marco.is_file():
            logger.error("La ruta de formulación no es un archivo válido: %s", ruta_marco)
            return

        if not ruta_destino.is_file():
            logger.error("La ruta destino no es un archivo válido: %s", ruta_destino)
            return

        if ruta_ejec.suffix.lower() != ".xlsx":
            logger.error("El archivo de ejecución debe ser .xlsx: %s", ruta_ejec)
            return

        if ruta_marco.suffix.lower() != ".xlsx":
            logger.error("El archivo de formulación debe ser .xlsx: %s", ruta_marco)
            return

        if ruta_destino.suffix.lower() != ".xlsm":
            logger.error("El archivo destino debe ser .xlsm: %s", ruta_destino)
            return

        fd, temp_name = tempfile.mkstemp(
            prefix=f".{ruta_destino.stem}_",
            suffix=".xlsm",
            dir=str(ruta_destino.parent),
        )
        os.close(fd)

        ruta_destino_temporal = Path(temp_name)

        shutil.copy2(str(ruta_destino), str(ruta_destino_temporal))

        logger.info("Inicializando COM...")
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        com_inicializado = True

        logger.info("Iniciando instancia aislada de Excel...")
        excel = com_call(lambda: win32.DispatchEx("Excel.Application"))

        com_call(lambda: setattr(excel, "Visible", False))
        com_call(lambda: setattr(excel, "DisplayAlerts", False))
        com_call(lambda: setattr(excel, "AskToUpdateLinks", False))
        com_call(lambda: setattr(excel, "ScreenUpdating", False))
        com_call(lambda: setattr(excel, "EnableEvents", False))
        com_call(lambda: setattr(excel, "DisplayStatusBar", False))

        try:
            com_call(lambda: setattr(excel, "AutomationSecurity", 3))
        except Exception as e:
            logger.warning("No se pudo establecer AutomationSecurity=3: %s", e)

        logger.info("Abriendo archivo de ejecución...")
        wb_ejec = com_call(lambda: excel.Workbooks.Open(
            str(ruta_ejec),
            UpdateLinks=False,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        ))
        pythoncom.PumpWaitingMessages()
        time.sleep(0.5)

        logger.info("Abriendo archivo de formulación...")
        wb_marco = com_call(lambda: excel.Workbooks.Open(
            str(ruta_marco),
            UpdateLinks=False,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        ))
        pythoncom.PumpWaitingMessages()
        time.sleep(0.5)

        logger.info("Abriendo copia temporal del archivo destino...")
        wb_destino = com_call(lambda: excel.Workbooks.Open(
            str(ruta_destino_temporal),
            UpdateLinks=False,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        ))
        pythoncom.PumpWaitingMessages()
        time.sleep(0.5)

        ws_origen_ejec = com_call(lambda: wb_ejec.Worksheets(1))
        ws_origen_marco = com_call(lambda: wb_marco.Worksheets(1))

        try:
            ws_destino_ejec = com_call(lambda: wb_destino.Worksheets("Estado Cierre Ejec"))
            ws_destino_marco = com_call(lambda: wb_destino.Worksheets("Estado Cierre Marco"))
        except Exception:
            nombres_hojas = []
            sheets = None

            try:
                sheets = wb_destino.Worksheets
                for i in range(1, sheets.Count + 1):
                    hoja = None
                    try:
                        hoja = sheets(i)
                        nombres_hojas.append(hoja.Name)
                    finally:
                        hoja = None
            finally:
                sheets = None

            logger.error(
                "No se encontraron las hojas destino esperadas. Hojas disponibles:"
            )
            for nombre in nombres_hojas:
                logger.error("  - %s", nombre)

            return

        logger.info("Limpiando contenido destino...")

        rango_destino_ejec = None
        rango_destino_marco = None

        try:
            rango_destino_ejec = com_call(lambda: ws_destino_ejec.Range("B1:M1000"))
            com_call(lambda: rango_destino_ejec.ClearContents())

            rango_destino_marco = com_call(lambda: ws_destino_marco.Range("B1:M1000"))
            com_call(lambda: rango_destino_marco.ClearContents())
        finally:
            rango_destino_ejec = None
            rango_destino_marco = None
            gc.collect()

        def _copiar_valores(ws_origen, ws_destino, rango="B1:M1000"):
            rango_origen = None
            rango_destino = None
            valores = None

            try:
                rango_origen = com_call(lambda: ws_origen.Range(rango))
                valores = com_call(lambda: rango_origen.Value)

                rango_destino = com_call(lambda: ws_destino.Range(rango))

                def _asignar():
                    rango_destino.Value = valores

                com_call(_asignar)

            finally:
                valores = None
                rango_origen = None
                rango_destino = None
                gc.collect()

        logger.info("Copiando Estado Cierre Ejec...")
        _copiar_valores(ws_origen_ejec, ws_destino_ejec)

        logger.info("Copiando Estado Cierre Marco...")
        _copiar_valores(ws_origen_marco, ws_destino_marco)

        logger.info("Guardando copia temporal del libro destino...")
        com_call(lambda: wb_destino.Save())

        logger.info("Cerrando copia temporal del libro destino...")
        com_call(lambda: wb_destino.Close(SaveChanges=False))
        wb_destino = None

        if not ruta_destino_temporal.exists():
            raise FileNotFoundError(
                f"No existe el archivo temporal guardado: {ruta_destino_temporal}"
            )

        if ruta_destino_temporal.stat().st_size <= 0:
            raise ValueError(
                f"El archivo temporal guardado está vacío: {ruta_destino_temporal}"
            )

        logger.info("Reemplazando archivo destino original...")
        os.replace(str(ruta_destino_temporal), str(ruta_destino))
        ruta_destino_temporal = None

        proceso_ok = True
        logger.info("Archivo guardado correctamente: %s", ruta_destino)

    except Exception as e:
        logger.exception("ERROR en Cierre Periodo: %s - %s", type(e).__name__, e)

    finally:
        logger.info("=== LIMPIANDO RECURSOS EN CIERRE ===")

        ws_origen_ejec = None
        ws_origen_marco = None
        ws_destino_ejec = None
        ws_destino_marco = None
        gc.collect()

        try:
            if wb_ejec is not None:
                com_call(lambda: wb_ejec.Close(SaveChanges=False))
        except Exception as e:
            logger.warning("No se pudo cerrar wb_ejec correctamente: %s", e)
        finally:
            wb_ejec = None

        try:
            if wb_marco is not None:
                com_call(lambda: wb_marco.Close(SaveChanges=False))
        except Exception as e:
            logger.warning("No se pudo cerrar wb_marco correctamente: %s", e)
        finally:
            wb_marco = None

        try:
            if wb_destino is not None:
                com_call(lambda: wb_destino.Close(SaveChanges=False))
        except Exception as e:
            logger.warning("No se pudo cerrar wb_destino correctamente: %s", e)
        finally:
            wb_destino = None

        try:
            if excel is not None:
                com_call(lambda: excel.Quit())
        except Exception as e:
            logger.warning("No se pudo cerrar Excel correctamente con Quit(): %s", e)
        finally:
            excel = None

        gc.collect()
        time.sleep(0.5)

        if com_inicializado:
            try:
                pythoncom.CoUninitialize()
                logger.info("COM desinicializado en Cierre")
            except Exception as e:
                logger.warning("No se pudo desinicializar COM correctamente: %s", e)

        if ruta_destino_temporal is not None:
            try:
                if ruta_destino_temporal.exists():
                    ruta_destino_temporal.unlink()
            except Exception as e:
                logger.warning(
                    "No se pudo eliminar el archivo temporal %s: %s",
                    ruta_destino_temporal,
                    e,
                )

        logger.info(
            "Tiempo total: %.2f segundos | Resultado: %s",
            time.time() - inicio,
            "OK" if proceso_ok else "ERROR",
        )