import os
import time
import tempfile
import shutil
import logging
import gc
from datetime import datetime

import win32com.client as win32
import pywintypes
import pythoncom


logger = logging.getLogger(__name__)

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846

# Constantes literales en vez de win32.constants.*:
# Con DispatchEx (late binding), win32com.client.constants puede no estar
# disponible si no existe caché de gencache/makepy.
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105

# msoAutomationSecurityForceDisable = 3
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def com_call(fn, reintentos=12, pausa=1.0):
    """
    Ejecuta una llamada COM con reintentos solo ante errores transitorios
    típicos de Office ocupado/rechazando llamadas.
    """
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


def actualizar_datos_iniciales(ruta_principal, anio, mes, fecha_cierre_sistema):
    inicio_tiempo = time.time()

    ruta_principal_abs = os.path.abspath(ruta_principal)
    RUTA = os.path.abspath(os.path.join(ruta_principal_abs, "Base FONAFE WEB al mes.xlsm"))
    HOJA = "Sistema Cierre"

    periodo_actual = mes
    anio_actual = anio
    ahora = datetime.now()

    fecha_actual = ahora.strftime("%d.%m.%Y")
    hora_actual = ahora.strftime("%H:%M:%S")

    com_inicializado = False
    excel = None
    wb = None
    ws = None

    rango_c4 = None
    rango_c5 = None
    rango_g2 = None
    rango_g3 = None
    rango_g4 = None

    ruta_temporal = None
    guardado_correcto = False
    reemplazo_correcto = False

    try:
        # Validaciones de rutas antes de abrir COM.
        if not os.path.isdir(ruta_principal_abs):
            raise NotADirectoryError(f"La carpeta principal no existe: {ruta_principal_abs}")

        if not os.path.exists(RUTA):
            raise FileNotFoundError(f"El archivo origen no existe: {RUTA}")

        if not os.path.isfile(RUTA):
            raise FileNotFoundError(f"La ruta origen no corresponde a un archivo válido: {RUTA}")

        # Inicialización COM explícita en el hilo actual.
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        com_inicializado = True

        excel = win32.DispatchEx("Excel.Application")

        # Configuración no interactiva.
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.DisplayStatusBar = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        # Crear archivo temporal en la misma carpeta para que os.replace()
        # sea atómico dentro del mismo volumen.
        fd, temp_name = tempfile.mkstemp(
            prefix="Base FONAFE WEB al mes_",
            suffix=".xlsm",
            dir=os.path.dirname(RUTA)
        )
        os.close(fd)
        ruta_temporal = os.path.abspath(temp_name)

        # Copia temporal. El archivo original no se modifica directamente.
        shutil.copy2(RUTA, ruta_temporal)

        wb = com_call(lambda: excel.Workbooks.Open(
            ruta_temporal,
            UpdateLinks=False,
            ReadOnly=False
        ))

        com_call(lambda: setattr(excel, "Calculation", XL_CALCULATION_MANUAL))

        ws = com_call(lambda: wb.Worksheets(HOJA))

        # Evitar encadenamientos COM innecesarios.
        rango_c4 = com_call(lambda: ws.Range("C4"))
        rango_c5 = com_call(lambda: ws.Range("C5"))
        rango_g2 = com_call(lambda: ws.Range("G2"))
        rango_g3 = com_call(lambda: ws.Range("G3"))
        rango_g4 = com_call(lambda: ws.Range("G4"))

        rango_c4.Value = periodo_actual
        rango_c5.Value = anio_actual

        rango_g2.Value = fecha_cierre_sistema
        rango_g3.Value = fecha_actual
        rango_g4.Value = hora_actual

        logger.info("Guardando 'Base FONAFE WEB al mes.xlsm' en archivo temporal...")
        com_call(lambda: wb.Save())
        guardado_correcto = True
        logger.info("Proceso en Excel completado con éxito.")

    except Exception as e:
        logger.exception("Error al actualizar datos iniciales: %s", e)
        raise

    finally:
        # Liberar referencias COM específicas primero.
        rango_g4 = None
        rango_g3 = None
        rango_g2 = None
        rango_c5 = None
        rango_c4 = None

        if excel is not None:
            try:
                com_call(lambda: setattr(excel, "Calculation", XL_CALCULATION_AUTOMATIC))
            except Exception as e:
                logger.warning("No se pudo restaurar el modo de cálculo de Excel: %s", e)

            try:
                excel.ScreenUpdating = True
                excel.EnableEvents = True
                excel.DisplayStatusBar = True
            except Exception as e:
                logger.warning("No se pudieron restaurar algunas propiedades de Excel: %s", e)

        ws = None

        if wb is not None:
            try:
                # Ya se hizo Save explícito. No guardar cambios adicionales al cerrar.
                com_call(lambda: wb.Close(SaveChanges=False))
            except Exception as e:
                logger.warning("No se pudo cerrar el workbook correctamente: %s", e)
            finally:
                wb = None

        if excel is not None:
            try:
                excel.Quit()
            except Exception as e:
                logger.warning("No se pudo cerrar la aplicación Excel correctamente: %s", e)
            finally:
                excel = None

        gc.collect()

        if com_inicializado:
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.warning("No se pudo finalizar COM correctamente: %s", e)

        # Reemplazo atómico solo si el guardado temporal fue exitoso.
        if ruta_temporal:
            if guardado_correcto:
                try:
                    os.replace(ruta_temporal, RUTA)
                    reemplazo_correcto = True
                    logger.info("Archivo modificado y reemplazado atómicamente: %s", RUTA)
                except Exception as e:
                    logger.warning(
                        "No se pudo reemplazar el archivo original. "
                        "El temporal quedó en: %s. Error: %s",
                        ruta_temporal,
                        e
                    )
            else:
                # Si falló el flujo antes del guardado correcto, no se reemplaza el original.
                try:
                    if os.path.exists(ruta_temporal):
                        os.remove(ruta_temporal)
                except Exception as e:
                    logger.warning(
                        "No se pudo eliminar el archivo temporal fallido: %s. Error: %s",
                        ruta_temporal,
                        e
                    )

    if guardado_correcto and reemplazo_correcto:
        logger.info(
            "Datos iniciales actualizados en %s segundos.",
            round(time.time() - inicio_tiempo, 2)
        )