import os
import time
from datetime import datetime
import win32com.client as win32
import pywintypes
import pythoncom


RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846


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


def actualizar_datos_iniciales(ruta_principal, anio, mes,fecha_cierre_sistema):
    RUTA = rf"{ruta_principal}\\Base FONAFE WEB al mes.xlsm"
    HOJA = "Sistema Cierre"

    periodo_actual = mes
    anio_actual = anio
    fecha_cierre_sistema = fecha_cierre_sistema
    ahora = datetime.now()

    fecha_actual = ahora.strftime("%d.%m.%Y")
    hora_actual = ahora.strftime("%H:%M:%S")

    # Inicializar COM en este hilo antes de usar objetos COM
    pythoncom.CoInitialize()
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    wb = None
    try:

        wb = com_call(lambda: excel.Workbooks.Open(
            RUTA,
            UpdateLinks=False,
            ReadOnly=False
        ))
        com_call(lambda: setattr(excel, "Calculation", win32.constants.xlCalculationManual))

        ws = wb.Worksheets(HOJA)

        ws.Range("C4").Value = periodo_actual
        ws.Range("C5").Value = anio_actual

        ws.Range("G2").Value = fecha_cierre_sistema
        ws.Range("G3").Value = fecha_actual
        ws.Range("G4").Value = hora_actual

        wb.Save()

        print("Proceso completado.")

    finally:

        try:
            com_call(lambda: setattr(excel, "Calculation", win32.constants.xlCalculationAutomatic))
        except:
            pass

        try:
            excel.ScreenUpdating = True
            excel.EnableEvents = True
        except Exception:
            pass

        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            excel.Quit()
        except Exception:
            pass

        # Desinicializar COM
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass