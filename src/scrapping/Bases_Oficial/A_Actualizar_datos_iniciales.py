import os
import time
import asyncio
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import win32com.client as win32
import pywintypes
import pythoncom

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846

# Se mantiene la lógica exacta de reintentos síncronos de COM
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

# Convertida a función principal asíncrona
async def actualizar_datos_iniciales(ruta_principal, anio, mes, fecha_cierre_sistema):
    inicio_tiempo = time.time()
    
    # Lógica interna síncrona aislada para el Executor
    def _actualizar_sync():
        RUTA = os.path.join(ruta_principal, "Base FONAFE WEB al mes.xlsm")
        HOJA = "Sistema Cierre"

        periodo_actual = mes
        anio_actual = anio
        ahora = datetime.now()

        fecha_actual = ahora.strftime("%d.%m.%Y")
        hora_actual = ahora.strftime("%H:%M:%S")

        # CRÍTICO: Inicializar COM en este hilo secundario antes de instanciar Excel
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

            # Asignación exacta de valores en las celdas correspondientes
            ws.Range("C4").Value = periodo_actual
            ws.Range("C5").Value = anio_actual

            ws.Range("G2").Value = fecha_cierre_sistema
            ws.Range("G3").Value = fecha_actual
            ws.Range("G4").Value = hora_actual

            print("Guardando 'Base FONAFE WEB al mes.xlsm'...")
            wb.Save()
            print("Proceso en Excel completado con éxito.")

        finally:
            # Limpieza segura de recursos dentro del hilo
            try:
                com_call(lambda: setattr(excel, "Calculation", win32.constants.xlCalculationAutomatic))
            except:
                pass

            try:
                excel.ScreenUpdating = True
                excel.EnableEvents = True
            except:
                pass

            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except:
                pass

            try:
                excel.Quit()
            except:
                pass

            # CRÍTICO: Liberar COM al finalizar el trabajo del hilo
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    # --- COORDINACIÓN ASÍNCRONA ---
    loop = asyncio.get_running_loop()
    
    # Forzamos max_workers=1 porque Excel Interop de escritorio no admite paralelismo real sobre sí mismo
    with ThreadPoolExecutor(max_workers=1) as executor:
        await loop.run_in_executor(executor, _actualizar_sync)
        
    print(f"✓ Datos iniciales actualizados en {round(time.time() - inicio_tiempo, 2)} segundos.")


