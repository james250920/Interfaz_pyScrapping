import os
import time
import asyncio
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import win32com.client as win32
import pywintypes
import pythoncom
import gc

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846

# Constantes literales en vez de win32.constants.* : con DispatchEx (late
# binding) el módulo win32com.client.constants solo se rellena si existe
# caché de gencache (EnsureDispatch/makepy). Sin esa caché, cualquier acceso
# a win32.constants.ALGO lanza AttributeError. Como este AttributeError no
# es un pywintypes.com_error, com_call() NO lo reintenta y, peor, el try/
# finally original no tenía ningún except, así que la excepción se propagaba
# sin capturar y mataba todo el pipeline en silencio.
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105

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
        
        # Se usa DispatchEx (no EnsureDispatch) para forzar SIEMPRE una instancia
        # nueva y aislada de Excel. EnsureDispatch consulta el ROT (Running Object
        # Table) de Windows y puede engancharse a un Excel que el usuario ya tenga
        # abierto manualmente, heredando diálogos modales pendientes o cambios sin
        # guardar — eso es lo que provoca los congelamientos intermitentes.
        excel = win32.DispatchEx("Excel.Application")
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
            
            com_call(lambda: setattr(excel, "Calculation", XL_CALCULATION_MANUAL))

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

        except Exception as e:
            # IMPORTANTE: antes no existía este except. Cualquier error
            # (incluso uno inesperado no relacionado a COM) se propagaba
            # sin control y mataba todo el pipeline sin mensaje claro.
            print(f"✗ ERROR al actualizar datos iniciales: {e}")
            raise

        finally:
            # Limpieza segura de recursos dentro del hilo
            try:
                com_call(lambda: setattr(excel, "Calculation", XL_CALCULATION_AUTOMATIC))
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

            wb = None
            excel = None
            gc.collect()

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