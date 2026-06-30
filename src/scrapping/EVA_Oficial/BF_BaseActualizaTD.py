import os
import time
import asyncio
from concurrent.futures import ThreadPoolExecutor
import win32com.client as win32
import pywintypes
import pythoncom

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846


# Se mantiene el envoltorio síncrono para colas COM ocupadas
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


# Convertida en la función principal asíncrona
async def actualizar_td(ruta_principal):
    inicio_total = time.time()

    # Subfunción síncrona interna aislada para el Pool de hilos
    def _ejecutar_refresh_sync():
        ruta = os.path.join(ruta_principal, "Base FONAFE WEB al mes.xlsm")
        errores = 0
        excel = None
        wb = None

        # Funciones auxiliares internas para asegurar el contexto del hilo
        def refrescar_conexiones_y_tablas_dinamicas(workbook):
            cont_errores = 0
            try:
                conexiones = workbook.Connections
            except Exception:
                conexiones = []

            for conexion in conexiones:
                try:
                    com_call(lambda: conexion.Refresh())
                except Exception as e:
                    cont_errores += 1
                    print(f"Aviso: no se pudo refrescar la conexión '{conexion.Name}': {e}")

            for hoja in workbook.Worksheets:
                try:
                    tablas = hoja.PivotTables()
                    total_tablas = tablas.Count
                except Exception:
                    continue

                for indice in range(1, total_tablas + 1):
                    try:
                        tabla = tablas.Item(indice)
                        com_call(lambda: tabla.RefreshTable())
                    except Exception as e:
                        cont_errores += 1
                        print(f"Aviso: no se pudo refrescar la tabla dinámica en '{hoja.Name}': {e}")
            return cont_errores

        try:
            # CRÍTICO: Inicializar COM en el entorno de este hilo secundario
            pythoncom.CoInitialize()

            if not os.path.exists(ruta):
                raise FileNotFoundError(f"No se encontró el archivo: {ruta}")

            # EnsureDispatch fuerza el enlace en un hilo aislado de forma segura
            excel = win32.gencache.EnsureDispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            print("Abriendo libro de trabajo en segundo plano...")
            wb = com_call(lambda: excel.Workbooks.Open(ruta))

            ws = wb.Sheets("DEP Y COLOC_SALDOF_EFECT Y EQUI")
            print("Actualizando conexiones y tablas dinámicas...")
            
            errores = refrescar_conexiones_y_tablas_dinamicas(wb)

            # Esperar a que terminen los procesos de actualización en segundo plano de Excel
            while excel.CalculationState != 0:
                time.sleep(0.5)

            print("Aplicando copiado especial de valores...")
            ws.Range("A44:B85").Copy()
            ws.Range("D44").PasteSpecial(Paste=-4163)  # xlPasteValues
            excel.CutCopyMode = False

            com_call(lambda: wb.Save())

            # Esperar confirmación de recálculo tras guardar
            while excel.CalculationState != 0:
                time.sleep(0.5)

            if errores == 0:
                print("✓ Proceso completado correctamente sin advertencias.")
            else:
                print(f"⚠ Proceso completado con {errores} advertencia(s) de refresco.")

        except Exception as e:
            print(f"✗ Error durante la ejecución del proceso COM: {e}")
        finally:
            # Control de cierre seguro sin com_call para evitar bloqueos si falló el puntero
            if wb:
                try:
                    wb.Close(SaveChanges=True)
                except Exception: pass
            if excel:
                try:
                    excel.Quit()
                except Exception: pass

            wb = excel = None
            
            # Forzar liberación de memoria RAM y desinicializar COM en el hilo
            import gc
            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except Exception: pass

    # --- COORDINACIÓN ASÍNCRONA ---
    loop = asyncio.get_running_loop()
    
    with ThreadPoolExecutor(max_workers=1) as executor:
        await loop.run_in_executor(executor, _ejecutar_refresh_sync)
        
    print(f"Tiempo total de actualización: {round(time.time() - inicio_total, 2)} segundos.")

