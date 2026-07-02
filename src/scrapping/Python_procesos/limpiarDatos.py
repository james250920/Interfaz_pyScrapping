import os
import time
import subprocess
import asyncio
from concurrent.futures import ThreadPoolExecutor
import win32com.client as win32
import pythoncom
import gc

# Modificada a función principal asíncrona
async def limpiar_datos(ruta_archivo, mes):
    inicio = time.time()
    destino_file = os.path.join(ruta_archivo, "Validación Data Marco y Ejecución al mes.xlsm")
    
    # Esta es la lógica síncrona original que interactúa con la API COM de Excel
    def _limpiar_datos_sync():
        excel = None
        wb = None
        try:
            print("Iniciando Excel mediante interfaz COM...")
            # CRÍTICO: Inicializa el entorno COM para este hilo específico del pool
            pythoncom.CoInitialize()
            
            excel = win32.gencache.EnsureDispatch("Excel.Application")
            
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False
            excel.DisplayStatusBar = False
            
            try:
                for wb_existente in excel.Workbooks:
                    if wb_existente.FullName == destino_file:
                        print("El archivo ya estaba abierto, cerrándolo...")
                        wb_existente.Close(SaveChanges=False)
                        break
            except:
                pass
            
            wb = excel.Workbooks.Open(
                destino_file,
                UpdateLinks=False,
                ReadOnly=False
            )

            try:
                ws_calidad = wb.Worksheets("Calidad de Data")
                ws_calidad.Range("E2").Value = mes
                print(f"Mes {mes} colocado en Calidad de Data!E2")
            except Exception as e:
                print(f"Error al actualizar hoja Calidad de Data: {e}")
            
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
            
            for hoja, rango in configuracion:
                try:
                    ws = wb.Worksheets(hoja)
                    print(f"Limpiando {hoja}")
                    ws.Range(rango).ClearContents()
                except Exception as e:
                    print(f"Error en hoja {hoja}: {e}")
            
            print("Guardando...")
            wb.Save()
            
        except Exception as e:
            print(f"Error durante el proceso interno: {e}")
            
        finally:
            print("\n=== INICIANDO LIMPIEZA DE RECURSOS ===")
            if wb:
                try:
                    wb.Close(SaveChanges=True)
                    print("✓ Libro cerrado")
                except Exception as e:
                    print(f"✗ No se pudo cerrar el libro: {e}")
            
            if excel:
                try:
                    excel.Quit()
                    print("✓ Excel cerrado")
                except Exception as e:
                    print(f"✗ Error al cerrar Excel: {e}")
            
            wb = None
            excel = None
            gc.collect()
            
            try:
                # CRÍTICO: Libera el entorno COM antes de cerrar el hilo
                pythoncom.CoUninitialize()
                print("✓ COM liberado")
            except Exception as e:
                print(f"✗ Error liberando COM: {e}")

    # --- COORDINACIÓN ASÍNCRONA ---
    loop = asyncio.get_running_loop()
    
    # Ejecutamos la lógica COM pesada en un hilo separado para que no congele la app
    # max_workers=1 es ideal aquí para evitar colisiones en la API de Excel de escritorio
    with ThreadPoolExecutor(max_workers=1) as executor:
        await loop.run_in_executor(executor, _limpiar_datos_sync)
        
    fin = time.time()
    print(f"\n✓ Proceso terminado en {round(fin - inicio, 2)} segundos")
