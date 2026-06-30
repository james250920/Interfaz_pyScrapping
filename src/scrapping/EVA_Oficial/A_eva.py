import os
import time
import asyncio
from concurrent.futures import ThreadPoolExecutor
import psutil
import pythoncom
import win32com.client as win32

# Convertida a función principal asíncrona
async def copiar_pegar(ruta_principal, anio, mes):
    inicio = time.time()

    # Subfunción síncrona interna para aislar el proceso COM en el pool
    def _procesar_sync():
        RUTA_DESTINO = os.path.join(ruta_principal, "Eva 2026 - Información al mes.xlsm")
        RUTA_ORIGEN = os.path.join(ruta_principal, "Base FONAFE WEB al mes.xlsm")

        OPERACIONES = [
            ("PRE ", "C4:Z8000", "BPRE", "AG4:BD8000"),
            ("FLU", "C4:Z8000", "BFLU", "AG4:BD8000"),
            ("ESF", "C4:Z8000", "BESF", "AG4:BD8000"),
            ("ERI", "C4:Z8000", "BERI", "AG4:BD8000"),
        ]

        def cerrar_excels():
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.info['name'] and 'EXCEL' in proc.info['name'].upper():
                        print(f"Cerrando Excel residual PID: {proc.info['pid']}")
                        proc.kill()
                except:
                    pass

        excel = None
        wb_origen = None
        wb_destino = None

        try:
            # 1. Limpieza inicial de procesos huerfanos
            cerrar_excels()
            time.sleep(2)

            # CRÍTICO: Registrar el subsistema COM para el entorno de este hilo
            pythoncom.CoInitialize()

            if not os.path.exists(RUTA_ORIGEN):
                raise FileNotFoundError(f"No existe:\n{RUTA_ORIGEN}")

            if not os.path.exists(RUTA_DESTINO):
                raise FileNotFoundError(f"No existe:\n{RUTA_DESTINO}")

            print("✓ Archivos verificados en disco.")

            # 2. Inicialización segura del Servidor de Excel
            excel = win32.DispatchEx("Excel.Application")
            
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False
            excel.Interactive = False

            time.sleep(1)

            print("Abriendo archivo origen...")
            wb_origen = excel.Workbooks.Open(RUTA_ORIGEN, UpdateLinks=0, ReadOnly=True)

            print("Abriendo archivo destino...")
            wb_destino = excel.Workbooks.Open(RUTA_DESTINO, UpdateLinks=0)

            # 3. Escritura de variables de cabecera
            hoja_bpre = wb_destino.Worksheets("BPRE")
            hoja_bpre.Range("A2").Value = anio
            hoja_bpre.Range("B2").Value = mes
            print(f"✓ Cabecera BPRE actualizada ({anio} - Mes: {mes})")

            # 4. Bloque de transferencia de datos masivos
            for hoja_orig, rango_orig, hoja_dest, rango_dest in OPERACIONES:
                try:
                    ws_origen = wb_origen.Worksheets(hoja_orig)
                    ws_destino = wb_destino.Worksheets(hoja_dest)

                    # Transferencia directa por matriz de memoria (Rápido)
                    datos = ws_origen.Range(rango_orig).Value
                    ws_destino.Range(rango_dest).Value = datos

                    print(f"  ✓ {hoja_orig} → {hoja_dest}")
                except Exception as e:
                    print(f"  ✗ Error transfiriendo pestaña {hoja_orig}: {e}")

            print("Guardando libro de destino...")
            wb_destino.Save()
            print("✓ Archivo maestro guardado correctamente.")

        except Exception as e:
            print(f"\n✗ ERROR GENERAL EN HILO DE TRANSFERENCIA:\n{e}")

        finally:
            # Clausura y liberación estructurada de objetos del hilo
            try:
                if wb_origen: wb_origen.Close(False)
            except: pass
            try:
                if wb_destino: wb_destino.Close(True)
            except: pass
            try:
                if excel: excel.Quit()
            except: pass

            excel = wb_origen = wb_destino = None
            
            # Forzado de recolección de basura de punteros COM
            import gc
            gc.collect()
            pythoncom.CoUninitialize()
            
            # Limpieza final de procesos retenidos en segundo plano
            cerrar_excels()

    # --- COORDINACIÓN ASÍNCRONA ---
    loop = asyncio.get_running_loop()
    
    # max_workers=1 asegura que no interfiera con otros libros COM abiertos en paralelo
    with ThreadPoolExecutor(max_workers=1) as executor:
        await loop.run_in_executor(executor, _procesar_sync)

    print(f"\nTiempo total de la operación: {round(time.time() - inicio, 2)} segundos")

