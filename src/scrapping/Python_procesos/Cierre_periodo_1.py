import time
import gc
import os
import asyncio
from concurrent.futures import ThreadPoolExecutor
import pythoncom
import win32com.client as win32
import pywintypes


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

# Convertida a función principal asíncrona
async def cierre_periodo(ruta_principal):
    # Subfunción síncrona interna que correrá en ThreadPoolExecutor
    def _cierre_sync():
        ruta_destino = os.path.join(ruta_principal, "Estado de Cierre al mes.xlsm")
        ruta_ejec = os.path.join(ruta_principal, "CIERRE", "Estado_de_Cierre_del_Periodo_Ejecucion.xlsx")
        ruta_marco = os.path.join(ruta_principal, "CIERRE", "Estado_de_Cierre_del_Periodo_Formulacion.xlsx")
        
        inicio = time.time()
        excel = None
        wb_ejec = None
        wb_marco = None
        wb_destino = None
        ws_origen_ejec = None
        ws_origen_marco = None
        ws_destino_ejec = None
        ws_destino_marco = None

        try:

            # Verificar que los archivos existen antes de intentar abrir Excel
            missing = []
            for p, label in ((ruta_ejec, 'CIERRE Ejecución'), (ruta_marco, 'CIERRE Marco'), (ruta_destino, 'Destino')):
                if not os.path.exists(p):
                    missing.append((label, p))

            if missing:
                print("✗ Archivos faltantes para cierre de periodo:")
                for lab, p in missing:
                    print(f"  - {lab}: {p}")
                return

            pythoncom.CoInitialize()

            excel = com_call(lambda: win32.DispatchEx("Excel.Application"))

            try:
                excel.Visible = False
                excel.DisplayAlerts = False
                excel.ScreenUpdating = False
                excel.EnableEvents = False
            except Exception:
                pass

            print("Abriendo archivos...")

            wb_ejec = com_call(lambda: excel.Workbooks.Open(ruta_ejec, UpdateLinks=False, ReadOnly=True))
            wb_marco = com_call(lambda: excel.Workbooks.Open(ruta_marco, UpdateLinks=False, ReadOnly=True))
            wb_destino = com_call(lambda: excel.Workbooks.Open(ruta_destino, UpdateLinks=False, ReadOnly=False))

            ws_origen_ejec = wb_ejec.Worksheets(1)
            ws_origen_marco = wb_marco.Worksheets(1)

            # Asegurarse de que las hojas destino existen
            try:
                ws_destino_ejec = wb_destino.Worksheets("Estado Cierre Ejec")
                ws_destino_marco = wb_destino.Worksheets("Estado Cierre Marco")
            except Exception:
                hojas = [ws.Name for ws in wb_destino.Worksheets]
                print("✗ No se encontraron las hojas destino esperadas en el libro destino. Hojas disponibles:")
                for h in hojas:
                    print(f"  - {h}")
                return

            print("Limpiando contenido...")

            com_call(lambda: ws_destino_ejec.Range("B1:M1000").Clear())
            com_call(lambda: ws_destino_marco.Range("B1:M1000").Clear())

            # NOTA IMPORTANTE: se reemplazó Copy()/PasteSpecial() por una
            # transferencia directa de valores (Range.Value = Range.Value).
            # Copy()/PasteSpecial usan el portapapeles de Windows, que es un
            # recurso ÚNICO y GLOBAL del sistema operativo. Como este script
            # corre EN PARALELO con "copiar_pegar_form_ejecu.py" (otra
            # instancia de Excel automatizada al mismo tiempo), ambos procesos
            # pueden pelear por el portapapeles y corromper la operación,
            # lo que produce errores intermitentes y poco descriptivos como
            # "(-2147352573, 'No se ha encontrado el miembro.', None, None)".
            # Al transferir por .Value se evita el portapapeles por completo
            # y además es más rápido.

            print("Copiando Estado Cierre Ejec...")

            def _copiar_valores(ws_origen, ws_destino, rango="B1:M1000"):
                valores = ws_origen.Range(rango).Value
                ws_destino.Range(rango).Value = valores

            com_call(lambda: _copiar_valores(ws_origen_ejec, ws_destino_ejec))

            print("Copiando Estado Cierre Marco...")

            com_call(lambda: _copiar_valores(ws_origen_marco, ws_destino_marco))

            com_call(lambda: wb_destino.Save())

            print("✓ Archivo guardado correctamente")

        except Exception as e:

            print(f"\n✗ ERROR: {e}")

        finally:
            # Soltar referencias a Worksheets/Ranges ANTES de cerrar libros
            # y salir de Excel. Si el garbage collector de Python no libera
            # estos objetos COM (RCW) antes del Quit(), EXCEL.EXE puede
            # quedar como proceso zombi en segundo plano aunque Quit() se
            # haya llamado "correctamente".
            ws_origen_ejec = None
            ws_origen_marco = None
            ws_destino_ejec = None
            ws_destino_marco = None
            gc.collect()

            try:
                if wb_ejec:
                    com_call(lambda: wb_ejec.Close(False))
            except Exception:
                pass

            try:
                if wb_marco:
                    com_call(lambda: wb_marco.Close(False))
            except Exception:
                pass

            try:
                if wb_destino:
                    com_call(lambda: wb_destino.Close(True))
            except Exception:
                pass

            try:
                if excel:
                    com_call(lambda: excel.Quit())
            except Exception:
                pass

            wb_ejec = None
            wb_marco = None
            wb_destino = None
            excel = None

            gc.collect()

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

            print(f"Tiempo total: {round(time.time() - inicio, 2)} segundos")
    
    # --- COORDINACIÓN ASÍNCRONA ---
    loop = asyncio.get_running_loop()
    
    # max_workers=1 para garantizar exclusión mutua en operaciones COM
    with ThreadPoolExecutor(max_workers=1) as executor:
        await loop.run_in_executor(executor, _cierre_sync)