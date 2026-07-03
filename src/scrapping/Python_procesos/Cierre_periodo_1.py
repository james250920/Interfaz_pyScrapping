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

# Función principal asíncrona corregida
async def cierre_periodo(ruta_principal):
    
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

            print("Inicializando COM...")
            pythoncom.CoInitialize()

            print("Iniciando instancia de Excel...")
            excel = com_call(lambda: win32.DispatchEx("Excel.Application"))

            try:
                excel.Visible = False
                excel.DisplayAlerts = False
                excel.ScreenUpdating = False
                excel.EnableEvents = False
            except Exception:
                pass

            print("Abriendo archivos...")

            print("Abriendo Estado_de_Cierre_del_Periodo_Ejecucion.xlsx...")
            wb_ejec = com_call(lambda: excel.Workbooks.Open(
                ruta_ejec,
                UpdateLinks=False,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True
            ))
            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)

            print("Abriendo Estado_de_Cierre_del_Periodo_Formulacion.xlsx...")
            wb_marco = com_call(lambda: excel.Workbooks.Open(
                ruta_marco,
                UpdateLinks=False,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True
            ))
            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)

            print("Abriendo Estado de Cierre al mes.xlsm...")
            wb_destino = com_call(lambda: excel.Workbooks.Open(
                ruta_destino,
                UpdateLinks=False,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True
            ))
            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)

            # Obtener hojas de origen de forma segura
            ws_origen_ejec = com_call(lambda: wb_ejec.Worksheets(1))
            ws_origen_marco = com_call(lambda: wb_marco.Worksheets(1))

            # Asegurarse de que las hojas destino existen
            try:
                ws_destino_ejec = com_call(lambda: wb_destino.Worksheets("Estado Cierre Ejec"))
                ws_destino_marco = com_call(lambda: wb_destino.Worksheets("Estado Cierre Marco"))
            except Exception:
                hojas = [ws.Name for ws in wb_destino.Worksheets]
                print("✗ No se encontraron las hojas destino esperadas en el libro destino. Hojas disponibles:")
                for h in hojas:
                    print(f"  - {h}")
                return

            print("Limpiando contenido...")
            com_call(lambda: ws_destino_ejec.Range("B1:M1000").ClearContents())
            com_call(lambda: ws_destino_marco.Range("B1:M1000").ClearContents())

            # --- SUBFUNCIÓN DE COPIADO BLINDADA ---
            def _copiar_valores(ws_origen, ws_destino, rango="B1:M1000"):
                # Se envuelven las operaciones nativas COM individuales en el com_call
                valores = com_call(lambda: ws_origen.Range(rango).Value)
                
                def _asignar():
                    ws_destino.Range(rango).Value = valores
                com_call(_asignar)

            print("Copiando Estado Cierre Ejec...")
            # Llamada directa (sin lambda anidado corrupto en com_call)
            _copiar_valores(ws_origen_ejec, ws_destino_ejec)

            print("Copiando Estado Cierre Marco...")
            _copiar_valores(ws_origen_marco, ws_destino_marco)

            print("Guardando libro destino...")
            com_call(lambda: wb_destino.Save())
            print("✓ Archivo guardado correctamente")

        except Exception as e:
            print(f"\n✗ ERROR en Cierre Periodo: {type(e).__name__} - {e}")

        finally:
            print("=== LIMPIANDO RECURSOS EN CIERRE ===")
            # Soltar referencias COM explícitamente para evitar bloqueos
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
            time.sleep(0.5)

            try:
                pythoncom.CoUninitialize()
                print("✓ COM desinicializado en Cierre")
            except Exception:
                pass

            print(f"Tiempo total: {round(time.time() - inicio, 2)} segundos")
    
    # --- COORDINACIÓN ASÍNCRONA ---
    loop = asyncio.get_running_loop()
    with ThreadPoolExecutor(max_workers=1) as executor:
        await loop.run_in_executor(executor, _cierre_sync)