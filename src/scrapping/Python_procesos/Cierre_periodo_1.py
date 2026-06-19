import time
import gc
import os
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

def cierre_periodo(ruta_principal):
    ruta_destino = rf"{ruta_principal}\\Estado de Cierre al mes.xlsm"
    ruta_ejec = rf"{ruta_principal}\\CIERRE\\Estado_de_Cierre_del_Periodo_Ejecucion.xlsx"
    ruta_marco = rf"{ruta_principal}\\CIERRE\\Estado_de_Cierre_del_Periodo_Formulacion.xlsx"
    
    inicio = time.time()
    excel = None
    wb_ejec = None
    wb_marco = None
    wb_destino = None

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

        excel = com_call(lambda: win32.gencache.EnsureDispatch("Excel.Application"))

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

        xlPasteAll = -4104

        print("Copiando Estado Cierre Ejec...")

        com_call(lambda: ws_origen_ejec.Range("B1:M1000").Copy())

        com_call(lambda: ws_destino_ejec.Range("B1").PasteSpecial(xlPasteAll))

        print("Copiando Estado Cierre Marco...")

        com_call(lambda: ws_origen_marco.Range("B1:M1000").Copy())

        com_call(lambda: ws_destino_marco.Range("B1").PasteSpecial(xlPasteAll))

        try:
            excel.CutCopyMode = False
        except Exception:
            pass

        com_call(lambda: wb_destino.Save())

        print("✓ Archivo guardado correctamente")

    except Exception as e:

        print(f"\n✗ ERROR: {e}")

    finally:
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