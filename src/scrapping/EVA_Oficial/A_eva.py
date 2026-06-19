import os
import time
import psutil
import pythoncom
import win32com.client as win32

def Copiar_Pegar(ruta_principal, anio, mes):
    RUTA_DESTINO = rf"{ruta_principal}\Eva 2026 - Información al mes.xlsm"
    RUTA_ORIGEN = rf"{ruta_principal}\Base FONAFE WEB al mes.xlsm"


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

                    print(f"Cerrando Excel PID: {proc.info['pid']}")

                    proc.kill()

            except:
                pass

    inicio = time.time()

    excel = None
    wb_origen = None
    wb_destino = None
    excel_pid = None

    try:

        cerrar_excels()

        time.sleep(3)

        pythoncom.CoInitialize()

        if not os.path.exists(RUTA_ORIGEN):
            raise FileNotFoundError(f"No existe:\n{RUTA_ORIGEN}")

        if not os.path.exists(RUTA_DESTINO):
            raise FileNotFoundError(f"No existe:\n{RUTA_DESTINO}")

        print("✓ Archivos encontrados")

        excel = win32.DispatchEx("Excel.Application")

        excel_pid = excel.Hwnd

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.Interactive = False

        time.sleep(2)

        print("Abriendo archivo origen...")

        wb_origen = excel.Workbooks.Open(
            RUTA_ORIGEN,
            UpdateLinks=0,
            ReadOnly=True
        )

        time.sleep(2)

        print("Abriendo archivo destino...")

        wb_destino = excel.Workbooks.Open(
            RUTA_DESTINO,
            UpdateLinks=0
        )

        time.sleep(2)

        hoja_bpre = wb_destino.Worksheets("BPRE")

        hoja_bpre.Range("A2").Value = anio
        hoja_bpre.Range("B2").Value = mes

        print("✓ BPRE actualizado")

        for hoja_orig, rango_orig, hoja_dest, rango_dest in OPERACIONES:

            try:

                ws_origen = wb_origen.Worksheets(hoja_orig)
                ws_destino = wb_destino.Worksheets(hoja_dest)

                datos = ws_origen.Range(rango_orig).Value

                ws_destino.Range(rango_dest).Value = datos

                print(f"✓ {hoja_orig} → {hoja_dest}")

            except Exception as e:

                print(f"✗ Error en {hoja_orig}: {e}")

        print("Guardando archivo...")

        wb_destino.Save()

        print("\n✓ Archivo guardado correctamente")

    except Exception as e:

        print(f"\n✗ ERROR GENERAL:\n{e}")

    finally:

        try:
            if wb_origen:
                wb_origen.Close(False)
        except:
            pass

        try:
            if wb_destino:
                wb_destino.Close(True)
        except:
            pass

        try:
            if excel:
                excel.Quit()
        except:
            pass

        cerrar_excels()

        pythoncom.CoUninitialize()

        print(f"\nTiempo total: {round(time.time() - inicio, 2)} segundos")