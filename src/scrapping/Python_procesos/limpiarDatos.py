import win32com.client as win32
import pythoncom
import time
import os
import subprocess

def limpiar_datos(ruta_archivo, mes):
    destino_file = f"{ruta_archivo}\\Validación Data Marco y Ejecución al mes.xlsm"
    inicio = time.time()
    
    excel = None
    wb = None
    
    try:
        print("Iniciando Excel...")
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
                    print(f"El archivo ya estaba abierto, cerrándolo...")
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
        print(f"Error durante el proceso: {e}")
        
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
        
        try:
            pythoncom.CoUninitialize()
            print("✓ COM liberado")
        except Exception as e:
            print(f"✗ Error liberando COM: {e}")
        
        try:
            time.sleep(1)
            subprocess.run(
                'taskkill /f /im excel.exe',
                shell=True,
                capture_output=True,
                text=True
            )
            print("✓ Procesos Excel forzados cerrados")
        except:
            pass
        
        fin = time.time()
        print(f"\n✓ Proceso terminado en {round(fin - inicio, 2)} segundos")
