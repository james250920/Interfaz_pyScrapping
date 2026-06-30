import os
import asyncio
from concurrent.futures import ThreadPoolExecutor
import win32com.client
import pythoncom

# Convertida a función principal asíncrona
async def limpiar_hojas_excel(ruta_archivo):
    
    # Subfunción síncrona interna para aislar el entorno COM en el hilo secundario
    def _limpiar_sync():
        archivo_completo = os.path.join(ruta_archivo, "Base FONAFE WEB al mes.xlsm")
        
        # CRÍTICO: Inicializa el entorno COM para este hilo específico del pool
        pythoncom.CoInitialize()
        
        excel = None
        wb = None

        try:
            print("Iniciando Excel en segundo plano mediante interfaz COM...")
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False    
            excel.ScreenUpdating = False   
            
            print(f"Abriendo archivo: {os.path.basename(archivo_completo)}...")
            wb = excel.Workbooks.Open(archivo_completo)

            hojas_a_limpiar = ["FBK-Alineado PRE", "FBK-Alineado FLU"]
            rango_borrar = "A5:AZ5000"

            for nombre_hoja in hojas_a_limpiar:
                try:
                    ws = wb.Worksheets(nombre_hoja)
                    ws.Range(rango_borrar).ClearContents()
                    print(f"  ✓ Rango {rango_borrar} limpiado en la hoja: {nombre_hoja}")
                except Exception as e_hoja:
                    print(f"  ⚠ Advertencia: No se pudo limpiar la hoja '{nombre_hoja}'. Error: {e_hoja}")

            wb.Close(SaveChanges=True)
            print("✓ Archivo guardado y cerrado correctamente.")

        except Exception as e:
            print(f"✗ Error crítico durante la ejecución en Excel: {e}")
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
                    
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception as e_quit:
                    print(f"Error al intentar cerrar la instancia de Excel: {e_quit}")
                    
            # CRÍTICO: Libera los recursos del entorno COM antes de cerrar el hilo
            pythoncom.CoUninitialize()

    # --- COORDINACIÓN ASÍNCRONA ---
    loop = asyncio.get_running_loop()
    
    # max_workers=1 asegura que no compartas memoria de la API de Excel de escritorio
    with ThreadPoolExecutor(max_workers=1) as executor:
        await loop.run_in_executor(executor, _limpiar_sync)
        
    print("Proceso de limpieza finalizado de forma asíncrona.")


