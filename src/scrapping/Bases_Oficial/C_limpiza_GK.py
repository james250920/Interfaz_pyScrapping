import win32com.client
import pythoncom

def limpiar_hojas_excel(ruta_archivo):
    ruta_archivo = rf"{ruta_archivo}\Base FONAFE WEB al mes.xlsm"
    pythoncom.CoInitialize()
    
    excel = None
    wb = None

    try:
        print("Iniciando Excel en segundo plano...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False    
        excel.ScreenUpdating = False   
        
        
        print("Abriendo archivo...")
        wb = excel.Workbooks.Open(ruta_archivo)

        hojas_a_limpiar = ["FBK-Alineado PRE", "FBK-Alineado FLU"]
        rango_borrar = "A5:AZ5000"

        for nombre_hoja in hojas_a_limpiar:
            try:
                ws = wb.Worksheets(nombre_hoja)
                ws.Range(rango_borrar).ClearContents()
                print(f"Rango limpiado en la hoja: {nombre_hoja}")
            except Exception as e_hoja:
                print(f"Advertencia: No se encontró o no se pudo limpiar la hoja '{nombre_hoja}'. Error: {e_hoja}")

        wb.Close(SaveChanges=True)
        print(" Archivo guardado y cerrado correctamente.")

    except Exception as e:
        print(f" Error crítico durante la ejecución: {e}")
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
                print(f"Error al intentar cerrar Excel: {e_quit}")
                
        pythoncom.CoUninitialize()
        print("Proceso finalizado.")
