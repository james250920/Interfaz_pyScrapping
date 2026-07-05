import os
import win32com.client
import pythoncom
import gc

def limpiar_hojas_excel(ruta_archivo):
    archivo_completo = os.path.join(ruta_archivo, "Base FONAFE WEB al mes.xlsm")
    
    # CRÍTICO: Inicializa el entorno COM
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    com_inicializado = True
    
    excel = None
    wb = None
    ruta_temporal = None

    try:
        import tempfile
        import shutil
        
        # Validar archivo origen
        if not os.path.exists(archivo_completo):
            raise FileNotFoundError(f"El archivo origen no existe: {archivo_completo}")

        fd, temp_name = tempfile.mkstemp(prefix="Base FONAFE WEB al mes_", suffix=".xlsm", dir=os.path.dirname(archivo_completo))
        os.close(fd)
        ruta_temporal = temp_name
        
        shutil.copy2(archivo_completo, ruta_temporal)

        print("Iniciando Excel en segundo plano mediante interfaz COM...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False    
        excel.ScreenUpdating = False   
        excel.EnableEvents = False
        
        print(f"Abriendo archivo temporal para limpieza...")
        wb = excel.Workbooks.Open(ruta_temporal, UpdateLinks=False)

        hojas_a_limpiar = ["FBK-Alineado PRE", "FBK-Alineado FLU"]
        rango_borrar = "A5:AZ5000"

        for nombre_hoja in hojas_a_limpiar:
            try:
                ws = wb.Worksheets(nombre_hoja)
                ws.Range(rango_borrar).ClearContents()
                print(f"  ✓ Rango {rango_borrar} limpiado en la hoja: {nombre_hoja}")
            except Exception as e_hoja:
                print(f"  ⚠ Advertencia: No se pudo limpiar la hoja '{nombre_hoja}'. Error: {e_hoja}")

        print("Guardando cambios en archivo temporal...")
        wb.Save()
        print("✓ Archivo guardado correctamente.")

    except Exception as e:
        print(f"✗ Error crítico durante la ejecución en Excel: {e}")
        raise
                
    finally:
        try:
            if excel is not None:
                excel.ScreenUpdating = True
                excel.EnableEvents = True
        except:
            pass
            
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
                
        if excel is not None:
            try:
                excel.Quit()
            except Exception as e_quit:
                print(f"Error al intentar cerrar la instancia de Excel: {e_quit}")
        wb = None
        excel = None
        gc.collect()
        
        if com_inicializado:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

        if ruta_temporal and os.path.exists(ruta_temporal):
            # Verificar si se guardó sin errores antes de reemplazar
            try:
                os.replace(ruta_temporal, archivo_completo)
                print(f"Archivo modificado y reemplazado atómicamente: {archivo_completo}")
            except Exception as e:
                print(f"⚠ Advertencia: no se pudo reemplazar el archivo original, el temporal quedó en: {ruta_temporal}")

    print("Proceso de limpieza finalizado.")


