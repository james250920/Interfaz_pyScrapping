import os
import time
import gc
import pythoncom
import win32com.client as win32
from win32com.client import constants
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def copiar_pegar_form_ejecu(ruta_archivo):
    destino_file = f"{ruta_archivo}\\Validación Data Marco y Ejecución al mes.xlsm"
    
    if not os.path.exists(destino_file):
        print(f"ERROR: No existe el archivo destino: {destino_file}")
        return False

    CONFIG = [
        {
            "nombre": "MARCO",
            "carpeta": f"{ruta_archivo}\\MARCO",
            "hoja_origen": 2,
            "columnas": "AB",
            "excluir": ("gastos_capital",),
            "mapa": {
                "Presu_Ingresos_Egresos_Formulacion": "Sistema-Form. PRE",
                "Flujo_de_Caja_Formulacion": "Sistema-Form. FC",
                "Estado_de_Situacion_Financiera_Formulacion": "Sistema-Form. ESF",
                "Estado_de_Resultados_Integrales_Formulacion": "Sistema-Form. ERI",
            }
        },
        {
            "nombre": "EJECUCION",
            "carpeta": f"{ruta_archivo}\\EJECUCION",
            "hoja_origen": 1,
            "columnas": "U",
            "excluir": ("gastos_capital", "depositos_colocaciones",),
            "mapa": {
                "Presu_Ingresos_Egresos_Ejecucion": "Sistema-Ejec. PRE",
                "Flujo_de_Caja_Ejecucion": "Sistema-Ejec. FC",
                "Balance_General_Ejecucion": "Sistema-Ejec. ESF",
                "Estado_Ganancias_Perdidas_Ejecucion": "Sistema-Ejec. ERI",
            }
        }
    ]

    inicio = time.time()
    
    excel = None
    wb_destino = None
    
    try:
        print("Inicializando COM...")
        pythoncom.CoInitialize()
        
        print("Iniciando Excel...")
        # Usar una instancia nueva evita reusar un Excel/COM de una ejecución anterior
        excel = win32.DispatchEx("Excel.Application")
        
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.DisplayStatusBar = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3
        
        print("\nAbriendo archivo destino...")
        wb_destino = excel.Workbooks.Open(
            destino_file,
            UpdateLinks=False,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True
        )

        # Darle un instante a Excel para estabilizar la colección de hojas
        time.sleep(0.5)
        pythoncom.PumpWaitingMessages()
        
        time.sleep(0.5)
        pythoncom.PumpWaitingMessages()
        
        cache_destino = {}
        
        for config in CONFIG:
            for hoja in config["mapa"].values():
                if hoja not in cache_destino:
                    try:
                        cache_destino[hoja] = wb_destino.Worksheets(hoja)
                        print(f"✓ Hoja destino cacheada: {hoja}")
                    except Exception as e:
                        print(f"✗ Error accediendo a hoja {hoja}: {e}")
                        for ws in wb_destino.Worksheets:
                            if ws.Name == hoja:
                                cache_destino[hoja] = ws
                                print(f"  Recuperada hoja: {hoja}")
                                break
        
        for config in CONFIG:
            print(f"\n{'=' * 60}")
            print(f"PROCESANDO: {config['nombre']}")
            print(f"{'=' * 60}")
            
            carpeta = config["carpeta"]
            mapa_destino = config["mapa"]
            hoja_origen_num = config["hoja_origen"]
            columnas = config["columnas"]
            excluir = tuple(x.lower() for x in config["excluir"])
            
            if not os.path.exists(carpeta):
                print(f"¡ATENCIÓN! Carpeta no encontrada: {carpeta}")
                continue
            
            archivos = [
                os.path.join(carpeta, f)
                for f in os.listdir(carpeta)
                if (f.lower().endswith((".xlsx", ".xlsm", ".xls"))
                    and not f.lower().startswith(excluir))
            ]
            
            if not archivos:
                print(f"No se encontraron archivos en: {carpeta}")
                continue
            
            for archivo in archivos:
                wb_origen = None
                
                try:
                    nombre_archivo = os.path.splitext(os.path.basename(archivo))[0]
                    print(f"\n  Procesando: {nombre_archivo}")
                    
                    hoja_destino_nombre = mapa_destino.get(nombre_archivo)
                    if not hoja_destino_nombre:
                        print(f"    Sin hoja destino definida para: {nombre_archivo}")
                        continue
                    
                    if hoja_destino_nombre not in cache_destino:
                        print(f"    ERROR: Hoja destino '{hoja_destino_nombre}' no está en cache")
                        continue

                    wb_origen = excel.Workbooks.Open(
                        archivo,
                        UpdateLinks=False,
                        ReadOnly=True,
                        IgnoreReadOnlyRecommended=True
                    )
                    
                    ws_origen = wb_origen.Worksheets(hoja_origen_num)
                    print(f"    Hoja origen: {ws_origen.Name}")
                    
                    ws_destino = cache_destino[hoja_destino_nombre]
                    
                    ultima_fila = ws_origen.Cells(ws_origen.Rows.Count, 1).End(constants.xlUp).Row
                    
                    if ultima_fila < 2:
                        print("    No hay datos para copiar")
                        wb_origen.Close(False)
                        wb_origen = None
                        continue
                    
                    rango_origen = f"A2:{columnas}{ultima_fila}"
                    rango_destino = f"A2:{columnas}{ultima_fila}"
                    
                    data = ws_origen.Range(rango_origen).Value
                    ws_destino.Range(rango_destino).ClearContents()
                    ws_destino.Range(rango_destino).Value = data
                    
                    print(f"    ✓ OK -> {hoja_destino_nombre} (Filas: {ultima_fila-1})")

                    pythoncom.PumpWaitingMessages()
                    
                except Exception as e:
                    print(f"    ✗ ERROR en {nombre_archivo}: {type(e).__name__} - {e}")
                    
                finally:
                    if wb_origen:
                        try:
                            wb_origen.Close(False)
                            wb_origen = None
                        except:
                            pass
        
        print("\nGuardando archivo destino...")
        wb_destino.Save()
        print("✓ Guardado exitoso")
        
        return True
        
    except Exception as e:
        print(f"\nERROR GENERAL: {type(e).__name__} - {e}")
        return False
        
    finally:
        print("\n=== LIMPIANDO RECURSOS ===")
        
        if wb_destino:
            try:
                wb_destino.Close(SaveChanges=True)
                print("✓ Archivo destino cerrado")
            except Exception as e:
                print(f"✗ Error cerrando destino: {e}")
        
        if excel:
            try:
                excel.Quit()
                print("✓ Excel cerrado")
            except Exception as e:
                print(f"✗ Error cerrando Excel: {e}")
        
        wb_destino = None
        excel = None
        
        try:
            gc.collect()
            time.sleep(0.5)
            pythoncom.CoUninitialize()
            print("✓ COM desinicializado")
        except:
            pass
        
        fin = time.time()
        print(f"\n✓ Tiempo total: {round(fin - inicio, 2)} segundos")