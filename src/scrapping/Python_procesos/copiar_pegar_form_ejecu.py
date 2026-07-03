import os
import time
import gc
import asyncio
from concurrent.futures import ThreadPoolExecutor
import pythoncom
import win32com.client as win32
import pywintypes
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

XL_UP = -4162  # constante nativa xlUp

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


# Función principal asíncrona robusta
async def copiar_pegar_form_ejecu(ruta_archivo):
    
    def _copiar_sync():
        destino_file = os.path.join(ruta_archivo, "Validación Data Marco y Ejecución al mes.xlsm")
        
        if not os.path.exists(destino_file):
            print(f"ERROR: No existe el archivo destino: {destino_file}")
            return False

        CONFIG = [
            {
                "nombre": "MARCO",
                "carpeta": os.path.join(ruta_archivo, "MARCO"),
                "hoja_origen": "ULTIMA_FORMULACION",  # <--- Detección dinámica inteligente
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
                "carpeta": os.path.join(ruta_archivo, "EJECUCION"),
                "hoja_origen": 1,
                "columnas": "U",
                "excluir": ("gastos_capital", "depositos_colocaciones"),
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
        cache_destino = {}
        
        try:
            print("Inicializando COM...")
            pythoncom.CoInitialize()
            
            print("Iniciando instancia de Excel...")
            excel = com_call(lambda: win32.DispatchEx("Excel.Application"))
            
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.DisplayStatusBar = False
            excel.AskToUpdateLinks = False
            excel.AutomationSecurity = 3
            
            print("\nAbriendo archivo destino...")
            wb_destino = com_call(lambda: excel.Workbooks.Open(
                destino_file,
                UpdateLinks=False,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True
            ))

            time.sleep(0.5)
            pythoncom.PumpWaitingMessages()
            
            # Almacenar hojas destino de manera segura en caché
            for config in CONFIG:
                for hoja in config["mapa"].values():
                    if hoja not in cache_destino:
                        try:
                            cache_destino[hoja] = com_call(lambda: wb_destino.Worksheets(hoja))
                            print(f"✓ Hoja destino cacheada: {hoja}")
                        except Exception as e:
                            print(f"✗ Error accediendo a hoja {hoja}: {e}")
            
            # Procesar las carpetas secuencialmente
            for config in CONFIG:
                print(f"\n{'=' * 60}")
                print(f"PROCESANDO: {config['nombre']}")
                print(f"{'=' * 60}")
                
                carpeta = config["carpeta"]
                mapa_destino = config["mapa"]
                hoja_origen_cfg = config["hoja_origen"]
                columnas = config["columnas"]
                excluir = tuple(x.lower() for x in config["excluir"])
                
                if not os.path.exists(carpeta):
                    print(f"¡ATENCIÓN! Carpeta no encontrada: {carpeta}")
                    continue
                
                archivos = [
                    os.path.join(carpeta, f)
                    for f in os.listdir(carpeta)
                    if (f.lower().endswith((".xlsx", ".xlsm", ".xls"))
                        and not f.lower().startswith(excluir)
                        and not f.startswith("~$"))
                ]
                
                if not archivos:
                    print(f"No se encontraron archivos válidos en: {carpeta}")
                    continue
                
                for archivo in archivos:
                    wb_origen = None
                    try:
                        nombre_archivo = os.path.splitext(os.path.basename(archivo))[0]
                        print(f"\n  Procesando: {nombre_archivo}")
                        
                        hoja_destino_nombre = mapa_destino.get(nombre_archivo)
                        if not hoja_destino_nombre or hoja_destino_nombre not in cache_destino:
                            print(f"    Saltando o sin mapa destino válido para: {nombre_archivo}")
                            continue
                        
                        wb_origen = com_call(lambda: excel.Workbooks.Open(
                            archivo,
                            UpdateLinks=False,
                            ReadOnly=True,
                            IgnoreReadOnlyRecommended=True
                        ))
                        
                        # --- DETECCIÓN INTELIGENTE DE HOJA ---
                        if hoja_origen_cfg == "ULTIMA_FORMULACION":
                            # Capturar nombres de hojas que cumplan con la estructura de openpyxl
                            nombres_hojas = [ws.Name for ws in wb_origen.Worksheets]
                            hojas_form = [h for h in nombres_hojas if h.startswith("Formulacion_")]
                            
                            if hojas_form:
                                # Ordenar por el entero final (ej: "Formulacion_3" -> 3) para obtener la última
                                hoja_final = sorted(hojas_form, key=lambda x: int(x.split('_')[1]))[-1]
                                ws_origen = com_call(lambda: wb_origen.Worksheets(hoja_final))
                            else:
                                ws_origen = com_call(lambda: wb_origen.Worksheets(1))
                        else:
                            ws_origen = com_call(lambda: wb_origen.Worksheets(hoja_origen_cfg))
                        
                        print(f"    Hoja origen detectada: '{ws_origen.Name}'")
                        ws_destino = cache_destino[hoja_destino_nombre]
                        
                        # Calcular última fila usando com_call de manera segura
                        def _get_last_row():
                            rcount = ws_origen.Rows.Count
                            return ws_origen.Cells(rcount, 1).End(XL_UP).Row
                        
                        ultima_fila = com_call(_get_last_row)
                        
                        if ultima_fila < 2:
                            print("    No existen datos para copiar (fila menor a 2)")
                            com_call(lambda: wb_origen.Close(False))
                            wb_origen = None
                            continue
                        
                        rango_origen = f"A2:{columnas}{ultima_fila}"
                        rango_destino = f"A2:{columnas}{ultima_fila}"
                        
                        # Lectura limpia directa por valores (sin portapapeles)
                        data = com_call(lambda: ws_origen.Range(rango_origen).Value)
                        
                        # Limpiar toda la data antigua para evitar filas remanentes si la data nueva es menor
                        rango_limpieza = f"A2:{columnas}50000"
                        com_call(lambda: ws_destino.Range(rango_limpieza).ClearContents())
                        
                        def _set_value():
                            ws_destino.Range(rango_destino).Value = data
                        com_call(_set_value)
                        
                        print(f"    ✓ OK -> {hoja_destino_nombre} (Filas transferidas: {ultima_fila-1})")
                        pythoncom.PumpWaitingMessages()
                        
                    except Exception as e:
                        print(f"    ✗ ERROR procesando {nombre_archivo}: {type(e).__name__} - {e}")
                    finally:
                        if wb_origen:
                            try:
                                com_call(lambda: wb_origen.Close(False))
                            except:
                                pass
                            wb_origen = None
            
            print("\nGuardando archivo destino...")
            com_call(lambda: wb_destino.Save())
            print("✓ Guardado general exitoso")
            return True
            
        except Exception as e:
            print(f"\nERROR CRÍTICO GENERAL: {type(e).__name__} - {e}")
            return False
            
        finally:
            print("\n=== LIMPIANDO RECURSOS EN COPIADO ===")
            try:
                cache_destino.clear()
            except:
                pass
            
            ws_origen = None
            ws_destino = None
            gc.collect()

            if wb_destino:
                try:
                    com_call(lambda: wb_destino.Close(False))
                    print("✓ Archivo destino cerrado")
                except Exception as e:
                    print(f"✗ Error cerrando libro destino: {e}")
            
            if excel:
                try:
                    com_call(lambda: excel.Quit())
                    print("✓ Proceso Excel cerrado exitosamente")
                except Exception as e:
                    print(f"✗ Error matando instancia de Excel: {e}")
            
            wb_destino = None
            excel = None
            gc.collect()
            time.sleep(0.5)
            
            try:
                pythoncom.CoUninitialize()
                print("✓ Sistema COM liberado")
            except:
                pass
            
            print(f"Tiempo total de copiado: {round(time.time() - inicio, 2)} segundos")
    
    # --- COORDINACIÓN ASÍNCRONA ---
    loop = asyncio.get_running_loop()
    with ThreadPoolExecutor(max_workers=1) as executor:
        return await loop.run_in_executor(executor, _copiar_sync)