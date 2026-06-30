import os
import time
import gc
import asyncio
from concurrent.futures import ThreadPoolExecutor
import tempfile
import shutil
import pythoncom
import pywintypes
import win32com.client as win32
import win32process
import win32api
import openpyxl

RPC_E_CALL_REJECTED         = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846

# Se mantiene la lógica exacta de reintentos síncronos de COM
def com_call(fn, reintentos=12, pausa=2.5):
    for intento in range(1, reintentos + 1):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER):
                print(f"    ⏳ Excel ocupado, reintento {intento}/{reintentos} en {pausa}s...")
                time.sleep(pausa)
            else:
                raise
    raise RuntimeError(f"Excel rechazó la llamada tras {reintentos} reintentos.")


def obtener_pid_excel(excel_app):
    try:
        hwnd = excel_app.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid
    except Exception:
        return None


def forzar_cierre_proceso(pid):
    if pid is None:
        return
    try:
        handle = win32api.OpenProcess(1, False, pid)
        win32api.TerminateProcess(handle, 0)
        win32api.CloseHandle(handle)
        print(f"    🔪 Proceso Excel (PID {pid}) terminado forzosamente.")
    except Exception:
        pass  


def limpiar_cache_genpy():
    for patron in [
        os.path.join(tempfile.gettempdir(), "gen_py"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "gen_py"),
    ]:
        if os.path.exists(patron):
            shutil.rmtree(patron, ignore_errors=True)


# Convertida en la función principal asíncrona
async def copiar_pegar_validacion_presupuesto(ruta_principal):
    inicio = time.time()

    # Subfunción síncrona interna para aislar el proceso en el Pool de Hilos
    def _ejecutar_proceso_sync():
        RUTA_VALIDACION = os.path.join(ruta_principal, "VALIDACION GASTO CAPITAL", "Validación_Gastos_Capital_Presupuesto.xlsx")
        RUTA_DESTINO    = os.path.join(ruta_principal, "Base FONAFE WEB al mes.xlsm")

        OPERACIONES = [
            ("Validacion_Comparativa", "A2:R5000",   "FBK-Alineado PRE", "A5:R5000"),
            ("Validacion_Comparativa", "AA2:AL5000", "FBK-Alineado PRE", "S5:AD5000"),
        ]

        # 1. Lectura de validación rápida usando openpyxl
        wb_leer   = openpyxl.load_workbook(RUTA_VALIDACION, data_only=True)
        hoja_leer = wb_leer.worksheets[0]
        valor_ap1 = hoja_leer["AP1"].value or 0
        valor_aq1 = hoja_leer["AQ1"].value or 0
        wb_leer.close()

        if valor_ap1 + valor_aq1 != 0:
            print(f"✗ No cumple. La suma de AP1+AQ1 = {valor_ap1 + valor_aq1}. Actualizar manualmente.")
            return None # Retorna None para indicar que el proceso COM no se inició

        print(f"✓ Validación cumplida: AP1({valor_ap1}) + AQ1({valor_aq1}) == 0\n")

        # 2. Inicialización del Bloque COM seguro
        excel = wb_FC = wb_destino = None
        pid_excel = None

        def buscar_hoja(workbook, nombre):
            for ws in workbook.Worksheets:
                if ws.Name.strip() == nombre.strip():
                    return ws
            return None

        try:
            # CRÍTICO: Registrar COM en el hilo secundario
            pythoncom.CoInitialize()
            limpiar_cache_genpy()

            excel     = win32.DispatchEx("Excel.Application")
            pid_excel = obtener_pid_excel(excel)
            print(f"Proceso Excel iniciado en segundo plano (PID: {pid_excel})\n")

            excel.Visible         = False
            excel.DisplayAlerts   = False
            excel.ScreenUpdating  = False
            excel.EnableEvents    = False

            print("Abriendo archivos...")
            wb_FC      = excel.Workbooks.Open(RUTA_VALIDACION, UpdateLinks=False, ReadOnly=True)
            wb_destino = excel.Workbooks.Open(RUTA_DESTINO,    UpdateLinks=False, ReadOnly=False)

            excel.Calculation         = -4135 # xlCalculationManual
            excel.CalculateBeforeSave = False

            errores = 0
            for hoja_orig, rango_orig, hoja_dest, rango_dest in OPERACIONES:
                ws_origen  = buscar_hoja(wb_FC, hoja_orig)
                ws_destino = buscar_hoja(wb_destino, hoja_dest)

                if ws_origen and ws_destino:
                    valor = com_call(lambda o=ws_origen, r=rango_orig: o.Range(r).Value)
                    com_call(lambda d=ws_destino, r=rango_dest, v=valor: setattr(d.Range(r), "Value", v))
                    print(f"  ✓ {hoja_orig!r:25s} → {hoja_dest!r}")
                else:
                    faltante = []
                    if not ws_origen:  faltante.append(f"origen={hoja_orig!r}")
                    if not ws_destino: faltante.append(f"destino={hoja_dest!r}")
                    print(f"  ✗ Hoja no encontrada: {', '.join(faltante)}")
                    errores += 1

            excel.Calculation = -4105 # xlCalculationAutomatic
            wb_destino.Save()

            if errores == 0:
                print("\n✓ Guardado exitoso — sin errores de mapeo.")
            else:
                print(f"\n⚠ Guardado completado con {errores} operación(es) fallida(s).")

        except Exception as e:
            print(f"\n✗ ERROR INESPERADO DURANTE EL PROCESO COM: {e}")

        finally:
            # Garantizar la liberación ordenada de recursos en el hilo
            for wb in (wb_FC, wb_destino):
                try:
                    if wb: wb.Close(False)
                except Exception: pass
            try:
                if excel: excel.Quit()
            except Exception: pass

            excel = wb_FC = wb_destino = None
            gc.collect()
            pythoncom.CoUninitialize()

        return pid_excel

    # --- COORDINACIÓN ASÍNCRONA ---
    loop = asyncio.get_running_loop()
    
    # max_workers=1 para aislar la instancia de automatización de Excel de escritorio
    with ThreadPoolExecutor(max_workers=1) as executor:
        pid_generado = await loop.run_in_executor(executor, _ejecutar_proceso_sync)

    # Si se pasó la validación y se creó el proceso, esperamos y cerramos asíncronamente
    if pid_generado is not None:
        await asyncio.sleep(1)
        forzar_cierre_proceso(pid_generado)

    print(f"\nTiempo total de ejecución: {round(time.time() - inicio, 2)} segundos")

