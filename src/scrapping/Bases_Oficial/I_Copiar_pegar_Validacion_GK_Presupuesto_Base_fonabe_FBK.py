import time
import gc
import os
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


def buscar_hoja(workbook, nombre):
    for ws in workbook.Worksheets:
        if ws.Name.strip() == nombre.strip():
            return ws
    return None

def copiar_pegar_validacion_presupuesto(ruta_principal):
    RUTA_VALIDACION = rf"{ruta_principal}\VALIDACION GASTO CAPITAL\Validación_Gastos_Capital_Presupuesto.xlsx"
    RUTA_DESTINO    = rf"{ruta_principal}\Base FONAFE WEB al mes.xlsm"

    OPERACIONES = [
        ("Validacion_Comparativa", "A2:R5000",   "FBK-Alineado PRE", "A5:R5000"),
        ("Validacion_Comparativa", "AA2:AL5000", "FBK-Alineado PRE", "S5:AD5000"),
    ]

    wb_leer   = openpyxl.load_workbook(RUTA_VALIDACION, data_only=True)
    hoja_leer = wb_leer.worksheets[0]
    valor_ap1 = hoja_leer["AP1"].value or 0
    valor_aq1 = hoja_leer["AQ1"].value or 0
    wb_leer.close()

    if valor_ap1 + valor_aq1 != 0:
        print(f"✗ No cumple. La suma de AP1+AQ1 = {valor_ap1 + valor_aq1}. Actualizar manualmente.")
    else:
        print(f"✓ Validación cumplida: AP1({valor_ap1}) + AQ1({valor_aq1}) == 0\n")

        inicio = time.time()
        excel = wb_FC = wb_destino = None
        pid_excel = None

        try:
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

            excel.Calculation         = -4135
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

            excel.Calculation = -4105    
            wb_destino.Save()

            if errores == 0:
                print("\n✓ Guardado exitoso — sin errores de mapeo.")
            else:
                print(f"\n⚠ Guardado completado con {errores} operación(es) fallida(s).")

        except Exception as e:
            print(f"\n✗ ERROR INESPERADO DURANTE EL PROCESO: {e}")

        finally:
            for wb in (wb_FC, wb_destino):
                try:
                    if wb:
                        wb.Close(False)
                except Exception:
                    pass
            try:
                if excel:
                    excel.Quit()
            except Exception:
                pass

            excel = wb_FC = wb_destino = None
            gc.collect()
            pythoncom.CoUninitialize()

            time.sleep(1)
            forzar_cierre_proceso(pid_excel)

            print(f"\nTiempo total de ejecución: {round(time.time() - inicio, 2)} segundos")