import os
import time
import gc
import tempfile
import shutil
import asyncio
from concurrent.futures import ThreadPoolExecutor
import pythoncom
import pywintypes
import win32com.client as win32
import win32process
import openpyxl


RPC_E_CALL_REJECTED         = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
VBA_E_IGNORE                 = -2146777998  # 0x800AC472, error transitorio de "Excel ocupado"

HRESULTS_REINTENTABLES = (
    RPC_E_CALL_REJECTED,
    RPC_E_SERVERCALL_RETRYLATER,
    VBA_E_IGNORE,
)


# -----------------------------
# COM SAFE CALL
# -----------------------------
def com_call(fn, reintentos=12, pausa=2.5):
    for intento in range(1, reintentos + 1):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult in HRESULTS_REINTENTABLES:
                print(f"⏳ Excel ocupado (hresult={e.hresult}), reintento {intento}/{reintentos}...")
                time.sleep(pausa)
            else:
                raise
    raise RuntimeError("Excel rechazó la operación tras reintentos.")


# -----------------------------
# LIMPIEZA TEMPORAL
# -----------------------------
def limpiar_cache_genpy():
    rutas = [
        os.path.join(tempfile.gettempdir(), "gen_py"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "gen_py"),
    ]

    for r in rutas:
        if os.path.exists(r):
            shutil.rmtree(r, ignore_errors=True)


# -----------------------------
# FUNCION PRINCIPAL (ASYNC + THREADPOOL)
# -----------------------------
async def copiar_pegar_validacion_Flujo_Caja(ruta_principal):

    RUTA_VALIDACION = os.path.join(
        ruta_principal,
        "VALIDACION GASTO CAPITAL",
        "Validacion_Gastos_Capital_Flujo_Caja.xlsx"
    )

    RUTA_DESTINO = os.path.join(
        ruta_principal,
        "Base FONAFE WEB al mes.xlsm"
    )

    OPERACIONES = [
        ("Validacion_Comparativa", "A2:Q5000",  "FBK-Alineado FLU", "A5:Q5000"),
        ("Validacion_Comparativa", "X2:AI5000", "FBK-Alineado FLU", "R5:AC5000"),
    ]

    # -----------------------------
    # VALIDACIÓN PREVIA (openpyxl — fuera del hilo COM)
    # -----------------------------
    wb_check = openpyxl.load_workbook(RUTA_VALIDACION, data_only=True)
    ws_check = wb_check.worksheets[0]

    ap1 = ws_check["AP1"].value or 0
    aq1 = ws_check["AQ1"].value or 0
    wb_check.close()
    wb_check = None

    if ap1 + aq1 != 0:
        print(f"✗ Validación fallida: AP1+AQ1 = {ap1 + aq1}")
        return

    print(f"✓ Validación OK: AP1({ap1}) + AQ1({aq1}) = 0")

    # -----------------------------
    # SUBFUNCION SINCRONA (aislada en hilo propio)
    # -----------------------------
    def _ejecutar_com_sync():
        inicio = time.time()

        excel = None
        wb_validacion = None
        wb_destino = None

        def buscar_hoja(workbook, nombre):
            for ws in workbook.Worksheets:
                if ws.Name.strip() == nombre.strip():
                    return ws
            return None

        try:
            # CRÍTICO: Inicializar COM en este hilo
            pythoncom.CoInitialize()
            limpiar_cache_genpy()

            excel = win32.DispatchEx("Excel.Application")

            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            print("Excel iniciado (Flujo Caja)")
            print("Abriendo archivos...")

            wb_validacion = com_call(lambda: excel.Workbooks.Open(
                RUTA_VALIDACION,
                UpdateLinks=False,
                ReadOnly=True
            ))

            wb_destino = com_call(lambda: excel.Workbooks.Open(
                RUTA_DESTINO,
                UpdateLinks=False,
                ReadOnly=False
            ))

            excel.Calculation = -4135  # manual

            errores = 0

            for hoja_o, rango_o, hoja_d, rango_d in OPERACIONES:

                ws_o = buscar_hoja(wb_validacion, hoja_o)
                ws_d = buscar_hoja(wb_destino, hoja_d)

                if ws_o and ws_d:
                    valor = com_call(lambda o=ws_o, r=rango_o: o.Range(r).Value)
                    com_call(lambda d=ws_d, r=rango_d, v=valor: setattr(d.Range(r), "Value", v))

                    print(f"  ✓ {hoja_o} → {hoja_d}")
                else:
                    print(f"  ✗ Hoja faltante: {hoja_o} o {hoja_d}")
                    errores += 1

            excel.Calculation = -4105  # automático
            com_call(lambda: wb_destino.Save())

            if errores == 0:
                print("✓ Proceso Flujo Caja completado sin errores")
            else:
                print(f"⚠ Proceso Flujo Caja con {errores} errores")

        except Exception as e:
            print(f"✗ ERROR en Flujo Caja: {e}")
            # Se relanza para que el pipeline se detenga: si esto falla, la
            # Base FONAFE queda sin la validación de Flujo de Caja copiada.
            raise

        finally:
            # -----------------------------
            # CIERRE SEGURO Y ORDENADO
            # -----------------------------
            try:
                if wb_validacion:
                    wb_validacion.Close(False)
            except:
                pass

            try:
                if wb_destino:
                    wb_destino.Close(False)
            except:
                pass

            try:
                if excel:
                    excel.Quit()
            except:
                pass

            wb_validacion = None
            wb_destino = None
            excel = None
            
            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except:
                pass

            print(f"Tiempo total Flujo Caja: {round(time.time() - inicio, 2)}s")

    # Ejecutar en hilo aislado (max_workers=1 garantiza exclusión mutua COM)
    loop = asyncio.get_running_loop()
    with ThreadPoolExecutor(max_workers=1) as executor:
        await loop.run_in_executor(executor, _ejecutar_com_sync)