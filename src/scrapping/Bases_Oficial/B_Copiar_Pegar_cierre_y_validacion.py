import os
import time
import gc
import asyncio
from concurrent.futures import ThreadPoolExecutor
import pythoncom
import pywintypes
import win32com.client as win32
import win32process
import win32api

RPC_E_CALL_REJECTED         = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
# 0x800AC472: "Excel no puede completar esta tarea con los recursos disponibles"
# Aparece de forma intermitente cuando el archivo aún conserva un handle/lock
# de la instancia de Excel anterior que se acaba de cerrar (condición de carrera).
VBA_E_IGNORE                 = -2146777998

HRESULTS_REINTENTABLES = (
    RPC_E_CALL_REJECTED,
    RPC_E_SERVERCALL_RETRYLATER,
    VBA_E_IGNORE,
)

# Se mantiene tu lógica original de reintentos síncronos, ampliando los
# códigos de error considerados transitorios/reintentables
def com_call(fn, reintentos=12, pausa=2.5):
    for intento in range(1, reintentos + 1):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult in HRESULTS_REINTENTABLES:
                print(f"    ⏳ Excel ocupado (hresult={e.hresult}), reintento {intento}/{reintentos} en {pausa}s...")
                time.sleep(pausa)
            else:
                raise
    raise RuntimeError(f"Excel rechazó la llamada tras {reintentos} reintentos.")


# Convertida a la función principal asíncrona
async def copiar_pegar_cierre_y_validacion(ruta_principal):
    inicio = time.time()

    # Subfunción interna síncrona que ejecutará el Executor de forma aislada
    def _ejecutar_proceso_com_sync():
        RUTA_CIERRE_1 = os.path.join(ruta_principal, "Estado de Cierre al mes.xlsm")
        RUTA_CIERRE_2 = os.path.join(ruta_principal, "Validación Data Marco y Ejecución al mes.xlsm")
        RUTA_DESTINO  = os.path.join(ruta_principal, "Base FONAFE WEB al mes.xlsm")

        OPS_CIERRE_1 = [
            ("Resumen Marco", "D5:R41",   "Sistema Cierre", "D52:R88"),
            ("Resumen Ejec",   "D5:R41",   "Sistema Cierre", "D9:R45"),
            ("Resumen Marco", "B1",       "Sistema Cierre", "D47"),
            ("Resumen Ejec",   "S5:AM41",  "Cierre FE",      "D7:X43"),
            ("Resumen Ejec",   "AN5:BB41", "Cierre FE",      "AB7:AP43"),
            ("Resumen Marco", "S5:X41",   "Cierre FE",      "D51:I87"),
            ("Resumen Marco", "Y5:AD41",  "Cierre FE",      "V51:AA87"),
            ("Resumen Marco", "AE5:AG41", "Cierre FE",      "AH51:AJ87"),
            ("Resumen Marco", "AH5:AJ41", "Cierre FE",      "AN51:AP87"),
        ]

        OPS_CIERRE_2 = [
            ("Presupuesto",   "G4:AD3999", "PRE ", "C4:Z3999"),
            ("Flujo de Caja", "G4:AD2343", "FLU",  "C4:Z2343"),
            ("ESF",           "G4:AD2215", "ESF",  "C4:Z2215"),
            ("ERI",           "G4:AD1519", "ERI",  "C4:Z1519"),
        ]

        def buscar_hoja_exacta(workbook, nombre):
            for ws in workbook.Worksheets:
                if ws.Name.strip() == nombre.strip():
                    return ws
            return None

        def ejecutar_operaciones(wb_origen, wb_destino, operaciones, etiqueta):
            errores = 0
            for hoja_orig, rango_orig, hoja_dest, rango_dest in operaciones:
                ws_origen  = buscar_hoja_exacta(wb_origen,  hoja_orig)
                ws_destino = buscar_hoja_exacta(wb_destino, hoja_dest)

                if ws_origen and ws_destino:
                    valor = com_call(lambda o=ws_origen, r=rango_orig: o.Range(r).Value)
                    com_call(lambda d=ws_destino, r=rango_dest, v=valor: setattr(d.Range(r), "Value", v))
                    print(f"  ✓ {hoja_orig!r:20s} → {hoja_dest!r}")
                else:
                    faltante = []
                    if not ws_origen:  faltante.append(f"origen={hoja_orig!r}")
                    if not ws_destino: faltante.append(f"destino={hoja_dest!r}")
                    print(f"  ✗ Hoja no encontrada: {', '.join(faltante)}")
                    errores += 1

            print(f"  [{etiqueta}] {len(operaciones) - errores}/{len(operaciones)} operaciones OK\n")
            return errores

        # --- INICIO DEL FLUJO OPERATIVO ---
        excel = wb_cierre1 = wb_cierre2 = wb_destino = None
        pid_excel = None

        try:
            # CRÍTICO: Inicialización COM requerida para este hilo del Executor
            pythoncom.CoInitialize()
            
            import tempfile, shutil
            for patron in [
                os.path.join(tempfile.gettempdir(), "gen_py"),
                os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "gen_py"),
            ]:
                if os.path.exists(patron):
                    shutil.rmtree(patron, ignore_errors=True)

            excel = win32.DispatchEx("Excel.Application")
            print(f"Proceso Excel iniciado en hilo secundario\n")

            excel.Visible        = False
            excel.DisplayAlerts  = False
            excel.ScreenUpdating = False
            excel.EnableEvents   = False

            print("Abriendo archivos de control...\n")
            wb_cierre1 = com_call(lambda: excel.Workbooks.Open(RUTA_CIERRE_1, UpdateLinks=False, ReadOnly=True))
            wb_cierre2 = com_call(lambda: excel.Workbooks.Open(RUTA_CIERRE_2, UpdateLinks=False, ReadOnly=True))
            wb_destino = com_call(lambda: excel.Workbooks.Open(RUTA_DESTINO,  UpdateLinks=False, ReadOnly=False))

            excel.Calculation        = -4135  # Manual
            excel.CalculateBeforeSave = False

            print("=" * 55)
            print("BLOQUE 1 · Estado de Cierre → Base FONAFE")
            print("=" * 55)
            errores1 = ejecutar_operaciones(wb_cierre1, wb_destino, OPS_CIERRE_1, "Bloque 1")

            print("=" * 55)
            print("BLOQUE 2 · Validación Data → Base FONAFE")
            print("=" * 55)
            errores2 = ejecutar_operaciones(wb_cierre2, wb_destino, OPS_CIERRE_2, "Bloque 2")

            excel.Calculation = -4105  # Automático
            com_call(lambda: wb_destino.Save())

            total_errores = errores1 + errores2
            if total_errores == 0:
                print("✓ Guardado exitoso — sin errores.")
            else:
                print(f"⚠ Guardado con {total_errores} operación(es) fallida(s).")

        except Exception as e:
            print(f"\n✗ ERROR INESPERADO EN MATRIZ COM: {e}")
            # IMPORTANTE: se relanza para que el pipeline se detenga y no
            # continúe como si el copiado de Cierre/Validación se hubiera
            # completado, cuando en realidad la Base FONAFE quedó incompleta.
            raise

        finally:
            # Asegurar cierre de libros y desinicialización
            for wb in (wb_cierre1, wb_cierre2, wb_destino):
                try:
                    if wb: wb.Close(False)
                except Exception: pass

            try:
                if excel: excel.Quit()
            except Exception: pass

            excel = wb_cierre1 = wb_cierre2 = wb_destino = None
            gc.collect()
            
            # CRÍTICO: Liberar el entorno COM antes de abandonar el hilo
            pythoncom.CoUninitialize()
            
        return pid_excel

    # --- COORDINACIÓN DEL LOOP ASÍNCRONO ---
    loop = asyncio.get_running_loop()
    
    # max_workers=1 previene colisiones de punteros en la memoria de la instancia de Excel abierta
    with ThreadPoolExecutor(max_workers=1) as executor:
        await loop.run_in_executor(executor, _ejecutar_proceso_com_sync)