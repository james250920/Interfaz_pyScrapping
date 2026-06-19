import win32com.client as win32
import time
import pywintypes
import pythoncom

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


def refrescar_conexiones_y_tablas_dinamicas(wb):
    errores = 0

    try:
        conexiones = wb.Connections
    except Exception:
        conexiones = []

    for conexion in conexiones:
        try:
            # usar reintentos para manejar RPC_E_CALL_REJECTED
            com_call(lambda: conexion.Refresh())
        except pywintypes.com_error as e:
            errores += 1
            print(f"Aviso: no se pudo refrescar la conexión '{conexion.Name}': {e}")
        except Exception as e:
            errores += 1
            print(f"Aviso: no se pudo refrescar la conexión '{conexion.Name}': {e}")

    for hoja in wb.Worksheets:
        try:
            tablas = hoja.PivotTables()
        except Exception:
            continue

        try:
            total_tablas = tablas.Count
        except Exception:
            total_tablas = 0

        for indice in range(1, total_tablas + 1):
            try:
                tabla = tablas.Item(indice)
                # reintentar refresh de tabla dinámica si Excel responde ocupado
                com_call(lambda: tabla.RefreshTable())
            except pywintypes.com_error as e:
                errores += 1
                print(f"Aviso: no se pudo refrescar la tabla dinámica en '{hoja.Name}': {e}")
            except Exception as e:
                errores += 1
                print(f"Aviso: no se pudo refrescar la tabla dinámica en '{hoja.Name}': {e}")

    return errores

def actualizar_td(ruta_principal):
    ruta = rf"{ruta_principal}\Base FONAFE WEB al mes.xlsm"

    # Inicializar COM en este hilo
    pythoncom.CoInitialize()

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = com_call(lambda: excel.Workbooks.Open(ruta))

        ws = wb.Sheets("DEP Y COLOC_SALDOF_EFECT Y EQUI")

        print("Actualizando conexiones y tablas dinámicas...")

        errores = refrescar_conexiones_y_tablas_dinamicas(wb)

        while excel.CalculationState != 0:
            time.sleep(1)

        ws.Range("A44:B85").Copy()

        ws.Range("D44").PasteSpecial(Paste=-4163)

        excel.CutCopyMode = False

        com_call(lambda: wb.Save())

        while excel.CalculationState != 0:
            time.sleep(1)

        if errores == 0:
            print("Proceso completado correctamente.")
        else:
            print(f"Proceso completado con {errores} advertencia(s) de refresco.")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        try:
            com_call(lambda: wb.Close(SaveChanges=True))
        except Exception:
            pass

        try:
            com_call(lambda: excel.Quit())
        except Exception:
            pass

        # Desinicializar COM
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass