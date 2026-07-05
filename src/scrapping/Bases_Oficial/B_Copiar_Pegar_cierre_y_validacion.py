import os
import time
import gc
import tempfile
import shutil
import logging

import pythoncom
import pywintypes
import win32com.client as win32


logger = logging.getLogger(__name__)

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846

# 0x800AC472: Excel ocupado / no puede completar la tarea.
VBA_E_IGNORE = -2146777998

HRESULTS_REINTENTABLES = (
    RPC_E_CALL_REJECTED,
    RPC_E_SERVERCALL_RETRYLATER,
    VBA_E_IGNORE,
)

XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105

# msoAutomationSecurityForceDisable = 3
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def com_call(fn, reintentos=12, pausa=2.5):
    """
    Ejecuta una llamada COM con reintentos síncronos únicamente ante errores
    transitorios habituales de Excel ocupado/rechazando llamadas.
    """
    ultimo_error = None

    for intento in range(1, reintentos + 1):
        try:
            return fn()
        except pywintypes.com_error as e:
            ultimo_error = e

            if e.hresult in HRESULTS_REINTENTABLES:
                if intento == reintentos:
                    raise

                print(
                    f"    ⏳ Excel ocupado (hresult={e.hresult}), "
                    f"reintento {intento}/{reintentos} en {pausa}s..."
                )
                time.sleep(pausa)
            else:
                raise

    raise RuntimeError(
        f"Excel rechazó la llamada tras {reintentos} reintentos."
    ) from ultimo_error


def copiar_pegar_cierre_y_validacion(ruta_principal):
    inicio = time.time()

    ruta_principal_abs = os.path.abspath(ruta_principal)

    RUTA_CIERRE_1 = os.path.abspath(
        os.path.join(ruta_principal_abs, "Estado de Cierre al mes.xlsm")
    )
    RUTA_CIERRE_2 = os.path.abspath(
        os.path.join(ruta_principal_abs, "Validación Data Marco y Ejecución al mes.xlsm")
    )
    RUTA_DESTINO = os.path.abspath(
        os.path.join(ruta_principal_abs, "Base FONAFE WEB al mes.xlsm")
    )

    OPS_CIERRE_1 = [
        ("Resumen Marco", "D5:R41",   "Sistema Cierre", "D52:R88"),
        ("Resumen Ejec",  "D5:R41",   "Sistema Cierre", "D9:R45"),
        ("Resumen Marco", "B1",       "Sistema Cierre", "D47"),
        ("Resumen Ejec",  "S5:AM41",  "Cierre FE",      "D7:X43"),
        ("Resumen Ejec",  "AN5:BB41", "Cierre FE",      "AB7:AP43"),
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
        """
        Busca hoja por nombre normalizado con strip(), manteniendo la lógica
        funcional original.
        """
        worksheets = None
        ws = None

        try:
            worksheets = workbook.Worksheets
            total_hojas = com_call(lambda: worksheets.Count)

            for indice in range(1, total_hojas + 1):
                ws = com_call(lambda i=indice: worksheets.Item(i))
                nombre_ws = com_call(lambda hoja=ws: hoja.Name)

                if nombre_ws.strip() == nombre.strip():
                    return ws

                ws = None

            return None

        finally:
            worksheets = None

    def ejecutar_operaciones(wb_origen, wb_destino, operaciones, etiqueta):
        """
        Ejecuta las operaciones de copiado de valores de forma estrictamente
        secuencial. No agrega concurrencia ni paralelismo.
        """
        errores = 0

        for hoja_orig, rango_orig, hoja_dest, rango_dest in operaciones:
            ws_origen = None
            ws_destino = None
            rango_origen_obj = None
            rango_destino_obj = None

            try:
                ws_origen = buscar_hoja_exacta(wb_origen, hoja_orig)
                ws_destino = buscar_hoja_exacta(wb_destino, hoja_dest)

                if ws_origen is not None and ws_destino is not None:
                    rango_origen_obj = com_call(
                        lambda o=ws_origen, r=rango_orig: o.Range(r)
                    )
                    valor = com_call(lambda rg=rango_origen_obj: rg.Value)

                    rango_destino_obj = com_call(
                        lambda d=ws_destino, r=rango_dest: d.Range(r)
                    )
                    com_call(lambda rg=rango_destino_obj, v=valor: setattr(rg, "Value", v))

                    print(f"  ✓ {hoja_orig!r:20s} → {hoja_dest!r}")
                else:
                    faltante = []

                    if ws_origen is None:
                        faltante.append(f"origen={hoja_orig!r}")

                    if ws_destino is None:
                        faltante.append(f"destino={hoja_dest!r}")

                    print(f"  ✗ Hoja no encontrada: {', '.join(faltante)}")
                    errores += 1

            finally:
                rango_destino_obj = None
                rango_origen_obj = None
                ws_destino = None
                ws_origen = None

        print(f"  [{etiqueta}] {len(operaciones) - errores}/{len(operaciones)} operaciones OK\n")
        return errores

    excel = None
    wb_cierre1 = None
    wb_cierre2 = None
    wb_destino = None

    ruta_temporal = None
    com_inicializado = False
    guardado_correcto = False
    reemplazo_correcto = False

    try:
        if not os.path.isdir(ruta_principal_abs):
            raise NotADirectoryError(f"La carpeta principal no existe: {ruta_principal_abs}")

        for ruta in (RUTA_CIERRE_1, RUTA_CIERRE_2, RUTA_DESTINO):
            if not os.path.exists(ruta):
                raise FileNotFoundError(f"El archivo no existe: {ruta}")

            if not os.path.isfile(ruta):
                raise FileNotFoundError(f"La ruta no corresponde a un archivo válido: {ruta}")

        # Inicialización COM explícita en el hilo actual.
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        com_inicializado = True

        # Crear copia temporal del destino antes de abrir Excel.
        fd, temp_name = tempfile.mkstemp(
            prefix="Base FONAFE WEB al mes_",
            suffix=".xlsm",
            dir=os.path.dirname(RUTA_DESTINO)
        )
        os.close(fd)

        ruta_temporal = os.path.abspath(temp_name)
        shutil.copy2(RUTA_DESTINO, ruta_temporal)

        excel = win32.DispatchEx("Excel.Application")
        print("Proceso Excel iniciado\n")

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.DisplayStatusBar = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        print("Abriendo archivos de control...\n")

        wb_cierre1 = com_call(
            lambda: excel.Workbooks.Open(
                RUTA_CIERRE_1,
                UpdateLinks=False,
                ReadOnly=True
            )
        )
        wb_cierre2 = com_call(
            lambda: excel.Workbooks.Open(
                RUTA_CIERRE_2,
                UpdateLinks=False,
                ReadOnly=True
            )
        )
        wb_destino = com_call(
            lambda: excel.Workbooks.Open(
                ruta_temporal,
                UpdateLinks=False,
                ReadOnly=False
            )
        )

        com_call(lambda: setattr(excel, "Calculation", XL_CALCULATION_MANUAL))
        com_call(lambda: setattr(excel, "CalculateBeforeSave", False))

        print("=" * 55)
        print("BLOQUE 1 · Estado de Cierre → Base FONAFE")
        print("=" * 55)
        errores1 = ejecutar_operaciones(
            wb_cierre1,
            wb_destino,
            OPS_CIERRE_1,
            "Bloque 1"
        )

        print("=" * 55)
        print("BLOQUE 2 · Validación Data → Base FONAFE")
        print("=" * 55)
        errores2 = ejecutar_operaciones(
            wb_cierre2,
            wb_destino,
            OPS_CIERRE_2,
            "Bloque 2"
        )

        com_call(lambda: setattr(excel, "Calculation", XL_CALCULATION_AUTOMATIC))

        print("Guardando archivo temporal...")
        com_call(lambda: wb_destino.Save())
        guardado_correcto = True

        total_errores = errores1 + errores2

        if total_errores == 0:
            print("✓ Guardado exitoso — sin errores.")
        else:
            print(f"⚠ Guardado con {total_errores} operación(es) fallida(s).")

    except Exception as e:
        print(f"\n✗ ERROR INESPERADO EN MATRIZ COM: {e}")
        logger.exception("Error inesperado en matriz COM")
        raise

    finally:
        if excel is not None:
            try:
                com_call(lambda: setattr(excel, "Calculation", XL_CALCULATION_AUTOMATIC))
            except Exception as e:
                logger.warning("No se pudo restaurar cálculo automático de Excel: %s", e)

            try:
                excel.ScreenUpdating = True
                excel.EnableEvents = True
                excel.DisplayStatusBar = True
            except Exception as e:
                logger.warning("No se pudieron restaurar propiedades de Excel: %s", e)

        # Cierre de libros. El destino se cierra sin guardar cambios adicionales
        # porque el guardado explícito ya ocurrió antes.
        if wb_destino is not None:
            try:
                com_call(lambda: wb_destino.Close(SaveChanges=False))
            except Exception as e:
                logger.warning("No se pudo cerrar el workbook destino: %s", e)
            finally:
                wb_destino = None

        if wb_cierre2 is not None:
            try:
                com_call(lambda: wb_cierre2.Close(SaveChanges=False))
            except Exception as e:
                logger.warning("No se pudo cerrar el workbook cierre 2: %s", e)
            finally:
                wb_cierre2 = None

        if wb_cierre1 is not None:
            try:
                com_call(lambda: wb_cierre1.Close(SaveChanges=False))
            except Exception as e:
                logger.warning("No se pudo cerrar el workbook cierre 1: %s", e)
            finally:
                wb_cierre1 = None

        if excel is not None:
            try:
                excel.Quit()
            except Exception as e:
                logger.warning("No se pudo cerrar Excel correctamente: %s", e)
            finally:
                excel = None

        gc.collect()

        if com_inicializado:
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.warning("No se pudo desinicializar COM correctamente: %s", e)

        # Reemplazo atómico solo si el archivo temporal se guardó correctamente.
        if ruta_temporal:
            if guardado_correcto:
                try:
                    os.replace(ruta_temporal, RUTA_DESTINO)
                    reemplazo_correcto = True
                    print(f"Archivo destino reemplazado atómicamente: {RUTA_DESTINO}")
                except Exception as e:
                    print(
                        "⚠ Advertencia: no se pudo reemplazar el archivo original, "
                        f"el temporal quedó en: {ruta_temporal}"
                    )
                    logger.warning(
                        "No se pudo reemplazar el archivo original %s con %s: %s",
                        RUTA_DESTINO,
                        ruta_temporal,
                        e
                    )
            else:
                try:
                    if os.path.exists(ruta_temporal):
                        os.remove(ruta_temporal)
                except Exception as e:
                    logger.warning(
                        "No se pudo eliminar el archivo temporal fallido %s: %s",
                        ruta_temporal,
                        e
                    )

    if guardado_correcto and reemplazo_correcto:
        print(f"✓ Copia y validación completadas en {round(time.time() - inicio, 2)} segundos.")

    return None