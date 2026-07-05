from src.scrapping.Python_procesos import (
    Crear_Descargar_0,
    Cierre_periodo_1,
    Actualizar_marcos,
    limpiarDatos,
    copiar_pegar_form_ejecu,
)

from src.scrapping.Bases_Oficial import (
    A_Actualizar_datos_iniciales as A,
    B_Copiar_Pegar_cierre_y_validacion as B,
    C_limpiza_GK as C,
    D_Crear_archivos_validacion_GK as D,
    E_Actualizar_Marcos as E,
    F_consolidar_Gastos_Capital_Flujo_de_caja as F,
    G_consolidar_Gastos_Capital_Presupuesto as G,
    H_Copiar_pegar_Validacion_GK_Flujo_Caja_Base_fonabe_FBK as H,
    I_Copiar_pegar_Validacion_GK_Presupuesto_Base_fonabe_FBK as I,
    J_desplegarFormula as J,
    K_cambiarFormato_llenarCeros as K,
    L_eliminar_Rango_Depositos_Colocaciones as L,
    M_Copiar_Pegar_Deposito_Colocaciones_Base as M,
    N_Formato_Depositos_Colocaciones_Bases as N,
)

from src.scrapping.EVA_Oficial import A_eva

import time
import os
import shutil
import datetime
import re
import inspect
from pathlib import Path
import asyncio


# ══════════════════════════════════════════════════════════════
# FUNCIONES EXISTENTES
# ══════════════════════════════════════════════════════════════


def formato_name_plantillas(ruta_principal, mes):
    ruta_base = Path(ruta_principal).expanduser().resolve()

    meses_texto = {
        1: "enero",
        2: "febrero",
        3: "marzo",
        4: "abril",
        5: "mayo",
        6: "junio",
        7: "julio",
        8: "agosto",
        9: "septiembre",
        10: "octubre",
        11: "noviembre",
        12: "diciembre",
    }

    try:
        mes_valido = meses_texto[int(mes)]
    except (ValueError, KeyError, TypeError):
        mes_valido = str(mes).lower()

    fecha_ejecucion = datetime.datetime.now().strftime("%d.%m.%Y")
    patron_fecha_final = r"\s\d{2}\.\d{2}\.\d{4}$"
    carpetas_excluidas = {"logs"}

    if not ruta_base.exists():
        print(f"Aviso: no existe la ruta principal: {ruta_base}", flush=True)
        return

    if not ruta_base.is_dir():
        print(f"Aviso: la ruta principal no es una carpeta válida: {ruta_base}", flush=True)
        return

    for elemento in sorted(os.listdir(ruta_base)):
        ruta_completa = ruta_base / elemento

        if ruta_completa.is_dir():
            if elemento.lower() in carpetas_excluidas:
                continue

            if re.search(patron_fecha_final, elemento):
                continue

            nuevo_nombre = f"{elemento} {fecha_ejecucion}"
            nueva_ruta = ruta_base / nuevo_nombre

            if nueva_ruta.exists():
                print(
                    f"Aviso: no se renombra '{elemento}' porque ya existe '{nuevo_nombre}'",
                    flush=True,
                )
                continue

            try:
                os.rename(str(ruta_completa), str(nueva_ruta))
            except PermissionError as e:
                print(f"Aviso: no se pudo renombrar la carpeta '{elemento}': {e}", flush=True)
            except OSError as e:
                print(f"Aviso: error renombrando carpeta '{elemento}': {e}", flush=True)

        elif ruta_completa.is_file():
            nombre_base, extension = os.path.splitext(elemento)

            if re.search(patron_fecha_final, nombre_base):
                continue

            nuevo_nombre = f"{nombre_base} de {mes_valido} {fecha_ejecucion}{extension}"
            nueva_ruta = ruta_base / nuevo_nombre

            if nueva_ruta.exists():
                print(
                    f"Aviso: no se renombra '{elemento}' porque ya existe '{nuevo_nombre}'",
                    flush=True,
                )
                continue

            try:
                os.rename(str(ruta_completa), str(nueva_ruta))
            except PermissionError as e:
                print(f"Aviso: no se pudo renombrar el archivo '{elemento}': {e}", flush=True)
            except OSError as e:
                print(f"Aviso: error renombrando archivo '{elemento}': {e}", flush=True)


def copiarPlantillas(ruta_principal):
    carpeta_origen = Path(
        os.environ.get(
            "RUTA_PLANTILLAS",
            r"C:\Users\james\OneDrive\Escritorio\plantilla",
        )
    ).expanduser().resolve()

    carpeta_destino = Path(ruta_principal).expanduser().resolve()

    if not carpeta_origen.exists():
        raise FileNotFoundError(f"No existe la carpeta de plantillas: {carpeta_origen}")

    if not carpeta_origen.is_dir():
        raise NotADirectoryError(f"La ruta de plantillas no es una carpeta: {carpeta_origen}")

    carpeta_destino.mkdir(parents=True, exist_ok=True)

    for nombre_archivo in sorted(os.listdir(carpeta_origen)):
        ruta_completa_origen = carpeta_origen / nombre_archivo

        if not ruta_completa_origen.is_file():
            continue

        destino_final = carpeta_destino / nombre_archivo
        destino_temporal = carpeta_destino / f".{nombre_archivo}.tmp"

        try:
            shutil.copy2(str(ruta_completa_origen), str(destino_temporal))
            os.replace(str(destino_temporal), str(destino_final))
            print(f"Copiado: {nombre_archivo}", flush=True)

        finally:
            try:
                if destino_temporal.exists():
                    destino_temporal.unlink()
            except OSError as e:
                print(f"Aviso: no se pudo eliminar temporal '{destino_temporal}': {e}", flush=True)

    print("¡Proceso de copia finalizado!", flush=True)


def _validar_resultado_no_async(nombre, resultado):
    """
    Evita que el pipeline continúe si una función que debe ser síncrona
    todavía devuelve una corrutina.
    """

    if inspect.isawaitable(resultado):
        try:
            if hasattr(resultado, "close"):
                resultado.close()
        except Exception:
            pass

        raise TypeError(
            f"[{nombre}] devolvió una corrutina/awaitable. "
            "Solo Crear_Descargar_0.crear_descargar puede ser async. "
            "Convierte esta función a síncrona antes de llamarla desde el pipeline."
        )


def ejecutar_con_tiempo(nombre, funcion, *args, **kwargs):
    inicio = time.perf_counter()
    print(f"[{nombre}] iniciando...", flush=True)

    resultado = funcion(*args, **kwargs)
    _validar_resultado_no_async(nombre, resultado)

    duracion = time.perf_counter() - inicio
    print(f"[{nombre}] finalizado en {duracion:.2f} segundos.", flush=True)

    if resultado is False:
        raise RuntimeError(f"[{nombre}] finalizó con resultado False")

    return resultado


async def _ejecutar_crear_descargar_async(ruta_principal, anio, mes):
    """
    Único punto async permitido.
    """

    resultado = Crear_Descargar_0.crear_descargar(ruta_principal, anio, mes)

    if inspect.isawaitable(resultado):
        resultado = await resultado

    if resultado is False:
        raise RuntimeError("[Crear_Descargar_0.crear_descargar] finalizó con resultado False")

    return resultado


def ejecutar_crear_descargar_con_tiempo(ruta_principal, anio, mes):
    """
    Ejecuta únicamente Crear_Descargar_0.crear_descargar con asyncio.run().
    No introduce paralelismo.
    """

    inicio = time.perf_counter()
    nombre = "Crear_Descargar_0.crear_descargar"

    print(f"[{nombre}] iniciando...", flush=True)

    resultado = asyncio.run(
        _ejecutar_crear_descargar_async(ruta_principal, anio, mes)
    )

    duracion = time.perf_counter() - inicio
    print(f"[{nombre}] finalizado en {duracion:.2f} segundos.", flush=True)

    return resultado


def _verificar_cancelacion(check_cancel):
    if check_cancel and check_cancel():
        raise Exception("Proceso cancelado por el usuario")


def _pipeline_sync(
    ruta_principal,
    anio,
    mes,
    fecha_cierre_sistema,
    reportar,
    check_cancel,
):
    """
    Coordinador principal estrictamente secuencial.

    Regla:
    - Crear_Descargar_0.crear_descargar puede ser async.
    - Todo lo demás debe ser síncrono.
    """

    _verificar_cancelacion(check_cancel)
    reportar("Descargando datos...")
    ejecutar_crear_descargar_con_tiempo(
        ruta_principal,
        anio,
        mes,
    )

    _verificar_cancelacion(check_cancel)
    reportar("Actualizando marcos...")
    ejecutar_con_tiempo(
        "Actualizar_marcos.actualizar_marco",
        Actualizar_marcos.actualizar_marco,
        ruta_principal,
    )

    _verificar_cancelacion(check_cancel)
    reportar("Limpiando datos...")
    ejecutar_con_tiempo(
        "limpiarDatos.limpiar_datos",
        limpiarDatos.limpiar_datos,
        ruta_principal,
        mes,
    )

    time.sleep(2)

    _verificar_cancelacion(check_cancel)
    reportar("Cierre...")
    ejecutar_con_tiempo(
        "Cierre_periodo_1.cierre_periodo",
        Cierre_periodo_1.cierre_periodo,
        ruta_principal,
    )

    _verificar_cancelacion(check_cancel)
    reportar("Copiando formulación/ejecución...")
    ejecutar_con_tiempo(
        "copiar_pegar_form_ejecu.copiar_pegar_form_ejecu",
        copiar_pegar_form_ejecu.copiar_pegar_form_ejecu,
        ruta_principal,
    )

    print("Proceso Python_procesos finalizado", flush=True)
# ═══ Fase 2: Bases_Oficia**═════════════════════════════════**═══════════


    # _verificar_cancelacion(check_cancel)
    # reportar("Actualizando datos iniciales...")
    # ejecutar_con_tiempo(
    #     "A.actualizar_datos_iniciales",
    #     A.actualizar_datos_iniciales,
    #     ruta_principal,
    #     anio,
    #     mes,
    #     fecha_cierre_sistema,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Copiando cierre y validación...")
    # ejecutar_con_tiempo(
    #     "B.copiar_pegar_cierre_y_validacion",
    #     B.copiar_pegar_cierre_y_validacion,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Limpiando hojas Excel...")
    # ejecutar_con_tiempo(
    #     "C.limpiar_hojas_excel",
    #     C.limpiar_hojas_excel,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Creando archivos validación GK...")
    # ejecutar_con_tiempo(
    #     "D.crear_archivos_validacion",
    #     D.crear_archivos_validacion,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Actualizando gastos de capital...")
    # ejecutar_con_tiempo(
    #     "E.actualizar_gastos_capital",
    #     E.actualizar_gastos_capital,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Consolidando GK flujo caja...")
    # ejecutar_con_tiempo(
    #     "F.consolidar_gk_flujo_caja",
    #     F.consolidar_gk_flujo_caja,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Consolidando GK presupuesto...")
    # ejecutar_con_tiempo(
    #     "G.consolidar_gk_presupuesto",
    #     G.consolidar_gk_presupuesto,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Validación Flujo de Caja...")
    # ejecutar_con_tiempo(
    #     "H.copiar_pegar_validacion_Flujo_Caja",
    #     H.copiar_pegar_validacion_Flujo_Caja,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Validación Presupuesto...")
    # ejecutar_con_tiempo(
    #     "I.copiar_pegar_validacion_presupuesto",
    #     I.copiar_pegar_validacion_presupuesto,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Desplegando fórmulas...")
    # ejecutar_con_tiempo(
    #     "J.desplegar_formulas",
    #     J.desplegar_formulas,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Cambiando formato...")
    # ejecutar_con_tiempo(
    #     "K.cambiar_formato",
    #     K.cambiar_formato,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Limpiando depósitos/colocaciones...")
    # ejecutar_con_tiempo(
    #     "L.limpiar_hojas_excel",
    #     L.limpiar_hojas_excel,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Copiando depósitos/colocaciones...")
    # ejecutar_con_tiempo(
    #     "M.copiar_pegar_deposiciones_colocaciones",
    #     M.copiar_pegar_deposiciones_colocaciones,
    #     ruta_principal,
    # )

    # _verificar_cancelacion(check_cancel)
    # reportar("Formato depósitos/colocaciones...")
    # ejecutar_con_tiempo(
    #     "N.formato_deposiciones_colocaciones",
    #     N.formato_deposiciones_colocaciones,
    #     ruta_principal,
    # )

    # print("BASE finalizado")

    # ═══ Fase 3: EVA_Oficial ════════════════════════════════════════════════

    # _verificar_cancelacion(check_cancel)
    # reportar("EVA: Copiar y pegar...")
    # ejecutar_con_tiempo(
    #     "A_eva.copiar_pegar",
    #     A_eva.copiar_pegar,
    #     ruta_principal,
    #     anio,
    #     mes,
    # )

def scrapping_main(
    ruta_principal,
    anio,
    mes,
    fecha_cierre_sistema,
    on_progreso=None,
    check_cancel=None,
):
    """
    Orquestador principal.

    Ejecución:
    - Secuencial.
    - async solo para Crear_Descargar_0.crear_descargar.
    - Sin ThreadPoolExecutor.
    - Sin run_in_executor.
    - Sin paralelismo sobre Excel ni archivos.

    Al finalizar:
    - Mata forzosamente el proceso principal del worker.
    """

    ruta_base = Path(ruta_principal).expanduser().resolve()

    def reportar(msg):
        print(f"[Progreso] {msg}", flush=True)

        if on_progreso:
            on_progreso(msg)

    try:
        if not ruta_base.exists():
            raise FileNotFoundError(f"No existe la ruta principal: {ruta_base}")

        if not ruta_base.is_dir():
            raise NotADirectoryError(
                f"La ruta principal no es una carpeta válida: {ruta_base}"
            )

        reportar("Iniciando proceso...")

        _verificar_cancelacion(check_cancel)
        reportar("Copiando plantillas...")
        ejecutar_con_tiempo(
            "copiarPlantillas",
            copiarPlantillas,
            str(ruta_base),
        )

        try:
            _pipeline_sync(
                str(ruta_base),
                anio,
                mes,
                fecha_cierre_sistema,
                reportar,
                check_cancel,
            )

        except Exception as e:
            if "Proceso cancelado" in str(e):
                reportar("Proceso cancelado de forma segura.")
                print(f"ERROR::{e}", flush=True)
                return

            reportar(f"Proceso detenido por error: {e}")
            print(f"ERROR::{e}", flush=True)
            return

        _verificar_cancelacion(check_cancel)

        reportar("Renombrando archivos y carpetas...")
        ejecutar_con_tiempo(
            "formato_name_plantillas",
            formato_name_plantillas,
            str(ruta_base),
            mes,
        )

        reportar("Proceso finalizado")
        print("Proceso finalizado", flush=True)
        print("DONE::Proceso finalizado correctamente", flush=True)

        return

    except Exception as e:
        reportar(f"Proceso detenido por error: {e}")
        print(f"ERROR::{e}", flush=True)
        return