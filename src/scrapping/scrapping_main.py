from src.scrapping.Python_procesos import Crear_Descargar_0, Cierre_periodo_1, Actualizar_marcos, limpiarDatos, copiar_pegar_form_ejecu
from src.scrapping.Bases_Oficial import A_Actualizar_datos_iniciales as A, B_Copiar_Pegar_cierre_y_validacion as B, C_limpiza_GK as C , D_Crear_archivos_validacion_GK as D,E_Actualizar_Marcos as E,F_consolidar_Gastos_Capital_Flujo_de_caja as F, G_consolidar_Gastos_Capital_Presupuesto as G,H_Copiar_pegar_Validacion_GK_Flujo_Caja_Base_fonabe_FBK as H, I_Copiar_pegar_Validacion_GK_Presupuesto_Base_fonabe_FBK as I,J_desplegarFormula as J, K_cambiarFormato_llenarCeros as K,L_eliminar_Rango_Depositos_Colocaciones as L,M_Copiar_Pegar_Deposito_Colocaciones_Base as M, N_Formato_Depositos_Colocaciones_Bases as N
from src.scrapping.EVA_Oficial import A_eva
import time
import os
import shutil
import datetime
import re
import asyncio
import sys


def eliminar_procesos_excel(forzar=False):
    if not forzar:
        return
    try:
        comando = "taskkill /f /im excel.exe"
        resultado = os.system(comando)
        if resultado == 0:
            print("Todos los procesos de Excel fueron eliminados (fallback).")
        else:
            print("No se encontraron procesos de Excel activos o no se pudieron cerrar.")
    except Exception as e:
        print(f"Ocurrió un error al intentar cerrar Excel: {e}")


def formato_name_plantillas(ruta_principal, mes):
    meses_texto = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    try:
        mes_valido = meses_texto[int(mes)]
    except (ValueError, KeyError):
        mes_valido = str(mes).lower()

    fecha_ejecucion = datetime.datetime.now().strftime("%d.%m.%Y")
    patron_fecha_final = r"\s\d{2}\.\d{2}\.\d{4}$"
    carpetas_excluidas = {"logs"}

    if not os.path.exists(ruta_principal):
        return

    for elemento in os.listdir(ruta_principal):
        ruta_completa = os.path.join(ruta_principal, elemento)
        
        if os.path.isdir(ruta_completa):
            if elemento.lower() in carpetas_excluidas:
                continue

            if re.search(patron_fecha_final, elemento):
                continue
                
            nuevo_nombre = f"{elemento} {fecha_ejecucion}"
            try:
                os.rename(ruta_completa, os.path.join(ruta_principal, nuevo_nombre))
            except PermissionError as e:
                print(f"Aviso: no se pudo renombrar la carpeta '{elemento}': {e}")
            
        elif os.path.isfile(ruta_completa):
            nombre_base, extension = os.path.splitext(elemento)
            
            if re.search(patron_fecha_final, nombre_base):
                continue
                
            nuevo_nombre = f"{nombre_base} de {mes_valido} {fecha_ejecucion}{extension}"
            try:
                os.rename(ruta_completa, os.path.join(ruta_principal, nuevo_nombre))
            except PermissionError as e:
                print(f"Aviso: no se pudo renombrar el archivo '{elemento}': {e}")


def copiarPlantillas(ruta_principal):
    # Ruta configurable: variable de entorno RUTA_PLANTILLAS o valor por defecto
    carpeta_origen = os.environ.get("RUTA_PLANTILLAS", r"Z:\Data_extraction\Plantillas")
    carpeta_destino = ruta_principal
    # Asegúrate de que la carpeta de destino exista, si no, la crea
    os.makedirs(carpeta_destino, exist_ok=True)

    # Listar y copiar los archivos
    for nombre_archivo in os.listdir(carpeta_origen):
        ruta_completa_origen = os.path.join(carpeta_origen, nombre_archivo)
        
        # Comprobamos que sea un archivo y no una subcarpeta
        if os.path.isfile(ruta_completa_origen):
            shutil.copy2(ruta_completa_origen, carpeta_destino)
            print(f"Copiado: {nombre_archivo}")

    print("¡Proceso de copia finalizado!")


def ejecutar_con_tiempo(nombre, funcion, *args, **kwargs):
    inicio = time.perf_counter()
    print(f"[{nombre}] iniciando...")
    resultado = funcion(*args, **kwargs)
    duracion = time.perf_counter() - inicio
    print(f"[{nombre}] finalizado en {duracion:.2f} segundos.")
    return resultado


async def _pipeline_async(ruta_principal, anio, mes, fecha_cierre_sistema, reportar, check_cancel):
    """Coordinador async principal — un solo event loop para todo el pipeline."""
    loop = asyncio.get_running_loop()

    def verificar_cancelacion():
        if check_cancel and check_cancel():
            raise Exception("Proceso cancelado por el usuario")

    # ═══ Fase 1: Python_procesos ═══════════════════════════════════════════
    verificar_cancelacion()
    reportar("Descargando datos...")
    await Crear_Descargar_0.crear_descargar(ruta_principal, anio, mes)

    verificar_cancelacion()
    reportar("Actualizando marcos...")
    await Actualizar_marcos.actualizar_marco(ruta_principal, max_workers=16)

    verificar_cancelacion()
    reportar("Limpiando datos...")
    await limpiarDatos.limpiar_datos(ruta_principal, mes)
    await asyncio.sleep(2)

    verificar_cancelacion()
    reportar("Cierre...")
    await Cierre_periodo_1.cierre_periodo(ruta_principal)

    verificar_cancelacion()
    reportar("Copiando formulación/ejecución...")
    await copiar_pegar_form_ejecu.copiar_pegar_form_ejecu(ruta_principal)
    print("Proceso Python_procesos finalizado")

    # ═══ Fase 2: Bases_Oficial ══════════════════════════════════════════════
    verificar_cancelacion()
    reportar("Actualizando datos iniciales...")
    await A.actualizar_datos_iniciales(ruta_principal, anio, mes, fecha_cierre_sistema)

    verificar_cancelacion()
    reportar("Copiando cierre y validación...")
    await B.copiar_pegar_cierre_y_validacion(ruta_principal)

    verificar_cancelacion()
    reportar("Limpiando hojas Excel...")
    await C.limpiar_hojas_excel(ruta_principal)

    verificar_cancelacion()
    reportar("Creando archivos validación GK...")
    await loop.run_in_executor(None, D.crear_archivos_validacion, ruta_principal)

    verificar_cancelacion()
    reportar("Actualizando gastos de capital...")
    await E.actualizar_gastos_capital(ruta_principal, max_workers=16)

    verificar_cancelacion()
    reportar("Consolidando GK flujo caja...")
    await F.consolidar_gk_flujo_caja(ruta_principal)

    verificar_cancelacion()
    reportar("Consolidando GK presupuesto...")
    await G.consolidar_gk_presupuesto(ruta_principal)
    

    verificar_cancelacion()
    reportar("Validación Flujo de Caja...")
    await H.copiar_pegar_validacion_Flujo_Caja(ruta_principal)

    verificar_cancelacion()
    reportar("Validación Presupuesto...")
    await I.copiar_pegar_validacion_presupuesto(ruta_principal)

    verificar_cancelacion()
    reportar("Desplegando fórmulas...")
    await J.desplegar_formulas(ruta_principal)

    verificar_cancelacion()
    reportar("Cambiando formato...")
    await loop.run_in_executor(None, K.cambiar_formato, ruta_principal)

    verificar_cancelacion()
    reportar("Limpiando depósitos/colocaciones...")
    await L.limpiar_hojas_excel(ruta_principal)

    verificar_cancelacion()
    reportar("Copiando depósitos/colocaciones...")
    await M.copiar_pegar_deposiciones_colocaciones(ruta_principal)

    verificar_cancelacion()
    reportar("Formato depósitos/colocaciones...")
    await N.formato_deposiciones_colocaciones(ruta_principal)

    print("BASE finalizado")

    # ═══ Fase 3: EVA_Oficial ════════════════════════════════════════════════
    verificar_cancelacion()
    # Se eliminan los taskkill por regla
    reportar("EVA: Copiar y pegar...")
    await A_eva.copiar_pegar(ruta_principal, anio, mes)



def scrapping_main(ruta_principal, anio, mes, fecha_cierre_sistema, on_progreso=None, check_cancel=None):
    """Orquestador principal del proceso de scrapping.

    Args:
        ruta_principal: Directorio de trabajo donde se procesan los archivos.
        anio: Año del período a procesar (ej: "2026").
        mes: Número de mes del período (ej: "5").
        fecha_cierre_sistema: Fecha de cierre formateada (ej: "01.07.2026").
        on_progreso: Callback opcional que recibe un str con el mensaje de avance.
        check_cancel: Callback para verificar si se solicitó la cancelación de UI.
    """
    def reportar(msg):
        print(f"[Progreso] {msg}")
        if on_progreso:
            on_progreso(msg)

    reportar("Iniciando proceso...")
    # Solo en caso extremo se limpia al inicio si se fuerza, pero omitimos por defecto
    eliminar_procesos_excel(forzar=False)

    reportar("Copiando plantillas...")
    ejecutar_con_tiempo("copiarPlantillas", copiarPlantillas, ruta_principal)

    # Un solo event loop para todas las operaciones asíncronas
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    try:
        asyncio.run(_pipeline_async(ruta_principal, anio, mes, fecha_cierre_sistema, reportar, check_cancel))
    except Exception as e:
        if "Proceso cancelado" in str(e):
            reportar("Proceso cancelado de forma segura.")
            raise
        else:
            raise

    if check_cancel and check_cancel():
        raise Exception("Proceso cancelado por el usuario")

    reportar("Renombrando archivos y carpetas...")
    formato_name_plantillas(ruta_principal, mes)
    print("Proceso finalizado")