from src.scrapping.Python_procesos import Crear_Descargar_0, Cierre_periodo_1, Actualizar_marcos, limpiarDatos, copiar_pegar_form_ejecu
from src.scrapping.Bases_Oficial import A_Actualizar_datos_iniciales as A, B_Copiar_Pegar_cierre_y_validacion as B, C_limpiza_GK as C , D_Crear_archivos_validacion_GK as D,E_Actualizar_Marcos as E,F_consolidar_Gastos_Capital_Flujo_de_caja as F, G_consolidar_Gastos_Capital_Presupuesto as G,H_Copiar_pegar_Validacion_GK_Flujo_Caja_Base_fonabe_FBK as H, I_Copiar_pegar_Validacion_GK_Presupuesto_Base_fonabe_FBK as I,J_desplegarFormula as J, K_cambiarFormato_llenarCeros as K,L_eliminar_Rango_Depositos_Colocaciones as L,M_Copiar_Pegar_Deposito_Colocaciones_Base as M, N_Formato_Depositos_Colocaciones_Bases as N
from src.scrapping.EVA_Oficial import A_eva, BF_BaseActualizaTD
import time
import os
import shutil
from datetime import datetime
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import getpass

def eliminar_procesos_excel():
    try:
        comando = "taskkill /f /im excel.exe"
        
        resultado = os.system(comando)

        if resultado == 0:
            print("Todos los procesos de Excel fueron eliminados correctamente.")
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

    fecha_ejecucion = datetime.now().strftime("%d.%m.%Y")
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
    carpeta_origen = r"Z:\Data_extraction\Plantillas"
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
def ejecutar_en_paralelo(trabajos):
    """Ejecuta funciones independientes en paralelo y espera a que todas terminen."""
    if not trabajos:
        return

    with ThreadPoolExecutor(max_workers=len(trabajos)) as executor:
        futuros = {
            executor.submit(funcion, *args, **kwargs): nombre
            for nombre, funcion, args, kwargs in trabajos
        }

        for futuro in as_completed(futuros):
            nombre = futuros[futuro]
            futuro.result()
            print(f"[{nombre}] finalizado en paralelo.")
def ejecutar_con_tiempo(nombre, funcion, *args, **kwargs):
    inicio = time.perf_counter()
    print(f"[{nombre}] iniciando...")
    resultado = funcion(*args, **kwargs)
    duracion = time.perf_counter() - inicio
    print(f"[{nombre}] finalizado en {duracion:.2f} segundos.")
    return resultado
def limpiar_procesos_menos_vscode():
    usuario = getpass.getuser()
    
    print(f"Iniciando limpieza de procesos para el usuario: {usuario}...")

    
    comando = (
        f'taskkill /F '
        f'/FI "USERNAME eq {usuario}" '
        f'/FI "IMAGENAME ne code.exe" '
        f'/FI "IMAGENAME ne python.exe" '
        f'/FI "IMAGENAME ne cmd.exe"'
    )
    resultado = os.system(comando)
    if resultado == 0:
        print("\nLimpieza completada con éxito. Solo quedaron VS Code y Python.")
    else:
        print("\nSe ejecutó el comando, pero algunos procesos protegidos no se pudieron cerrar.")

# ruta_principal = r"Z:\Data_extraction\TEST"
# anio = "2026"
# mes = "5"
# fecha_cierre_sistema = "11.06.2026"


def scrapping_main(ruta_principal, anio, mes, fecha_cierre_sistema):
    ruta_principal = rf"{ruta_principal}"
    eliminar_procesos_excel()
    print("Procesos de Excel eliminados. Iniciando proceso principal...")
    ejecutar_con_tiempo("copiarPlantillas", copiarPlantillas, ruta_principal)
    print("Python_procesos")
    ejecutar_con_tiempo("Crear_Descargar_0.crear_descargar", Crear_Descargar_0.crear_descargar, ruta_principal, anio, mes)
    ejecutar_con_tiempo("Actualizar_marcos.actualizar_marco", Actualizar_marcos.actualizar_marco, ruta_principal)
    ejecutar_con_tiempo("limpiarDatos.limpiar_datos", limpiarDatos.limpiar_datos, ruta_principal, mes)
    time.sleep(2)

    def ejecutar_cierre():
        ejecutar_con_tiempo("Cierre_periodo_1.cierre_periodo", Cierre_periodo_1.cierre_periodo, ruta_principal)

    def ejecutar_copia():
        ejecutar_con_tiempo("copiar_pegar_form_ejecu.copiar_pegar_form_ejecu", copiar_pegar_form_ejecu.copiar_pegar_form_ejecu, ruta_principal)

    hilo_cierre = threading.Thread(target=ejecutar_cierre)
    hilo_copia = threading.Thread(target=ejecutar_copia)

    hilo_cierre.start()
    hilo_copia.start()

    hilo_cierre.join()
    hilo_copia.join()

    print("Proceso Python_procesos finalizado")

    print("Bases_Oficial")

    ejecutar_con_tiempo("A.actualizar_datos_iniciales", A.actualizar_datos_iniciales, ruta_principal, anio, mes, fecha_cierre_sistema)

    ejecutar_con_tiempo("B.copiar_pegar_cierre_y_validacion", B.copiar_pegar_cierre_y_validacion, ruta_principal)
    print("empezamos a limpiar las hojas de excel")
    ejecutar_con_tiempo("C.limpiar_hojas_excel", C.limpiar_hojas_excel, ruta_principal)
    ejecutar_con_tiempo("D.crear_archivos_validacion", D.crear_archivos_validacion, ruta_principal)


    ejecutar_con_tiempo("E.actualizar_gastos_capital", E.actualizar_gastos_capital, ruta_principal, max_workers=16)

    ejecutar_con_tiempo("F/G consolidar_gk", ejecutar_en_paralelo, [
        ("F.consolidar_gk_flujo_caja", F.consolidar_gk_flujo_caja, (ruta_principal,), {}),
        ("G.consolidar_gk_presupuesto", G.consolidar_gk_presupuesto, (ruta_principal,), {}),
    ])

    ejecutar_con_tiempo("H.copiar_pegar_validacion_Flujo_Caja", H.copiar_pegar_validacion_Flujo_Caja, ruta_principal)

    ejecutar_con_tiempo("I.copiar_pegar_validacion_presupuesto", I.copiar_pegar_validacion_presupuesto, ruta_principal)

    ejecutar_con_tiempo("K.cambiar_formato", K.cambiar_formato, ruta_principal)

    ejecutar_con_tiempo("L.limpiar_hojas_excel", L.limpiar_hojas_excel, ruta_principal)

    ejecutar_con_tiempo("M.copiar_pegar_deposiciones_colocaciones", M.copiar_pegar_deposiciones_colocaciones, ruta_principal)

    ejecutar_con_tiempo("N.formato_deposiciones_colocaciones", N.formato_deposiciones_colocaciones, ruta_principal)

    print("Proceso finalizado")

    print("EVA_Oficial")

    A_eva.Copiar_Pegar(ruta_principal, anio, mes)

    ejecutar_con_tiempo("BF_BaseActualizaTD.actualizar_td", BF_BaseActualizaTD.actualizar_td, ruta_principal)
    print("Proceso finalizado")
    print("Renombrando archivos y carpetas...")
    formato_name_plantillas(ruta_principal, mes)
    eliminar_procesos_excel()
