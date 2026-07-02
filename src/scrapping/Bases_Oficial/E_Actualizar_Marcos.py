import os
import time
import asyncio
from concurrent.futures import ThreadPoolExecutor
import pandas as pd
from openpyxl import load_workbook

# Modificada a función asíncrona principal
async def actualizar_gastos_capital(ruta_principal: str, max_workers: int = 8):
    ruta = os.path.join(ruta_principal, "MARCO")
    CATEGORIA_ORDEN = {"M": 1, "I": 2, "P": 3, "S": 4, "3": 5}

    # Función síncrona interna aislada para el Pool de Hilos
    def procesar_archivo(ruta_archivo: str) -> dict:
        archivo = os.path.basename(ruta_archivo)
        resultado = {
            "archivo": archivo,
            "ok": False,
            "hoja": "",
            "eliminadas": 0,
            "error": ""
        }

        # Intentos repetidos por si Windows bloquea temporalmente el archivo en disco
        for intento in range(3):
            try:
                # Forzamos .copy() para evitar fugas de memoria compartida entre hilos
                df = pd.read_excel(ruta_archivo, sheet_name=0, header=0, engine="openpyxl").copy()

                col_empresa      = df.columns[2]   # columna C
                col_modificacion = df.columns[5]   # columna F

                df[col_empresa] = df[col_empresa].astype(str).str.strip()
                df[col_modificacion] = df[col_modificacion].astype(str).str.strip()

                df["_orden"] = (
                    df[col_modificacion]
                    .map(CATEGORIA_ORDEN)
                    .fillna(0)
                )

                idx_max = df.groupby(col_empresa)["_orden"].idxmax()

                mejor_mod = (
                    df.loc[idx_max, [col_empresa, col_modificacion]]
                    .set_index(col_empresa)[col_modificacion]
                    .copy()
                )

                mascara = df.apply(
                    lambda r: mejor_mod.get(r[col_empresa], "") == r[col_modificacion],
                    axis=1
                )

                eliminadas = int((~mascara).sum())

                df_filtrado = (
                    df[mascara]
                    .drop(columns=["_orden"])
                    .reset_index(drop=True)
                    .copy()
                )

                wb = load_workbook(ruta_archivo)
                nombre_copia = f"Gastos_Capital_{len(wb.worksheets) + 1}"
                ws_copia = wb.create_sheet(title=nombre_copia)

                # Encabezados
                ws_copia.append(list(df_filtrado.columns))

                # Datos
                for row in df_filtrado.itertuples(index=False, name=None):
                    ws_copia.append(list(row))

                temp_file = ruta_archivo + ".tmp"
                wb.save(temp_file)
                wb.close() # Cierre explícito del puntero físico del archivo
                os.replace(temp_file, ruta_archivo)

                resultado.update({
                    "ok": True,
                    "hoja": nombre_copia,
                    "eliminadas": eliminadas,
                    "error": ""
                })
                break # Éxito, rompemos el ciclo de reintentos

            except PermissionError:
                # Si el sistema operativo retiene el archivo, pausa breve y reintenta
                time.sleep(0.5)
                resultado["error"] = "Archivo bloqueado temporalmente por I/O"
            except Exception as e:
                resultado["error"] = str(e)
                break

        return resultado

    # Coordinación interna del flujo asíncrono
    async def actualizar_archivos_async_inner(carpeta: str, max_workers: int = 8):
        if not os.path.isdir(carpeta):
            print(f"La carpeta no existe: {carpeta}")
            return

        # Filtro estricto: único, sin temporales ocultos de Excel (~$)
        archivos_unicos = list(set([
            os.path.join(carpeta, f)
            for f in os.listdir(carpeta)
            if f.lower().endswith((".xlsx", ".xls", ".xlsm"))
            and f.startswith("Gastos_Capital")
            and not f.startswith("~$")
        ]))

        if not archivos_unicos:
            print("No se encontraron archivos Gastos_Capital válidos.")
            return

        total = len(archivos_unicos)
        workers = min(max_workers, total)

        print(f"Archivos Gastos_Capital únicos: {total} | Max Workers: {workers}\n")

        t_inicio = time.perf_counter()
        completados = 0
        errores = 0

        loop = asyncio.get_running_loop()

        with ThreadPoolExecutor(max_workers=workers) as executor:
            # Lista de tareas mapeadas de forma asíncrona
            tareas = [
                loop.run_in_executor(executor, procesar_archivo, ruta_archivo)
                for ruta_archivo in archivos_unicos
            ]

            # Procesamiento iterativo conforme se completan en segundo plano
            for tarea_futura in asyncio.as_completed(tareas):
                res = await tarea_futura
                completados += 1
                estado = f"[{completados}/{total}]"

                if res["ok"]:
                    print(
                        f"  {estado} OK  {res['archivo']}  "
                        f"-> '{res['hoja']}' | eliminadas: {res['eliminadas']}"
                    )
                else:
                    errores += 1
                    print(
                        f"  {estado} ERR {res['archivo']}  "
                        f"->  {res['error']}"
                    )

        t_fin = time.perf_counter()
        print(
            f"\nCompletado en {t_fin - t_inicio:.2f}s | "
            f"OK: {completados - errores} | "
            f"Errores: {errores}"
        )

    # Invocación interna asíncrona
    await actualizar_archivos_async_inner(ruta, max_workers=max_workers)

