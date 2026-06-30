import os
import time
import asyncio
from concurrent.futures import ThreadPoolExecutor
import pandas as pd
from openpyxl import load_workbook

async def actualizar_marco(ruta: str, max_workers: int = 8):
    ruta = os.path.join(ruta, "MARCO")
    CATEGORIA_ORDEN = {"M": 1, "I": 2, "P": 3, "S": 4, "3": 5}

    def procesar_archivo(ruta_archivo: str) -> dict:
        archivo = os.path.basename(ruta_archivo)
        resultado = {"archivo": archivo, "ok": False, "hoja": "", "eliminadas": 0, "error": ""}

        # Intentos repetidos por si el archivo está bloqueado en disco por Windows
        for intento in range(3):
            try:
                # Forzamos una copia limpia del DataFrame para evitar hilos compartiendo memoria
                df = pd.read_excel(ruta_archivo, sheet_name=0, header=0, engine="openpyxl").copy()

                col_empresa     = df.columns[2]   # columna C
                col_modificacion = df.columns[5]  # columna F

                df[col_empresa]      = df[col_empresa].astype(str).str.strip()
                df[col_modificacion] = df[col_modificacion].astype(str).str.strip()

                # ── Mapear orden numérico
                df["_orden"] = df[col_modificacion].map(CATEGORIA_ORDEN).fillna(0)

                idx_max = df.groupby(col_empresa)["_orden"].idxmax()
                # .copy() aquí es vital para romper cualquier referencia circular entre hilos
                mejor_mod = df.loc[idx_max, [col_empresa, col_modificacion]].set_index(col_empresa)[col_modificacion].copy()

                # ── Filtrar filas a conservar
                mascara = df.apply(
                    lambda r: mejor_mod.get(r[col_empresa], "") == r[col_modificacion], axis=1
                )
                eliminadas = int((~mascara).sum())
                df_filtrado = df[mascara].drop(columns=["_orden"]).reset_index(drop=True).copy()

                # ── Escribir hoja copia con openpyxl
                wb = load_workbook(ruta_archivo)
                nombre_copia = f"Formulacion_{len(wb.worksheets) + 1}"
                ws_copia = wb.create_sheet(title=nombre_copia)

                # Encabezados
                ws_copia.append(list(df_filtrado.columns))
                # Datos fila a fila
                for row in df_filtrado.itertuples(index=False, name=None):
                    ws_copia.append(list(row))

                wb.save(ruta_archivo)
                wb.close() # Forzamos el cierre explícito del puntero del archivo
                
                resultado.update({"ok": True, "hoja": nombre_copia, "eliminadas": eliminadas, "error": ""})
                break # Éxito, salimos del bucle de reintentos
                
            except PermissionError:
                # Si el archivo está bloqueado por el OS, esperamos una fracción de segundo y reintentamos
                time.sleep(0.5)
                resultado["error"] = "Archivo bloqueado en disco (I/O Concurrente)"
            except Exception as e:
                resultado["error"] = str(e)
                break # Si es otro error de código, no reintentamos

        return resultado

    async def actualizar_marcos_async_inner(carpeta: str, max_workers: int = 8):
        if not os.path.isdir(carpeta):
            print(f"La carpeta no existe: {carpeta}")
            return

        # Evitamos archivos duplicados usando un set antes de pasarlo a lista
        archivos_unicos = list(set([
            os.path.join(carpeta, f)
            for f in os.listdir(carpeta)
            if f.lower().endswith((".xlsx", ".xls", ".xlsm"))
            and not f.startswith("Gastos_Capital")
            and not f.startswith("~$") # IGNORAR archivos temporales ocultos de Excel
        ]))

        if not archivos_unicos:
            print("No se encontraron archivos Excel válidos para procesar.")
            return

        total   = len(archivos_unicos)
        workers = min(max_workers, total)
        print(f"Archivos únicos a procesar: {total} | Max Workers: {workers}\n")

        t_inicio    = time.perf_counter()
        completados = errores = 0

        loop = asyncio.get_running_loop()

        with ThreadPoolExecutor(max_workers=workers) as executor:
            tareas = [
                loop.run_in_executor(executor, procesar_archivo, ruta_archivo)
                for ruta_archivo in archivos_unicos
            ]
            
            for tarea_futura in asyncio.as_completed(tareas):
                res = await tarea_futura
                completados += 1
                estado = f"[{completados}/{total}]"

                if res["ok"]:
                    print(f"  {estado} OK  {res['archivo']}  ->  '{res['hoja']}' | eliminadas: {res['eliminadas']}")
                else:
                    errores += 1
                    print(f"  {estado} ERR {res['archivo']}  ->  {res['error']}")

        t_fin = time.perf_counter()
        print(f"\nCompletado en {t_fin - t_inicio:.2f}s | OK: {completados - errores} | Errores: {errores}")

    await actualizar_marcos_async_inner(ruta, max_workers=max_workers)