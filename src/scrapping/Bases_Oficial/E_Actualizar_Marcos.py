import os
import time
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def actualizar_gastos_capital(ruta_principal: str, max_workers: int = 8):
    ruta = rf"{ruta_principal}\MARCO"
    CATEGORIA_ORDEN = {"M": 1, "I": 2, "P": 3, "S": 4, "3": 5}


    def procesar_archivo(ruta_archivo: str) -> dict:

        archivo = os.path.basename(ruta_archivo)

        resultado = {
            "archivo": archivo,
            "ok": False,
            "hoja": "",
            "eliminadas": 0,
            "error": ""
        }

        try:
            df = pd.read_excel(
                ruta_archivo,
                sheet_name=0,
                header=0,
                engine="openpyxl"
            )

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
            )

            wb = load_workbook(ruta_archivo)

            nombre_copia = f"Gastos_Capital_{len(wb.worksheets) + 1}"

            ws_copia = wb.create_sheet(title=nombre_copia)

            # Encabezados
            ws_copia.append(list(df_filtrado.columns))

            # Datos
            for row in df_filtrado.itertuples(index=False, name=None):
                ws_copia.append(list(row))

            wb.save(ruta_archivo)

            resultado.update({
                "ok": True,
                "hoja": nombre_copia,
                "eliminadas": eliminadas
            })

        except Exception as e:
            resultado["error"] = str(e)

        return resultado


    def actualizar_archivos(carpeta: str, max_workers: int = 8):

        if not os.path.isdir(carpeta):
            print(f"La carpeta no existe: {carpeta}")
            return

        archivos = [
            os.path.join(carpeta, f)
            for f in os.listdir(carpeta)
            if f.lower().endswith((".xlsx", ".xls", ".xlsm"))
            and f.startswith("Gastos_Capital")
        ]

        if not archivos:
            print("No se encontraron archivos Gastos_Capital.")
            return

        total = len(archivos)
        workers = min(max_workers, total)

        print(f"Archivos: {total} | Hilos: {workers}\n")

        t_inicio = time.perf_counter()

        completados = 0
        errores = 0

        with ThreadPoolExecutor(max_workers=workers) as executor:

            futuros = {
                executor.submit(procesar_archivo, ruta): ruta
                for ruta in archivos
            }

            for futuro in as_completed(futuros):

                res = futuro.result()

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


    actualizar_archivos(ruta, max_workers=16)

# ruta = r"D:\GESTION PRESUPUESTAL\MARCO"

# actualizar_gastos_capital(ruta, max_workers=16)