import os
from openpyxl import load_workbook


def limpiar_hojas_excel(ruta_principal):
    ruta_archivo = rf"{ruta_principal}\Base FONAFE WEB al mes.xlsm"
    try:

        print("Abriendo archivo Excel...")

        wb = load_workbook(
            ruta_archivo,
            keep_vba=True
        )

        hojas_a_limpiar = [
            "Depósitos y Colocaciones"
        ]

        rango_borrar = "A5:L6000"

        for nombre_hoja in hojas_a_limpiar:

            try:

                if nombre_hoja not in wb.sheetnames:

                    print(
                        f"La hoja '{nombre_hoja}' "
                        f"no existe."
                    )

                    continue

                ws = wb[nombre_hoja]

                for fila in ws[rango_borrar]:

                    for celda in fila:

                        celda.value = None

                        celda._style = celda.parent['A1']._style

                        celda.comment = None

                        celda.hyperlink = None

                print(
                    f"Rango borrado completamente "
                    f"en hoja: {nombre_hoja}"
                )

            except Exception as e_hoja:

                print(
                    f"No se pudo limpiar "
                    f"'{nombre_hoja}'. "
                    f"Error: {e_hoja}"
                )

        print("Guardando archivo...")

        wb.save(ruta_archivo)

        print("Archivo guardado correctamente.")

    except PermissionError:

        print(
            "Cierre el archivo Excel antes "
            "de ejecutar."
        )

    except Exception as e:

        print(
            f"Error crítico: {e}"
        )

    finally:

        print("Proceso finalizado.")
