import os
from openpyxl import Workbook

def crear_archivos_validacion(ruta_principal):
    ruta = rf"{ruta_principal}\VALIDACION GASTO CAPITAL\Validacion_Gastos_Capital_Flujo_Caja.xlsx"
    ruta2 = rf"{ruta_principal}\VALIDACION GASTO CAPITAL\Validación_Gastos_Capital_Presupuesto.xlsx"

    if not os.path.exists(ruta) or not os.path.exists(ruta2):

        wb = Workbook()
        wb2 = Workbook()

        ws = wb.active
        ws.title = "Hoja1"

        ws2 = wb2.active
        ws2.title = "Hoja1"

        wb.save(ruta)
        wb2.save(ruta2)

        print("Archivos creados")

    else:
        print("Los archivos ya existen")