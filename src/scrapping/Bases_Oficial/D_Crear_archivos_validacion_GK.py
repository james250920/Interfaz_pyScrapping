import os
import tempfile
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

        dir_ruta = os.path.dirname(ruta)
        os.makedirs(dir_ruta, exist_ok=True)
        
        fd1, temp1 = tempfile.mkstemp(suffix=".xlsx", dir=dir_ruta)
        os.close(fd1)
        wb.save(temp1)
        wb.close()
        os.replace(temp1, ruta)
        
        fd2, temp2 = tempfile.mkstemp(suffix=".xlsx", dir=dir_ruta)
        os.close(fd2)
        wb2.save(temp2)
        wb2.close()
        os.replace(temp2, ruta2)

        print("Archivos creados de forma segura.")

    else:
        print("Los archivos ya existen")