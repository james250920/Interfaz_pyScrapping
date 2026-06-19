import os
import sys
from openpyxl import load_workbook

def obtener_ultima_fila(ws, columna='A', fila_minima=5):
    for fila in range(ws.max_row, fila_minima - 1, -1):
        valor = ws[f"{columna}{fila}"].value
        if valor is not None and str(valor).strip() != '':
            return fila
    return fila_minima

def formato_deposiciones_colocaciones(ruta_principal):
    ruta_destino = rf"{ruta_principal}\Base FONAFE WEB al mes.xlsm"

    if not os.path.exists(ruta_destino):
        print("No existe archivo destino.")
        return

    wb = None
    try:
        print("Abriendo archivo (esto puede demorar por las macros)...")
        wb = load_workbook(ruta_destino, keep_vba=True)
        ws = wb["Depósitos y Colocaciones"]

        ultima_fila = obtener_ultima_fila(ws, columna='A', fila_minima=5)
        print(f"Última fila detectada: {ultima_fila}")

        print("Aplicando formatos...")
        for fila in range(5, ultima_fila + 1):
            
            for col_letra in ['H', 'I', 'J']:
                ws[f"{col_letra}{fila}"].number_format = '#,##0.00'
            
            ws[f"L{fila}"].number_format = 'dd/mm/yyyy'

        print("Guardando archivo...")
        print(">>> ATENCIÓN: El guardado puede demorar unos minutos. Por favor, espera... <<<")
        wb.save(ruta_destino)
        
        print(f"\n¡Formatos aplicados correctamente hasta la fila {ultima_fila}!")

    except PermissionError:
        print("\nERROR: El archivo Excel está abierto. Ciérrelo antes de ejecutar el programa.")
    except Exception as e:
        print(f"\nERROR Inesperado: {e}")
    finally:
        print("Cerrando archivo en memoria...")
        if wb:
            wb.close()
        
        print("Finalizando proceso...")

