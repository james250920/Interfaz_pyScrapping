import os
import asyncio
from concurrent.futures import ThreadPoolExecutor
from openpyxl import load_workbook

# Convertida a función principal asíncrona con nombre único
async def limpiar_hojas_excel(ruta_principal):
    ruta_archivo = os.path.join(ruta_principal, "Base FONAFE WEB al mes.xlsm")
    
    # Subfunción síncrona interna que correrá aislada en el Pool de Hilos
    def _limpiar_sync():
        try:
            print("Abriendo archivo Excel en segundo plano...")
            wb = load_workbook(ruta_archivo, keep_vba=True)

            hojas_a_limpiar = ["Depósitos y Colocaciones"]
            rango_borrar = "A5:L6000"

            for nombre_hoja in hojas_a_limpiar:
                try:
                    if nombre_hoja not in wb.sheetnames:
                        print(f"  ✗ La hoja '{nombre_hoja}' no existe.")
                        continue

                    ws = wb[nombre_hoja]
                    
                    # Obtener el estilo base de A1 una sola vez fuera de los bucles para mejorar rendimiento
                    estilo_base = ws['A1']._style

                    # Limpieza masiva e iterativa celda por celda
                    for fila in ws[rango_borrar]:
                        for celda in fila:
                            celda.value = None
                            celda._style = estilo_base
                            celda.comment = None
                            celda.hyperlink = None

                    print(f"  ✓ Rango {rango_borrar} borrado a fondo en hoja: {nombre_hoja}")

                except Exception as e_hoja:
                    print(f"  ⚠ No se pudo limpiar '{nombre_hoja}'. Error: {e_hoja}")

            print("Guardando archivo en el disco...")
            wb.save(ruta_archivo)
            wb.close()  # Cierre explícito para liberar la RAM de las 72,000 celdas
            print("✓ Archivo guardado correctamente sin bloquear el flujo principal.")

        except PermissionError:
            print("✗ Error: El archivo está abierto. Ciérrelo en Excel de escritorio antes de ejecutar.")
        except Exception as e:
            print(f"✗ Error crítico en el hilo de openpyxl: {e}")

    # --- COORDINACIÓN DEL LOOP ASÍNCRONO ---
    loop = asyncio.get_running_loop()
    
    with ThreadPoolExecutor(max_workers=1) as executor:
        await loop.run_in_executor(executor, _limpiar_sync)
        
    print("Proceso de desinfección y formateo openpyxl finalizado.")


