import pandas as pd
import os
import locale
import asyncio
from openpyxl.styles import Font, PatternFill, Alignment


# -----------------------------
# CONFIG
# -----------------------------
def obtener_separador_formula():
    try:
        return ';' if locale.localeconv()['decimal_point'] == ',' else ','
    except:
        return ';'


# -----------------------------
# UTILIDADES EXCEL
# -----------------------------
def obtener_ultima_fila(ws, columna='A'):
    for fila in range(ws.max_row, 1, -1):
        if ws[f'{columna}{fila}'].value not in (None, ''):
            return fila
    return 1


def aplicar_estilo_encabezados(worksheet):
    relleno = PatternFill(fill_type='solid', start_color='C0C0C0', end_color='C0C0C0')
    fuente = Font(bold=True)
    alineacion = Alignment(horizontal='center', vertical='center')

    for celda in worksheet[1]:
        if celda.value:
            celda.fill = relleno
            celda.font = fuente
            celda.alignment = alineacion


def aplicar_formato_numerico(worksheet):
    rangos = [('F', 'R'), ('X', 'AL')]
    ultima_fila = worksheet.max_row

    for col_inicio, col_fin in rangos:
        for fila in range(2, ultima_fila + 1):
            for celda in worksheet[f'{col_inicio}{fila}:{col_fin}{fila}'][0]:
                if celda.value is not None:
                    celda.number_format = '#,##0.00'


# -----------------------------
# VALIDACIONES EXCEL
# -----------------------------
def agregar_validaciones(worksheet):
    sep = obtener_separador_formula()

    worksheet["AN1"] = "VALIDACION SIGLAS"
    worksheet["AO1"] = "VALIDACION RUBROS"

    # FIX importante COUNTIF
    worksheet["AP1"] = f'=COUNTIF(AN:AN{sep}"FALSE")'
    worksheet["AQ1"] = f'=COUNTIF(AO:AO{sep}"FALSE")'

    ultima_fila = obtener_ultima_fila(worksheet, 'A')

    for fila in range(2, ultima_fila + 1):
        worksheet[f"AN{fila}"] = f"=A{fila}=T{fila}"
        worksheet[f"AO{fila}"] = f"=C{fila}=U{fila}"


# -----------------------------
# PROCESO PRINCIPAL
# -----------------------------
async def consolidar_gk_presupuesto(ruta_principal):

    ruta_ejecucion = rf"{ruta_principal}\EJECUCION\Gastos_Capital_Ejecucion_Presupuesto.xlsx"
    ruta_marco = rf"{ruta_principal}\MARCO\Gastos_Capital_Formulacion_Presupuesto.xlsx"
    ruta_destino = rf"{ruta_principal}\VALIDACION GASTO CAPITAL\Validación_Gastos_Capital_Presupuesto.xlsx"

    try:
        print("Leyendo archivos...")

        df_marco = pd.read_excel(ruta_marco, sheet_name=1)
        df_ejecucion = pd.read_excel(ruta_ejecucion, sheet_name=0)

        df_marco = df_marco.dropna(how='all').dropna(axis=1, how='all')
        df_ejecucion = df_ejecucion.dropna(how='all').dropna(axis=1, how='all')

        # -----------------------------
        # REGLAS
        # -----------------------------
        reglas = {
            "PROGRAMA DE INVERSIONES": ("1", "PGI"),
            "PROYECTOS DE INVERSION": ("2", "PI"),
            "GASTOS DE CAPITAL NO LIGADOS A PROYECTOS": ("3", "NL"),
            "INVERSION FINANCIERA": ("4", "IF"),
            "OTROS": ("5", "OT"),
            "TOTAL GASTOS DE CAPITAL": ("6", "GC")
        }

        codigos = []
        codigo_actual = ""
        ultima_regla_num = 0
        ultima_regla_texto = None

        for valor in df_marco.get("NOMBRE_GASTO", []):
            texto = str(valor).strip().upper()

            if texto in reglas:
                numero, abrev = reglas[texto]
                numero_int = int(numero)

                esperado = 1 if ultima_regla_num == 6 else ultima_regla_num + 1

                if numero_int == esperado:

                    if texto == ultima_regla_texto:
                        codigos.append(abrev)
                    else:
                        codigos.append(numero)

                    codigo_actual = abrev

                    if texto != ultima_regla_texto:
                        ultima_regla_num = numero_int
                else:
                    codigos.append(codigo_actual)

                ultima_regla_texto = texto
            else:
                codigos.append(codigo_actual)
                ultima_regla_texto = None

        # -----------------------------
        # INSERTAR CODIGO
        # -----------------------------
        if "SIGLAS" in df_marco.columns:
            pos = df_marco.columns.get_loc("SIGLAS") + 1
        else:
            pos = 0

        if len(codigos) == len(df_marco):
            df_marco.insert(pos, "CODIGO", codigos)
        else:
            df_marco["CODIGO"] = codigos[:len(df_marco)]

        # -----------------------------
        # COLUMNAS
        # -----------------------------
        columnas = [
            'SIGLAS',
            'CODIGO',
            'NOMBRE_GASTO',
            'TIPO_PROYECTO',
            'CODIGO_UNICO',
            'COSTO_INVERSION'
        ]

        columnas_ejec = [
            'SIGLAS',
            'NOMBRE_GASTO',
            'TIPO_PROYECTO',
            'CODIGO_UNICO',
            'COSTO_INVERSION',
            'MONTO_INICIAL_AL_MES_CORTE',
            'EJEC_AÑO_ANTERIOR'
        ]

        columnas_montos = [f'MONTO{i}' for i in range(1, 13)]

        columnas_final_marco = columnas + columnas_montos
        columnas_final_ejec = columnas_ejec + columnas_montos

        # -----------------------------
        # FILTRADO SEGURO
        # -----------------------------
        df_marco = df_marco[[c for c in columnas_final_marco if c in df_marco.columns]]
        df_ejecucion = df_ejecucion[[c for c in columnas_final_ejec if c in df_ejecucion.columns]]

        # -----------------------------
        # ALINEAR FILAS
        # -----------------------------
        max_len = max(len(df_marco), len(df_ejecucion))

        df_marco = df_marco.reindex(range(max_len))
        df_ejecucion = df_ejecucion.reindex(range(max_len))

        espacio = pd.DataFrame({' ': [''] * max_len})

        df_final = pd.concat(
            [df_marco.reset_index(drop=True),
             espacio,
             df_ejecucion.reset_index(drop=True)],
            axis=1
        )

        # -----------------------------
        # EXPORTAR
        # -----------------------------
        os.makedirs(os.path.dirname(ruta_destino), exist_ok=True)

        temp_file = ruta_destino.replace(".xlsx", "_temp.xlsx")
        with pd.ExcelWriter(temp_file, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Validacion_Comparativa', index=False)

            ws = writer.sheets['Validacion_Comparativa']

            agregar_validaciones(ws)
            aplicar_formato_numerico(ws)
            aplicar_estilo_encabezados(ws)
            
        os.replace(temp_file, ruta_destino)
        print(f"OK: Archivo generado en:\n{ruta_destino}")

    except KeyError as e:
        print(f"ERROR columna faltante: {e}")

    except PermissionError:
        print("ERROR: Cierra el archivo antes de ejecutar.")

    except Exception as e:
        print(f"Error inesperado: {e}")