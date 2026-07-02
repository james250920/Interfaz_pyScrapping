import pandas as pd
import os
import locale
import asyncio
from openpyxl.styles import Font, PatternFill, Alignment


# -----------------------------
# CONFIG EXCEL
# -----------------------------
def obtener_separador_formula():
    try:
        configuracion = locale.localeconv()
        return ';' if configuracion['decimal_point'] == ',' else ','
    except:
        return ';'


# -----------------------------
# UTILIDADES EXCEL
# -----------------------------
def obtener_ultima_fila(ws, columna='A'):
    for fila in range(ws.max_row, 1, -1):
        valor = ws[f'{columna}{fila}'].value
        if valor not in (None, ''):
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
    ultima_fila = obtener_ultima_fila(worksheet, 'A')

    for col_inicio, col_fin in rangos:
        for fila in range(2, ultima_fila + 1):
            for celda in worksheet[f'{col_inicio}{fila}:{col_fin}{fila}'][0]:
                if celda.value is not None:
                    celda.number_format = '#,##0.00'


# -----------------------------
# VALIDACIONES
# -----------------------------
def agregar_validaciones(worksheet):
    sep = obtener_separador_formula()

    worksheet["AN1"] = "VALIDACION SIGLAS"
    worksheet["AO1"] = "VALIDACION RUBROS"

    # ✔ CORREGIDO COUNTIF
    worksheet["AP1"] = f'=COUNTIF(AN:AN{sep}"FALSE")'
    worksheet["AQ1"] = f'=COUNTIF(AO:AO{sep}"FALSE")'

    ultima_fila = obtener_ultima_fila(worksheet, 'A')

    for fila in range(2, ultima_fila + 1):
        worksheet[f"AN{fila}"] = f"=A{fila}=T{fila}"
        worksheet[f"AO{fila}"] = f"=C{fila}=U{fila}"


# -----------------------------
# PROCESO PRINCIPAL
# -----------------------------
async def consolidar_gk_flujo_caja(ruta_principal):

    ruta_ejecucion = rf"{ruta_principal}\EJECUCION\Gastos_Capital_Ejecucion_Flujo_Caja.xlsx"
    ruta_marco = rf"{ruta_principal}\MARCO\Gastos_Capital_Formulacion_Flujo_Caja.xlsx"
    ruta_destino = rf"{ruta_principal}\VALIDACION GASTO CAPITAL\Validacion_Gastos_Capital_Flujo_Caja.xlsx"

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
                    codigos.append(abrev if texto != ultima_regla_texto else numero)
                    codigo_actual = abrev
                    ultima_regla_num = numero_int if texto != ultima_regla_texto else ultima_regla_num
                else:
                    codigos.append(codigo_actual)

                ultima_regla_texto = texto
            else:
                codigos.append(codigo_actual)

        # -----------------------------
        # INSERTAR COLUMNA CODIGO
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
        columnas_base = [
            'SIGLAS',
            'CODIGO',
            'NOMBRE_GASTO',
            'TIPO_PROYECTO',
            'CODIGO_UNICO'
        ]

        columnas_montos = [f'MONTO{i}' for i in range(1, 13)]
        columnas_finales = columnas_base + columnas_montos

        df_marco = df_marco[[c for c in columnas_finales if c in df_marco.columns]]
        df_ejecucion = df_ejecucion[[c for c in columnas_finales if c != "CODIGO" and c in df_ejecucion.columns]]

        # -----------------------------
        # ALINEAR TAMAÑO
        # -----------------------------
        max_len = max(len(df_marco), len(df_ejecucion))

        df_marco = df_marco.reindex(range(max_len))
        df_ejecucion = df_ejecucion.reindex(range(max_len))

        espacio = pd.DataFrame({' ': [''] * max_len})

        df_final = pd.concat(
            [df_marco, espacio, espacio, df_ejecucion],
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
        print(f"OK: Archivo generado en {ruta_destino}")

    except KeyError as e:
        print(f"Falta columna: {e}")

    except PermissionError:
        print("Cierra el archivo de salida antes de ejecutar.")

    except Exception as e:
        print(f"Error inesperado: {e}")