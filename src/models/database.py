import os

def cargar_datos_csv(ruta_archivo: str) -> list:
    datos_por_archivo = {
        "listAnio.txt": [
            "2021", "2022", "2023", "2024", "2025", "2026", "2027", "2028", "2029", "2030"
        ],
        "lisMes.txt": [
            "1, ENERO",
            "2, FEBRERO",
            "3, MARZO",
            "4, ABRIL",
            "5, MAYO",
            "6, JUNIO",
            "7, JULIO",
            "8, AGOSTO",
            "9, SETIEMBRE",
            "10, OCTUBRE",
            "11, NOVIEMBRE",
            "12, DICIEMBRE",
            "13, CIERRE",
            "14, AUDITORIA",
            "16, ADECUACION NIIF",
        ],
    }

    nombre_archivo = os.path.basename(ruta_archivo)
    return datos_por_archivo.get(nombre_archivo, [])