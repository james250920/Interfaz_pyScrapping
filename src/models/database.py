def obtener_datos_periodo(tipo: str) -> list:
    """Devuelve los datos estáticos de años o meses según el tipo solicitado.

    Args:
        tipo: Clave del conjunto de datos. Valores válidos: "anios", "meses".

    Returns:
        Lista de cadenas con los valores del período, o lista vacía si el tipo
        no es reconocido.
    """
    datos = {
        "anios": [
            "2021", "2022", "2023", "2024", "2025",
            "2026", "2027", "2028", "2029", "2030",
            "2031","2032","2033","2034", "2035", 
            "2036", "2037", "2038", "2039", "2040",
        ],
        "meses": [
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

    return datos.get(tipo, [])