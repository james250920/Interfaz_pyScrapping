import os
from PySide6.QtWidgets import QComboBox, QLabel
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap

AZUL = "#115b82"
GRIS_BORDE = "#e2e2e2"
GRIS_BG = "#f8f8f8"

def crear_menubar(placeholder: str, datos: list, on_change=None) -> QComboBox:
    combo = QComboBox()
    
    # En Qt, para tener un "Label/Placeholder" flotante o inicial, agregamos un elemento vacío 
    # o usamos un truco de estilo. Lo más limpio para mantener tu lógica es añadir los datos:
    combo.addItems(datos)
    combo.setFixedWidth(148)
    combo.setFixedHeight(40)
    
    # Estilo CSS para replicar Flet
    combo.setStyleSheet(f"""
        QComboBox {{
            border: 1px solid {GRIS_BORDE};
            border-radius: 10px;
            background-color: white;
            padding-left: 14px;
            color: #333333;
            font-size: 13px;
        }}
        QComboBox::drop-down {{
            border: 0px;
        }}
        QComboBox:focus {{
            border: 1px solid {AZUL};
        }}
    """)
    
    if on_change:
        combo.currentTextChanged.connect(on_change)
        
    return combo


def crear_imagen(ruta_base: str, src: str, width: int = 240, height: int = 180) -> QLabel:
    label = QLabel()
    label.setFixedSize(width, height)
    label.setAlignment(Qt.AlignCenter)
    
    # Construimos la ruta hacia la carpeta assets
    ruta_img = os.path.join(ruta_base, "src", "assets", "images", src)
    
    if os.path.exists(ruta_img):
        pixmap = QPixmap(ruta_img).scaled(width, height, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        label.setPixmap(pixmap)
    else:
        label.setText(f"[ {src} no encontrada ]")
        label.setStyleSheet("color: white; font-size: 11px;")
        
    return label