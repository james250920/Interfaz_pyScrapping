import os
from PySide6.QtWidgets import QComboBox, QLabel, QProgressBar, QListView
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap

from src.views.theme import (AZUL, GRIS_BORDE, GRIS_BG, FONT_FAMILY, 
                             DORADO, DORADO_HOVER, OSCURO_CARD, OSCURO_2, GRIS_TEXTO, RADIUS_MD)

def crear_menubar(placeholder: str, datos: list, on_change=None, default_val=None) -> QComboBox:
    combo = QComboBox()
    combo.addItems(datos)
    combo.setFixedWidth(148)
    combo.setFixedHeight(44)
    
    if default_val:
        combo.setCurrentText(str(default_val))
    
    # ListView customizado para el dropdown
    list_view = QListView()
    list_view.setStyleSheet(f"""
        QListView {{
            background-color: {OSCURO_CARD};
            color: white;
            border: 1px solid {DORADO};
            border-radius: {RADIUS_MD};
            padding: 4px;
            outline: 0;
        }}
        QListView::item {{
            padding: 8px;
            border-radius: 4px;
        }}
        QListView::item:hover {{
            background-color: {DORADO_HOVER};
            color: {OSCURO_CARD};
        }}
        QListView::item:selected {{
            background-color: {DORADO};
            color: {OSCURO_CARD};
        }}
    """)
    combo.setView(list_view)
    
    # Estilo CSS premium
    combo.setStyleSheet(f"""
        QComboBox {{
            border: 1px solid {GRIS_BORDE};
            border-radius: {RADIUS_MD};
            background-color: {GRIS_BG};
            padding-left: 14px;
            color: #333333;
            font-size: 13px;
            font-family: {FONT_FAMILY};
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 30px;
            border-left-width: 0px;
        }}
        QComboBox::down-arrow {{
            image: none; /* Podríamos usar un SVG aquí si hubiera */
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 5px solid {GRIS_TEXTO};
            margin-right: 14px;
        }}
        QComboBox:hover {{
            border: 1px solid {DORADO};
        }}
        QComboBox:focus {{
            border: 1px solid {DORADO};
            background-color: white;
        }}
        QComboBox QAbstractItemView {{
            border: 1px solid {DORADO};
            border-radius: {RADIUS_MD};
            background-color: {OSCURO_CARD};
            selection-background-color: {DORADO};
        }}
    """)
    
    if on_change:
        combo.currentTextChanged.connect(on_change)
        
    return combo


def crear_imagen(ruta_base: str, src: str, width: int = 300, height: int = 230) -> QLabel:
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


def crear_barra_progreso() -> QProgressBar:
    barra = QProgressBar()
    barra.setFixedHeight(12)
    barra.setTextVisible(False)
    barra.setStyleSheet(f"""
        QProgressBar {{
            border: none;
            border-radius: 6px;
            background-color: {GRIS_BG};
        }}
        QProgressBar::chunk {{
            background-color: {DORADO};
            border-radius: 6px;
        }}
    """)
    return barra