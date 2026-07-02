import os
from PySide6.QtWidgets import (
    QComboBox,
    QLabel,
    QProgressBar,
    QListView,
    QSizePolicy,
    QAbstractItemView,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap

from src.views.theme import (AZUL, GRIS_BORDE, GRIS_BG, FONT_FAMILY, 
                             DORADO, DORADO_HOVER, OSCURO_CARD, OSCURO_2, GRIS_TEXTO, RADIUS_MD)

def crear_menubar(placeholder: str, datos: list, on_change=None, default_val=None) -> QComboBox:
    combo = QComboBox()

    # Datos
    combo.addItems(datos)

    # Tamaño
    combo.setMinimumHeight(44)
    combo.setMaxVisibleItems(6)

    # Evita problemas de tamaño del popup
    combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

    # Mucho más estable que AdjustToContents
    combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)

    # Vista personalizada
    view = QListView()

    view.setUniformItemSizes(True)
    view.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
    view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    view.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

    combo.setView(view)

    # Valor por defecto
    if default_val is not None:
        combo.setCurrentText(str(default_val))

    # Evento
    if on_change:
        combo.currentTextChanged.connect(on_change)

    # Tu stylesheet
    combo.setStyleSheet(f"""
QComboBox {{
    background-color: #FFFFFF;
    border: 1px solid {GRIS_BORDE};
    border-radius: 10px;
    padding: 8px 38px 8px 12px;
    color: #2F2F2F;
    font-size: 13px;
    font-family: {FONT_FAMILY};
}}

QComboBox:hover {{
    border: 1px solid {DORADO};
    background: #FCFCFC;
}}

QComboBox:focus {{
    border: 2px solid {DORADO};
    background: white;
}}

QComboBox:disabled {{
    background: #F4F4F4;
    color: #A0A0A0;
}}

QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 32px;
    border: none;
    background: transparent;
}}

QComboBox::down-arrow {{
    image: none;
    width: 0px;
    height: 0px;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #666666;
    margin-right: 10px;
}}

QComboBox QAbstractItemView {{
    background: {OSCURO_CARD};
    color: white;
    border: 1px solid {DORADO};
    border-radius: 8px;
    outline: none;
    selection-background-color: {DORADO};
    selection-color: {OSCURO_CARD};
}}

QComboBox QAbstractItemView::item {{
    height: 34px;
    padding-left: 12px;
    padding-right: 12px;
    border: none;
}}

QComboBox QAbstractItemView::item:hover {{
    background: {DORADO_HOVER};
    color: {OSCURO_CARD};
}}

QComboBox QAbstractItemView::item:selected {{
    background: {DORADO};
    color: {OSCURO_CARD};
    font-weight: bold;
}}
""")

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