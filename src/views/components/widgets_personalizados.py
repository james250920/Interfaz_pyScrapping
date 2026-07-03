import os
import qtawesome as qta
from PySide6.QtWidgets import (
    QComboBox,
    QLabel,
    QProgressBar,
    QListView,
    QSizePolicy,
    QAbstractItemView,
    QWidget,
    QPushButton,
    QHBoxLayout,
    QVBoxLayout,
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QPixmap

from src.views.theme import *


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

    # Stylesheet — tarjeta blanca con borde sutil, texto oscuro legible
    combo.setStyleSheet(f"""
QComboBox {{
    background-color: {BLANCO};
    border: 1px solid {GRIS_BORDE};
    border-radius: 10px;
    padding: 8px 38px 8px 14px;
    color: {TEXTO};
    font-size: 13px;
    font-family: {FONT_FAMILY};
}}

QComboBox:hover {{
    border: 1px solid {ROJO};
    background: {BLANCO_SUAVE};
}}

QComboBox:focus {{
    border: 1.5px solid {ROJO};
    background: {BLANCO};
}}

QComboBox:disabled {{
    background: {GRIS_BG};
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
    border-top: 6px solid {TEXTO_SECUNDARIO};
    margin-right: 12px;
}}

QComboBox QAbstractItemView {{
    background: {BLANCO};
    color: {TEXTO};
    border: 1px solid {GRIS_BORDE};
    border-radius: 8px;
    outline: none;
    padding: 4px;
    selection-background-color: {ROJO};
    selection-color: {BLANCO};
}}

QComboBox QAbstractItemView::item {{
    height: 34px;
    padding-left: 12px;
    padding-right: 12px;
    border-radius: 6px;
    border: none;
}}

QComboBox QAbstractItemView::item:hover {{
    background: {GRIS_BG};
    color: {TEXTO};
}}

QComboBox QAbstractItemView::item:selected {{
    background: {ROJO};
    color: {BLANCO};
    font-weight: bold;
}}
""")

    return combo


def crear_imagen(ruta_base: str, src: str, width: int = 300, height: int = 230) -> QLabel:
    label = QLabel()
    label.setFixedSize(width, height)
    label.setAlignment(Qt.AlignCenter)

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
    barra.setFixedHeight(10)
    barra.setTextVisible(False)
    barra.setStyleSheet(f"""
        QProgressBar {{
            border: none;
            border-radius: 5px;
            background-color: {GRIS_BORDE};
        }}
        QProgressBar::chunk {{
            background-color: {ROJO};
            border-radius: 5px;
        }}
    """)
    return barra


# ══════════════════════════════════════════════════════════════════
# ICONOS (Font Awesome vía qtawesome) — reemplazo de emojis
# ══════════════════════════════════════════════════════════════════
# qtawesome dibuja los íconos como vectores y permite fijar su color
# exacto en hex, por lo que no dependen de la fuente emoji del SO
# (en Ubuntu/Linux los emojis suelen verse distinto o no renderizar).
# Instalar con: pip install qtawesome --break-system-packages

def crear_icono_label(nombre_icono: str, color: str, tam: int = 16) -> QLabel:
    """QLabel con un ícono de Font Awesome del color y tamaño indicados."""
    lbl = QLabel()
    icono = qta.icon(nombre_icono, color=color)
    lbl.setPixmap(icono.pixmap(QSize(tam, tam)))
    lbl.setFixedSize(tam, tam)
    return lbl


def crear_fila_icono_texto(nombre_icono: str, texto: str, color_icono: str,
                            color_texto: str, tam_icono: int = 14,
                            tam_fuente: int = 13, negrita: bool = True) -> QWidget:
    """Fila horizontal: ícono + texto (reemplaza patrones tipo '📁 Texto')."""
    fila = QWidget()
    fila.setStyleSheet("background: transparent;")
    layout = QHBoxLayout(fila)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(8)

    layout.addWidget(crear_icono_label(nombre_icono, color_icono, tam_icono))

    lbl_texto = QLabel(texto)
    peso = "700" if negrita else "400"
    lbl_texto.setStyleSheet(f"""
        color: {color_texto};
        font-size: {tam_fuente}px;
        font-weight: {peso};
        font-family: {FONT_FAMILY};
    """)
    layout.addWidget(lbl_texto)
    layout.addStretch()

    return fila


def crear_boton_icono(texto: str, nombre_icono: str, color_fondo: str,
                       color_texto: str = BLANCO, color_icono: str = None,
                       alto: int = 44, radio: int = 10,
                       color_hover: str = None, tam_icono: int = 14) -> QPushButton:
    """QPushButton con ícono Font Awesome + texto (ej. botón 'Examinar', 'Iniciar')."""
    boton = QPushButton(f"  {texto}")
    boton.setIcon(qta.icon(nombre_icono, color=color_icono or color_texto))
    boton.setIconSize(QSize(tam_icono, tam_icono))
    boton.setFixedHeight(alto)
    boton.setCursor(Qt.PointingHandCursor)
    boton.setStyleSheet(f"""
        QPushButton {{
            background-color: {color_fondo};
            color: {color_texto};
            border: none;
            border-radius: {radio}px;
            padding: 0 18px;
            font-size: 13px;
            font-weight: 700;
            font-family: {FONT_FAMILY};
            text-align: center;
        }}
        QPushButton:hover {{
            background-color: {color_hover or color_fondo};
        }}
        QPushButton:disabled {{
            background-color: {TEXTO_SECUNDARIO};
            color: rgba(255,255,255,0.6);
        }}
    """)
    return boton


class BotonIconoHover(QPushButton):
    """Botón pequeño (usado en la barra de título: minimizar, maximizar,
    cerrar) cuyo ícono cambia de color al pasar el mouse.
    Reemplaza los caracteres '─', '▢', '✕' del diseño original."""

    def __init__(self, nombre_icono: str, color_normal: str, color_hover: str,
                 fondo_hover: str, tam_icono: int = 12, tam_boton: int = 32, parent=None):
        super().__init__(parent)
        self._nombre_icono = nombre_icono
        self._color_normal = color_normal
        self._color_hover = color_hover
        self._tam_icono = tam_icono

        self.setFixedSize(tam_boton, tam_boton)
        self.setCursor(Qt.PointingHandCursor)
        self.setIconSize(QSize(tam_icono, tam_icono))
        self.setIcon(qta.icon(nombre_icono, color=color_normal))
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                border: none;
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background-color: {fondo_hover};
            }}
        """)

    def enterEvent(self, event):
        self.setIcon(qta.icon(self._nombre_icono, color=self._color_hover))
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.setIcon(qta.icon(self._nombre_icono, color=self._color_normal))
        super().leaveEvent(event)


def crear_item_check(texto: str) -> QWidget:
    """Fila con ícono de check en círculo rojo + texto blanco.
    Usado en la tarjeta 'Archivos a extraer' del panel derecho."""
    fila = QWidget()
    fila.setStyleSheet("background: transparent;")
    layout = QHBoxLayout(fila)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(10)

    circulo = QLabel()
    circulo.setFixedSize(20, 20)
    circulo.setAlignment(Qt.AlignCenter)
    circulo.setStyleSheet(f"background-color: {ROJO}; border-radius: 10px;")
    circulo.setPixmap(qta.icon("fa5s.check", color=BLANCO).pixmap(QSize(10, 10)))

    lbl = QLabel(texto)
    lbl.setStyleSheet(f"""
        color: {BLANCO};
        font-size: 13px;
        font-family: {FONT_FAMILY};
    """)
    lbl.setWordWrap(False)

    layout.addWidget(circulo)
    layout.addWidget(lbl, 1)

    return fila