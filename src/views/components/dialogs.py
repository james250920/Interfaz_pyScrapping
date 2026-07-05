import os
from PySide6.QtWidgets import QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QPixmap

from src.views.theme import *


class StyledDialog(QDialog):
    def __init__(self, parent=None, title="", message="", tipo="info", icon_path=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedSize(400, 200)
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        # Definir color según tipo
        color_tipo = AZUL_INFO
        if tipo == "success":
            color_tipo = VERDE_EXITO
        elif tipo == "error":
            color_tipo = ROJO
        elif tipo == "warning":
            color_tipo = NARANJA_ADVERTENCIA

        self.init_ui(title, message, color_tipo, icon_path)

    def init_ui(self, title, message, color, icon_path):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Contenedor principal
        self.container = QDialog(self)
        self.container.setStyleSheet(f"""
            QDialog {{
                background-color: {BLANCO};
                border: 1px solid {color};
                border-radius: 14px;
            }}
        """)
        container_layout = QVBoxLayout(self.container)
        container_layout.setContentsMargins(24, 24, 24, 24)
        container_layout.setSpacing(16)

        # Header (Icono + Título)
        header_layout = QHBoxLayout()

        if icon_path and os.path.exists(icon_path):
            icon_label = QLabel()
            pixmap = QPixmap(icon_path).scaled(24, 24, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            icon_label.setPixmap(pixmap)
            header_layout.addWidget(icon_label)

        title_label = QLabel(title)
        title_label.setStyleSheet(f"""
            color: {color};
            font-family: {FONT_FAMILY_TITLE};
            font-size: 16px;
            font-weight: bold;
        """)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        container_layout.addLayout(header_layout)

        # Mensaje
        msg_label = QLabel(message)
        msg_label.setWordWrap(True)
        msg_label.setStyleSheet(f"""
            color: {TEXTO};
            font-family: {FONT_FAMILY};
            font-size: 13px;
        """)
        container_layout.addWidget(msg_label)

        container_layout.addStretch()

        # Botones
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        btn_ok = QPushButton("Aceptar")
        btn_ok.setFixedSize(100, 36)
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: {BLANCO};
                border: none;
                border-radius: 6px;
                font-family: {FONT_FAMILY_TITLE};
                font-size: 13px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {ROJO_HOVER if color == ROJO else color};
            }}
        """)
        btn_ok.clicked.connect(self.accept)
        btn_layout.addWidget(btn_ok)

        container_layout.addLayout(btn_layout)

        layout.addWidget(self.container)

    def showEvent(self, event):
        super().showEvent(event)
        self._centrar_dialogo()

    def _centrar_dialogo(self):
        destino = None

        parent = self.parentWidget()
        if parent is not None:
            destino = parent.frameGeometry().center()
        else:
            pantalla = QApplication.primaryScreen()
            if pantalla is not None:
                destino = pantalla.availableGeometry().center()

        if destino is None:
            return

        geometria = self.frameGeometry()
        geometria.moveCenter(destino)
        self.move(geometria.topLeft())


def show_info(parent, title, message):
    dlg = StyledDialog(parent, title, message, "info")
    dlg.exec()


def show_success(parent, title, message):
    dlg = StyledDialog(parent, title, message, "success")
    dlg.exec()


def show_error(parent, title, message):
    dlg = StyledDialog(parent, title, message, "error")
    dlg.exec()


def show_warning(parent, title, message):
    dlg = StyledDialog(parent, title, message, "warning")
    dlg.exec()