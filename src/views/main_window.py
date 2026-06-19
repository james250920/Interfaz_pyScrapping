import os
from PySide6.QtWidgets import (QMainWindow, QWidget, QLabel, QLineEdit, 
                             QPushButton, QHBoxLayout, QVBoxLayout, QMessageBox)
from PySide6.QtCore import Qt

# Importamos tus componentes y modelos adaptados
from src.views.components.widgets_personalizados import crear_menubar, crear_imagen
from src.models.database import cargar_datos_csv
from src.scrapping.ScrappingMain import scrapping_main
import datetime 


OSCURO = "#1a1a2e"
GRIS_BG = "#f8f8f8"
GRIS_BORDE = "#e2e2e2"

class MainWindow(QMainWindow):
    def __init__(self, ruta_raiz):
        super().__init__()
        self.ruta_raiz = ruta_raiz # Guardamos la ruta del proyecto
        self.setWindowTitle("Sistema de extracción")
        self.setFixedSize(900, 540)
        
        # Carga de datos usando tu modelo
        self.datos_anos = cargar_datos_csv(os.path.join(self.ruta_raiz, "listAnio.txt"))
        self.datos_meses = cargar_datos_csv(os.path.join(self.ruta_raiz, "lisMes.txt"))
        
        self.init_ui()

    def init_ui(self):
        widget_central = QWidget()
        layout_principal = QHBoxLayout(widget_central)
        layout_principal.setContentsMargins(0, 0, 0, 0)
        layout_principal.setSpacing(0)

        # ── PANEL IZQUIERDO ──────────────────────────────────────────────
        panel_izquierdo = QWidget()
        panel_izquierdo.setStyleSheet("background-color: white;")
        layout_izquierdo = QVBoxLayout(panel_izquierdo)
        layout_izquierdo.setContentsMargins(40, 36, 40, 36)
        layout_izquierdo.setSpacing(0)

        etiqueta_config = QLabel("CONFIGURACIÓN DE CARGA")
        etiqueta_config.setStyleSheet("color: rgba(0,0,0,0.45); font-size: 10px; font-weight: 600; font-family: 'Sora', 'Segoe UI';")

        etiqueta_dir = QLabel("Directorio de trabajo")
        etiqueta_dir.setStyleSheet("color: rgba(0,0,0,0.55); font-size: 11px; font-weight: 500;")

        self.campo_ruta = QLineEdit()
        self.campo_ruta.setPlaceholderText("Selecciona o escribe un directorio...")
        self.campo_ruta.setFixedHeight(44)
        self.campo_ruta.setStyleSheet(f"""
            QLineEdit {{
                border: 1px solid {GRIS_BORDE};
                border-radius: 10px;
                background-color: {GRIS_BG};
                padding-left: 14px;
                color: black;
                font-size: 13px;
            }}
        """)

        etiqueta_periodo = QLabel("Período de análisis")
        etiqueta_periodo.setStyleSheet("color: rgba(0,0,0,0.55); font-size: 11px; font-weight: 500;")

        # Inyección de tus Menubars reutilizables
        self.combo_ano = crear_menubar("Año", self.datos_anos, on_change=self.on_change_ano)
        self.combo_mes = crear_menubar("Mes", self.datos_meses, on_change=self.on_change_mes)

        layout_combos = QHBoxLayout()
        layout_combos.addWidget(self.combo_ano)
        layout_combos.addWidget(self.combo_mes)
        layout_combos.setSpacing(14)

        boton_iniciar = QPushButton("  Iniciar")
        boton_iniciar.setFixedSize(160, 46)
        boton_iniciar.setCursor(Qt.PointingHandCursor)
        boton_iniciar.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border-radius: 10px;
                font-size: 14px;
                font-weight: 600;
                font-family: 'Sora', 'Segoe UI';
            }
            QPushButton:hover { background-color: #006abc; }
        """)
        boton_iniciar.clicked.connect(self.on_click_iniciar)

        # Construcción del Layout Izquierdo
        layout_izquierdo.addWidget(etiqueta_config)
        layout_izquierdo.addSpacing(16)
        layout_izquierdo.addWidget(etiqueta_dir)
        layout_izquierdo.addSpacing(6)
        layout_izquierdo.addWidget(self.campo_ruta)
        layout_izquierdo.addSpacing(16)
        layout_izquierdo.addWidget(etiqueta_periodo)
        layout_izquierdo.addSpacing(6)
        layout_izquierdo.addLayout(layout_combos)
        layout_izquierdo.addSpacing(20)
        layout_izquierdo.addWidget(boton_iniciar, 0, Qt.AlignLeft)
        layout_izquierdo.addStretch()

        # ── PANEL DERECHO (Oscuro) ───────────────────────────────────────
        panel_derecho = QWidget()
        panel_derecho.setFixedWidth(340)
        panel_derecho.setStyleSheet(f"background-color: {OSCURO};")
        layout_derecho = QVBoxLayout(panel_derecho)
        layout_derecho.setContentsMargins(24, 24, 24, 24)

        # Inyección de tu componente Imagen reutilizable
        imagen_libros = crear_imagen(self.ruta_raiz, "Img.png", width=240, height=180)
        layout_derecho.addWidget(imagen_libros, 0, Qt.AlignCenter)

        # Ensamblado
        layout_principal.addWidget(panel_izquierdo)
        layout_principal.addWidget(panel_derecho)
        self.setCentralWidget(widget_central)

    # ── Métodos de Eventos ───────────────────────────────────────────
    def on_change_ano(self, valor):
        print(f"Año seleccionado: {valor}")

    def on_change_mes(self, valor):
        print(f"Mes seleccionado: {valor}")

    def on_click_iniciar(self):
        directorio = self.campo_ruta.text().strip()
        if not directorio:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)
            msg.setText("Por favor selecciona un directorio")
            msg.setWindowTitle("Advertencia")
            msg.exec()
            return
        anio = self.combo_ano.currentText()
        mes = self.combo_mes.currentText()
        fecha_cierre_sistema = datetime.datetime.now().strftime("%d.%m.%Y")
        print(f"Iniciando con directorio: {directorio}")
        scrapping_main(directorio, anio, mes, fecha_cierre_sistema)