import os
from PySide6.QtWidgets import (QMainWindow, QWidget, QLabel, QLineEdit, 
                             QPushButton, QHBoxLayout, QVBoxLayout, QMessageBox)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QIcon

# Importamos componentes, modelos y tema centralizado
from src.views.components.widgets_personalizados import crear_menubar
from src.views.theme import OSCURO, GRIS_BG, GRIS_BORDE, AZUL_PRIMARIO, AZUL_HOVER, GRIS_DESHABILITADO
from src.models.database import obtener_datos_periodo
from src.scrapping.scrapping_main import scrapping_main
import datetime


class ScrappingWorker(QThread):
    """Hilo de fondo para ejecutar el scrapping sin congelar la UI."""
    finished = Signal()
    error = Signal(str)
    progress = Signal(str)

    def __init__(self, directorio, anio, mes, fecha_cierre):
        super().__init__()
        self.directorio = directorio
        self.anio = anio
        self.mes = mes
        self.fecha_cierre = fecha_cierre

    def run(self):
        try:
            scrapping_main(
                self.directorio, self.anio, self.mes, self.fecha_cierre,
                on_progreso=self.progress.emit,
            )
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self, ruta_raiz):
        super().__init__()
        self.ruta_raiz = ruta_raiz # Guardamos la ruta del proyecto
        self._worker = None  # Referencia al hilo de scrapping
        self.setWindowTitle("Sistema de extracción")
        self.setFixedSize(900, 540)
        
        # Configurar el icono de la ventana
        ruta_icono = os.path.join(self.ruta_raiz, "src", "assets", "icons", "icon.svg")
        self.setWindowIcon(QIcon(ruta_icono))
        
        # Carga de datos usando el modelo
        self.datos_anos = obtener_datos_periodo("anios")
        self.datos_meses = obtener_datos_periodo("meses")
        
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

        self.boton_iniciar = QPushButton("  Iniciar")
        self.boton_iniciar.setFixedSize(160, 46)
        self.boton_iniciar.setCursor(Qt.PointingHandCursor)
        self.boton_iniciar.setStyleSheet(f"""
            QPushButton {{
                background-color: {AZUL_PRIMARIO};
                color: white;
                border-radius: 10px;
                font-size: 14px;
                font-weight: 600;
                font-family: 'Sora', 'Segoe UI';
            }}
            QPushButton:hover {{ background-color: {AZUL_HOVER}; }}
            QPushButton:disabled {{ background-color: {GRIS_DESHABILITADO}; }}
        """)
        self.boton_iniciar.clicked.connect(self.on_click_iniciar)

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
        layout_izquierdo.addWidget(self.boton_iniciar, 0, Qt.AlignLeft)
        layout_izquierdo.addStretch()

        # ── PANEL DERECHO (Oscuro) ───────────────────────────────────────
        panel_derecho = QWidget()
        panel_derecho.setFixedWidth(340)
        panel_derecho.setObjectName("panel_derecho")
        path_fondo = os.path.join(self.ruta_raiz, "src", "assets", "images", "fondo.png").replace("\\", "/")
        panel_derecho.setStyleSheet(f"""
            QWidget#panel_derecho {{
                border-image: url("{path_fondo}") 0 0 0 0 stretch stretch;
                background-color: {OSCURO};
            }}
        """)
        layout_derecho = QVBoxLayout(panel_derecho)
        layout_derecho.setContentsMargins(24, 24, 24, 24)



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
        # Extraer solo el número del mes: "5, MAYO" → "5"
        mes_texto = self.combo_mes.currentText()
        mes = mes_texto.split(",")[0].strip()
        fecha_cierre_sistema = datetime.datetime.now().strftime("%d.%m.%Y")

        print(f"Iniciando con directorio: {directorio}, año: {anio}, mes: {mes}")

        # Deshabilitar el botón para evitar doble ejecución
        self.boton_iniciar.setEnabled(False)
        self.boton_iniciar.setText("  Ejecutando...")

        # Ejecutar scrapping en hilo de fondo para no congelar la UI
        self._worker = ScrappingWorker(directorio, anio, mes, fecha_cierre_sistema)
        self._worker.finished.connect(self._on_scrapping_finished)
        self._worker.error.connect(self._on_scrapping_error)
        self._worker.progress.connect(self._on_scrapping_progress)
        self._worker.start()

    def _on_scrapping_finished(self):
        """Se ejecuta cuando el scrapping termina exitosamente."""
        self.boton_iniciar.setEnabled(True)
        self.boton_iniciar.setText("  Iniciar")
        QMessageBox.information(self, "Completado", "El proceso de scrapping finalizó correctamente.")

    def _on_scrapping_error(self, mensaje):
        """Se ejecuta cuando el scrapping falla."""
        self.boton_iniciar.setEnabled(True)
        self.boton_iniciar.setText("  Iniciar")
        QMessageBox.critical(self, "Error", f"Error durante el scrapping:\n{mensaje}")

    def _on_scrapping_progress(self, mensaje):
        """Actualiza el texto del botón con el paso actual del proceso."""
        self.boton_iniciar.setText(f"  {mensaje}")