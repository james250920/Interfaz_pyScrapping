import os
import datetime
from PySide6.QtWidgets import (QMainWindow, QWidget, QLabel, QLineEdit, 
                             QPushButton, QHBoxLayout, QVBoxLayout, QFileDialog)
from PySide6.QtCore import Qt, QThread, Signal, QPoint
from PySide6.QtGui import QIcon, QMouseEvent, QPixmap

# Importamos componentes, modelos y tema centralizado
from src.views.components.widgets_personalizados import crear_menubar, crear_imagen, crear_barra_progreso
from src.views.components.dialogs import show_success, show_error, show_warning, show_info
from src.views.theme import (OSCURO, OSCURO_2, GRIS_BG, GRIS_BORDE, GRIS_DESHABILITADO,
                            DORADO, DORADO_HOVER, BLANCO, GRIS_TEXTO, RADIUS_MD, FONT_FAMILY)
from src.models.database import obtener_datos_periodo
from src.scrapping.scrapping_main import scrapping_main


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
                check_cancel=self.isInterruptionRequested
            )
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self, ruta_raiz):
        super().__init__()
        self.ruta_raiz = ruta_raiz # Guardamos la ruta del proyecto
        self._worker = None  # Referencia al hilo de scrapping
        self.setWindowTitle("Zeus Excels - Sistema de Extracción")
        self.setFixedSize(1000, 600)
        
        # Eliminar barra de título nativa y hacer esquinas redondeadas
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # Variables para mover la ventana custom
        self._drag_pos = None

        # Configurar el icono de la ventana
        ruta_icono = os.path.join(self.ruta_raiz, "src", "assets", "icons", "icon.svg")
        self.setWindowIcon(QIcon(ruta_icono))
        
        # Carga de datos usando el modelo
        self.datos_anos = obtener_datos_periodo("anios")
        self.datos_meses = obtener_datos_periodo("meses")
        
        self.init_ui()

    def init_ui(self):
        # Contenedor principal para poder aplicar border-radius a toda la ventana
        self.main_container = QWidget(self)
        self.main_container.setObjectName("MainContainer")
        self.main_container.setStyleSheet(f"""
            QWidget#MainContainer {{
                background-color: {BLANCO};
                border-radius: {RADIUS_MD};
                border: 1px solid {GRIS_BORDE};
            }}
        """)
        
        # Layout del main container
        container_layout = QVBoxLayout(self.main_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)
        
        # ── CUSTOM TITLE BAR ─────────────────────────────────────────────
        title_bar = QWidget()
        title_bar.setFixedHeight(40)
        title_bar.setStyleSheet(f"""
            QWidget {{
                background-color: {OSCURO};
                border-top-left-radius: {RADIUS_MD};
                border-top-right-radius: {RADIUS_MD};
            }}
        """)
        title_layout = QHBoxLayout(title_bar)
        title_layout.setContentsMargins(16, 0, 16, 0)
        
        # Eventos para mover la ventana desde la barra de título
        title_bar.mousePressEvent = self.title_bar_mousePressEvent
        title_bar.mouseMoveEvent = self.title_bar_mouseMoveEvent
        
        app_title = QLabel("ZEUS EXCELS")
        app_title.setStyleSheet(f"color: {DORADO}; font-weight: bold; font-family: {FONT_FAMILY}; letter-spacing: 1px;")
        
        btn_close = QPushButton("✕")
        btn_close.setFixedSize(30, 30)
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {GRIS_TEXTO};
                border: none;
                font-weight: bold;
                font-size: 14px;
            }}
            QPushButton:hover {{
                color: white;
                background-color: #ef4444;
                border-radius: 4px;
            }}
        """)
        btn_close.clicked.connect(self.close)
        
        btn_min = QPushButton("─")
        btn_min.setFixedSize(30, 30)
        btn_min.setCursor(Qt.PointingHandCursor)
        btn_min.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {GRIS_TEXTO};
                border: none;
                font-weight: bold;
                font-size: 14px;
            }}
            QPushButton:hover {{
                color: white;
                background-color: rgba(255, 255, 255, 0.1);
                border-radius: 4px;
            }}
        """)
        btn_min.clicked.connect(self.showMinimized)
        
        title_layout.addWidget(app_title)
        title_layout.addStretch()
        title_layout.addWidget(btn_min)
        title_layout.addWidget(btn_close)
        
        container_layout.addWidget(title_bar)

        # ── CONTENIDO PRINCIPAL (Paneles Izquierdo y Derecho) ────────────
        content_widget = QWidget()
        layout_principal = QHBoxLayout(content_widget)
        layout_principal.setContentsMargins(0, 0, 0, 0)
        layout_principal.setSpacing(0)

        # ── PANEL IZQUIERDO (Formulario) ─────────────────────────────────
        panel_izquierdo = QWidget()
        panel_izquierdo.setStyleSheet(f"background-color: {BLANCO}; border-bottom-left-radius: {RADIUS_MD};")
        layout_izquierdo = QVBoxLayout(panel_izquierdo)
        layout_izquierdo.setContentsMargins(40, 40, 40, 40)
        layout_izquierdo.setSpacing(24)

        # Header del Formulario
        header_form = QWidget()
        header_layout = QVBoxLayout(header_form)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(8)
        
        lbl_titulo_form = QLabel("Configuración de Carga")
        lbl_titulo_form.setStyleSheet(f"color: {OSCURO}; font-size: 24px; font-weight: bold; font-family: {FONT_FAMILY};")
        
        lbl_subtitulo = QLabel("Especifica los parámetros para iniciar la extracción")
        lbl_subtitulo.setStyleSheet(f"color: {GRIS_TEXTO}; font-size: 13px; font-family: {FONT_FAMILY};")
        
        header_layout.addWidget(lbl_titulo_form)
        header_layout.addWidget(lbl_subtitulo)

        # Card: Directorio
        card_dir = QWidget()
        card_dir.setStyleSheet(f"""
            QWidget {{
                background-color: {GRIS_BG};
                border-radius: {RADIUS_MD};
            }}
        """)
        layout_card_dir = QVBoxLayout(card_dir)
        layout_card_dir.setContentsMargins(20, 20, 20, 20)
        layout_card_dir.setSpacing(12)
        
        etiqueta_dir = QLabel("Directorio de trabajo")
        etiqueta_dir.setStyleSheet(f"color: {OSCURO}; font-size: 13px; font-weight: 600; font-family: {FONT_FAMILY};")

        layout_input_dir = QHBoxLayout()
        self.campo_ruta = QLineEdit()
        self.campo_ruta.setPlaceholderText("C:/Ruta/a/tu/carpeta...")
        self.campo_ruta.setFixedHeight(44)
        self.campo_ruta.setStyleSheet(f"""
            QLineEdit {{
                border: 1px solid {GRIS_BORDE};
                border-radius: {RADIUS_MD};
                background-color: {BLANCO};
                padding-left: 14px;
                color: #333;
                font-size: 13px;
                font-family: {FONT_FAMILY};
            }}
            QLineEdit:focus {{
                border: 1px solid {DORADO};
            }}
        """)
        
        self.btn_browse = QPushButton("📁")
        self.btn_browse.setFixedSize(44, 44)
        self.btn_browse.setCursor(Qt.PointingHandCursor)
        self.btn_browse.setStyleSheet(f"""
            QPushButton {{
                background-color: {BLANCO};
                border: 1px solid {GRIS_BORDE};
                border-radius: {RADIUS_MD};
                font-size: 16px;
            }}
            QPushButton:hover {{
                border: 1px solid {DORADO};
                background-color: rgba(255, 196, 0, 0.1);
            }}
        """)
        self.btn_browse.clicked.connect(self.browse_folder)
        
        layout_input_dir.addWidget(self.campo_ruta)
        layout_input_dir.addWidget(self.btn_browse)
        
        layout_card_dir.addWidget(etiqueta_dir)
        layout_card_dir.addLayout(layout_input_dir)

        # Card: Período
        card_periodo = QWidget()
        card_periodo.setStyleSheet(f"""
            QWidget {{
                background-color: {GRIS_BG};
                border-radius: {RADIUS_MD};
            }}
        """)
        layout_card_per = QVBoxLayout(card_periodo)
        layout_card_per.setContentsMargins(20, 20, 20, 20)
        layout_card_per.setSpacing(12)
        
        etiqueta_periodo = QLabel("Período de análisis")
        etiqueta_periodo.setStyleSheet(f"color: {OSCURO}; font-size: 13px; font-weight: 600; font-family: {FONT_FAMILY};")

        # Inyección de Menubars reutilizables
        anio_actual = str(datetime.datetime.now().year)
        self.combo_ano = crear_menubar("Año", self.datos_anos, default_val=anio_actual)
        self.combo_mes = crear_menubar("Mes", self.datos_meses)

        layout_combos = QHBoxLayout()
        layout_combos.addWidget(self.combo_ano)
        layout_combos.addWidget(self.combo_mes)
        layout_combos.setSpacing(14)
        
        layout_card_per.addWidget(etiqueta_periodo)
        layout_card_per.addLayout(layout_combos)

        # Botón Iniciar
        self.boton_iniciar = QPushButton("⚡ INICIAR EXTRACCIÓN")
        self.boton_iniciar.setFixedHeight(50)
        self.boton_iniciar.setCursor(Qt.PointingHandCursor)
        self.boton_iniciar.setStyleSheet(f"""
            QPushButton {{
                background-color: {DORADO};
                color: {OSCURO};
                border-radius: {RADIUS_MD};
                font-size: 14px;
                font-weight: 800;
                font-family: {FONT_FAMILY};
                letter-spacing: 1px;
            }}
            QPushButton:hover {{ 
                background-color: {DORADO_HOVER}; 
            }}
            QPushButton:disabled {{ 
                background-color: {GRIS_DESHABILITADO}; 
                color: rgba(255,255,255,0.5);
            }}
        """)
        self.boton_iniciar.clicked.connect(self.on_click_iniciar)

        # Progreso
        self.progreso_widget = QWidget()
        self.progreso_widget.setVisible(False)
        layout_progreso = QVBoxLayout(self.progreso_widget)
        layout_progreso.setContentsMargins(0, 0, 0, 0)
        layout_progreso.setSpacing(8)
        
        self.lbl_progreso = QLabel("Preparando...")
        self.lbl_progreso.setStyleSheet(f"color: {GRIS_TEXTO}; font-size: 12px; font-family: {FONT_FAMILY};")
        
        self.barra_progreso = crear_barra_progreso()
        
        layout_progreso.addWidget(self.lbl_progreso)
        layout_progreso.addWidget(self.barra_progreso)

        # Construcción del Layout Izquierdo
        layout_izquierdo.addWidget(header_form)
        layout_izquierdo.addWidget(card_dir)
        layout_izquierdo.addWidget(card_periodo)
        layout_izquierdo.addStretch()
        layout_izquierdo.addWidget(self.progreso_widget)
        layout_izquierdo.addWidget(self.boton_iniciar)


        # ── PANEL DERECHO (Visual Hero) ──────────────────────────────────
        panel_derecho = QWidget()
        panel_derecho.setFixedWidth(450)
        panel_derecho.setObjectName("panel_derecho")
        
        path_fondo = os.path.join(self.ruta_raiz, "src", "assets", "images", "fondo.png").replace("\\", "/")
        path_logo = os.path.join(self.ruta_raiz, "src", "assets", "images", "Img.png")
        
        panel_derecho.setStyleSheet(f"""
            QWidget#panel_derecho {{
                border-image: url("{path_fondo}") 0 0 0 0 stretch stretch;
                border-bottom-right-radius: {RADIUS_MD};
            }}
        """)
        
        # Overlay oscuro
        overlay = QWidget(panel_derecho)
        overlay.setStyleSheet(f"background-color: rgba(15, 15, 26, 0.7); border-bottom-right-radius: {RADIUS_MD};")
        overlay.resize(450, 600)
        
        layout_derecho = QVBoxLayout(overlay)
        layout_derecho.setContentsMargins(40, 40, 40, 40)
        
        layout_derecho.addStretch()
        
        # Logo Central
        if os.path.exists(path_logo):
            logo_label = QLabel()
            pixmap = QPixmap(path_logo).scaled(320, 320, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
            logo_label.setAlignment(Qt.AlignCenter)
            layout_derecho.addWidget(logo_label)
        
        layout_derecho.addStretch()
        
        # Footer / Tagline
        lbl_tagline = QLabel("Optimizado para alto rendimiento")
        lbl_tagline.setAlignment(Qt.AlignCenter)
        lbl_tagline.setStyleSheet(f"color: rgba(255,255,255,0.5); font-size: 11px; font-family: {FONT_FAMILY};")
        layout_derecho.addWidget(lbl_tagline)

        # Ensamblado Final
        layout_principal.addWidget(panel_izquierdo)
        layout_principal.addWidget(panel_derecho)
        
        container_layout.addWidget(content_widget)
        
        self.setCentralWidget(self.main_container)

    # ── Métodos de Título y Movimiento ───────────────────────────────
    def title_bar_mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def title_bar_mouseMoveEvent(self, event: QMouseEvent):
        if event.buttons() == Qt.LeftButton and self._drag_pos is not None:
            self.move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Directorio")
        if folder:
            self.campo_ruta.setText(folder)

    # ── Métodos de Eventos ───────────────────────────────────────────
    def on_click_iniciar(self):
        directorio = self.campo_ruta.text().strip()
        if not directorio:
            show_warning(self, "Advertencia", "Por favor selecciona un directorio de trabajo válido.")
            return
            
        anio = self.combo_ano.currentText()
        mes_texto = self.combo_mes.currentText()
        mes = mes_texto.split(",")[0].strip()
        fecha_cierre_sistema = datetime.datetime.now().strftime("%d.%m.%Y")

        # Modificación de UI estado de carga
        self.boton_iniciar.setEnabled(False)
        self.boton_iniciar.setText("EJECUTANDO...")
        self.campo_ruta.setEnabled(False)
        self.btn_browse.setEnabled(False)
        self.combo_ano.setEnabled(False)
        self.combo_mes.setEnabled(False)
        
        self.progreso_widget.setVisible(True)
        self.barra_progreso.setRange(0, 0) # Loading infinito

        # Ejecutar scrapping en hilo de fondo
        self._worker = ScrappingWorker(directorio, anio, mes, fecha_cierre_sistema)
        self._worker.finished.connect(self._on_scrapping_finished)
        self._worker.finished.connect(self._release_worker)
        self._worker.error.connect(self._on_scrapping_error)
        self._worker.error.connect(self._release_worker)
        self._worker.progress.connect(self._on_scrapping_progress)
        self._worker.start()

    def _reset_ui_state(self):
        self.boton_iniciar.setEnabled(True)
        self.boton_iniciar.setText("⚡ INICIAR EXTRACCIÓN")
        self.campo_ruta.setEnabled(True)
        self.btn_browse.setEnabled(True)
        self.combo_ano.setEnabled(True)
        self.combo_mes.setEnabled(True)
        self.progreso_widget.setVisible(False)
        self.barra_progreso.setRange(0, 100)

    def _on_scrapping_finished(self):
        self._reset_ui_state()
        show_success(self, "Completado", "El proceso de extracción finalizó correctamente con todos los excels procesados.")

    def _on_scrapping_error(self, mensaje):
        self._reset_ui_state()
        if "Proceso cancelado" not in str(mensaje):
            show_error(self, "Error de Ejecución", f"Ocurrió un problema durante la extracción:\n{mensaje}")

    def _on_scrapping_progress(self, mensaje):
        self.lbl_progreso.setText(mensaje)

    def _release_worker(self):
        """Limpia la referencia al worker cuando termina (éxito o error)."""
        self._worker = None

    # ── Cierre seguro de la ventana ─────────────────────────────────
    def closeEvent(self, event):
        """Intercepta el cierre para terminar el worker si sigue activo."""
        if self._worker is not None and self._worker.isRunning():
            # 1. Pedir al hilo que se detenga de forma cooperativa
            self._worker.requestInterruption()

            # 2. Esperar hasta 10 segundos para que termine limpiamente (Excel necesita tiempo)
            terminado = self._worker.wait(10000)

            if not terminado:
                # 3. Si no terminó, forzar la terminación del hilo
                # Esto puede dejar Excel abierto, pero la UI queda cerrable.
                print("[Advertencia] Worker no terminó a tiempo — forzando término.")
                self._worker.terminate()
                self._worker.wait(1000)

        event.accept()