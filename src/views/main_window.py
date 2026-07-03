import os
import datetime
from PySide6.QtWidgets import (QMainWindow, QWidget, QLabel, QLineEdit,
                             QPushButton, QHBoxLayout, QVBoxLayout, QFileDialog,
                             QGraphicsDropShadowEffect, QFrame,QSizePolicy)
from PySide6.QtCore import Qt, QThread, Signal, QSize
from PySide6.QtGui import QIcon, QMouseEvent, QPixmap, QColor, QPainter
from PySide6.QtSvg import QSvgRenderer

# Importamos componentes, modelos y tema centralizado
from src.views.components.widgets_personalizados import (
    crear_menubar, crear_imagen, crear_barra_progreso, crear_item_check,
    crear_fila_icono_texto, crear_boton_icono, BotonIconoHover,
)
from src.views.components.dialogs import show_success, show_error, show_warning, show_info
from src.views.theme import *
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
        self.ruta_raiz = ruta_raiz  # Guardamos la ruta del proyecto
        self._worker = None  # Referencia al hilo de scrapping
        self.setWindowTitle("Zeus Excels - Sistema de Extracción")
        self.setFixedSize(1120, 640)

        # Eliminar barra de título nativa y hacer esquinas redondeadas
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        # Variables para mover la ventana custom
        self._drag_pos = None
        self._is_maximized = False

        # Configurar el icono de la ventana
        ruta_icono = os.path.join(self.ruta_raiz, "src", "assets", "icons", "icon.svg")
        self.setWindowIcon(QIcon(ruta_icono))

        # Carga de datos usando el modelo
        self.datos_anos = obtener_datos_periodo("anios")
        self.datos_meses = obtener_datos_periodo("meses")

        self.init_ui()

    # ══════════════════════════════════════════════════════════════
    # CONSTRUCCIÓN DE LA UI
    # ══════════════════════════════════════════════════════════════
    def init_ui(self):
        # Contenedor principal para poder aplicar border-radius a toda la ventana
        self.main_container = QWidget(self)
        self.main_container.setObjectName("MainContainer")
        self.main_container.setStyleSheet(f"""
            QWidget#MainContainer {{
                background-color: {BLANCO};
                border-radius: 12px;
                border: 1px solid {GRIS_BORDE};
            }}
        """)

        container_layout = QVBoxLayout(self.main_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)

        container_layout.addWidget(self._crear_title_bar())

        # ── CONTENIDO PRINCIPAL (Paneles Izquierdo y Derecho) ────────
        content_widget = QWidget()
        layout_principal = QHBoxLayout(content_widget)
        layout_principal.setContentsMargins(0, 0, 0, 0)
        layout_principal.setSpacing(0)

        layout_principal.addWidget(self._crear_panel_izquierdo(), 1)
        layout_principal.addWidget(self._crear_panel_derecho())

        container_layout.addWidget(content_widget, 1)
        container_layout.addWidget(self._crear_footer())

        self.setCentralWidget(self.main_container)

    # ── TITLE BAR ─────────────────────────────────────────────────
    def _crear_title_bar(self) -> QWidget:
        title_bar = QWidget()
        title_bar.setFixedHeight(48)
        title_bar.setStyleSheet(f"""
            QWidget {{
                background-color: {BLANCO};
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                border-bottom: 1px solid {GRIS_BORDE};
            }}
        """)
        title_layout = QHBoxLayout(title_bar)
        title_layout.setContentsMargins(20, 0, 12, 0)
        title_layout.setSpacing(10)

        # Eventos para mover la ventana desde la barra de título
        title_bar.mousePressEvent = self.title_bar_mousePressEvent
        title_bar.mouseMoveEvent = self.title_bar_mouseMoveEvent
        title_bar.mouseDoubleClickEvent = lambda e: self.toggle_maximizar()

        # Icono pequeño (mismo icono de la app, ej. SVG de marca)
        icono_lbl = QLabel()
        ruta_icono = os.path.join(self.ruta_raiz, "src", "assets", "icons", "icon.svg")
        if os.path.exists(ruta_icono):
            pixmap = QPixmap(ruta_icono).scaled(18, 18, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            icono_lbl.setPixmap(pixmap)
        title_layout.addWidget(icono_lbl)

        app_title = QLabel("SISTEMA DE EXTRACCIÓN ZEUS")
        app_title.setStyleSheet(f"""
            color: {ROJO};
            font-weight: bold;
            font-size: 13px;
            font-family: {FONT_FAMILY_TITLE};
            letter-spacing: 0.5px;
        """)
        title_layout.addWidget(app_title)
        title_layout.addStretch()

        # Botones de ventana (min / max / close) — íconos Font Awesome,
        # cambian de color al pasar el mouse (ver BotonIconoHover)
        btn_min = BotonIconoHover(
            "fa5s.window-minimize", TEXTO_SECUNDARIO, TEXTO, "rgba(0,0,0,0.06)"
        )
        btn_min.clicked.connect(self.showMinimized)

        btn_max = BotonIconoHover(
            "fa5s.window-maximize", TEXTO_SECUNDARIO, TEXTO, "rgba(0,0,0,0.06)", tam_icono=10
        )
        btn_max.clicked.connect(self.toggle_maximizar)

        btn_close = BotonIconoHover(
            "fa5s.times", TEXTO_SECUNDARIO, BLANCO, "#ef4444", tam_icono=13
        )
        btn_close.clicked.connect(self.close)

        title_layout.addWidget(btn_min)
        title_layout.addWidget(btn_max)
        title_layout.addWidget(btn_close)

        return title_bar

    # ── PANEL IZQUIERDO (Formulario) ─────────────────────────────
    def _crear_panel_izquierdo(self) -> QWidget:
        panel = QWidget()
        panel.setStyleSheet(f"background-color: {BLANCO};")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(48, 36, 48, 12)
        layout.setSpacing(10)

        layout.addWidget(self._crear_logo())
        layout.addWidget(self._crear_header_formulario())
        layout.addWidget(self._crear_card_directorio())
        layout.addWidget(self._crear_card_periodo())
        layout.addStretch()
        layout.addWidget(self._crear_progreso())
        layout.addWidget(self._crear_boton_iniciar())

        return panel

    def _crear_logo(self) -> QWidget:
        """Logo institucional. Se carga desde un único SVG
        (logo_fonafe.svg) que ya trae integrados el ícono y el texto
        'CORPORACIÓN FONAFE' — no se arma con labels sueltos."""
        cont = QWidget()
        row = QHBoxLayout(cont)
        row.setContentsMargins(0, 0, 0, 0)

        lbl_logo = QLabel()
        alto_logo = 80 # alto objetivo del logo dentro del panel

        ruta_svg = os.path.join(self.ruta_raiz, "src", "assets", "images", "logo_fonafe.svg")

        if os.path.exists(ruta_svg):
            renderer = QSvgRenderer(ruta_svg)
            tam_base = renderer.defaultSize()

            if tam_base.width() > 0 and tam_base.height() > 0:
                ancho_logo = int(alto_logo * tam_base.width() / tam_base.height())
            else:
                ancho_logo = 160

            # Se renderiza a 2x y se marca el devicePixelRatio para que
            # el SVG se vea nítido en pantallas HiDPI (evita el
            # pixelado que da usar QPixmap(path).scaled() directamente).
            escala = 3
            pixmap = QPixmap(ancho_logo * escala, alto_logo * escala)
            pixmap.fill(Qt.transparent)
            painter = QPainter(pixmap)
            renderer.render(painter)
            painter.end()
            pixmap.setDevicePixelRatio(escala)

            lbl_logo.setPixmap(pixmap)
            lbl_logo.setFixedSize(ancho_logo, alto_logo)
        else:
            # Fallback de texto si el SVG no está presente, para no
            # romper la ventana mientras se agrega el asset definitivo.
            lbl_logo.setText("CORPORACIÓN FONAFE")
            lbl_logo.setStyleSheet(f"""
                color: {ROJO};
                font-size: 16px;
                font-weight: 800;
                font-family: {FONT_FAMILY_TITLE};
            """)

        row.addWidget(lbl_logo)
        row.addStretch()
        return cont

    def _crear_header_formulario(self) -> QWidget:
        header = QWidget()
        layout = QVBoxLayout(header)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        lbl_titulo = QLabel("Configuración de Carga")
        lbl_titulo.setStyleSheet(f"""
            color: {TEXTO};
            font-size: 25px;
            font-weight: 800;
            font-family: {FONT_FAMILY_TITLE};
        """)

        linea = QFrame()
        linea.setFixedSize(46, 4)
        linea.setStyleSheet(f"background-color: {ROJO}; border-radius: 2px;")

        lbl_subtitulo = QLabel("Especifica los parámetros para iniciar la extracción de información.")
        lbl_subtitulo.setWordWrap(True)
        lbl_subtitulo.setStyleSheet(f"""
            color: {TEXTO_SECUNDARIO};
            font-size: 13px;
            font-family: {FONT_FAMILY};
        """)

        layout.addWidget(lbl_titulo)
        layout.addWidget(linea)
        layout.addWidget(lbl_subtitulo)
        return header

    def _crear_tarjeta_base(self) -> QWidget:
        """Tarjeta gris clara reutilizable con sombra suave."""
        card = QWidget()
        card.setStyleSheet(f"""
            QWidget {{
                background-color: {GRIS_BG};
                border-radius: 14px;
            }}
        """)
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setXOffset(0)
        shadow.setYOffset(4)
        shadow.setColor(QColor(0, 0, 0, 15))
        card.setGraphicsEffect(shadow)
        return card

    def _crear_card_directorio(self) -> QWidget:
        card = self._crear_tarjeta_base()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        etiqueta = crear_fila_icono_texto(
            "fa5s.folder", "Directorio de trabajo",
            color_icono=TEXTO, color_texto=TEXTO
        )

        fila = QHBoxLayout()
        fila.setSpacing(10)

        self.campo_ruta = QLineEdit()
        self.campo_ruta.setPlaceholderText("D:/Ruta/a/tu/carpeta...")
        self.campo_ruta.setFixedHeight(44)
        self.campo_ruta.setStyleSheet(f"""
            QLineEdit {{
                border: 1px solid {GRIS_BORDE};
                border-radius: 10px;
                background-color: {BLANCO};
                padding-left: 14px;
                color: {TEXTO};
                font-size: 13px;
                font-family: {FONT_FAMILY};
            }}
            QLineEdit:focus {{
                border: 1.5px solid {ROJO};
            }}
        """)

        self.btn_browse = crear_boton_icono(
            "Examinar", "fa5s.folder-open",
            color_fondo=ROJO, color_texto=BLANCO, color_hover=ROJO_HOVER, alto=44
        )
        self.btn_browse.clicked.connect(self.browse_folder)

        fila.addWidget(self.campo_ruta)
        fila.addWidget(self.btn_browse)

        layout.addWidget(etiqueta)
        layout.addLayout(fila)
        return card

    def _crear_card_periodo(self) -> QWidget:
        card = self._crear_tarjeta_base()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        etiqueta = crear_fila_icono_texto(
            "fa5s.calendar-alt", "Periodo de análisis",
            color_icono=TEXTO, color_texto=TEXTO
        )

        anio_actual = str(datetime.datetime.now().year)
        self.combo_ano = crear_menubar("Año", self.datos_anos, default_val=anio_actual)
        self.combo_mes = crear_menubar("Mes", self.datos_meses)

        # Sub-bloques con etiqueta propia (Año / Mes), como en el diseño
        bloque_ano = QVBoxLayout()
        bloque_ano.setSpacing(6)
        lbl_ano = QLabel("Año")
        lbl_ano.setStyleSheet(f"color: {TEXTO_SECUNDARIO}; font-size: 11px; font-family: {FONT_FAMILY};")
        bloque_ano.addWidget(lbl_ano)
        bloque_ano.addWidget(self.combo_ano)

        bloque_mes = QVBoxLayout()
        bloque_mes.setSpacing(6)
        lbl_mes = QLabel("Mes")
        lbl_mes.setStyleSheet(f"color: {TEXTO_SECUNDARIO}; font-size: 11px; font-family: {FONT_FAMILY};")
        bloque_mes.addWidget(lbl_mes)
        bloque_mes.addWidget(self.combo_mes)

        fila_combos = QHBoxLayout()
        fila_combos.setSpacing(18)
        fila_combos.addLayout(bloque_ano, 1)
        fila_combos.addLayout(bloque_mes, 1)

        layout.addWidget(etiqueta)
        layout.addLayout(fila_combos)
        return card

    def _crear_progreso(self) -> QWidget:
        self.progreso_widget = QWidget()
        self.progreso_widget.setVisible(False)
        layout = QVBoxLayout(self.progreso_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        self.lbl_progreso = QLabel("Preparando...")
        self.lbl_progreso.setStyleSheet(f"color: {TEXTO_SECUNDARIO}; font-size: 12px; font-family: {FONT_FAMILY};")

        self.barra_progreso = crear_barra_progreso()

        layout.addWidget(self.lbl_progreso)
        layout.addWidget(self.barra_progreso)
        return self.progreso_widget

    def _crear_boton_iniciar(self) -> QPushButton:
        self.boton_iniciar = crear_boton_icono(
            "INICIAR EXTRACCIÓN", "fa5s.bolt",
            color_fondo=ROJO, color_texto=BLANCO, color_hover=ROJO_HOVER,
            alto=52, radio=12, tam_icono=15
        )
        self.boton_iniciar.setStyleSheet(self.boton_iniciar.styleSheet() + f"""
            QPushButton {{ font-size: 14px; font-weight: 800; letter-spacing: 1px;
                           font-family: {FONT_FAMILY_TITLE}; }}
        """)
        self.boton_iniciar.clicked.connect(self.on_click_iniciar)

        shadow_btn = QGraphicsDropShadowEffect(self)
        shadow_btn.setBlurRadius(20)
        shadow_btn.setXOffset(0)
        shadow_btn.setYOffset(6)
        shadow_btn.setColor(QColor(200, 16, 46, 70))
        self.boton_iniciar.setGraphicsEffect(shadow_btn)

        return self.boton_iniciar

    # ── PANEL DERECHO (Visual Hero) ───────────────────────────────
    def _crear_panel_derecho(self) -> QWidget:
        panel = QWidget()
        panel.setFixedWidth(600)
        panel.setObjectName("panel_derecho")

        # Si existe una imagen de fondo (textura/globo), se usa como base;
        # si no, se aplica un degradado rojo corporativo.
        ruta_fondo = os.path.join(self.ruta_raiz, "src", "assets", "images", "fondo.png").replace("\\", "/")

        if os.path.exists(ruta_fondo):
            panel.setStyleSheet(f"""
                QWidget#panel_derecho {{
                    border-image: url("{ruta_fondo}") 0 0 0 0 stretch stretch;
                    border-bottom-right-radius: 12px;
                }}
            """)
        else:
            panel.setStyleSheet(f"""
                QWidget#panel_derecho {{
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                        stop:0 {ROJO}, stop:1 {ROJO_OSCURO});
                    border-bottom-right-radius: 12px;
                }}
            """)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(36, 40, 36, 32)
        layout.setSpacing(0)

        # Título
        lbl_titulo = QLabel("SISTEMA DE\nEXTRACCIÓN ZEUS")
        lbl_titulo.setStyleSheet(f"""
            color: {TEXTO_SECUNDARIO};
            font-size: 26px;
            font-weight: 800;
            font-family: {FONT_FAMILY_TITLE};
            line-height: 120%;
        """)
        layout.addWidget(lbl_titulo)

        lbl_subtitulo = QLabel("EXTRACCIÓN AVANZADA DE DATOS")
        lbl_subtitulo.setStyleSheet(f"""
            color: {ROJO_SUAVE};
            font-size: 12px;
            font-weight: 700;
            font-family: {FONT_FAMILY};
            letter-spacing: 1px;
            margin-top: 8px;
        """)
        layout.addWidget(lbl_subtitulo)

        linea = QFrame()
        linea.setFixedSize(46, 3)
        linea.setStyleSheet(f"background-color: {BLANCO}; border-radius: 1px; margin-top: 14px;")
        layout.addWidget(linea)

        lbl_desc = QLabel("Automatiza la extracción de información financiera\n y presupuestal desde el SISFONAFE.")
        lbl_desc.setWordWrap(True)
        lbl_desc.setStyleSheet(f"""
            color: {TEXTO_SECUNDARIO};
            font-size: 13px;
            font-family: {FONT_FAMILY};
            letter-spacing: 1px;
            margin-top: 14px;
        """)
        layout.addWidget(lbl_desc)

        layout.addStretch()

        layout.addWidget(self._crear_card_archivos())

        return panel

    def _crear_card_archivos(self) -> QWidget:
        card = QWidget()
        card.setStyleSheet(f"""
            QWidget {{
                background-color: {OVERLAY_OSCURO};
                border: 1px solid {BORDE_OVERLAY};
                border-radius: 14px;
            }}
        """)

        # La card solo ocupará el ancho necesario
        card.setSizePolicy(
            QSizePolicy.Policy.Maximum,
            QSizePolicy.Policy.Preferred
        )

        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        lbl_titulo = QLabel("Archivos a extraer")
        lbl_titulo.setStyleSheet(f"""
            color: {BLANCO};
            font-size: 14px;
            font-weight: 400;
            font-family: {FONT_FAMILY_TITLE};
        """)
        lbl_titulo.setSizePolicy(
            QSizePolicy.Policy.Maximum,
            QSizePolicy.Policy.Fixed
        )
        layout.addWidget(lbl_titulo)

        items = [
            "Estado de Situación Financiera",
            "Estado de Resultados Integrales",
            "Presupuesto de Ingresos y Egresos",
            "Flujo de Caja",
            "Gastos de Capital",
            "Depósitos y Colocaciones",
            "y otros reportes asociados",
        ]

        for texto in items:
            item = crear_item_check(texto)
            item.setSizePolicy(
                QSizePolicy.Policy.Maximum,
                QSizePolicy.Policy.Fixed
            )
            layout.addWidget(item)

        layout.addStretch()

        return card

    # ── FOOTER ────────────────────────────────────────────────────
    def _crear_footer(self) -> QWidget:
        footer = QWidget()
        footer.setFixedHeight(34)
        footer.setStyleSheet(f"""
            background-color: {GRIS_BG};
            border-top: 1px solid {GRIS_BORDE};
            border-bottom-left-radius: 12px;
            border-bottom-right-radius: 12px;
        """)
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(20, 0, 20, 0)

        lbl = QLabel("Uso interno • Versión 1.0.0 • Área Corporativa de Presupuesto")
        lbl.setStyleSheet(f"color: {TEXTO_SECUNDARIO}; font-size: 11px; font-family: {FONT_FAMILY};")
        layout.addStretch()
        layout.addWidget(lbl)
        layout.addStretch()

        return footer

    # ══════════════════════════════════════════════════════════════
    # Métodos de título / movimiento de ventana
    # ══════════════════════════════════════════════════════════════
    def title_bar_mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def title_bar_mouseMoveEvent(self, event: QMouseEvent):
        if event.buttons() == Qt.LeftButton and self._drag_pos is not None:
            self.move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()

    def toggle_maximizar(self):
        if self._is_maximized:
            self.showNormal()
        else:
            self.showMaximized()
        self._is_maximized = not self._is_maximized

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Directorio")
        if folder:
            self.campo_ruta.setText(folder)

    # ══════════════════════════════════════════════════════════════
    # Métodos de eventos / lógica de negocio (sin cambios funcionales)
    # ══════════════════════════════════════════════════════════════
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
        self.boton_iniciar.setText("  EJECUTANDO...")
        self.campo_ruta.setEnabled(False)
        self.btn_browse.setEnabled(False)
        self.combo_ano.setEnabled(False)
        self.combo_mes.setEnabled(False)

        self.progreso_widget.setVisible(True)
        self.barra_progreso.setRange(0, 0)  # Loading infinito

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
        self.boton_iniciar.setText("  INICIAR EXTRACCIÓN")
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
            self._worker.requestInterruption()
            terminado = self._worker.wait(10000)

            if not terminado:
                print("[Advertencia] Worker no terminó a tiempo — forzando término.")
                self._worker.terminate()
                self._worker.wait(1000)

        event.accept()