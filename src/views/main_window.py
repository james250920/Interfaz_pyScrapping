import os
import sys
import json
import base64
import datetime
from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QHBoxLayout,
    QVBoxLayout,
    QFileDialog,
    QGraphicsDropShadowEffect,
    QFrame,
    QSizePolicy,
)
from PySide6.QtCore import Qt, QProcess, QProcessEnvironment, QTimer
from PySide6.QtGui import QIcon, QMouseEvent, QPixmap, QColor, QPainter
from PySide6.QtSvg import QSvgRenderer

from src.views.components.widgets_personalizados import (
    crear_menubar,
    crear_barra_progreso,
    crear_item_check,
    crear_fila_icono_texto,
    crear_boton_icono,
    BotonIconoHover,
)
from src.views.components.dialogs import (
    show_success,
    show_error,
    show_warning,
)
from src.views.theme import *
from src.models.database import obtener_datos_periodo


class MainWindow(QMainWindow):
    def __init__(self, ruta_raiz):
        super().__init__()
        self.ruta_raiz = ruta_raiz

        self._process = None
        self._cerrando = False
        self._ultimo_error_proceso = None
        self._process_stdout_buffer = ""
        self._proceso_marco_done = False
        self._success_mostrado = False

        self.setWindowTitle("Zeus Excels - Sistema de Extracción")
        self.setFixedSize(1120, 640)

        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self._drag_pos = None
        self._is_maximized = False

        ruta_icono = os.path.join(
            self.ruta_raiz,
            "src",
            "assets",
            "icons",
            "icon.svg",
        )
        self.setWindowIcon(QIcon(ruta_icono))

        self.datos_anos = obtener_datos_periodo("anios")
        self.datos_meses = obtener_datos_periodo("meses")

        self.init_ui()

    # ══════════════════════════════════════════════════════════════
    # CONSTRUCCIÓN DE LA UI
    # ══════════════════════════════════════════════════════════════

    def init_ui(self):
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

        title_bar.mousePressEvent = self.title_bar_mousePressEvent
        title_bar.mouseMoveEvent = self.title_bar_mouseMoveEvent
        title_bar.mouseDoubleClickEvent = lambda e: self.toggle_maximizar()

        icono_lbl = QLabel()
        ruta_icono = os.path.join(
            self.ruta_raiz,
            "src",
            "assets",
            "icons",
            "icon.svg",
        )

        if os.path.exists(ruta_icono):
            pixmap = QPixmap(ruta_icono).scaled(
                18,
                18,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation,
            )
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

        btn_min = BotonIconoHover(
            "fa5s.window-minimize",
            TEXTO_SECUNDARIO,
            TEXTO,
            "rgba(0,0,0,0.06)",
        )
        btn_min.clicked.connect(self.showMinimized)

        btn_max = BotonIconoHover(
            "fa5s.window-maximize",
            TEXTO_SECUNDARIO,
            TEXTO,
            "rgba(0,0,0,0.06)",
            tam_icono=10,
        )
        btn_max.clicked.connect(self.toggle_maximizar)

        btn_close = BotonIconoHover(
            "fa5s.times",
            TEXTO_SECUNDARIO,
            BLANCO,
            "#ef4444",
            tam_icono=13,
        )
        btn_close.clicked.connect(self.close)

        title_layout.addWidget(btn_min)
        title_layout.addWidget(btn_max)
        title_layout.addWidget(btn_close)

        return title_bar

    # ── PANEL IZQUIERDO ───────────────────────────────────────────

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
        cont = QWidget()
        row = QHBoxLayout(cont)
        row.setContentsMargins(0, 0, 0, 0)

        lbl_logo = QLabel()
        alto_logo = 80

        ruta_svg = os.path.join(
            self.ruta_raiz,
            "src",
            "assets",
            "images",
            "logo_fonafe.svg",
        )

        if os.path.exists(ruta_svg):
            renderer = QSvgRenderer(ruta_svg)
            tam_base = renderer.defaultSize()

            if tam_base.width() > 0 and tam_base.height() > 0:
                ancho_logo = int(alto_logo * tam_base.width() / tam_base.height())
            else:
                ancho_logo = 160

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
        linea.setStyleSheet(f"""
            background-color: {ROJO};
            border-radius: 2px;
        """)

        lbl_subtitulo = QLabel(
            "Especifica los parámetros para iniciar la extracción de información."
        )
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
            "fa5s.folder",
            "Directorio de trabajo",
            color_icono=TEXTO,
            color_texto=TEXTO,
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
            "Examinar",
            "fa5s.folder-open",
            color_fondo=ROJO,
            color_texto=BLANCO,
            color_hover=ROJO_HOVER,
            alto=44,
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
            "fa5s.calendar-alt",
            "Periodo de análisis",
            color_icono=TEXTO,
            color_texto=TEXTO,
        )

        anio_actual = str(datetime.datetime.now().year)

        self.combo_ano = crear_menubar(
            "Año",
            self.datos_anos,
            default_val=anio_actual,
        )
        self.combo_mes = crear_menubar("Mes", self.datos_meses)

        bloque_ano = QVBoxLayout()
        bloque_ano.setSpacing(6)

        lbl_ano = QLabel("Año")
        lbl_ano.setStyleSheet(f"""
            color: {TEXTO_SECUNDARIO};
            font-size: 11px;
            font-family: {FONT_FAMILY};
        """)

        bloque_ano.addWidget(lbl_ano)
        bloque_ano.addWidget(self.combo_ano)

        bloque_mes = QVBoxLayout()
        bloque_mes.setSpacing(6)

        lbl_mes = QLabel("Mes")
        lbl_mes.setStyleSheet(f"""
            color: {TEXTO_SECUNDARIO};
            font-size: 11px;
            font-family: {FONT_FAMILY};
        """)

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
        self.lbl_progreso.setStyleSheet(f"""
            color: {TEXTO_SECUNDARIO};
            font-size: 12px;
            font-family: {FONT_FAMILY};
        """)

        self.barra_progreso = crear_barra_progreso()

        layout.addWidget(self.lbl_progreso)
        layout.addWidget(self.barra_progreso)

        return self.progreso_widget

    def _crear_boton_iniciar(self) -> QPushButton:
        self.boton_iniciar = crear_boton_icono(
            "INICIAR EXTRACCIÓN",
            "fa5s.bolt",
            color_fondo=ROJO,
            color_texto=BLANCO,
            color_hover=ROJO_HOVER,
            alto=52,
            radio=12,
            tam_icono=15,
        )

        self.boton_iniciar.setStyleSheet(self.boton_iniciar.styleSheet() + f"""
            QPushButton {{
                font-size: 14px;
                font-weight: 800;
                letter-spacing: 1px;
                font-family: {FONT_FAMILY_TITLE};
            }}
        """)

        self.boton_iniciar.clicked.connect(self.on_click_iniciar)

        shadow_btn = QGraphicsDropShadowEffect(self)
        shadow_btn.setBlurRadius(20)
        shadow_btn.setXOffset(0)
        shadow_btn.setYOffset(6)
        shadow_btn.setColor(QColor(200, 16, 46, 70))

        self.boton_iniciar.setGraphicsEffect(shadow_btn)

        return self.boton_iniciar

    # ── PANEL DERECHO ─────────────────────────────────────────────

    def _crear_panel_derecho(self) -> QWidget:
        panel = QWidget()
        panel.setFixedWidth(600)
        panel.setObjectName("panel_derecho")

        ruta_fondo = os.path.join(
            self.ruta_raiz,
            "src",
            "assets",
            "images",
            "fondo.png",
        ).replace("\\", "/")

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
                    background: qlineargradient(
                        x1:0,
                        y1:0,
                        x2:1,
                        y2:1,
                        stop:0 {ROJO},
                        stop:1 {ROJO_OSCURO}
                    );
                    border-bottom-right-radius: 12px;
                }}
            """)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(36, 40, 36, 32)
        layout.setSpacing(0)

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
        linea.setStyleSheet(f"""
            background-color: {BLANCO};
            border-radius: 1px;
            margin-top: 14px;
        """)
        layout.addWidget(linea)

        lbl_desc = QLabel(
            "Automatiza la extracción de información financiera\n"
            " y presupuestal desde el SISFONAFE."
        )
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

        card.setSizePolicy(
            QSizePolicy.Policy.Maximum,
            QSizePolicy.Policy.Preferred,
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
            QSizePolicy.Policy.Fixed,
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
                QSizePolicy.Policy.Fixed,
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

        lbl = QLabel(
            "Uso interno • Versión 1.0.0 • Área Corporativa de Presupuesto"
        )
        lbl.setStyleSheet(f"""
            color: {TEXTO_SECUNDARIO};
            font-size: 11px;
            font-family: {FONT_FAMILY};
        """)

        layout.addStretch()
        layout.addWidget(lbl)
        layout.addStretch()

        return footer

    # ══════════════════════════════════════════════════════════════
    # MOVIMIENTO DE VENTANA
    # ══════════════════════════════════════════════════════════════

    def title_bar_mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            self._drag_pos = (
                event.globalPosition().toPoint()
                - self.frameGeometry().topLeft()
            )
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
        folder = QFileDialog.getExistingDirectory(
            self,
            "Seleccionar Directorio",
        )

        if folder:
            self.campo_ruta.setText(folder)

    # ══════════════════════════════════════════════════════════════
    # LÓGICA DE EJECUCIÓN CON QPROCESS
    # ══════════════════════════════════════════════════════════════

    def on_click_iniciar(self):
        directorio = self.campo_ruta.text().strip()

        if not directorio:
            show_warning(
                self,
                "Advertencia",
                "Por favor selecciona un directorio de trabajo válido.",
            )
            return

        anio = self.combo_ano.currentText().strip()
        mes_texto = self.combo_mes.currentText().strip()

        if not anio or not mes_texto:
            show_warning(
                self,
                "Advertencia",
                "Por favor selecciona un año y un mes válidos.",
            )
            return

        if self._process is not None:
            show_warning(
                self,
                "Advertencia",
                "Ya hay un proceso en ejecución.",
            )
            return

        mes = mes_texto.split(",")[0].strip()
        fecha_cierre_sistema = datetime.datetime.now().strftime("%d.%m.%Y")

        self.boton_iniciar.setEnabled(False)
        self.boton_iniciar.setText("  EJECUTANDO...")
        self.campo_ruta.setEnabled(False)
        self.btn_browse.setEnabled(False)
        self.combo_ano.setEnabled(False)
        self.combo_mes.setEnabled(False)

        self.progreso_widget.setVisible(True)
        self.barra_progreso.setRange(0, 0)
        self.lbl_progreso.setText("Iniciando proceso...")

        self._ultimo_error_proceso = None
        self._process_stdout_buffer = ""

        self._proceso_marco_done = False
        self._success_mostrado = False

        payload = {
            "directorio": directorio,
            "anio": anio,
            "mes": mes,
            "fecha_cierre": fecha_cierre_sistema,
        }

        payload_json = json.dumps(payload, ensure_ascii=False)
        payload_b64 = base64.b64encode(payload_json.encode("utf-8")).decode("utf-8")

        programa, argumentos = self._crear_comando_scrapping(payload_b64)

        process = QProcess(self)
        process.setProcessChannelMode(QProcess.MergedChannels)

        env = QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONIOENCODING", "utf-8")
        env.insert("PYTHONUTF8", "1")
        process.setProcessEnvironment(env)

        process.readyReadStandardOutput.connect(self._on_process_output)
        process.finished.connect(self._on_process_finished)
        process.errorOccurred.connect(self._on_process_error)

        self._process = process
        self._proceso_marco_done = False

        self._process.start(programa, argumentos)

    def _crear_comando_scrapping(self, payload_b64):
        """
        Devuelve el ejecutable y argumentos para lanzar el proceso hijo.

        En desarrollo:
            python main.py --scrapping-worker payload

        En PyInstaller:
            app.exe --scrapping-worker payload
        """

        if getattr(sys, "frozen", False):
            programa = sys.executable
            argumentos = [
                "--scrapping-worker",
                payload_b64,
            ]
            return programa, argumentos

        programa = sys.executable

        script_principal = os.path.abspath(sys.argv[0])

        if not script_principal or not os.path.exists(script_principal):
            script_principal = os.path.join(self.ruta_raiz, "main.py")

        argumentos = [
            script_principal,
            "--scrapping-worker",
            payload_b64,
        ]

        return programa, argumentos

    def _on_process_output(self):
        if self._process is None:
            return

        data = self._process.readAllStandardOutput()
        texto = bytes(data).decode("utf-8", errors="replace")

        self._process_stdout_buffer += texto

        lineas = self._process_stdout_buffer.splitlines(keepends=True)

        if lineas and not lineas[-1].endswith(("\n", "\r")):
            self._process_stdout_buffer = lineas.pop()
        else:
            self._process_stdout_buffer = ""

        for linea in lineas:
            self._procesar_linea_proceso(linea.strip())

    def _procesar_linea_proceso(self, linea):
        if not linea:
            return

        print(linea)

        if linea.startswith("PROGRESS::"):
            mensaje = linea.replace("PROGRESS::", "", 1)
            self.lbl_progreso.setText(mensaje)

            if "Proceso finalizado" in mensaje:
                self._marcar_scrapping_finalizado()

            return

        if linea.startswith("ERROR::"):
            mensaje = linea.replace("ERROR::", "", 1)
            self._ultimo_error_proceso = mensaje
            return

        if linea.startswith("DONE::"):
            self._marcar_scrapping_finalizado()
            return

        if "[Progreso] Proceso finalizado" in linea:
            self._marcar_scrapping_finalizado()
            return

        if linea.strip() == "Proceso finalizado":
            self._marcar_scrapping_finalizado()
            return

    def _marcar_scrapping_finalizado(self):
        if self._proceso_marco_done:
            return

        self._proceso_marco_done = True
        self._ultimo_error_proceso = None
        self._success_mostrado = True

        self.lbl_progreso.setText("Proceso finalizado correctamente.")
        self.barra_progreso.setRange(0, 1)
        self.barra_progreso.setValue(1)

        QTimer.singleShot(0, self._mostrar_dialogo_exito)

    def _mostrar_dialogo_exito(self):
        show_success(
            self,
            "Completado",
            "El proceso de extracción finalizó correctamente con todos los excels procesados.",
        )

        self._terminar_y_cerrar()

    def _terminar_y_cerrar(self):
        """Cierra el proceso hijo si sigue corriendo y luego cierra la aplicación."""
        if self._process is not None:
            try:
                self._process.finished.disconnect(self._on_process_finished)
            except Exception:
                pass

            try:
                if self._process.state() != QProcess.NotRunning:
                    self._process.kill()
                    self._process.waitForFinished(3000)
            except Exception:
                pass

            try:
                self._process.deleteLater()
            except Exception:
                pass

            self._process = None

        from PySide6.QtWidgets import QApplication

        app = QApplication.instance()
        if app is not None:
            app.quit()
        else:
            self.close()

    def _on_process_finished(self, exit_code, exit_status):
        proceso_finalizado = self.sender()

        if self._process is not None and proceso_finalizado is not self._process:
            return

        if self._process_stdout_buffer:
            self._procesar_linea_proceso(self._process_stdout_buffer.strip())
            self._process_stdout_buffer = ""

        self._process = None

        try:
            proceso_finalizado.deleteLater()
        except Exception:
            pass

        if self._cerrando:
            return

        if self._success_mostrado:
            return

        if exit_code == 0 or self._proceso_marco_done:
            self._success_mostrado = True

            show_success(
                self,
                "Completado",
                "El proceso de extracción finalizó correctamente con todos los excels procesados.",
            )

            self._terminar_y_cerrar()
        else:
            mensaje = self._ultimo_error_proceso or (
                "El proceso de scrapping terminó con error."
            )

            show_error(
                self,
                "Error de Ejecución",
                f"Ocurrió un problema durante la extracción:\n{mensaje}",
            )

    def _on_process_error(self, error):
        if self._cerrando:
            return

        self._ultimo_error_proceso = str(error)

    # ══════════════════════════════════════════════════════════════
    # CIERRE SEGURO
    # ══════════════════════════════════════════════════════════════

    def closeEvent(self, event):
        if self._process is None or self._process.state() == QProcess.NotRunning:
            event.accept()
            return

        event.ignore()
        return