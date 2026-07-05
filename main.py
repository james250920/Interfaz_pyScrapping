import sys
import os
import json
import base64

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

from src.views.main_window import MainWindow


def configurar_salida_utf8():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def ejecutar_scrapping_worker():
    configurar_salida_utf8()

    from src.scrapping.scrapping_main import scrapping_main

    if len(sys.argv) < 3:
        print("ERROR::No se recibió payload para scrapping", flush=True)
        return

    try:
        payload_b64 = sys.argv[2]
        payload_json = base64.b64decode(payload_b64.encode("utf-8")).decode("utf-8")
        payload = json.loads(payload_json)

        def reportar(mensaje):
            print(f"PROGRESS::{mensaje}", flush=True)

        scrapping_main(
            payload["directorio"],
            payload["anio"],
            payload["mes"],
            payload["fecha_cierre"],
            on_progreso=reportar,
            check_cancel=None,
        )

        print("DONE::Proceso finalizado correctamente", flush=True)

    except Exception as e:
        print(f"ERROR::{e}", flush=True)
        return


def main():
    configurar_salida_utf8()

    if len(sys.argv) > 1 and sys.argv[1] == "--scrapping-worker":
        ejecutar_scrapping_worker()
        return

    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    os.environ["PYTHONIOENCODING"] = "utf-8"
    os.environ["PYTHONUTF8"] = "1"

    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    if getattr(sys, "frozen", False):
        ruta_raiz = sys._MEIPASS
    else:
        ruta_raiz = os.path.dirname(os.path.abspath(__file__))

    ventana = MainWindow(ruta_raiz)
    ventana.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()