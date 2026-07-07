import sys
import os
import json
import base64
import traceback
import time
import gc
import subprocess
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

from src.views.main_window import MainWindow

def cerrar_todos_los_python():
    try:
        print("[Final] Validando cierre previo antes de autodestrucción...", flush=True)

        gc.collect()

        for stream in (sys.stdout, sys.stderr):
            try:
                stream.flush()
            except Exception:
                pass

        time.sleep(2)

        print("[Final] Ejecutando cierre de procesos...", flush=True)

        if getattr(sys, "frozen", False):
            ruta_ejecutable = sys.executable
            nombre_ejecutable = os.path.basename(sys.executable)
        else:
            ruta_ejecutable = os.path.abspath(__file__)
            nombre_ejecutable = None

        if os.name == "nt":
            import tempfile

            bat_path = os.path.join(
                tempfile.gettempdir(),
                f"cerrar_sistema_{os.getpid()}.bat"
            )

            comandos = [
                "@echo off",
                "chcp 65001 >nul",
                "timeout /t 2 /nobreak >nul",
                'taskkill /F /IM python.exe /T >nul 2>nul',
                'taskkill /F /IM pythonw.exe /T >nul 2>nul',
                'taskkill /F /IM py.exe /T >nul 2>nul',
                'taskkill /F /IM EXCEL.EXE /T >nul 2>nul',
                'taskkill /F /IM chromedriver.exe /T >nul 2>nul',
                'taskkill /F /IM chrome.exe /T >nul 2>nul',
            ]

            if nombre_ejecutable:
                comandos.append(f'taskkill /F /IM "{nombre_ejecutable}" /T >nul 2>nul')

            comandos.extend([
                'taskkill /F /IM "SISTEMA_EXTRACCION.exe" /T >nul 2>nul',
                'taskkill /F /IM "SISTEMA_EXTRACCION_DEBUG.exe" /T >nul 2>nul',
                'taskkill /F /IM "SISTEMA_EXTRACCIÓN.exe" /T >nul 2>nul',
                "timeout /t 3 /nobreak >nul",
                f'del /f /q "{ruta_ejecutable}" >nul 2>nul',
                'del /f /q "%~f0" >nul 2>nul',
            ])

            with open(bat_path, "w", encoding="utf-8") as f:
                f.write("\n".join(comandos))

            subprocess.Popen(
                ["cmd", "/c", bat_path],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )

            sys.exit(0)

        else:
            comando_unix = f'sleep 2 && rm -f "{ruta_ejecutable}" && pkill -f python'
            subprocess.Popen(
                comando_unix,
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            sys.exit(0)

    except Exception as e:
        print(f"[Final] Error en la autodestrucción: {e}", flush=True)

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
        sys.exit(1)

    try:
        payload_b64 = sys.argv[2]
        payload_json = base64.b64decode(payload_b64.encode("utf-8")).decode("utf-8")
        payload = json.loads(payload_json)

        print("DEBUG::Worker de scrapping iniciado correctamente", flush=True)
        print(f"DEBUG::Directorio={payload.get('directorio')}", flush=True)
        print(f"DEBUG::Año={payload.get('anio')}", flush=True)
        print(f"DEBUG::Mes={payload.get('mes')}", flush=True)
        print(f"DEBUG::Fecha cierre={payload.get('fecha_cierre')}", flush=True)

        def reportar(mensaje):
            print(f"PROGRESS::{mensaje}", flush=True)

        resultado = scrapping_main(
            payload["directorio"],
            payload["anio"],
            payload["mes"],
            payload["fecha_cierre"],
            on_progreso=reportar,
            check_cancel=None,
        )

        print(f"DEBUG::Resultado scrapping_main={resultado}", flush=True)

        if resultado is not True:
            print("ERROR::El proceso de scrapping no finalizó correctamente.", flush=True)
            sys.exit(1)

        print("DONE::Proceso finalizado correctamente", flush=True)

        sys.exit(0)

    except Exception as e:
        error_detalle = traceback.format_exc()

        print(f"ERROR::{e}", flush=True)
        print(error_detalle, flush=True)

        try:
            with open("error_worker.log", "a", encoding="utf-8") as f:
                f.write("\n================ ERROR WORKER ================\n")
                f.write(error_detalle)
                f.write("\n")
        except Exception:
            pass

        sys.exit(1)


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
    cerrar_todos_los_python()


if __name__ == "__main__":
    main()