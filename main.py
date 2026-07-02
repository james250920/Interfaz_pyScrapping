import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from src.views.main_window import MainWindow


def main():
    # Optimizaciones de rendimiento de renderizado UI
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    
    app = QApplication(sys.argv)
    
    # Optimizaciones adicionales para la aplicación
    app.setStyle("Fusion") # Estilo moderno y ligero por defecto
    
    # Obtenemos la ruta raíz absoluta del proyecto actual
    ruta_raiz = os.path.dirname(os.path.abspath(__file__))
    
    # Pasamos la ruta raíz a la ventana para que busque los archivos correctamente
    ventana = MainWindow(ruta_raiz)
    ventana.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()