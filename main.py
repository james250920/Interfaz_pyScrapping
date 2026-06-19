import sys
import os
from PySide6.QtWidgets import QApplication
from src.views.main_window import MainWindow


def main():
    # Obtenemos la ruta raíz absoluta del proyecto actual
    ruta_raiz = os.path.dirname(os.path.abspath(__file__))
    
    app = QApplication(sys.argv)
    
    # Pasamos la ruta raíz a la ventana para que busque los archivos correctamente
    ventana = MainWindow(ruta_raiz)
    ventana.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()