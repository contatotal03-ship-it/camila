import sys
import os
from PyQt6.QtWidgets import QApplication # type: ignore
from PyQt6.QtGui import QIcon # type: ignore
from ui.main_window import MainWindow # type: ignore

def main():
    app = QApplication(sys.argv)
    
    # Opcional: configurar fuente global
    font = app.font()
    font.setPointSize(10)
    app.setFont(font)
    
    if hasattr(sys, '_MEIPASS'):
        logo_path = os.path.join(sys._MEIPASS, "ui", "logo.png") # type: ignore
    else:
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ui", "logo.png")
        
    if os.path.exists(logo_path):
        app.setWindowIcon(QIcon(logo_path))
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
