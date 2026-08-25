import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from PyQt6.QtWidgets import QApplication, QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton  # type: ignore
from PyQt6.QtGui import QIcon, QFont  # type: ignore
from PyQt6.QtCore import Qt  # type: ignore
from ui.main_window import MainWindow  # type: ignore
import seguridad

def _mostrar_dialogo_activacion_qt(app):
    """Ventana de activación de licencia para Camila Conciliador."""
    dlg = QDialog()
    dlg.setWindowTitle("Camila Conciliador — Activación de Licencia")
    dlg.setFixedSize(420, 280)
    dlg.setStyleSheet("background-color: #0f0f1a; color: #f0f0f0; font-family: 'Segoe UI';")
    dlg.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.WindowCloseButtonHint)

    layout = QVBoxLayout(dlg)
    layout.setContentsMargins(30, 30, 30, 20)
    layout.setSpacing(12)

    lbl_titulo = QLabel("Camila Conciliador")
    lbl_titulo.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
    lbl_titulo.setStyleSheet("color: #0f8bfb;")
    lbl_titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
    layout.addWidget(lbl_titulo)

    lbl_sub = QLabel("Ingresa tu Serial de Activación para continuar")
    lbl_sub.setStyleSheet("color: #9ca3b8; font-size: 12px;")
    lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
    layout.addWidget(lbl_sub)

    entry = QLineEdit()
    entry.setPlaceholderText("Ej: CAMI-XXXX-XXXX")
    entry.setAlignment(Qt.AlignmentFlag.AlignCenter)
    entry.setStyleSheet("background: #1a1a40; border: 1px solid #3a3a5a; border-radius: 6px; padding: 8px; font-size: 14px; color: white;")
    entry.setFixedHeight(40)
    layout.addWidget(entry)

    lbl_estado = QLabel("")
    lbl_estado.setStyleSheet("color: #ff4757; font-size: 11px;")
    lbl_estado.setAlignment(Qt.AlignmentFlag.AlignCenter)
    lbl_estado.setWordWrap(True)
    layout.addWidget(lbl_estado)

    resultado = {"ok": False}

    def _activar():
        serial = entry.text().strip()
        btn.setEnabled(False)
        btn.setText("Validando...")
        lbl_estado.setStyleSheet("color: #ffb020; font-size: 11px;")
        lbl_estado.setText("Conectando con el servidor de licencias...")
        app.processEvents()
        ok, msg, _ = seguridad.activar_serial_cloud(serial, "Camila Conciliador")
        if ok:
            lbl_estado.setStyleSheet("color: #00d4aa; font-size: 11px;")
            lbl_estado.setText("✅ " + msg)
            resultado["ok"] = True
            from PyQt6.QtCore import QTimer
            QTimer.singleShot(1200, dlg.accept)
        else:
            lbl_estado.setStyleSheet("color: #ff4757; font-size: 11px;")
            lbl_estado.setText("❌ " + msg)
            btn.setEnabled(True)
            btn.setText("Activar Licencia")

    btn = QPushButton("Activar Licencia")
    btn.setFixedHeight(40)
    btn.setStyleSheet("background: #0f8bfb; color: white; border-radius: 6px; font-size: 14px; font-weight: bold;")
    btn.clicked.connect(_activar)
    layout.addWidget(btn)

    lbl_ayuda = QLabel("Para obtener tu serial contacta a tu administrador.")
    lbl_ayuda.setStyleSheet("color: #6b7280; font-size: 10px;")
    lbl_ayuda.setAlignment(Qt.AlignmentFlag.AlignCenter)
    layout.addWidget(lbl_ayuda)

    dlg.exec()
    return resultado["ok"]

def main():
    app = QApplication(sys.argv)

    font = app.font()
    font.setPointSize(10)
    app.setFont(font)

    if hasattr(sys, '_MEIPASS'):
        logo_path = os.path.join(sys._MEIPASS, "ui", "logo.png")  # type: ignore
    else:
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ui", "logo.png")

    if os.path.exists(logo_path):
        app.setWindowIcon(QIcon(logo_path))

    # ── VALIDACIÓN DE LICENCIA ────────────────────────────────────────
    valido, _ = seguridad.validar_licencia("Camila Conciliador")
    if not valido:
        activado = _mostrar_dialogo_activacion_qt(app)
        if not activado:
            sys.exit(0)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()
