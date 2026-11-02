import sys
import os
from PyQt6.QtWidgets import ( # type: ignore
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QStackedWidget, QFileDialog, QLineEdit,
    QTextEdit, QFrame, QMessageBox, QProgressDialog
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize, QTimer # type: ignore
from PyQt6.QtGui import QFont, QColor # type: ignore

import auto_updater

# Importar la logica de ventas y compras
try:
    from core.ventas_logic import procesar_conciliacion_ventas # type: ignore
    from core.compras_logic import procesar_conciliacion_compras # type: ignore
    from ui.vendedores_widget import VendedoresWidget # type: ignore
except ImportError:
    import sys
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from core.ventas_logic import procesar_conciliacion_ventas # type: ignore
    from core.compras_logic import procesar_conciliacion_compras # type: ignore
    from ui.vendedores_widget import VendedoresWidget # type: ignore


SIDEBAR_BTN_STYLE = (
    "QPushButton {{"
    "  background-color: transparent;"
    "  color: {color};"
    "  text-align: left;"
    "  padding: 15px 20px;"
    "  font-size: 14px;"
    "  font-weight: 500;"
    "  border: none;"
    "  border-left: 3px solid transparent;"
    "  border-bottom: 1px solid #131e29;"
    "}}"
    "QPushButton:hover {{"
    "  background-color: #121e2b;"
    "  color: #e0e6ed;"
    "  border-left-color: #1e3a5f;"
    "}}"
    "QPushButton:checked {{"
    "  background-color: #0f253d;"
    "  color: {accent};"
    "  border-left: 3px solid {accent};"
    "  font-weight: 600;"
    "}}"
)

SIDEBAR_LOGO_STYLE = (
    "QLabel {{"
    "  color: #ffffff;"
    "  font-size: 18px;"
    "  font-weight: 800;"
    "  padding: 8px 10px;"
    "  letter-spacing: 1px;"
    "}}"
)

SIDEBAR_TAG_STYLE = (
    "QLabel {{"
    "  color: #3a5060;"
    "  font-size: 10px;"
    "  font-weight: 500;"
    "  padding: 2px 10px;"
    "  letter-spacing: 1px;"
    "}}"
)


class WorkerThread(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, ruta_sis, ruta_dian, ruta_salida, ruta_dian_detalle=None, tipo_registro="Ventas", **kwargs):
        super().__init__()
        self.ruta_sis = ruta_sis
        self.ruta_dian = ruta_dian
        self.ruta_salida = ruta_salida
        self.ruta_dian_detalle = ruta_dian_detalle
        self.ruta_referencias = kwargs.get('ruta_referencias', None)
        self.tipo_registro = tipo_registro

    def run(self):
        try:
            if self.tipo_registro == "Ventas":
                stats = procesar_conciliacion_ventas(self.ruta_sis, self.ruta_dian, self.ruta_salida, self.ruta_dian_detalle)
            elif self.tipo_registro == "Compras":
                stats = procesar_conciliacion_compras(self.ruta_sis, self.ruta_dian, self.ruta_salida, self.ruta_referencias)
            else:
                raise Exception("Tipo de registro desconocido en WorkerThread.")
                
            self.finished.emit(stats)
        except Exception as e:
            self.error.emit(str(e))


class VentasWidget(QWidget):
    def __init__(self):
        super().__init__()
        
        self.ruta_sistema = ""
        self.ruta_dian = ""
        self.ruta_dian_detalle = ""
        self.ruta_salida = ""
        
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(32, 28, 32, 28)
        layout.setSpacing(0)
        
        # Titulo
        title = QLabel("Conciliacion de Ventas")
        title.setObjectName("TitleLabel")
        subtitle = QLabel("Compara las facturas del ERP vs Token DIAN")
        subtitle.setObjectName("SubtitleLabel")
        
        layout.addWidget(title)
        layout.addSpacing(2)
        layout.addWidget(subtitle)
        layout.addSpacing(16)
        
        # Seccion de Archivo Sistema
        sis_layout = QHBoxLayout()
        self.sis_btn = QPushButton(" Cargar Excel Sistema ")
        self.sis_btn.setProperty("class", "SecondaryButton")
        self.sis_btn.setMinimumHeight(42)
        self.sis_btn.clicked.connect(self.cargar_sistema)
        
        self.sis_path = QLineEdit()
        self.sis_path.setReadOnly(True)
        self.sis_path.setPlaceholderText("Ruta del archivo VENTAS & NOTAS.xlsx...")
        self.sis_path.setMinimumHeight(42)
        
        sis_layout.addWidget(self.sis_path)
        sis_layout.addWidget(self.sis_btn)
        
        # Seccion de Archivo DIAN
        dian_layout = QHBoxLayout()
        self.dian_btn = QPushButton(" Cargar Excel DIAN ")
        self.dian_btn.setProperty("class", "SecondaryButton")
        self.dian_btn.setMinimumHeight(42)
        self.dian_btn.clicked.connect(self.cargar_dian)
        
        self.dian_path = QLineEdit()
        self.dian_path.setReadOnly(True)
        self.dian_path.setPlaceholderText("Ruta del archivo TOKEN DIAN.xlsx...")
        self.dian_path.setMinimumHeight(42)
        
        dian_layout.addWidget(self.dian_path)
        dian_layout.addWidget(self.dian_btn)

        # Seccion de Detalle DIAN (Opcional)
        dian_det_layout = QHBoxLayout()
        self.dian_det_btn = QPushButton(" Cargar Detalle DIAN (Opcional) ")
        self.dian_det_btn.setProperty("class", "SecondaryButton")
        self.dian_det_btn.setMinimumHeight(42)
        self.dian_det_btn.clicked.connect(self.cargar_dian_detalle)
        
        self.dian_det_path = QLineEdit()
        self.dian_det_path.setReadOnly(True)
        self.dian_det_path.setPlaceholderText("Para comparar transporte (Resumen_Facturas_Completo.xlsx)...")
        self.dian_det_path.setMinimumHeight(42)
        
        dian_det_layout.addWidget(self.dian_det_path)
        dian_det_layout.addWidget(self.dian_det_btn)
        
        layout.addLayout(sis_layout)
        layout.addSpacing(8)
        layout.addLayout(dian_layout)
        layout.addSpacing(8)
        layout.addLayout(dian_det_layout)
        
        # Botones Procesar y Limpiar
        layout.addSpacing(14)
        btn_box = QHBoxLayout()
        btn_box.setSpacing(12)
        btn_box.addStretch()

        self.procesar_btn = QPushButton("Procesar y Generar Reporte")
        self.procesar_btn.setProperty("class", "PrimaryButton")
        self.procesar_btn.setMinimumHeight(48)
        self.procesar_btn.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        self.procesar_btn.clicked.connect(self.procesar)
        self.procesar_btn.setEnabled(False)

        self.limpiar_btn = QPushButton("  🧹 Limpiar Todo  ")
        self.limpiar_btn.setProperty("class", "DangerButton")
        self.limpiar_btn.setMinimumHeight(48)
        self.limpiar_btn.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
        self.limpiar_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.limpiar_btn.setToolTip("Limpiar todas las rutas cargadas y resultados para empezar de cero")
        self.limpiar_btn.clicked.connect(self.limpiar_todo)

        btn_box.addWidget(self.procesar_btn)
        btn_box.addWidget(self.limpiar_btn)
        btn_box.addStretch()

        layout.addLayout(btn_box)
        layout.addSpacing(14)

        # Resultados
        self.result_text = QTextEdit()
        self.result_text.setObjectName("ResultText")
        self.result_text.setReadOnly(True)
        self.result_text.setPlaceholderText("Resultados apareceran aqui...")
        
        layout.addWidget(self.result_text)
        
        self.setLayout(layout)
        
    def validar_botones(self):
        if self.ruta_sistema and self.ruta_dian:
            self.procesar_btn.setEnabled(True)
            
    def cargar_sistema(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo del Sistema", "", "Excel Files (*.xls *.xlsx)")
        if file:
            self.ruta_sistema = file
            self.sis_path.setText(file)
            self.validar_botones()
            
    def cargar_dian(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo de la DIAN", "", "Excel Files (*.xls *.xlsx)")
        if file:
            self.ruta_dian = file
            self.dian_path.setText(file)
            self.validar_botones()

    def cargar_dian_detalle(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar detalle de la DIAN", "", "Excel Files (*.xls *.xlsx)")
        if file:
            self.ruta_dian_detalle = file
            self.dian_det_path.setText(file)

    def limpiar_todo(self):
        """Limpia todas las rutas cargadas, campos de texto y resultados."""
        self.ruta_sistema = ""
        self.ruta_dian = ""
        self.ruta_dian_detalle = ""
        self.ruta_salida = ""
        self.sis_path.clear()
        self.dian_path.clear()
        self.dian_det_path.clear()
        if hasattr(self, "ruta_referencias"):
            self.ruta_referencias = None
        if hasattr(self, "ref_path"):
            self.ref_path.clear()
        self.result_text.clear()
        self.procesar_btn.setEnabled(False)
            
    def procesar(self):
        # Preguntar donde guardar el resultado
        self.ruta_salida, _ = QFileDialog.getSaveFileName(self, "Guardar Reporte", "Reporte_Conciliador_Ventas.xlsx", "Excel Files (*.xlsx)")
        if not self.ruta_salida:
            return
            
        self.result_text.append("Iniciando procesamiento...\nPor favor espera, esto puede tomar unos segundos.")
        self.procesar_btn.setEnabled(False)
        
        self.thread = WorkerThread(self.ruta_sistema, self.ruta_dian, self.ruta_salida, self.ruta_dian_detalle, "Ventas")
        self.thread.finished.connect(self.on_success)
        self.thread.error.connect(self.on_error)
        self.thread.start()

    def on_success(self, stats):
        self.procesar_btn.setEnabled(True)
        msg = (
            f"Exito!\n\n"
            f"Estadisticas:\n"
            f"   Facturas cruzadas: {stats['total_procesadas']}\n"
            f"   Cuadre Perfecto: {stats['cuadre']}\n"
            f"   Diferencias: {stats['diferencias']}\n\n"
            f"Guardado en:\n{self.ruta_salida}"
        )
        self.result_text.setPlainText(msg)
        os.startfile(self.ruta_salida) # type: ignore

    def on_error(self, err_msg):
        self.procesar_btn.setEnabled(True)
        self.result_text.setPlainText(f"Ocurrio un error:\n\n{err_msg}")
        QMessageBox.critical(self, "Error", f"Fallo al procesar: {err_msg}")

class ComprasWidget(VentasWidget):
    def init_ui(self):
        super().init_ui()
        title = self.findChild(QLabel, "TitleLabel")
        if title: title.setText("Conciliacion de Compras / Gastos")
        subtitle = self.findChild(QLabel, "SubtitleLabel")
        if subtitle: subtitle.setText("Compara las Facturas Recibidas en ERP vs Token DIAN (Recibidos)")
        
        self.ruta_referencias = None
        
        # Ocultar detalle DIAN (es de ventas)
        self.dian_det_btn.setVisible(False)
        self.dian_det_path.setVisible(False)
        
        # Crear UI para Referencias
        ref_layout = QHBoxLayout()
        self.ref_btn = QPushButton(" Cargar Excel Referencias (Odoo) ")
        self.ref_btn.setProperty("class", "SecondaryButton")
        self.ref_btn.setMinimumHeight(42)
        self.ref_btn.clicked.connect(self.cargar_referencias)
        
        self.ref_path = QLineEdit()
        self.ref_path.setReadOnly(True)
        self.ref_path.setPlaceholderText("Para asociar Referencia y Centro de Costo (Opcional)...")
        self.ref_path.setMinimumHeight(42)
        
        ref_layout.addWidget(self.ref_path)
        ref_layout.addWidget(self.ref_btn)
        
        # Insertar justo antes del boton procesar
        layout = self.layout()
        layout.insertLayout(5, ref_layout)
        layout.insertSpacing(5, 8)
        
    def cargar_referencias(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo de Referencias", "", "Excel Files (*.xls *.xlsx)")
        if file:
            self.ruta_referencias = file
            self.ref_path.setText(file)

    def procesar(self):
        self.ruta_salida, _ = QFileDialog.getSaveFileName(self, "Guardar Reporte", "Reporte_Conciliador_Compras.xlsx", "Excel Files (*.xlsx)")
        if not self.ruta_salida: return
        self.result_text.append("Iniciando procesamiento de compras...")
        self.procesar_btn.setEnabled(False)
        self.thread = WorkerThread(self.ruta_sistema, self.ruta_dian, self.ruta_salida, None, "Compras", ruta_referencias=self.ruta_referencias)
        self.thread.finished.connect(self.on_success)
        self.thread.error.connect(self.on_error)
        self.thread.start()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Conciliador Financiero v{auto_updater.get_local_version()}")
        self.setMinimumSize(1024, 680)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # ── Sidebar ────────────────────────────────────────────
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(220)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)

        # Logo / Brand area
        logo_container = QFrame()
        logo_container.setStyleSheet(
            "QFrame { background-color: #0b1219; border-bottom: 1px solid #131e29; }"
        )
        logo_layout = QVBoxLayout(logo_container)
        logo_layout.setContentsMargins(14, 20, 14, 18)
        logo_layout.setSpacing(4)

        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        from PyQt6.QtGui import QPixmap # type: ignore
        img_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
        pixmap = QPixmap(img_path)
        if not pixmap.isNull():
            logo_label.setPixmap(pixmap.scaled(
                QSize(120, 120),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            ))
            logo_layout.addWidget(logo_label)
        else:
            logo_text = QLabel("CONCILIADOR")
            logo_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
            logo_text.setStyleSheet(
                "color: #ffffff; font-size: 18px; font-weight: 800; letter-spacing: 2px; padding: 0px;"
            )
            logo_layout.addWidget(logo_text)

        tag_label = QLabel("FINANCIERO")
        tag_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tag_label.setStyleSheet(
            "color: #3a5565; font-size: 9px; font-weight: 600; letter-spacing: 3px; padding: 0px;"
        )
        logo_layout.addWidget(tag_label)

        sidebar_layout.addWidget(logo_container)

        # Navigation section label
        nav_label = QLabel("  NAVEGACION")
        nav_label.setStyleSheet(
            "color: #2a3e4e; font-size: 9px; font-weight: 700; "
            "letter-spacing: 2px; padding: 14px 0px 6px 0px;"
        )
        sidebar_layout.addWidget(nav_label)

        # Nav buttons
        self.btn_ventas = QPushButton("  Ventas")
        self.btn_ventas.setCheckable(True)
        self.btn_ventas.setChecked(True)
        self.btn_ventas.setMinimumHeight(46)

        self.btn_compras = QPushButton("  Compras")
        self.btn_compras.setCheckable(True)
        self.btn_compras.setMinimumHeight(46)

        self.btn_vendedores = QPushButton("  Vendedores")
        self.btn_vendedores.setCheckable(True)
        self.btn_vendedores.setMinimumHeight(46)
        
        sidebar_layout.addWidget(self.btn_ventas)
        sidebar_layout.addWidget(self.btn_compras)
        sidebar_layout.addWidget(self.btn_vendedores)
        sidebar_layout.addStretch()

        # Bottom section
        bottom_line = QFrame()
        bottom_line.setStyleSheet("background-color: #131e29; max-height: 1px;")
        bottom_line.setFixedHeight(1)
        sidebar_layout.addWidget(bottom_line)

        self.btn_config = QPushButton("  Configuracion")
        self.btn_config.setObjectName("BtnConfig")
        self.btn_config.setMinimumHeight(44)
        self.btn_config.setStyleSheet(
            "QPushButton {"
            "  background-color: transparent; color: #4a6070;"
            "  text-align: left; padding: 12px 20px; font-size: 13px;"
            "  border: none; border-top: 1px solid #131e29;"
            "}"
            "QPushButton:hover { color: #7faabb; background-color: #0e1820; }"
        )
        self.btn_config.clicked.connect(self.abrir_configuracion)
        sidebar_layout.addWidget(self.btn_config)

        self.btn_update = QPushButton(f"  🌐 v{auto_updater.get_local_version()} | Actualizar")
        self.btn_update.setObjectName("BtnUpdateSidebar")
        self.btn_update.setMinimumHeight(44)
        self.btn_update.setStyleSheet(
            "QPushButton {"
            "  background-color: transparent; color: #00d4aa;"
            "  text-align: left; padding: 12px 20px; font-size: 12px; font-weight: bold;"
            "  border: none; border-top: 1px solid #131e29;"
            "}"
            "QPushButton:hover { color: #ffffff; background-color: #0c2b2b; }"
        )
        self.btn_update.clicked.connect(lambda: auto_updater.check_for_updates_threaded(self, silent=False))
        sidebar_layout.addWidget(self.btn_update)

        # ── Stack ──────────────────────────────────────────────
        self.stack = QStackedWidget()
        self.ventas_widget = VentasWidget()
        self.compras_widget = ComprasWidget()
        self.vendedores_widget = VendedoresWidget()
        self.stack.addWidget(self.ventas_widget)
        self.stack.addWidget(self.compras_widget)
        self.stack.addWidget(self.vendedores_widget)
        
        self.btn_ventas.clicked.connect(self.show_ventas)
        self.btn_compras.clicked.connect(self.show_compras)
        self.btn_vendedores.clicked.connect(self.show_vendedores)
        
        main_layout.addWidget(sidebar)
        main_layout.addWidget(self.stack)
        self.load_styles()
        QTimer.singleShot(3500, lambda: auto_updater.check_for_updates_threaded(self, silent=True))

    def load_styles(self):
        style_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "styles.qss")
        if os.path.exists(style_path):
            with open(style_path, "r", encoding='utf-8') as f:
                self.setStyleSheet(f.read())

    def show_ventas(self):
        self.stack.setCurrentIndex(0)
        self.btn_ventas.setChecked(True)
        self.btn_compras.setChecked(False)
        self.btn_vendedores.setChecked(False)
        
    def show_compras(self):
        self.stack.setCurrentIndex(1)
        self.btn_compras.setChecked(True)
        self.btn_ventas.setChecked(False)
        self.btn_vendedores.setChecked(False)

    def show_vendedores(self):
        self.stack.setCurrentIndex(2)
        self.btn_vendedores.setChecked(True)
        self.btn_ventas.setChecked(False)
        self.btn_compras.setChecked(False)

    def abrir_configuracion(self):
        try:
            from ui.settings_dialog import SettingsDialog # type: ignore
            dialog = SettingsDialog(self)
            dialog.exec()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo abrir la ventana de configuracion: {e}")
