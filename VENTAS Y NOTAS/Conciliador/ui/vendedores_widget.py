import os
from PyQt6.QtWidgets import (  # type: ignore
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QLineEdit, QTextEdit, QFileDialog, QMessageBox, QFrame,
    QGraphicsDropShadowEffect, QDialog, QListWidget, QGroupBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize  # type: ignore
from PyQt6.QtGui import QFont, QColor  # type: ignore

try:
    from core.vendedores_logic import (  # type: ignore
        procesar_dashboard_vendedores,
        procesar_dashboard_nc_asiento,
        procesar_informe_ventas_asiento,
        _leer_excel_inteligente,
        cargar_config_vendedores,
        guardar_config_vendedores
    )
except ImportError:
    import sys
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from core.vendedores_logic import (  # type: ignore
        procesar_dashboard_vendedores,
        procesar_dashboard_nc_asiento,
        procesar_informe_ventas_asiento,
        _leer_excel_inteligente,
        cargar_config_vendedores,
        guardar_config_vendedores
    )


CARD_STYLE = (
    "QFrame {"
    "  background-color: #111d29;"
    "  border: 1px solid #1a2736;"
    "  border-radius: 12px;"
    "  padding: 0px;"
    "}"
)

SECTION_HEADER_STYLE = (
    "color: #4fc3f7;"
    "font-size: 11px;"
    "font-weight: 700;"
    "letter-spacing: 2px;"
    "text-transform: uppercase;"
    "padding: 0px;"
    "margin-bottom: 2px;"
)

FILE_ICON_STYLE = (
    "color: #1976d2;"
    "font-size: 20px;"
    "padding: 0px;"
)

STATUS_DOT_GREEN = (
    "display: inline-block;"
    "width: 8px; height: 8px;"
    "border-radius: 4px;"
    "background-color: #4caf50;"
)

INFO_STYLE = (
    "color: #4a6572;"
    "font-size: 11px;"
    "font-style: italic;"
)

PLACEHOLDER_STYLE = (
    "color: #2d4050;"
    "font-size: 12px;"
)


class ConfigVendedoresDialog(QDialog):
    """Diálogo modal para configurar qué vendedores pertenecen a Repuestos y a Vehículos Jaime en tiempo real."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configuración de Vendedores Especiales")
        self.setMinimumSize(700, 530)
        self.setStyleSheet("""
            QDialog {
                background-color: #0b131b;
                color: #e2e8f0;
            }
            QGroupBox {
                border: 1px solid #1e293b;
                border-radius: 8px;
                margin-top: 14px;
                font-weight: bold;
                color: #38bdf8;
                padding: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 4px;
            }
            QListWidget {
                background-color: #111d29;
                border: 1px solid #1e293b;
                border-radius: 6px;
                color: #f1f5f9;
                font-size: 12px;
                padding: 4px;
            }
            QListWidget::item {
                padding: 6px;
                border-radius: 4px;
            }
            QListWidget::item:selected {
                background-color: #0284c7;
                color: #ffffff;
            }
            QLineEdit {
                background-color: #111d29;
                border: 1px solid #334155;
                border-radius: 6px;
                color: #ffffff;
                padding: 6px 10px;
                font-size: 12px;
            }
            QLineEdit:focus {
                border: 1px solid #38bdf8;
            }
            QPushButton {
                background-color: #1e293b;
                border: 1px solid #334155;
                color: #e2e8f0;
                border-radius: 6px;
                padding: 6px 14px;
                font-size: 12px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #334155;
                color: #38bdf8;
            }
            QPushButton#btnGuardar {
                background-color: #059669;
                border: 1px solid #10b981;
                color: #ffffff;
                font-weight: bold;
            }
            QPushButton#btnGuardar:hover {
                background-color: #10b981;
            }
            QPushButton#btnEliminar {
                background-color: #450a0a;
                border: 1px solid #991b1b;
                color: #fca5a5;
            }
            QPushButton#btnEliminar:hover {
                background-color: #991b1b;
                color: #ffffff;
            }
        """)
        self.init_ui()
        self.cargar_datos()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(14)

        header_lbl = QLabel("⚙️ Reglas de Clasificación de Vendedores en Tiempo Real")
        header_lbl.setStyleSheet("font-size: 15px; font-weight: bold; color: #38bdf8;")
        sub_lbl = QLabel(
            "Configura qué vendedores se agrupan en las columnas exclusivas de 'REPUESTOS' o 'VEHÍCULOS JAIME'.\n"
            "Todos los cambios se aplican de inmediato en los informes por Centro de Costo (Sede)."
        )
        sub_lbl.setStyleSheet("color: #94a3b8; font-size: 12px;")
        sub_lbl.setWordWrap(True)
        layout.addWidget(header_lbl)
        layout.addWidget(sub_lbl)

        # Contenedor de dos columnas
        cols_layout = QHBoxLayout()
        cols_layout.setSpacing(16)

        # Columna 1: REPUESTOS
        grp_rep = QGroupBox("🛠️ Vendedores de REPUESTOS")
        rep_layout = QVBoxLayout(grp_rep)
        self.list_rep = QListWidget()
        rep_layout.addWidget(self.list_rep)

        add_rep_layout = QHBoxLayout()
        self.txt_rep = QLineEdit()
        self.txt_rep.setPlaceholderText("Nombre del vendedor...")
        btn_add_rep = QPushButton("➕ Agregar")
        btn_add_rep.clicked.connect(self.agregar_repuesto)
        self.txt_rep.returnPressed.connect(self.agregar_repuesto)
        add_rep_layout.addWidget(self.txt_rep)
        add_rep_layout.addWidget(btn_add_rep)
        rep_layout.addLayout(add_rep_layout)

        btn_del_rep = QPushButton("🗑️ Eliminar Seleccionado")
        btn_del_rep.setObjectName("btnEliminar")
        btn_del_rep.clicked.connect(lambda: self.eliminar_item(self.list_rep))
        rep_layout.addWidget(btn_del_rep)
        cols_layout.addWidget(grp_rep)

        # Columna 2: VEHÍCULOS JAIME
        grp_jaime = QGroupBox("🚗 Vendedores VEHÍCULOS JAIME")
        jaime_layout = QVBoxLayout(grp_jaime)
        self.list_jaime = QListWidget()
        jaime_layout.addWidget(self.list_jaime)

        add_jaime_layout = QHBoxLayout()
        self.txt_jaime = QLineEdit()
        self.txt_jaime.setPlaceholderText("Nombre del vendedor...")
        btn_add_jaime = QPushButton("➕ Agregar")
        btn_add_jaime.clicked.connect(self.agregar_jaime)
        self.txt_jaime.returnPressed.connect(self.agregar_jaime)
        add_jaime_layout.addWidget(self.txt_jaime)
        add_jaime_layout.addWidget(btn_add_jaime)
        jaime_layout.addLayout(add_jaime_layout)

        btn_del_jaime = QPushButton("🗑️ Eliminar Seleccionado")
        btn_del_jaime.setObjectName("btnEliminar")
        btn_del_jaime.clicked.connect(lambda: self.eliminar_item(self.list_jaime))
        jaime_layout.addWidget(btn_del_jaime)
        cols_layout.addWidget(grp_jaime)

        layout.addLayout(cols_layout)

        # Botonera Inferior
        bottom_layout = QHBoxLayout()
        btn_reset = QPushButton("🔄 Restaurar Predeterminados")
        btn_reset.setToolTip("Restaurar lista por defecto (Mailyn y Liliana en Repuestos; Jaime Orozco en Vehículos Jaime)")
        btn_reset.clicked.connect(self.restaurar_defaults)
        bottom_layout.addWidget(btn_reset)

        bottom_layout.addStretch()

        btn_cancelar = QPushButton("Cancelar")
        btn_cancelar.clicked.connect(self.reject)
        bottom_layout.addWidget(btn_cancelar)

        btn_guardar = QPushButton("💾 Guardar y Aplicar")
        btn_guardar.setObjectName("btnGuardar")
        btn_guardar.clicked.connect(self.guardar)
        bottom_layout.addWidget(btn_guardar)

        layout.addLayout(bottom_layout)

    def cargar_datos(self):
        cfg = cargar_config_vendedores()
        self.list_rep.clear()
        for v in cfg.get("vendedores_repuestos", []):
            self.list_rep.addItem(v.strip().upper())

        self.list_jaime.clear()
        for v in cfg.get("vendedores_vehiculos_jaime", []):
            self.list_jaime.addItem(v.strip().upper())

    def agregar_repuesto(self):
        nom = self.txt_rep.text().strip().upper()
        if nom:
            items = [self.list_rep.item(i).text() for i in range(self.list_rep.count())]
            if nom not in items:
                self.list_rep.addItem(nom)
            self.txt_rep.clear()

    def agregar_jaime(self):
        nom = self.txt_jaime.text().strip().upper()
        if nom:
            items = [self.list_jaime.item(i).text() for i in range(self.list_jaime.count())]
            if nom not in items:
                self.list_jaime.addItem(nom)
            self.txt_jaime.clear()

    def eliminar_item(self, list_widget):
        for item in list_widget.selectedItems():
            list_widget.takeItem(list_widget.row(item))

    def restaurar_defaults(self):
        res = QMessageBox.question(
            self, "Confirmar restauración",
            "¿Desea restaurar los vendedores predeterminados?\n\n"
            "• Repuestos: MAILYN YORGELY CONTRERAS CARDENAS, LILIANA DEL PILAR MENDOZA GUEVARA\n"
            "• Vehículos Jaime: JAIME EDUARDO  OROZCO CHAVARRO",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if res == QMessageBox.StandardButton.Yes:
            self.list_rep.clear()
            self.list_rep.addItem("MAILYN YORGELY CONTRERAS CARDENAS")
            self.list_rep.addItem("LILIANA DEL PILAR MENDOZA GUEVARA")

            self.list_jaime.clear()
            self.list_jaime.addItem("JAIME EDUARDO  OROZCO CHAVARRO")

    def guardar(self):
        reps = [self.list_rep.item(i).text().strip().upper() for i in range(self.list_rep.count()) if self.list_rep.item(i).text().strip()]
        jaimes = [self.list_jaime.item(i).text().strip().upper() for i in range(self.list_jaime.count()) if self.list_jaime.item(i).text().strip()]

        nueva_cfg = {
            "vendedores_repuestos": reps,
            "vendedores_vehiculos_jaime": jaimes
        }
        if guardar_config_vendedores(nueva_cfg):
            QMessageBox.information(
                self, "Configuración Guardada",
                "Las reglas de clasificación de vendedores se han guardado y aplicado correctamente."
            )
            self.accept()
        else:
            QMessageBox.warning(self, "Error", "No se pudo guardar la configuración.")


class VendedorWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, ruta_archivo, ruta_salida, ruta_asiento_nc=None, ruta_dian=None):
        super().__init__()
        self.ruta_archivo = ruta_archivo
        self.ruta_salida = ruta_salida
        self.ruta_asiento_nc = ruta_asiento_nc
        self.ruta_dian = ruta_dian

    def run(self):
        try:
            stats = procesar_dashboard_vendedores(self.ruta_archivo, self.ruta_salida,
                                                  ruta_asiento_nc=self.ruta_asiento_nc,
                                                  ruta_dian=self.ruta_dian)
            self.finished.emit(stats)
        except Exception as e:
            self.error.emit(str(e))


class NotasCreditoWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, ruta_asiento, ruta_salida):
        super().__init__()
        self.ruta_asiento = ruta_asiento
        self.ruta_salida = ruta_salida

    def run(self):
        try:
            stats = procesar_dashboard_nc_asiento(self.ruta_asiento, self.ruta_salida)
            self.finished.emit(stats)
        except Exception as e:
            self.error.emit(str(e))


class InformeVentasWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, ruta_archivo, ruta_salida, ruta_asiento_nc=None, ruta_tb=None, ruta_gastos=None):
        super().__init__()
        self.ruta_archivo = ruta_archivo
        self.ruta_salida = ruta_salida
        self.ruta_asiento_nc = ruta_asiento_nc
        self.ruta_tb = ruta_tb
        self.ruta_gastos = ruta_gastos

    def run(self):
        try:
            stats = procesar_informe_ventas_asiento(
                self.ruta_archivo, self.ruta_salida,
                ruta_asiento_nc=self.ruta_asiento_nc,
                ruta_tb=self.ruta_tb,
                ruta_gastos=self.ruta_gastos
            )
            self.finished.emit(stats)
        except Exception as e:
            self.error.emit(str(e))


class VendedoresWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.ruta_archivo = ""
        self.ruta_salida = ""
        self.ruta_salida_ventas = ""
        self.ruta_asiento_nc = ""
        self.ruta_salida_nc = ""
        self.ruta_dian = ""
        self.ruta_tb = ""
        self.ruta_gastos = ""
        self.init_ui()

    def init_ui(self):
        root = QVBoxLayout()
        root.setContentsMargins(32, 28, 32, 28)
        root.setSpacing(0)

        # ── Header ─────────────────────────────────────────────
        title = QLabel("Dashboard de Vendedores")
        title.setObjectName("TitleLabel")

        subtitle = QLabel(
            "Analisis de rendimiento por vendedor, cliente, almacen y factura"
        )
        subtitle.setObjectName("SubtitleLabel")

        root.addWidget(title)
        root.addSpacing(2)
        root.addWidget(subtitle)
        root.addSpacing(16)

        # ── Input Card ─────────────────────────────────────────
        input_card = QFrame()
        input_card.setStyleSheet(CARD_STYLE)
        input_card_layout = QVBoxLayout(input_card)
        input_card_layout.setContentsMargins(22, 20, 22, 22)
        input_card_layout.setSpacing(12)

        # Section label
        section_lbl = QLabel("ARCHIVO DE ENTRADA")
        section_lbl.setStyleSheet(SECTION_HEADER_STYLE)
        input_card_layout.addWidget(section_lbl)

        # Description
        desc_lbl = QLabel(
            "Selecciona el archivo de Ventas ('informe ventas.xlsx' / Asiento Contable o 'Pedido de venta.xlsx')"
        )
        desc_lbl.setStyleSheet("color: #6a8090; font-size: 12px;")
        desc_lbl.setWordWrap(True)
        input_card_layout.addWidget(desc_lbl)

        input_card_layout.addSpacing(4)

        # File selector row
        input_row = QHBoxLayout()
        input_row.setSpacing(10)

        self.file_path = QLineEdit()
        self.file_path.setReadOnly(True)
        self.file_path.setPlaceholderText(
            "Asiento Contable de Ventas o Pedido de Venta..."
        )
        self.file_path.setMinimumHeight(42)

        self.file_btn = QPushButton("  Examinar...  ")
        self.file_btn.setProperty("class", "SecondaryButton")
        self.file_btn.setMinimumHeight(42)
        self.file_btn.setMinimumWidth(130)
        self.file_btn.clicked.connect(self.cargar_archivo)

        input_row.addWidget(self.file_path)
        input_row.addWidget(self.file_btn)
        input_card_layout.addLayout(input_row)

        # Info label (file name, size)
        self.info_label = QLabel("")
        self.info_label.setStyleSheet(INFO_STYLE)
        input_card_layout.addWidget(self.info_label)

        ncc_row = QHBoxLayout()
        ncc_row.setSpacing(10)
        ncc_lbl = QLabel("Notas Crédito:")
        ncc_lbl.setStyleSheet("color: #4fc3f7; font-size: 11px; font-weight: bold;")
        ncc_lbl.setFixedWidth(100)
        self.ncc_file_path = QLineEdit()
        self.ncc_file_path.setReadOnly(True)
        self.ncc_file_path.setPlaceholderText("Asiento contable (opcional, para restar Notas Crédito)...")
        self.ncc_file_path.setMinimumHeight(32)
        self.ncc_file_btn = QPushButton("  Examinar NCC...  ")
        self.ncc_file_btn.setProperty("class", "SecondaryButton")
        self.ncc_file_btn.setMinimumHeight(32)
        self.ncc_file_btn.setMinimumWidth(120)
        self.ncc_file_btn.clicked.connect(self.cargar_archivo_nc)
        ncc_row.addWidget(ncc_lbl)
        ncc_row.addWidget(self.ncc_file_path)
        ncc_row.addWidget(self.ncc_file_btn)
        input_card_layout.addLayout(ncc_row)

        dian_row = QHBoxLayout()
        dian_row.setSpacing(10)
        dian_lbl = QLabel("Token / DIAN:")
        dian_lbl.setStyleSheet("color: #00e676; font-size: 11px; font-weight: bold;")
        dian_lbl.setFixedWidth(100)
        self.dian_file_path = QLineEdit()
        self.dian_file_path.setReadOnly(True)
        self.dian_file_path.setPlaceholderText("Excel descargado con el Token DIAN (fcbf...xlsx)...")
        self.dian_file_path.setMinimumHeight(32)
        self.dian_file_btn = QPushButton("  Examinar DIAN...  ")
        self.dian_file_btn.setProperty("class", "SecondaryButton")
        self.dian_file_btn.setMinimumHeight(32)
        self.dian_file_btn.setMinimumWidth(120)
        self.dian_file_btn.clicked.connect(self.cargar_archivo_dian)
        dian_row.addWidget(dian_lbl)
        dian_row.addWidget(self.dian_file_path)
        dian_row.addWidget(self.dian_file_btn)
        input_card_layout.addLayout(dian_row)

        tb_row = QHBoxLayout()
        tb_row.setSpacing(10)
        tb_lbl = QLabel("Balance Odoo:")
        tb_lbl.setStyleSheet("color: #ffb74d; font-size: 11px; font-weight: bold;")
        tb_lbl.setFixedWidth(100)
        self.tb_file_path = QLineEdit()
        self.tb_file_path.setReadOnly(True)
        self.tb_file_path.setPlaceholderText("Trial Balance / Balance de Comprobación Odoo (opcional, para Otros Ingresos)...")
        self.tb_file_path.setMinimumHeight(32)
        self.tb_file_btn = QPushButton("  Examinar Balance...  ")
        self.tb_file_btn.setProperty("class", "SecondaryButton")
        self.tb_file_btn.setMinimumHeight(32)
        self.tb_file_btn.setMinimumWidth(130)
        self.tb_file_btn.clicked.connect(self.cargar_archivo_tb)
        tb_row.addWidget(tb_lbl)
        tb_row.addWidget(self.tb_file_path)
        tb_row.addWidget(self.tb_file_btn)
        input_card_layout.addLayout(tb_row)

        gastos_row = QHBoxLayout()
        gastos_row.setSpacing(10)
        gastos_lbl = QLabel("Gastos CC:")
        gastos_lbl.setStyleSheet("color: #ba68c8; font-size: 11px; font-weight: bold;")
        gastos_lbl.setFixedWidth(100)
        self.gastos_file_path = QLineEdit()
        self.gastos_file_path.setReadOnly(True)
        self.gastos_file_path.setPlaceholderText("Gastos por Centro de Costos / Sedes (opcional, para Estado de Resultados y Gastos por CC)...")
        self.gastos_file_path.setMinimumHeight(32)
        self.gastos_file_btn = QPushButton("  Examinar Gastos...  ")
        self.gastos_file_btn.setProperty("class", "SecondaryButton")
        self.gastos_file_btn.setMinimumHeight(32)
        self.gastos_file_btn.setMinimumWidth(130)
        self.gastos_file_btn.clicked.connect(self.cargar_archivo_gastos)
        gastos_row.addWidget(gastos_lbl)
        gastos_row.addWidget(self.gastos_file_path)
        gastos_row.addWidget(self.gastos_file_btn)
        input_card_layout.addLayout(gastos_row)

        root.addWidget(input_card)
        root.addSpacing(16)

        # ── Buttons (Generate & Reset) ─────────────────────────
        btn_container = QFrame()
        btn_container.setStyleSheet("background: transparent; border: none;")
        btn_layout = QHBoxLayout(btn_container)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(12)
        btn_layout.addStretch()

        self.procesar_ventas_btn = QPushButton("   📊 Generar Informe de Ventas (Centros de Costo)   ")
        self.procesar_ventas_btn.setStyleSheet(
            "QPushButton {"
            "  background-color: #059669; color: #ffffff; border-radius: 8px; font-weight: bold; font-size: 12px;"
            "  padding: 10px 16px; border: 1px solid #10b981;"
            "}"
            "QPushButton:hover { background-color: #10b981; }"
            "QPushButton:disabled { background-color: #1a2f2b; color: #3a5c53; border-color: #1a2f2b; }"
        )
        self.procesar_ventas_btn.setMinimumHeight(48)
        self.procesar_ventas_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.procesar_ventas_btn.setToolTip("Genera informe estructurado de Ventas por Centro de Costo y Vendedor (Exentas vs Gravadas)")
        self.procesar_ventas_btn.clicked.connect(self.procesar_ventas)
        self.procesar_ventas_btn.setEnabled(False)

        self.config_btn = QPushButton("   ⚙️ Reglas Vendedores (Repuestos / Jaime)   ")
        self.config_btn.setProperty("class", "SecondaryButton")
        self.config_btn.setMinimumHeight(48)
        self.config_btn.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        self.config_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.config_btn.setToolTip("Configurar qué vendedores se asignan a Repuestos y cuáles a Vehículos Jaime en tiempo real")
        self.config_btn.clicked.connect(self.abrir_config_vendedores)

        self.procesar_nc_btn = QPushButton("   📄 Reporte Notas Crédito   ")
        self.procesar_nc_btn.setProperty("class", "SecondaryButton")
        self.procesar_nc_btn.setMinimumHeight(48)
        self.procesar_nc_btn.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
        self.procesar_nc_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.procesar_nc_btn.setToolTip("Genera un dashboard exclusivo de Notas Crédito desde el Asiento Contable, recuperando vendedores por cliente")
        self.procesar_nc_btn.clicked.connect(self.procesar_nc)
        self.procesar_nc_btn.setEnabled(False)

        self.procesar_btn = QPushButton("   🎯 Dashboard Vendedores (Pedidos)   ")
        self.procesar_btn.setProperty("class", "PrimaryButton")
        self.procesar_btn.setMinimumHeight(48)
        self.procesar_btn.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
        self.procesar_btn.clicked.connect(self.procesar)
        self.procesar_btn.setEnabled(False)

        self.limpiar_btn = QPushButton("   🧹 Limpiar Todo   ")
        self.limpiar_btn.setProperty("class", "DangerButton")
        self.limpiar_btn.setMinimumHeight(48)
        self.limpiar_btn.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
        self.limpiar_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.limpiar_btn.setToolTip("Limpiar todas las rutas, campos y resultados para empezar desde cero")
        self.limpiar_btn.clicked.connect(self.limpiar_todo)

        btn_layout.addWidget(self.procesar_ventas_btn)
        btn_layout.addWidget(self.config_btn)
        btn_layout.addWidget(self.procesar_nc_btn)
        btn_layout.addWidget(self.procesar_btn)
        btn_layout.addWidget(self.limpiar_btn)
        btn_layout.addStretch()
        root.addWidget(btn_container)
        root.addSpacing(12)

        # ── Results Card ───────────────────────────────────────
        results_card = QFrame()
        results_card.setStyleSheet(CARD_STYLE)
        results_card_layout = QVBoxLayout(results_card)
        results_card_layout.setContentsMargins(22, 20, 22, 22)
        results_card_layout.setSpacing(10)

        section_lbl2 = QLabel("RESULTADOS")
        section_lbl2.setStyleSheet(SECTION_HEADER_STYLE)
        results_card_layout.addWidget(section_lbl2)

        self.result_text = QTextEdit()
        self.result_text.setObjectName("ResultText")
        self.result_text.setReadOnly(True)
        self.result_text.setMinimumHeight(220)
        self.result_text.setPlaceholderText(
            "El dashboard generado incluira:\n\n"
            "  1. Resumen Ejecutivo  -  KPIs + Ranking de vendedores\n"
            "  2. Clientes por Vendedor  -  Almacen, Factura, Producto, Total\n"
            "  3. Detalle Completo  -  Todos los pedidos con todas las columnas\n"
            "  4. Por Almacen  -  Distribucion de pedidos por punto de venta\n"
            "  5. Top Productos  -  Los 10 mas vendidos\n"
            "  6. Evolucion Semanal  -  Tendencia temporal\n\n"
            "Carga un archivo y haz clic en 'Generar Dashboard'..."
        )
        results_card_layout.addWidget(self.result_text)

        root.addWidget(results_card)
        root.addStretch()

        self.setLayout(root)

    def cargar_archivo(self):
        file, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo de Ventas (Asiento Contable o Pedidos)",
            "",
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )
        if file:
            self.ruta_archivo = file
            self.file_path.setText(file)

            fname = os.path.basename(file)
            fsize = os.path.getsize(file) / 1024
            self.info_label.setText(
                f"  {fname}   |   {fsize:.1f} KB"
            )

            self.procesar_ventas_btn.setEnabled(True)
            self.procesar_btn.setEnabled(True)

            diag_str = f"📁 Archivo cargado: {fname} ({fsize:.1f} KB)\n"
            try:
                df_prev = _leer_excel_inteligente(file)
                total_rows = len(df_prev)
                cols_low = [str(c).lower().strip() for c in df_prev.columns]
                
                tiene_producto = any('roducto' in c for c in cols_low)
                tiene_cc = any('costo' in c or 'sede' in c for c in cols_low)
                
                col_num = next((c for c in df_prev.columns if any(kw in str(c).lower() for kw in ['número', 'numero', 'factura', 'referencia'])), None)
                n_fact = total_rows
                n_nc = 0
                if col_num:
                    s_nums = df_prev[col_num].astype(str)
                    n_nc = int(s_nums.str.upper().str.contains(r'^NC|^NCC').sum())
                    n_fact = total_rows - n_nc

                if tiene_producto:
                    diag_str += f"\n✅ Tipo detectado: Pedidos de Venta (sale.order) con {total_rows} filas."
                    diag_str += "\n   Listo para generar Dashboard Vendedores con margen y productos."
                elif tiene_cc:
                    diag_str += f"\n✅ Tipo detectado: Facturas y Asientos de Ventas ({total_rows} registros totales)."
                    if n_nc > 0:
                        diag_str += f"\n   • Facturas de venta emitidas: {n_fact}"
                        diag_str += f"\n   • Notas Crédito (NC/NCC): {n_nc}"
                    else:
                        diag_str += f"\n   • Facturas de venta emitidas: {n_fact}"
                    diag_str += "\n   Listo para clasificar por Centro de Costo (Exentas vs Gravadas)."
                else:
                    diag_str += f"\nℹ️ Registros detectados: {total_rows} filas."
            except Exception as e:
                diag_str += f"\nℹ️ Archivo verificado y listo para procesar."

            diag_str += (
                "\n\nPresione cualquiera de los botones de arriba para procesar:\n"
                "  • '📊 Generar Informe de Ventas (Centros de Costo)': Matriz Exentas vs Gravadas\n"
                "  • '🎯 Dashboard Vendedores (Pedidos)': Dashboard integral de facturación"
            )
            self.result_text.setPlainText(diag_str)

    def procesar_ventas(self):
        if not self.ruta_archivo:
            return
        self.ruta_salida_ventas, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar Informe de Ventas por Centro de Costo",
            "Informe_Ventas_Por_Centro_Costo.xlsx",
            "Archivos Excel (*.xlsx)"
        )
        if not self.ruta_salida_ventas:
            return

        self.result_text.setPlainText(
            "Procesando Informe de Ventas por Centro de Costo y Vendedor...\n"
            "• Validando y corrigiendo inconsistencias de sedes y vendedores\n"
            "• Clasificando en Ventas Exentas (0%) vs Gravadas (19%)\n"
            "• Generando matrices con fórmulas vivas de Excel y subtotales por sede...\n"
            "Esto puede tardar unos segundos."
        )
        self.procesar_ventas_btn.setEnabled(False)
        self.procesar_ventas_btn.setText("   Procesando Ventas...   ")

        self.thread_ventas = InformeVentasWorker(
            self.ruta_archivo,
            self.ruta_salida_ventas,
            ruta_asiento_nc=getattr(self, 'ruta_asiento_nc', None),
            ruta_tb=getattr(self, 'ruta_tb', None),
            ruta_gastos=getattr(self, 'ruta_gastos', None)
        )
        self.thread_ventas.finished.connect(self.on_success_ventas)
        self.thread_ventas.error.connect(self.on_error_ventas)
        self.thread_ventas.start()

    def on_success_ventas(self, stats):
        self.procesar_ventas_btn.setEnabled(True)
        self.procesar_ventas_btn.setText("   📊 Generar Informe de Ventas (Centros de Costo)   ")
        self.procesar_btn.setEnabled(True)
        self.procesar_btn.setText("   🎯 Dashboard Vendedores (Pedidos)   ")

        nc_info = ""
        bruto_info = ""
        if stats.get('n_facturas_pos', 0) > 0 and stats.get('n_nc_total', 0) > 0:
            bruto_info = (
                f"   • Total Ventas Brutas: ${stats['total_ventas_brutas']:,.0f} ({stats['n_facturas_pos']} facturas)\n"
                f"   • Total Notas Crédito (-): -${abs(stats['total_nc']):,.0f} ({stats['n_nc_total']} notas cruzadas)\n"
            )
        elif stats.get('n_nc_cruzadas', 0) > 0:
            nc_info = f"   • Notas Crédito cruzadas (-): {stats['n_nc_cruzadas']} (-${abs(stats['total_nc_cruzadas']):,.0f})\n"

        msg = (
            f"🎉 Informe de Ventas generado con éxito!\n\n"
            f"Métricas de la Facturación:\n"
            f"   • Documentos procesados: {stats['total_facturas']}\n"
            f"{bruto_info}"
            f"{nc_info}"
            f"   • Total Facturación Neta: ${stats['total_ventas']:,.0f}\n"
            f"   • Ventas Exentas (Tarifa 0%): ${stats['total_exentas']:,.0f} ({stats['pct_exento']*100:.1f}%)\n"
            f"   • Base Gravada (Tarifa 19%): ${stats['total_base_gravada']:,.0f}\n"
            f"   • Total IVA Facturado: ${stats['total_iva']:,.0f}\n"
            f"   • Total Ventas Gravadas: ${stats['total_gravado']:,.0f} ({stats['pct_gravado']*100:.1f}%)\n"
            f"   • Sedes / Centros de Costo: {stats['total_sedes']}\n"
            f"   • Vendedores activos: {stats['total_vendedores']}\n\n"
            f"Hojas generadas en el archivo:\n"
            f"   1. Resumen Ventas CC y Vendedor - Matriz con separación de Ventas Brutas y Notas por vendedor\n"
            f"   2. Consolidado por Sede - Resumen ejecutivo de facturación bruta, notas y tributario\n"
            f"   3. Detalle Ventas - Base de datos completa con factura referenciada y columnas de venta bruta vs nota\n"
            f"   4. Detalle Notas Crédito - Auditoría detallada de cada nota crédito cruzada con su factura de origen\n\n"
            f"Guardado en:\n{stats['ruta_guardada']}"
        )
        if stats.get("tb_cargado"):
            msg += (
                f"\n\nConciliación con Balance Odoo (Otros Ingresos y Taller):\n"
                f"   • Total Otros Ingresos Auditados: ${stats.get('total_otros_ingresos', 0):,.0f}\n"
                f"   • Liquidación Taller (Auteco + Autotécnica): ${stats.get('liq_taller_total', 0):,.0f}\n"
                f"   • Devoluciones Cuenta 4175 (Base NCs): ${stats.get('dev_4175', 0):,.0f}\n"
                f"   • Hoja 'Otros Ingresos y Balance' incluida con éxito en el Excel."
            )
        self.result_text.setPlainText(msg)

        try:
            os.startfile(stats['ruta_guardada'])  # type: ignore
        except Exception:
            pass

    def on_error_ventas(self, err_msg):
        self.procesar_ventas_btn.setEnabled(True)
        self.procesar_ventas_btn.setText("   📊 Generar Informe de Ventas (Centros de Costo)   ")
        self.procesar_btn.setEnabled(True)
        self.procesar_btn.setText("   🎯 Dashboard Vendedores (Pedidos)   ")
        err_traducido = self._traducir_error(err_msg)
        self.result_text.setPlainText(f"Error al procesar ventas:\n\n{err_traducido}")
        QMessageBox.critical(
            self, "Error", f"No se pudo generar el informe de ventas:\n\n{err_traducido}"
        )

    def procesar(self):
        self.ruta_salida, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar Dashboard de Vendedores",
            "Dashboard_Vendedores.xlsx",
            "Archivos Excel (*.xlsx)"
        )
        if not self.ruta_salida:
            return

        self.result_text.setPlainText(
            "Procesando datos...\n"
            "Esto puede tomar unos segundos dependiendo del tamano del archivo."
        )
        self.procesar_btn.setEnabled(False)
        self.procesar_btn.setText("   Procesando...   ")

        self.thread = VendedorWorker(self.ruta_archivo, self.ruta_salida,
                                      ruta_asiento_nc=getattr(self, 'ruta_asiento_nc', None),
                                      ruta_dian=getattr(self, 'ruta_dian', None))
        self.thread.finished.connect(self.on_success)
        self.thread.error.connect(self.on_error)
        self.thread.start()

    def _traducir_error(self, err_msg):
        msg_str = str(err_msg)
        if "Permission denied" in msg_str or "Errno 13" in msg_str:
            import re
            m = re.search(r"'(.*?)'", msg_str)
            fname = os.path.basename(m.group(1)) if m else "el archivo de Excel"
            return (
                f"⚠️ El archivo '{fname}' está actualmente ABIERTO en Microsoft Excel.\n\n"
                "Windows no permite modificar archivos mientras Excel los tiene abiertos.\n\n"
                "Solución fácil:\n"
                "1. Ve a Excel y CIERRA esa ventana.\n"
                "2. Vuelve a pulsar el botón y se generará de inmediato."
            )
        if "is not unique" in msg_str:
            return (
                f"Error de estructura en el archivo Excel: {msg_str}\n\n"
                "Una de las columnas tiene nombre duplicado en el archivo de entrada."
            )
        if "NoneType" in msg_str and "upper" in msg_str:
            return (
                "⚠️ Diagnóstico del archivo: Se detectó una fila donde el Centro de Costo o Vendedor venía en blanco (sin sede asignada).\n\n"
                "El sistema ya se configuró para asignar la sede automáticamente por el vendedor o marcarlo como 'SIN ASIGNAR' para que nunca se interrumpa."
            )
        if "KeyError" in msg_str:
            return (
                f"⚠️ No se encontró una columna requerida en el archivo:\n{msg_str}\n\n"
                "Verifica que el archivo exportado contenga las columnas de Odoo habituales."
            )
        return msg_str

    def on_success(self, stats):
        self.procesar_btn.setEnabled(True)
        self.procesar_btn.setText("   🎯 Dashboard Vendedores (Pedidos)   ")
        self.procesar_ventas_btn.setEnabled(True)

        if stats.get('es_asiento_ventas'):
            self.on_success_ventas(stats)
            return

        if 'ruta_guardada' in stats and stats['ruta_guardada']:
            self.ruta_salida = stats['ruta_guardada']

        v_brutas = stats.get('ventas_brutas', stats['total_ventas'])
        v_nc = stats.get('total_nc', 0.0)
        v_netas = stats.get('ventas_netas', stats['total_ventas'])

        breakdown_str = f"   • Ventas Brutas: ${v_brutas:,.0f}\n"
        if v_nc != 0.0:
            breakdown_str += f"   • Notas Crédito (-): -${abs(v_nc):,.0f}\n"
        breakdown_str += f"   • Ventas Netas: ${v_netas:,.0f}\n"

        hojas_str = (
            f"Hojas generadas en el archivo:\n"
            f"   1. Resumen Ejecutivo (KPIs con Ventas Brutas/NC/Netas + Ranking)\n"
            f"   2. Clientes por Vendedor (Almacen, Factura, Coste, Utilidad)\n"
            f"   3. Detalle Completo (Todos los pedidos con columnas completas)\n"
            f"   4. Por Almacen (Distribucion por punto de venta)\n"
            f"   5. Top Productos (Mas vendidos con margen)\n"
            f"   6. Evolucion Semanal (Tendencia temporal Top 8)"
        )

        dian_info_str = ""
        dian_stats = stats.get('dian_stats', None)
        if dian_stats:
            hojas_str += "\n   7. Conciliación DIAN vs Sistema (Auditoría cruzada Facturas y NC con motivos)"
            iva_str = f" | IVA: ${dian_stats['facturas_iva']:,.0f}" if 'facturas_iva' in dian_stats else ""
            nc_iva_str = f" | IVA: -${dian_stats['nc_iva']:,.0f}" if 'nc_iva' in dian_stats else ""
            neto_iva_str = f" | IVA Neto: ${dian_stats['neto_iva_dian']:,.0f}" if 'neto_iva_dian' in dian_stats else ""
            dif_iva_str = f" (IVA: ${dian_stats['dif_global_iva']:,.0f})" if 'dif_global_iva' in dian_stats else ""
            dian_info_str = (
                f"\n\nConciliación Oficial DIAN (con IVA):\n"
                f"   • Facturas DIAN: {dian_stats['facturas_count']} (${dian_stats['facturas_total']:,.0f}{iva_str})\n"
                f"   • Notas Crédito DIAN: {dian_stats['nc_count']} (-${dian_stats['nc_total']:,.0f}{nc_iva_str})\n"
                f"   • Total Neto DIAN: ${dian_stats['neto_dian']:,.0f}{neto_iva_str}\n"
                f"   • Facturas solo en DIAN (no en pedidos): {dian_stats['solo_dian_count']} (${dian_stats['solo_dian_total']:,.0f})\n"
                f"   • Diferencia global explicada: ${dian_stats['dif_global']:,.0f}{dif_iva_str}"
            )

        aviso_ruta = ""
        if 'ruta_guardada' in stats and stats['ruta_guardada'] != getattr(self, '_ruta_original_pedida', ''):
            if os.path.basename(stats['ruta_guardada']) != "Dashboard_Vendedores.xlsx":
                aviso_ruta = f"\n(Nota: Se guardó con número incremental porque el anterior estaba abierto en Excel)\n"

        msg = (
            f"🎉 Dashboard de Vendedores generado con éxito!\n\n"
            f"Métricas del Dashboard:\n"
            f"{breakdown_str}"
            f"   • Coste total: ${stats['total_coste']:,.0f}\n"
            f"   • Utilidad: ${stats['total_utilidad']:,.0f}\n"
            f"   • Margen real global: {stats['margen_real'] * 100:.1f}%\n"
            f"   • Vendedores activos: {stats['total_vendedores']}\n"
            f"   • Clientes atendidos: {stats['total_clientes']}\n"
            f"   • Pedidos procesados: {stats['total_pedidos']}\n"
            f"   • Top vendedor: {stats['top_vendedor']}\n"
            f"{dian_info_str}\n"
            f"{hojas_str}\n"
            f"{aviso_ruta}"
            f"Guardado en:\n{self.ruta_salida}"
        )
        self.result_text.setPlainText(msg)

        try:
            os.startfile(self.ruta_salida)  # type: ignore
        except Exception:
            pass

    def on_error(self, err_msg):
        self.procesar_btn.setEnabled(True)
        self.procesar_btn.setText("   🎯 Dashboard Vendedores (Pedidos)   ")
        self.procesar_ventas_btn.setEnabled(True)
        err_traducido = self._traducir_error(err_msg)
        self.result_text.setPlainText(f"Error al procesar:\n\n{err_traducido}")
        QMessageBox.critical(
            self, "Archivo Bloqueado en Excel" if "ABIERTO" in err_traducido else "Error",
            f"No se pudo generar el dashboard:\n\n{err_traducido}"
        )

    def cargar_archivo_tb(self):
        file, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar Balance de Comprobación / Trial Balance Odoo",
            "",
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )
        if file:
            self.ruta_tb = file
            fname = os.path.basename(file)
            self.tb_file_path.setText(fname)
            self.result_text.setPlainText(
                f"⚖️ Balance de Comprobación cargado: {fname}\n\n"
                "• El sistema extraerá automáticamente las cuentas de Otros Ingresos:\n"
                "  - Cesantías comercial movilidad y repuestos (41800103 / 41800104)\n"
                "  - Comisiones (42253001)\n"
                "  - Intereses de financiación (42100501)\n"
                "  - Netforms / Administrativos (42352001)\n"
                "  - Arriendo UMA (42354001 - Supermotos UMA)\n"
                "  - Descuentos comerciales condicionados (42104001)\n"
                "  - Ajuste al peso (42958101) y Sobrantes de caja (42955301)\n"
                "  - Liquidación de Taller (Auteco y Autotécnica bajo 42354001)\n\n"
                "• Se generará la hoja ejecutiva 'Otros Ingresos y Balance' con el Estado Integral de Ingresos "
                "y la conciliación forense de diferencias."
            )

    def cargar_archivo_gastos(self):
        file, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar Gastos por Centro de Costo (con pestañas por sede)",
            "",
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )
        if file:
            self.ruta_gastos = file
            fname = os.path.basename(file)
            self.gastos_file_path.setText(fname)
            self.result_text.setPlainText(
                f"📊 Gastos por Centro de Costo cargados: {fname}\n\n"
                "• El sistema integrará las pestañas de sedes (Consolidado, San José, Granada Mobility, San Martín, Villavicencio, Puerto Inírida, Carreño, Granada TVS).\n"
                "• Conciliará los gastos (Clase 5: Administrativos 51, Ventas 52, Financieros 53).\n"
                "• Generará automáticamente las hojas 'Gastos por Centro Costo' y 'Estado de Resultados' con diseño pastel ejecutivo y fórmulas vivas de Excel."
            )

    def cargar_archivo_dian(self):
        file, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar Reporte DIAN / Token",
            "",
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )
        if file:
            self.ruta_dian = file
            fname = os.path.basename(file)
            self.dian_file_path.setText(fname)
            self.result_text.setPlainText(
                f"Reporte DIAN cargado: {fname}\n\n"
                "• Se filtrarán automáticamente solo Facturas Electrónicas y Notas Crédito (excluyendo nóminas y eventos).\n"
                "• Se generará la hoja 'Conciliación DIAN vs Sistema' con la comparativa de totales, comentarios explicativos "
                "del motivo de las diferencias y las facturas detalladas."
            )

    def cargar_archivo_nc(self):
        file, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar Asiento Contable (para NCC)",
            "",
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )
        if file:
            self.ruta_asiento_nc = file
            fname = os.path.basename(file)
            self.ncc_file_path.setText(fname)
            self.procesar_nc_btn.setEnabled(True)
            self.result_text.setPlainText(
                f"📄 Notas Crédito cargadas: {fname}\n\n"
                "• Al procesar el Informe de Ventas, estas notas crédito se cruzarán automáticamente por Centro de Costo y Vendedor.\n"
                "• El sistema recuperará el vendedor original de cada factura revertida.\n\n"
                "Opciones disponibles:\n"
                "  1. '📊 Generar Informe de Ventas (Centros de Costo)': Matriz consolidada de Ventas Netas con Notas Crédito cruzadas.\n"
                "  2. '📄 Reporte Notas Crédito': Reporte exclusivo de las Notas Crédito."
            )

    def procesar_nc(self):
        if not self.ruta_asiento_nc:
            return
        self.ruta_salida_nc, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar Reporte de Notas Crédito",
            "Reporte_Notas_Credito.xlsx",
            "Archivos Excel (*.xlsx)"
        )
        if not self.ruta_salida_nc:
            return

        self.result_text.setPlainText(
            "Procesando Notas Crédito desde Asiento Contable...\n"
            "• Filtrando únicamente notas crédito (NCC/NC)\n"
            "• Excluyendo borradores preliminares y cancelados\n"
            "• Recuperando vendedores basándose en el historial de facturación de cada cliente\n"
            "• Aplicando resaltado visual en color ámbar para auditoría...\n"
            "Esto puede tardar unos segundos."
        )
        self.procesar_nc_btn.setEnabled(False)
        self.procesar_nc_btn.setText("   Procesando Notas...   ")

        self.thread_nc = NotasCreditoWorker(self.ruta_asiento_nc, self.ruta_salida_nc)
        self.thread_nc.finished.connect(self.on_success_nc)
        self.thread_nc.error.connect(self.on_error_nc)
        self.thread_nc.start()

    def on_success_nc(self, stats):
        self.procesar_nc_btn.setEnabled(True)
        self.procesar_nc_btn.setText("   📄 Generar Reporte de Notas Crédito   ")

        if 'ruta_guardada' in stats and stats['ruta_guardada']:
            self.ruta_salida_nc = stats['ruta_guardada']

        msg = (
            f"🎉 Reporte de Notas Crédito generado con éxito!\n\n"
            f"Métricas del Reporte:\n"
            f"   • Notas Crédito procesadas (efectivas): {stats['total_nc']}\n"
            f"   • Vendedores asignados por Cliente: {stats['nc_recuperadas']} (resaltadas en ámbar en el Excel)\n"
            f"   • Notas excluidas (Borrador / Cancelado): {stats['nc_excluidas']}\n"
            f"   • Total Devoluciones NC: ${abs(stats['total_total']):,.0f}\n"
            f"   • Base Imponible NC: ${abs(stats['total_base']):,.0f}\n"
            f"   • Vendedores con NC: {stats['total_vendedores']}\n"
            f"   • Ticket promedio NC: ${abs(stats['ticket_prom']):,.0f}\n\n"
            f"Hojas generadas en el archivo:\n"
            f"   1. Resumen NCC  -  KPIs, ranking de vendedores y gráfica de distribución\n"
            f"   2. Detalle NCC  -  Listado completo de notas con columna 'Origen Vendedor' y alertas en ámbar\n"
            f"   3. Por Vendedor  -  Agrupado por vendedor con desglose de clientes\n\n"
            f"Guardado en:\n{self.ruta_salida_nc}"
        )
        self.result_text.setPlainText(msg)

        try:
            os.startfile(self.ruta_salida_nc)  # type: ignore
        except Exception:
            pass

    def on_error_nc(self, err_msg):
        self.procesar_nc_btn.setEnabled(True)
        self.procesar_nc_btn.setText("   📄 Generar Reporte de Notas Crédito   ")
        err_traducido = self._traducir_error(err_msg)
        self.result_text.setPlainText(f"Error al procesar notas crédito:\n\n{err_traducido}")
        QMessageBox.critical(
            self, "Archivo Bloqueado en Excel" if "ABIERTO" in err_traducido else "Error",
            f"No se pudo generar el reporte de notas crédito:\n\n{err_traducido}"
        )

    def abrir_config_vendedores(self):
        dlg = ConfigVendedoresDialog(self)
        dlg.exec()

    def limpiar_todo(self):
        """Limpia todas las rutas, campos de texto y resultados para empezar desde cero."""
        self.ruta_archivo = ""
        self.ruta_salida = ""
        self.ruta_salida_ventas = ""
        self.ruta_asiento_nc = ""
        self.ruta_salida_nc = ""
        self.ruta_dian = ""
        self.ruta_tb = ""
        self.ruta_gastos = ""
        self.file_path.clear()
        self.tb_file_path.clear()
        self.gastos_file_path.clear()
        self.ncc_file_path.clear()
        self.dian_file_path.clear()
        self.info_label.setText("")
        self.result_text.clear()
        self.procesar_ventas_btn.setEnabled(False)
        self.procesar_ventas_btn.setText("   📊 Generar Informe de Ventas (Centros de Costo)   ")
        self.procesar_btn.setEnabled(False)
        self.procesar_btn.setText("   🎯 Dashboard Vendedores (Pedidos)   ")
        self.procesar_nc_btn.setEnabled(False)
        self.procesar_nc_btn.setText("   📄 Reporte Notas Crédito   ")
