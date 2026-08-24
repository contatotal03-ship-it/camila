import os
from PyQt6.QtWidgets import (  # type: ignore
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QLineEdit, QTextEdit, QFileDialog, QMessageBox, QFrame,
    QGraphicsDropShadowEffect
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize  # type: ignore
from PyQt6.QtGui import QFont, QColor  # type: ignore

try:
    from core.vendedores_logic import procesar_dashboard_vendedores, procesar_dashboard_nc_asiento  # type: ignore
except ImportError:
    import sys
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from core.vendedores_logic import procesar_dashboard_vendedores, procesar_dashboard_nc_asiento  # type: ignore


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


class VendedorWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, ruta_archivo, ruta_salida, ruta_asiento_nc=None):
        super().__init__()
        self.ruta_archivo = ruta_archivo
        self.ruta_salida = ruta_salida
        self.ruta_asiento_nc = ruta_asiento_nc

    def run(self):
        try:
            stats = procesar_dashboard_vendedores(self.ruta_archivo, self.ruta_salida,
                                                  ruta_asiento_nc=self.ruta_asiento_nc)
            self.finished.emit(stats)
        except Exception as e:
            self.error.emit(str(e))


class VendedoresWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.ruta_archivo = ""
        self.ruta_salida = ""
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
            "Selecciona el archivo 'Pedido de venta (sale.order).xlsx' desde cualquier ubicacion"
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
            "Ningun archivo seleccionado..."
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
        ncc_lbl = QLabel("NCC:")
        ncc_lbl.setStyleSheet("color: #4fc3f7; font-size: 11px; font-weight: bold;")
        ncc_lbl.setFixedWidth(40)
        self.ncc_file_path = QLineEdit()
        self.ncc_file_path.setReadOnly(True)
        self.ncc_file_path.setPlaceholderText("Asiento contable (opcional, para NCC)...")
        self.ncc_file_path.setMinimumHeight(32)
        self.ncc_file_btn = QPushButton("  NCC...  ")
        self.ncc_file_btn.setProperty("class", "SecondaryButton")
        self.ncc_file_btn.setMinimumHeight(32)
        self.ncc_file_btn.setMinimumWidth(80)
        self.ncc_file_btn.clicked.connect(self.cargar_archivo_nc)
        ncc_row.addWidget(ncc_lbl)
        ncc_row.addWidget(self.ncc_file_path)
        ncc_row.addWidget(self.ncc_file_btn)
        input_card_layout.addLayout(ncc_row)

        root.addWidget(input_card)
        root.addSpacing(16)

        # ── Generate Button ────────────────────────────────────
        btn_container = QFrame()
        btn_container.setStyleSheet("background: transparent; border: none;")
        btn_layout = QHBoxLayout(btn_container)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.addStretch()

        self.procesar_btn = QPushButton("   Generar Dashboard de Vendedores   ")
        self.procesar_btn.setProperty("class", "PrimaryButton")
        self.procesar_btn.setMinimumHeight(48)
        self.procesar_btn.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        self.procesar_btn.clicked.connect(self.procesar)
        self.procesar_btn.setEnabled(False)

        btn_layout.addWidget(self.procesar_btn)
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
            "Seleccionar archivo de Pedidos de Venta",
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

            self.procesar_btn.setEnabled(True)
            self.result_text.setPlainText(
                f"Archivo cargado correctamente: {fname}\n\n"
                "Haz clic en 'Generar Dashboard de Vendedores' para procesar..."
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
                                      ruta_asiento_nc=getattr(self, 'ruta_asiento_nc', None))
        self.thread.finished.connect(self.on_success)
        self.thread.error.connect(self.on_error)
        self.thread.start()

    def on_success(self, stats):
        self.procesar_btn.setEnabled(True)
        self.procesar_btn.setText("   Generar Dashboard de Vendedores   ")

        msg = (
            f"Dashboard generado con exito!\n\n"
            f"Metricas:\n"
            f"   Vendedores analizados: {stats['total_vendedores']}\n"
            f"   Clientes totales: {stats['total_clientes']}\n"
            f"   Pedidos procesados: {stats['total_pedidos']}\n"
            f"   Ventas totales: ${stats['total_ventas']:,.0f}\n"
            f"   Coste total: ${stats['total_coste']:,.0f}\n"
            f"   Utilidad: ${stats['total_utilidad']:,.0f}\n"
            f"   Margen real: {stats['margen_real']*100:.1f}%\n"
            f"   Top vendedor: {stats['top_vendedor']}\n\n"
            f"El archivo incluye 6 hojas:\n"
            f"   1. Resumen Ejecutivo (KPIs + Ranking con Coste/Utilidad)\n"
            f"   2. Clientes por Vendedor (Almacen, Factura, Coste, Utilidad)\n"
            f"   3. Detalle Completo (todos los pedidos con Coste)\n"
            f"   4. Por Almacen (distribucion)\n"
            f"   5. Top Productos (Ventas, Coste, Utilidad, Margen)\n"
            f"   6. Evolucion Semanal\n\n"
            f"Guardado en:\n{self.ruta_salida}"
        )
        self.result_text.setPlainText(msg)

        try:
            os.startfile(self.ruta_salida)  # type: ignore
        except Exception:
            pass

    def on_error(self, err_msg):
        self.procesar_btn.setEnabled(True)
        self.procesar_btn.setText("   Generar Dashboard de Vendedores   ")
        self.result_text.setPlainText(f"Error al procesar:\n\n{err_msg}")
        QMessageBox.critical(
            self, "Error",
            f"No se pudo generar el dashboard:\n\n{err_msg}"
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
            self.ncc_file_path.setText(os.path.basename(file))
