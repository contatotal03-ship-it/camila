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


class VendedoresWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.ruta_archivo = ""
        self.ruta_salida = ""
        self.ruta_asiento_nc = ""
        self.ruta_salida_nc = ""
        self.ruta_dian = ""
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

        root.addWidget(input_card)
        root.addSpacing(16)

        # ── Buttons (Generate & Reset) ─────────────────────────
        btn_container = QFrame()
        btn_container.setStyleSheet("background: transparent; border: none;")
        btn_layout = QHBoxLayout(btn_container)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(12)
        btn_layout.addStretch()

        self.procesar_btn = QPushButton("   Generar Dashboard de Vendedores   ")
        self.procesar_btn.setProperty("class", "PrimaryButton")
        self.procesar_btn.setMinimumHeight(48)
        self.procesar_btn.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        self.procesar_btn.clicked.connect(self.procesar)
        self.procesar_btn.setEnabled(False)

        self.procesar_nc_btn = QPushButton("   📄 Generar Reporte de Notas Crédito   ")
        self.procesar_nc_btn.setProperty("class", "SecondaryButton")
        self.procesar_nc_btn.setMinimumHeight(48)
        self.procesar_nc_btn.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
        self.procesar_nc_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.procesar_nc_btn.setToolTip("Genera un dashboard exclusivo de Notas Crédito desde el Asiento Contable, recuperando vendedores por cliente")
        self.procesar_nc_btn.clicked.connect(self.procesar_nc)
        self.procesar_nc_btn.setEnabled(False)

        self.limpiar_btn = QPushButton("   🧹 Limpiar Todo   ")
        self.limpiar_btn.setProperty("class", "DangerButton")
        self.limpiar_btn.setMinimumHeight(48)
        self.limpiar_btn.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
        self.limpiar_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.limpiar_btn.setToolTip("Limpiar todas las rutas, campos y resultados para empezar desde cero")
        self.limpiar_btn.clicked.connect(self.limpiar_todo)

        btn_layout.addWidget(self.procesar_btn)
        btn_layout.addWidget(self.procesar_nc_btn)
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
                                      ruta_asiento_nc=getattr(self, 'ruta_asiento_nc', None),
                                      ruta_dian=getattr(self, 'ruta_dian', None))
        self.thread.finished.connect(self.on_success)
        self.thread.error.connect(self.on_error)
        self.thread.start()

    def on_success(self, stats):
        self.procesar_btn.setEnabled(True)
        self.procesar_btn.setText("   Generar Dashboard de Vendedores   ")

        v_brutas = stats.get('ventas_brutas', stats['total_ventas'])
        t_nc = stats.get('total_nc', 0)
        v_netas = stats['total_ventas']

        breakdown_str = f"   • Ventas Brutas: ${v_brutas:,.0f}\n"
        if t_nc != 0:
            breakdown_str += f"   • Notas Crédito (restadas): -${abs(t_nc):,.0f}\n"
        breakdown_str += f"   • Ventas Netas (Facturación): ${v_netas:,.0f}\n"

        dian_info_str = ""
        hojas_str = (
            f"El archivo incluye las hojas estándar:\n"
            f"   1. Resumen Ejecutivo (KPIs con Ventas Brutas/NC/Netas + Ranking)\n"
            f"   2. Clientes por Vendedor (Almacen, Factura, Coste, Utilidad)\n"
            f"   3. Detalle Completo (todos los pedidos con Coste)\n"
            f"   4. Por Almacen (distribucion)\n"
            f"   5. Top Productos (Ventas, Coste, Utilidad, Margen)\n"
            f"   6. Evolucion Semanal\n"
        )
        if 'dian_stats' in stats and stats['dian_stats']:
            ds = stats['dian_stats']
            dian_info_str = (
                f"\n🔍 Conciliación con DIAN:\n"
                f"   • Facturas DIAN: {ds.get('facturas_count', 0)} (${ds.get('facturas_total', 0):,.0f})\n"
                f"   • Notas Crédito DIAN: {ds.get('nc_count', 0)} (${ds.get('nc_total', 0):,.0f})\n"
                f"   • Neto Oficial DIAN: ${ds.get('neto_dian', 0):,.0f}\n"
                f"   • Facturas solo en DIAN (no en pedidos): {ds.get('solo_dian_count', 0)} (${ds.get('solo_dian_total', 0):,.0f})\n"
                f"   • Facturas con diferencias de valor: {ds.get('diff_val_count', 0)}\n"
            )
            hojas_str += f"   7. Conciliación DIAN vs Sistema (Totales, comentarios explicativos y auditoría)\n"

        msg = (
            f"🎉 Dashboard generado con éxito!\n\n"
            f"Métricas Financieras:\n"
            f"{breakdown_str}"
            f"   • Coste total: ${stats['total_coste']:,.0f}\n"
            f"   • Utilidad: ${stats['total_utilidad']:,.0f}\n"
            f"   • Margen real: {stats['margen_real']*100:.1f}%\n"
            f"   • Vendedores analizados: {stats['total_vendedores']}\n"
            f"   • Clientes totales: {stats['total_clientes']}\n"
            f"   • Pedidos procesados: {stats['total_pedidos']}\n"
            f"   • Top vendedor: {stats['top_vendedor']}\n"
            f"{dian_info_str}\n"
            f"{hojas_str}\n"
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
                f"Asiento Contable cargado: {fname}\n\n"
                "Opciones disponibles:\n"
                "  1. Clic en '📄 Generar Reporte de Notas Crédito' para procesar únicamente las notas de este archivo.\n"
                "     - Asignará automáticamente vendedores por historial de cliente (filas resaltadas en ámbar).\n"
                "     - Excluirá borradores y cancelados.\n\n"
                "  2. O carga arriba 'Pedido de venta' si deseas generar el Dashboard Integral de Vendedores con notas combinadas."
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
        self.result_text.setPlainText(f"Error al procesar notas crédito:\n\n{err_msg}")
        QMessageBox.critical(
            self, "Error",
            f"No se pudo generar el reporte de notas crédito:\n\n{err_msg}"
        )

    def limpiar_todo(self):
        """Limpia todas las rutas, campos de texto y resultados para empezar desde cero."""
        self.ruta_archivo = ""
        self.ruta_salida = ""
        self.ruta_asiento_nc = ""
        self.ruta_salida_nc = ""
        self.ruta_dian = ""
        self.file_path.clear()
        self.ncc_file_path.clear()
        self.dian_file_path.clear()
        self.info_label.setText("")
        self.result_text.clear()
        self.procesar_btn.setEnabled(False)
        self.procesar_btn.setText("   Generar Dashboard de Vendedores   ")
        self.procesar_nc_btn.setEnabled(False)
        self.procesar_nc_btn.setText("   📄 Generar Reporte de Notas Crédito   ")
