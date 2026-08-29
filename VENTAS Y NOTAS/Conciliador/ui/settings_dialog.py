from PyQt6.QtWidgets import ( # type: ignore
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QPlainTextEdit, QPushButton, QMessageBox, QTabWidget, QWidget
)
from PyQt6.QtCore import Qt # type: ignore
from core.config_manager import load_config, CONFIG_FILE # type: ignore
import json
import os

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent) # type: ignore
        self.setWindowTitle("Configuración de Hojas Sistema")
        self.setMinimumSize(500, 400)
        self.setWindowModality(Qt.WindowModality.WindowModal)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.init_ui()
        self.load_settings()

    def showEvent(self, event):
        super().showEvent(event)
        self.raise_()
        self.activateWindow()

    def init_ui(self):
        layout = QVBoxLayout(self)
        
        self.tabs = QTabWidget()
        
        # Tab Ventas
        self.tab_ventas = QWidget()
        v_layout = QVBoxLayout(self.tab_ventas)
        v_layout.addWidget(QLabel("Escribe los nombres de las hojas de VENTAS (una por línea):"))
        self.txt_ventas = QPlainTextEdit()
        v_layout.addWidget(self.txt_ventas)
        
        v_layout.addWidget(QLabel("Nombre del GRUPO DIAN para Ventas (ej: Emitido):"))
        self.txt_grupo_ventas = QPlainTextEdit()
        self.txt_grupo_ventas.setMaximumHeight(30)
        v_layout.addWidget(self.txt_grupo_ventas)
        
        v_layout.addWidget(QLabel("Tipos de Documento DIAN válidos para Ventas (ej: Factura electrónica de Venta):"))
        self.txt_tipos_ventas = QPlainTextEdit()
        self.txt_tipos_ventas.setMaximumHeight(60)
        v_layout.addWidget(self.txt_tipos_ventas)
        
        v_layout.addWidget(QLabel("Palabras clave para EXCLUIR del Token DIAN (ej: soporte, peaje):"))
        self.txt_dian_ventas = QPlainTextEdit()
        self.txt_dian_ventas.setMaximumHeight(80)
        v_layout.addWidget(self.txt_dian_ventas)
        
        v_layout.addWidget(QLabel("Cuentas de IVA en Ventas (ej: 2408):"))
        self.txt_iva_ventas = QPlainTextEdit()
        self.txt_iva_ventas.setMaximumHeight(60)
        v_layout.addWidget(self.txt_iva_ventas)
        
        v_layout.addWidget(QLabel("Cuentas a EXCLUIR en Ventas (ej: 6, 14):"))
        self.txt_excluir_ventas = QPlainTextEdit()
        self.txt_excluir_ventas.setMaximumHeight(60)
        v_layout.addWidget(self.txt_excluir_ventas)
        
        self.tabs.addTab(self.tab_ventas, "📊 Ventas")
        
        # Tab Compras
        self.tab_compras = QWidget()
        c_layout = QVBoxLayout(self.tab_compras)
        c_layout.addWidget(QLabel("Escribe los nombres de las hojas de COMPRAS/GASTOS (una por línea):"))
        self.txt_compras = QPlainTextEdit()
        c_layout.addWidget(self.txt_compras)
        
        c_layout.addWidget(QLabel("Nombre del GRUPO DIAN para Compras (ej: Recibido):"))
        self.txt_grupo_compras = QPlainTextEdit()
        self.txt_grupo_compras.setMaximumHeight(30)
        c_layout.addWidget(self.txt_grupo_compras)
        
        c_layout.addWidget(QLabel("Tipos de Documento DIAN válidos para Compras (ej: Factura electrónica de Venta):"))
        self.txt_tipos_compras = QPlainTextEdit()
        self.txt_tipos_compras.setMaximumHeight(60)
        c_layout.addWidget(self.txt_tipos_compras)
        
        c_layout.addWidget(QLabel("Palabras clave para INCLUIR del Token DIAN (ej: soporte, peaje):"))
        self.txt_dian_compras = QPlainTextEdit()
        self.txt_dian_compras.setMaximumHeight(80)
        c_layout.addWidget(self.txt_dian_compras)
        
        c_layout.addWidget(QLabel("Cuentas de IVA en Compras (ej: 2408):"))
        self.txt_iva_compras = QPlainTextEdit()
        self.txt_iva_compras.setMaximumHeight(60)
        c_layout.addWidget(self.txt_iva_compras)
        
        c_layout.addWidget(QLabel("Cuentas de TOTAL (Proveedores/Retención) en Compras (ej: 2205, 2365):"))
        self.txt_total_compras = QPlainTextEdit()
        self.txt_total_compras.setMaximumHeight(60)
        c_layout.addWidget(self.txt_total_compras)
        
        self.tabs.addTab(self.tab_compras, "🛒 Compras")
        
        layout.addWidget(self.tabs)
        
        # Botones
        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("Guardar Cambios")
        self.btn_save.clicked.connect(self.save_settings)
        self.btn_cancel = QPushButton("Cancelar")
        self.btn_cancel.clicked.connect(self.reject)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_cancel)
        btn_layout.addWidget(self.btn_save)
        layout.addLayout(btn_layout)

    def load_settings(self):
        config = load_config()
        ventas = config.get("ventas", {}).get("hojas_validas", [])
        compras = config.get("compras", {}).get("hojas_validas", [])
        dian_ventas = config.get("ventas", {}).get("dian_excluir", [])
        dian_compras = config.get("compras", {}).get("dian_incluir", [])
        grupo_ventas = config.get("ventas", {}).get("dian_grupo", "Emitido")
        grupo_compras = config.get("compras", {}).get("dian_grupo", "Recibido")
        
        self.txt_ventas.setPlainText("\n".join(ventas))
        self.txt_compras.setPlainText("\n".join(compras))
        self.txt_dian_ventas.setPlainText("\n".join(dian_ventas))
        self.txt_dian_compras.setPlainText("\n".join(dian_compras))
        self.txt_grupo_ventas.setPlainText(grupo_ventas)
        self.txt_grupo_compras.setPlainText(grupo_compras)
        
        tipos_ventas = config.get("ventas", {}).get("dian_tipos_validos", [])
        tipos_compras = config.get("compras", {}).get("dian_tipos_validos", [])
        self.txt_tipos_ventas.setPlainText("\n".join(tipos_ventas))
        self.txt_tipos_compras.setPlainText("\n".join(tipos_compras))
        
        iva_ventas = config.get("ventas", {}).get("cuentas_iva", [])
        excluir_ventas = config.get("ventas", {}).get("cuentas_excluir", [])
        iva_compras = config.get("compras", {}).get("cuentas_iva", [])
        total_compras = config.get("compras", {}).get("cuentas_total", [])
        
        self.txt_iva_ventas.setPlainText("\n".join(iva_ventas))
        self.txt_excluir_ventas.setPlainText("\n".join(excluir_ventas))
        self.txt_iva_compras.setPlainText("\n".join(iva_compras))
        self.txt_total_compras.setPlainText("\n".join(total_compras))

    def save_settings(self):
        ventas_list = [line.strip() for line in self.txt_ventas.toPlainText().split("\n") if line.strip()]
        compras_list = [line.strip() for line in self.txt_compras.toPlainText().split("\n") if line.strip()]
        dian_ventas_list = [line.strip().lower() for line in self.txt_dian_ventas.toPlainText().split("\n") if line.strip()]
        dian_compras_list = [line.strip().lower() for line in self.txt_dian_compras.toPlainText().split("\n") if line.strip()]
        
        grupo_ventas = self.txt_grupo_ventas.toPlainText().strip()
        grupo_compras = self.txt_grupo_compras.toPlainText().strip()
        
        tipos_ventas_list = [line.strip() for line in self.txt_tipos_ventas.toPlainText().split("\n") if line.strip()]
        tipos_compras_list = [line.strip() for line in self.txt_tipos_compras.toPlainText().split("\n") if line.strip()]
        
        iva_ventas_list = [line.strip() for line in self.txt_iva_ventas.toPlainText().split("\n") if line.strip()]
        excluir_ventas_list = [line.strip() for line in self.txt_excluir_ventas.toPlainText().split("\n") if line.strip()]
        iva_compras_list = [line.strip() for line in self.txt_iva_compras.toPlainText().split("\n") if line.strip()]
        total_compras_list = [line.strip() for line in self.txt_total_compras.toPlainText().split("\n") if line.strip()]
        
        new_config = {
            "compras": {
                "hojas_validas": compras_list,
                "dian_incluir": dian_compras_list,
                "dian_grupo": grupo_compras,
                "dian_tipos_validos": tipos_compras_list,
                "cuentas_iva": iva_compras_list,
                "cuentas_total": total_compras_list
            },
            "ventas": {
                "hojas_validas": ventas_list,
                "dian_excluir": dian_ventas_list,
                "dian_grupo": grupo_ventas,
                "dian_tipos_validos": tipos_ventas_list,
                "cuentas_iva": iva_ventas_list,
                "cuentas_excluir": excluir_ventas_list
            }
        }
        
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(new_config, f, indent=4, ensure_ascii=False)
            QMessageBox.information(self, "Configuración", "¡Configuración guardada correctamente!")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo guardar la configuración: {e}")
