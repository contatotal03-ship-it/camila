import sys
import os

# Add the project directory to sys.path
sys.path.append(r"c:\Users\Admin\Downloads\camila\VENTAS Y NOTAS\Conciliador")

from core.compras_logic import procesar_conciliacion_compras

ruta_sistema = r"C:\Users\Admin\Downloads\MARZO ODOO.xlsx"
ruta_dian = r"C:\Users\Admin\Downloads\TOKEN DIAN..xlsx"
ruta_salida = r"C:\Users\Admin\Downloads\Reporte_Conciliador_Compras_REVISADO.xlsx"

try:
    print("Iniciando conciliación de compras...")
    resultado = procesar_conciliacion_compras(ruta_sistema, ruta_dian, ruta_salida)
    print(f"Conciliación completada: {resultado}")
    print(f"Reporte generado en: {ruta_salida}")
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
