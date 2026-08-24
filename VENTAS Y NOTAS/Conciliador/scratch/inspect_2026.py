import sys
import os
import pandas as pd

# Añadir el directorio actual al path para importar core
sys.path.append(os.getcwd())

from core.excel_utils import robust_read_excel, robust_ExcelFile
from core.config_manager import get_hojas_ventas

def inspect():
    ruta_contabilidad = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD 2026.xlsx'
    ruta_enviados = r'C:\Users\Admin\Downloads\camila\Enviados 2026.xlsx'
    
    print("--- INSPECCIONANDO CONTABILIDAD 2026 ---")
    try:
        xls = robust_ExcelFile(ruta_contabilidad)
        print(f"Hojas encontradas: {xls.sheet_names}")
        hojas_config = get_hojas_ventas()
        print(f"Hojas configuradas para Ventas: {hojas_config}")
        
        for sheet in xls.sheet_names:
            if sheet.strip().upper() in hojas_config:
                print(f"\nAnalizando Hoja: {sheet}")
                df = robust_read_excel(xls, sheet_name=sheet, nrows=10)
                print("Columnas:", df.columns.tolist())
                print(df.head(3))
            else:
                print(f"Hoja '{sheet}' ignorada (no está en la configuración de Ventas)")
    except Exception as e:
        print(f"Error leyendo Contabilidad: {e}")

    print("\n--- INSPECCIONANDO ENVIADOS 2026 (DIAN) ---")
    try:
        df_dian = robust_read_excel(ruta_enviados, nrows=10)
        print("Columnas DIAN:", df_dian.columns.tolist())
        print(df_dian.head(3))
    except Exception as e:
        print(f"Error leyendo Enviados: {e}")

if __name__ == '__main__':
    inspect()
