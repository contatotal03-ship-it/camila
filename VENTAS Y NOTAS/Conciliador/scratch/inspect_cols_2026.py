import sys
import os
import pandas as pd

sys.path.append(os.getcwd())

from core.excel_utils import robust_read_excel, robust_ExcelFile
from core.config_manager import get_hojas_ventas

def inspect():
    ruta_contabilidad = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD 2026.xlsx'
    
    print("--- INSPECCIONANDO COLUMNAS EN CONTABILIDAD 2026 ---")
    try:
        xls = robust_ExcelFile(ruta_contabilidad)
        hojas_config = [h.strip().upper() for h in get_hojas_ventas()]
        
        for sheet in xls.sheet_names:
            if sheet.strip().upper() in hojas_config:
                print(f"\n--- Hoja: {sheet} ---")
                df = robust_read_excel(xls, sheet_name=sheet, nrows=5)
                print("Columnas encontradas:", df.columns.tolist())
                # Verificar columnas clave
                cols = df.columns.tolist()
                has_doc = 'Documento' in cols
                has_haber = any('Haber' in c for c in cols)
                has_debe = any('Debe' in c for c in cols)
                has_cuenta = any('cuenta' in c.lower() for c in cols)
                
                print(f"¿Tiene 'Documento'?: {has_doc}")
                print(f"¿Tiene 'Haber'?: {has_haber}")
                print(f"¿Tiene 'Debe'?: {has_debe}")
                print(f"¿Tiene 'cuenta'?: {has_cuenta}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == '__main__':
    inspect()
