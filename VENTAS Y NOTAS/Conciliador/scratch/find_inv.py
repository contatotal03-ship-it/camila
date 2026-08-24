import sys
import os
import pandas as pd

sys.path.append(os.getcwd())
from core.excel_utils import robust_read_excel, robust_ExcelFile

def find_invoice():
    ruta = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD 2026.xlsx'
    xls = robust_ExcelFile(ruta)
    for s in xls.sheet_names:
        try:
            df = robust_read_excel(xls, sheet_name=s)
            if 'Documento' in df.columns:
                docs = df['Documento'].astype(str).unique()
                for d in docs:
                    if '7649' in d:
                        print(f"Factura 7649 encontrada en Hoja: {s}, Documento exacto: '{d}'")
        except: pass

if __name__ == '__main__':
    find_invoice()
