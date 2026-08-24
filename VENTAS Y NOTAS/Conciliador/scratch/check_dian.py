import sys
import os
import pandas as pd

sys.path.append(os.getcwd())
from core.excel_utils import robust_read_excel

def check():
    df = robust_read_excel(r'C:\Users\Admin\Downloads\camila\Enviados 2026.xlsx')
    col_fol = next(c for c in df.columns if 'folio' in c.lower())
    found = df[df[col_fol].astype(str).str.contains('7729', na=False)]
    if not found.empty:
        print("Encontrado en DIAN:")
        print(found[[col_fol, 'Prefijo', 'Total']])
    else:
        print("Folio 7729 NO existe en el archivo de la DIAN.")

if __name__ == '__main__':
    check()
