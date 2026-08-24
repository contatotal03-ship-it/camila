import sys
sys.path.append(r'c:\Users\Admin\Downloads\camila\VENTAS Y NOTAS\Conciliador')
from core.excel_utils import robust_read_excel
from core.config_manager import get_hojas_compras, get_cuentas_total_compras
import pandas as pd
import numpy as np

ruta_sistema = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\marzo 2026 odoo.xlsx'
hojas_validas = get_hojas_compras()
xls_sistema = pd.ExcelFile(ruta_sistema)

for sheet in xls_sistema.sheet_names:
    if sheet.strip().upper() not in hojas_validas:
        continue
        
    df_sheet = robust_read_excel(xls_sistema, sheet_name=sheet)
    df_sheet.columns = [str(c).strip() for c in df_sheet.columns]
    
    col_haber = next((c for c in df_sheet.columns if 'Haber' in c), None)
    col_debe = next((c for c in df_sheet.columns if 'Debe' in c), None)
    col_cuenta = next((c for c in df_sheet.columns if 'cuenta' in c.lower()), None)
    col_nombre = next((c for c in df_sheet.columns if 'nombre' in c.lower()), None)
    
    def clean_money(val):
        if pd.isna(val): return 0.0
        if isinstance(val, (int, float)): return float(val)
        val_str = str(val).replace('$', '').replace('.', '').replace(',', '.').replace('\xa0', '').strip()
        try: return float(val_str)
        except: return 0.0
        
    df_sheet['Monto_Debe'] = df_sheet[col_debe].apply(clean_money)
    df_sheet['Monto_Haber'] = df_sheet[col_haber].apply(clean_money)
    
    nom_upper = df_sheet[col_nombre].astype(str).str.strip().str.upper()
    es_total = nom_upper.str.startswith('TOTAL') | nom_upper.str.startswith('SUMA')
    doc_vacio = df_sheet['Documento'].isna() | df_sheet['Documento'].astype(str).str.strip().str.upper().isin(['', 'NAN', 'NONE'])
    
    df_filtered = df_sheet[~(es_total & doc_vacio)].copy().reset_index(drop=True)
    
    print(f'Sheet {sheet}:')
    print(f'  Original Rows: {len(df_sheet)}, Debe={df_sheet["Monto_Debe"].sum():,.2f}, Haber={df_sheet["Monto_Haber"].sum():,.2f}')
    print(f'  Filtered Rows: {len(df_filtered)}, Debe={df_filtered["Monto_Debe"].sum():,.2f}, Haber={df_filtered["Monto_Haber"].sum():,.2f}')
