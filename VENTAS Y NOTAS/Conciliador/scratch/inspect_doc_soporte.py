import sys
sys.path.append(r'c:\Users\Admin\Downloads\camila\VENTAS Y NOTAS\Conciliador')
from core.excel_utils import robust_read_excel
import pandas as pd

ruta_sistema = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\marzo 2026 odoo.xlsx'
df = robust_read_excel(ruta_sistema, sheet_name='DOC. SOPORTE')
df.columns = [str(c).strip() for c in df.columns]

col_haber = next((c for c in df.columns if 'Haber' in c), None)
col_debe = next((c for c in df.columns if 'Debe' in c), None)
col_cuenta = next((c for c in df.columns if 'cuenta' in c.lower()), None)
col_nombre = next((c for c in df.columns if 'nombre' in c.lower()), None)

def clean_money(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).replace('$', '').replace('.', '').replace(',', '.').replace('\xa0', '').strip()
    try: return float(val_str)
    except: return 0.0
    
df['Monto_Debe'] = df[col_debe].apply(clean_money)
df['Monto_Haber'] = df[col_haber].apply(clean_money)
df['Cod_Cuenta'] = df[col_cuenta].astype(str).str.replace('.0', '', regex=False).str.strip()

# Filtrar totales
nom_upper = df[col_nombre].astype(str).str.strip().str.upper()
es_total = nom_upper.str.startswith('TOTAL') | nom_upper.str.startswith('SUMA')
doc_vacio = df['Documento'].isna() | df['Documento'].astype(str).str.strip().str.upper().isin(['', 'NAN', 'NONE'])
df_filtered = df[~(es_total & doc_vacio)].copy().reset_index(drop=True)

print("Total Debe en df_filtered:", df_filtered['Monto_Debe'].sum())
print("Total Haber en df_filtered:", df_filtered['Monto_Haber'].sum())

# Vamos a ver las cuentas y sus montos sumados
df_cta = df_filtered.groupby(['Cod_Cuenta']).agg({
    'Monto_Debe': 'sum',
    'Monto_Haber': 'sum'
}).reset_index()
print("\n--- SUMA POR CUENTA CONTABLE ---")
print(df_cta)
