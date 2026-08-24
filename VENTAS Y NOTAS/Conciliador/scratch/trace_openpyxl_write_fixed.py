import sys
sys.path.append(r'c:\Users\Admin\Downloads\camila\VENTAS Y NOTAS\Conciliador')
from core.excel_utils import robust_read_excel
import openpyxl

ruta_sistema = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\marzo 2026 odoo  - ORIGINAL.xlsx'
ruta_ref = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\PROVEEDORES REFERENCIA MARZO.xlsx'

df_ref = robust_read_excel(ruta_ref)
wb = openpyxl.load_workbook(ruta_sistema)

# Crear un mapa rápido de Número -> (Referencia, Centro de costo)
ref_map = {}
for _, r in df_ref.iterrows():
    num = str(r.get('Número', '')).strip()
    if num:
        ref_map[num] = (str(r.get('Referencia', '')).strip(), str(r.get('Centro de costo', '')).strip())

ws = wb['Compras']
max_col = ws.max_column

# 1. Encontrar la fila de cabecera y el índice de las columnas de interés
header_row_idx = None
col_doc_idx = None
col_ref_idx = None
col_cc_idx = None

for row in range(1, 15):
    row_vals = [str(ws.cell(row, col).value).strip().lower() for col in range(1, max_col + 1)]
    if 'documento' in row_vals:
        header_row_idx = row
        col_doc_idx = row_vals.index('documento') + 1
        
        if 'referencia' in row_vals:
            col_ref_idx = row_vals.index('referencia') + 1
        if 'centro de costo' in row_vals or 'centro_costo' in row_vals:
            col_cc_idx = next((i for i, val in enumerate(row_vals) if 'centro' in val or 'cc' in val), None)
            if col_cc_idx is not None:
                col_cc_idx += 1
        break

# Si no existen, crearlas
if not col_ref_idx or not col_cc_idx:
    last_header_col = ws.max_column
    if not col_ref_idx:
        col_ref_idx = last_header_col + 1
        ws.cell(header_row_idx, col_ref_idx).value = 'Referencia'
        col_cc_idx = last_header_col + 2
        ws.cell(header_row_idx, col_cc_idx).value = 'Centro de costo'
    elif not col_cc_idx:
        col_cc_idx = max(col_ref_idx + 1, last_header_col + 1)
        ws.cell(header_row_idx, col_cc_idx).value = 'Centro de costo'

# Llenar/actualizar con la NUEVA lógica de detección de cuenta
from core.config_manager import get_cuentas_total_compras
cuentas_cabecera = get_cuentas_total_compras()

col_cta_idx = None
col_nombre_idx = None
row_vals_hdr = [str(ws.cell(header_row_idx, col).value).strip().lower() for col in range(1, ws.max_column + 1)]

# Detección robusta de Código de la cuenta
for i, val in enumerate(row_vals_hdr, 1):
    if 'código' in val or 'codigo' in val:
        col_cta_idx = i
        break

if col_cta_idx is None:
    for i, val in enumerate(row_vals_hdr, 1):
        if 'cuenta' in val and 'etiqueta' not in val:
            col_cta_idx = i
            break

for i, val in enumerate(row_vals_hdr, 1):
    if 'nombre' in val:
        col_nombre_idx = i

print("NUEVO col_cta_idx:", col_cta_idx)
print("col_nombre_idx:", col_nombre_idx)

curr_ref = ""
curr_cc = "SIN CENTRO DE COSTOS"

for row in range(header_row_idx + 1, ws.max_row + 1):
    doc_val = str(ws.cell(row, col_doc_idx).value or '').strip().replace('.0', '')
    cta_val = str(ws.cell(row, col_cta_idx).value or '').strip().replace('.0', '') if col_cta_idx else ''
    nombre_val = str(ws.cell(row, col_nombre_idx).value or '').strip().upper() if col_nombre_idx else ''
    
    # Detectar si esta fila es cabecera
    is_head = False
    if doc_val and doc_val not in ['', 'NAN', 'NONE']:
        # Omitir fechas
        if not (doc_val.startswith('20') and '-' in doc_val and sum(c.isdigit() for c in doc_val) >= 4):
            if any(cta_val.startswith(c) for c in cuentas_cabecera):
                is_head = True
                
    # Omitir filas de totales
    if nombre_val.startswith('TOTAL') or nombre_val.startswith('SUMA'):
        is_head = False
        
    if is_head:
        if doc_val in ref_map:
            curr_ref, curr_cc = ref_map[doc_val]
            if not curr_cc or curr_cc in ['', 'nan', 'NaN', 'None', '0', '0.0']:
                curr_cc = "SIN CENTRO DE COSTOS"
        else:
            curr_ref = ""
            curr_cc = "SIN CENTRO DE COSTOS"
            
    if row in [322, 327]:
        print(f"Row {row}: doc_val='{doc_val}', is_head={is_head}, curr_ref='{curr_ref}', curr_cc='{curr_cc}'")
