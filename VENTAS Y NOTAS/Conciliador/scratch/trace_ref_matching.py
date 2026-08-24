import sys
sys.path.append(r'c:\Users\Admin\Downloads\camila\VENTAS Y NOTAS\Conciliador')
from core.excel_utils import robust_read_excel
import openpyxl

ruta_sistema = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\marzo 2026 odoo  - ORIGINAL.xlsx'
ruta_ref = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\PROVEEDORES REFERENCIA MARZO.xlsx'

df_ref = robust_read_excel(ruta_ref)
print("df_ref columns:", list(df_ref.columns))

# Crear el ref_map exactamente como lo hace guardar_referencias_en_sistema
ref_map = {}
for _, r in df_ref.iterrows():
    num = str(r.get('Número', '')).strip()
    if num:
        ref_map[num] = (str(r.get('Referencia', '')).strip(), str(r.get('Centro de costo', '')).strip())

print("Total keys in ref_map:", len(ref_map))
print("Some keys in ref_map:", list(ref_map.keys())[:10])

# Vamos a ver si 'FE67253963' o 'H130002772' están en ref_map
print("'FE67253963' in ref_map:", 'FE67253963' in ref_map)
if 'FE67253963' in ref_map:
    print("Value for 'FE67253963':", ref_map['FE67253963'])

print("'H130002772' in ref_map:", 'H130002772' in ref_map)
if 'H130002772' in ref_map:
    print("Value for 'H130002772':", ref_map['H130002772'])

# Leer hoja de Odoo
wb = openpyxl.load_workbook(ruta_sistema)
ws = wb['Compras']
# Encontrar la columna de Documento
header_row_idx = None
col_doc_idx = None
for row in range(1, 15):
    row_vals = [str(ws.cell(row, col).value).strip().lower() for col in range(1, ws.max_column + 1)]
    if 'documento' in row_vals:
        header_row_idx = row
        col_doc_idx = row_vals.index('documento') + 1
        break

print("header_row_idx:", header_row_idx, "col_doc_idx:", col_doc_idx)

# Ver las filas 322 y 327 y ver qué valores de documento tienen y qué buscaría el script
for r in [322, 327]:
    doc_val = str(ws.cell(r, col_doc_idx).value or '').strip().replace('.0', '')
    print(f"Fila {r}: doc_val='{doc_val}' (repr={repr(ws.cell(r, col_doc_idx).value)})")
    # ¿Es cabecera?
    # Buscar cuenta
    col_cta_idx = None
    row_vals_hdr = [str(ws.cell(header_row_idx, col).value).strip().lower() for col in range(1, ws.max_column + 1)]
    for i, val in enumerate(row_vals_hdr, 1):
        if 'cuenta' in val or 'código' in val or 'codigo' in val:
            col_cta_idx = i
            break
    cta_val = str(ws.cell(r, col_cta_idx).value or '').strip().replace('.0', '') if col_cta_idx else ''
    print(f"  cta_val='{cta_val}'")
    
    # Comprobar si coincide con cuentas_cabecera (get_cuentas_total_compras())
    from core.config_manager import get_cuentas_total_compras
    cuentas_cabecera = get_cuentas_total_compras()
    is_head = False
    if doc_val and doc_val not in ['', 'NAN', 'NONE']:
        if not (doc_val.startswith('20') and '-' in doc_val and sum(c.isdigit() for c in doc_val) >= 4):
            if any(cta_val.startswith(c) for c in cuentas_cabecera):
                is_head = True
    print(f"  is_head={is_head}")
    if is_head:
        if doc_val in ref_map:
            print(f"  -> MATCH ENCONTRADO EN MAPA: {ref_map[doc_val]}")
        else:
            print(f"  -> NO match en mapa para '{doc_val}'")
