import sys
sys.path.append(r'c:\Users\Admin\Downloads\camila\VENTAS Y NOTAS\Conciliador')
from core.compras_logic import procesar_conciliacion_compras
from core.excel_utils import robust_read_excel
import openpyxl
import os

ruta_dian = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\dian marzo.xlsx'
ruta_sistema = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\marzo 2026 odoo  - ORIGINAL.xlsx'
ruta_ref = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\PROVEEDORES REFERENCIA MARZO.xlsx'

ruta_dest_sistema = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\marzo 2026 odoo.xlsx'
ruta_output = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD\marzo\Reporte_Conciliador_Compras.xlsx'

if os.path.exists(ruta_output):
    try:
        os.remove(ruta_output)
        print("Reporte anterior eliminado.")
    except Exception as e:
        print("No se pudo eliminar reporte anterior:", e)

# Ejecutar procesar_conciliacion_compras
print("Ejecutando conciliacion completa...")
try:
    procesar_conciliacion_compras(
        ruta_sistema=ruta_sistema,
        ruta_dian=ruta_dian,
        ruta_salida=ruta_output,
        ruta_referencias=ruta_ref
    )
    print("Ejecución finalizada correctamente.")
except Exception as e:
    import traceback
    traceback.print_exc()
    print("Error:", e)

# Verificar sumas y diferencias de las tablas en el reporte generado
wb = openpyxl.load_workbook(ruta_output, data_only=True)

for sheet_name in ['Resumen por Cuenta y CC', 'Resumen por Hoja de Compras']:
    ws = wb[sheet_name]
    print(f'\n--- HOJA: {sheet_name} ---')
    
    current_table = 'Desconocido'
    data_rows = []
    
    for r in range(1, ws.max_row + 1):
        v_b = ws.cell(r, 2).value # Col B
        v_d = ws.cell(r, 4).value # Col D (Debe)
        v_e = ws.cell(r, 5).value # Col E (Haber)
        
        if v_b and ('CENTRO DE COSTO:' in str(v_b) or 'HOJA COMPRAS:' in str(v_b)):
            current_table = v_b
            data_rows = []
            continue
            
        if v_b == 'TOTALES':
            s_debe = sum(r[0] for r in data_rows)
            s_haber = sum(r[1] for r in data_rows)
            diff = s_debe - s_haber
            print(f'Tabla {current_table}:')
            print(f'  Suma manual Debe: {s_debe:,.2f}, Haber: {s_haber:,.2f}, Dif: {diff:,.2f}')
            continue
            
        if v_b == 'DIFERENCIA':
            continue
            
        # Si es una fila de datos
        if v_b and v_b != 'Código de la cuenta' and v_b != 'TOTALES' and v_b != 'DIFERENCIA':
            try:
                val_d = float(v_d) if v_d and v_d != '-' else 0.0
                val_e = float(v_e) if v_e and v_e != '-' else 0.0
                data_rows.append((val_d, val_e))
            except:
                pass
