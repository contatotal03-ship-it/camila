import pandas as pd
import numpy as np

f_tmp = r'C:\Users\Admin\Downloads\camila\VENTAS DESCARGADAS DE LA DIAN VENTAS\Resumen_Facturas_Completo.xlsx'
try:
    xls_det = pd.ExcelFile(f_tmp)
    df_det = pd.read_excel(xls_det, sheet_name='Det_Facturas_Emitidas')
    df_det.columns = [str(c).strip() for c in df_det.columns]

    c_d = next((c for c in df_det.columns if 'DESCRIP' in c.upper()), None)
    c_v = next((c for c in df_det.columns if 'SUBTOTAL' in c.upper() and 'SIN' in c.upper()), 
               next((c for c in df_det.columns if 'TOTAL' in c.upper() and 'LINEA' in c.upper()), None))

    df_det['DESC_UP'] = df_det[c_d].astype(str).str.upper()
    df_det['RAW_VAL'] = df_det[c_v]

    # Filtramos por las que dice TRASPORTE o TRANSPORTE
    mask_tr = df_det['DESC_UP'].str.contains('TRANSPORTE|TRASPORTE', na=False, regex=True)
    df_tr = df_det[mask_tr].copy()

    df_tr['VAL_NUM'] = pd.to_numeric(df_tr['RAW_VAL'], errors='coerce')
    suma_num = df_tr['VAL_NUM'].sum()
    print('Suma numerica (coerce) Facturas Emitidas:', suma_num)
    
    # Hacer lo mismo para Notas Credito
    df_nc = pd.read_excel(xls_det, sheet_name='Det_Notas_Credito_Emitidas')
    df_nc.columns = [str(c).strip() for c in df_nc.columns]
    c_d_nc = next((c for c in df_nc.columns if 'DESCRIP' in c.upper()), None)
    c_v_nc = next((c for c in df_nc.columns if 'SUBTOTAL' in c.upper() and 'SIN' in c.upper()), 
                  next((c for c in df_nc.columns if 'TOTAL' in c.upper() and 'LINEA' in c.upper()), None))
    df_nc['DESC_UP'] = df_nc[c_d_nc].astype(str).str.upper()
    mask_nc = df_nc['DESC_UP'].str.contains('TRANSPORTE|TRASPORTE', na=False, regex=True)
    df_tr_nc = df_nc[mask_nc].copy()
    df_tr_nc['VAL_NUM'] = pd.to_numeric(df_tr_nc[c_v_nc], errors='coerce')
    suma_nc = df_tr_nc['VAL_NUM'].sum()
    print('Suma numerica (coerce) Notas Credito:', suma_nc)
    print('SUMA TOTAL AMBOS:', suma_num + suma_nc)

except Exception as e:
    print('Error:', str(e))
