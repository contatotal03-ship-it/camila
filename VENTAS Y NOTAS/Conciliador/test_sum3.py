import pandas as pd
f_tmp = r'C:\Users\Admin\Downloads\camila\VENTAS DESCARGADAS DE LA DIAN VENTAS\Resumen_Facturas_Completo.xlsx'
try:
    xls_det = pd.ExcelFile(f_tmp)
    df_det = pd.read_excel(xls_det, sheet_name='Det_Facturas_Emitidas')
    c_v = next((c for c in df_det.columns if 'SUBTOTAL' in str(c).upper() and 'SIN' in str(c).upper()), None)
    c_d = next((c for c in df_det.columns if 'DESCRIP' in str(c).upper()), None)
    df_det['VAL'] = pd.to_numeric(df_det[c_v], errors='coerce').fillna(0)
    df_270 = df_det[(df_det['VAL'] >= 270199) & (df_det['VAL'] <= 270201)]
    if not df_270.empty:
        print('Encontrado valor ~270200:')
        for idx, row in df_270.iterrows():
            print(f"{row[c_d]} -> {row['VAL']}")
    else:
        print('No hay ningun valor ~270200 en toda la hoja.')
except Exception as e:
    print('Error:', e)
