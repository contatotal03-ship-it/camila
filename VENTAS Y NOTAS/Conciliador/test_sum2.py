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

    mask_tr = df_det['DESC_UP'].str.contains('TRANSPORTE|TRASPORTE', na=False, regex=True)
    df_tr = df_det[mask_tr].copy()

    df_tr['VAL_NUM'] = pd.to_numeric(df_tr['RAW_VAL'], errors='coerce')
    
    # Filas que perdieron su valor o son cero
    print('Total sumado num:', df_tr['VAL_NUM'].sum())
    
    # Que pasa con los NaN?
    df_nan = df_tr[df_tr['VAL_NUM'].isna()]
    if not df_nan.empty:
        print('NaN encontrados:', len(df_nan))
        for idx, row in df_nan.iterrows():
            print(f"[{idx}] {row['RAW_VAL']}")
            
    # Intento parsear comas
    df_tr['VAL_CLEAN'] = df_tr['RAW_VAL'].astype(str).str.replace(r'[^\d,\.-]', '', regex=True)
    df_tr['VAL_CLEAN'] = df_tr['VAL_CLEAN'].str.replace(',', '.')
    df_tr['VAL_NUM_2'] = pd.to_numeric(df_tr['VAL_CLEAN'], errors='coerce')
    print('Suma despues de clean:', df_tr['VAL_NUM_2'].sum())
    
except Exception as e:
    print('Error:', str(e))
