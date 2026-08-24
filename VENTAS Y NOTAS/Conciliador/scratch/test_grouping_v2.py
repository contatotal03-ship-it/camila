import pandas as pd
import zipfile
import io
import os

# Mock the configuration functions
def get_cuentas_total_compras():
    return ["2205", "2335", "2365", "2367", "2368"]

def repair_excel(path_or_file):
    with open(path_or_file, 'rb') as f:
        file_data = f.read()
    in_buffer = io.BytesIO(file_data)
    out_buffer = io.BytesIO()
    with zipfile.ZipFile(in_buffer, 'r') as zin:
        with zipfile.ZipFile(out_buffer, 'w') as zout:
            for item in zin.infolist():
                if item.filename == 'xl/styles.xml': continue
                zout.writestr(item, zin.read(item.filename))
    out_buffer.seek(0)
    return out_buffer

def test_new_logic(path):
    repaired = repair_excel(path)
    df = pd.read_excel(repaired, sheet_name="FACTURA DE COMPRA", engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    
    col_cuenta = next((c for c in df.columns if 'cuenta' in c.lower() and 'código' in c.lower()), 'Código de la cuenta')
    df['Cod_Cuenta'] = df[col_cuenta].astype(str).str.replace('.0', '', regex=False).str.strip()
    
    cuentas_cabecera = get_cuentas_total_compras()

    def is_new_transaction(row):
        doc = str(row['Documento']).strip().upper()
        cta = str(row['Cod_Cuenta']).strip()
        if doc in ['', 'NAN', 'NONE']: return False
        if doc.startswith('20') and '-' in doc and sum(c.isdigit() for c in doc) >= 4:
            return False
        if any(cta.startswith(c) for c in cuentas_cabecera):
            return True
        return False

    df['_IsHead'] = df.apply(is_new_transaction, axis=1)
    if not df.empty: df.loc[0, '_IsHead'] = True
    df['_GroupID'] = df['_IsHead'].cumsum()
    
    # Show results for rows 40-70
    cols = ['Documento', 'Cod_Cuenta', '_IsHead', '_GroupID']
    if 'Etiqueta de la cuenta' in df.columns: cols.append('Etiqueta de la cuenta')
    if 'Nombre' in df.columns: cols.append('Nombre')
    
    print(df[cols].iloc[40:70].to_string())

test_new_logic(r"C:\Users\Admin\Downloads\MARZO ODOO.xlsx")
