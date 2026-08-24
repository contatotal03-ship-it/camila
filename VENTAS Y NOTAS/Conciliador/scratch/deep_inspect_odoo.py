import pandas as pd
import zipfile
import io
import os

def repair_excel(path_or_file):
    if isinstance(path_or_file, str) and os.path.exists(path_or_file):
        with open(path_or_file, 'rb') as f:
            file_data = f.read()
    elif hasattr(path_or_file, 'read'):
        if hasattr(path_or_file, 'seek'):
            path_or_file.seek(0)
        file_data = path_or_file.read()
    else:
        return None
    
    in_buffer = io.BytesIO(file_data)
    out_buffer = io.BytesIO()
    
    with zipfile.ZipFile(in_buffer, 'r') as zin:
        with zipfile.ZipFile(out_buffer, 'w') as zout:
            for item in zin.infolist():
                if item.filename == 'xl/styles.xml':
                    continue
                zout.writestr(item, zin.read(item.filename))
    
    out_buffer.seek(0)
    return out_buffer

def inspect_file(path):
    print(f"\n--- Inspecting: {path} ---")
    
    try:
        repaired_buffer = repair_excel(path)
        xl = pd.ExcelFile(repaired_buffer, engine='openpyxl')
        sheets = xl.sheet_names
        
        for sheet in sheets:
            if sheet.strip().upper() not in ["FACTURA DE COMPRA", "COMPRAS"]:
                continue
                
            print(f"\nSheet: {sheet}")
            repaired_buffer.seek(0)
            # Read first 100 rows to find patterns
            df = pd.read_excel(repaired_buffer, sheet_name=sheet, nrows=100, engine='openpyxl')
            
            # Print columns with their index
            for i, col in enumerate(df.columns):
                print(f"{i}: {col}")
            
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', 1000)
            print("\nSample Data (Rows 0-20):")
            print(df.head(20).to_string())
            
            print("\nCheck for NaNs in 'Documento':")
            print(df['Documento'].isna().value_counts())
            
    except Exception as e:
        print(f"Error: {e}")

inspect_file(r"C:\Users\Admin\Downloads\MARZO ODOO.xlsx")
