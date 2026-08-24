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
    if not os.path.exists(path):
        print(f"File not found: {path}")
        return
    
    try:
        # Try normal first
        try:
            xl = pd.ExcelFile(path)
            repaired = False
        except:
            print("Normal read failed, trying repair...")
            repaired_buffer = repair_excel(path)
            xl = pd.ExcelFile(repaired_buffer, engine='openpyxl')
            repaired = True
            
        sheets = xl.sheet_names
        print(f"Sheets found: {sheets}")
        
        valid_sheets = ["COMPRAS", "FACTURA DE COMPRA", "DOCUMENTO EQUIVALENTE A FACTURA", "DOC. SOPORTE"]
        
        for sheet in sheets:
            if sheet.strip().upper() not in valid_sheets:
                print(f"Skipping sheet: {sheet}")
                continue
                
            print(f"\nSheet: {sheet}")
            if repaired:
                repaired_buffer.seek(0)
                df = pd.read_excel(repaired_buffer, sheet_name=sheet, nrows=50, engine='openpyxl')
            else:
                df = pd.read_excel(path, sheet_name=sheet, nrows=50)
                
            print("Columns:")
            print(df.columns.tolist())
            print("\nFirst 30 rows:")
            # Use to_string to see all columns
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', 1000)
            print(df.head(30).to_string())
    except Exception as e:
        print(f"Error reading {path}: {e}")
        import traceback
        traceback.print_exc()

inspect_file(r"C:\Users\Admin\Downloads\MARZO ODOO.xlsx")
