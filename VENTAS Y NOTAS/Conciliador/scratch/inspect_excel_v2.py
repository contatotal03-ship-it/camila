import pandas as pd
import zipfile
import io
import os

def robust_read_excel(path_or_file, **kwargs):
    try:
        return pd.read_excel(path_or_file, **kwargs)
    except Exception as e:
        error_msg = str(e)
        if "could not read stylesheet" in error_msg or "stylesheet" in error_msg.lower():
            print(f"Detectado error de estilos en Excel. Intentando auto-reparación...")
            try:
                if isinstance(path_or_file, str) and os.path.exists(path_or_file):
                    with open(path_or_file, 'rb') as f:
                        file_data = f.read()
                elif hasattr(path_or_file, 'read'):
                    if hasattr(path_or_file, 'seek'):
                        path_or_file.seek(0)
                    file_data = path_or_file.read()
                else:
                    raise e
                
                in_buffer = io.BytesIO(file_data)
                out_buffer = io.BytesIO()
                
                with zipfile.ZipFile(in_buffer, 'r') as zin:
                    with zipfile.ZipFile(out_buffer, 'w') as zout:
                        for item in zin.infolist():
                            if item.filename == 'xl/styles.xml':
                                continue
                            zout.writestr(item, zin.read(item.filename))
                
                out_buffer.seek(0)
                if 'engine' not in kwargs:
                    kwargs['engine'] = 'openpyxl'
                return pd.read_excel(out_buffer, **kwargs)
            except Exception as repair_e:
                print(f"La auto-reparación falló: {repair_e}")
                raise e
        else:
            raise e

def inspect_file(path):
    print(f"\n--- Inspecting: {path} ---")
    if not os.path.exists(path):
        print(f"File not found: {path}")
        return
    
    try:
        # Get sheet names
        try:
            xl = pd.ExcelFile(path)
            sheets = xl.sheet_names
        except:
            # If ExcelFile fails, we might need the repair logic for it too but let's just try reading directly if possible
            sheets = ['Sheet1'] # Fallback
            
        for sheet in sheets:
            print(f"\nSheet: {sheet}")
            df = robust_read_excel(path, sheet_name=sheet, nrows=20)
            print("Columns:")
            print(df.columns.tolist())
            print("\nFirst 10 rows:")
            # Displaying more rows to see the structure of transactions
            print(df.head(15).to_string())
    except Exception as e:
        print(f"Error reading {path}: {e}")

inspect_file(r"C:\Users\Admin\Downloads\MARZO ODOO.xlsx")
