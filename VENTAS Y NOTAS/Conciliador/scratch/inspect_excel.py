import pandas as pd
import os

def inspect_file(path):
    print(f"\n--- Inspecting: {path} ---")
    if not os.path.exists(path):
        print(f"File not found: {path}")
        return
    
    try:
        # Load the excel file
        xl = pd.ExcelFile(path)
        print(f"Sheets: {xl.sheet_names}")
        
        for sheet in xl.sheet_names[:2]: # Check first 2 sheets
            print(f"\nSheet: {sheet}")
            df = pd.read_excel(path, sheet_name=sheet, nrows=10)
            print("Columns:")
            print(df.columns.tolist())
            print("\nFirst 5 rows:")
            print(df.head(5).to_string())
    except Exception as e:
        print(f"Error reading {path}: {e}")

inspect_file(r"C:\Users\Admin\Downloads\MARZO ODOO.xlsx")
inspect_file(r"C:\Users\Admin\Downloads\TOKEN DIAN..xlsx")
inspect_file(r"C:\Users\Admin\Downloads\Reporte_Conciliador_Compras.xlsx")
