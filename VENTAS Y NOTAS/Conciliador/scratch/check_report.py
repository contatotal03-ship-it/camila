import pandas as pd
import os

def check_report(query):
    path = r"C:\Users\Admin\Downloads\Reporte_Conciliador_Compras.xlsx"
    xl = pd.ExcelFile(path)
    for sheet in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        mask = df.astype(str).apply(lambda x: x.str.contains(query, case=False, na=False)).any(axis=1)
        matches = df[mask]
        if not matches.empty:
            print(f"\nMatches for '{query}' in sheet '{sheet}':")
            print(matches.to_string())

check_report("1156")
check_report("BILL/2026/03/0014")
