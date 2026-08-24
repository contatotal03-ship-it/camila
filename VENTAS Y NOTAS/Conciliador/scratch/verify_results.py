import pandas as pd
import os

def check_results(path, query):
    xl = pd.ExcelFile(path)
    for sheet in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        mask = df.astype(str).apply(lambda x: x.str.contains(query, case=False, na=False)).any(axis=1)
        matches = df[mask]
        if not matches.empty:
            print(f"\nMatches for '{query}' in sheet '{sheet}':")
            print(matches[['Numero_Factura', 'Nombre_Tercero', 'IVA_Sistema', 'IVA_DIAN', 'Total_Sistema', 'Total_DIAN', 'Estado']].to_string())

path = r"C:\Users\Admin\Downloads\Reporte_Conciliador_Compras_REVISADO.xlsx"
check_results(path, "FELE1156")
check_results(path, "NUMERA")
