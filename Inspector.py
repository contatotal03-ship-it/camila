import pandas as pd
import json

ruta = r"C:\Users\Admin\Downloads\camila\VENTAS Y NOTAS\VENTAS  & NOTAS.xlsx"
df = pd.read_excel(ruta, sheet_name='FACTURACION SAN JOSE GUAVIARE')

# Print columns
print("COLUMNS:")
print(df.columns.tolist())

# Print first 10 rows safely
print("\nFIRST 10 ROWS:")
print(df.head(10).to_string())
