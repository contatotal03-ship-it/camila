import pandas as pd
import os

def search_token_anywhere(query):
    path = r"C:\Users\Admin\Downloads\TOKEN DIAN..xlsx"
    df = pd.read_excel(path)
    # Search in all columns
    mask = df.astype(str).apply(lambda x: x.str.contains(query, case=False, na=False)).any(axis=1)
    matches = df[mask]
    print(f"\nMatches for '{query}' anywhere:")
    print(matches[['Prefijo', 'Folio', 'Nombre Emisor', 'Total']].to_string())

search_token_anywhere("FELE")
search_token_anywhere("1156")
