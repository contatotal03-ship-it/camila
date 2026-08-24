import pandas as pd
import os

def search_token(query):
    path = r"C:\Users\Admin\Downloads\TOKEN DIAN..xlsx"
    df = pd.read_excel(path)
    # Search in Nombre Emisor
    matches = df[df['Nombre Emisor'].str.contains(query, case=False, na=False)]
    print(f"\nMatches for '{query}':")
    print(matches[['Prefijo', 'Folio', 'Nombre Emisor', 'Total']].to_string())

search_token("MACARENA")
search_token("NUMERA")
search_token("AUTOTECNICA")
