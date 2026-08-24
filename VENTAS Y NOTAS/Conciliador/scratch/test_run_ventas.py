import sys
import os
import pandas as pd

sys.path.append(os.getcwd())

from core.ventas_logic import procesar_conciliacion_ventas

def test_run():
    ruta_contabilidad = r'C:\Users\Admin\Downloads\camila\CONTABILIDAD 2026.xlsx'
    ruta_enviados = r'C:\Users\Admin\Downloads\camila\Enviados 2026.xlsx'
    ruta_salida = r'C:\Users\Admin\Downloads\camila\Reporte_Prueba_Ventas.xlsx'
    
    print("Iniciando procesamiento de ventas...")
    try:
        resultado = procesar_conciliacion_ventas(ruta_contabilidad, ruta_enviados, ruta_salida)
        print("\n--- RESULTADO DEL PROCESAMIENTO ---")
        print(f"Total facturas procesadas: {resultado['total_procesadas']}")
        print(f"Cuadre Perfecto: {resultado['cuadre']}")
        print(f"Diferencias en Valores: {resultado['diferencias']}")
        print(f"Solo en Sistema: {resultado['solo_sis']}")
        print(f"Solo en DIAN: {resultado['solo_dian']}")
        
        # Analizar una muestra de las diferencias si existen
        df_res = pd.read_excel(ruta_salida, sheet_name='Resumen Total Cruce')
        if not df_res.empty:
            print("\nMuestra de facturas con problemas (Estado != Cuadre Perfecto):")
            problemas = df_res[df_res['Estado'] != 'Cuadre Perfecto'].head(10)
            print(problemas[['Numero_Factura', 'IVA_Sistema', 'IVA_DIAN', 'Estado']])
            
    except Exception as e:
        print(f"Error durante el proceso: {e}")

if __name__ == '__main__':
    test_run()
