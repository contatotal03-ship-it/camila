import pandas as pd # type: ignore
import numpy as np # type: ignore
import datetime
import os
import difflib
from .config_manager import get_hojas_compras, get_dian_compras_incluir, get_cuentas_iva_compras, get_cuentas_total_compras, get_dian_grupo_compras, get_dian_grupo_ventas, get_dian_tipos_validos_compras # type: ignore
from .excel_utils import robust_read_excel, robust_ExcelFile # type: ignore

def calcular_similitud_nombres(nom1, nom2):
    if not nom1 or not nom2: return 0.0
    s1 = str(nom1).strip().upper()
    s2 = str(nom2).strip().upper()
    
    # 1. Secuencia exacta (difflib)
    score_seq = difflib.SequenceMatcher(None, s1, s2).ratio()
    
    # 2. Mismas palabras pero en distinto orden (Jaccard sobre palabras)
    # Limpiamos caracteres que puedan interferir en la separación
    w1 = set(s1.replace('.', '').replace(',', '').split())
    w2 = set(s2.replace('.', '').replace(',', '').split())
    
    if w1 and w2:
        intersection = w1.intersection(w2)
        union = w1.union(w2)
        score_set = len(intersection) / len(union)
    else:
        score_set = 0.0
        
    return max(score_seq, score_set)

def guardar_referencias_en_sistema(ruta_sistema, df_ref, hojas_validas):
    """
    Agrega o actualiza las columnas 'Referencia' y 'Centro de costo' al archivo original
    de Odoo (ruta_sistema) preservando el formato original de todas las celdas y estilos.
    """
    if df_ref is None:
        return
        
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border
    
    try:
        # Cargar el libro preservando estilos y formato original
        wb = openpyxl.load_workbook(ruta_sistema)
        hubo_cambios = False
        
        # Crear un mapa rápido de Número -> (Referencia, Centro de costo)
        ref_map = {}
        for _, r in df_ref.iterrows():
            num = str(r.get('Número', '')).strip()
            if num:
                ref_map[num] = (str(r.get('Referencia', '')).strip(), str(r.get('Centro de costo', '')).strip())
                
        for sheet_name in wb.sheetnames:
            if sheet_name.strip().upper() not in hojas_validas:
                continue
                
            ws = wb[sheet_name]
            
            # 1. Encontrar la fila de cabecera y el índice de las columnas de interés
            # Buscamos la primera fila que contenga la palabra 'Documento' o 'Impuestos'
            header_row_idx = None
            col_doc_idx = None
            col_ref_idx = None
            col_cc_idx = None
            max_col = ws.max_column
            
            for row in range(1, min(ws.max_row + 1, 15)):
                row_vals = [str(ws.cell(row, col).value).strip().lower() for col in range(1, max_col + 1)]
                if 'documento' in row_vals:
                    header_row_idx = row
                    col_doc_idx = row_vals.index('documento') + 1
                    
                    if 'referencia' in row_vals:
                        col_ref_idx = row_vals.index('referencia') + 1
                    if 'centro de costo' in row_vals or 'centro_costo' in row_vals:
                        col_cc_idx = next((i for i, val in enumerate(row_vals) if 'centro' in val or 'cc' in val), None)
                        if col_cc_idx is not None:
                            col_cc_idx += 1
                    break
            
            if not header_row_idx or not col_doc_idx:
                continue # No es una hoja de formato contable válido
                
            # 2. Si no existen las columnas, las creamos al final absoluto del libro
            # para evitar sobreescribir la tabla de resumen de impuestos aplicados a la derecha.
            if not col_ref_idx or not col_cc_idx:
                last_header_col = ws.max_column
                        
                # Asignamos las nuevas columnas justo después del final de toda la hoja
                if not col_ref_idx:
                    col_ref_idx = last_header_col + 1
                    ws.cell(header_row_idx, col_ref_idx).value = 'Referencia'
                    col_cc_idx = last_header_col + 2
                    ws.cell(header_row_idx, col_cc_idx).value = 'Centro de costo'
                elif not col_cc_idx:
                    col_cc_idx = max(col_ref_idx + 1, last_header_col + 1)
                    ws.cell(header_row_idx, col_cc_idx).value = 'Centro de costo'
                    
                # Copiar estilo de la columna 7 (que representa el estilo de la tabla principal)
                # si las nuevas columnas se colocan más a la derecha.
                col_estilo = 7 if col_ref_idx > 7 else col_ref_idx - 1
                estilo_origen = ws.cell(header_row_idx, col_estilo)
                for col_idx in [col_ref_idx, col_cc_idx]:
                    target = ws.cell(header_row_idx, col_idx)
                    if estilo_origen.has_style:
                        target.font = Font(name=estilo_origen.font.name, size=estilo_origen.font.size,
                                           bold=estilo_origen.font.bold, italic=estilo_origen.font.italic,
                                           color=estilo_origen.font.color)
                        if estilo_origen.fill:
                            target.fill = PatternFill(fill_type=estilo_origen.fill.fill_type,
                                                      start_color=estilo_origen.fill.start_color,
                                                      end_color=estilo_origen.fill.end_color)
                        if estilo_origen.border:
                            target.border = Border(left=estilo_origen.border.left, right=estilo_origen.border.right,
                                                   top=estilo_origen.border.top, bottom=estilo_origen.border.bottom)
                        if estilo_origen.alignment:
                            target.alignment = Alignment(horizontal=estilo_origen.alignment.horizontal,
                                                         vertical=estilo_origen.alignment.vertical)
                
            # 3. Llenar/actualizar los valores en cada fila de datos de la hoja
            # Primero leemos qué cuentas inician transacción
            cuentas_cabecera = get_cuentas_total_compras()
            
            # Encontrar índices de código de cuenta
            col_cta_idx = None
            col_nombre_idx = None
            row_vals_hdr = [str(ws.cell(header_row_idx, col).value).strip().lower() for col in range(1, ws.max_column + 1)]
            
            # Buscamos primero código/codigo de cuenta para evitar falsos positivos con 'etiqueta de la cuenta'
            for i, val in enumerate(row_vals_hdr, 1):
                if 'código' in val or 'codigo' in val:
                    col_cta_idx = i
                    break
            
            # Si no se encuentra estricto, buscar 'cuenta' que no sea etiqueta
            if col_cta_idx is None:
                for i, val in enumerate(row_vals_hdr, 1):
                    if 'cuenta' in val and 'etiqueta' not in val:
                        col_cta_idx = i
                        break
                        
            for i, val in enumerate(row_vals_hdr, 1):
                if 'nombre' in val:
                    col_nombre_idx = i
            
            curr_ref = ""
            curr_cc = "SIN CENTRO DE COSTOS"
            
            for row in range(header_row_idx + 1, ws.max_row + 1):
                doc_val = str(ws.cell(row, col_doc_idx).value or '').strip().replace('.0', '')
                cta_val = str(ws.cell(row, col_cta_idx).value or '').strip().replace('.0', '') if col_cta_idx else ''
                nombre_val = str(ws.cell(row, col_nombre_idx).value or '').strip().upper() if col_nombre_idx else ''
                
                # Detectar si esta fila es cabecera
                is_head = False
                if doc_val and doc_val not in ['', 'NAN', 'NONE']:
                    # Omitir fechas
                    if not (doc_val.startswith('20') and '-' in doc_val and sum(c.isdigit() for c in doc_val) >= 4):
                        if sheet_name.strip().upper() == 'DOC. SOPORTE':
                            is_head = True
                        elif any(cta_val.startswith(c) for c in cuentas_cabecera):
                            is_head = True
                            
                # Omitir filas de totales
                if nombre_val.startswith('TOTAL') or nombre_val.startswith('SUMA'):
                    is_head = False
                    
                if is_head:
                    # Buscar en el mapa
                    if doc_val in ref_map:
                        curr_ref, curr_cc = ref_map[doc_val]
                        if not curr_cc or curr_cc in ['', 'nan', 'NaN', 'None', '0', '0.0']:
                            curr_cc = "SIN CENTRO DE COSTOS"
                    else:
                        curr_ref = ""
                        curr_cc = "SIN CENTRO DE COSTOS"
                
                # Escribir en la celda
                ws.cell(row, col_ref_idx).value = curr_ref
                ws.cell(row, col_cc_idx).value = curr_cc
                
                # Copiar estilo de celda principal
                col_estilo = 7 if col_ref_idx > 7 else col_ref_idx - 1
                estilo_origen = ws.cell(row, col_estilo)
                for col_idx in [col_ref_idx, col_cc_idx]:
                    target = ws.cell(row, col_idx)
                    if estilo_origen.has_style:
                        target.font = Font(name=estilo_origen.font.name, size=estilo_origen.font.size,
                                           bold=estilo_origen.font.bold, italic=estilo_origen.font.italic,
                                           color=estilo_origen.font.color)
                        if estilo_origen.fill:
                            target.fill = PatternFill(fill_type=estilo_origen.fill.fill_type,
                                                      start_color=estilo_origen.fill.start_color,
                                                      end_color=estilo_origen.fill.end_color)
                        if estilo_origen.border:
                            target.border = Border(left=estilo_origen.border.left, right=estilo_origen.border.right,
                                                   top=estilo_origen.border.top, bottom=estilo_origen.border.bottom)
                        if estilo_origen.alignment:
                            target.alignment = Alignment(horizontal=estilo_origen.alignment.horizontal,
                                                         vertical=estilo_origen.alignment.vertical)
                            
            hubo_cambios = True
            
        if hubo_cambios:
            wb.save(ruta_sistema)
            print("Archivo de Odoo guardado con éxito y formato 100% preservado.")
            
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error escribiendo en archivo original de Odoo: {e}")

def escribir_una_tabla(ws, r_start, c_start, titulo, df_g, header_fill, sub_header_fill, total_fill, header_font, normal_font, bold_font, thin_border):
    from openpyxl.styles import Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    # 1. Escribir y mezclar fila de título (ej: HOJA COMPRAS: Compras)
    ws.merge_cells(start_row=r_start, start_column=c_start, end_row=r_start, end_column=c_start + 3)
    title_cell = ws.cell(r_start, c_start)
    title_cell.value = str(titulo).upper()
    title_cell.fill = header_fill
    title_cell.font = header_font
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Aplicar borde a la celda del título
    for c in range(c_start, c_start + 4):
        ws.cell(r_start, c).border = thin_border
        if c > c_start:
            ws.cell(r_start, c).fill = header_fill
            
    # 2. Escribir fila de subcabeceras (columnas)
    headers = ['Código de la cuenta', 'Etiqueta de la cuenta', 'Debe', 'Haber']
    for idx, h in enumerate(headers):
        cell = ws.cell(r_start + 1, c_start + idx)
        cell.value = h
        cell.fill = sub_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
        
    # 3. Escribir filas de datos
    start_data_row = r_start + 2
    for idx, (_, row) in enumerate(df_g.iterrows()):
        curr_row = start_data_row + idx
        
        # Código
        cell_cod = ws.cell(curr_row, c_start)
        cell_cod.value = str(row['Cod_Cuenta']).strip()
        cell_cod.font = normal_font
        cell_cod.alignment = Alignment(horizontal='left', vertical='center')
        cell_cod.border = thin_border
        
        # Etiqueta
        cell_nom = ws.cell(curr_row, c_start + 1)
        cell_nom.value = str(row['Nombre_Cuenta']).strip()
        cell_nom.font = normal_font
        cell_nom.alignment = Alignment(horizontal='left', vertical='center')
        cell_nom.border = thin_border
        
        # Debe
        cell_debe = ws.cell(curr_row, c_start + 2)
        debe_val = float(row['Monto_Debe'])
        cell_debe.value = debe_val if debe_val != 0 else "-"
        cell_debe.font = normal_font
        cell_debe.alignment = Alignment(horizontal='right', vertical='center')
        cell_debe.border = thin_border
        if debe_val != 0:
            cell_debe.number_format = '$#,##0'
            
        # Haber
        cell_haber = ws.cell(curr_row, c_start + 3)
        haber_val = float(row['Monto_Haber'])
        cell_haber.value = haber_val if haber_val != 0 else "-"
        cell_haber.font = normal_font
        cell_haber.alignment = Alignment(horizontal='right', vertical='center')
        cell_haber.border = thin_border
        if haber_val != 0:
            cell_haber.number_format = '$#,##0'
            
    # 4. Escribir fila de TOTALES
    totals_row = start_data_row + len(df_g)
    
    # Combinar Código y Etiqueta para TOTALES
    ws.merge_cells(start_row=totals_row, start_column=c_start, end_row=totals_row, end_column=c_start + 1)
    tot_label_cell = ws.cell(totals_row, c_start)
    tot_label_cell.value = "TOTALES"
    tot_label_cell.fill = total_fill
    tot_label_cell.font = header_font
    tot_label_cell.alignment = Alignment(horizontal='left', vertical='center')
    
    # Aplicar borde y relleno a las celdas mezcladas
    ws.cell(totals_row, c_start).border = thin_border
    ws.cell(totals_row, c_start + 1).border = thin_border
    ws.cell(totals_row, c_start + 1).fill = total_fill
    
    # Fórmulas de Suma
    col_letter_debe = get_column_letter(c_start + 2)
    col_letter_haber = get_column_letter(c_start + 3)
    
    cell_tot_debe = ws.cell(totals_row, c_start + 2)
    if len(df_g) > 0:
        cell_tot_debe.value = f"=SUM({col_letter_debe}{start_data_row}:{col_letter_debe}{totals_row - 1})"
    else:
        cell_tot_debe.value = 0
    cell_tot_debe.fill = total_fill
    cell_tot_debe.font = header_font
    cell_tot_debe.alignment = Alignment(horizontal='right', vertical='center')
    cell_tot_debe.border = thin_border
    cell_tot_debe.number_format = '$#,##0'
    
    cell_tot_haber = ws.cell(totals_row, c_start + 3)
    if len(df_g) > 0:
        cell_tot_haber.value = f"=SUM({col_letter_haber}{start_data_row}:{col_letter_haber}{totals_row - 1})"
    else:
        cell_tot_haber.value = 0
    cell_tot_haber.fill = total_fill
    cell_tot_haber.font = header_font
    cell_tot_haber.alignment = Alignment(horizontal='right', vertical='center')
    cell_tot_haber.border = thin_border
    cell_tot_haber.number_format = '$#,##0'
    
    # 5. Escribir fila de DIFERENCIA
    diff_row = totals_row + 1
    
    # Combinar Código y Etiqueta para DIFERENCIA
    ws.merge_cells(start_row=diff_row, start_column=c_start, end_row=diff_row, end_column=c_start + 1)
    diff_label_cell = ws.cell(diff_row, c_start)
    diff_label_cell.value = "DIFERENCIA"
    diff_label_cell.fill = total_fill
    diff_label_cell.font = header_font
    diff_label_cell.alignment = Alignment(horizontal='left', vertical='center')
    
    ws.cell(diff_row, c_start).border = thin_border
    ws.cell(diff_row, c_start + 1).border = thin_border
    ws.cell(diff_row, c_start + 1).fill = total_fill
    
    # Fórmula de Diferencia (Debe - Haber)
    cell_diff = ws.cell(diff_row, c_start + 2)
    cell_diff.value = f"={col_letter_debe}{totals_row} - {col_letter_haber}{totals_row}"
    cell_diff.fill = total_fill
    cell_diff.font = header_font
    cell_diff.alignment = Alignment(horizontal='right', vertical='center')
    cell_diff.border = thin_border
    cell_diff.number_format = '$#,##0'
    
    # Celda vacía Haber en DIFERENCIA
    cell_diff_haber = ws.cell(diff_row, c_start + 3)
    cell_diff_haber.value = "-"
    cell_diff_haber.fill = total_fill
    cell_diff_haber.font = header_font
    cell_diff_haber.alignment = Alignment(horizontal='right', vertical='center')
    cell_diff_haber.border = thin_border

def escribir_tablas_resumen_grid(writer, df_sistema, campo_grupo, nombre_hoja, titulo_prefijo):
    """
    Escribe resúmenes apilados verticalmente (uno debajo del otro) en una sola columna de tablas.
    Esto evita el scroll horizontal y hace que la lectura sea fluida y ordenada.
    """
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    # Crear la hoja vacía en openpyxl
    wb = writer.book
    if nombre_hoja in wb.sheetnames:
        del wb[nombre_hoja]
    ws = wb.create_sheet(title=nombre_hoja)
    ws.sheet_view.showGridLines = True  # Mantener las líneas de cuadrícula para que se vea profesional
    
    if df_sistema.empty:
        ws.cell(1, 1).value = "No hay datos para mostrar."
        return
        
    # Obtener grupos únicos y ordenarlos para consistencia
    grupos = sorted(df_sistema[campo_grupo].unique())
    
    # Estilos
    header_fill = PatternFill(start_color="1B263B", end_color="1B263B", fill_type="solid") # Dark Blue
    sub_header_fill = PatternFill(start_color="2E4057", end_color="2E4057", fill_type="solid") # Medium Blue
    total_fill = PatternFill(start_color="1B263B", end_color="1B263B", fill_type="solid") # Dark Blue
    
    header_font = Font(name="Calibri", size=11, color="FFFFFF", bold=True)
    normal_font = Font(name="Calibri", size=11, color="000000")
    bold_font = Font(name="Calibri", size=11, color="000000", bold=True)
    
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    
    current_row = 2
    
    for g in grupos:
        # Agrupar y procesar datos del grupo actual
        df_g = df_sistema[df_sistema[campo_grupo] == g].groupby(['Cod_Cuenta', 'Nombre_Cuenta']).agg({
            'Monto_Debe': 'sum',
            'Monto_Haber': 'sum'
        }).reset_index()
        df_g = df_g[(df_g['Monto_Debe'] > 0) | (df_g['Monto_Haber'] > 0)].copy().reset_index(drop=True)
        
        if df_g.empty:
            continue
            
        # Escribir la tabla del grupo actual en la columna B (columna 2)
        escribir_una_tabla(ws, current_row, 2, f"{titulo_prefijo}{g}", df_g, header_fill, sub_header_fill, total_fill, header_font, normal_font, bold_font, thin_border)
        
        # Avanzar el puntero de fila (altura de la tabla + 6 filas de espacio)
        current_row += len(df_g) + 6
        
    # Autoajustar anchos de columnas
    for col in range(1, 7):
        col_letter = get_column_letter(col)
        if col == 1:  # Espaciador izquierdo
            ws.column_dimensions[col_letter].width = 4
        elif col == 2:  # Código de la cuenta
            ws.column_dimensions[col_letter].width = 20
        elif col == 3:  # Etiqueta de la cuenta
            ws.column_dimensions[col_letter].width = 45
        elif col in [4, 5]:  # Debe y Haber
            ws.column_dimensions[col_letter].width = 16

def procesar_conciliacion_compras(ruta_sistema, ruta_dian, ruta_salida, ruta_referencias=None):
    """
    Lee los archivos del Sistema y de la DIAN, aplica las reglas de negocio,
    cruza la información y genera un archivo de Excel con el resultado.
    """
    
    # 1. ------------- PROCESAR DIAN -------------
    try:
        df_dian = robust_read_excel(ruta_dian)
        
        # Eliminamos todos los filtros tempranos (ni por Grupo ni por tipo_doc) para garantizar 
        # que las 296 facturas que el usuario subió en el Token se contemplen absolutamente todas en el reporte.
        prefijo_str = df_dian['Prefijo'].astype(str).replace(['nan', 'NaN', 'None'], '').str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
        folio_str = df_dian['Folio'].astype(str).replace(['nan', 'NaN', 'None'], '').str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
        
        df_dian['Numero_Factura'] = prefijo_str + folio_str
        df_dian['Folio_DIAN_Puro'] = folio_str
        df_dian['PrefijoFolio_DIAN'] = df_dian['Numero_Factura']
        df_dian['IVA_DIAN'] = pd.to_numeric(df_dian['IVA'], errors='coerce').fillna(0)
        df_dian['Total_DIAN'] = pd.to_numeric(df_dian['Total'], errors='coerce').fillna(0)
        
        # Identificar Tercero (Proveedor)
        col_emisor = next((c for c in df_dian.columns if 'emisor' in c.lower() and 'nombre' in c.lower()), 'Nombre Emisor')
        col_receptor = next((c for c in df_dian.columns if 'receptor' in c.lower() and 'nombre' in c.lower()), 'Nombre Receptor')
        col_nit_receptor = next((c for c in df_dian.columns if 'receptor' in c.lower() and 'nit' in c.lower()), 'NIT Receptor')
        
        if col_emisor not in df_dian.columns: df_dian[col_emisor] = ""
            
        nombre_dian = df_dian[col_emisor].fillna('').astype(str).str.strip()
        if 'Grupo' in df_dian.columns:
            nombre_dian = np.where(df_dian['Grupo'].astype(str).str.strip().str.upper() == 'EMITIDO', 
                                   df_dian[col_receptor].fillna(df_dian[col_nit_receptor]).fillna('').astype(str).str.strip(), 
                                   nombre_dian)
        df_dian['Nombre_Receptor_DIAN'] = nombre_dian
        
        if 'Tipo de documento' not in df_dian.columns:
            df_dian['Tipo de documento'] = ""
            
        # AGRUPAR POR FACTURA Y TERCERO para evitar sumar duplicados de distintas personas
        dian_agrupado = df_dian.groupby(['Numero_Factura', 'Nombre_Receptor_DIAN']).agg({
            'IVA_DIAN': 'sum',
            'Total_DIAN': 'sum',
            'PrefijoFolio_DIAN': 'first',
            'Folio_DIAN_Puro': 'first',
            'Tipo de documento': 'first'
        }).reset_index()
    except Exception as e:
        raise Exception(f"Error procesando token DIAN: {str(e)}")

    # 2. ------------- PROCESAR SISTEMA ----------
    try:
        df_ref = None
        if ruta_referencias:
            try:
                df_ref = robust_read_excel(ruta_referencias)
                cols_req = [c for c in ['Número', 'Referencia', 'Centro de costo'] if c in df_ref.columns]
                if 'Número' in cols_req:
                    df_ref = df_ref[cols_req].copy()
                    df_ref['Número'] = df_ref['Número'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    
                    # Eliminar duplicados de la columna 'Número' para evitar duplicidad de registros contables
                    # Priorizamos filas que tengan un Centro de costo válido (no vacío ni nulo)
                    if 'Centro de costo' in df_ref.columns:
                        df_ref['_has_cc'] = df_ref['Centro de costo'].notna() & (df_ref['Centro de costo'].astype(str).str.strip() != '')
                        df_ref = df_ref.sort_values(by='_has_cc', ascending=False)
                        df_ref.drop(columns=['_has_cc'], errors='ignore', inplace=True)
                    df_ref = df_ref.drop_duplicates(subset=['Número'], keep='first')
            except Exception as e:
                print(f"No se pudo cargar el archivo de referencias: {e}")
                df_ref = None

        hojas_validas = get_hojas_compras()
        # Escribir Referencia y Centro de costo directamente en el archivo original preservando su formato
        if df_ref is not None:
            try:
                guardar_referencias_en_sistema(ruta_sistema, df_ref, hojas_validas)
            except Exception as e:
                print(f"No se pudo guardar en el archivo original de Odoo: {e}")

        xls_sistema = robust_ExcelFile(ruta_sistema)
        lista_dfs_sistema = []
        
        all_sheets = pd.read_excel(ruta_sistema, sheet_name=None)
        hubo_cambios_sistema = False
        
        for sheet in xls_sistema.sheet_names:
            if sheet.strip().upper() not in hojas_validas:
                continue
                
            df_sheet = robust_read_excel(xls_sistema, sheet_name=sheet)
            df_sheet.columns = [str(c).strip() for c in df_sheet.columns]
            
            if 'Documento' not in df_sheet.columns:
                continue
                
            # Agregar Referencia y Centro de Costo si existe el archivo
            if df_ref is not None:
                df_sheet['_doc_str'] = df_sheet['Documento'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                # Para evitar duplicar columnas si el usuario ya lo habia corrido antes
                if 'Referencia' in df_sheet.columns: df_sheet = df_sheet.drop(columns=['Referencia'])
                if 'Centro de costo' in df_sheet.columns: df_sheet = df_sheet.drop(columns=['Centro de costo'])
                
                df_sheet = pd.merge(df_sheet, df_ref, left_on='_doc_str', right_on='Número', how='left')
                df_sheet.drop(columns=['_doc_str', 'Número'], inplace=True, errors='ignore')
                
                # Actualizar all_sheets para guardarlo después
                all_sheets[sheet] = df_sheet.copy()
                hubo_cambios_sistema = True
                
            # Inicialización robusta por defecto
            if 'Referencia' not in df_sheet.columns:
                df_sheet['Referencia'] = ''
            if 'Centro de costo' not in df_sheet.columns:
                df_sheet['Centro de costo'] = 'SIN CENTRO DE COSTOS'
                
            # Asegurar y normalizar Centro de costo en la hoja actual
            if 'Centro de costo' not in df_sheet.columns:
                df_sheet['Centro de costo'] = 'SIN CENTRO DE COSTOS'
            else:
                df_sheet['Centro de costo'] = df_sheet['Centro de costo'].fillna('').astype(str).str.strip()
                df_sheet['Centro de costo'] = np.where(df_sheet['Centro de costo'].isin(['', 'nan', 'NaN', 'None', '0', '0.0']), 'SIN CENTRO DE COSTOS', df_sheet['Centro de costo'])
                
            col_nombre = dict(zip(df_sheet.columns, df_sheet.columns)).get('Nombre', 
                            next((c for c in df_sheet.columns if 'nombre' in c.lower()), None))
            if not col_nombre: continue

            # Purga de filas de Totales Generales que vengan sin documento para no alterar los valores del último bloque
            nom_upper = df_sheet[col_nombre].astype(str).str.strip().str.upper()
            es_total = nom_upper.str.startswith('TOTAL') | nom_upper.str.startswith('SUMA')
            doc_vacio = df_sheet['Documento'].isna() | df_sheet['Documento'].astype(str).str.strip().str.upper().isin(['', 'NAN', 'NONE'])
            df_sheet = df_sheet[~(es_total & doc_vacio)].copy().reset_index(drop=True)

            col_haber = next((c for c in df_sheet.columns if 'Haber' in c), None)
            col_debe = next((c for c in df_sheet.columns if 'Debe' in c), None)
            col_cuenta = dict(zip(df_sheet.columns, df_sheet.columns)).get('Código de la cuenta', 
                            next((c for c in df_sheet.columns if 'cuenta' in c.lower()), None))
            col_etiqueta = dict(zip(df_sheet.columns, df_sheet.columns)).get('Etiqueta de la cuenta', 
                            next((c for c in df_sheet.columns if 'etiqueta' in c.lower()), None))

            if col_etiqueta and col_etiqueta in df_sheet.columns:
                df_sheet['Nombre_Cuenta'] = df_sheet[col_etiqueta].fillna('').astype(str).str.strip()
            else:
                df_sheet['Nombre_Cuenta'] = ''

            if not col_haber or not col_debe or not col_cuenta: continue

            df_sheet['Cod_Cuenta'] = df_sheet[col_cuenta].astype(str).str.replace('.0', '', regex=False).str.strip()
            
            # Definir qué cuentas inician una transacción (Cabeceras de Proveedores/Gastos por pagar)
            cuentas_cabecera = get_cuentas_total_compras()

            def is_new_transaction(row):
                doc = str(row['Documento']).strip().upper()
                cta = str(row['Cod_Cuenta']).strip()
                
                if doc in ['', 'NAN', 'NONE']: return False
                # Saltamos fechas que a veces Odoo pone en la columna Documento
                if doc.startswith('20') and '-' in doc and sum(c.isdigit() for c in doc) >= 4:
                    return False
                
                # Para la hoja DOC. SOPORTE, la presencia de un documento válido (no fecha, no vacío) inicia la transacción
                if sheet == 'DOC. SOPORTE':
                    return True
                
                # Una nueva transacción empieza con una cuenta de pasivo (proveedores) configurada como cuenta de total
                if any(cta.startswith(c) for c in cuentas_cabecera):
                    return True
                return False
            
            df_sheet['_IsHead'] = df_sheet.apply(is_new_transaction, axis=1)
            # Forzar la primera fila como head si no hay ninguna detectada para no perder datos iniciales
            if not df_sheet.empty and not df_sheet['_IsHead'].any():
                df_sheet.loc[0, '_IsHead'] = True
            
            df_sheet['_GroupID'] = df_sheet['_IsHead'].cumsum()

            # Propagar Referencia y Centro de costo de la cabecera a todas las líneas del grupo (por factura)
            if 'Centro de costo' in df_sheet.columns:
                df_sheet['_TempCC'] = np.where(df_sheet['_IsHead'], df_sheet['Centro de costo'], np.nan)
                df_sheet['Centro de costo'] = df_sheet.groupby('_GroupID')['_TempCC'].transform('first')
                df_sheet.drop(columns=['_TempCC'], inplace=True, errors='ignore')
                # Llenar nulos residuales
                df_sheet['Centro de costo'] = df_sheet['Centro de costo'].fillna('SIN CENTRO DE COSTOS')
                
            if 'Referencia' in df_sheet.columns:
                df_sheet['_TempRef'] = np.where(df_sheet['_IsHead'], df_sheet['Referencia'], np.nan)
                df_sheet['Referencia'] = df_sheet.groupby('_GroupID')['_TempRef'].transform('first')
                df_sheet.drop(columns=['_TempRef'], inplace=True, errors='ignore')

            def clean_money(val):
                if pd.isna(val): return 0.0
                if isinstance(val, (int, float)): return float(val)
                val_str = str(val).replace('$', '').replace('.', '').replace(',', '.').strip()
                try: return float(val_str)
                except: return 0.0

            df_sheet['Monto_Debe'] = df_sheet[col_debe].apply(clean_money)
            df_sheet['Monto_Haber'] = df_sheet[col_haber].apply(clean_money)

            def safe_row_join(row):
                return ' '.join([str(v) for v in row if pd.notna(v)])
            col_textos = [col_nombre]
            if col_etiqueta: col_textos.append(col_etiqueta)
            if 'Referencia' in df_sheet.columns and str(sheet).strip().upper() != 'DOC. SOPORTE':
                col_textos.append('Referencia')
            if 'Documento' in df_sheet.columns: col_textos.append('Documento')
            df_sheet['_FilaText'] = df_sheet[col_textos].apply(safe_row_join, axis=1)
            
            def safe_group_join(grp):
                try:
                    arr = grp.dropna().astype(str).unique() if hasattr(grp, 'unique') else grp.dropna().astype(str)
                    return ' '.join([str(v) for v in arr if pd.notna(v) and str(v).strip()]).upper()
                except: return ''
            df_sheet['Texto_Busqueda_Sis'] = df_sheet.groupby('_GroupID')['_FilaText'].transform(safe_group_join)

            # Extraer el Tercero de la línea de cabecera. 
            # Priorizamos 'Etiqueta de la cuenta' si existe, sino usamos 'Nombre'
            col_tercero_src = col_etiqueta if col_etiqueta else col_nombre
            df_sheet['Nombre_Original'] = df_sheet[col_nombre].astype(str).str.strip()
            df_sheet['_TempTercero'] = np.where(df_sheet['_IsHead'], df_sheet[col_tercero_src], np.nan)
            df_sheet['Tercero_Sistema'] = df_sheet.groupby('_GroupID')['_TempTercero'].transform('first')
            
            # También extraemos el Documento real de la cabecera para evitar que se use el número interno de las líneas de gasto
            df_sheet['_TempDoc'] = np.where(df_sheet['_IsHead'], df_sheet['Documento'], np.nan)
            df_sheet['Documento_Cabecera'] = df_sheet.groupby('_GroupID')['_TempDoc'].transform('first')

            df_sheet['Es_Devolucion_Fila'] = df_sheet['Nombre_Original'].str.upper().str.contains('REVERSIÓN|REVERSION', regex=True, na=False)
            df_sheet['Es_Devolucion'] = df_sheet.groupby('_GroupID')['Es_Devolucion_Fila'].transform('any')
            df_sheet['Hoja_Origen_Sistema'] = sheet
            lista_dfs_sistema.append(df_sheet)
            
        if not lista_dfs_sistema:
            raise Exception("No se encontraron hojas válidas con información de facturas en el Sistema.")
            
        # Eliminamos la escritura del archivo de origen para no destruir su formato.
        # El input file (ruta_sistema) no debe ser alterado.

        df_sistema = pd.concat(lista_dfs_sistema, ignore_index=True)
        # Generar identificador de bloque contable
        df_sistema['Identificador_Unico'] = df_sistema['Hoja_Origen_Sistema'] + "_" + df_sistema['_GroupID'].astype(str)
        
        # Agrupar IVA respetando la configuracion
        cuentas_iva = get_cuentas_iva_compras()
        df_iva_sis = df_sistema[df_sistema['Cod_Cuenta'].str.startswith(tuple(cuentas_iva), na=False)].copy()
        df_iva_sis['Monto_IVA'] = np.where(df_iva_sis['Es_Devolucion'], df_iva_sis['Monto_Haber'], df_iva_sis['Monto_Debe'])
        iva_por_grupo = df_iva_sis.groupby('Identificador_Unico')['Monto_IVA'].sum().reset_index()
        iva_por_grupo.rename(columns={'Monto_IVA': 'IVA_Sistema'}, inplace=True)
        
        # Agrupar Total respetando estrictamente las cuentas_total de la usuaria
        cuentas_total = get_cuentas_total_compras()
        df_total_sis = df_sistema[df_sistema['Cod_Cuenta'].str.startswith(tuple(cuentas_total), na=False)].copy()
        df_total_sis['Monto_Total_Conf'] = np.where(df_total_sis['Es_Devolucion'], df_total_sis['Monto_Debe'], df_total_sis['Monto_Haber'])
        total_por_grupo = df_total_sis.groupby('Identificador_Unico')['Monto_Total_Conf'].sum().reset_index()
        
        # Salvavidas: si por error usaron una cuenta no configurada (ej. 2335 en vez de 2205), sacar el total sumando el Debe
        total_debe = df_sistema.groupby('Identificador_Unico')['Monto_Debe'].sum().reset_index()
        total_por_grupo = pd.merge(total_debe, total_por_grupo, on='Identificador_Unico', how='left')
        total_por_grupo['Monto_Total_Conf'] = total_por_grupo['Monto_Total_Conf'].fillna(0)
        
        total_por_grupo['Total_Sistema'] = np.where(total_por_grupo['Monto_Total_Conf'] > 0, 
                                                    total_por_grupo['Monto_Total_Conf'], 
                                                    total_por_grupo['Monto_Debe'])
        total_por_grupo = total_por_grupo[['Identificador_Unico', 'Total_Sistema']]
        
        agg_dict = {
            'Tercero_Sistema': 'first',
            'Nombre_Original': 'first',
            'Hoja_Origen_Sistema': 'first',
            'Texto_Busqueda_Sis': 'first',
            'Documento_Cabecera': 'first'
        }
        if 'Referencia' in df_sistema.columns:
            agg_dict['Referencia'] = 'first'
        if 'Centro de costo' in df_sistema.columns:
            agg_dict['Centro de costo'] = 'first'

        df_agrup_extra = df_sistema.groupby('Identificador_Unico').agg(agg_dict).reset_index()
        df_agrup_extra.rename(columns={'Documento_Cabecera': 'Documento_Sistema_Original'}, inplace=True)
        
        sis_consolidado = pd.merge(df_agrup_extra, iva_por_grupo, on='Identificador_Unico', how='left')
        sis_consolidado = pd.merge(sis_consolidado, total_por_grupo, on='Identificador_Unico', how='left')
        
        # Llenar nulos de forma específica por tipo
        cols_num = ['IVA_Sistema', 'Total_Sistema']
        sis_consolidado[cols_num] = sis_consolidado[cols_num].fillna(0)
        sis_consolidado['Hoja_Origen_Sistema'] = sis_consolidado['Hoja_Origen_Sistema'].fillna('')
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise Exception(f"Error procesando reporte del sistema: {str(e)}")

    # 3. ------------- CONCILIAR (CRUCE) MASIVO (POR TEXTO NORMALIZADO) ---------
    try:
        import re
        resultados = []
        matched_sis = set()
        matched_dian = set()
        
        dian_dict = dian_agrupado.to_dict('index')
        sis_dict = sis_consolidado.to_dict('index')
        
        # Funcion que normaliza los strings extrayendo ABSOLUTAMENTE todo lo que no sea una letra o digito.
        # Esto soluciona de raiz el problema de pandas quitando los ceros a la izquierda (0001992 -> 1992)
        # o los guiones que agregan a mano en contabilidad (H-13 1992).
        def norm(s): return re.sub(r'[^A-Z1-9]', '', str(s).upper()) if pd.notna(s) else ''
        
        for k, v in dian_dict.items():
            pre_fol = str(v.get('PrefijoFolio_DIAN', '')).strip().upper()
            v['PrefijoFolio_DIAN_norm'] = norm(pre_fol)
            v['PrefijoFolio_DIAN_up'] = pre_fol
            
            fol = str(v.get('Folio_DIAN_Puro', '')).strip().upper()
            v['Folio_DIAN_Puro_norm'] = norm(fol)
            
            v['Total_DIAN_fl'] = float(v.get('Total_DIAN', 0))
            v['Nombre_Receptor_DIAN_up'] = str(v.get('Nombre_Receptor_DIAN', '')).strip().upper()
            
        for k, v in sis_dict.items():
            texto = str(v.get('Texto_Busqueda_Sis', '')).upper()
            v['Texto_Busqueda_Sis_up'] = texto
            v['Texto_Busqueda_Sis_norm'] = norm(texto)
            v['Total_Sistema_fl'] = float(v.get('Total_Sistema', 0))
            v['Tercero_Sistema_up'] = str(v.get('Tercero_Sistema', '')).upper()
        
        # 1. Match Preciso: Prefijo+Folio Normalizado
        for d_idx, d_row in dian_dict.items():
            prefijo_folio = d_row['PrefijoFolio_DIAN_norm']
            if not prefijo_folio or len(prefijo_folio) < 2: continue
            
            dian_total = d_row['Total_DIAN_fl']
            dian_receptor = d_row['Nombre_Receptor_DIAN_up']
            
            candidatos = []
            for s_idx, s_row in sis_dict.items():
                if s_idx in matched_sis: continue
                # Busqueda literal super eficiente, ignorando ceros, rayas y espacios de ambos lados.
                if prefijo_folio in s_row['Texto_Busqueda_Sis_norm']:
                    score_nombre = calcular_similitud_nombres(dian_receptor, s_row['Tercero_Sistema_up'])
                    diff_total = abs(s_row['Total_Sistema_fl'] - dian_total)
                    candidatos.append((s_idx, score_nombre, diff_total))
            
            if candidatos:
                candidatos.sort(key=lambda x: (-x[1], x[2]))
                mejor = candidatos[0]
                if mejor[1] > 0.6 or mejor[2] <= 100:
                    matched_sis.add(mejor[0])
                    matched_dian.add(d_idx)
                    resultados.append({**sis_dict[mejor[0]], **d_row, '_TipoCruce': 'Prefijo+Folio Coincidente'})
                elif len(candidatos) == 1:
                    matched_sis.add(mejor[0])
                    matched_dian.add(d_idx)
                    resultados.append({**sis_dict[mejor[0]], **d_row, '_TipoCruce': 'Prefijo+Folio Forzado'})

        # 2. Match por Solo Folio Normalizado
        for d_idx, d_row in dian_dict.items():
            if d_idx in matched_dian: continue
            
            folio_puro = d_row['Folio_DIAN_Puro_norm']
            if not folio_puro or len(folio_puro) < 3: 
                continue 
                
            dian_total = d_row['Total_DIAN_fl']
            dian_receptor = d_row['Nombre_Receptor_DIAN_up']
            
            candidatos = []
            for s_idx, s_row in sis_dict.items():
                if s_idx in matched_sis: continue
                if folio_puro in s_row['Texto_Busqueda_Sis_norm']:
                    score_nombre = calcular_similitud_nombres(dian_receptor, s_row['Tercero_Sistema_up'])
                    diff_total = abs(s_row['Total_Sistema_fl'] - dian_total)
                    if diff_total <= 50 or score_nombre > 0.8:
                        candidatos.append((s_idx, score_nombre, diff_total))
            
            if candidatos:
                candidatos.sort(key=lambda x: (-x[1], x[2]))
                mejor = candidatos[0]
                matched_sis.add(mejor[0])
                matched_dian.add(d_idx)
                resultados.append({**sis_dict[mejor[0]], **d_row, '_TipoCruce': 'Solo Folio Coincidente'})

        # 3. Match por Valor y Nombre
        for d_idx, d_row in dian_dict.items():
            if d_idx in matched_dian: continue
            dian_total = d_row['Total_DIAN_fl']
            dian_receptor = d_row['Nombre_Receptor_DIAN_up']
            
            candidatos = []
            for s_idx, s_row in sis_dict.items():
                if s_idx in matched_sis: continue
                diff_total = abs(s_row['Total_Sistema_fl'] - dian_total)
                if diff_total <= 10:
                    score = calcular_similitud_nombres(dian_receptor, s_row['Tercero_Sistema_up'])
                    if score > 0.75:
                        candidatos.append((s_idx, score, diff_total))
            
            if candidatos:
                candidatos.sort(key=lambda x: (-x[1], x[2]))
                mejor = candidatos[0]
                matched_sis.add(mejor[0])
                matched_dian.add(d_idx)
                
                s_row_mejor = sis_dict[mejor[0]]
                ref_sis = str(s_row_mejor.get('Referencia', '')).strip().upper()
                doc_sis = str(s_row_mejor.get('Documento_Sistema_Original', '')).strip().upper()
                pref_dian = str(d_row.get('PrefijoFolio_DIAN', '')).strip().upper()
                hoja_sis = str(s_row_mejor.get('Hoja_Origen_Sistema', '')).strip().upper()
                
                def son_parejas(r_sis, d_sis, p_dian, h_sis):
                    import re
                    def norm_str(s): return re.sub(r'[^A-Z0-9]', '', str(s).upper())
                    
                    np_dian = norm_str(p_dian)
                    if not np_dian: return False
                    
                    # 1. Si la pref_dian está contenida en el nombre de la hoja (como DOCUMENTO EQUIVALENTE A FACTUR.DEFL8980808)
                    if np_dian in norm_str(h_sis):
                        return True
                        
                    # 2. Si la referencia del sistema y la pref_dian son muy parecidas
                    nr_sis = norm_str(r_sis)
                    if nr_sis:
                        if nr_sis == np_dian: return True
                        if nr_sis in np_dian or np_dian in nr_sis: return True
                        # Similitud de Levenshtein/difflib
                        if difflib.SequenceMatcher(None, nr_sis, np_dian).ratio() >= 0.8:
                            return True
                            
                    # 3. Si el documento original del sistema y pref_dian son muy parecidos
                    nd_sis = norm_str(d_sis)
                    if nd_sis:
                        if nd_sis == np_dian: return True
                        if nd_sis in np_dian or np_dian in nd_sis: return True
                        if difflib.SequenceMatcher(None, nd_sis, np_dian).ratio() >= 0.8:
                            return True
                            
                    return False
                
                if son_parejas(ref_sis, doc_sis, pref_dian, hoja_sis):
                    resultados.append({**s_row_mejor, **d_row, '_TipoCruce': 'Cruce por Pareja de Factura'})
                else:
                    resultados.append({**s_row_mejor, **d_row, '_TipoCruce': 'Valor y Nombre (Sin Numero)'})

        # Insertar los sobrantes de forma clasificada
        for s_idx, s_row in sis_dict.items():
            if s_idx not in matched_sis:
                resultados.append({**s_row, '_TipoCruce': 'Solo en Sistema'})
                
        for d_idx, d_row in dian_dict.items():
            if d_idx not in matched_dian:
                resultados.append({**d_row, '_TipoCruce': 'Solo en DIAN'})
                
        cruce = pd.DataFrame(resultados)
        
        # Generar Numero_Factura correcto usando DIAN si la encontró, sino poner Documento de base del Sistema
        if 'Documento_Sistema_Original' not in cruce.columns: cruce['Documento_Sistema_Original'] = ''
        cruce['Nombre_Tercero'] = cruce['Tercero_Sistema'].fillna('')
        def refinar_numero_factura(row):
            tc = row.get('_TipoCruce', '')
            
            # Si no es un cruce con número coincidente, NO usamos el número de DIAN
            # (evita que parezcan idénticos en la hoja de Diferencias y Alertas)
            usar_dian = True
            if tc == 'Valor y Nombre (Sin Numero)':
                usar_dian = False
                
            # Si hay algo en DIAN y está permitido usarlo, eso es lo sagrado
            if usar_dian:
                pre_fol_dian = str(row.get('PrefijoFolio_DIAN', '')).strip().upper()
                if pre_fol_dian and pre_fol_dian not in ['0', '0.0', 'NAN', 'NONE']:
                    return pre_fol_dian
            
            # Si no hay DIAN, o no está permitido usarlo, usamos el del sistema pero con esteroides
            doc_sis = str(row.get('Documento_Sistema_Original', '')).strip().upper()
            if doc_sis in ['', 'NAN', 'NONE']: return ''
            
            # Si el documento del sistema parece un ID interno de Odoo (BILL/, INV/ o numero pequeño)
            # intentamos buscar la factura real
            internal_patterns = ['BILL/', 'INV/', 'FC/', 'NC/', 'ND/']
            is_internal = any(p in doc_sis for p in internal_patterns) or (doc_sis.isdigit() and len(doc_sis) < 7)
            
            if is_internal:
                hoja = str(row.get('Hoja_Origen_Sistema', '')).strip().upper()
                # Para DOC SOPORTE, usamos el documento original como solicitó el usuario
                if hoja == 'DOC. SOPORTE':
                    return doc_sis
                else:
                    # Para las demás compras, extraer de la nueva columna 'Referencia'
                    ref = str(row.get('Referencia', '')).strip().upper()
                    if ref and ref not in ['NAN', 'NONE', '0', '0.0']:
                        return ref
            
            return doc_sis

        cruce['Numero_Factura'] = cruce.apply(refinar_numero_factura, axis=1)
        
        # Limpiar valores NA originados por combinaciones a lo bruto y normalizar Centro de costo
        for c in cruce.columns:
            if cruce[c].dtype == 'object':
                cruce[c] = cruce[c].fillna('')
            else:
                cruce[c] = cruce[c].fillna(0)
                
        if 'Centro de costo' in cruce.columns:
            cruce['Centro de costo'] = cruce['Centro de costo'].fillna('').astype(str).str.strip()
            cruce['Centro de costo'] = np.where(cruce['Centro de costo'].isin(['', 'nan', 'NaN', 'None', '0', '0.0', 0, 0.0]), 'SIN CENTRO DE COSTOS', cruce['Centro de costo'])
        
        cruce['IVA_Sistema'] = cruce['IVA_Sistema'].round(0)
        cruce['IVA_DIAN'] = cruce['IVA_DIAN'].round(0)
        cruce['Total_Sistema'] = cruce['Total_Sistema'].round(0)
        cruce['Total_DIAN'] = cruce['Total_DIAN'].round(0)
        cruce['Dif_IVA'] = cruce['IVA_Sistema'] - cruce['IVA_DIAN']
        cruce['Dif_Total'] = cruce['Total_Sistema'] - cruce['Total_DIAN']
        
        def categorizar(row):
            tc = row.get('_TipoCruce', '')
            if tc == 'Prefijo+Folio Forzado':
                return 'Validar: Extrapolación de Nro / Proveedor Distinto'
            elif tc == 'Valor y Nombre (Sin Numero)':
                return 'Validar: Cruce Mismo Valor/Nombre (Sin Nro Factura)'
            
            if tc == 'Solo en Sistema': return 'Solo en Sistema'
            elif tc == 'Solo en DIAN': return 'Solo en DIAN'
            else:
                if abs(row['Dif_IVA']) > 10 or abs(row['Dif_Total']) > 10: return 'Diferencia en Valores'
                else: return 'Cuadre Perfecto'
                    
        cruce['Estado'] = cruce.apply(categorizar, axis=1)
        
        # Limpiar cualquier residuo de tipos mixtos en columnas de texto
        columnas_texto = ['Numero_Factura', 'PrefijoFolio_DIAN', 'Hoja_Origen_Sistema', 'Nombre_Tercero', 'Nombre_Receptor_DIAN']
        for col in columnas_texto:
            if col in cruce.columns:
                cruce[col] = cruce[col].astype(str).replace(['0', '0.0', 'nan', 'NaN', 'None'], '').str.strip()

        # Generar DataFrame de Resumen por Cuenta y CC
        df_res_cuentas = pd.DataFrame()
        if not df_sistema.empty:
            # Asegurar normalización final de CC
            df_sistema['Centro de costo'] = df_sistema['Centro de costo'].fillna('').astype(str).str.strip()
            df_sistema['Centro de costo'] = np.where(df_sistema['Centro de costo'].isin(['', 'nan', 'NaN', 'None', '0', '0.0', 0, 0.0]), 'SIN CENTRO DE COSTOS', df_sistema['Centro de costo'])
            
            df_res_cuentas = df_sistema.groupby(['Cod_Cuenta', 'Nombre_Cuenta', 'Centro de costo']).agg({
                'Monto_Debe': 'sum',
                'Monto_Haber': 'sum'
            }).reset_index()
            df_res_cuentas['Neto'] = df_res_cuentas['Monto_Debe'] - df_res_cuentas['Monto_Haber']
            df_res_cuentas.rename(columns={
                'Cod_Cuenta': 'Código de la Cuenta',
                'Nombre_Cuenta': 'Etiqueta de la Cuenta',
                'Centro de costo': 'Centro de Costos',
                'Monto_Debe': 'Total Debe',
                'Monto_Haber': 'Total Haber',
                'Neto': 'Saldo Neto'
            }, inplace=True)
            for col in ['Total Debe', 'Total Haber', 'Saldo Neto']:
                df_res_cuentas[col] = df_res_cuentas[col].round(0)

        # Generar DataFrame de Resumen por Hoja de Compras y CC
        df_res_hojas = pd.DataFrame()
        if not df_sistema.empty:
            df_res_hojas = df_sistema.groupby(['Hoja_Origen_Sistema', 'Centro de costo']).agg({
                'Monto_Debe': 'sum',
                'Monto_Haber': 'sum'
            }).reset_index()
            df_res_hojas['Neto'] = df_res_hojas['Monto_Debe'] - df_res_hojas['Monto_Haber']
            df_res_hojas.rename(columns={
                'Hoja_Origen_Sistema': 'Hoja de Compras',
                'Centro de costo': 'Centro de Costos',
                'Monto_Debe': 'Total Debe',
                'Monto_Haber': 'Total Haber',
                'Neto': 'Saldo Neto'
            }, inplace=True)
            for col in ['Total Debe', 'Total Haber', 'Saldo Neto']:
                df_res_hojas[col] = df_res_hojas[col].round(0)

        # 4. ------------- GENERAR EXCEL -------------
        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            df_diff = cruce[cruce['Estado'].str.contains('Diferencia|Validar', na=False)]
            df_solo_sis = cruce[cruce['Estado'] == 'Solo en Sistema']
            df_solo_dian = cruce[cruce['Estado'] == 'Solo en DIAN']
            df_cuadre = cruce[cruce['Estado'].str.contains('Cuadre Perfecto', na=False)]
            
            cols_export = ['Numero_Factura', 'Nombre_Tercero', 'Nombre_Receptor_DIAN', 
                           'Hoja_Origen_Sistema', 'PrefijoFolio_DIAN', 
                           'IVA_Sistema', 'IVA_DIAN', 'Dif_IVA', 
                           'Total_Sistema', 'Total_DIAN', 'Dif_Total', 'Estado']
            
            if 'Referencia' in cruce.columns:
                cols_export.insert(1, 'Referencia')
            if 'Centro de costo' in cruce.columns:
                cols_export.insert(2, 'Centro de costo')
            
            hojas_a_crear = [
                ('Diferencias y Alertas', df_diff), 
                ('Solo en DIAN', df_solo_dian), 
                ('Solo en Sistema', df_solo_sis), 
                ('Cuadre Perfecto', df_cuadre), 
                ('Resumen Total Cruce', cruce)
            ]
            
            from openpyxl.styles import PatternFill, Font, Alignment # type: ignore
            header_fill = PatternFill(start_color="1B263B", end_color="1B263B", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True, size=11)
            alternating_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
            alert_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid") # Amarillo suave
 
            for nombre_hoja, dataframe in hojas_a_crear:
                if not dataframe.empty or nombre_hoja in ['Resumen Total Cruce']:
                    cols_to_write = list(cols_export)
                    
                    if nombre_hoja == 'Solo en DIAN' and 'Tipo de documento' in dataframe.columns:
                        cols_to_write.insert(0, 'Tipo de documento')
                        
                    cols_to_write = [c for c in cols_to_write if c in dataframe.columns]
                    
                    dataframe[cols_to_write].to_excel(writer, sheet_name=nombre_hoja, index=False)
                    worksheet = writer.sheets[nombre_hoja] # type: ignore
                    worksheet.sheet_view.showGridLines = False
                    
                    for cell in worksheet[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    
                    worksheet.auto_filter.ref = worksheet.dimensions
                    
                    # Determinar dinámicamente las columnas de moneda
                    currency_cols = []
                    for i, col_name in enumerate(cols_to_write, 1):
                        if any(term in col_name for term in ['IVA', 'Total', 'Dif_Total', 'Dif_IVA', 'Debe', 'Haber', 'Neto', 'Saldo Neto']):
                            currency_cols.append(i)
                            
                    for row_idx, row_cells in enumerate(worksheet.iter_rows(min_row=2), 2):
                        # Zebra stripes
                        if row_idx % 2 == 0:
                            for cell in row_cells: cell.fill = alternating_fill
                        
                        # Resaltar alertas
                        if 'Estado' in cols_to_write:
                            estado_val = str(row_cells[-1].value)
                            if 'Validar' in estado_val or 'Diferencia' in estado_val:
                                for cell in row_cells: cell.fill = alert_fill
 
                    for col_idx, col in enumerate(worksheet.columns, 1):
                        max_length = len(str(col[0].value)) + 4
                        column_letter = col[0].column_letter
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                            except: pass
                            # Formato moneda
                            if col_idx in currency_cols and cell.row > 1: cell.number_format = '#,##0'
                        worksheet.column_dimensions[column_letter].width = min(max_length + 3, 50)
                        
                    for row_idx in range(1, worksheet.max_row + 1):
                        worksheet.row_dimensions[row_idx].height = 22
            
            # Generar los resúmenes en cuadrícula premium (estilo Dashboard) tal como lo solicitó el usuario
            escribir_tablas_resumen_grid(writer, df_sistema, 'Centro de costo', 'Resumen por Cuenta y CC', 'CENTRO DE COSTO: ')
            escribir_tablas_resumen_grid(writer, df_sistema, 'Hoja_Origen_Sistema', 'Resumen por Hoja de Compras', 'HOJA COMPRAS: ')
            
        stats_diferencias = len(cruce[cruce['Estado'].str.contains('Diferencia|Validar', na=False) & ~cruce['Estado'].str.contains('Manual', na=False)])
        return {"total_procesadas": len(cruce), "cuadre": len(df_cuadre), "diferencias": stats_diferencias, "solo_sis": len(df_solo_sis), "solo_dian": len(df_solo_dian)}
    except Exception as e:
        raise Exception(f"Error generando reporte y cruce de datos: {str(e)}")
