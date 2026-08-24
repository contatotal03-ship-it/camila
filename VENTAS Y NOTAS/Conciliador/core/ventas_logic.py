import pandas as pd # type: ignore
import numpy as np # type: ignore
import datetime
import os
from .config_manager import get_hojas_ventas, get_dian_ventas_excluir, get_cuentas_iva_ventas, get_cuentas_excluir_ventas, get_dian_grupo_ventas, get_dian_tipos_validos_ventas, get_cuentas_transporte_ventas # type: ignore
from .excel_utils import robust_read_excel, robust_ExcelFile # type: ignore

def procesar_conciliacion_ventas(ruta_sistema, ruta_dian, ruta_salida, ruta_dian_detalle=None):
    """
    Lee los archivos del Sistema y de la DIAN, aplica las reglas de negocio,
    cruza la información y genera un archivo de Excel con el resultado.
    """
    
    # 1. ------------- PROCESAR DIAN (PRINCIPAL) -------------
    try:
        df_dian = robust_read_excel(ruta_dian)
        cols_dian = {col.lower().strip(): col for col in df_dian.columns}
        col_pre = next((v for k, v in cols_dian.items() if k in ['prefijo', 'pre', 'pref']), None)
        col_fol = next((v for k, v in cols_dian.items() if k in ['folio', 'fol']), None)
        col_iva = next((v for k, v in cols_dian.items() if k == 'iva'), None)
        col_total = next((v for k, v in cols_dian.items() if k == 'total'), None)
        col_grupo = next((v for k, v in cols_dian.items() if k == 'grupo'), None)
        col_tipo = next((c for c in df_dian.columns if 'tipo' in c.lower() and 'doc' in c.lower()), 
                        next((c for c in df_dian.columns if 'tipo' in c.lower()), None))
        col_nombre = next((v for k, v in cols_dian.items() if 'receptor' in k or 'nombre' in k), None)

        if not col_fol or not col_total:
            raise Exception("No se encontraron Folio o Total en el archivo DIAN.")

        if col_grupo:
            grupo_ventas = str(get_dian_grupo_ventas()).strip().lower()
            mask_grupo = df_dian[col_grupo].astype(str).str.strip().str.lower() == grupo_ventas
            if col_tipo:
                terminos_excluir = get_dian_ventas_excluir()
                pattern = '|'.join(terminos_excluir) if terminos_excluir else 'XXXXX'
                df_dian = df_dian[mask_grupo & ~df_dian[col_tipo].astype(str).str.contains(pattern, case=False, na=False)].copy()
            else:
                df_dian = df_dian[mask_grupo].copy()
            
        p_part = df_dian[col_pre].astype(str).replace(['nan', 'NaN', 'None', '0', '0.0'], '').str.replace(r'\.0$', '', regex=True).str.strip() if col_pre else ''
        f_part = df_dian[col_fol].astype(str).replace(['nan', 'NaN', 'None'], '').str.replace(r'\.0$', '', regex=True).str.strip()
        df_dian['Numero_Factura'] = p_part + f_part
        
        tv = get_dian_tipos_validos_ventas()
        if tv and col_tipo:
            tv_l = [t.lower().strip() for t in tv]
            df_dian = df_dian[df_dian[col_tipo].astype(str).str.strip().str.lower().apply(lambda x: any(v in x or x in v for v in tv_l))].copy()
        
        df_dian['IVA_DIAN'] = pd.to_numeric(df_dian[col_iva], errors='coerce').fillna(0)
        df_dian['Total_DIAN'] = pd.to_numeric(df_dian[col_total], errors='coerce').fillna(0)
        df_dian['Nombre_Receptor_DIAN'] = df_dian[col_nombre].fillna('').astype(str).str.strip() if col_nombre else ''
        df_dian['Prefijo_Dian'] = p_part
        
        dian_agrupado = df_dian.groupby('Numero_Factura').agg({
            'IVA_DIAN': 'sum', 'Total_DIAN': 'sum', 'Nombre_Receptor_DIAN': 'first', 'Prefijo_Dian': 'first'
        }).reset_index()
    except Exception as e:
        raise Exception(f"Error DIAN: {str(e)}")

    # --- PROCESAR DIAN DETALLE ---
    dian_transporte_agrupado = pd.DataFrame(columns=['Numero_Factura', 'Transporte_DIAN'])
    if ruta_dian_detalle:
        try:
            xls_det = robust_ExcelFile(ruta_dian_detalle)
            lista_det = []
            for sheet in [s for s in xls_det.sheet_names if 'DET_' in s.upper()]:
                df_det = robust_read_excel(xls_det, sheet_name=sheet)
                df_det.columns = [str(c).strip() for c in df_det.columns]
                c_f = next((c for c in df_det.columns if 'FOLIO' in c.upper()), None)
                c_p = next((c for c in df_det.columns if 'PREFIJO' in c.upper()), None)
                c_d = next((c for c in df_det.columns if 'DESCRIP' in c.upper()), None)
                c_v = next((c for c in df_det.columns if 'SUBTOTAL' in c.upper() and 'SIN' in c.upper()), 
                           next((c for c in df_det.columns if 'TOTAL' in c.upper() and 'LINEA' in c.upper()), None))
                if c_f and c_d and c_v:
                    pp = df_det[c_p].astype(str).replace(['nan', 'NaN', 'None', '0', '0.0'], '').str.replace(r'\.0$', '', regex=True).str.strip() if c_p else ''
                    ff = df_det[c_f].astype(str).replace(['nan', 'NaN', 'None'], '').str.replace(r'\.0$', '', regex=True).str.strip()
                    df_det['Num_F'] = pp + ff
                    mask_tr = df_det[c_d].astype(str).str.upper().str.contains('TRANSPORTE|TRASPORTE', na=False, regex=True)
                    df_det['Val_T'] = np.where(mask_tr, pd.to_numeric(df_det[c_v], errors='coerce').fillna(0), 0)
                    lista_det.append(df_det[['Num_F', 'Val_T']])
            if lista_det:
                dian_transporte_agrupado = pd.concat(lista_det, ignore_index=True).groupby('Num_F')['Val_T'].sum().reset_index()
                dian_transporte_agrupado.columns = ['Numero_Factura', 'Transporte_DIAN']
        except Exception: pass

    # 2. ------------- PROCESAR SISTEMA ----------
    try:
        xls_sis = robust_ExcelFile(ruta_sistema)
        lista_sis = []
        hojas_v = [h.strip().upper() for h in get_hojas_ventas()]
        for sheet in xls_sis.sheet_names:
            if sheet.strip().upper() not in hojas_v: continue
            df_s = robust_read_excel(xls_sis, sheet_name=sheet)
            df_s.columns = [str(c).strip() for c in df_s.columns]
            if 'Documento' not in df_s.columns: continue
            
            def is_head(v):
                s = str(v).strip().upper()
                if s in ['', 'NAN', 'NONE']: return False
                
                import re
                # Excluir SOLAMENTE si es una fecha pura (ej: 2026-01-29 o 29/01/2026)
                if re.match(r'^\d{2,4}[-/]\d{1,2}[-/]\d{2,4}$', s): return False
                
                # Aumentar el límite de longitud para permitir nombres de terceros que actúan como cabecera
                if len(s) > 60: return False
                
                # Una cabecera es válida si tiene letras o números
                return any(c.isalnum() for c in s)


            df_s['_IsHead'] = df_s['Documento'].apply(is_head)
            
            # Nueva lógica: Identificar cabeceras reales
            df_s['_IsRealHead'] = False
            last_was_empty = True
            for idx in df_s.index:
                val = str(df_s.at[idx, 'Documento']).strip().upper()
                if val in ['', 'NAN', 'NONE']:
                    last_was_empty = True
                    continue
                
                if is_head(val):
                    # Es cabecera real si tiene letras O si viene tras un espacio vacío
                    if any(c.isalpha() for c in val) or last_was_empty:
                        df_s.at[idx, '_IsRealHead'] = True
                last_was_empty = False

            df_s['_GroupID'] = df_s['_IsRealHead'].cumsum()
            # Aseguramos que Num_Factura_Sis se propague a TODAS las filas del grupo
            def get_group_name(group):
                first_val = str(group.iloc[0]).strip().upper()
                return first_val if is_head(first_val) else np.nan

            df_s['Num_Factura_Sis'] = df_s.groupby('_GroupID')['Documento'].transform(get_group_name)

            
            def clean_m(v):
                if pd.isna(v): return 0.0
                if isinstance(v, (int, float)): return float(v)
                s = str(v).replace('$', '').strip()
                # Eliminar caracteres invisibles como \xa0
                s = "".join(c for c in s if c.isdigit() or c in [',', '.', '-'])
                if not s: return 0.0
                try:
                    # Formato estándar: punto mil, coma decimal -> 1.234,56
                    if ',' in s and '.' in s:
                        if s.rfind('.') < s.rfind(','): # El punto está antes que la coma
                            s = s.replace('.', '').replace(',', '.')
                        else: # La coma está antes que el punto
                            s = s.replace(',', '')
                    elif ',' in s: # Solo coma
                        s = s.replace(',', '.')
                    return float(s)
                except: return 0.0
            
            cH, cD = next((c for c in df_s.columns if 'Haber' in c), None), next((c for c in df_s.columns if 'Debe' in c), None)
            cC, cE = next((c for c in df_s.columns if 'cuenta' in c.lower()), None), next((c for c in df_s.columns if 'etiqueta' in c.lower()), None)
            if not cH or not cD or not cC: continue
            
            df_s['M_Debe'], df_s['M_Haber'] = df_s[cD].apply(clean_m), df_s[cH].apply(clean_m)
            # Limpiar código de cuenta de puntos y decimales
            df_s['Cod_C'] = df_s[cC].astype(str).str.replace(r'\.0$', '', regex=True).str.replace('.', '', regex=False).str.strip()

            
            # Cuentas Críticas
            cIVA = tuple(get_cuentas_iva_ventas()) # Ej: '2408'
            cTR = get_cuentas_transporte_ventas() # Ej: ['413572']
            
            # Identificar NC
            df_s['Es_NC'] = df_s['Num_Factura_Sis'].str.startswith(('NCC', 'NC'), na=False)
            
            # Valores por línea
            # Para Facturas: Sumamos Haber. Para NC: Sumamos Debe.
            df_s['M_Cont'] = np.where(df_s['Es_NC'], df_s['M_Debe'], df_s['M_Haber'])
            
            # TOTAL SISTEMA: Solo sumamos cuentas de Ingresos (4) e IVA (24)
            # Esto evita sumar Inventarios (14) o Costos (6) que inflan el total.
            mask_total = df_s['Cod_C'].str.startswith(('4', '24'))
            df_s['Val_Total_Sis'] = np.where(mask_total, df_s['M_Cont'], 0)
            
            # IVA SISTEMA: Solo cuentas de IVA
            mask_iva = df_s['Cod_C'].str.startswith(cIVA)
            df_s['Val_IVA_Sis'] = np.where(mask_iva, df_s['M_Cont'], 0)
            
            # TRANSPORTE SISTEMA: Cuentas específicas o etiqueta
            mask_tr = df_s['Cod_C'].isin(cTR) | df_s[cE].astype(str).str.upper().str.contains('TRANSPORTE', na=False)
            df_s['Val_Tr_Sis'] = np.where(mask_tr, df_s['M_Cont'], 0)
            
            df_s['Tercero_S'] = df_s.groupby('_GroupID')[cE].transform('first')
            df_s['Hoja_S'] = sheet
            lista_sis.append(df_s)
            
        df_all = pd.concat(lista_sis, ignore_index=True).dropna(subset=['Num_Factura_Sis'])
        
        # Agrupar por factura
        sis_con = df_all.groupby('Num_Factura_Sis').agg({
            'Val_Total_Sis': 'sum',
            'Val_IVA_Sis': 'sum',
            'Val_Tr_Sis': 'sum',
            'Tercero_S': 'first',
            'Hoja_S': 'first'
        }).reset_index()
        sis_con.columns = ['Numero_Factura', 'Total_Sistema', 'IVA_Sistema', 'Transporte_Sistema', 'Tercero_Sistema', 'Hoja_Origen_Sistema']
        
    except Exception as e: raise Exception(f"Error Sistema: {str(e)}")

    # 3. ------------- CRUCE ---------
    # Limpieza para el cruce: convertir a string y quitar espacios
    sis_con['Match_Key'] = sis_con['Numero_Factura'].astype(str).str.strip().str.upper()
    dian_agrupado['Match_Key'] = dian_agrupado['Numero_Factura'].astype(str).str.strip().str.upper()
    
    # Intentar cruce exacto primero
    cruce = pd.merge(sis_con, dian_agrupado, on='Match_Key', how='outer', suffixes=('_Sis', '_Dian')).fillna(0)
    
    # Para las que no cruzaron, intentar cruce por "Solo Números"
    def extraer_numeros(s):
        import re
        nums = re.findall(r'\d+', str(s))
        return "".join(nums).lstrip('0') if nums else s

    solo_sis = cruce[cruce['Numero_Factura_Dian'] == 0].copy()
    solo_dian = cruce[cruce['Numero_Factura_Sis'] == 0].copy()
    cruce_exito = cruce[(cruce['Numero_Factura_Dian'] != 0) & (cruce['Numero_Factura_Sis'] != 0)].copy()
    
    if not solo_sis.empty and not solo_dian.empty:
        solo_sis['Num_Only'] = solo_sis['Match_Key'].apply(extraer_numeros)
        solo_dian['Num_Only'] = solo_dian['Match_Key'].apply(extraer_numeros)
        
        # Filtramos para evitar cruces falsos con números vacíos o muy cortos
        solo_sis_valid = solo_sis[solo_sis['Num_Only'].str.len() > 1]
        solo_dian_valid = solo_dian[solo_dian['Num_Only'].str.len() > 1]
        
        re_cruce = pd.merge(
            solo_sis_valid[['Num_Only', 'Total_Sistema', 'IVA_Sistema', 'Transporte_Sistema', 'Tercero_Sistema', 'Hoja_Origen_Sistema', 'Numero_Factura_Sis']], 
            solo_dian_valid[['Num_Only', 'IVA_DIAN', 'Total_DIAN', 'Nombre_Receptor_DIAN', 'Numero_Factura_Dian', 'Prefijo_Dian']], 
            on='Num_Only', how='inner'
        )
        
        if not re_cruce.empty:
            # Eliminar las que re-cruzaron de las listas de solo uno
            sis_recruzados = re_cruce['Numero_Factura_Sis'].values
            dian_recruzados = re_cruce['Numero_Factura_Dian'].values
            solo_sis = solo_sis[~solo_sis['Numero_Factura_Sis'].isin(sis_recruzados)]
            solo_dian = solo_dian[~solo_dian['Numero_Factura_Dian'].isin(dian_recruzados)]
            # Añadir las re-cruzadas al éxito
            re_cruce['Match_Key'] = re_cruce['Numero_Factura_Sis']
            cruce_exito = pd.concat([cruce_exito, re_cruce], ignore_index=True)
            
    cruce = pd.concat([cruce_exito, solo_sis, solo_dian], ignore_index=True).fillna(0)
    
    # Asegurar que existan todas las columnas necesarias
    for col in ['Transporte_DIAN', 'Transporte_Sistema', 'IVA_Sistema', 'IVA_DIAN', 'Total_Sistema', 'Total_DIAN', 'Numero_Factura_Sis', 'Numero_Factura_Dian']:
        if col not in cruce.columns:
            cruce[col] = 0
            
    if 'Prefijo_Dian' not in cruce.columns:
        cruce['Prefijo_Dian'] = ''
    
    # Restaurar Numero_Factura (preferencia al del Sistema si existe)
    cruce['Numero_Factura'] = np.where(cruce['Numero_Factura_Sis'] != 0, cruce['Numero_Factura_Sis'], cruce['Numero_Factura_Dian'])
    
    # Procesar transporte DIAN si existe
    if not dian_transporte_agrupado.empty:
        dian_transporte_agrupado['Match_Key'] = dian_transporte_agrupado['Numero_Factura'].astype(str).str.strip().str.upper()
        cruce = pd.merge(cruce, dian_transporte_agrupado[['Match_Key', 'Transporte_DIAN']], on='Match_Key', how='outer', suffixes=('', '_new')).fillna(0)
        
        if 'Transporte_DIAN_new' in cruce.columns:
            cruce['Transporte_DIAN'] = np.where(cruce['Transporte_DIAN_new'] != 0, cruce['Transporte_DIAN_new'], cruce['Transporte_DIAN'])
            cruce = cruce.drop(columns=['Transporte_DIAN_new'])
            
        # Restaurar Número de Factura para los que entraron solo por transporte DIAN
        cruce['Numero_Factura'] = np.where(cruce['Numero_Factura'] == 0, cruce['Match_Key'], cruce['Numero_Factura'])
        # Asegurar que se vean como Solo en DIAN
        cruce['Hoja_Origen_Sistema'] = np.where(cruce['Hoja_Origen_Sistema'] == 0, '0', cruce['Hoja_Origen_Sistema'])
        cruce['Numero_Factura_Dian'] = np.where(cruce['Numero_Factura_Dian'] == 0, cruce['Match_Key'], cruce['Numero_Factura_Dian'])

    for c in ['IVA_Sistema', 'IVA_DIAN', 'Total_Sistema', 'Total_DIAN', 'Transporte_Sistema', 'Transporte_DIAN']: 
        cruce[c] = pd.to_numeric(cruce[c], errors='coerce').fillna(0).round(0)
        
    cruce['Dif_IVA'], cruce['Dif_Total'] = cruce['IVA_Sistema'] - cruce['IVA_DIAN'], cruce['Total_Sistema'] - cruce['Total_DIAN']
    
    def categorizar(r):
        if r['Numero_Factura'] not in sis_con['Numero_Factura'].values: return 'Solo en DIAN'
        if r['Numero_Factura'] not in dian_agrupado['Numero_Factura'].values: return 'Solo en Sistema'
        return 'Diferencia en Valores' if abs(r['Dif_IVA']) > 10 or abs(r['Dif_Total']) > 10 else 'Cuadre Perfecto'
    cruce['Estado'] = cruce.apply(categorizar, axis=1)

    def asignar_cat(r):
        if str(r['Numero_Factura']).upper().startswith(('NC', 'NCC')): return 'NOTAS CREDITO'
        iva = abs(r['IVA_Sistema']) > 1 or abs(r['IVA_DIAN']) > 1
        if 'REPUESTOS' in str(r['Hoja_Origen_Sistema']).upper(): return 'REPUESTOS GRAVADOS' if iva else 'REPUESTOS EXENTOS'
        return 'GRAVADAS' if iva else 'EXENTAS'
    cruce['Categoria_Resumen'] = cruce.apply(asignar_cat, axis=1)

    # Base = Total - IVA
    cruce['Base_Sistema'] = cruce['Total_Sistema'] - cruce['IVA_Sistema']
    cruce['Base_DIAN'] = cruce['Total_DIAN'] - cruce['IVA_DIAN']
    def get_hoja_resumen(r):
        hoja_sis = str(r.get('Hoja_Origen_Sistema', '')).strip().upper()
        if hoja_sis in ['0', '0.0', 'NAN', 'NONE', '']:
            prefijo = str(r.get('Prefijo_Dian', '')).strip().upper()
            if prefijo and prefijo not in ['0', '0.0', 'NAN', 'NONE']:
                return f"SOLO EN DIAN - {prefijo}"
            return "SOLO EN DIAN"
        return hoja_sis
        
    cruce['Hoja_Resumen'] = cruce.apply(get_hoja_resumen, axis=1)
    
    # Resumen General
    res_data = []
    # Ordenar las hojas. Ponemos "SOLO EN DIAN" al final
    hojas_todas = cruce['Hoja_Resumen'].unique()
    hojas_sis = sorted([h for h in hojas_todas if not h.startswith('SOLO EN DIAN')])
    hojas_dian = sorted([h for h in hojas_todas if h.startswith('SOLO EN DIAN')])
    hojas_u = hojas_sis + hojas_dian
    
    for h in hojas_u:
        df_h = cruce[cruce['Hoja_Resumen'] == h]
        for c in ['EXENTAS', 'GRAVADAS', 'REPUESTOS EXENTOS', 'REPUESTOS GRAVADOS', 'NOTAS CREDITO']:
            df_c = df_h[df_h['Categoria_Resumen'] == c]
            if df_c.empty: continue
            res_data.append({
                'SEDE / HOJA': h, 
                'CONCEPTO': c, 
                'TOTAL SISTEMA': df_c['Total_Sistema'].sum(), 
                'IVA SISTEMA': df_c['IVA_Sistema'].sum(), 
                'TOTAL DIAN': df_c['Total_DIAN'].sum(), 
                'IVA DIAN': df_c['IVA_DIAN'].sum(), 
                'DIF. TOTAL': df_c['Total_Sistema'].sum() - df_c['Total_DIAN'].sum(), 
                'DIF. IVA': df_c['IVA_Sistema'].sum() - df_c['IVA_DIAN'].sum()
            })
            
        res_data.append({
            'SEDE / HOJA': f'TOTAL {h}', 
            'CONCEPTO': 'TOTAL', 
            'TOTAL SISTEMA': df_h['Total_Sistema'].sum(), 
            'IVA SISTEMA': df_h['IVA_Sistema'].sum(), 
            'TOTAL DIAN': df_h['Total_DIAN'].sum(), 
            'IVA DIAN': df_h['IVA_DIAN'].sum(), 
            'DIF. TOTAL': df_h['Total_Sistema'].sum() - df_h['Total_DIAN'].sum(), 
            'DIF. IVA': df_h['IVA_Sistema'].sum() - df_h['IVA_DIAN'].sum()
        })
        # Fila en blanco para legibilidad
        res_data.append({k: '' for k in ['SEDE / HOJA', 'CONCEPTO', 'TOTAL SISTEMA', 'IVA SISTEMA', 'TOTAL DIAN', 'IVA DIAN', 'DIF. TOTAL', 'DIF. IVA']})

    # Cuadro Consolidado de Transporte
    res_tr_data = []
    for h in hojas_u:
        df_h = cruce[cruce['Hoja_Resumen'] == h]
        s_tr, d_tr = df_h['Transporte_Sistema'].sum(), df_h['Transporte_DIAN'].sum()
        if abs(s_tr) > 1 or abs(d_tr) > 1:
            res_tr_data.append({'SEDE / HOJA': h, 'TRANSPORTE SISTEMA': s_tr, 'TRANSPORTE DIAN': d_tr, 'DIFERENCIA': s_tr - d_tr})
    df_tr_summary = pd.DataFrame(res_tr_data)
    
    # Detalle Factura por Factura de Transporte
    mask_tr_detalle = (abs(cruce['Transporte_Sistema']) > 0) | (abs(cruce['Transporte_DIAN']) > 0)
    df_tr_detalle = cruce[mask_tr_detalle].copy()
    
    df_tr_detalle['Tipo_Documento'] = np.where(
        df_tr_detalle['Categoria_Resumen'] == 'NOTAS CREDITO', 
        'Devolución / Nota Crédito', 
        'Factura'
    )
    df_tr_detalle['Diferencia_Transporte'] = df_tr_detalle['Transporte_Sistema'] - df_tr_detalle['Transporte_DIAN']
    
    columnas_detalle_tr = [
        'Numero_Factura', 'Tipo_Documento', 'Hoja_Resumen', 
        'Transporte_Sistema', 'Transporte_DIAN', 'Diferencia_Transporte'
    ]
    if not df_tr_detalle.empty:
        df_tr_out = df_tr_detalle[columnas_detalle_tr].sort_values(by=['Tipo_Documento', 'Hoja_Resumen', 'Numero_Factura'])
        df_tr_out.rename(columns={
            'Numero_Factura': 'NÚMERO FACTURA',
            'Tipo_Documento': 'TIPO DOCUMENTO',
            'Hoja_Resumen': 'SEDE / HOJA',
            'Transporte_Sistema': 'TRANSPORTE SISTEMA',
            'Transporte_DIAN': 'TRANSPORTE DIAN',
            'Diferencia_Transporte': 'DIFERENCIA'
        }, inplace=True)
    else:
        df_tr_out = pd.DataFrame()

    # 4. ------------- EXCEL -------------
    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
        pd.DataFrame(res_data).to_excel(writer, sheet_name='Resumen General', startrow=1, index=False)
        start_tr = len(res_data) + 5
        if not df_tr_summary.empty:
            df_tr_summary.to_excel(writer, sheet_name='Resumen General', startrow=start_tr, index=False)
            
        # Poner los títulos de los cuadros
        ws_resumen = writer.sheets['Resumen General']
        ws_resumen['A1'] = "RESUMEN CONSOLIDADO DE VENTAS E IMPUESTOS"
        if not df_tr_summary.empty:
            ws_resumen.cell(row=start_tr, column=1).value = "CUADRO CONSOLIDADO DE TRANSPORTE"
            
        cols = ['Numero_Factura', 'Tercero_Sistema', 'Nombre_Receptor_DIAN', 'Hoja_Origen_Sistema', 'IVA_Sistema', 'IVA_DIAN', 'Dif_IVA', 'Total_Sistema', 'Total_DIAN', 'Dif_Total', 'Estado', 'Categoria_Resumen']
        for s, f in [('Cuadre Perfecto', 'Cuadre Perfecto'), ('Diferencias Valores', 'Diferencia en Valores'), ('Solo en DIAN', 'Solo en DIAN'), ('Solo en Sistema', 'Solo en Sistema')]:
            cruce[cruce['Estado'] == f][cols].to_excel(writer, sheet_name=s, index=False)
        cruce[cols + ['Transporte_Sistema', 'Transporte_DIAN']].to_excel(writer, sheet_name='Resumen Total Cruce', index=False)
        
        # Nueva hoja de Detalle de Transporte
        if not df_tr_out.empty:
            df_tr_out.to_excel(writer, sheet_name='Detalle Transporte', index=False)
            
        from openpyxl.styles import PatternFill, Font # type: ignore
        for sn in writer.sheets:
            ws = writer.sheets[sn]
            
            if sn == 'Resumen General':
                ws['A1'].font = Font(bold=True, size=12, color="FFFFFF")
                ws['A1'].fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
                if not df_tr_summary.empty:
                    c_titulo = ws.cell(row=start_tr, column=1)
                    c_titulo.font = Font(bold=True, size=12, color="FFFFFF")
                    c_titulo.fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
                
                for cell in ws[2]:
                    cell.fill, cell.font = PatternFill(start_color="1B263B", end_color="1B263B", fill_type="solid"), Font(color="FFFFFF", bold=True)
                if not df_tr_summary.empty:
                    for cell in ws[start_tr + 1]:
                        cell.fill, cell.font = PatternFill(start_color="1B263B", end_color="1B263B", fill_type="solid"), Font(color="FFFFFF", bold=True)
            else:
                for cell in ws[1]:
                    cell.fill, cell.font = PatternFill(start_color="1B263B", end_color="1B263B", fill_type="solid"), Font(color="FFFFFF", bold=True)
                    
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 20
                for cell in col:
                    if sn == 'Resumen General' and (cell.row == 1 or (not df_tr_summary.empty and cell.row == start_tr)):
                        continue
                    if cell.row > 1:
                        if isinstance(cell.value, (int, float)): cell.number_format = '#,##0'
                        if 'TOTAL' in str(ws.cell(cell.row, 1).value).upper():
                            cell.font = Font(bold=True)
                            cell.fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    return {"total_procesadas": len(cruce), "cuadre": len(cruce[cruce['Estado']=='Cuadre Perfecto']), "diferencias": len(cruce[cruce['Estado']=='Diferencia en Valores']), "solo_sis": len(cruce[cruce['Estado']=='Solo en Sistema']), "solo_dian": len(cruce[cruce['Estado']=='Solo en DIAN'])}
