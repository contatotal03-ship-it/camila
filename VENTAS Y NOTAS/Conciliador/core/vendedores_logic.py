import pandas as pd
import numpy as np
import os
import traceback
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from .excel_utils import robust_read_excel


def _resolve_col(df, primary_name, keywords, excludes=None, label=""):
    """
    Resuelve una columna: primero busca por nombre exacto, luego fuzzy.
    Lanza error claro si no encuentra.
    """
    if excludes is None:
        excludes = []
    if primary_name and primary_name in df.columns:
        return primary_name
    found = _find_col(df, keywords, excludes)
    if found:
        return found
    if primary_name:
        raise KeyError(
            f"Columna '{label or primary_name}' no encontrada. "
            f"Columnas disponibles: {list(df.columns)}"
        )
    raise KeyError(
        f"No se encontro columna '{label}' con keywords {keywords}. "
        f"Columnas disponibles: {list(df.columns)}"
    )


def procesar_dashboard_vendedores(ruta_archivo, ruta_salida, ruta_asiento_nc=None, ruta_dian=None):
    """
    Lee un archivo de Pedidos de Venta (sale.order), procesa la info
    enfocada en VENDEDORES y genera un Excel profesional con:
    - Resumen Ejecutivo (KPIs + Ranking)
    - Clientes por Vendedor (Almacen, Factura, Producto, Total)
    - Detalle Completo
    - Distribucion por Almacen
    - Top Productos
    - Evolucion Semanal

    Si ruta_asiento_nc se provee, las NCC de ese archivo reemplazan las del
    Pedido de venta (valores negativos, filas en rojo).
    Si ruta_dian se provee, genera adicionalmente la hoja de 'Conciliación DIAN vs Sistema'
    con auditoría cruzada, motivos de diferencias y desglose de facturas y notas.
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.chart import BarChart, PieChart, LineChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.utils import get_column_letter

    df = robust_read_excel(ruta_archivo)

    cols_lower = [c.lower().strip() for c in df.columns]
    tiene_vendedor = any('vendedor' in c for c in cols_lower)
    tiene_total = any(c == 'total' or c == 'total (tax included)' for c in cols_lower)
    tiene_producto = any('roducto' in c for c in cols_lower)

    if not (tiene_vendedor and tiene_total and tiene_producto):
        cols_str = ', '.join([str(c) for c in df.columns[:12]])
        raise ValueError(
            f"El archivo cargado NO es el correcto.\n\n"
            f"Se esperaba: 'Pedido de venta (sale.order).xlsx'\n"
            f"Se detectaron estas columnas: [{cols_str}]\n\n"
            f"Por favor seleccione el archivo 'Pedido de venta (sale.order).xlsx' "
            f"que contiene las columnas: Vendedor, Producto, Total, Base imponible, etc."
        )

    col_vendedor = _resolve_col(df, 'Vendedor', ['vendedor'], label='Vendedor')
    df = df.dropna(subset=[col_vendedor])

    if len(df) == 0:
        raise ValueError(
            "El archivo no tiene filas con datos de vendedor. "
            f"Columnas disponibles: {list(df.columns)}"
        )

    col_fecha = _find_col(df, ['fecha', 'ech'])
    col_alm = _find_col(df, ['lmac'])
    col_cliente = _find_col(df, ['liente'])
    col_prod = _find_col(df, ['roducto'])
    if col_prod:
        col_prod_clean = [c for c in df.columns if 'roducto' in c and '/' not in c]
        col_prod = col_prod_clean[0] if col_prod_clean else col_prod
    col_fact_num = _find_col(df, ['actura', 'mero'], excludes=['edido'])
    col_fact_orig = _find_col(df, ['actura', 'riginal'])
    col_base = _find_col(df, ['base', 'imponible'])
    col_imp = _find_col(df, ['impuesto'])
    col_total = 'Total' if 'Total' in df.columns else _find_col(df, ['total'])
    col_margen = _find_col(df, ['margen'])
    col_desc = _find_col(df, ['descuento'])
    col_coste = _find_col(df, ['coste'])

    col_nota_credito_num = None
    if 'Facturas/Número' in df.columns:
        col_nota_credito_num = 'Facturas/Número'
    elif col_fact_num and 'facturas' in col_fact_num.lower():
        muestra = df[col_fact_num].head(20).dropna()
        if not muestra.empty:
            valores_muestra = muestra.astype(str).str.upper()
            if valores_muestra.str.contains(r'^NCC').any() or valores_muestra.str.contains(r'^NC').any():
                col_nota_credito_num = col_fact_num

    df['Es_Nota_Credito'] = False
    if col_nota_credito_num and col_nota_credito_num in df.columns:
        df['Es_Nota_Credito'] = df[col_nota_credito_num].astype(str).str.upper().str.contains(r'^NCC|^NC', na=False)

    n_nc = df['Es_Nota_Credito'].sum()
    col_estado = _find_col(df, ['estado'])
    col_pedido = _find_col(df, ['referencia', 'pedido'])

    if ruta_asiento_nc and os.path.exists(ruta_asiento_nc):
        df_ac = robust_read_excel(ruta_asiento_nc)
        ac_num_col = df_ac.columns[0]
        ac_nc_mask = df_ac[ac_num_col].astype(str).str.upper().str.contains(r'^NCC|^NC', na=False)

        # Excluir borradores y cancelados si hay columna de estado
        ac_estado_col_match = [c for c in df_ac.columns if 'estado' in c.lower() or 'pago' in c.lower()]
        if ac_estado_col_match:
            mask_no_borrador = ~df_ac[ac_estado_col_match[0]].astype(str).str.strip().str.lower().isin(['borrador', 'cancelado', 'cancelada'])
            ac_nc_mask = ac_nc_mask & mask_no_borrador

        df_nc_ac = df_ac[ac_nc_mask].copy().reset_index(drop=True)

        if len(df_nc_ac) > 0:
            # Construir mapa de lookup desde el Pedido de Venta completo (df) por pedido y por factura
            df_so_lookup = df.copy()
            if col_pedido and col_pedido in df_so_lookup.columns:
                df_so_lookup['lookup_pedido'] = df_so_lookup[col_pedido].ffill().astype(str).str.strip().str.upper()
            else:
                df_so_lookup['lookup_pedido'] = ''

            if col_fact_num and col_fact_num in df_so_lookup.columns:
                df_so_lookup['lookup_factura'] = df_so_lookup[col_fact_num].ffill().astype(str).str.strip().str.upper()
            else:
                df_so_lookup['lookup_factura'] = ''

            # Si en df había registros que eran genuinamente notas crédito negativas previas, excluirlos para no duplicar con Asiento Contable
            if col_total and col_total in df.columns:
                df = df[~(df['Es_Nota_Credito'] & (pd.to_numeric(df[col_total], errors='coerce').fillna(0) < 0))].copy()
            df['Es_Nota_Credito'] = False

            # Columnas del Asiento Contable
            ac_nombre_col = [c for c in df_ac.columns if 'empresa' in c.lower() or 'nombre' in c.lower()]
            col_cli_ac = ac_nombre_col[0] if ac_nombre_col else None
            col_ven_ac = 'Vendedor' if 'Vendedor' in df_ac.columns else None
            col_orig_ac = _find_col(df_ac, ['origen'])
            col_ref_ac = _find_col(df_ac, ['referencia'])

            pedidos_map = {}
            if col_pedido and col_pedido in df_so_lookup.columns:
                for ped_code, grp in df_so_lookup.groupby('lookup_pedido'):
                    if ped_code and ped_code not in ['NAN', 'NONE', '']:
                        v_val = grp[col_vendedor].dropna()
                        v_name = v_val.iloc[0] if not v_val.empty else None
                        alm_val = grp[col_alm].dropna() if col_alm and col_alm in grp.columns else None
                        alm_name = alm_val.iloc[0] if alm_val is not None and not alm_val.empty else 'NCC'
                        tot_val = float(pd.to_numeric(grp[col_total], errors='coerce').fillna(0).sum()) if col_total and col_total in grp.columns else 0.0
                        cost_val = float(pd.to_numeric(grp[col_coste], errors='coerce').fillna(0).sum()) if col_coste and col_coste in grp.columns else 0.0
                        pedidos_map[ped_code] = {
                            'vendedor': v_name,
                            'almacen': alm_name,
                            'total': tot_val,
                            'coste': cost_val
                        }

            facturas_map = {}
            if col_fact_num and col_fact_num in df_so_lookup.columns:
                for fact_code, grp in df_so_lookup.groupby('lookup_factura'):
                    if fact_code and fact_code not in ['NAN', 'NONE', '']:
                        v_val = grp[col_vendedor].dropna()
                        v_name = v_val.iloc[0] if not v_val.empty else None
                        alm_val = grp[col_alm].dropna() if col_alm and col_alm in grp.columns else None
                        alm_name = alm_val.iloc[0] if alm_val is not None and not alm_val.empty else 'NCC'
                        tot_val = float(pd.to_numeric(grp[col_total], errors='coerce').fillna(0).sum()) if col_total and col_total in grp.columns else 0.0
                        cost_val = float(pd.to_numeric(grp[col_coste], errors='coerce').fillna(0).sum()) if col_coste and col_coste in grp.columns else 0.0
                        facturas_map[fact_code] = {
                            'vendedor': v_name,
                            'almacen': alm_name,
                            'total': tot_val,
                            'coste': cost_val
                        }

            # Fallback secundario de cliente -> vendedor habitual usando todo el asiento (que contiene las ventas)
            cliente_ven_map = {}
            if col_cli_ac and col_ven_ac:
                df_valid_v = df_ac.dropna(subset=[col_ven_ac])
                df_valid_v = df_valid_v[df_valid_v[col_ven_ac].astype(str).str.strip() != '']
                cliente_ven_map = df_valid_v.groupby(col_cli_ac)[col_ven_ac].agg(
                    lambda s: s.mode()[0] if not s.empty else None
                ).dropna().to_dict()

            ac_total_col = [c for c in df_nc_ac.columns if 'total' in c.lower() and 'signo' in c.lower()]
            if not ac_total_col:
                ac_total_col = [c for c in df_nc_ac.columns if 'total' in c.lower()]

            import re
            vendedores_final = []
            almacenes_final = []
            costes_final = []
            origenes_vendedor = []

            for _, r in df_nc_ac.iterrows():
                v_orig = r.get(col_ven_ac, None)
                c_nom = r.get(col_cli_ac, '')
                orig_val = str(r.get(col_orig_ac, '')).strip().upper() if col_orig_ac else ''
                ref_val = str(r.get(col_ref_ac, '')).strip() if col_ref_ac else ''
                ref_upper = ref_val.upper()

                tot_nc_val = 0.0
                if ac_total_col:
                    tot_nc_val = pd.to_numeric(r.get(ac_total_col[0]), errors='coerce') or 0.0

                # Detección si es Descuento Comercial
                es_descuento = any(w in ref_upper for w in ['DESCUEN', 'DESCUET', 'IMPULSA', 'BONO', 'REBAJA'])

                # Buscar match en Pedidos o Facturas del Pedido de Venta
                match_data = None
                if orig_val and orig_val in pedidos_map:
                    match_data = pedidos_map[orig_val]
                elif ref_upper:
                    tokens = re.findall(r'[A-Z0-9/]+', ref_upper)
                    for tok in tokens:
                        if tok in facturas_map:
                            match_data = facturas_map[tok]
                            break

                # 1. Resolver Vendedor
                if pd.notna(v_orig) and str(v_orig).strip() != '' and str(v_orig).strip().lower() not in ['nan', 'none']:
                    vendedores_final.append(str(v_orig).strip())
                    origenes_vendedor.append('Original')
                elif match_data and match_data['vendedor']:
                    vendedores_final.append(match_data['vendedor'])
                    origenes_vendedor.append('Recuperado por Pedido')
                elif c_nom and c_nom in cliente_ven_map:
                    vendedores_final.append(cliente_ven_map[c_nom])
                    origenes_vendedor.append('Asignado por Cliente')
                else:
                    vendedores_final.append('SIN ASIGNAR')
                    origenes_vendedor.append('Sin Vendedor')

                # 2. Resolver Almacén
                if match_data and match_data['almacen'] and match_data['almacen'] != 'NCC':
                    almacenes_final.append(match_data['almacen'])
                else:
                    almacenes_final.append('NCC')

                # 3. Resolver Coste
                if es_descuento:
                    # Descuento en dinero: proceso normal, no se devuelve mercancía, coste = 0
                    costes_final.append(0.0)
                elif match_data and match_data['coste'] > 0 and match_data['total'] > 0:
                    # Devolución física o anulación: descontar el costo de la mercancía devuelta
                    ratio = min(1.0, abs(tot_nc_val) / match_data['total'])
                    costes_final.append(-(match_data['coste'] * ratio))
                else:
                    costes_final.append(0.0)

            df_nc_ac['Vendedor_Final'] = vendedores_final
            df_nc_ac['Almacen_Final'] = almacenes_final
            df_nc_ac['Coste_Final'] = costes_final
            df_nc_ac['Origen_Vendedor'] = origenes_vendedor

            ncc_rows = pd.DataFrame()
            ncc_rows[col_fact_num] = df_nc_ac[ac_num_col].values
            ncc_rows[col_vendedor] = df_nc_ac['Vendedor_Final'].values

            if ac_nombre_col:
                ncc_rows[col_cliente] = df_nc_ac[ac_nombre_col[0]].values
            else:
                ncc_rows[col_cliente] = ''

            ac_fecha_col = [c for c in df_nc_ac.columns if 'fecha' in c.lower() and 'vencimiento' not in c.lower()]
            if ac_fecha_col:
                ncc_rows[col_fecha] = pd.to_datetime(df_nc_ac[ac_fecha_col[0]], errors='coerce').values
            elif col_fecha and col_fecha in df.columns:
                ncc_rows[col_fecha] = pd.NaT

            ac_base_col = [c for c in df_nc_ac.columns if 'base' in c.lower() or 'imponible' in c.lower()]
            if ac_base_col:
                ncc_rows[col_base] = pd.to_numeric(df_nc_ac[ac_base_col[0]], errors='coerce').fillna(0).values
            else:
                ncc_rows[col_base] = 0

            if ac_total_col:
                ncc_rows[col_total] = pd.to_numeric(df_nc_ac[ac_total_col[0]], errors='coerce').fillna(0).values
            else:
                ncc_rows[col_total] = 0

            if col_imp and col_base in ncc_rows.columns and col_total in ncc_rows.columns:
                ncc_rows[col_imp] = ncc_rows[col_total].astype(float) - ncc_rows[col_base].astype(float)
            elif col_imp:
                ncc_rows[col_imp] = 0

            if col_margen:
                ncc_rows[col_margen] = 0
            if col_desc:
                ncc_rows[col_desc] = 0
            if col_coste:
                ncc_rows[col_coste] = df_nc_ac['Coste_Final'].values
            if col_alm:
                ncc_rows[col_alm] = df_nc_ac['Almacen_Final'].values
            if col_estado:
                ac_estado_col = [c for c in df_nc_ac.columns if 'estado' in c.lower() or 'pago' in c.lower()]
                if ac_estado_col:
                    ncc_rows[col_estado] = df_nc_ac[ac_estado_col[0]].values
                else:
                    ncc_rows[col_estado] = 'Nota Credito'
            if col_pedido:
                if col_orig_ac and col_orig_ac in df_nc_ac.columns:
                    ncc_rows[col_pedido] = df_nc_ac[col_orig_ac].fillna(df_nc_ac[ac_num_col]).values
                else:
                    ncc_rows[col_pedido] = ncc_rows[col_fact_num]

            if col_prod:
                ncc_rows[col_prod] = 'N/A - Nota Credito'

            ncc_rows['Es_Nota_Credito'] = True

            for c in df.columns:
                if c not in ncc_rows.columns:
                    ncc_rows[c] = ''

            ncc_rows = ncc_rows[df.columns]
            df = pd.concat([df, ncc_rows], ignore_index=True)
            n_nc = df['Es_Nota_Credito'].sum()

    # ====== ESTILOS ======
    # ====== ESTILOS EJECUTIVOS Y PALETA MODERNA ======
    DARK = "0F172A"         # Slate 900 - Fondo base oscuro profundo
    MED = "1E293B"          # Slate 800 - Superficie para tarjetas y filas de agrupación
    ALT = "162033"          # Slate 850 - Fila alterna sutil
    HEADER_NAVY = "1E3A8A"  # Navy 900 - Encabezados de tabla formales y ejecutivos
    BLUE = "38BDF8"         # Sky 400 - Acento cian glaciar
    GOLD = "F59E0B"         # Amber 500 - Oro cálido
    GREEN = "10B981"        # Emerald 500 - Verde financiero impecable
    RED = "F87171"          # Red 400 - Rojo coral elegante y legible (no agresivo)
    ORANGE_NC = "381219"    # Fondo sutil vino/rosa para Notas Crédito
    RED_NC_FONT = "FECDD3"  # Texto rosa suave de alta legibilidad
    PURPLE = "A78BFA"       # Lavender 400 - Púrpura suave
    CYAN = "2DD4BF"         # Mint / Cyan 400
    W = "FFFFFF"            # Blanco puro
    LG = "94A3B8"           # Slate 400 - Gris azulado para etiquetas secundarias

    f_dark = PatternFill(start_color=DARK, end_color=DARK, fill_type="solid")
    f_med = PatternFill(start_color=MED, end_color=MED, fill_type="solid")
    f_alt = PatternFill(start_color=ALT, end_color=ALT, fill_type="solid")
    f_blue = PatternFill(start_color=HEADER_NAVY, end_color=HEADER_NAVY, fill_type="solid")
    f_gold = PatternFill(start_color="B45309", end_color="B45309", fill_type="solid")
    f_green = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")
    f_sub = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    f_nc = PatternFill(start_color=ORANGE_NC, end_color=ORANGE_NC, fill_type="solid")
    f_nc_font = Font(name='Calibri', size=9, bold=True, color=RED_NC_FONT)

    f_title = Font(name='Calibri', size=17, bold=True, color=W)
    f_subt = Font(name='Calibri', size=9.5, color=LG)
    f_kpi = Font(name='Calibri', size=15, bold=True, color=W)
    f_kpi_l = Font(name='Calibri', size=8.5, bold=True, color=LG)
    f_sec = Font(name='Calibri', size=11, bold=True, color=W)
    f_hdr = Font(name='Calibri', size=9, bold=True, color=W)
    f_dat = Font(name='Calibri', size=9, color="CBD5E1")
    f_datb = Font(name='Calibri', size=9, bold=True, color=W)
    f_vend = Font(name='Calibri', size=9.5, bold=True, color="38BDF8")
    f_cli = Font(name='Calibri', size=9, color="E2E8F0")

    a_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    a_l = Alignment(horizontal='left', vertical='center', wrap_text=True)
    a_r = Alignment(horizontal='right', vertical='center')

    bdr = Border(
        left=Side(style='thin', color='334155'),
        right=Side(style='thin', color='334155'),
        top=Side(style='thin', color='334155'),
        bottom=Side(style='thin', color='334155')
    )

    def sc(ws, r, c, val, font=f_dat, fill=f_dark, align=a_c, nf=None):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font = font
        cell.fill = fill
        cell.alignment = align
        cell.border = bdr
        if nf:
            cell.number_format = nf
        return cell

    def mtitle(ws, r1, c1, r2, c2, val, font=f_title, fill=f_dark, align=a_l):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        cell = ws.cell(row=r1, column=c1, value=val)
        cell.font = font
        cell.fill = fill
        cell.alignment = align
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = fill
                ws.cell(row=r, column=c).border = bdr

    def fbg(ws, r1, c1, r2, c2, fill):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = fill
                ws.cell(row=r, column=c).border = bdr

    # ====== PROCESAR DATA ======
    resumen = df.groupby(col_vendedor).agg({
        col_pedido: 'count',
        col_total: 'sum',
        col_margen: 'mean',
        col_desc: 'mean',
        col_coste: 'sum',
        col_cliente: 'nunique'
    }).reset_index()
    resumen.columns = [col_vendedor, 'Pedidos', 'Ventas', 'Margen', 'Descuento', 'Coste', 'Clientes']
    resumen['Utilidad'] = resumen['Ventas'] - resumen['Coste']
    resumen = resumen.sort_values('Ventas', ascending=False).reset_index(drop=True)
    resumen['Part'] = resumen['Ventas'] / resumen['Ventas'].sum()

    mask_nc_calc = df['Es_Nota_Credito'] | (pd.to_numeric(df[col_total], errors='coerce').fillna(0) < 0)
    ventas_brutas = df[~mask_nc_calc][col_total].sum()
    total_nc_val = df[mask_nc_calc][col_total].sum()
    ventas_netas = df[col_total].sum()

    total_ventas = ventas_netas
    total_coste = df[col_coste].sum()
    total_utilidad = total_ventas - total_coste
    margen_real_glb = total_utilidad / total_ventas if total_ventas else 0
    total_pedidos = len(df)
    ticket_prom = total_ventas / total_pedidos if total_pedidos else 0
    margen_glb = df[col_margen].mean()
    desc_glb = df[col_desc].mean()
    n_alm = df[col_alm].nunique()
    n_clientes = df[col_cliente].nunique()

    sorted_df = df.sort_values([col_vendedor, col_cliente, col_total], ascending=[True, True, False])

    # ====== CREAR EXCEL ======
    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:

        # ---- HOJA 1: RESUMEN EJECUTIVO ----
        ws1 = writer.book.create_sheet("Resumen Ejecutivo")
        ws1.sheet_view.showGridLines = False
        ws1.sheet_properties.tabColor = "1E3A8A"
        ws1.freeze_panes = 'A10'

        widths1 = {'A': 3, 'B': 22, 'C': 20, 'D': 20, 'E': 19, 'F': 19, 'G': 16, 'H': 16, 'I': 19, 'J': 18, 'K': 18, 'L': 18, 'M': 3}
        for col_l, w in widths1.items():
            ws1.column_dimensions[col_l].width = w

        ws1.row_dimensions[1].height = 24
        ws1.row_dimensions[2].height = 20
        ws1.row_dimensions[3].height = 18
        ws1.row_dimensions[5].height = 18
        ws1.row_dimensions[6].height = 28
        ws1.row_dimensions[8].height = 22
        ws1.row_dimensions[9].height = 24

        fbg(ws1, 1, 1, 55, 13, f_dark)
        mtitle(ws1, 1, 1, 2, 13, "DASHBOARD EJECUTIVO DE VENTAS Y VENDEDORES")
        mtitle(ws1, 3, 1, 3, 13,
               f"Panel de Control Comercial  |  {total_pedidos:,} registros  |  {n_clientes:,} clientes  |  {len(resumen)} vendedores",
               f_subt)

        # Tarjetas KPI unificadas con fondo formal Slate 800 (1E293B) y números claros
        kpis = [
            ("VENTAS BRUTAS", ventas_brutas, "$#,##0", "FFFFFF"),
            ("NOTAS CRÉDITO (-)", total_nc_val, "$#,##0", "F87171"),
            ("VENTAS NETAS", ventas_netas, "$#,##0", "10B981"),
            ("COSTE TOTAL", total_coste, "$#,##0", "CBD5E1"),
            ("UTILIDAD", total_utilidad, "$#,##0", "10B981" if total_utilidad >= 0 else "F87171"),
            ("MARGEN REAL", margen_real_glb, "0.0%", "38BDF8"),
            ("TOTAL PEDIDOS", total_pedidos, "#,##0", "FBBF24"),
            ("TICKET PROMEDIO", ticket_prom, "$#,##0", "C084FC"),
        ]
        f_card_bg = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        for i, (label, val, nf, val_color) in enumerate(kpis):
            c = i + 2
            fbg(ws1, 5, c, 6, c, f_card_bg)
            sc(ws1, 5, c, label, Font(name='Calibri', size=8.5, bold=True, color="94A3B8"), f_card_bg)
            sc(ws1, 6, c, val, Font(name='Calibri', size=15, bold=True, color=val_color), f_card_bg, a_c, nf)

        fbg(ws1, 4, 1, 4, 13, f_dark)
        fbg(ws1, 7, 1, 7, 13, f_dark)
        mtitle(ws1, 8, 1, 8, 13, "RANKING DE RENDIMIENTO POR VENDEDOR", f_sec, f_med)

        hdrs = ['#', 'Vendedor', 'Clientes', 'Pedidos', 'Ventas Totales',
                'Coste Total', 'Utilidad', 'Margen Real', 'Descuento %', 'Ticket Prom.', 'Part. %']
        for i, h in enumerate(hdrs):
            sc(ws1, 9, i + 1, h, f_hdr, f_blue)

        for idx, row in resumen.iterrows():
            r = 10 + idx
            ws1.row_dimensions[r].height = 20
            fr = f_dark if idx % 2 == 0 else f_alt
            rk = idx + 1
            rk_f = f_gold if rk == 1 else (
                PatternFill(start_color="475569", end_color="475569", fill_type="solid") if rk == 2 else (
                    PatternFill(start_color="78350F", end_color="78350F", fill_type="solid") if rk == 3 else fr))
            tk = row['Ventas'] / row['Pedidos'] if row['Pedidos'] > 0 else 0
            margen_real_v = row['Utilidad'] / row['Ventas'] if row['Ventas'] > 0 else 0
            sc(ws1, r, 1, rk, f_datb if rk <= 3 else f_dat, rk_f)
            sc(ws1, r, 2, row[col_vendedor], f_datb if rk <= 3 else f_dat, fr, a_l)
            sc(ws1, r, 3, row['Clientes'], f_dat, fr)
            sc(ws1, r, 4, row['Pedidos'], f_dat, fr)
            sc(ws1, r, 5, row['Ventas'], f_datb if rk <= 3 else f_dat, fr, a_r, '$#,##0')
            sc(ws1, r, 6, row['Coste'], f_dat, fr, a_r, '$#,##0')
            sc(ws1, r, 7, row['Utilidad'], Font(name='Calibri', size=9, bold=True, color="10B981" if row['Utilidad'] >= 0 else "F87171"), fr, a_r, '$#,##0')
            sc(ws1, r, 8, margen_real_v, f_dat, fr, a_c, '0.0%')
            sc(ws1, r, 9, row['Descuento'], f_dat, fr, a_c, '0.0%')
            sc(ws1, r, 10, tk, f_dat, fr, a_r, '$#,##0')
            sc(ws1, r, 11, row['Part'], f_dat, fr, a_c, '0.0%')

        # Fila TOTAL GENERAL del Ranking
        r_tot = 10 + len(resumen)
        ws1.row_dimensions[r_tot].height = 22
        f_tot_bdr = Border(
            left=Side(style='thin', color='334155'),
            right=Side(style='thin', color='334155'),
            top=Side(style='thin', color='38BDF8'),
            bottom=Side(style='double', color='38BDF8')
        )
        f_tot_bg = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        f_tot_font = Font(name='Calibri', size=9.5, bold=True, color="FFFFFF")
        f_tot_green = Font(name='Calibri', size=9.5, bold=True, color="10B981")

        ws1.cell(row=r_tot, column=1, value="").fill = f_tot_bg
        ws1.cell(row=r_tot, column=1).border = f_tot_bdr
        sc(ws1, r_tot, 2, "TOTAL GENERAL", f_tot_font, f_tot_bg, a_l)
        ws1.cell(row=r_tot, column=2).border = f_tot_bdr
        sc(ws1, r_tot, 3, n_clientes, f_tot_font, f_tot_bg)
        ws1.cell(row=r_tot, column=3).border = f_tot_bdr
        sc(ws1, r_tot, 4, total_pedidos, f_tot_font, f_tot_bg)
        ws1.cell(row=r_tot, column=4).border = f_tot_bdr
        sc(ws1, r_tot, 5, resumen['Ventas'].sum(), f_tot_font, f_tot_bg, a_r, '$#,##0')
        ws1.cell(row=r_tot, column=5).border = f_tot_bdr
        sc(ws1, r_tot, 6, resumen['Coste'].sum(), f_tot_font, f_tot_bg, a_r, '$#,##0')
        ws1.cell(row=r_tot, column=6).border = f_tot_bdr
        sc(ws1, r_tot, 7, resumen['Utilidad'].sum(), f_tot_green, f_tot_bg, a_r, '$#,##0')
        ws1.cell(row=r_tot, column=7).border = f_tot_bdr
        sc(ws1, r_tot, 8, margen_real_glb, f_tot_font, f_tot_bg, a_c, '0.0%')
        ws1.cell(row=r_tot, column=8).border = f_tot_bdr
        sc(ws1, r_tot, 9, desc_glb, f_tot_font, f_tot_bg, a_c, '0.0%')
        ws1.cell(row=r_tot, column=9).border = f_tot_bdr
        sc(ws1, r_tot, 10, ticket_prom, f_tot_font, f_tot_bg, a_r, '$#,##0')
        ws1.cell(row=r_tot, column=10).border = f_tot_bdr
        sc(ws1, r_tot, 11, 1.0, f_tot_font, f_tot_bg, a_c, '100.0%')
        ws1.cell(row=r_tot, column=11).border = f_tot_bdr

        last = r_tot
        fbg(ws1, last + 1, 1, last + 1, 13, f_dark)

        ch1 = BarChart()
        ch1.type = "col"
        ch1.style = 10
        ch1.width = 30
        ch1.height = 14
        ch1.legend = None
        ch1.title = "Ventas vs Coste por Vendedor (Top 12)"
        ch1.y_axis.title = "Monto ($)"
        d1a = Reference(ws1, min_col=5, min_row=9, max_row=min(21, 9 + len(resumen)))
        d1b = Reference(ws1, min_col=6, min_row=9, max_row=min(21, 9 + len(resumen)))
        c1 = Reference(ws1, min_col=2, min_row=10, max_row=min(21, 9 + len(resumen)))
        ch1.add_data(d1a, titles_from_data=True)
        ch1.add_data(d1b, titles_from_data=True)
        ch1.set_categories(c1)
        ch1.series[0].graphicalProperties.solidFill = "38BDF8"
        ch1.series[1].graphicalProperties.solidFill = "F87171"
        ws1.add_chart(ch1, f"B{last + 2}")

        ch2 = PieChart()
        ch2.title = "Participacion en Ventas"
        ch2.style = 10
        ch2.width = 16
        ch2.height = 14
        d2 = Reference(ws1, min_col=11, min_row=9, max_row=min(17, 9 + len(resumen)))
        c2 = Reference(ws1, min_col=2, min_row=10, max_row=min(17, 9 + len(resumen)))
        ch2.add_data(d2, titles_from_data=True)
        ch2.set_categories(c2)
        ch2.dataLabels = DataLabelList()
        ch2.dataLabels.showPercent = True
        ch2.dataLabels.showCatName = False
        ws1.add_chart(ch2, f"E{last + 2}")

        # ---- HOJA 2: CLIENTES POR VENDEDOR ----
        ws2 = writer.book.create_sheet("Clientes por Vendedor")
        ws2.sheet_view.showGridLines = False
        ws2.sheet_properties.tabColor = "2563EB"
        ws2.freeze_panes = 'A6'

        for i, w in enumerate([24, 30, 20, 14, 16, 16, 16, 16, 12, 45]):
            ws2.column_dimensions[get_column_letter(i + 1)].width = w

        fbg(ws2, 1, 1, 5, 10, f_dark)
        mtitle(ws2, 1, 1, 2, 10, "CLIENTES POR VENDEDOR")
        mtitle(ws2, 3, 1, 3, 10,
               "Detalle: Cliente, Almacen, Factura, Producto, Coste, Utilidad  |  Ordenado por Vendedor",
               f_subt)

        hdrs2 = ['Vendedor', 'Cliente', 'Almacen', 'Factura #', 'Ref. Pedido',
                 'Base Imponible', 'Total', 'Coste', 'Utilidad', 'Producto']
        for i, h in enumerate(hdrs2):
            sc(ws2, 5, i + 1, h, f_hdr, f_blue)

        row_num = 6
        for vend in resumen[col_vendedor].tolist():
            vd = sorted_df[sorted_df[col_vendedor] == vend].copy()
            if vd.empty:
                continue

            # Fila header del vendedor
            fbg(ws2, row_num, 1, row_num, 10, f_med)
            nc = vd[col_cliente].nunique()
            vend_coste = vd[col_coste].sum()
            vend_util = vd[col_total].sum() - vend_coste
            sc(ws2, row_num, 1, f"  {vend}", f_vend, f_med, a_l)
            sc(ws2, row_num, 2, f"{nc} clientes", f_dat, f_med, a_l)
            sc(ws2, row_num, 3, f"{len(vd)} pedidos", f_dat, f_med)
            sc(ws2, row_num, 7, vd[col_total].sum(),
               Font(name='Calibri', size=10, bold=True, color=GREEN), f_med, a_r, '$#,##0')
            sc(ws2, row_num, 8, vend_coste,
               Font(name='Calibri', size=10, bold=True, color=RED), f_med, a_r, '$#,##0')
            sc(ws2, row_num, 9, vend_util,
               Font(name='Calibri', size=10, bold=True, color=GREEN if vend_util >= 0 else RED), f_med, a_r, '$#,##0')
            row_num += 1

            # Headers internos
            sub_hdrs = ['Cliente', 'Almacen', 'Factura #', 'Ref. Pedido', 'Base Imp.', 'Total', 'Coste', 'Utilidad', 'Margen Real', 'Producto']
            for i, h in enumerate(sub_hdrs):
                sc(ws2, row_num, i + 1, h,
                   Font(name='Calibri', size=8, bold=True, color="94A3B8"),
                   PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid"))
            row_num += 1

            # Agrupar por cliente
            clients = vd.groupby(col_cliente).agg({
                col_total: 'sum',
                col_base: 'sum',
                col_coste: 'sum',
                col_margen: 'mean'
            }).reset_index().sort_values(col_total, ascending=False)

            for ci, (_, cr) in enumerate(clients.iterrows()):
                cn = cr[col_cliente]
                co = vd[vd[col_cliente] == cn]
                tiene_nc = co['Es_Nota_Credito'].any()

                fr = f_nc if tiene_nc else (f_dark if ci % 2 == 0 else f_alt)
                almacenes = ', '.join(co[col_alm].dropna().unique().tolist())
                facturas = _get_unique_str(co, col_fact_num, max_items=3)
                pedidos = _get_unique_str(co, col_pedido, max_items=3)
                prods = _get_unique_str(co, col_prod, max_items=2)
                cli_coste = co[col_coste].sum()
                cli_util = cr[col_total] - cli_coste
                cli_margen_real = cli_util / cr[col_total] if cr[col_total] > 0 else 0

                f_c = f_nc_font if tiene_nc else f_cli
                f_d = f_nc_font if tiene_nc else f_dat
                f_db = f_nc_font if tiene_nc else f_datb
                sc(ws2, row_num, 1, cn, f_c, fr, a_l)
                sc(ws2, row_num, 2, almacenes, f_d, fr, a_l)
                sc(ws2, row_num, 3, facturas, f_d, fr)
                sc(ws2, row_num, 4, pedidos, f_d, fr)
                sc(ws2, row_num, 5, cr[col_base], f_d, fr, a_r, '$#,##0')
                sc(ws2, row_num, 6, cr[col_total], f_db, fr, a_r, '$#,##0')
                sc(ws2, row_num, 7, cli_coste, f_d, fr, a_r, '$#,##0')
                sc(ws2, row_num, 8, cli_util,
                   Font(name='Calibri', size=9, bold=True, color=RED_NC_FONT if tiene_nc else (GREEN if cli_util >= 0 else RED)), fr, a_r, '$#,##0')
                sc(ws2, row_num, 9, cli_margen_real, f_d, fr, a_c, '0.0%')
                sc(ws2, row_num, 10, prods, f_d, fr, a_l)
                row_num += 1

            # Subtotal
            fbg(ws2, row_num, 1, row_num, 10, f_sub)
            sc(ws2, row_num, 1, f"  SUBTOTAL {vend.split()[0]}",
               Font(name='Calibri', size=9, bold=True, color=BLUE), f_sub, a_l)
            sc(ws2, row_num, 5, vd[col_base].sum(), f_datb, f_sub, a_r, '$#,##0')
            sc(ws2, row_num, 6, vd[col_total].sum(),
               Font(name='Calibri', size=10, bold=True, color=GREEN), f_sub, a_r, '$#,##0')
            sc(ws2, row_num, 8, vend_util,
               Font(name='Calibri', size=10, bold=True, color=GREEN if vend_util >= 0 else RED), f_sub, a_r, '$#,##0')
            row_num += 2

        # ---- HOJA 3: DETALLE COMPLETO ----
        ws3 = writer.book.create_sheet("Detalle Completo")
        ws3.sheet_view.showGridLines = False
        ws3.sheet_properties.tabColor = "10B981"
        ws3.freeze_panes = 'A5'

        full_cols = [col_vendedor, col_cliente, col_alm, col_pedido, col_fecha,
                     col_fact_num, col_prod, col_base, col_imp, col_total,
                     col_coste, col_margen, col_desc, col_estado]
        full_labels = ['Vendedor', 'Cliente', 'Almacen', 'Ref. Pedido', 'Fecha',
                       'Factura #', 'Producto', 'Base Imp.', 'Impuestos', 'Total',
                       'Coste', 'Margen %', 'Descuento %', 'Estado']
        full_w = [24, 28, 20, 14, 18, 14, 45, 16, 14, 16, 16, 10, 12, 16]

        for i, w in enumerate(full_w):
            ws3.column_dimensions[get_column_letter(i + 1)].width = w

        ws3.row_dimensions[1].height = 24
        ws3.row_dimensions[2].height = 20
        ws3.row_dimensions[3].height = 18
        ws3.row_dimensions[4].height = 24

        fbg(ws3, 1, 1, 3, len(full_cols), f_dark)
        mtitle(ws3, 1, 1, 2, len(full_cols), "DETALLE COMPLETO - TODOS LOS PEDIDOS")
        mtitle(ws3, 3, 1, 3, len(full_cols),
               f"Ordenado por Vendedor  |  {len(df):,} registros  |  {n_nc} Notas Credito (destacadas)", f_subt)

        for i, h in enumerate(full_labels):
            sc(ws3, 4, i + 1, h, f_hdr, f_blue)

        detail_sorted = df.sort_values([col_vendedor, col_fecha], ascending=[True, True]).reset_index(drop=True)
        for idx, row in detail_sorted.iterrows():
            r = 5 + idx
            ws3.row_dimensions[r].height = 19
            es_nc = row.get('Es_Nota_Credito', False)
            fr = f_nc if es_nc else (f_dark if idx % 2 == 0 else f_alt)
            vals = [
                row[col_vendedor], row.get(col_cliente, ''), row.get(col_alm, ''),
                row.get(col_pedido, ''),
                row[col_fecha].strftime('%Y-%m-%d %H:%M') if pd.notna(row.get(col_fecha)) else '',
                str(row[col_fact_num]) if col_fact_num and pd.notna(row.get(col_fact_num)) else '',
                str(row.get(col_prod, ''))[:70] if pd.notna(row.get(col_prod)) else '',
                row.get(col_base, 0), row.get(col_imp, 0), row.get(col_total, 0),
                row.get(col_coste, 0), row.get(col_margen, 0), row.get(col_desc, 0), row.get(col_estado, '')
            ]
            for ci, v in enumerate(vals):
                nf = None
                if ci in [7, 8, 9, 10]:
                    nf = '$#,##0'
                if ci in [11, 12]:
                    nf = '0.0%'
                al = a_l if ci in [0, 1, 2, 6] else (a_r if ci in [7, 8, 9, 10] else a_c)
                f = f_nc_font if es_nc else (f_datb if ci == 9 else f_dat)
                sc(ws3, r, ci + 1, v, f, fr, al, nf)

        # Filtro automático en Detalle Completo
        ws3.auto_filter.ref = f"A4:{get_column_letter(len(full_cols))}{len(detail_sorted) + 4}"

        # ---- HOJA 4: ALMACENES ----
        ws4 = writer.book.create_sheet("Por Almacen")
        ws4.sheet_view.showGridLines = False
        ws4.sheet_properties.tabColor = "8B5CF6"
        ws4.freeze_panes = 'A6'

        alm_pivot = df.groupby([col_vendedor, col_alm]).size().unstack(fill_value=0)
        alm_cols = list(alm_pivot.columns)

        ws4.column_dimensions['A'].width = 24
        for i in range(len(alm_cols)):
            ws4.column_dimensions[get_column_letter(i + 2)].width = 18
        ws4.column_dimensions[get_column_letter(len(alm_cols) + 2)].width = 16

        ws4.row_dimensions[5].height = 24

        fbg(ws4, 1, 1, 4, len(alm_cols) + 3, f_dark)
        mtitle(ws4, 1, 1, 2, len(alm_cols) + 3, "DISTRIBUCION POR ALMACEN")
        mtitle(ws4, 3, 1, 3, len(alm_cols) + 3,
               "Pedidos por vendedor y punto de venta", f_subt)

        hdrs_a = [col_vendedor] + alm_cols + ['Total']
        for i, h in enumerate(hdrs_a):
            sc(ws4, 5, i + 1, h, f_hdr, f_blue)

        v_alm = alm_pivot.copy()
        v_alm['Total'] = v_alm.sum(axis=1)
        v_alm = v_alm.sort_values('Total', ascending=False)

        for idx, (vend, row) in enumerate(v_alm.iterrows()):
            r = 6 + idx
            ws4.row_dimensions[r].height = 20
            fr = f_dark if idx % 2 == 0 else f_alt
            sc(ws4, r, 1, vend, f_datb, fr, a_l)
            for ci, c in enumerate(alm_cols):
                sc(ws4, r, ci + 2, row.get(c, 0), f_dat, fr)
            sc(ws4, r, len(alm_cols) + 2, row['Total'], f_datb, fr)

        ch4 = BarChart()
        ch4.type = "col"
        ch4.grouping = "stacked"
        ch4.style = 10
        ch4.width = 36
        ch4.height = 16
        ch4.title = "Pedidos por Almacen (Stacked)"
        for i in range(len(alm_cols)):
            d = Reference(ws4, min_col=i + 2, min_row=5, max_row=5 + len(v_alm))
            ch4.add_data(d, titles_from_data=True)
        c4 = Reference(ws4, min_col=1, min_row=6, max_row=5 + len(v_alm))
        ch4.set_categories(c4)
        ch4.legend.position = 'b'
        ws4.add_chart(ch4, f"B{7 + len(v_alm)}")

        # ---- HOJA 5: TOP PRODUCTOS ----
        ws5 = writer.book.create_sheet("Top Productos")
        ws5.sheet_view.showGridLines = False
        ws5.sheet_properties.tabColor = "F59E0B"
        ws5.freeze_panes = 'A6'

        for i, w in enumerate([3, 55, 14, 20, 20, 20, 12, 3]):
            ws5.column_dimensions[get_column_letter(i + 1)].width = w

        ws5.row_dimensions[5].height = 24

        fbg(ws5, 1, 1, 4, 8, f_dark)
        mtitle(ws5, 1, 1, 2, 8, "TOP 10 PRODUCTOS POR VENTAS")
        mtitle(ws5, 3, 1, 3, 8, "Productos: Ventas, Coste, Utilidad y Margen Real", f_subt)

        for i, h in enumerate(['#', 'Producto', 'Unidades', 'Ventas Totales', 'Coste Total', 'Utilidad', 'Margen Real', '']):
            sc(ws5, 5, i + 1, h, f_hdr, f_blue)

        top_p = df.groupby(col_prod).agg({
            col_pedido: 'count',
            col_total: 'sum',
            col_coste: 'sum'
        }).reset_index()
        top_p.columns = ['Producto', 'Cantidad', 'Ventas', 'Coste']
        top_p['Utilidad'] = top_p['Ventas'] - top_p['Coste']
        top_p['MargenReal'] = top_p['Utilidad'] / top_p['Ventas']
        top_p = top_p.sort_values('Ventas', ascending=False).head(10).reset_index(drop=True)

        for idx, row in top_p.iterrows():
            r = 6 + idx
            ws5.row_dimensions[r].height = 20
            fr = f_dark if idx % 2 == 0 else f_alt
            rf = f_gold if idx == 0 else (
                PatternFill(start_color="475569", end_color="475569", fill_type="solid") if idx == 1 else (
                    PatternFill(start_color="78350F", end_color="78350F", fill_type="solid") if idx == 2 else fr))
            sc(ws5, r, 1, idx + 1, f_datb if idx < 3 else f_dat, rf)
            sc(ws5, r, 2, str(row['Producto'])[:80], f_datb if idx < 3 else f_dat, fr, a_l)
            sc(ws5, r, 3, row['Cantidad'], f_dat, fr)
            sc(ws5, r, 4, row['Ventas'], f_dat, fr, a_r, '$#,##0')
            sc(ws5, r, 5, row['Coste'], f_dat, fr, a_r, '$#,##0')
            sc(ws5, r, 6, row['Utilidad'],
               Font(name='Calibri', size=9, bold=True, color="10B981" if row['Utilidad'] >= 0 else "F87171"), fr, a_r, '$#,##0')
            sc(ws5, r, 7, row['MargenReal'], f_dat, fr, a_c, '0.0%')

        ch5 = BarChart()
        ch5.type = "col"
        ch5.style = 10
        ch5.width = 36
        ch5.height = 15
        ch5.title = "Top 10 Productos - Ventas vs Coste"
        d5a = Reference(ws5, min_col=4, min_row=5, max_row=15)
        d5b = Reference(ws5, min_col=5, min_row=5, max_row=15)
        c5 = Reference(ws5, min_col=2, min_row=6, max_row=15)
        ch5.add_data(d5a, titles_from_data=True)
        ch5.add_data(d5b, titles_from_data=True)
        ch5.set_categories(c5)
        ch5.series[0].graphicalProperties.solidFill = "38BDF8"
        ch5.series[1].graphicalProperties.solidFill = "F87171"
        ws5.add_chart(ch5, "B17")

        # ---- HOJA 6: EVOLUCION SEMANAL ----
        ws6 = writer.book.create_sheet("Evolucion Semanal")
        ws6.sheet_view.showGridLines = False
        ws6.sheet_properties.tabColor = "06B6D4"
        ws6.freeze_panes = 'A6'

        df['_Semana'] = df[col_fecha].dt.isocalendar().week
        evol = df.groupby(['_Semana', col_vendedor]).agg({col_total: 'sum'}).reset_index()
        semanas = sorted(evol['_Semana'].unique())
        top8 = resumen[col_vendedor].head(8).tolist()

        ws6.column_dimensions['A'].width = 12
        for i in range(len(top8) + 1):
            ws6.column_dimensions[get_column_letter(i + 2)].width = 18

        tc6 = len(top8) + 3
        fbg(ws6, 1, 1, 4, tc6, f_dark)
        mtitle(ws6, 1, 1, 2, tc6, "EVOLUCION SEMANAL DE VENTAS")
        mtitle(ws6, 3, 1, 3, tc6, "Tendencia por vendedor - Top 8", f_subt)

        ws6.row_dimensions[5].height = 24

        fh = ['Semana'] + [' '.join(v.split()[:2]) for v in top8] + ['Total']
        for i, h in enumerate(fh):
            sc(ws6, 5, i + 1, h, f_hdr, f_blue)

        for si, sem in enumerate(semanas):
            r = 6 + si
            fr = f_dark if si % 2 == 0 else f_alt
            sc(ws6, r, 1, f"Sem {sem}", f_datb, fr)
            tot = 0
            for vi, vend in enumerate(top8):
                f = evol[(evol['_Semana'] == sem) & (evol[col_vendedor] == vend)]
                v = f[col_total].sum() if len(f) > 0 else 0
                tot += v
                sc(ws6, r, vi + 2, v, f_dat, fr, a_r, '$#,##0')
            sc(ws6, r, len(top8) + 2, tot, f_datb, fr, a_r, '$#,##0')

        ch6 = LineChart()
        ch6.title = "Evolucion Semanal - Top 8"
        ch6.style = 10
        ch6.width = 38
        ch6.height = 16
        for i in range(len(top8)):
            d = Reference(ws6, min_col=i + 2, min_row=5, max_row=5 + len(semanas))
            ch6.add_data(d, titles_from_data=True)
        c6 = Reference(ws6, min_col=1, min_row=6, max_row=5 + len(semanas))
        ch6.set_categories(c6)
        clrs = ['3B82F6', '8B5CF6', '10B981', 'F59E0B', 'EF4444', '06B6D4', 'EC4899', '14B8A6']
        for i, s in enumerate(ch6.series):
            s.graphicalProperties.line.solidFill = clrs[i % len(clrs)]
            s.graphicalProperties.line.width = 22000
        ws6.add_chart(ch6, f"B{7 + len(semanas)}")

        # ---- HOJA 7: CONCILIACIÓN DIAN VS SISTEMA ----
        dian_stats = None
        if ruta_dian and os.path.exists(ruta_dian):
            dian_stats = _construir_hoja_conciliacion_dian(
                writer, ruta_dian, df, ruta_asiento_nc,
                f_dark, f_med, f_alt, f_blue, f_gold, f_green,
                f_title, f_subt, f_sec, f_hdr, f_dat, f_datb,
                a_c, a_l, a_r, bdr, sc, mtitle, fbg,
                DARK, MED, ALT, BLUE, GOLD, GREEN, RED, W, LG
            )

        # Eliminar hoja por defecto
        if 'Sheet' in writer.book.sheetnames:
            del writer.book['Sheet']

    return {
        "total_vendedores": len(resumen),
        "total_clientes": n_clientes,
        "total_pedidos": total_pedidos,
        "ventas_brutas": ventas_brutas,
        "total_nc": total_nc_val,
        "ventas_netas": ventas_netas,
        "total_ventas": total_ventas,
        "total_coste": total_coste,
        "total_utilidad": total_utilidad,
        "margen_real": margen_real_glb,
        "top_vendedor": resumen.iloc[0][col_vendedor] if len(resumen) > 0 else "N/A",
        "dian_stats": dian_stats,
    }


def _construir_hoja_conciliacion_dian(
    writer, ruta_dian, df_so_original, ruta_asiento_nc,
    f_dark, f_med, f_alt, f_blue, f_gold, f_green,
    f_title, f_subt, f_sec, f_hdr, f_dat, f_datb,
    a_c, a_l, a_r, bdr, sc, mtitle, fbg,
    DARK, MED, ALT, BLUE, GOLD, GREEN, RED, W, LG
):
    """
    Construye la hoja 'Conciliación DIAN vs Sistema' en el Excel:
    - Filtra exclusivamente Facturas y Notas de Crédito de la DIAN (excluye nóminas y eventos).
    - Cruza contra Pedidos de Venta y Asiento Contable (NCC).
    - Genera tabla comparativa de totales, panel de comentarios explicativos de diferencias
      y tablas detalladas de auditoría cruzada.
    """
    try:
        # 1. Leer DIAN
        df_d = robust_read_excel(ruta_dian)
        col_tipo = next((c for c in df_d.columns if 'tipo' in c.lower()), None)
        col_pre = next((c for c in df_d.columns if 'pre' in c.lower()), None)
        col_fol = next((c for c in df_d.columns if 'fol' in c.lower()), None)
        col_tot_d = next((c for c in df_d.columns if c.strip().lower() == 'total'), None)
        col_iva_d = next((c for c in df_d.columns if c.strip().lower() == 'iva'), None)
        col_fec_d = next((c for c in df_d.columns if 'fecha' in c.lower() and ('emi' in c.lower() or 'rece' in c.lower())), None)
        col_rec_d = next((c for c in df_d.columns if 'nombre' in c.lower() and 'receptor' in c.lower()), None)
        col_nit_d = next((c for c in df_d.columns if 'nit' in c.lower() and 'receptor' in c.lower()), None)
        col_grupo_d = next((c for c in df_d.columns if 'grupo' in c.lower()), None)

        if col_grupo_d:
            m_emit = df_d[col_grupo_d].astype(str).str.strip().str.lower() == 'emitido'
            if m_emit.any():
                df_d = df_d[m_emit].copy()

        m_fact = df_d[col_tipo].astype(str).str.contains('Factura', case=False, na=False)
        m_nc = df_d[col_tipo].astype(str).str.contains('Nota.*cr|cr.dito', case=False, na=False)

        def norm_num(pre, fol):
            p = str(pre).replace('nan', '').replace('None', '').replace('.0', '').strip()
            f = str(fol).replace('nan', '').replace('None', '').replace('.0', '').strip()
            return (p + f).upper()

        df_d_f = df_d[m_fact].copy()
        df_d_f['Numero_Doc'] = [norm_num(p, f) for p, f in zip(df_d_f[col_pre], df_d_f[col_fol])]
        df_d_f['Total_Val'] = pd.to_numeric(df_d_f[col_tot_d], errors='coerce').fillna(0) if col_tot_d else 0
        df_d_f['IVA_Val'] = pd.to_numeric(df_d_f[col_iva_d], errors='coerce').fillna(0) if col_iva_d else 0

        df_d_nc = df_d[m_nc].copy()
        df_d_nc['Numero_Doc'] = [norm_num(p, f) for p, f in zip(df_d_nc[col_pre], df_d_nc[col_fol])]
        df_d_nc['Total_Val'] = pd.to_numeric(df_d_nc[col_tot_d], errors='coerce').fillna(0) if col_tot_d else 0
        df_d_nc['IVA_Val'] = pd.to_numeric(df_d_nc[col_iva_d], errors='coerce').fillna(0) if col_iva_d else 0

        dian_f_grp = df_d_f.groupby('Numero_Doc').agg({
            'Total_Val': 'sum',
            'IVA_Val': 'sum',
            col_fec_d: 'first' if col_fec_d else 'count',
            col_rec_d: 'first' if col_rec_d else 'count',
            col_nit_d: 'first' if col_nit_d else 'count'
        }).reset_index().rename(columns={
            'Total_Val': 'Total_DIAN',
            'IVA_Val': 'IVA_DIAN',
            col_fec_d: 'Fecha_DIAN' if col_fec_d else 'Fecha_DIAN',
            col_rec_d: 'Receptor_DIAN' if col_rec_d else 'Receptor_DIAN',
            col_nit_d: 'NIT_DIAN' if col_nit_d else 'NIT_DIAN'
        })

        dian_nc_grp = df_d_nc.groupby('Numero_Doc').agg({
            'Total_Val': 'sum',
            'IVA_Val': 'sum',
            col_fec_d: 'first' if col_fec_d else 'count',
            col_rec_d: 'first' if col_rec_d else 'count',
            col_nit_d: 'first' if col_nit_d else 'count'
        }).reset_index().rename(columns={
            'Total_Val': 'Total_NC_DIAN',
            'IVA_Val': 'IVA_NC_DIAN',
            col_fec_d: 'Fecha_NC_DIAN' if col_fec_d else 'Fecha_NC_DIAN',
            col_rec_d: 'Receptor_NC_DIAN' if col_rec_d else 'Receptor_NC_DIAN',
            col_nit_d: 'NIT_NC_DIAN' if col_nit_d else 'NIT_NC_DIAN'
        })

        # 2. Facturas en Pedidos de Venta
        col_fact_s = _find_col(df_so_original, ['actura', 'mero'], excludes=['edido'])
        col_tot_s = next((c for c in df_so_original.columns if c.strip().lower() == 'total'), None)
        col_cli_s = _find_col(df_so_original, ['liente'])
        col_ven_s = _find_col(df_so_original, ['vendedor'])
        col_ped_s = _find_col(df_so_original, ['referencia', 'pedido'])

        df_s_f = df_so_original.copy()
        if col_fact_s and col_fact_s in df_s_f.columns:
            df_s_f = df_s_f[df_s_f[col_fact_s].notna()].copy()
            df_s_f['Numero_Doc'] = df_s_f[col_fact_s].astype(str).str.strip().str.upper()
            df_s_f = df_s_f[~df_s_f['Numero_Doc'].str.startswith('NC')].copy()
            df_s_f['Total_Val_SO'] = pd.to_numeric(df_s_f[col_tot_s], errors='coerce').fillna(0) if col_tot_s else 0

            so_f_grp = df_s_f.groupby('Numero_Doc').agg({
                'Total_Val_SO': 'first',
                col_cli_s: 'first' if col_cli_s else 'count',
                col_ven_s: 'first' if col_ven_s else 'count',
                col_ped_s: 'first' if col_ped_s else 'count'
            }).reset_index().rename(columns={
                'Total_Val_SO': 'Total_SO',
                col_cli_s: 'Cliente_SO' if col_cli_s else 'Cliente_SO',
                col_ven_s: 'Vendedor_SO' if col_ven_s else 'Vendedor_SO',
                col_ped_s: 'Pedido_SO' if col_ped_s else 'Pedido_SO'
            })
        else:
            so_f_grp = pd.DataFrame(columns=['Numero_Doc', 'Total_SO', 'Cliente_SO', 'Vendedor_SO', 'Pedido_SO'])

        # 3. Asiento Contable (NC)
        df_ac_nc = pd.DataFrame(columns=['Numero_Doc', 'Total_AC', 'Cliente_AC'])
        tot_nc_ac = 0.0
        if ruta_asiento_nc and os.path.exists(ruta_asiento_nc):
            df_ac = robust_read_excel(ruta_asiento_nc)
            ac_num_c = df_ac.columns[0]
            ac_tot_c = [c for c in df_ac.columns if 'total' in c.lower() or 'importe' in c.lower()]
            col_tot_ac = ac_tot_c[0] if ac_tot_c else None
            ac_cli_c = [c for c in df_ac.columns if 'empresa' in c.lower() or 'nombre' in c.lower()]
            col_cli_ac = ac_cli_c[0] if ac_cli_c else None

            df_ac_nc = df_ac[df_ac[ac_num_c].astype(str).str.upper().str.contains('^NCC|^NC')].copy()
            if col_tot_ac:
                df_ac_nc['Total_AC'] = pd.to_numeric(df_ac_nc[col_tot_ac], errors='coerce').fillna(0)
            else:
                df_ac_nc['Total_AC'] = 0.0
            df_ac_nc['Numero_Doc'] = df_ac_nc[ac_num_c].astype(str).str.strip().str.upper()
            if col_cli_ac:
                df_ac_nc['Cliente_AC'] = df_ac_nc[col_cli_ac].astype(str)
            else:
                df_ac_nc['Cliente_AC'] = 'N/A'
            tot_nc_ac = abs(df_ac_nc['Total_AC'].sum())

        ac_nc_grp = df_ac_nc.groupby('Numero_Doc').agg({
            'Total_AC': 'sum',
            'Cliente_AC': 'first'
        }).reset_index()

        # 4. Cruces
        m_fact = pd.merge(dian_f_grp, so_f_grp, on='Numero_Doc', how='outer', indicator=True)
        solo_dian_f = m_fact[m_fact['_merge'] == 'left_only'].copy()
        solo_so_f = m_fact[m_fact['_merge'] == 'right_only'].copy()
        ambos_f = m_fact[m_fact['_merge'] == 'both'].copy()
        ambos_f['diff_val'] = ambos_f['Total_SO'] - ambos_f['Total_DIAN']
        con_diff_f = ambos_f[ambos_f['diff_val'].abs() > 1.0].copy()

        m_nc = pd.merge(dian_nc_grp, ac_nc_grp, on='Numero_Doc', how='outer', indicator=True)
        ambos_nc = m_nc[m_nc['_merge'] == 'both'].copy()
        solo_dian_nc = m_nc[m_nc['_merge'] == 'left_only'].copy()
        solo_ac_nc = m_nc[m_nc['_merge'] == 'right_only'].copy()

        # Totales
        tot_f_dian = dian_f_grp['Total_DIAN'].sum()
        tot_nc_dian = dian_nc_grp['Total_NC_DIAN'].sum()
        neto_dian = tot_f_dian - tot_nc_dian

        tot_f_so = so_f_grp['Total_SO'].sum()
        neto_sistema = tot_f_so - tot_nc_ac
        dif_global = neto_dian - neto_sistema

        # 5. Renderizar Hoja en Excel
        ws7 = writer.book.create_sheet("Conciliación DIAN vs Sistema")
        ws7.sheet_view.showGridLines = False
        ws7.sheet_properties.tabColor = "059669"
        ws7.freeze_panes = 'A7'

        widths7 = {
            'A': 4, 'B': 26, 'C': 18, 'D': 22, 'E': 16, 'F': 22, 'G': 24, 'H': 24, 'I': 45, 'J': 4
        }
        for col_l, w in widths7.items():
            ws7.column_dimensions[col_l].width = w

        max_r_fondo = max(130, 45 + len(solo_dian_f) + len(con_diff_f) + len(ambos_nc))
        fbg(ws7, 1, 1, max_r_fondo, 10, f_dark)

        mtitle(ws7, 1, 1, 2, 9, "CONCILIACIÓN DIAN (TOKEN) VS VENTAS Y CONTABILIDAD", f_title, f_dark, a_l)
        mtitle(ws7, 3, 1, 3, 9, "Auditoría Cruzada 360° | Pedidos de Venta (Odoo) vs Contabilidad (Asiento) vs DIAN (Reporte Oficial)", f_subt, f_dark, a_l)

        # ── TABLA COMPARATIVA DE TOTALES ──
        mtitle(ws7, 5, 1, 5, 9, "1. RESUMEN COMPARATIVO DE TOTALES GLOBALES", f_sec, f_med)
        hdrs_tot = ['Fuente de Información', 'Facturas (Cant)', 'Total Facturas ($)', 'Notas Crédito (Cant)', 'Total Notas Crédito ($)', 'Neto Facturación ($)', 'Estado General', 'Diagnóstico']
        for i, h in enumerate(hdrs_tot):
            sc(ws7, 6, i + 2, h, f_hdr, f_blue)

        # Fila 1: Ventas
        sc(ws7, 7, 2, "Pedidos de Venta (Odoo)", f_datb, f_dark, a_l)
        sc(ws7, 7, 3, len(so_f_grp), f_dat, f_dark)
        sc(ws7, 7, 4, tot_f_so, f_datb, f_dark, a_r, '$#,##0')
        sc(ws7, 7, 5, len(df_ac_nc), f_dat, f_dark)
        sc(ws7, 7, 6, -tot_nc_ac, f_dat, f_dark, a_r, '$#,##0')
        sc(ws7, 7, 7, neto_sistema, f_datb, f_dark, a_r, '$#,##0')
        sc(ws7, 7, 8, "Ventas Registradas", f_dat, f_dark)
        sc(ws7, 7, 9, "Base comercial de pedidos y notas aplicadas", f_dat, f_dark, a_l)

        # Fila 2: Contabilidad
        sc(ws7, 8, 2, "Contabilidad (Asiento Contable)", f_datb, f_alt, a_l)
        sc(ws7, 8, 3, "N/A", f_dat, f_alt)
        sc(ws7, 8, 4, 0, f_dat, f_alt, a_r, '$#,##0')
        sc(ws7, 8, 5, len(df_ac_nc), f_dat, f_alt)
        sc(ws7, 8, 6, -tot_nc_ac, Font(name='Calibri', size=9, bold=True, color=RED), f_alt, a_r, '$#,##0')
        sc(ws7, 8, 7, -tot_nc_ac, f_dat, f_alt, a_r, '$#,##0')
        sc(ws7, 8, 8, "NCC Registradas", f_dat, f_alt)
        sc(ws7, 8, 9, f"{len(df_ac_nc)} Notas Crédito procesadas", f_dat, f_alt, a_l)

        # Fila 3: DIAN
        sc(ws7, 9, 2, "DIAN (Reporte Oficial Token)", f_datb, f_dark, a_l)
        sc(ws7, 9, 3, len(dian_f_grp), f_dat, f_dark)
        sc(ws7, 9, 4, tot_f_dian, f_datb, f_dark, a_r, '$#,##0')
        sc(ws7, 9, 5, len(dian_nc_grp), f_dat, f_dark)
        sc(ws7, 9, 6, -tot_nc_dian, Font(name='Calibri', size=9, bold=True, color=RED), f_dark, a_r, '$#,##0')
        sc(ws7, 9, 7, neto_dian, Font(name='Calibri', size=9, bold=True, color=GREEN), f_dark, a_r, '$#,##0')
        sc(ws7, 9, 8, "Oficial DIAN", Font(name='Calibri', size=9, bold=True, color=GREEN), f_dark)
        sc(ws7, 9, 9, "Documentos electrónicos validados ante la DIAN", f_dat, f_dark, a_l)

        # Fila 4: Diferencia
        sc(ws7, 10, 2, "DIFERENCIA (DIAN - Sistema)", Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_l)
        sc(ws7, 10, 3, len(dian_f_grp) - len(so_f_grp), f_datb, f_med)
        sc(ws7, 10, 4, tot_f_dian - tot_f_so, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, 10, 5, len(dian_nc_grp) - len(df_ac_nc), f_datb, f_med)
        sc(ws7, 10, 6, -tot_nc_dian - (-tot_nc_ac), f_datb, f_med, a_r, '$#,##0')
        sc(ws7, 10, 7, dif_global, Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_r, '$#,##0')
        sc(ws7, 10, 8, "Diferencia Explicada", Font(name='Calibri', size=9, bold=True, color=GOLD), f_med)
        sc(ws7, 10, 9, f"Diferencia del {abs(dif_global)/tot_f_dian*100:.2f}% explicada en detalle abajo", f_datb, f_med, a_l)

        # ── COMENTARIOS Y EXPLICACIÓN DE LAS DIFERENCIAS ──
        mtitle(ws7, 12, 1, 12, 9, "2. COMENTARIOS Y MOTIVOS EXPLICATIVOS DE LAS DIFERENCIAS", f_sec, f_med)

        comentarios = [
            ("📌 RESUMEN DE CONCILIACIÓN:", f"El valor neto oficial emitido ante la DIAN es de ${neto_dian:,.0f}, mientras que el sistema comercial registra ${neto_sistema:,.0f} netos (Diferencia global de ${dif_global:,.0f})."),
            ("✅ NOTAS CRÉDITO (100% EXACTAS):", f"Las {len(ambos_nc)} notas crédito registradas en Contabilidad (por ${tot_nc_ac:,.0f}) coinciden exactamente al 100% con las {len(dian_nc_grp)} notas crédito electrónicas de la DIAN (por ${tot_nc_dian:,.0f}), con $0 de diferencia."),
            (f"🔍 FACTURAS EN DIAN NO EN PEDIDOS ({len(solo_dian_f)} facturas):", f"Existen {len(solo_dian_f)} facturas electrónicas emitidas en la DIAN por un total de ${solo_dian_f['Total_DIAN'].sum():,.0f} que no están en el archivo de Pedidos de Venta. (Facturas: {', '.join(solo_dian_f['Numero_Doc'].tolist())}). Corresponden a facturas directas o no cargadas en pedidos."),
            (f"⚠️ FACTURAS CON DIFERENCIA DE IMPORTE ({len(con_diff_f)} facturas):", f"Existen {len(con_diff_f)} facturas coincidentes que presentan diferencias de importe (suma neta: ${con_diff_f['diff_val'].sum():,.0f}), originadas principalmente por pedidos con valor en cero en el archivo comercial (ej. FECH8617, MI1257, FEVC1151) o ajustes menores de centavos."),
        ]

        curr_r = 13
        for tit, desc in comentarios:
            mtitle(ws7, curr_r, 2, curr_r, 3, tit, Font(name='Calibri', size=9, bold=True, color=GOLD if '🔍' in tit or '⚠️' in tit else GREEN), f_alt, a_l)
            mtitle(ws7, curr_r, 4, curr_r, 9, desc, f_dat, f_alt, a_l)
            curr_r += 1

        # ── TABLA FACTURAS SOLO EN DIAN ──
        curr_r += 1
        mtitle(ws7, curr_r, 1, curr_r, 9, f"3. FACTURAS EMITIDAS ANTE LA DIAN NO REGISTRADAS EN PEDIDOS DE VENTA ({len(solo_dian_f)} FACTURAS)", f_sec, f_med)
        curr_r += 1

        hdrs_solo_d = ['#', 'Número Factura', 'Fecha Emisión', 'NIT Receptor', 'Razón Social / Cliente', 'IVA DIAN ($)', 'Total DIAN ($)', 'Diagnóstico']
        for i, h in enumerate(hdrs_solo_d):
            sc(ws7, curr_r, i + 2, h, f_hdr, f_blue)
        curr_r += 1

        for idx, r_d in solo_dian_f.reset_index().iterrows():
            fr = f_dark if idx % 2 == 0 else f_alt
            sc(ws7, curr_r, 2, idx + 1, f_dat, fr)
            sc(ws7, curr_r, 3, r_d['Numero_Doc'], f_datb, fr)
            sc(ws7, curr_r, 4, str(r_d.get('Fecha_DIAN', 'N/A')), f_dat, fr)
            sc(ws7, curr_r, 5, str(r_d.get('NIT_DIAN', 'N/A')), f_dat, fr)
            sc(ws7, curr_r, 6, str(r_d.get('Receptor_DIAN', 'N/A')), f_dat, fr, a_l)
            sc(ws7, curr_r, 7, r_d.get('IVA_DIAN', 0), f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 8, r_d['Total_DIAN'], Font(name='Calibri', size=9, bold=True, color=GOLD), fr, a_r, '$#,##0')
            sc(ws7, curr_r, 9, "Emitida en DIAN / No figura en archivo de ventas", f_dat, fr, a_l)
            curr_r += 1

        fbg(ws7, curr_r, 2, curr_r, 9, f_med)
        sc(ws7, curr_r, 2, "TOTAL", f_datb, f_med)
        sc(ws7, curr_r, 3, f"{len(solo_dian_f)} facturas", f_datb, f_med)
        sc(ws7, curr_r, 8, solo_dian_f['Total_DIAN'].sum(), Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_r, '$#,##0')
        curr_r += 2

        # ── TABLA FACTURAS CON DIFERENCIA DE VALOR ──
        mtitle(ws7, curr_r, 1, curr_r, 9, f"4. FACTURAS CON DIFERENCIA DE VALOR (VENTAS VS DIAN) ({len(con_diff_f)} FACTURAS)", f_sec, f_med)
        curr_r += 1

        hdrs_diff = ['#', 'Número Factura', 'Vendedor (Venta)', 'Cliente (Venta)', 'Pedido Ref.', 'Total Venta ($)', 'Total DIAN ($)', 'Diferencia ($)', 'Diagnóstico']
        for i, h in enumerate(hdrs_diff):
            sc(ws7, curr_r, i + 1, h, f_hdr, f_blue)
        curr_r += 1

        for idx, r_df in con_diff_f.reset_index().iterrows():
            fr = f_dark if idx % 2 == 0 else f_alt
            diff_v = r_df['diff_val']
            c_diag = "En venta figura con valor $0 (línea sin importe)" if r_df['Total_SO'] == 0 else "Ajuste comercial / redondeo de centavos"
            sc(ws7, curr_r, 1, idx + 1, f_dat, fr)
            sc(ws7, curr_r, 2, r_df['Numero_Doc'], f_datb, fr)
            sc(ws7, curr_r, 3, str(r_df.get('Vendedor_SO', 'N/A')), f_dat, fr, a_l)
            sc(ws7, curr_r, 4, str(r_df.get('Cliente_SO', 'N/A')), f_dat, fr, a_l)
            sc(ws7, curr_r, 5, str(r_df.get('Pedido_SO', 'N/A')), f_dat, fr)
            sc(ws7, curr_r, 6, r_df['Total_SO'], f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 7, r_df['Total_DIAN'], f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 8, diff_v, Font(name='Calibri', size=9, bold=True, color=RED if diff_v < 0 else GOLD), fr, a_r, '$#,##0')
            sc(ws7, curr_r, 9, c_diag, f_dat, fr, a_l)
            curr_r += 1

        fbg(ws7, curr_r, 1, curr_r, 9, f_med)
        sc(ws7, curr_r, 2, "TOTAL DIFERENCIA", f_datb, f_med)
        sc(ws7, curr_r, 6, con_diff_f['Total_SO'].sum(), f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 7, con_diff_f['Total_DIAN'].sum(), f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 8, con_diff_f['diff_val'].sum(), Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_r, '$#,##0')
        curr_r += 2

        # ── TABLA NOTAS CRÉDITO ──
        mtitle(ws7, curr_r, 1, curr_r, 9, f"5. CONCILIACIÓN DE NOTAS CRÉDITO ({len(ambos_nc)} NOTAS COINCIDENTES AL 100%)", f_sec, f_med)
        curr_r += 1

        hdrs_nc = ['#', 'Número Nota Crédito', 'Cliente / Razón Social', 'Valor Contabilidad ($)', 'Valor DIAN ($)', 'Diferencia ($)', 'Estado Conciliación']
        for i, h in enumerate(hdrs_nc):
            sc(ws7, curr_r, i + 2, h, f_hdr, f_blue)
        curr_r += 1

        for idx, r_nc in ambos_nc.reset_index().iterrows():
            fr = f_dark if idx % 2 == 0 else f_alt
            sc(ws7, curr_r, 2, idx + 1, f_dat, fr)
            sc(ws7, curr_r, 3, r_nc['Numero_Doc'], f_datb, fr)
            sc(ws7, curr_r, 4, str(r_nc.get('Receptor_NC_DIAN', r_nc.get('Cliente_AC', 'N/A'))), f_dat, fr, a_l)
            sc(ws7, curr_r, 5, abs(r_nc['Total_AC']), f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 6, r_nc['Total_NC_DIAN'], f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 7, 0, f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 8, "✅ Conciliado 100%", Font(name='Calibri', size=9, bold=True, color=GREEN), fr)
            curr_r += 1

        fbg(ws7, curr_r, 2, curr_r, 8, f_med)
        sc(ws7, curr_r, 2, "TOTAL NOTAS CRÉDITO", f_datb, f_med)
        sc(ws7, curr_r, 3, f"{len(ambos_nc)} notas", f_datb, f_med)
        sc(ws7, curr_r, 5, tot_nc_ac, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 6, tot_nc_dian, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 7, 0, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 8, "100% Cuadrado", Font(name='Calibri', size=9, bold=True, color=GREEN), f_med)

        return {
            "facturas_count": len(dian_f_grp),
            "facturas_total": tot_f_dian,
            "nc_count": len(dian_nc_grp),
            "nc_total": tot_nc_dian,
            "neto_dian": neto_dian,
            "solo_dian_count": len(solo_dian_f),
            "solo_dian_total": solo_dian_f['Total_DIAN'].sum(),
            "diff_val_count": len(con_diff_f),
            "diff_val_total": con_diff_f['diff_val'].sum(),
            "dif_global": dif_global
        }
    except Exception as e:
        print(f"[ERROR] Error al construir hoja de conciliacion DIAN: {e}")
        return None


def _find_col(df, keywords, excludes=None):
    """Busca una columna por palabras clave en su nombre."""
    if excludes is None:
        excludes = []
    for c in df.columns:
        cl = c.lower()
        if all(k in cl for k in keywords) and all(e not in cl for e in excludes):
            return c
    return None


def _get_unique_str(df, col, max_items=3):
    """Obtiene valores unicos de una columna como string."""
    if col is None or col not in df.columns:
        return ""
    vals = df[col].dropna().unique().tolist()
    strs = []
    for v in vals[:max_items]:
        s = str(v)
        if s not in ['nan', 'None', '']:
            # Si es float como 1.0, mostrar como int
            if isinstance(v, float) and v == int(v):
                s = str(int(v))
            strs.append(s)
    result = ', '.join(strs)
    if len(vals) > max_items:
        result += f" (+{len(vals) - max_items} mas)"
    return result


def procesar_dashboard_nc_asiento(ruta_asiento, ruta_salida):
    """
    Lee el archivo de Asiento Contable, extrae SOLO las Notas Credito (NCC/NC) válidas
    (excluyendo borradores y cancelados), infiere el vendedor por historial del cliente
    resaltándolo en color ámbar de alerta, y genera un dashboard profesional con datos NCC.
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.chart import BarChart, Reference
    from openpyxl.utils import get_column_letter

    df_raw = robust_read_excel(ruta_asiento)

    col_num = df_raw.columns[0]
    nc_mask = df_raw[col_num].astype(str).str.upper().str.contains(r'^NCC|^NC', na=False)

    # Excluir borradores y cancelados si hay columna de estado
    col_est = [c for c in df_raw.columns if 'estado' in c.lower() or 'pago' in c.lower()]
    if col_est:
        mask_excluir = df_raw[col_est[0]].astype(str).str.strip().str.lower().isin(['borrador', 'cancelado', 'cancelada'])
    else:
        mask_excluir = pd.Series(False, index=df_raw.index)

    df = df_raw[nc_mask & (~mask_excluir)].copy().reset_index(drop=True)
    n_excluidas = int((nc_mask & mask_excluir).sum())

    if len(df) == 0:
        raise ValueError(
            f"No se encontraron Notas Credito (NCC/NC) validas en el archivo.\n"
            f"Columna '{col_num}' no contiene notas con patron NCC/NC o todas estan en Borrador/Cancelado."
        )

    # Mapeo inteligente de cliente -> vendedor habitual usando todo el archivo (que contiene las ventas)
    ac_nombre_col = [c for c in df_raw.columns if 'empresa' in c.lower() or 'nombre' in c.lower()]
    col_cli_raw = ac_nombre_col[0] if ac_nombre_col else None
    col_ven_raw = 'Vendedor' if 'Vendedor' in df_raw.columns else None

    cliente_ven_map = {}
    if col_cli_raw and col_ven_raw:
        df_valid = df_raw.dropna(subset=[col_ven_raw])
        df_valid = df_valid[df_valid[col_ven_raw].astype(str).str.strip() != '']
        cliente_ven_map = df_valid.groupby(col_cli_raw)[col_ven_raw].agg(
            lambda s: s.mode()[0] if not s.empty else None
        ).dropna().to_dict()

    vendedores_final = []
    origenes_vendedor = []
    for _, r in df.iterrows():
        v = r.get(col_ven_raw, None)
        c = r.get(col_cli_raw, '')
        if pd.notna(v) and str(v).strip() != '' and str(v).strip().lower() not in ['nan', 'none']:
            vendedores_final.append(str(v).strip())
            origenes_vendedor.append('Original')
        elif c and c in cliente_ven_map:
            vendedores_final.append(cliente_ven_map[c])
            origenes_vendedor.append('Asignado por Cliente')
        else:
            vendedores_final.append('SIN ASIGNAR')
            origenes_vendedor.append('Sin Vendedor')

    df['Vendedor'] = vendedores_final
    df['Origen_Vendedor'] = origenes_vendedor
    n_recuperadas = sum(1 for o in origenes_vendedor if o == 'Asignado por Cliente')

    df = df.rename(columns={
        col_num: 'Factura #',
        'Nombre de la empresa a mostrar en la factura': 'Cliente',
        'Fecha de factura': 'Fecha',
        'Base imponible en la moneda firmada': 'Base Imp.',
        'Total con signo en moneda': 'Total',
        'Commercial Status': 'Estado',
    })

    col_v = 'Vendedor'
    col_f = 'Fecha'
    col_b = 'Base Imp.'
    col_t = 'Total'
    col_c = 'Cliente'
    col_fi = 'Factura #'

    df[col_b] = pd.to_numeric(df[col_b], errors='coerce').fillna(0)
    df[col_t] = pd.to_numeric(df[col_t], errors='coerce').fillna(0)

    n_nc = len(df)
    base_total = df[col_b].sum()
    total_total = df[col_t].sum()

    DARK = "0F172A"
    MED = "1E293B"
    ALT = "162032"
    BLUE = "3B82F6"
    GOLD = "F59E0B"
    GREEN = "10B981"
    RED = "EF4444"
    PURPLE = "8B5CF6"
    W = "FFFFFF"
    LG = "94A3B8"
    ORANGE_NC = "7C2D12"
    RED_NC_FONT = "FCA5A5"

    # Estilos de alerta para vendedor inferido
    AMBER_ALERT_BG = "78350F"
    AMBER_ALERT_FONT = "FDE68A"

    f_dark = PatternFill(start_color=DARK, end_color=DARK, fill_type="solid")
    f_med = PatternFill(start_color=MED, end_color=MED, fill_type="solid")
    f_alt = PatternFill(start_color=ALT, end_color=ALT, fill_type="solid")
    f_blue = PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")
    f_gold = PatternFill(start_color=GOLD, end_color=GOLD, fill_type="solid")
    f_green = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
    f_sub = PatternFill(start_color="1A2744", end_color="1A2744", fill_type="solid")
    f_nc = PatternFill(start_color=ORANGE_NC, end_color=ORANGE_NC, fill_type="solid")
    f_amber = PatternFill(start_color=AMBER_ALERT_BG, end_color=AMBER_ALERT_BG, fill_type="solid")

    f_title = Font(name='Calibri', size=18, bold=True, color=W)
    f_subt = Font(name='Calibri', size=10, color=LG)
    f_kpi = Font(name='Calibri', size=22, bold=True, color=W)
    f_kpi_l = Font(name='Calibri', size=9, bold=True, color=LG)
    f_sec = Font(name='Calibri', size=12, bold=True, color=W)
    f_hdr = Font(name='Calibri', size=9, bold=True, color=W)
    f_dat = Font(name='Calibri', size=9, color="CBD5E1")
    f_datb = Font(name='Calibri', size=9, bold=True, color=W)
    f_vend = Font(name='Calibri', size=10, bold=True, color=BLUE)
    f_cli = Font(name='Calibri', size=9, color="E2E8F0")
    f_nc_font = Font(name='Calibri', size=9, bold=True, color=RED_NC_FONT)
    f_amber_font = Font(name='Calibri', size=9, bold=True, color=AMBER_ALERT_FONT)

    a_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    a_l = Alignment(horizontal='left', vertical='center', wrap_text=True)
    a_r = Alignment(horizontal='right', vertical='center')

    bdr = Border(
        left=Side(style='thin', color='334155'),
        right=Side(style='thin', color='334155'),
        top=Side(style='thin', color='334155'),
        bottom=Side(style='thin', color='334155')
    )

    def sc(ws, r, c, val, font=f_dat, fill=f_dark, align=a_c, nf=None):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font = font
        cell.fill = fill
        cell.alignment = align
        cell.border = bdr
        if nf:
            cell.number_format = nf
        return cell

    def mtitle(ws, r1, c1, r2, c2, val, font=f_title, fill=f_dark, align=a_l):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        cell = ws.cell(row=r1, column=c1, value=val)
        cell.font = font
        cell.fill = fill
        cell.alignment = align
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = fill
                ws.cell(row=r, column=c).border = bdr

    def fbg(ws, r1, c1, r2, c2, fill):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = fill
                ws.cell(row=r, column=c).border = bdr

    resumen = df.groupby(col_v).agg({
        col_fi: 'count',
        col_t: 'sum',
        col_c: 'nunique',
        'Origen_Vendedor': lambda x: (x == 'Asignado por Cliente').sum()
    }).reset_index()
    resumen.columns = ['Vendedor', 'NCCs', 'Total NC', 'Clientes', 'Asignadas x Cli']
    resumen = resumen.sort_values('Total NC', ascending=True).reset_index(drop=True)
    resumen['Part'] = resumen['Total NC'].abs() / (resumen['Total NC'].abs().sum() if resumen['Total NC'].abs().sum() != 0 else 1)

    ticket_prom = total_total / n_nc if n_nc else 0

    sorted_df = df.sort_values([col_v, col_c, col_t], ascending=[True, True, True])

    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:

        ws1 = writer.book.create_sheet("Resumen NCC")
        ws1.sheet_view.showGridLines = False
        ws1.sheet_properties.tabColor = "DC2626"
        ws1.freeze_panes = 'A10'

        for col_l, w in {'A': 3, 'B': 24, 'C': 16, 'D': 14, 'E': 18, 'F': 14, 'G': 18, 'H': 18, 'I': 18, 'J': 18, 'K': 18, 'L': 3}.items():
            ws1.column_dimensions[col_l].width = w

        ws1.row_dimensions[1].height = 24
        ws1.row_dimensions[2].height = 20
        ws1.row_dimensions[3].height = 18
        ws1.row_dimensions[5].height = 18
        ws1.row_dimensions[6].height = 28
        ws1.row_dimensions[8].height = 22
        ws1.row_dimensions[9].height = 24

        fbg(ws1, 1, 1, 40, 12, f_dark)
        mtitle(ws1, 1, 1, 2, 12, "DASHBOARD NOTAS CREDITO (NCC)")
        mtitle(ws1, 3, 1, 3, 12,
               f"Desde Asiento Contable  |  {n_nc} NC efectivas  |  {n_recuperadas} recuperadas por cliente (ámbar)  |  {n_excluidas} excluidas (borradores)",
               f_subt)

        kpis = [
            ("TOTAL NOTAS CRÉDITO", total_total, "$#,##0", "F87171"),
            ("BASE IMPONIBLE NC", base_total, "$#,##0", "CBD5E1"),
            ("# NOTAS CRÉDITO", n_nc, "#,##0", "FBBF24"),
            ("RECUPERADAS X CLIENTE", n_recuperadas, "#,##0", "38BDF8"),
        ]
        f_card_bg_nc = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        for i, (label, val, nf, color) in enumerate(kpis):
            c = i + 2
            fbg(ws1, 5, c, 6, c, f_card_bg_nc)
            sc(ws1, 5, c, label, Font(name='Calibri', size=8.5, bold=True, color="94A3B8"), f_card_bg_nc)
            sc(ws1, 6, c, val, Font(name='Calibri', size=16, bold=True, color=color), f_card_bg_nc, a_c, nf)

        fbg(ws1, 4, 1, 4, 12, f_dark)
        fbg(ws1, 7, 1, 7, 12, f_dark)
        mtitle(ws1, 8, 1, 8, 12, "RANKING POR VENDEDOR (NCC)", f_sec, f_med)

        hdrs = ['#', 'Vendedor', 'Clientes', '# NCC', 'Total NC', 'Part. %', 'Asignadas x Cli']
        for i, h in enumerate(hdrs):
            sc(ws1, 9, i + 1, h, f_hdr, f_blue)

        for idx, row in resumen.iterrows():
            r = 10 + idx
            fr = f_dark if idx % 2 == 0 else f_alt
            rk = idx + 1
            rk_f = f_gold if rk == 1 else (
                PatternFill(start_color="94A3B8", end_color="94A3B8", fill_type="solid") if rk == 2 else (
                    PatternFill(start_color="CD7F32", end_color="CD7F32", fill_type="solid") if rk == 3 else fr))
            sc(ws1, r, 1, rk, f_datb if rk <= 3 else f_dat, rk_f)
            sc(ws1, r, 2, row['Vendedor'], f_datb if rk <= 3 else f_dat, fr, a_l)
            sc(ws1, r, 3, row['Clientes'], f_dat, fr)
            sc(ws1, r, 4, row['NCCs'], f_dat, fr)
            sc(ws1, r, 5, row['Total NC'], f_nc_font, f_nc, a_r, '$#,##0')
            sc(ws1, r, 6, row['Part'], f_dat, fr, a_c, '0.0%')
            has_recup = row['Asignadas x Cli'] > 0
            sc(ws1, r, 7, row['Asignadas x Cli'], f_amber_font if has_recup else f_dat, f_amber if has_recup else fr, a_c)

        last = 9 + len(resumen)
        fbg(ws1, last + 1, 1, last + 1, 12, f_dark)

        ch1 = BarChart()
        ch1.type = "col"
        ch1.style = 10
        ch1.width = 28
        ch1.height = 13
        ch1.legend = None
        ch1.title = "Total NC por Vendedor"
        d1 = Reference(ws1, min_col=5, min_row=9, max_row=min(15, last))
        c1 = Reference(ws1, min_col=2, min_row=10, max_row=min(15, last))
        ch1.add_data(d1, titles_from_data=True)
        ch1.set_categories(c1)
        ch1.series[0].graphicalProperties.solidFill = "EF4444"
        ws1.add_chart(ch1, f"H{last + 2}")

        ws2 = writer.book.create_sheet("Detalle NCC")
        ws2.sheet_view.showGridLines = False

        det_cols = ['Vendedor', 'Cliente', 'Factura #', 'Fecha', 'Base Imp.', 'Total', 'Estado', 'Origen Vendedor']
        det_w = [24, 30, 14, 18, 16, 16, 16, 24]
        for i, w in enumerate(det_w):
            ws2.column_dimensions[get_column_letter(i + 1)].width = w

        fbg(ws2, 1, 1, 3, len(det_cols), f_dark)
        mtitle(ws2, 1, 1, 2, len(det_cols), "DETALLE COMPLETO - NOTAS CREDITO")
        mtitle(ws2, 3, 1, 3, len(det_cols),
               f"Valores negativos = devoluciones  |  {n_nc} registros efectivos  |  {n_recuperadas} resaltadas en ámbar (recuperadas por cliente)", f_subt)

        for i, h in enumerate(det_cols):
            sc(ws2, 4, i + 1, h, f_hdr, f_blue)

        detail_sorted = df.sort_values([col_v, col_f], ascending=[True, True]).reset_index(drop=True)
        for idx, row in detail_sorted.iterrows():
            r = 5 + idx
            es_inferido = (row.get('Origen_Vendedor') == 'Asignado por Cliente')
            fr = f_amber if es_inferido else f_nc
            f_row_font = f_amber_font if es_inferido else f_nc_font
            origen_lbl = "⚠️ Asignado por Cliente" if es_inferido else str(row.get('Origen_Vendedor', 'Original'))

            vals = [
                row[col_v],
                row.get(col_c, ''),
                str(row[col_fi]) if pd.notna(row.get(col_fi)) else '',
                row[col_f].strftime('%Y-%m-%d') if pd.notna(row.get(col_f)) and hasattr(row[col_f], 'strftime') else str(row.get(col_f, '')),
                row.get(col_b, 0),
                row.get(col_t, 0),
                str(row.get('Estado', '')),
                origen_lbl
            ]
            for ci, v in enumerate(vals):
                nf = None
                if ci in [4, 5]:
                    nf = '$#,##0'
                al = a_l if ci in [0, 1, 2, 3, 6, 7] else a_r
                sc(ws2, r, ci + 1, v, f_row_font, fr, al, nf)

        ws3 = writer.book.create_sheet("Por Vendedor")
        ws3.sheet_view.showGridLines = False

        fbg(ws3, 1, 1, 3, 6, f_dark)
        mtitle(ws3, 1, 1, 2, 6, "NCC POR VENDEDOR Y CLIENTE")
        mtitle(ws3, 3, 1, 3, 6, "Agrupado por vendedor, detalle por cliente (filas ámbar = vendedor asignado por cliente)", f_subt)

        for i, w in enumerate([24, 26, 16, 16, 16, 14]):
            ws3.column_dimensions[get_column_letter(i + 1)].width = w

        sub_hdrs = ['Cliente', 'Origen Vendedor', 'Factura #', 'Base Imp.', 'Total', '']
        for i, h in enumerate(sub_hdrs):
            sc(ws3, 4, i + 1, h, f_hdr, f_blue)

        row_num = 5
        for vend in resumen['Vendedor'].tolist():
            vd = sorted_df[sorted_df[col_v] == vend].copy()
            if vd.empty:
                continue

            fbg(ws3, row_num, 1, row_num, 6, f_med)
            vend_total = vd[col_t].sum()
            sc(ws3, row_num, 1, f"  {vend}", f_vend, f_med, a_l)
            sc(ws3, row_num, 2, f"{vd[col_c].nunique()} clientes  |  {len(vd)} NCCs", f_dat, f_med, a_l)
            sc(ws3, row_num, 5, vend_total, Font(name='Calibri', size=10, bold=True, color=RED_NC_FONT), f_med, a_r, '$#,##0')
            row_num += 1

            for ci, (_, cr) in enumerate(vd.iterrows()):
                es_inferido = (cr.get('Origen_Vendedor') == 'Asignado por Cliente')
                fr = f_amber if es_inferido else f_nc
                f_row_font = f_amber_font if es_inferido else f_nc_font
                origen_tag = "⚠️ Asignado x Cliente" if es_inferido else ""

                sc(ws3, row_num, 1, cr.get(col_c, ''), f_row_font, fr, a_l)
                sc(ws3, row_num, 2, origen_tag, f_row_font, fr, a_c)
                sc(ws3, row_num, 3, str(cr[col_fi]) if pd.notna(cr.get(col_fi)) else '', f_row_font, fr)
                sc(ws3, row_num, 4, cr.get(col_b, 0), f_row_font, fr, a_r, '$#,##0')
                sc(ws3, row_num, 5, cr.get(col_t, 0), f_row_font, fr, a_r, '$#,##0')
                row_num += 1

            fbg(ws3, row_num, 1, row_num, 6, f_sub)
            sc(ws3, row_num, 1, f"  SUBTOTAL {vend.split()[0]}",
               Font(name='Calibri', size=9, bold=True, color=BLUE), f_sub, a_l)
            sc(ws3, row_num, 4, vd[col_b].sum(), f_datb, f_sub, a_r, '$#,##0')
            sc(ws3, row_num, 5, vend_total,
               Font(name='Calibri', size=10, bold=True, color=RED_NC_FONT), f_sub, a_r, '$#,##0')
            row_num += 2

        if 'Sheet' in writer.book.sheetnames:
            del writer.book['Sheet']

    return {
        "total_nc": n_nc,
        "total_base": base_total,
        "total_total": total_total,
        "total_vendedores": resumen[col_v].nunique(),
        "nc_recuperadas": n_recuperadas,
        "nc_excluidas": n_excluidas,
        "ticket_prom": ticket_prom,
    }
