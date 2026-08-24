import pandas as pd
import numpy as np
import os
import traceback
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


def procesar_dashboard_vendedores(ruta_archivo, ruta_salida, ruta_asiento_nc=None):
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
        df_nc_ac = df_ac[ac_nc_mask].copy().reset_index(drop=True)

        if len(df_nc_ac) > 0:
            df = df[df['Es_Nota_Credito'] == False].copy()

            ncc_rows = pd.DataFrame()
            ncc_rows[col_fact_num] = df_nc_ac[ac_num_col].values
            ncc_rows[col_vendedor] = df_nc_ac['Vendedor'].values

            ac_nombre_col = [c for c in df_nc_ac.columns if 'empresa' in c.lower() or 'nombre' in c.lower()]
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

            ac_total_col = [c for c in df_nc_ac.columns if 'total' in c.lower() and 'signo' in c.lower()]
            if not ac_total_col:
                ac_total_col = [c for c in df_nc_ac.columns if 'total' in c.lower()]
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
                ncc_rows[col_coste] = 0
            if col_alm:
                ncc_rows[col_alm] = 'NCC'
            if col_estado:
                ac_estado_col = [c for c in df_nc_ac.columns if 'estado' in c.lower() or 'pago' in c.lower()]
                if ac_estado_col:
                    ncc_rows[col_estado] = df_nc_ac[ac_estado_col[0]].values
                else:
                    ncc_rows[col_estado] = 'Nota Credito'
            if col_pedido:
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
    DARK = "0F172A"
    MED = "1E293B"
    ALT = "162032"
    BLUE = "3B82F6"
    GOLD = "F59E0B"
    GREEN = "10B981"
    RED = "EF4444"
    ORANGE_NC = "7C2D12"
    RED_NC_FONT = "FCA5A5"
    PURPLE = "8B5CF6"
    CYAN = "06B6D4"
    W = "FFFFFF"
    LG = "94A3B8"

    f_dark = PatternFill(start_color=DARK, end_color=DARK, fill_type="solid")
    f_med = PatternFill(start_color=MED, end_color=MED, fill_type="solid")
    f_alt = PatternFill(start_color=ALT, end_color=ALT, fill_type="solid")
    f_blue = PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")
    f_gold = PatternFill(start_color=GOLD, end_color=GOLD, fill_type="solid")
    f_green = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
    f_sub = PatternFill(start_color="1A2744", end_color="1A2744", fill_type="solid")
    f_nc = PatternFill(start_color=ORANGE_NC, end_color=ORANGE_NC, fill_type="solid")
    f_nc_font = Font(name='Calibri', size=9, bold=True, color=RED_NC_FONT)

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

    total_ventas = df[col_total].sum()
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

        widths1 = {'A': 3, 'B': 24, 'C': 18, 'D': 18, 'E': 18, 'F': 18, 'G': 18, 'H': 18, 'I': 18, 'J': 18, 'K': 18, 'L': 18, 'M': 3}
        for col_l, w in widths1.items():
            ws1.column_dimensions[col_l].width = w

        fbg(ws1, 1, 1, 50, 13, f_dark)
        mtitle(ws1, 1, 1, 2, 13, "DASHBOARD DE VENDEDORES")
        mtitle(ws1, 3, 1, 3, 13,
               f"Panel de Control Ejecutivo  |  {total_pedidos} registros  |  {n_clientes} clientes  |  Periodo Junio 2026",
               f_subt)

        kpis = [
            ("Ventas Totales", total_ventas, "$#,##0", BLUE),
            ("Coste Total", total_coste, "$#,##0", RED),
            ("Utilidad", total_utilidad, "$#,##0", GREEN),
            ("Margen Real", margen_real_glb, "0.0%", GREEN),
            ("Total Pedidos", total_pedidos, "#,##0", GOLD),
            ("Ticket Promedio", ticket_prom, "$#,##0", PURPLE),
        ]
        for i, (label, val, nf, color) in enumerate(kpis):
            c = i + 2
            fk = PatternFill(start_color=color + "22", end_color=color + "22", fill_type="solid")
            fbg(ws1, 5, c, 5, c, fk)
            sc(ws1, 5, c, label, f_kpi_l, fk)
            sc(ws1, 6, c, val, Font(name='Calibri', size=20, bold=True, color=color), fk, a_c, nf)

        fbg(ws1, 4, 1, 4, 13, f_dark)
        fbg(ws1, 7, 1, 7, 13, f_dark)
        mtitle(ws1, 8, 1, 8, 13, "RANKING DE VENDEDORES", f_sec, f_med)

        hdrs = ['#', 'Vendedor', 'Clientes', 'Pedidos', 'Ventas Totales',
                'Coste Total', 'Utilidad', 'Margen Real', 'Descuento %', 'Ticket Prom.', 'Part. %']
        for i, h in enumerate(hdrs):
            sc(ws1, 9, i + 1, h, f_hdr, f_blue)

        for idx, row in resumen.iterrows():
            r = 10 + idx
            fr = f_dark if idx % 2 == 0 else f_alt
            rk = idx + 1
            rk_f = f_gold if rk == 1 else (
                PatternFill(start_color="94A3B8", end_color="94A3B8", fill_type="solid") if rk == 2 else (
                    PatternFill(start_color="CD7F32", end_color="CD7F32", fill_type="solid") if rk == 3 else fr))
            tk = row['Ventas'] / row['Pedidos'] if row['Pedidos'] > 0 else 0
            margen_real_v = row['Utilidad'] / row['Ventas'] if row['Ventas'] > 0 else 0
            sc(ws1, r, 1, rk, f_datb if rk <= 3 else f_dat, rk_f)
            sc(ws1, r, 2, row[col_vendedor], f_datb if rk <= 3 else f_dat, fr, a_l)
            sc(ws1, r, 3, row['Clientes'], f_dat, fr)
            sc(ws1, r, 4, row['Pedidos'], f_dat, fr)
            sc(ws1, r, 5, row['Ventas'], f_datb if rk <= 3 else f_dat, fr, a_r, '$#,##0')
            sc(ws1, r, 6, row['Coste'], f_dat, fr, a_r, '$#,##0')
            sc(ws1, r, 7, row['Utilidad'], f_dat, fr, a_r, '$#,##0')
            sc(ws1, r, 8, margen_real_v, f_dat, fr, a_c, '0.0%')
            sc(ws1, r, 9, row['Descuento'], f_dat, fr, a_c, '0.0%')
            sc(ws1, r, 10, tk, f_dat, fr, a_r, '$#,##0')
            sc(ws1, r, 11, row['Part'], f_dat, fr, a_c, '0.0%')

        last = 9 + len(resumen)
        fbg(ws1, last + 1, 1, last + 1, 13, f_dark)

        ch1 = BarChart()
        ch1.type = "col"
        ch1.style = 10
        ch1.width = 30
        ch1.height = 14
        ch1.legend = None
        ch1.title = "Ventas vs Coste por Vendedor (Top 12)"
        ch1.y_axis.title = "Monto ($)"
        d1a = Reference(ws1, min_col=5, min_row=9, max_row=min(21, last))
        d1b = Reference(ws1, min_col=6, min_row=9, max_row=min(21, last))
        c1 = Reference(ws1, min_col=2, min_row=10, max_row=min(21, last))
        ch1.add_data(d1a, titles_from_data=True)
        ch1.add_data(d1b, titles_from_data=True)
        ch1.set_categories(c1)
        ch1.series[0].graphicalProperties.solidFill = "3B82F6"
        ch1.series[1].graphicalProperties.solidFill = "EF4444"
        ws1.add_chart(ch1, f"B{last + 2}")

        ch2 = PieChart()
        ch2.title = "Participacion en Ventas"
        ch2.style = 10
        ch2.width = 16
        ch2.height = 14
        d2 = Reference(ws1, min_col=11, min_row=9, max_row=min(17, last))
        c2 = Reference(ws1, min_col=2, min_row=10, max_row=min(17, last))
        ch2.add_data(d2, titles_from_data=True)
        ch2.set_categories(c2)
        ch2.dataLabels = DataLabelList()
        ch2.dataLabels.showPercent = True
        ch2.dataLabels.showCatName = False
        ws1.add_chart(ch2, f"E{last + 2}")

        # ---- HOJA 2: CLIENTES POR VENDEDOR ----
        ws2 = writer.book.create_sheet("Clientes por Vendedor")
        ws2.sheet_view.showGridLines = False

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

        full_cols = [col_vendedor, col_cliente, col_alm, col_pedido, col_fecha,
                     col_fact_num, col_prod, col_base, col_imp, col_total,
                     col_coste, col_margen, col_desc, col_estado]
        full_labels = ['Vendedor', 'Cliente', 'Almacen', 'Ref. Pedido', 'Fecha',
                       'Factura #', 'Producto', 'Base Imp.', 'Impuestos', 'Total',
                       'Coste', 'Margen %', 'Descuento %', 'Estado']
        full_w = [24, 28, 20, 14, 18, 14, 45, 16, 14, 16, 16, 10, 12, 16]

        for i, w in enumerate(full_w):
            ws3.column_dimensions[get_column_letter(i + 1)].width = w

        fbg(ws3, 1, 1, 3, len(full_cols), f_dark)
        mtitle(ws3, 1, 1, 2, len(full_cols), "DETALLE COMPLETO - TODOS LOS PEDIDOS")
        mtitle(ws3, 3, 1, 3, len(full_cols),
               f"Ordenado por Vendedor  |  {len(df)} registros  |  {n_nc} Notas Credito (rojo)", f_subt)

        for i, h in enumerate(full_labels):
            sc(ws3, 4, i + 1, h, f_hdr, f_blue)

        detail_sorted = df.sort_values([col_vendedor, col_fecha], ascending=[True, True]).reset_index(drop=True)
        for idx, row in detail_sorted.iterrows():
            r = 5 + idx
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

        # ---- HOJA 4: ALMACENES ----
        ws4 = writer.book.create_sheet("Por Almacen")
        ws4.sheet_view.showGridLines = False

        alm_pivot = df.groupby([col_vendedor, col_alm]).size().unstack(fill_value=0)
        alm_cols = list(alm_pivot.columns)

        ws4.column_dimensions['A'].width = 24
        for i in range(len(alm_cols)):
            ws4.column_dimensions[get_column_letter(i + 2)].width = 18
        ws4.column_dimensions[get_column_letter(len(alm_cols) + 2)].width = 16

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

        for i, w in enumerate([3, 55, 14, 20, 20, 20, 12, 3]):
            ws5.column_dimensions[get_column_letter(i + 1)].width = w

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
            fr = f_dark if idx % 2 == 0 else f_alt
            rf = f_gold if idx == 0 else (
                PatternFill(start_color="94A3B8", end_color="94A3B8", fill_type="solid") if idx == 1 else (
                    PatternFill(start_color="CD7F32", end_color="CD7F32", fill_type="solid") if idx == 2 else fr))
            sc(ws5, r, 1, idx + 1, f_datb if idx < 3 else f_dat, rf)
            sc(ws5, r, 2, str(row['Producto'])[:80], f_datb if idx < 3 else f_dat, fr, a_l)
            sc(ws5, r, 3, row['Cantidad'], f_dat, fr)
            sc(ws5, r, 4, row['Ventas'], f_dat, fr, a_r, '$#,##0')
            sc(ws5, r, 5, row['Coste'], f_dat, fr, a_r, '$#,##0')
            sc(ws5, r, 6, row['Utilidad'],
               Font(name='Calibri', size=9, bold=True, color=GREEN if row['Utilidad'] >= 0 else RED), fr, a_r, '$#,##0')
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
        ch5.series[0].graphicalProperties.solidFill = "3B82F6"
        ch5.series[1].graphicalProperties.solidFill = "EF4444"
        ws5.add_chart(ch5, "B17")

        # ---- HOJA 6: EVOLUCION SEMANAL ----
        ws6 = writer.book.create_sheet("Evolucion Semanal")
        ws6.sheet_view.showGridLines = False

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

        # Eliminar hoja por defecto
        if 'Sheet' in writer.book.sheetnames:
            del writer.book['Sheet']

    return {
        "total_vendedores": len(resumen),
        "total_clientes": n_clientes,
        "total_pedidos": total_pedidos,
        "total_ventas": total_ventas,
        "total_coste": total_coste,
        "total_utilidad": total_utilidad,
        "margen_real": margen_real_glb,
        "top_vendedor": resumen.iloc[0][col_vendedor] if len(resumen) > 0 else "N/A",
    }


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
    Lee el archivo de Asiento Contable, extrae SOLO las Notas Credito (NCC)
    y genera un dashboard igual al de vendedores pero unicamente con datos NCC.
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.utils import get_column_letter

    df_raw = robust_read_excel(ruta_asiento)

    col_num = df_raw.columns[0]
    nc_mask = df_raw[col_num].astype(str).str.upper().str.contains(r'^NCC|^NC', na=False)
    df = df_raw[nc_mask].copy()

    if len(df) == 0:
        raise ValueError(
            f"No se encontraron Notas Credito (NCC/NC) en el archivo.\n"
            f"Columna '{col_num}' no contiene valores con patron NCC/NC."
        )

    df = df.rename(columns={
        col_num: 'Factura #',
        'Nombre de la empresa a mostrar en la factura': 'Cliente',
        'Vendedor': 'Vendedor',
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

    df = df.dropna(subset=[col_v])

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

    f_dark = PatternFill(start_color=DARK, end_color=DARK, fill_type="solid")
    f_med = PatternFill(start_color=MED, end_color=MED, fill_type="solid")
    f_alt = PatternFill(start_color=ALT, end_color=ALT, fill_type="solid")
    f_blue = PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")
    f_gold = PatternFill(start_color=GOLD, end_color=GOLD, fill_type="solid")
    f_green = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
    f_sub = PatternFill(start_color="1A2744", end_color="1A2744", fill_type="solid")
    f_nc = PatternFill(start_color=ORANGE_NC, end_color=ORANGE_NC, fill_type="solid")

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
        col_c: 'nunique'
    }).reset_index()
    resumen.columns = ['Vendedor', 'NCCs', 'Total NC', 'Clientes']
    resumen = resumen.sort_values('Total NC', ascending=False).reset_index(drop=True)
    resumen['Part'] = resumen['Total NC'].abs() / resumen['Total NC'].abs().sum()

    ticket_prom = total_total / n_nc if n_nc else 0

    sorted_df = df.sort_values([col_v, col_c, col_t], ascending=[True, True, False])

    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:

        ws1 = writer.book.create_sheet("Resumen NCC")
        ws1.sheet_view.showGridLines = False

        for col_l, w in {'A': 3, 'B': 24, 'C': 18, 'D': 18, 'E': 18, 'F': 18, 'G': 18, 'H': 18, 'I': 18, 'J': 18, 'K': 18, 'L': 3}.items():
            ws1.column_dimensions[col_l].width = w

        fbg(ws1, 1, 1, 40, 12, f_dark)
        mtitle(ws1, 1, 1, 2, 12, "DASHBOARD NOTAS CREDITO (NCC)")
        mtitle(ws1, 3, 1, 3, 12,
               f"Desde Asiento Contable  |  {n_nc} Notas Credito  |  {resumen[col_v].nunique()} vendedores",
               f_subt)

        kpis = [
            ("Total NC", total_total, "$#,##0", RED),
            ("Base NC", base_total, "$#,##0", PURPLE),
            ("# Notas Credito", n_nc, "#,##0", GOLD),
            ("Ticket Promedio", ticket_prom, "$#,##0", BLUE),
        ]
        for i, (label, val, nf, color) in enumerate(kpis):
            c = i + 2
            fk = PatternFill(start_color=color + "22", end_color=color + "22", fill_type="solid")
            fbg(ws1, 5, c, 5, c, fk)
            sc(ws1, 5, c, label, f_kpi_l, fk)
            sc(ws1, 6, c, val, Font(name='Calibri', size=20, bold=True, color=color), fk, a_c, nf)

        fbg(ws1, 4, 1, 4, 12, f_dark)
        fbg(ws1, 7, 1, 7, 12, f_dark)
        mtitle(ws1, 8, 1, 8, 12, "RANKING POR VENDEDOR (NCC)", f_sec, f_med)

        hdrs = ['#', 'Vendedor', 'Clientes', '# NCC', 'Total NC', 'Part. %']
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
        ws1.add_chart(ch1, f"G{last + 2}")

        ws2 = writer.book.create_sheet("Detalle NCC")
        ws2.sheet_view.showGridLines = False

        det_cols = ['Vendedor', 'Cliente', 'Factura #', 'Fecha', 'Base Imp.', 'Total', 'Estado']
        det_w = [24, 30, 14, 18, 16, 16, 16]
        for i, w in enumerate(det_w):
            ws2.column_dimensions[get_column_letter(i + 1)].width = w

        fbg(ws2, 1, 1, 3, len(det_cols), f_dark)
        mtitle(ws2, 1, 1, 2, len(det_cols), "DETALLE COMPLETO - NOTAS CREDITO")
        mtitle(ws2, 3, 1, 3, len(det_cols),
               f"Valores negativos = devoluciones  |  {n_nc} registros", f_subt)

        for i, h in enumerate(det_cols):
            sc(ws2, 4, i + 1, h, f_hdr, f_blue)

        detail_sorted = df.sort_values([col_v, col_f], ascending=[True, True]).reset_index(drop=True)
        for idx, row in detail_sorted.iterrows():
            r = 5 + idx
            fr = f_nc
            vals = [
                row[col_v],
                row.get(col_c, ''),
                str(row[col_fi]) if pd.notna(row.get(col_fi)) else '',
                row[col_f].strftime('%Y-%m-%d') if pd.notna(row.get(col_f)) and hasattr(row[col_f], 'strftime') else str(row.get(col_f, '')),
                row.get(col_b, 0),
                row.get(col_t, 0),
                str(row.get('Estado', ''))
            ]
            for ci, v in enumerate(vals):
                nf = None
                if ci in [4, 5]:
                    nf = '$#,##0'
                al = a_l if ci in [0, 1, 2, 3, 6] else a_r
                sc(ws2, r, ci + 1, v, f_nc_font, fr, al, nf)

        ws3 = writer.book.create_sheet("Por Vendedor")
        ws3.sheet_view.showGridLines = False

        fbg(ws3, 1, 1, 3, 6, f_dark)
        mtitle(ws3, 1, 1, 2, 6, "NCC POR VENDEDOR Y CLIENTE")
        mtitle(ws3, 3, 1, 3, 6, "Agrupado por vendedor, detalle por cliente", f_subt)

        for i, w in enumerate([24, 30, 20, 16, 16, 14]):
            ws3.column_dimensions[get_column_letter(i + 1)].width = w

        sub_hdrs = ['Vendedor', 'Cliente', 'Factura #', 'Base Imp.', 'Total', '']
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
                fr = f_nc
                sc(ws3, row_num, 1, cr.get(col_c, ''), f_nc_font, fr, a_l)
                sc(ws3, row_num, 2, '', f_nc_font, fr)
                sc(ws3, row_num, 3, str(cr[col_fi]) if pd.notna(cr.get(col_fi)) else '', f_nc_font, fr)
                sc(ws3, row_num, 4, cr.get(col_b, 0), f_nc_font, fr, a_r, '$#,##0')
                sc(ws3, row_num, 5, cr.get(col_t, 0), f_nc_font, fr, a_r, '$#,##0')
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
        "ticket_prom": ticket_prom,
    }
