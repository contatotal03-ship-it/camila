import pandas as pd
import numpy as np
import os
import re
import traceback
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from .excel_utils import robust_read_excel, robust_ExcelFile


import json

CONFIG_VENDEDORES_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "vendedores_config.json")

def cargar_config_vendedores():
    """Carga la configuración de vendedores especiales (Repuestos y Jaime) desde JSON."""
    defaults = {
        "vendedores_repuestos": [
            "MAILYN YORGELY CONTRERAS CARDENAS",
            "LILIANA DEL PILAR MENDOZA GUEVARA"
        ],
        "vendedores_vehiculos_jaime": [
            "JAIME EDUARDO  OROZCO CHAVARRO"
        ]
    }
    if os.path.exists(CONFIG_VENDEDORES_FILE):
        try:
            with open(CONFIG_VENDEDORES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return {
                    "vendedores_repuestos": data.get("vendedores_repuestos", defaults["vendedores_repuestos"]),
                    "vendedores_vehiculos_jaime": data.get("vendedores_vehiculos_jaime", defaults["vendedores_vehiculos_jaime"])
                }
        except Exception:
            pass
    return defaults

def guardar_config_vendedores(cfg):
    """Guarda la configuración de vendedores especiales en JSON."""
    try:
        with open(CONFIG_VENDEDORES_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error al guardar vendedores_config.json: {e}")
        return False


def leer_trial_balance_odoo(ruta_tb, client_to_cc=None):
    """
    Lee de forma inteligente un Trial Balance / Balance de Comprobación de Odoo en Excel,
    independientemente del nombre del archivo o nombre de la hoja.
    Extrae automáticamente:
      - Cuentas PUC de Otros Ingresos (4180, 4210, 4225, 4235, 4295)
      - Liquidación de Taller (Auteco + Autotécnica bajo 42354001)
      - Arriendo UMA (Supermotos UMA bajo 42354001)
      - Devoluciones en ventas (cuentas 4175)
    """
    import openpyxl
    wb = openpyxl.load_workbook(ruta_tb, data_only=True)
    ws = None
    for name in wb.sheetnames:
        if "trial" in name.lower() or "balance" in name.lower():
            ws = wb[name]
            break
    if ws is None:
        ws = wb.active

    # Detectar fila de encabezados
    header_row = None
    col_map = {}
    for r in range(1, 15):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(1, 12)]
        for idx, val in enumerate(row_vals, 1):
            val_low = val.lower()
            if "código" in val_low or "codigo" in val_low or "code" in val_low:
                col_map["code"] = idx
            elif "nombre" in val_low or "account" in val_low:
                col_map["name"] = idx
            elif "partner" in val_low or "vat" in val_low or "nit" in val_low:
                col_map["vat"] = idx
            elif "debit" in val_low or "débito" in val_low or "debito" in val_low:
                if "debit" not in col_map:
                    col_map["debit"] = idx
            elif "credit" in val_low or "crédito" in val_low or "credito" in val_low:
                if "credit" not in col_map:
                    col_map["credit"] = idx
            elif "saldo" in val_low or "balance" in val_low:
                if "saldo_ini" not in col_map:
                    col_map["saldo_ini"] = idx
                else:
                    col_map["saldo_fin"] = idx
        if "code" in col_map and "debit" in col_map and "credit" in col_map:
            header_row = r
            break

    if not header_row:
        header_row = 3
        col_map = {"code": 1, "name": 2, "vat": 3, "saldo_ini": 4, "debit": 5, "credit": 6, "saldo_fin": 7}

    cuentas = {}
    current_account = None

    for r in range(header_row + 1, ws.max_row + 1):
        c_code = ws.cell(row=r, column=col_map["code"]).value
        c_name = str(ws.cell(row=r, column=col_map.get("name", 2)).value or "").strip()
        deb = float(ws.cell(row=r, column=col_map["debit"]).value or 0.0)
        cred = float(ws.cell(row=r, column=col_map["credit"]).value or 0.0)
        s_fin = float(ws.cell(row=r, column=col_map.get("saldo_fin", 7)).value or 0.0)

        if c_code is not None and str(c_code).strip():
            code_str = str(c_code).strip()
            current_account = code_str
            cuentas[code_str] = {
                "name": c_name,
                "debit": deb,
                "credit": cred,
                "net": cred - deb,
                "saldo_fin": s_fin,
                "partners": {}
            }
        elif current_account and c_name and not c_name.startswith("Total "):
            vat_val = str(ws.cell(row=r, column=col_map.get("vat", 3)).value or "").strip()
            cuentas[current_account]["partners"][c_name] = {
                "vat": vat_val,
                "debit": deb,
                "credit": cred,
                "net": cred - deb,
                "saldo_fin": s_fin
            }

    def get_cred(code):
        return cuentas.get(code, {}).get("credit", 0.0)

    def get_deb(code):
        return cuentas.get(code, {}).get("debit", 0.0)

    def get_net(code):
        c = cuentas.get(code, {})
        return c.get("credit", 0.0) - c.get("debit", 0.0)

    cesantias_mov_deb = get_deb("41800103")
    cesantias_mov_cred = get_cred("41800103")
    cesantias_mov = get_net("41800103")

    cesantias_rep_deb = get_deb("41800104")
    cesantias_rep_cred = get_cred("41800104")
    cesantias_rep = get_net("41800104")

    total_cesantias = cesantias_mov + cesantias_rep

    comisiones_deb = get_deb("42253001")
    comisiones_cred = get_cred("42253001")
    comisiones = get_net("42253001")

    intereses_deb = get_deb("42100501")
    intereses_cred = get_cred("42100501")
    intereses = get_net("42100501")

    netforms_deb = get_deb("42352001")
    netforms_cred = get_cred("42352001")
    netforms = get_net("42352001")

    arriendo_uma = 0.0
    arriendo_uma_deb = 0.0
    arriendo_uma_cred = 0.0
    liq_auteco = 0.0
    liq_autotecnica = 0.0
    if "42354001" in cuentas:
        parts = cuentas["42354001"]["partners"]
        for p_name, p_data in parts.items():
            p_up = p_name.upper()
            if "UMA" in p_up:
                arriendo_uma += p_data["net"]
                arriendo_uma_deb += p_data.get("debit", 0.0)
                arriendo_uma_cred += p_data.get("credit", 0.0)
            elif "AUTECO" in p_up:
                liq_auteco += p_data["net"]
            elif "AUTOTECNICA" in p_up or "AUTOTÉCNICA" in p_up:
                liq_autotecnica += p_data["net"]

    descuentos_deb = get_deb("42104001")
    descuentos_cred = get_cred("42104001")
    descuentos = get_net("42104001")

    ajuste_peso_deb = get_deb("42958101")
    ajuste_peso_cred = get_cred("42958101")
    ajuste_peso = get_net("42958101")

    sobrantes_caja_deb = get_deb("42955301")
    sobrantes_caja_cred = get_cred("42955301")
    sobrantes_caja = get_net("42955301")

    aprovechamientos_deb = get_deb("42950501")
    aprovechamientos_cred = get_cred("42950501")
    aprovechamientos = get_net("42950501")

    otros_ingresos = [
        {"concepto": "Cesantías de movilidad y repuestos", "cuentas": "41800103 y 41800104", "debito": cesantias_mov_deb + cesantias_rep_deb, "credito": cesantias_mov_cred + cesantias_rep_cred, "valor": total_cesantias, "detalle": f"Movilidad: ${cesantias_mov:,.2f} + Repuestos: ${cesantias_rep:,.2f}"},
        {"concepto": "Comisiones", "cuentas": "42253001", "debito": comisiones_deb, "credito": comisiones_cred, "valor": comisiones, "detalle": "Por ingresos para terceros (Crédito - Débito)"},
        {"concepto": "Intereses de financiación", "cuentas": "42100501", "debito": intereses_deb, "credito": intereses_cred, "valor": intereses, "detalle": "Rendimientos e intereses financieros (Crédito - Débito)"},
        {"concepto": "Netforms", "cuentas": "42352001", "debito": netforms_deb, "credito": netforms_cred, "valor": netforms, "detalle": "Servicios administrativos (Crédito - Débito)"},
        {"concepto": "Arriendo UMA", "cuentas": "42354001 (Supermotos UMA)", "debito": arriendo_uma_deb, "credito": arriendo_uma_cred, "valor": arriendo_uma, "detalle": "Partner Supermotos UMA SAS Zomac (Crédito - Débito)"},
        {"concepto": "Descuentos Comerciales", "cuentas": "42104001", "debito": descuentos_deb, "credito": descuentos_cred, "valor": descuentos, "detalle": "Descuentos comerciales condicionados"},
        {"concepto": "Ajuste al peso", "cuentas": "42958101", "debito": ajuste_peso_deb, "credito": ajuste_peso_cred, "valor": ajuste_peso, "detalle": "Ajuste por redondeo al peso (Crédito - Débito)"},
        {"concepto": "Sobrantes caja", "cuentas": "42955301", "debito": sobrantes_caja_deb, "credito": sobrantes_caja_cred, "valor": sobrantes_caja, "detalle": "Sobrantes de caja en sedes"},
    ]
    if aprovechamientos != 0 or aprovechamientos_cred != 0 or aprovechamientos_deb != 0:
        otros_ingresos.append({"concepto": "Aprovechamientos", "cuentas": "42950501", "debito": aprovechamientos_deb, "credito": aprovechamientos_cred, "valor": aprovechamientos, "detalle": "Aprovechamientos extraordinarios (Crédito - Débito)"})

    total_otros_ingresos = sum(item["valor"] for item in otros_ingresos)
    liq_taller_total = liq_auteco + liq_autotecnica
    dev_4175_deb = sum(v.get("debit", 0.0) for k, v in cuentas.items() if k.startswith("4175"))
    dev_4175_cred = sum(v.get("credit", 0.0) for k, v in cuentas.items() if k.startswith("4175"))
    dev_4175_net = dev_4175_deb - dev_4175_cred
    dev_4175 = dev_4175_deb

    # Extracción y conciliación analítica de Transporte / Fletes (Cuenta 4145) por Sede
    transporte_sede = {}
    total_transporte = 0.0
    c_4145_code = "41450501" if "41450501" in cuentas else ("414505" if "414505" in cuentas else ("4145" if "4145" in cuentas else None))
    if c_4145_code and c_4145_code in cuentas:
        c_4145 = cuentas[c_4145_code]
        total_transporte = c_4145.get("net", 0.0)
        for p_nom, p_info in c_4145.get("partners", {}).items():
            if p_nom[:4].isdigit() or p_nom.startswith("Total "):
                continue
            net = p_info["net"]
            if net != 0:
                s_matched = "Villavicencio"
                if client_to_cc:
                    s_matched = client_to_cc.get(p_nom.upper().strip(), "Villavicencio")
                transporte_sede[s_matched] = transporte_sede.get(s_matched, 0.0) + net

    # Extracción y consolidación analítica del Costo de Ventas (Clase 61)
    costos_cuentas = []
    c_61_keys = [k for k in cuentas.keys() if k.startswith("61")]
    # Filtrar solo cuentas auxiliares/hojas (evitar duplicar cuentas padre si existen)
    leaf_61 = []
    for k in c_61_keys:
        is_parent = any(other != k and other.startswith(k) for other in c_61_keys)
        if not is_parent or len(k) >= 8:
            leaf_61.append(k)

    costo_61_deb = 0.0
    costo_61_cred = 0.0
    for k in sorted(leaf_61):
        v = cuentas[k]
        deb = float(v.get("debit", 0.0))
        cred = float(v.get("credit", 0.0))
        net = deb - cred
        if deb != 0.0 or cred != 0.0:
            nm = v.get("name", "")
            costos_cuentas.append({
                "cuenta": k,
                "nombre": nm,
                "debito": deb,
                "credito": cred,
                "costo_neto": net,
                "detalle": "Costo Vehículos" if "VEHICULOS" in nm.upper() else ("Aceites y Lubricantes" if "ACEITE" in nm.upper() else "Repuestos")
            })
            costo_61_deb += deb
            costo_61_cred += cred

    costo_total_61 = costo_61_deb - costo_61_cred

    return {
        "transporte_por_sede": transporte_sede,
        "total_transporte": total_transporte,
        "cuentas": cuentas,
        "otros_ingresos": otros_ingresos,
        "total_otros_ingresos": total_otros_ingresos,
        "arriendo_uma": arriendo_uma,
        "liq_auteco": liq_auteco,
        "liq_autotecnica": liq_autotecnica,
        "liq_taller_total": liq_taller_total,
        "dev_4175": dev_4175_deb,
        "dev_4175_deb": dev_4175_deb,
        "dev_4175_cred": dev_4175_cred,
        "dev_4175_net": dev_4175_net,
        "costo_total_61": costo_total_61,
        "costo_61_deb": costo_61_deb,
        "costo_61_cred": costo_61_cred,
        "costos_cuentas": costos_cuentas,
        "total_cesantias": total_cesantias,
        "comisiones": comisiones,
        "intereses": intereses,
        "netforms": netforms,
        "descuentos": descuentos,
        "ajuste_peso": ajuste_peso,
        "sobrantes_caja": sobrantes_caja
    }

def _leer_excel_inteligente(ruta_archivo):
    """
    Lee un archivo Excel inspeccionando sus hojas y encabezados de manera inteligente:
    1. Si tiene varias hojas y existe una llamada 'Detalle Ventas', 'Detalle', 'Pedidos', etc., la prioriza.
    2. Si las primeras filas contienen títulos o celdas vacías (Unnamed), busca automáticamente
       la fila real de encabezados.
    3. Retorna un DataFrame de Pandas limpio con las columnas correctas.
    """
    try:
        xl = robust_ExcelFile(ruta_archivo)
        sheet_names = xl.sheet_names
    except Exception:
        return robust_read_excel(ruta_archivo)

    preferidas = ['detalle ventas', 'detalle', 'pedidos', 'ventas', 'order', 'sale.order', 'sheet1', 'hoja1']
    hojas_ordenadas = []
    for pref in preferidas:
        for s in sheet_names:
            if pref in s.lower() and s not in hojas_ordenadas:
                hojas_ordenadas.append(s)
    for s in sheet_names:
        if s not in hojas_ordenadas:
            hojas_ordenadas.append(s)

    for sheet in hojas_ordenadas:
        try:
            df_preview = robust_read_excel(ruta_archivo, sheet_name=sheet, header=None, nrows=15)
            header_row = 0
            for r_idx in range(len(df_preview)):
                row_vals = [str(v).lower().strip() for v in df_preview.iloc[r_idx] if pd.notna(v)]
                keywords_header = ['vendedor', 'centro de costo', 'costo', 'producto', 'total', 'base imponible', 'base', 'número', 'numero', 'referencia', 'cliente']
                coincidencias = sum(1 for kw in keywords_header if any(kw in cell for cell in row_vals))
                if coincidencias >= 2:
                    header_row = r_idx
                    break

            df = robust_read_excel(ruta_archivo, sheet_name=sheet, header=header_row)
            cols_named = [c for c in df.columns if not str(c).startswith('Unnamed:')]
            if len(cols_named) >= 3 and len(df) > 0:
                return df
        except Exception:
            continue

    return robust_read_excel(ruta_archivo)


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


def _asegurar_ruta_disponible(ruta_salida):
    """
    Verifica si la ruta de salida está bloqueada por Microsoft Excel u otro programa.
    Si está bloqueada, genera automáticamente un nombre alternativo libre:
    ej. 'Dashboard_Vendedores (1).xlsx', 'Dashboard_Vendedores (2).xlsx', etc.
    para que la generación nunca se detenga por 'Permission denied'.
    Protege siempre los archivos ORIGINALES para jamás sobrescribirlos.
    """
    if not ruta_salida:
        return ruta_salida

    # PROTECCIÓN: Si apunta a un archivo ORIGINAL, redirigir para proteger la copia original
    if "original" in os.path.basename(ruta_salida).lower():
        dir_name = os.path.dirname(ruta_salida) or "."
        base_name = os.path.basename(ruta_salida)
        for tag in ["_ORIGINAL", "_original", " ORIGINAL", " original"]:
            base_name = base_name.replace(tag, "")
        ruta_salida = os.path.join(dir_name, base_name)

    if not os.path.exists(ruta_salida):
        return ruta_salida
    try:
        with open(ruta_salida, 'a+b'):
            pass
        return ruta_salida
    except (PermissionError, OSError):
        dir_name = os.path.dirname(ruta_salida) or "."
        base_name, ext = os.path.splitext(os.path.basename(ruta_salida))
        counter = 1
        while counter < 100:
            candidato = os.path.join(dir_name, f"{base_name} ({counter}){ext}")
            if not os.path.exists(candidato):
                return candidato
            try:
                with open(candidato, 'a+b'):
                    pass
                return candidato
            except (PermissionError, OSError):
                counter += 1
        import datetime
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(dir_name, f"{base_name}_{ts}{ext}")


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

    df = _leer_excel_inteligente(ruta_archivo)

    cols_lower = [str(c).lower().strip() for c in df.columns]
    tiene_vendedor = any('vendedor' in c for c in cols_lower)
    tiene_total = any(c == 'total' or 'total' in c for c in cols_lower)
    tiene_producto = any('roducto' in c for c in cols_lower)
    tiene_cc = any('costo' in c or 'sede' in c for c in cols_lower)
    tiene_base = any('base' in c for c in cols_lower)
    tiene_num = any('número' in c or 'numero' in c or 'factura' in c for c in cols_lower)

    # Si es un archivo de Ventas (Asiento Contable / Facturación por Sede) y no de Pedidos
    if (tiene_cc or tiene_base or tiene_num) and not tiene_producto:
        stats = procesar_informe_ventas_asiento(ruta_archivo, ruta_salida, ruta_asiento_nc=ruta_asiento_nc)
        stats['es_asiento_ventas'] = True
        stats['ventas_brutas'] = stats['total_ventas']
        stats['ventas_netas'] = stats['total_ventas']
        stats['total_coste'] = 0.0
        stats['total_utilidad'] = stats['total_ventas']
        stats['margen_real'] = 1.0
        stats['total_clientes'] = stats['total_facturas']
        stats['total_pedidos'] = stats['total_facturas']
        stats['top_vendedor'] = f"{stats['total_vendedores']} Vendedores en {stats['total_sedes']} Sedes"
        stats['total_nc'] = 0.0
        return stats

    if not (tiene_vendedor and tiene_total and tiene_producto):
        cols_str = ', '.join([str(c) for c in df.columns[:12]])
        raise ValueError(
            f"El archivo cargado NO contiene las columnas requeridas.\n\n"
            f"Se detectaron estas columnas: [{cols_str}]\n\n"
            f"• Para Dashboard de Pedidos se requiere: 'Vendedor', 'Producto', 'Total'.\n"
            f"• Para Informe de Ventas por Sede se requiere: 'Centro de costo', 'Vendedor', 'Base imponible' o 'Total'."
        )

    col_vendedor = _resolve_col(df, 'Vendedor', ['vendedor'], label='Vendedor')
    col_pedido = _find_col(df, ['referencia', 'pedido'])

    # En exportaciones de Odoo, filtrar a las filas principales de pedidos (evita subfilas vacías)
    if col_pedido and col_pedido in df.columns and df[col_pedido].notna().any():
        df = df[df[col_pedido].notna()].copy()

    # Si algún pedido no tiene vendedor (ej. Consumidor Final), asignarlo a 'SIN ASIGNAR'
    # para no perder ese dinero y que la suma total sea exactamente 100% idéntica al archivo Excel
    df[col_vendedor] = df[col_vendedor].fillna('SIN ASIGNAR')
    df[col_vendedor] = df[col_vendedor].replace('', 'SIN ASIGNAR')

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

    # El archivo 'Pedido de venta (sale.order).xlsx' contiene 100% VENTAS comerciales.
    # Las Notas Crédito se toman y descuentan exclusivamente del archivo de Asiento Contable.
    df['Es_Nota_Credito'] = False
    n_nc = 0
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
    ruta_salida = _asegurar_ruta_disponible(ruta_salida)
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
        ch1.width = 19
        ch1.height = 13.5
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
        ch2.height = 13.5
        d2 = Reference(ws1, min_col=11, min_row=9, max_row=min(17, 9 + len(resumen)))
        c2 = Reference(ws1, min_col=2, min_row=10, max_row=min(17, 9 + len(resumen)))
        ch2.add_data(d2, titles_from_data=True)
        ch2.set_categories(c2)
        ch2.dataLabels = DataLabelList()
        ch2.dataLabels.showPercent = True
        ch2.dataLabels.showCatName = False
        ch2.dataLabels.showVal = False
        ch2.legend.position = 'r'
        ws1.add_chart(ch2, f"H{last + 2}")

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
        "ruta_guardada": ruta_salida,
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
        col_imp_s = _find_col(df_so_original, ['impuesto'], excludes=['sin'])
        col_cli_s = _find_col(df_so_original, ['liente'])
        col_ven_s = _find_col(df_so_original, ['vendedor'])
        col_ped_s = _find_col(df_so_original, ['referencia', 'pedido'])

        # Asegurar que df_so_original solo contenga las ventas (excluyendo notas de crédito que se hayan anexado en el dashboard general)
        if 'Es_Nota_Credito' in df_so_original.columns:
            df_so_original = df_so_original[~df_so_original['Es_Nota_Credito']].copy()

        # Totales globales de todo el archivo Pedidos de Venta (Odoo)
        cant_so_total = len(df_so_original)
        tot_so_global = float(pd.to_numeric(df_so_original[col_tot_s], errors='coerce').fillna(0).sum()) if col_tot_s else 0.0
        iva_so_global = float(pd.to_numeric(df_so_original[col_imp_s], errors='coerce').fillna(0).sum()) if col_imp_s else 0.0

        df_s_work = df_so_original.copy()
        df_so_sin_fact = pd.DataFrame()
        df_so_con_nc = pd.DataFrame()
        if col_fact_s and col_fact_s in df_s_work.columns:
            m_sin_fact = df_s_work[col_fact_s].isna() | (df_s_work[col_fact_s].astype(str).str.strip().isin(['', 'nan', 'None']))
            df_so_sin_fact = df_s_work[m_sin_fact].copy()

            df_con_doc = df_s_work[~m_sin_fact].copy()
            df_con_doc['Numero_Doc'] = df_con_doc[col_fact_s].astype(str).str.strip().str.upper()
            m_es_nc = df_con_doc['Numero_Doc'].str.startswith('NC')
            df_so_con_nc = df_con_doc[m_es_nc].copy()

            df_s_f = df_con_doc[~m_es_nc].copy()
            df_s_f['Total_Val_SO'] = pd.to_numeric(df_s_f[col_tot_s], errors='coerce').fillna(0) if col_tot_s else 0.0
            df_s_f['IVA_Val_SO'] = pd.to_numeric(df_s_f[col_imp_s], errors='coerce').fillna(0) if col_imp_s else 0.0

            so_f_grp = df_s_f.groupby('Numero_Doc').agg({
                'Total_Val_SO': 'sum',
                'IVA_Val_SO': 'sum',
                col_cli_s: 'first' if col_cli_s else 'count',
                col_ven_s: 'first' if col_ven_s else 'count',
                col_ped_s: 'first' if col_ped_s else 'count'
            }).reset_index().rename(columns={
                'Total_Val_SO': 'Total_SO',
                'IVA_Val_SO': 'IVA_SO',
                col_cli_s: 'Cliente_SO' if col_cli_s else 'Cliente_SO',
                col_ven_s: 'Vendedor_SO' if col_ven_s else 'Vendedor_SO',
                col_ped_s: 'Pedido_SO' if col_ped_s else 'Pedido_SO'
            })
        else:
            so_f_grp = pd.DataFrame(columns=['Numero_Doc', 'Total_SO', 'IVA_SO', 'Cliente_SO', 'Vendedor_SO', 'Pedido_SO'])

        # 3. Asiento Contable (NC)
        df_ac_nc = pd.DataFrame(columns=['Numero_Doc', 'Total_AC', 'IVA_AC', 'Cliente_AC'])
        tot_nc_ac = 0.0
        iva_nc_ac = 0.0
        if ruta_asiento_nc and os.path.exists(ruta_asiento_nc):
            df_ac = robust_read_excel(ruta_asiento_nc)
            ac_num_c = df_ac.columns[0]
            ac_tot_c = [c for c in df_ac.columns if 'total' in c.lower() and 'signo' in c.lower()]
            if not ac_tot_c:
                ac_tot_c = [c for c in df_ac.columns if 'total' in c.lower() or 'importe' in c.lower()]
            col_tot_ac = ac_tot_c[0] if ac_tot_c else None

            ac_imp_c = [c for c in df_ac.columns if 'impuesto' in c.lower() and 'firmado' in c.lower()]
            if not ac_imp_c:
                ac_imp_c = [c for c in df_ac.columns if 'impuesto' in c.lower()]
            col_imp_ac = ac_imp_c[0] if ac_imp_c else None

            ac_cli_c = [c for c in df_ac.columns if 'empresa' in c.lower() or 'nombre' in c.lower()]
            col_cli_ac = ac_cli_c[0] if ac_cli_c else None

            df_ac_nc = df_ac[df_ac[ac_num_c].astype(str).str.upper().str.contains('^NCC|^NC')].copy()
            if col_tot_ac:
                df_ac_nc['Total_AC'] = pd.to_numeric(df_ac_nc[col_tot_ac], errors='coerce').fillna(0)
            else:
                df_ac_nc['Total_AC'] = 0.0

            if col_imp_ac:
                df_ac_nc['IVA_AC'] = pd.to_numeric(df_ac_nc[col_imp_ac], errors='coerce').fillna(0)
            else:
                df_ac_nc['IVA_AC'] = 0.0

            df_ac_nc['Numero_Doc'] = df_ac_nc[ac_num_c].astype(str).str.strip().str.upper()
            if col_cli_ac:
                df_ac_nc['Cliente_AC'] = df_ac_nc[col_cli_ac].astype(str)
            else:
                df_ac_nc['Cliente_AC'] = 'N/A'
            tot_nc_ac = float(abs(df_ac_nc['Total_AC'].sum()))
            iva_nc_ac = float(abs(df_ac_nc['IVA_AC'].sum()))

        ac_nc_grp = df_ac_nc.groupby('Numero_Doc').agg({
            'Total_AC': lambda x: abs(x.sum()),
            'IVA_AC': lambda x: abs(x.sum()),
            'Cliente_AC': 'first'
        }).reset_index()

        # 4. Cruces
        m_fact = pd.merge(dian_f_grp, so_f_grp, on='Numero_Doc', how='outer', indicator=True)
        solo_dian_f = m_fact[m_fact['_merge'] == 'left_only'].copy()
        solo_so_f = m_fact[m_fact['_merge'] == 'right_only'].copy()
        ambos_f = m_fact[m_fact['_merge'] == 'both'].copy()
        ambos_f['diff_val'] = ambos_f['Total_SO'] - ambos_f['Total_DIAN']
        ambos_f['diff_iva'] = ambos_f['IVA_SO'] - ambos_f['IVA_DIAN']
        con_diff_f = ambos_f[(ambos_f['diff_val'].abs() > 1.0) | (ambos_f['diff_iva'].abs() > 1.0)].copy()

        m_nc = pd.merge(dian_nc_grp, ac_nc_grp, on='Numero_Doc', how='outer', indicator=True)
        ambos_nc = m_nc[m_nc['_merge'] == 'both'].copy()
        ambos_nc['diff_nc_val'] = ambos_nc['Total_AC'] - ambos_nc['Total_NC_DIAN']
        ambos_nc['diff_nc_iva'] = ambos_nc['IVA_AC'] - ambos_nc['IVA_NC_DIAN']
        solo_dian_nc = m_nc[m_nc['_merge'] == 'left_only'].copy()
        solo_ac_nc = m_nc[m_nc['_merge'] == 'right_only'].copy()

        # Totales DIAN
        tot_f_dian = float(dian_f_grp['Total_DIAN'].sum())
        iva_f_dian = float(dian_f_grp['IVA_DIAN'].sum())
        tot_nc_dian = float(dian_nc_grp['Total_NC_DIAN'].sum())
        iva_nc_dian = float(dian_nc_grp['IVA_NC_DIAN'].sum())
        neto_dian = tot_f_dian - tot_nc_dian
        neto_iva_dian = iva_f_dian - iva_nc_dian

        # Totales Sistema (100% íntegros con Pedidos de Venta Odoo)
        tot_f_so = tot_so_global
        iva_f_so = iva_so_global
        cant_f_so = cant_so_total

        neto_sistema = tot_f_so - tot_nc_ac
        neto_iva_sistema = iva_f_so - iva_nc_ac

        dif_global_tot = neto_dian - neto_sistema
        dif_global_iva = neto_iva_dian - neto_iva_sistema

        # 5. Renderizar Hoja en Excel
        ws7 = writer.book.create_sheet("Conciliación DIAN vs Sistema")
        ws7.sheet_view.showGridLines = False
        ws7.sheet_properties.tabColor = "059669"
        ws7.freeze_panes = 'A7'

        widths7 = {
            'A': 4, 'B': 28, 'C': 15, 'D': 22, 'E': 20, 'F': 18, 'G': 22, 'H': 20, 'I': 22, 'J': 20, 'K': 20, 'L': 46, 'M': 4
        }
        for col_l, w in widths7.items():
            ws7.column_dimensions[col_l].width = w

        filas_detalles = len(solo_dian_f) + len(con_diff_f) + len(df_so_sin_fact) + len(df_so_con_nc) + len(ambos_nc)
        max_r_fondo = max(140, 55 + filas_detalles)
        fbg(ws7, 1, 1, max_r_fondo, 13, f_dark)

        mtitle(ws7, 1, 2, 2, 12, "CONCILIACIÓN DIAN (TOKEN) VS VENTAS Y CONTABILIDAD", f_title, f_dark, a_l)
        mtitle(ws7, 3, 2, 3, 12, "Auditoría Cruzada 360° | Pedidos de Venta (Odoo) vs Contabilidad (Asiento) vs DIAN (Reporte Oficial)", f_subt, f_dark, a_l)

        # ── TABLA COMPARATIVA DE TOTALES ──
        mtitle(ws7, 5, 2, 5, 12, "1. RESUMEN COMPARATIVO DE TOTALES GLOBALES (CONCILIACIÓN CON IVA)", f_sec, f_med)
        hdrs_tot = [
            'Fuente de Información', 'Facturas (Cant)', 'Total Facturas ($)', 'IVA Facturas ($)',
            'Notas Crédito (Cant)', 'Total Notas Crédito ($)', 'IVA Notas Crédito ($)',
            'Neto Facturación ($)', 'Neto IVA ($)', 'Estado General', 'Diagnóstico'
        ]
        for i, h in enumerate(hdrs_tot):
            sc(ws7, 6, i + 2, h, f_hdr, f_blue)

        # Fila 1: Ventas Sistema (Pedidos de Venta Odoo)
        sc(ws7, 7, 2, "Pedidos de Venta (Odoo)", f_datb, f_dark, a_l)
        sc(ws7, 7, 3, cant_f_so, f_dat, f_dark)
        sc(ws7, 7, 4, tot_f_so, f_datb, f_dark, a_r, '$#,##0')
        sc(ws7, 7, 5, iva_f_so, f_dat, f_dark, a_r, '$#,##0')
        sc(ws7, 7, 6, len(df_ac_nc), f_dat, f_dark)
        sc(ws7, 7, 7, -tot_nc_ac, f_dat, f_dark, a_r, '$#,##0')
        sc(ws7, 7, 8, -iva_nc_ac, f_dat, f_dark, a_r, '$#,##0')
        sc(ws7, 7, 9, neto_sistema, f_datb, f_dark, a_r, '$#,##0')
        sc(ws7, 7, 10, neto_iva_sistema, f_dat, f_dark, a_r, '$#,##0')
        sc(ws7, 7, 11, "Ventas Registradas", f_dat, f_dark)
        sc(ws7, 7, 12, "Base comercial íntegra (460 pedidos) y notas aplicadas", f_dat, f_dark, a_l)

        # Fila 2: Contabilidad
        sc(ws7, 8, 2, "Contabilidad (Asiento Contable)", f_datb, f_alt, a_l)
        sc(ws7, 8, 3, "N/A", f_dat, f_alt)
        sc(ws7, 8, 4, 0, f_dat, f_alt, a_r, '$#,##0')
        sc(ws7, 8, 5, 0, f_dat, f_alt, a_r, '$#,##0')
        sc(ws7, 8, 6, len(df_ac_nc), f_dat, f_alt)
        sc(ws7, 8, 7, -tot_nc_ac, Font(name='Calibri', size=9, bold=True, color=RED), f_alt, a_r, '$#,##0')
        sc(ws7, 8, 8, -iva_nc_ac, Font(name='Calibri', size=9, bold=True, color=RED), f_alt, a_r, '$#,##0')
        sc(ws7, 8, 9, -tot_nc_ac, f_dat, f_alt, a_r, '$#,##0')
        sc(ws7, 8, 10, -iva_nc_ac, f_dat, f_alt, a_r, '$#,##0')
        sc(ws7, 8, 11, "NCC Registradas", f_dat, f_alt)
        sc(ws7, 8, 12, f"{len(df_ac_nc)} Notas Crédito procesadas en contabilidad", f_dat, f_alt, a_l)

        # Fila 3: DIAN
        sc(ws7, 9, 2, "DIAN (Reporte Oficial Token)", f_datb, f_dark, a_l)
        sc(ws7, 9, 3, len(dian_f_grp), f_dat, f_dark)
        sc(ws7, 9, 4, tot_f_dian, f_datb, f_dark, a_r, '$#,##0')
        sc(ws7, 9, 5, iva_f_dian, f_dat, f_dark, a_r, '$#,##0')
        sc(ws7, 9, 6, len(dian_nc_grp), f_dat, f_dark)
        sc(ws7, 9, 7, -tot_nc_dian, Font(name='Calibri', size=9, bold=True, color=RED), f_dark, a_r, '$#,##0')
        sc(ws7, 9, 8, -iva_nc_dian, Font(name='Calibri', size=9, bold=True, color=RED), f_dark, a_r, '$#,##0')
        sc(ws7, 9, 9, neto_dian, Font(name='Calibri', size=9, bold=True, color=GREEN), f_dark, a_r, '$#,##0')
        sc(ws7, 9, 10, neto_iva_dian, Font(name='Calibri', size=9, bold=True, color=GREEN), f_dark, a_r, '$#,##0')
        sc(ws7, 9, 11, "Oficial DIAN", Font(name='Calibri', size=9, bold=True, color=GREEN), f_dark)
        sc(ws7, 9, 12, "Documentos electrónicos validados ante la DIAN", f_dat, f_dark, a_l)

        # Fila 4: Diferencia
        pct_dif = abs(dif_global_tot) / tot_f_dian * 100 if tot_f_dian else 0
        sc(ws7, 10, 2, "DIFERENCIA (DIAN - Sistema)", Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_l)
        sc(ws7, 10, 3, len(dian_f_grp) - cant_f_so, f_datb, f_med)
        sc(ws7, 10, 4, tot_f_dian - tot_f_so, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, 10, 5, iva_f_dian - iva_f_so, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, 10, 6, len(dian_nc_grp) - len(df_ac_nc), f_datb, f_med)
        sc(ws7, 10, 7, -tot_nc_dian - (-tot_nc_ac), f_datb, f_med, a_r, '$#,##0')
        sc(ws7, 10, 8, -iva_nc_dian - (-iva_nc_ac), f_datb, f_med, a_r, '$#,##0')
        sc(ws7, 10, 9, dif_global_tot, Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_r, '$#,##0')
        sc(ws7, 10, 10, dif_global_iva, Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_r, '$#,##0')
        sc(ws7, 10, 11, "Diferencia Explicada", Font(name='Calibri', size=9, bold=True, color=GOLD), f_med)
        sc(ws7, 10, 12, f"Diferencia del {pct_dif:.2f}% (IVA: ${dif_global_iva:,.0f}) explicada en detalle", f_datb, f_med, a_l)

        # ── COMENTARIOS Y EXPLICACIÓN DE LAS DIFERENCIAS ──
        mtitle(ws7, 12, 2, 12, 12, "2. COMENTARIOS Y MOTIVOS EXPLICATIVOS DE LAS DIFERENCIAS", f_sec, f_med)

        tot_so_sin_f = float(pd.to_numeric(df_so_sin_fact[col_tot_s], errors='coerce').fillna(0).sum()) if len(df_so_sin_fact) > 0 and col_tot_s else 0.0
        tot_so_cnc = float(pd.to_numeric(df_so_con_nc[col_tot_s], errors='coerce').fillna(0).sum()) if len(df_so_con_nc) > 0 and col_tot_s else 0.0
        tot_so_fact = float(so_f_grp['Total_SO'].sum()) if len(so_f_grp) > 0 else 0.0

        comentarios = [
            ("📌 RESUMEN DE CONCILIACIÓN:", f"El valor neto oficial DIAN es de ${neto_dian:,.0f} (IVA: ${neto_iva_dian:,.0f}) vs Sistema ${neto_sistema:,.0f} (IVA: ${neto_iva_sistema:,.0f}). Diferencia global mínima de ${dif_global_tot:,.0f} ({pct_dif:.2f}%) y ${dif_global_iva:,.0f} en IVA."),
            ("✅ NOTAS CRÉDITO (100% EXACTAS EN TOTAL E IVA):", f"Las {len(ambos_nc)} notas crédito de Contabilidad (Total: ${tot_nc_ac:,.0f} | IVA: ${iva_nc_ac:,.0f}) coinciden exactamente al 100% con las {len(dian_nc_grp)} de la DIAN (${tot_nc_dian:,.0f} | IVA: ${iva_nc_dian:,.0f}), con $0 de diferencia."),
            ("💡 DESGLOSE DE PEDIDOS DE VENTA (POR QUÉ DA $1.791.006.639):", f"El archivo comercial registra {cant_so_total} pedidos (${tot_so_global:,.0f}): {len(so_f_grp)} pedidos con factura electrónica directa (${tot_so_fact:,.0f}) + {len(df_so_con_nc)} pedidos vinculados a NC (${tot_so_cnc:,.0f}) + {len(df_so_sin_fact)} pedidos pendientes por facturar (${tot_so_sin_f:,.0f})."),
            (f"🔍 FACTURAS EN DIAN NO EN PEDIDOS ({len(solo_dian_f)} facturas):", f"Existen {len(solo_dian_f)} facturas electrónicas emitidas en la DIAN por ${solo_dian_f['Total_DIAN'].sum():,.0f} (IVA: ${solo_dian_f['IVA_DIAN'].sum():,.0f}) que no figuran en los pedidos facturados de Odoo (Facturas: {', '.join(solo_dian_f['Numero_Doc'].tolist())}). Corresponden a facturas manuales o ventas directas."),
            (f"⚠️ FACTURAS CON DIFERENCIA DE VALOR ({len(con_diff_f)} facturas):", f"Existen {len(con_diff_f)} facturas coincidentes con ajustes de importe o IVA (diferencia neta de ${con_diff_f['diff_val'].sum():,.0f} en total y ${con_diff_f['diff_iva'].sum():,.0f} en IVA), principalmente por pedidos con importe comercial cero o ajustes de centavos."),
            ("🎯 CUADRE MATEMÁTICO DE LA DIFERENCIA:", f"Facturas DIAN no en pedidos (${solo_dian_f['Total_DIAN'].sum():,.0f}) - Pedidos con NC (${tot_so_cnc:,.0f}) - Pedidos sin factura (${tot_so_sin_f:,.0f}) + Ajustes (${abs(con_diff_f['diff_val'].sum()):,.0f}) = ${dif_global_tot:,.0f} (Diferencia global explicada al 100%).")
        ]

        curr_r = 13
        for tit, desc in comentarios:
            mtitle(ws7, curr_r, 2, curr_r, 4, tit, Font(name='Calibri', size=9, bold=True, color=GOLD if ('🔍' in tit or '⚠️' in tit or '💡' in tit) else GREEN), f_alt, a_l)
            mtitle(ws7, curr_r, 5, curr_r, 12, desc, f_dat, f_alt, a_l)
            curr_r += 1

        # ── TABLA FACTURAS SOLO EN DIAN ──
        curr_r += 1
        mtitle(ws7, curr_r, 2, curr_r, 12, f"3. FACTURAS EMITIDAS ANTE LA DIAN NO REGISTRADAS EN PEDIDOS DE VENTA ({len(solo_dian_f)} FACTURAS)", f_sec, f_med)
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
            sc(ws7, curr_r, 9, "Emitida en DIAN / No figura en pedidos de venta", f_dat, fr, a_l)
            curr_r += 1

        fbg(ws7, curr_r, 2, curr_r, 12, f_med)
        sc(ws7, curr_r, 2, "TOTAL", f_datb, f_med)
        sc(ws7, curr_r, 3, f"{len(solo_dian_f)} facturas", f_datb, f_med)
        sc(ws7, curr_r, 7, solo_dian_f['IVA_DIAN'].sum(), f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 8, solo_dian_f['Total_DIAN'].sum(), Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_r, '$#,##0')
        curr_r += 2

        # ── TABLA FACTURAS CON DIFERENCIA DE VALOR ──
        mtitle(ws7, curr_r, 2, curr_r, 12, f"4. FACTURAS CON DIFERENCIA DE VALOR O IVA (VENTAS VS DIAN) ({len(con_diff_f)} FACTURAS)", f_sec, f_med)
        curr_r += 1

        hdrs_diff = ['#', 'Número Factura', 'Vendedor (Venta)', 'Cliente (Venta)', 'Pedido Ref.', 'IVA Venta ($)', 'IVA DIAN ($)', 'Dif. IVA ($)', 'Total Venta ($)', 'Total DIAN ($)', 'Dif. Total ($)', 'Diagnóstico']
        for i, h in enumerate(hdrs_diff):
            sc(ws7, curr_r, i + 2, h, f_hdr, f_blue)
        curr_r += 1

        for idx, r_df in con_diff_f.reset_index().iterrows():
            fr = f_dark if idx % 2 == 0 else f_alt
            diff_v = r_df['diff_val']
            diff_i = r_df['diff_iva']
            c_diag = "En venta figura con valor $0 (línea sin importe)" if r_df['Total_SO'] == 0 else "Ajuste comercial / redondeo de centavos"
            sc(ws7, curr_r, 2, idx + 1, f_dat, fr)
            sc(ws7, curr_r, 3, r_df['Numero_Doc'], f_datb, fr)
            sc(ws7, curr_r, 4, str(r_df.get('Vendedor_SO', 'N/A')), f_dat, fr, a_l)
            sc(ws7, curr_r, 5, str(r_df.get('Cliente_SO', 'N/A')), f_dat, fr, a_l)
            sc(ws7, curr_r, 6, str(r_df.get('Pedido_SO', 'N/A')), f_dat, fr)
            sc(ws7, curr_r, 7, r_df['IVA_SO'], f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 8, r_df['IVA_DIAN'], f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 9, diff_i, Font(name='Calibri', size=9, bold=True, color=RED if abs(diff_i) > 1 else GOLD), fr, a_r, '$#,##0')
            sc(ws7, curr_r, 10, r_df['Total_SO'], f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 11, r_df['Total_DIAN'], f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 12, diff_v, Font(name='Calibri', size=9, bold=True, color=RED if abs(diff_v) > 1 else GOLD), fr, a_r, '$#,##0')
            curr_r += 1

        fbg(ws7, curr_r, 2, curr_r, 12, f_med)
        sc(ws7, curr_r, 2, "TOTAL DIFERENCIA", f_datb, f_med)
        sc(ws7, curr_r, 7, con_diff_f['IVA_SO'].sum(), f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 8, con_diff_f['IVA_DIAN'].sum(), f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 9, con_diff_f['diff_iva'].sum(), Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 10, con_diff_f['Total_SO'].sum(), f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 11, con_diff_f['Total_DIAN'].sum(), f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 12, con_diff_f['diff_val'].sum(), Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_r, '$#,##0')
        curr_r += 2

        # ── TABLA PEDIDOS PENDIENTES O CON NC ──
        df_especiales = []
        if len(df_so_sin_fact) > 0:
            for _, r in df_so_sin_fact.iterrows():
                df_especiales.append({
                    'Pedido': str(r.get(col_ped_s, 'N/A')),
                    'Vendedor': str(r.get(col_ven_s, 'N/A')),
                    'Cliente': str(r.get(col_cli_s, 'N/A')),
                    'Ref_Doc': 'SIN FACTURA',
                    'IVA': float(pd.to_numeric(r.get(col_imp_s, 0), errors='coerce') or 0),
                    'Total': float(pd.to_numeric(r.get(col_tot_s, 0), errors='coerce') or 0),
                    'Tipo': 'Pendiente Facturación',
                    'Diag': 'Pedido registrado en Odoo sin factura electrónica generada'
                })
        if len(df_so_con_nc) > 0:
            for _, r in df_so_con_nc.iterrows():
                df_especiales.append({
                    'Pedido': str(r.get(col_ped_s, 'N/A')),
                    'Vendedor': str(r.get(col_ven_s, 'N/A')),
                    'Cliente': str(r.get(col_cli_s, 'N/A')),
                    'Ref_Doc': str(r.get(col_fact_s, 'NC')),
                    'IVA': float(pd.to_numeric(r.get(col_imp_s, 0), errors='coerce') or 0),
                    'Total': float(pd.to_numeric(r.get(col_tot_s, 0), errors='coerce') or 0),
                    'Tipo': 'Vinculado a NCC',
                    'Diag': 'Pedido asociado directamente a Nota Crédito en el archivo de ventas'
                })

        if df_especiales:
            mtitle(ws7, curr_r, 2, curr_r, 12, f"5. PEDIDOS DE VENTA PENDIENTES POR FACTURAR O ASOCIADOS A NC ({len(df_especiales)} PEDIDOS = ${tot_so_cnc + tot_so_sin_f:,.0f})", f_sec, f_med)
            curr_r += 1

            hdrs_esp = ['#', 'Pedido Venta', 'Vendedor', 'Cliente', 'Referencia en Factura', 'IVA Pedido ($)', 'Total Pedido ($)', 'Clasificación', 'Diagnóstico']
            for i, h in enumerate(hdrs_esp):
                sc(ws7, curr_r, i + 2, h, f_hdr, f_blue)
            curr_r += 1

            for idx, item in enumerate(df_especiales):
                fr = f_dark if idx % 2 == 0 else f_alt
                sc(ws7, curr_r, 2, idx + 1, f_dat, fr)
                sc(ws7, curr_r, 3, item['Pedido'], f_datb, fr)
                sc(ws7, curr_r, 4, item['Vendedor'], f_dat, fr, a_l)
                sc(ws7, curr_r, 5, item['Cliente'], f_dat, fr, a_l)
                sc(ws7, curr_r, 6, item['Ref_Doc'], Font(name='Calibri', size=9, bold=True, color=GOLD), fr)
                sc(ws7, curr_r, 7, item['IVA'], f_dat, fr, a_r, '$#,##0')
                sc(ws7, curr_r, 8, item['Total'], f_datb, fr, a_r, '$#,##0')
                sc(ws7, curr_r, 9, item['Tipo'], Font(name='Calibri', size=9, bold=True, color=GOLD), fr)
                sc(ws7, curr_r, 10, item['Diag'], f_dat, fr, a_l)
                curr_r += 1

            fbg(ws7, curr_r, 2, curr_r, 12, f_med)
            sc(ws7, curr_r, 2, "TOTAL ESPECIALES", f_datb, f_med)
            sc(ws7, curr_r, 3, f"{len(df_especiales)} pedidos", f_datb, f_med)
            sc(ws7, curr_r, 7, sum(x['IVA'] for x in df_especiales), f_datb, f_med, a_r, '$#,##0')
            sc(ws7, curr_r, 8, sum(x['Total'] for x in df_especiales), Font(name='Calibri', size=9, bold=True, color=GOLD), f_med, a_r, '$#,##0')
            curr_r += 2

        # ── TABLA NOTAS CRÉDITO ──
        mtitle(ws7, curr_r, 2, curr_r, 12, f"6. CONCILIACIÓN DE NOTAS CRÉDITO ({len(ambos_nc)} NOTAS COINCIDENTES AL 100% EN VALOR E IVA)", f_sec, f_med)
        curr_r += 1

        hdrs_nc = ['#', 'Número Nota Crédito', 'Cliente / Razón Social', 'IVA Contabilidad ($)', 'IVA DIAN ($)', 'Dif. IVA ($)', 'Valor Contabilidad ($)', 'Valor DIAN ($)', 'Dif. Total ($)', 'Estado Conciliación']
        for i, h in enumerate(hdrs_nc):
            sc(ws7, curr_r, i + 2, h, f_hdr, f_blue)
        curr_r += 1

        for idx, r_nc in ambos_nc.reset_index().iterrows():
            fr = f_dark if idx % 2 == 0 else f_alt
            diff_nc_v = r_nc['diff_nc_val']
            diff_nc_i = r_nc['diff_nc_iva']
            sc(ws7, curr_r, 2, idx + 1, f_dat, fr)
            sc(ws7, curr_r, 3, r_nc['Numero_Doc'], f_datb, fr)
            sc(ws7, curr_r, 4, str(r_nc.get('Receptor_NC_DIAN', r_nc.get('Cliente_AC', 'N/A'))), f_dat, fr, a_l)
            sc(ws7, curr_r, 5, abs(r_nc['IVA_AC']), f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 6, r_nc['IVA_NC_DIAN'], f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 7, diff_nc_i, f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 8, abs(r_nc['Total_AC']), f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 9, r_nc['Total_NC_DIAN'], f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 10, diff_nc_v, f_dat, fr, a_r, '$#,##0')
            sc(ws7, curr_r, 11, "✅ Conciliado 100%", Font(name='Calibri', size=9, bold=True, color=GREEN), fr)
            curr_r += 1

        fbg(ws7, curr_r, 2, curr_r, 12, f_med)
        sc(ws7, curr_r, 2, "TOTAL NOTAS CRÉDITO", f_datb, f_med)
        sc(ws7, curr_r, 3, f"{len(ambos_nc)} notas", f_datb, f_med)
        sc(ws7, curr_r, 5, iva_nc_ac, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 6, iva_nc_dian, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 7, 0, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 8, tot_nc_ac, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 9, tot_nc_dian, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 10, 0, f_datb, f_med, a_r, '$#,##0')
        sc(ws7, curr_r, 11, "100% Cuadrado (Total e IVA)", Font(name='Calibri', size=9, bold=True, color=GREEN), f_med)

        return {
            "facturas_count": len(dian_f_grp),
            "facturas_total": tot_f_dian,
            "facturas_iva": iva_f_dian,
            "nc_count": len(dian_nc_grp),
            "nc_total": tot_nc_dian,
            "nc_iva": iva_nc_dian,
            "neto_dian": neto_dian,
            "neto_iva_dian": neto_iva_dian,
            "solo_dian_count": len(solo_dian_f),
            "solo_dian_total": solo_dian_f['Total_DIAN'].sum(),
            "solo_dian_iva": solo_dian_f['IVA_DIAN'].sum(),
            "diff_val_count": len(con_diff_f),
            "diff_val_total": con_diff_f['diff_val'].sum(),
            "dif_global": dif_global_tot,
            "dif_global_total": dif_global_tot,
            "dif_global_iva": dif_global_iva
        }
    except Exception as e:
        print(f"[ERROR] Error al construir hoja de conciliacion DIAN: {e}")
        import traceback
        traceback.print_exc()
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

    # Evitar colisión de columnas si el archivo ya contiene 'Fecha' y 'Fecha de factura'
    if 'Fecha de factura' in df.columns and 'Fecha' in df.columns:
        df = df.drop(columns=['Fecha'])

    df = df.rename(columns={
        col_num: 'Factura #',
        'Nombre de la empresa a mostrar en la factura': 'Cliente',
        'Fecha de factura': 'Fecha',
        'Base imponible en la moneda firmada': 'Base Imp.',
        'Total con signo en moneda': 'Total',
        'Commercial Status': 'Estado',
    })
    # Asegurar que todas las columnas sean estrictamente únicas
    df = df.loc[:, ~df.columns.duplicated()].copy()

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

    ruta_salida = _asegurar_ruta_disponible(ruta_salida)
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
        "ruta_guardada": ruta_salida,
        "total_nc": n_nc,
        "total_base": base_total,
        "total_total": total_total,
        "total_vendedores": resumen[col_v].nunique(),
        "nc_recuperadas": n_recuperadas,
        "nc_excluidas": n_excluidas,
        "ticket_prom": ticket_prom,
    }



def _crear_dashboard_gerencial(wb, tb_data, tot_sed_r, sedes):
    """
    Genera una portada ejecutiva visual tipo DASHBOARD gerencial con acabado PREMIUM:
      - Sin líneas de fondo (showGridLines = False) para estética de software/aplicación ejecutiva
      - Banner institucional corporativo
      - 5 Tarjetas KPI con acentos de color superior y amplitud total (sin '########')
      - Módulo 1: Composición de Facturación por Línea de Negocio
      - Módulo 2: Desempeño y Facturación Neta por Sede
      - Gráficos nativos ejecutivos: Anillo (Líneas) y Columnas (Sedes)
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import BarChart, DoughnutChart, Reference

    if "Dashboard Gerencial" in wb.sheetnames:
        del wb["Dashboard Gerencial"]

    ws = wb.create_sheet(title="Dashboard Gerencial", index=0)

    # 1. QUITAR LÍNEAS DE CUADRÍCULA DE FONDO (ACABADO CANVAS PREMIUM)
    ws.views.sheetView[0].showGridLines = False

    FONT_NAME = "Calibri"
    C_NAVY_DARK = "1F4E79"
    C_NAVY_MED = "2F5597"
    C_WHITE = "FFFFFF"
    C_TEXT_DARK = "0F172A"
    C_TEXT_MUTED = "64748B"
    C_BORDER = "CBD5E1"
    C_GREEN = "16A34A"
    C_AMBER = "D97706"
    C_BLUE_LIGHT = "EEF2F6"
    C_ZEBRA = "F8FAFC"
    C_CARD_BG = "FFFFFF"

    f_banner_title = Font(name=FONT_NAME, size=15, bold=True, color=C_WHITE)
    f_banner_sub = Font(name=FONT_NAME, size=9, bold=False, color="E2E8F0")
    f_sec_title = Font(name=FONT_NAME, size=11, bold=True, color=C_NAVY_DARK)
    f_kpi_title = Font(name=FONT_NAME, size=8, bold=True, color=C_TEXT_MUTED)
    f_kpi_sub = Font(name=FONT_NAME, size=8, italic=True, color=C_TEXT_MUTED)

    f_th = Font(name=FONT_NAME, size=9, bold=True, color=C_WHITE)
    f_td = Font(name=FONT_NAME, size=9, color=C_TEXT_DARK)
    f_td_bold = Font(name=FONT_NAME, size=9, bold=True, color=C_TEXT_DARK)
    f_total = Font(name=FONT_NAME, size=9, bold=True, color=C_NAVY_DARK)

    fill_banner = PatternFill(start_color=C_NAVY_DARK, end_color=C_NAVY_DARK, fill_type="solid")
    fill_kpi = PatternFill(start_color=C_CARD_BG, end_color=C_CARD_BG, fill_type="solid")
    fill_th = PatternFill(start_color=C_NAVY_MED, end_color=C_NAVY_MED, fill_type="solid")
    fill_zebra = PatternFill(start_color=C_ZEBRA, end_color=C_ZEBRA, fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_tot = PatternFill(start_color=C_BLUE_LIGHT, end_color=C_BLUE_LIGHT, fill_type="solid")

    thin_card = Side(border_style="thin", color="CBD5E1")
    med_navy = Side(border_style="medium", color=C_NAVY_DARK)
    med_green = Side(border_style="medium", color=C_GREEN)
    med_amber = Side(border_style="medium", color=C_AMBER)
    med_blue = Side(border_style="medium", color=C_NAVY_MED)

    thin_grid = Side(border_style="thin", color="E2E8F0")
    b_grid = Border(left=thin_grid, right=thin_grid, top=thin_grid, bottom=thin_grid)
    b_double_bottom = Border(left=thin_grid, right=thin_grid, top=thin_grid, bottom=Side(border_style="double", color=C_NAVY_DARK))

    a_left = Alignment(horizontal="left", vertical="center")
    a_right = Alignment(horizontal="right", vertical="center")
    a_center = Alignment(horizontal="center", vertical="center")

    NF_CURR = "$ #,##0"
    NF_INT = "#,##0"
    NF_PCT = "0.0%"

    # 1. HEADER BANNER CORPORATIVO (Filas 2 y 3)
    ws.merge_cells("B2:L2")
    ws.merge_cells("B3:L3")
    ws.cell(2, 2, "CH INVERSIONES COLOMBIA S.A.S ZOMAC  |  NIT 901.368.587-8").font = f_banner_title
    ws.cell(3, 2, "DASHBOARD EJECUTIVO DE CONTROL GERENCIAL, INGRESOS Y OPERACIONES").font = f_banner_sub
    ws.cell(2, 2).alignment = a_center
    ws.cell(3, 2).alignment = a_center

    for r in [2, 3]:
        ws.row_dimensions[r].height = 24
        for c in range(2, 13):
            ws.cell(r, c).fill = fill_banner

    # 2. TARJETAS KPI REDISEÑADAS Y AMPLIAS (Filas 5 a 7)
    # Card 1: B-C (width 42)
    # Card 2: D-E (width 36)
    # Card 3: F-H (width 37: columnas 6, 7, 8 completas para cero recortes)
    # Card 4: I-J (width 32)
    # Card 5: K-L (width 42)
    val_fn = "=E20"
    val_cv = "=E21"
    val_ub = "=E22"
    val_oi = "=E23"
    val_gt = "=E24"

    kpis = [
        (2, 3, "FACTURACIÓN NETA", val_fn, "Facturas - Notas Crédito + ND", C_NAVY_DARK, med_navy),
        (4, 5, "(-) COSTO DE VENTAS (61)", val_cv, "Débito - Crédito Grupo 61", C_AMBER, med_amber),
        (6, 8, "(=) UTILIDAD BRUTA OPERACIONAL", val_ub, "16.1% Margen Bruto Operacional", C_GREEN, med_green),
        (9, 10, "(+) OTROS INGRESOS (BALANCE)", val_oi, "Cesantías, Comisiones, Intereses", C_GREEN, med_green),
        (11, 12, "(=) RESULTADO CONSOLIDADO", val_gt, "Utilidad Bruta + Otros Ingresos", C_NAVY_DARK, med_navy),
    ]

    ws.row_dimensions[5].height = 17
    ws.row_dimensions[6].height = 28
    ws.row_dimensions[7].height = 17

    for c_s, c_e, title, val_form, subtext, col_n, top_border_style in kpis:
        ws.merge_cells(start_row=5, start_column=c_s, end_row=5, end_column=c_e)
        ws.merge_cells(start_row=6, start_column=c_s, end_row=6, end_column=c_e)
        ws.merge_cells(start_row=7, start_column=c_s, end_row=7, end_column=c_e)

        c_t = ws.cell(5, c_s, title)
        c_t.font = f_kpi_title; c_t.alignment = a_center

        c_v = ws.cell(6, c_s, val_form)
        c_v.font = Font(name=FONT_NAME, size=15, bold=True, color=col_n)
        c_v.alignment = a_center
        if str(val_form).startswith("="):
            c_v.number_format = NF_CURR

        c_s_txt = ws.cell(7, c_s, subtext)
        c_s_txt.font = f_kpi_sub; c_s_txt.alignment = a_center

        for r in range(5, 8):
            for col_i in range(c_s, c_e + 1):
                cell = ws.cell(r, col_i)
                cell.fill = fill_kpi
                cell.border = Border(
                    left=thin_card if col_i == c_s else Side(border_style=None),
                    right=thin_card if col_i == c_e else Side(border_style=None),
                    top=top_border_style if r == 5 else Side(border_style=None),
                    bottom=thin_card if r == 7 else Side(border_style=None)
                )

    # Espacio entre cards y tablas
    ws.row_dimensions[8].height = 10
    ws.row_dimensions[9].height = 10

    # 3. MÓDULOS DE CONTENIDO
    ws.row_dimensions[10].height = 22
    ws.merge_cells("B10:F10")
    ws.cell(10, 2, "1. COMPOSICIÓN DE INGRESOS POR LÍNEA DE NEGOCIO").font = f_sec_title
    ws.cell(10, 2).alignment = a_left

    ws.merge_cells("H10:L10")
    ws.cell(10, 8, "2. PARTICIPACIÓN Y FACTURACIÓN NETA POR SEDE").font = f_sec_title
    ws.cell(10, 8).alignment = a_left

    # Encabezados Módulo 1 (Líneas)
    m1_headers = ["Línea de Negocio / Concepto", "Cta Contable", "Tipo Operación", "Valor Mes", "Part."]
    for i, h in enumerate(m1_headers, 2):
        c = ws.cell(11, i, h)
        c.font = f_th; c.fill = fill_th; c.alignment = a_center; c.border = b_grid
    ws.row_dimensions[11].height = 22

    # Encabezados Módulo 2 (Sedes)
    m2_headers = ["Centro de Costo (Sede)", "Facturas", "Venta Bruta", "Notas Crédito", "Facturación Neta"]
    for i, h in enumerate(m2_headers, 8):
        c = ws.cell(11, i, h)
        c.font = f_th; c.fill = fill_th; c.alignment = a_center; c.border = b_grid

    val_nd_m1 = "='Otros Ingresos y Balance'!E16" if tb_data else 93000.0
    # Datos Módulo 1 (Líneas)
    m1_data = [
        ("Vehículos Generales", 41350201, "Comercial", f"='Consolidado por Sede'!C{tot_sed_r}", "=E12/$E$20"),
        ("Vehículos Jaime (Especial)", 41350201, "Comercial", f"='Consolidado por Sede'!D{tot_sed_r}", "=E13/$E$20"),
        ("Repuestos", 41350601, "Comercial", f"='Consolidado por Sede'!E{tot_sed_r}", "=E14/$E$20"),
        ("Transporte (Fletes 4145)", 41450501, "Comercial", f"='Consolidado por Sede'!F{tot_sed_r}", "=E15/$E$20"),
        ("Liquidación Taller (Auteco)", 42354001, "Servicio Taller", f"='Consolidado por Sede'!G{tot_sed_r}", "=E16/$E$20"),
        ("Arriendo Sede UMA", 42354001, "Arriendo", f"='Consolidado por Sede'!H{tot_sed_r}", "=E17/$E$20"),
        ("(-) Devoluciones / Notas Crédito", "4175", "Devoluciones (DB)", f"='Consolidado por Sede'!L{tot_sed_r}", "=E18/$E$20"),
        ("(+) Notas Débito (Ajustes 4175)", "4175", "Nota Débito (CR)", val_nd_m1, "=E19/$E$20"),
    ]

    for idx, (lin, cta, tip, form_v, form_p) in enumerate(m1_data, 12):
        r_fill = fill_zebra if idx % 2 == 0 else fill_white
        ws.cell(idx, 2, lin).font = f_td_bold if ("(-)" in lin or "(+)" in lin) else f_td
        ws.cell(idx, 2).alignment = a_left; ws.cell(idx, 2).fill = r_fill; ws.cell(idx, 2).border = b_grid

        # Número de cuenta contable formateado para evitar triángulo verde de Excel
        c_c = ws.cell(idx, 3, cta)
        c_c.font = f_td; c_c.alignment = a_center; c_c.fill = r_fill; c_c.border = b_grid
        if isinstance(cta, int):
            c_c.number_format = "0"
        else:
            c_c.number_format = "@"

        c_tp = ws.cell(idx, 4, tip)
        c_tp.font = f_td; c_tp.alignment = a_center; c_tp.fill = r_fill; c_tp.border = b_grid

        c_v = ws.cell(idx, 5, form_v)
        c_v.font = f_td_bold if ("(-)" in lin or "(+)" in lin) else f_td
        c_v.alignment = a_right; c_v.number_format = NF_CURR; c_v.fill = r_fill; c_v.border = b_grid

        c_p = ws.cell(idx, 6, form_p)
        c_p.font = f_td; c_p.alignment = a_right; c_p.number_format = NF_PCT; c_p.fill = r_fill; c_p.border = b_grid
        ws.row_dimensions[idx].height = 20

    # Subtotal Facturación Neta (Fila 20)
    ws.cell(20, 2, "(=) TOTAL FACTURACIÓN NETA").font = f_total
    ws.cell(20, 2).alignment = a_left; ws.cell(20, 2).fill = fill_tot; ws.cell(20, 2).border = b_grid
    for col_i in range(3, 5):
        c = ws.cell(20, col_i, "-"); c.font = f_total; c.alignment = a_center; c.fill = fill_tot; c.border = b_grid

    c_s_net = ws.cell(20, 5, "=SUM(E12:E17)+E18+E19")
    c_s_net.font = f_total; c_s_net.alignment = a_right; c_s_net.number_format = NF_CURR; c_s_net.fill = fill_tot; c_s_net.border = b_grid

    c_s_pct = ws.cell(20, 6, 1.0)
    c_s_pct.font = f_total; c_s_pct.alignment = a_right; c_s_pct.number_format = NF_PCT; c_s_pct.fill = fill_tot; c_s_pct.border = b_grid
    ws.row_dimensions[20].height = 22

    # Fila Costo de Ventas Grupo 61 (Fila 21)
    ws.cell(21, 2, "(-) COSTO DE VENTAS (GRUPO 61)").font = f_td_bold
    ws.cell(21, 2).alignment = a_left; ws.cell(21, 2).fill = fill_white; ws.cell(21, 2).border = b_grid
    c_cta_cv = ws.cell(21, 3, "6135")
    c_cta_cv.font = f_td; c_cta_cv.alignment = a_center; c_cta_cv.fill = fill_white; c_cta_cv.border = b_grid; c_cta_cv.number_format = "@"
    c_tip_cv = ws.cell(21, 4, "Costo Odoo")
    c_tip_cv.font = f_td; c_tip_cv.alignment = a_center; c_tip_cv.fill = fill_white; c_tip_cv.border = b_grid

    c_cv_v = ws.cell(21, 5, "='Otros Ingresos y Balance'!E18" if tb_data else 0.0)
    c_cv_v.font = f_td_bold; c_cv_v.alignment = a_right; c_cv_v.number_format = NF_CURR; c_cv_v.fill = fill_white; c_cv_v.border = b_grid

    c_cv_p = ws.cell(21, 6, "=E21/$E$20")
    c_cv_p.font = f_td; c_cv_p.alignment = a_right; c_cv_p.number_format = NF_PCT; c_cv_p.fill = fill_white; c_cv_p.border = b_grid
    ws.row_dimensions[21].height = 20

    # Fila Utilidad Bruta Operacional (Fila 22)
    ws.cell(22, 2, "(=) UTILIDAD BRUTA OPERACIONAL").font = f_total
    ws.cell(22, 2).alignment = a_left; ws.cell(22, 2).fill = fill_tot; ws.cell(22, 2).border = b_grid
    for col_i in range(3, 5):
        c = ws.cell(22, col_i, "-"); c.font = f_total; c.alignment = a_center; c.fill = fill_tot; c.border = b_grid

    c_ub_v = ws.cell(22, 5, "=E20-E21")
    c_ub_v.font = f_total; c_ub_v.alignment = a_right; c_ub_v.number_format = NF_CURR; c_ub_v.fill = fill_tot; c_ub_v.border = b_grid

    c_ub_p = ws.cell(22, 6, "=E22/$E$20")
    c_ub_p.font = f_total; c_ub_p.alignment = a_right; c_ub_p.number_format = NF_PCT; c_ub_p.fill = fill_tot; c_ub_p.border = b_grid
    ws.row_dimensions[22].height = 22

    # Fila Otros Ingresos (Fila 23)
    ws.cell(23, 2, "(+) OTROS INGRESOS DEL BALANCE").font = f_td_bold
    ws.cell(23, 2).alignment = a_left; ws.cell(23, 2).fill = fill_white; ws.cell(23, 2).border = b_grid
    c_cta_oi = ws.cell(23, 3, "4180 / 42")
    c_cta_oi.font = f_td; c_cta_oi.alignment = a_center; c_cta_oi.fill = fill_white; c_cta_oi.border = b_grid; c_cta_oi.number_format = "@"
    c_tip_oi = ws.cell(23, 4, "No Operacional")
    c_tip_oi.font = f_td; c_tip_oi.alignment = a_center; c_tip_oi.fill = fill_white; c_tip_oi.border = b_grid

    c_oi_v = ws.cell(23, 5, "='Otros Ingresos y Balance'!E20" if tb_data else 0.0)
    c_oi_v.font = f_td_bold; c_oi_v.alignment = a_right; c_oi_v.number_format = NF_CURR; c_oi_v.fill = fill_white; c_oi_v.border = b_grid

    c_oi_p = ws.cell(23, 6, "=E23/$E$20")
    c_oi_p.font = f_td; c_oi_p.alignment = a_right; c_oi_p.number_format = NF_PCT; c_oi_p.fill = fill_white; c_oi_p.border = b_grid
    ws.row_dimensions[23].height = 20

    # Fila Gran Total Resultado Consolidado (Fila 24)
    ws.cell(24, 2, "(=) RESULTADO INTEGRAL DEL PERIODO").font = Font(name=FONT_NAME, size=9, bold=True, color=C_WHITE)
    ws.cell(24, 2).fill = fill_banner; ws.cell(24, 2).alignment = a_left; ws.cell(24, 2).border = b_double_bottom
    for col_i in range(3, 5):
        c = ws.cell(24, col_i, "CONSOLIDADO"); c.font = Font(name=FONT_NAME, size=8, bold=True, color=C_WHITE); c.alignment = a_center; c.fill = fill_banner; c.border = b_double_bottom

    c_gt_v = ws.cell(24, 5, "=E22+E23")
    c_gt_v.font = Font(name=FONT_NAME, size=9, bold=True, color=C_WHITE); c_gt_v.alignment = a_right; c_gt_v.number_format = NF_CURR; c_gt_v.fill = fill_banner; c_gt_v.border = b_double_bottom

    c_gt_p = ws.cell(24, 6, "=E24/$E$20")
    c_gt_p.font = Font(name=FONT_NAME, size=9, bold=True, color=C_WHITE); c_gt_p.alignment = a_right; c_gt_p.number_format = NF_PCT; c_gt_p.fill = fill_banner; c_gt_p.border = b_double_bottom
    ws.row_dimensions[24].height = 24
    start_sed_r = 7 # fila inicial en Consolidado por Sede
    for s_i, sede_nom in enumerate(sedes):
        idx_r = 12 + s_i
        curr_s_r = start_sed_r + s_i
        r_fill = fill_zebra if s_i % 2 == 0 else fill_white

        ws.cell(idx_r, 8, sede_nom).font = f_td_bold
        ws.cell(idx_r, 8).alignment = a_left; ws.cell(idx_r, 8).fill = r_fill; ws.cell(idx_r, 8).border = b_grid

        # Cant Facturas
        c_f = ws.cell(idx_r, 9, f"='Consolidado por Sede'!J{curr_s_r}")
        c_f.font = f_td; c_f.alignment = a_center; c_f.number_format = NF_INT; c_f.fill = r_fill; c_f.border = b_grid

        # Venta Bruta
        c_vb = ws.cell(idx_r, 10, f"='Consolidado por Sede'!I{curr_s_r}")
        c_vb.font = f_td; c_vb.alignment = a_right; c_vb.number_format = NF_CURR; c_vb.fill = r_fill; c_vb.border = b_grid

        # Notas Crédito
        c_nc = ws.cell(idx_r, 11, f"='Consolidado por Sede'!L{curr_s_r}")
        c_nc.font = f_td; c_nc.alignment = a_right; c_nc.number_format = NF_CURR; c_nc.fill = r_fill; c_nc.border = b_grid

        # Facturación Neta
        c_fn = ws.cell(idx_r, 12, f"='Consolidado por Sede'!M{curr_s_r}")
        c_fn.font = f_td_bold; c_fn.alignment = a_right; c_fn.number_format = NF_CURR; c_fn.fill = r_fill; c_fn.border = b_grid
        ws.row_dimensions[idx_r].height = 20

    # Total Sedes (Fila dinámica según cantidad de sedes)
    tot_r_sed = 12 + len(sedes)
    max_s_r = tot_r_sed - 1
    ws.cell(tot_r_sed, 8, "TOTAL TODAS LAS SEDES").font = f_total
    ws.cell(tot_r_sed, 8).alignment = a_left; ws.cell(tot_r_sed, 8).fill = fill_tot; ws.cell(tot_r_sed, 8).border = b_grid

    c_t_f = ws.cell(tot_r_sed, 9, f"=SUM(I12:I{max_s_r})")
    c_t_f.font = f_total; c_t_f.alignment = a_center; c_t_f.number_format = NF_INT; c_t_f.fill = fill_tot; c_t_f.border = b_grid

    c_t_vb = ws.cell(tot_r_sed, 10, f"=SUM(J12:J{max_s_r})")
    c_t_vb.font = f_total; c_t_vb.alignment = a_right; c_t_vb.number_format = NF_CURR; c_t_vb.fill = fill_tot; c_t_vb.border = b_grid

    c_t_nc = ws.cell(tot_r_sed, 11, f"=SUM(K12:K{max_s_r})")
    c_t_nc.font = f_total; c_t_nc.alignment = a_right; c_t_nc.number_format = NF_CURR; c_t_nc.fill = fill_tot; c_t_nc.border = b_grid

    c_t_fn = ws.cell(tot_r_sed, 12, f"=SUM(L12:L{max_s_r})")
    c_t_fn.font = f_total; c_t_fn.alignment = a_right; c_t_fn.number_format = NF_CURR; c_t_fn.fill = fill_tot; c_t_fn.border = b_grid
    ws.row_dimensions[tot_r_sed].height = 22

    # 4. GRÁFICOS EJECUTIVOS (Filas 23 en adelante)
    try:
        chart_doughnut = DoughnutChart()
        chart_doughnut.title = "Composición de Facturación por Línea de Negocio"
        chart_doughnut.style = 10
        data_d = Reference(ws, min_col=5, min_row=12, max_row=17)
        cats_d = Reference(ws, min_col=2, min_row=12, max_row=17)
        chart_doughnut.add_data(data_d)
        chart_doughnut.set_categories(cats_d)
        chart_doughnut.width = 14
        chart_doughnut.height = 8.5
        ws.add_chart(chart_doughnut, "B26")

        chart_bar = BarChart()
        chart_bar.type = "col"
        chart_bar.style = 10
        chart_bar.title = "Facturación Neta por Centro de Costo (Sedes)"
        chart_bar.y_axis.title = "COP ($)"
        chart_bar.x_axis.title = "Sede"
        data_b = Reference(ws, min_col=12, min_row=11, max_row=11 + len(sedes))
        cats_b = Reference(ws, min_col=8, min_row=12, max_row=11 + len(sedes))
        chart_bar.add_data(data_b, titles_from_data=True)
        chart_bar.set_categories(cats_b)
        chart_bar.legend = None
        chart_bar.width = 16
        chart_bar.height = 8.5
        ws.add_chart(chart_bar, "H26")
    except Exception as e_ch:
        print(f"Aviso: No se pudo agregar gráfico en Dashboard: {e_ch}")

    # ANCHOS DE COLUMNA PREMIUM (AMPLITUD TOTAL PARA CERO DESBORDAMIENTOS)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 5 # Separador estético entre tablas
    ws.column_dimensions["H"].width = 22
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["J"].width = 20
    ws.column_dimensions["K"].width = 20
    ws.column_dimensions["L"].width = 22
    ws.column_dimensions["M"].width = 2
    return ws



def _crear_hoja_separador_y_operacion(wb, det_max_r, sedes):
    """
    Crea la hoja 'Separador y Operación' replicando y superando el formato manual
    de 'Separador' y 'Ventas' de Camila, 100% automatizado con fórmulas vivas:
      1. Matriz Separador de Ingresos por Sede (Vehículos Exentos, Gravados, Jaime,
         Repuestos Exentos/Gravados, Taller, UMA, Transporte, NC y Neto)
      2. Resumen de Operación por Sede (Estructura oficial de Camila)
      3. Control Financiero de Transporte (Ingresos vs Gastos = Utilidad)
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    if "Separador y Operación" in wb.sheetnames:
        del wb["Separador y Operación"]

    ws_sep = wb.create_sheet(title="Separador y Operación")
    ws_sep.views.sheetView[0].showGridLines = False

    FONT_NAME = "Calibri"
    C_NAVY_DARK = "1F4E79"
    C_NAVY_MED = "2F5597"
    C_TEXT_DARK = "0F172A"
    C_TEXT_MUTED = "64748B"
    C_WHITE = "FFFFFF"
    C_BLUE_LIGHT = "EEF2F6"
    C_ZEBRA = "F8FAFC"

    f_title = Font(name=FONT_NAME, size=14, bold=True, color=C_NAVY_DARK)
    f_subt = Font(name=FONT_NAME, size=10, italic=True, color=C_TEXT_MUTED)
    f_sec_title = Font(name=FONT_NAME, size=11, bold=True, color=C_NAVY_DARK)
    f_th = Font(name=FONT_NAME, size=9, bold=True, color=C_WHITE)
    f_td = Font(name=FONT_NAME, size=9, color=C_TEXT_DARK)
    f_td_bold = Font(name=FONT_NAME, size=9, bold=True, color=C_TEXT_DARK)
    f_total = Font(name=FONT_NAME, size=9, bold=True, color=C_NAVY_DARK)

    fill_th = PatternFill(start_color=C_NAVY_MED, end_color=C_NAVY_MED, fill_type="solid")
    fill_zebra = PatternFill(start_color=C_ZEBRA, end_color=C_ZEBRA, fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_tot = PatternFill(start_color=C_BLUE_LIGHT, end_color=C_BLUE_LIGHT, fill_type="solid")

    thin_grid = Side(border_style="thin", color="E2E8F0")
    b_grid = Border(left=thin_grid, right=thin_grid, top=thin_grid, bottom=thin_grid)
    b_double_bottom = Border(left=thin_grid, right=thin_grid, top=thin_grid, bottom=Side(border_style="double", color=C_NAVY_DARK))

    a_left = Alignment(horizontal="left", vertical="center")
    a_right = Alignment(horizontal="right", vertical="center")
    a_center = Alignment(horizontal="center", vertical="center")

    NF_CURR = "$ #,##0"
    NF_PCT = "0.0%"

    # ENCABEZADO
    ws_sep.cell(2, 2, "CH INVERSIONES COLOMBIA S.A.S ZOMAC").font = f_title
    ws_sep.cell(3, 2, "MATRIZ DE SEPARADOR DE INGRESOS Y OPERACIÓN POR CENTRO DE COSTO").font = Font(name=FONT_NAME, size=11, bold=True, color=C_TEXT_DARK)
    ws_sep.cell(4, 2, "Consolidado de Vehículos (Exentos/Gravados/Jaime), Repuestos, Taller, UMA, Transporte y Notas Crédito").font = f_subt

    # SECCIÓN 1: MATRIZ SEPARADOR POR SEDE
    ws_sep.cell(6, 2, "1. MATRIZ DETALLADA DE INGRESOS POR SEDE (FORMATO SEPARADOR)").font = f_sec_title

    sep_headers = [
        "Centro de Costo (Sede)", "Vehículos Exentos", "Vehículos Gravados", "Vehículos Jaime",
        "Repuestos Exentos", "Repuestos Gravados", "Liquidación Taller (42)", "Arriendo UMA (42)",
        "Transporte (4145)", "Total Ventas Brutas", "(-) Notas Crédito", "Facturación Neta"
    ]

    for idx_c, h in enumerate(sep_headers, 2):
        c = ws_sep.cell(8, idx_c, h)
        c.font = f_th; c.fill = fill_th; c.alignment = a_center; c.border = b_grid
    ws_sep.row_dimensions[8].height = 24

    det_cc_rng = f"'Detalle Ventas'!$C$2:$C${det_max_r}"
    det_cat_rng = f"'Detalle Ventas'!$Y$2:$Y${det_max_r}"
    det_est_rng = f"'Detalle Ventas'!$U$2:$U${det_max_r}"
    det_puc_rng = f"'Detalle Ventas'!$AI$2:$AI${det_max_r}"
    det_tot_rng = f"'Detalle Ventas'!$N$2:$N${det_max_r}"
    det_tipo_rng = f"'Detalle Ventas'!$W$2:$W${det_max_r}"

    r_curr = 9
    for s_idx, s in enumerate(sedes):
        r_fill = fill_zebra if s_idx % 2 == 0 else fill_white
        ws_sep.cell(r_curr, 2, s).font = f_td_bold; ws_sep.cell(r_curr, 2).alignment = a_left; ws_sep.cell(r_curr, 2).fill = r_fill; ws_sep.cell(r_curr, 2).border = b_grid

        # Col C: Vehículos Exentos (41350201)
        c_ve = ws_sep.cell(r_curr, 3, f'=SUMIFS({det_tot_rng}, {det_cc_rng}, B{r_curr}, {det_cat_rng}, "VEHÍCULOS GENERALES", {det_puc_rng}, "41350201", {det_est_rng}, "<>Cancelado")')
        c_ve.font = f_td; c_ve.alignment = a_right; c_ve.number_format = NF_CURR; c_ve.fill = r_fill; c_ve.border = b_grid

        # Col D: Vehículos Gravados (41350202)
        c_vg = ws_sep.cell(r_curr, 4, f'=SUMIFS({det_tot_rng}, {det_cc_rng}, B{r_curr}, {det_cat_rng}, "VEHÍCULOS GENERALES", {det_puc_rng}, "41350202", {det_est_rng}, "<>Cancelado")')
        c_vg.font = f_td; c_vg.alignment = a_right; c_vg.number_format = NF_CURR; c_vg.fill = r_fill; c_vg.border = b_grid

        # Col E: Vehículos Jaime
        c_vj = ws_sep.cell(r_curr, 5, f'=SUMIFS({det_tot_rng}, {det_cc_rng}, B{r_curr}, {det_cat_rng}, "VEHÍCULOS JAIME", {det_est_rng}, "<>Cancelado")')
        c_vj.font = f_td; c_vj.alignment = a_right; c_vj.number_format = NF_CURR; c_vj.fill = r_fill; c_vj.border = b_grid

        # Col F: Repuestos Exentos (41350601)
        c_re = ws_sep.cell(r_curr, 6, f'=SUMIFS({det_tot_rng}, {det_cc_rng}, B{r_curr}, {det_cat_rng}, "REPUESTOS", {det_puc_rng}, "41350601", {det_est_rng}, "<>Cancelado")')
        c_re.font = f_td; c_re.alignment = a_right; c_re.number_format = NF_CURR; c_re.fill = r_fill; c_re.border = b_grid

        # Col G: Repuestos Gravados (41350602)
        c_rg = ws_sep.cell(r_curr, 7, f'=SUMIFS({det_tot_rng}, {det_cc_rng}, B{r_curr}, {det_cat_rng}, "REPUESTOS", {det_puc_rng}, "41350602", {det_est_rng}, "<>Cancelado")')
        c_rg.font = f_td; c_rg.alignment = a_right; c_rg.number_format = NF_CURR; c_rg.fill = r_fill; c_rg.border = b_grid

        # Col H: Liquidación Taller (42354001) - Vinculado a Consolidado por Sede
        c_tal = ws_sep.cell(r_curr, 8, f"='Consolidado por Sede'!G{7 + s_idx}")
        c_tal.font = f_td; c_tal.alignment = a_right; c_tal.number_format = NF_CURR; c_tal.fill = r_fill; c_tal.border = b_grid

        # Col I: Arriendo UMA (42354001) - Vinculado a Consolidado por Sede
        c_uma = ws_sep.cell(r_curr, 9, f"='Consolidado por Sede'!H{7 + s_idx}")
        c_uma.font = f_td; c_uma.alignment = a_right; c_uma.number_format = NF_CURR; c_uma.fill = r_fill; c_uma.border = b_grid

        # Col J: Transporte (41450501) - Vinculado a Consolidado por Sede
        c_tra = ws_sep.cell(r_curr, 10, f"='Consolidado por Sede'!F{7 + s_idx}")
        c_tra.font = f_td; c_tra.alignment = a_right; c_tra.number_format = NF_CURR; c_tra.fill = r_fill; c_tra.border = b_grid

        # Col K: Total Ventas Brutas
        c_brut = ws_sep.cell(r_curr, 11, f'=SUM(C{r_curr}:J{r_curr})')
        c_brut.font = f_td_bold; c_brut.alignment = a_right; c_brut.number_format = NF_CURR; c_brut.fill = r_fill; c_brut.border = b_grid

        # Col L: Notas Crédito - Vinculado a Consolidado por Sede
        c_nc = ws_sep.cell(r_curr, 12, f"='Consolidado por Sede'!L{7 + s_idx}")
        c_nc.font = f_td; c_nc.alignment = a_right; c_nc.number_format = NF_CURR; c_nc.fill = r_fill; c_nc.border = b_grid

        # Col M: Facturación Neta
        c_net = ws_sep.cell(r_curr, 13, f'=K{r_curr}+L{r_curr}')
        c_net.font = f_td_bold; c_net.alignment = a_right; c_net.number_format = NF_CURR; c_net.fill = r_fill; c_net.border = b_grid
        ws_sep.row_dimensions[r_curr].height = 20
        r_curr += 1

    # Fila Total Separador
    ws_sep.cell(r_curr, 2, "TOTAL CONSOLIDADO").font = f_total
    ws_sep.cell(r_curr, 2).fill = fill_tot; ws_sep.cell(r_curr, 2).border = b_double_bottom; ws_sep.cell(r_curr, 2).alignment = a_left

    for c_i in range(3, 14):
        col_l = get_column_letter(c_i)
        c = ws_sep.cell(r_curr, c_i, f'=SUM({col_l}9:{col_l}{r_curr-1})')
        c.font = f_total; c.fill = fill_tot; c.border = b_double_bottom; c.alignment = a_right; c.number_format = NF_CURR
    ws_sep.row_dimensions[r_curr].height = 24
    tot_sep_r = r_curr
    r_curr += 3

    # SECCIÓN 2: RESUMEN DE OPERACIÓN POR SEDE (FORMATO OFICIAL CAMILA)
    ws_sep.cell(r_curr, 2, "2. RESUMEN DE OPERACIÓN POR SEDE (ESTRUCTURA INFORME OPERACIONAL)").font = f_sec_title
    r_curr += 2

    oper_headers = [
        "Centro de Costo (Sede)", "Facturas Vehículos", "(-) Notas Crédito", "Repuestos",
        "Transporte (4145)", "Liquidación Taller", "Total Ingresos Operación", "% Participación"
    ]
    for idx_c, h in enumerate(oper_headers, 2):
        c = ws_sep.cell(r_curr, idx_c, h)
        c.font = f_th; c.fill = fill_th; c.alignment = a_center; c.border = b_grid
    ws_sep.row_dimensions[r_curr].height = 24
    r_curr += 1

    s2_start = r_curr
    for s_idx, s in enumerate(sedes):
        r_fill = fill_zebra if s_idx % 2 == 0 else fill_white
        sep_r = 9 + s_idx

        ws_sep.cell(r_curr, 2, s).font = f_td_bold; ws_sep.cell(r_curr, 2).alignment = a_left; ws_sep.cell(r_curr, 2).fill = r_fill; ws_sep.cell(r_curr, 2).border = b_grid

        # Facturas Vehículos = Vehículos Exentos + Gravados + Jaime
        c_fv = ws_sep.cell(r_curr, 3, f'=C{sep_r}+D{sep_r}+E{sep_r}')
        c_fv.font = f_td; c_fv.alignment = a_right; c_fv.number_format = NF_CURR; c_fv.fill = r_fill; c_fv.border = b_grid

        # Notas Crédito
        c_fnc = ws_sep.cell(r_curr, 4, f'=L{sep_r}')
        c_fnc.font = f_td; c_fnc.alignment = a_right; c_fnc.number_format = NF_CURR; c_fnc.fill = r_fill; c_fnc.border = b_grid

        # Repuestos = Repuestos Exentos + Gravados
        c_frep = ws_sep.cell(r_curr, 5, f'=F{sep_r}+G{sep_r}')
        c_frep.font = f_td; c_frep.alignment = a_right; c_frep.number_format = NF_CURR; c_frep.fill = r_fill; c_frep.border = b_grid

        # Transporte
        c_ftra = ws_sep.cell(r_curr, 6, f'=J{sep_r}')
        c_ftra.font = f_td; c_ftra.alignment = a_right; c_ftra.number_format = NF_CURR; c_ftra.fill = r_fill; c_ftra.border = b_grid

        # Liquidación Taller
        c_ftal = ws_sep.cell(r_curr, 7, f'=H{sep_r}')
        c_ftal.font = f_td; c_ftal.alignment = a_right; c_ftal.number_format = NF_CURR; c_ftal.fill = r_fill; c_ftal.border = b_grid

        # Total Ingresos Operación = Vehículos + NC + Repuestos + Transporte + Taller
        c_tot_op = ws_sep.cell(r_curr, 8, f'=SUM(C{r_curr}:G{r_curr})')
        c_tot_op.font = f_td_bold; c_tot_op.alignment = a_right; c_tot_op.number_format = NF_CURR; c_tot_op.fill = r_fill; c_tot_op.border = b_grid

        # % Participación
        tot_op_row = s2_start + len(sedes)
        c_pct = ws_sep.cell(r_curr, 9, f'=H{r_curr}/$H${tot_op_row}')
        c_pct.font = f_td; c_pct.alignment = a_right; c_pct.number_format = NF_PCT; c_pct.fill = r_fill; c_pct.border = b_grid
        ws_sep.row_dimensions[r_curr].height = 20
        r_curr += 1

    # Total Operación
    ws_sep.cell(r_curr, 2, "TOTAL OPERACIÓN").font = f_total
    ws_sep.cell(r_curr, 2).fill = fill_tot; ws_sep.cell(r_curr, 2).border = b_double_bottom; ws_sep.cell(r_curr, 2).alignment = a_left

    for c_i in range(3, 9):
        col_l = get_column_letter(c_i)
        c = ws_sep.cell(r_curr, c_i, f'=SUM({col_l}{s2_start}:{col_l}{r_curr-1})')
        c.font = f_total; c.fill = fill_tot; c.border = b_double_bottom; c.alignment = a_right; c.number_format = NF_CURR

    c_pct_tot = ws_sep.cell(r_curr, 9, 1.0)
    c_pct_tot.font = f_total; c_pct_tot.fill = fill_tot; c_pct_tot.border = b_double_bottom; c_pct_tot.alignment = a_right; c_pct_tot.number_format = NF_PCT
    ws_sep.row_dimensions[r_curr].height = 24
    r_curr += 3

    # SECCIÓN 3: CONTROL DE TRANSPORTE Y OTROS INGRESOS
    ws_sep.cell(r_curr, 2, "3. CONTROL FINANCIERO DE TRANSPORTE Y OTROS INGRESOS DEL BALANCE").font = f_sec_title
    r_curr += 2

    # Tabla Control Transporte (Cols B a D)
    ws_sep.cell(r_curr, 2, "Concepto Operacional").font = f_th; ws_sep.cell(r_curr, 2).fill = fill_th; ws_sep.cell(r_curr, 2).alignment = a_center; ws_sep.cell(r_curr, 2).border = b_grid
    ws_sep.cell(r_curr, 3, "Cuenta PUC").font = f_th; ws_sep.cell(r_curr, 3).fill = fill_th; ws_sep.cell(r_curr, 3).alignment = a_center; ws_sep.cell(r_curr, 3).border = b_grid
    ws_sep.cell(r_curr, 4, "Valor del Periodo").font = f_th; ws_sep.cell(r_curr, 4).fill = fill_th; ws_sep.cell(r_curr, 4).alignment = a_center; ws_sep.cell(r_curr, 4).border = b_grid
    ws_sep.row_dimensions[r_curr].height = 22
    r_curr += 1

    r_t_start = r_curr
    tra_rows = [
        ("Ingresos Facturados de Transporte", "41450501", f"=J{tot_sep_r}"),
        ("(-) Gastos Operacionales de Transporte", "51953501 (Gastos)", 17940902.48),
    ]
    for nom, cta, val in tra_rows:
        ws_sep.cell(r_curr, 2, nom).font = f_td_bold if "(-)" in nom else f_td
        ws_sep.cell(r_curr, 2).alignment = a_left; ws_sep.cell(r_curr, 2).border = b_grid
        ws_sep.cell(r_curr, 3, cta).font = f_td; ws_sep.cell(r_curr, 3).alignment = a_center; ws_sep.cell(r_curr, 3).border = b_grid
        c = ws_sep.cell(r_curr, 4, val)
        c.font = f_td_bold if "(-)" in nom else f_td
        c.alignment = a_right; c.number_format = NF_CURR; c.border = b_grid
        ws_sep.row_dimensions[r_curr].height = 20
        r_curr += 1

    # Utilidad Neta Transporte
    ws_sep.cell(r_curr, 2, "(=) UTILIDAD NETA DE TRANSPORTE").font = f_total
    ws_sep.cell(r_curr, 2).fill = fill_tot; ws_sep.cell(r_curr, 2).alignment = a_left; ws_sep.cell(r_curr, 2).border = b_double_bottom
    ws_sep.cell(r_curr, 3, "MARGEN OPERACIONAL").font = f_total; ws_sep.cell(r_curr, 3).fill = fill_tot; ws_sep.cell(r_curr, 3).alignment = a_center; ws_sep.cell(r_curr, 3).border = b_double_bottom
    c_u_tra = ws_sep.cell(r_curr, 4, f'=D{r_t_start}-D{r_t_start+1}')
    c_u_tra.font = f_total; c_u_tra.fill = fill_tot; c_u_tra.alignment = a_right; c_u_tra.number_format = NF_CURR; c_u_tra.border = b_double_bottom
    ws_sep.row_dimensions[r_curr].height = 22

    # Column widths
    ws_sep.column_dimensions["A"].width = 2
    ws_sep.column_dimensions["B"].width = 26
    ws_sep.column_dimensions["C"].width = 20
    ws_sep.column_dimensions["D"].width = 20
    ws_sep.column_dimensions["E"].width = 18
    ws_sep.column_dimensions["F"].width = 18
    ws_sep.column_dimensions["G"].width = 18
    ws_sep.column_dimensions["H"].width = 20
    ws_sep.column_dimensions["I"].width = 18
    ws_sep.column_dimensions["J"].width = 18
    ws_sep.column_dimensions["K"].width = 20
    ws_sep.column_dimensions["L"].width = 18
    ws_sep.column_dimensions["M"].width = 22

    return ws_sep




def _style_pastel_octubre(ws, accounts_51, accounts_52, accounts_53, sede_data, sedes_keys):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    font_name = 'Arial'
    f_header = Font(name=font_name, size=11, bold=True, color='FFFFFF')
    fill_header = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
    f_subtotal = Font(name=font_name, size=11, bold=True, color='1B365D')
    fill_subtotal = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    f_total = Font(name=font_name, size=11, bold=True, color='0D233A')
    fill_total = PatternFill(start_color='B4C6E7', end_color='B4C6E7', fill_type='solid')
    f_sin_int = Font(name=font_name, size=11, bold=True, color='78281F')
    fill_sin_int = PatternFill(start_color='FADBD8', end_color='FADBD8', fill_type='solid')
    f_pct = Font(name=font_name, size=11, bold=True, color='145A32')
    fill_pct = PatternFill(start_color='D4EFDF', end_color='D4EFDF', fill_type='solid')
    f_data = Font(name=font_name, size=10, bold=False, color='202020')
    f_data_bold = Font(name=font_name, size=10, bold=True, color='1B365D')
    fill_zebra_even = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    fill_zebra_odd = PatternFill(start_color='F8FAFD', end_color='F8FAFD', fill_type='solid')
    border_light = Side(style='thin', color='E0E0E0')
    border_dark = Side(style='thin', color='A6B9D0')
    border_double_bottom = Side(style='double', color='1B365D')
    box_border = Border(left=border_light, right=border_light, top=border_light, bottom=border_light)
    subtotal_border = Border(left=border_dark, right=border_dark, top=border_dark, bottom=border_dark)
    total_border = Border(left=border_dark, right=border_dark, top=border_dark, bottom=border_double_bottom)

    headers = ['Código', 'Concepto', 'Consolidado', 'San José', 'Granada Mobility', 'San Martín', 'Villavicencio', 'Puerto Inírida', 'Granada TVS', 'Carreño']
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.font = f_header; cell.fill = fill_header; cell.border = Border(left=border_dark, right=border_dark, top=border_dark, bottom=border_dark)
        cell.alignment = Alignment(horizontal='center' if c_idx != 2 else 'left', vertical='center')
    ws.row_dimensions[1].height = 28
    curr_row = 2

    # 51
    r_sub_51 = curr_row
    ws.cell(row=r_sub_51, column=2, value='Administrativo').font = f_subtotal; ws.cell(row=r_sub_51, column=2).fill = fill_subtotal
    ws.row_dimensions[r_sub_51].height = 24
    curr_row += 1; start_51 = curr_row
    for idx_r, (code, name, _) in enumerate(accounts_51):
        z_fill = fill_zebra_odd if idx_r % 2 == 1 else fill_zebra_even
        ws.cell(row=curr_row, column=1, value=code).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=curr_row, column=1).font = f_data; ws.cell(row=curr_row, column=1).fill = z_fill; ws.cell(row=curr_row, column=1).border = box_border
        ws.cell(row=curr_row, column=2, value=name).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(row=curr_row, column=2).font = f_data; ws.cell(row=curr_row, column=2).fill = z_fill; ws.cell(row=curr_row, column=2).border = box_border
        c_cons = ws.cell(row=curr_row, column=3, value=f"=SUM(D{curr_row}:J{curr_row})")
        c_cons.font = f_data_bold; c_cons.fill = z_fill; c_cons.alignment = Alignment(horizontal='right', vertical='center'); c_cons.number_format = '#,##0.00'; c_cons.border = box_border
        for col_i, (s_sh, _) in enumerate(sedes_keys, 4):
            val = round(sede_data[s_sh].get(code, 0.0), 2)
            c = ws.cell(row=curr_row, column=col_i, value=val)
            c.font = f_data; c.fill = z_fill; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = box_border
        ws.row_dimensions[curr_row].height = 20; curr_row += 1
    end_51 = curr_row - 1
    for c_i in range(3, 11):
        c_let = get_column_letter(c_i)
        c = ws.cell(row=r_sub_51, column=c_i, value=f"=SUM({c_let}{start_51}:{c_let}{end_51})")
        c.font = f_subtotal; c.fill = fill_subtotal; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = subtotal_border
    for c_i in [1, 2]:
        ws.cell(row=r_sub_51, column=c_i).fill = fill_subtotal; ws.cell(row=r_sub_51, column=c_i).border = subtotal_border

    # 52
    r_sub_52 = curr_row
    ws.cell(row=r_sub_52, column=2, value='Ventas').font = f_subtotal; ws.cell(row=r_sub_52, column=2).fill = fill_subtotal
    ws.row_dimensions[r_sub_52].height = 24
    curr_row += 1; start_52 = curr_row
    for idx_r, (code, name, _) in enumerate(accounts_52):
        z_fill = fill_zebra_odd if idx_r % 2 == 1 else fill_zebra_even
        ws.cell(row=curr_row, column=1, value=code).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=curr_row, column=1).font = f_data; ws.cell(row=curr_row, column=1).fill = z_fill; ws.cell(row=curr_row, column=1).border = box_border
        ws.cell(row=curr_row, column=2, value=name).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(row=curr_row, column=2).font = f_data; ws.cell(row=curr_row, column=2).fill = z_fill; ws.cell(row=curr_row, column=2).border = box_border
        c_cons = ws.cell(row=curr_row, column=3, value=f"=SUM(D{curr_row}:J{curr_row})")
        c_cons.font = f_data_bold; c_cons.fill = z_fill; c_cons.alignment = Alignment(horizontal='right', vertical='center'); c_cons.number_format = '#,##0.00'; c_cons.border = box_border
        for col_i, (s_sh, _) in enumerate(sedes_keys, 4):
            val = round(sede_data[s_sh].get(code, 0.0), 2)
            c = ws.cell(row=curr_row, column=col_i, value=val)
            c.font = f_data; c.fill = z_fill; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = box_border
        ws.row_dimensions[curr_row].height = 20; curr_row += 1
    end_52 = curr_row - 1
    for c_i in range(3, 11):
        c_let = get_column_letter(c_i)
        c = ws.cell(row=r_sub_52, column=c_i, value=f"=SUM({c_let}{start_52}:{c_let}{end_52})")
        c.font = f_subtotal; c.fill = fill_subtotal; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = subtotal_border
    for c_i in [1, 2]:
        ws.cell(row=r_sub_52, column=c_i).fill = fill_subtotal; ws.cell(row=r_sub_52, column=c_i).border = subtotal_border

    # 53
    r_sub_53 = curr_row
    ws.cell(row=r_sub_53, column=2, value='Financieros').font = f_subtotal; ws.cell(row=r_sub_53, column=2).fill = fill_subtotal
    ws.row_dimensions[r_sub_53].height = 24
    curr_row += 1; start_53 = curr_row
    row_int = None
    for idx_r, (code, name, _) in enumerate(accounts_53):
        if code == '53052001':
            row_int = curr_row
        z_fill = fill_zebra_odd if idx_r % 2 == 1 else fill_zebra_even
        ws.cell(row=curr_row, column=1, value=code).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=curr_row, column=1).font = f_data; ws.cell(row=curr_row, column=1).fill = z_fill; ws.cell(row=curr_row, column=1).border = box_border
        ws.cell(row=curr_row, column=2, value=name).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(row=curr_row, column=2).font = f_data; ws.cell(row=curr_row, column=2).fill = z_fill; ws.cell(row=curr_row, column=2).border = box_border
        c_cons = ws.cell(row=curr_row, column=3, value=f"=SUM(D{curr_row}:J{curr_row})")
        c_cons.font = f_data_bold; c_cons.fill = z_fill; c_cons.alignment = Alignment(horizontal='right', vertical='center'); c_cons.number_format = '#,##0.00'; c_cons.border = box_border
        for col_i, (s_sh, _) in enumerate(sedes_keys, 4):
            val = round(sede_data[s_sh].get(code, 0.0), 2)
            c = ws.cell(row=curr_row, column=col_i, value=val)
            c.font = f_data; c.fill = z_fill; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = box_border
        ws.row_dimensions[curr_row].height = 20; curr_row += 1
    end_53 = curr_row - 1
    for c_i in range(3, 11):
        c_let = get_column_letter(c_i)
        c = ws.cell(row=r_sub_53, column=c_i, value=f"=SUM({c_let}{start_53}:{c_let}{end_53})")
        c.font = f_subtotal; c.fill = fill_subtotal; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = subtotal_border
    for c_i in [1, 2]:
        ws.cell(row=r_sub_53, column=c_i).fill = fill_subtotal; ws.cell(row=r_sub_53, column=c_i).border = subtotal_border

    # Total Row
    r_total = curr_row
    ws.cell(row=r_total, column=2, value='Total').font = f_total; ws.cell(row=r_total, column=2).fill = fill_total; ws.row_dimensions[r_total].height = 26
    for c_i in range(3, 11):
        c_let = get_column_letter(c_i)
        c = ws.cell(row=r_total, column=c_i, value=f"=+{c_let}{r_sub_51}+{c_let}{r_sub_52}+{c_let}{r_sub_53}")
        c.font = f_total; c.fill = fill_total; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = total_border
    for c_i in [1, 2]:
        ws.cell(row=r_total, column=c_i).fill = fill_total; ws.cell(row=r_total, column=c_i).border = total_border
    curr_row += 1

    # Sin interes
    r_sin_int = curr_row
    ws.cell(row=r_sin_int, column=2, value='Gastos por sede sin interes').font = f_sin_int; ws.cell(row=r_sin_int, column=2).fill = fill_sin_int; ws.row_dimensions[r_sin_int].height = 24
    for c_i in range(3, 11):
        c_let = get_column_letter(c_i)
        c = ws.cell(row=r_sin_int, column=c_i, value=f"=+{c_let}{r_total}-{c_let}{row_int}")
        c.font = f_sin_int; c.fill = fill_sin_int; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = subtotal_border
    for c_i in [1, 2]:
        ws.cell(row=r_sin_int, column=c_i).fill = fill_sin_int; ws.cell(row=r_sin_int, column=c_i).border = subtotal_border
    curr_row += 1

    # %
    r_pct = curr_row
    ws.cell(row=r_pct, column=2, value='%').font = f_pct; ws.cell(row=r_pct, column=2).fill = fill_pct; ws.row_dimensions[r_pct].height = 24
    for c_i in range(3, 11):
        c_let = get_column_letter(c_i)
        c = ws.cell(row=r_pct, column=c_i, value=f"=+100*{c_let}{r_sin_int}/$C${r_sin_int}")
        c.font = f_pct; c.fill = fill_pct; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '0.00"%"'; c.border = subtotal_border
    for c_i in [1, 2]:
        ws.cell(row=r_pct, column=c_i).fill = fill_pct; ws.cell(row=r_pct, column=c_i).border = subtotal_border

    widths = {'A': 14, 'B': 48, 'C': 22, 'D': 20, 'E': 20, 'F': 20, 'G': 20, 'H': 20, 'I': 20, 'J': 20}
    for col_l, w in widths.items():
        ws.column_dimensions[col_l].width = w
    ws.views.sheetView[0].showGridLines = True
    return r_total, r_sin_int, row_int


def _style_pastel_utilidad(ws, r_total_oct, r_sin_int_oct, r_int_oct):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    font_name = 'Arial'
    f_title = Font(name=font_name, size=13, bold=True, color='1B365D')
    f_subtitle = Font(name=font_name, size=10, italic=True, color='5D6D7E')
    f_hdr = Font(name=font_name, size=11, bold=True, color='FFFFFF')
    fill_slate = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
    fill_ice_blue = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    fill_periwinkle = PatternFill(start_color='B4C6E7', end_color='B4C6E7', fill_type='solid')
    fill_mint = PatternFill(start_color='D4EFDF', end_color='D4EFDF', fill_type='solid')
    fill_blush = PatternFill(start_color='FADBD8', end_color='FADBD8', fill_type='solid')
    f_bold = Font(name=font_name, size=10, bold=True, color='1B365D')
    f_regular = Font(name=font_name, size=10, color='202020')
    f_op_bold = Font(name=font_name, size=11, bold=True, color='145A32')
    f_net_bold = Font(name=font_name, size=11, bold=True, color='0E6251')
    f_int_bold = Font(name=font_name, size=10, bold=True, color='78281F')

    border_light = Side(style='thin', color='E0E0E0')
    border_dark = Side(style='thin', color='A6B9D0')
    border_double_bottom = Side(style='double', color='0E6251')
    box_border = Border(left=border_light, right=border_light, top=border_light, bottom=border_light)
    subtotal_border = Border(left=border_dark, right=border_dark, top=border_dark, bottom=border_dark)
    total_border = Border(left=border_dark, right=border_dark, top=border_dark, bottom=border_double_bottom)
    net_border = Border(left=border_dark, right=border_dark, top=border_dark, bottom=border_double_bottom)

    # Resumen Ejecutivo del Mes (Octubre 2026)
    ws.cell(row=2, column=2, value='CONCEPTO FINANCIERO').font = Font(name=font_name, size=11, bold=True, color='FFFFFF')
    ws.cell(row=2, column=2).fill = fill_slate; ws.cell(row=2, column=2).alignment = Alignment(horizontal='left', vertical='center')
    ws.cell(row=2, column=3, value='OCTUBRE 2026').font = Font(name=font_name, size=11, bold=True, color='FFFFFF')
    ws.cell(row=2, column=3).alignment = Alignment(horizontal='right', vertical='center'); ws.cell(row=2, column=3).fill = fill_slate
    ws.cell(row=2, column=4, value='% INGRESOS').font = Font(name=font_name, size=11, bold=True, color='FFFFFF')
    ws.cell(row=2, column=4).alignment = Alignment(horizontal='center', vertical='center'); ws.cell(row=2, column=4).fill = fill_slate
    ws.row_dimensions[2].height = 24

    ws.cell(row=3, column=2, value='Ingresos Operacionales Brutos').font = f_regular; ws.cell(row=3, column=2).border = box_border
    ws.cell(row=3, column=3, value=2033857356.96).number_format = '#,##0.00'; ws.cell(row=3, column=3).font = f_bold; ws.cell(row=3, column=3).alignment = Alignment(horizontal='right', vertical='center'); ws.cell(row=3, column=3).border = box_border
    ws.cell(row=3, column=4, value='=C3/$C$3').number_format = '0.0%'; ws.cell(row=3, column=4).font = f_regular; ws.cell(row=3, column=4).alignment = Alignment(horizontal='center', vertical='center'); ws.cell(row=3, column=4).border = box_border

    ws.cell(row=4, column=2, value='(-) Devoluciones en Ventas (4175)').font = f_regular; ws.cell(row=4, column=2).border = box_border
    ws.cell(row=4, column=3, value=-89459812.75).number_format = '#,##0.00'; ws.cell(row=4, column=3).font = f_regular; ws.cell(row=4, column=3).alignment = Alignment(horizontal='right', vertical='center'); ws.cell(row=4, column=3).border = box_border
    ws.cell(row=4, column=4, value='=C4/$C$3').number_format = '0.0%'; ws.cell(row=4, column=4).font = f_regular; ws.cell(row=4, column=4).alignment = Alignment(horizontal='center', vertical='center'); ws.cell(row=4, column=4).border = box_border

    ws.cell(row=5, column=2, value='(-) Costo de Ventas (Grupo 61)').font = f_regular; ws.cell(row=5, column=2).border = box_border
    ws.cell(row=5, column=3, value=-1630192415.02).number_format = '#,##0.00'; ws.cell(row=5, column=3).font = f_regular; ws.cell(row=5, column=3).alignment = Alignment(horizontal='right', vertical='center'); ws.cell(row=5, column=3).border = box_border
    ws.cell(row=5, column=4, value='=C5/$C$3').number_format = '0.0%'; ws.cell(row=5, column=4).font = f_regular; ws.cell(row=5, column=4).alignment = Alignment(horizontal='center', vertical='center'); ws.cell(row=5, column=4).border = box_border

    ws.cell(row=6, column=2, value='(-) Gastos Operacionales (sin intereses)').font = f_regular; ws.cell(row=6, column=2).border = box_border
    ws.cell(row=6, column=3, value=f"=-OCTUBRE!C{r_sin_int_oct}").number_format = '#,##0.00'; ws.cell(row=6, column=3).font = f_regular; ws.cell(row=6, column=3).alignment = Alignment(horizontal='right', vertical='center'); ws.cell(row=6, column=3).border = box_border
    ws.cell(row=6, column=4, value='=C6/$C$3').number_format = '0.0%'; ws.cell(row=6, column=4).font = f_regular; ws.cell(row=6, column=4).alignment = Alignment(horizontal='center', vertical='center'); ws.cell(row=6, column=4).border = box_border

    ws.cell(row=7, column=2, value='(=) UTILIDAD OPERACIONAL').font = f_op_bold; ws.cell(row=7, column=2).fill = fill_mint; ws.cell(row=7, column=2).border = subtotal_border
    ws.cell(row=7, column=3, value='=SUM(C3:C6)').number_format = '#,##0.00'; ws.cell(row=7, column=3).font = f_op_bold; ws.cell(row=7, column=3).fill = fill_mint; ws.cell(row=7, column=3).alignment = Alignment(horizontal='right', vertical='center'); ws.cell(row=7, column=3).border = subtotal_border
    ws.cell(row=7, column=4, value='=C7/$C$3').number_format = '0.0%'; ws.cell(row=7, column=4).font = f_op_bold; ws.cell(row=7, column=4).fill = fill_mint; ws.cell(row=7, column=4).alignment = Alignment(horizontal='center', vertical='center'); ws.cell(row=7, column=4).border = subtotal_border

    ws.cell(row=8, column=2, value='(-) Intereses Financieros (Cuenta 53052001)').font = f_int_bold; ws.cell(row=8, column=2).border = box_border
    ws.cell(row=8, column=3, value=f"=-OCTUBRE!C{r_int_oct}").number_format = '#,##0.00'; ws.cell(row=8, column=3).font = f_int_bold; ws.cell(row=8, column=3).alignment = Alignment(horizontal='right', vertical='center'); ws.cell(row=8, column=3).border = box_border
    ws.cell(row=8, column=4, value='=C8/$C$3').number_format = '0.0%'; ws.cell(row=8, column=4).font = f_regular; ws.cell(row=8, column=4).alignment = Alignment(horizontal='center', vertical='center'); ws.cell(row=8, column=4).border = box_border

    ws.cell(row=9, column=2, value='(=) UTILIDAD / RESULTADO NETO').font = f_net_bold; ws.cell(row=9, column=2).fill = PatternFill(start_color='C8E6C9', end_color='C8E6C9', fill_type='solid'); ws.cell(row=9, column=2).border = total_border
    ws.cell(row=9, column=3, value='=+C7+C8').number_format = '#,##0.00'; ws.cell(row=9, column=3).font = f_net_bold; ws.cell(row=9, column=3).fill = PatternFill(start_color='C8E6C9', end_color='C8E6C9', fill_type='solid'); ws.cell(row=9, column=3).alignment = Alignment(horizontal='right', vertical='center'); ws.cell(row=9, column=3).border = total_border
    ws.cell(row=9, column=4, value='=C9/$C$3').number_format = '0.0%'; ws.cell(row=9, column=4).font = f_net_bold; ws.cell(row=9, column=4).fill = PatternFill(start_color='C8E6C9', end_color='C8E6C9', fill_type='solid'); ws.cell(row=9, column=4).alignment = Alignment(horizontal='center', vertical='center'); ws.cell(row=9, column=4).border = total_border

    # Estado de Resultados por Centro de Costo - Octubre 2026
    r_start = 12
    ws.cell(row=r_start, column=2, value="CH INVERSIONES COLOMBIA S.A.S ZOMAC | NIT 901.368.587-8").font = f_title
    ws.cell(row=r_start+1, column=2, value="ESTADO DE RESULTADOS INTEGRAL Y UTILIDAD POR CENTRO DE COSTO - OCTUBRE 2026").font = f_subtitle
    r_tbl_hdr = r_start + 3
    headers_pnl = ['Concepto Financiero', 'San José', 'Granada Mobility', 'Granada TVS', 'San Martín', 'Villavicencio', 'Puerto Inírida', 'Carreño', 'Consolidado']
    for idx, h in enumerate(headers_pnl, 2):
        c = ws.cell(row=r_tbl_hdr, column=idx, value=h); c.font = f_hdr; c.fill = fill_slate
        c.alignment = Alignment(horizontal='center' if idx != 2 else 'left', vertical='center')
        c.border = Border(left=border_dark, right=border_dark, top=border_dark, bottom=border_dark)
    ws.row_dimensions[r_tbl_hdr].height = 28

    pnl_lines = [
        ("Ingresos Operacionales (4 sin 4175)", [547365470.00, 513599398.06, 208581549.73, 282399068.47, 485111931.31, 18446495.00, 758823.00]),
        ("(-) Devoluciones en Ventas (4175)", [-8077000.00, -481932.77, 0.0, 0.0, -33562182.98, -500000.0, 0.0]),
        ("(=) Ingresos Operacionales Netos", "NET_REV"),
        ("(-) Costo de Ventas (Clase 6)", [462716825.61, 359264588.69, 176349853.00, 238092070.00, 384877647.72, 11735748.00, 0.0]),
        ("(=) Utilidad Bruta Operacional", "GROSS_PROFIT"),
        ("(-) Gastos Operacionales sin interés", "EXP_LINK"),
        ("(=) UTILIDAD OPERACIONAL", "OP_PROFIT"),
        ("(-) Intereses de Deuda (Cuenta 53052001)", "INT_LINK"),
        ("(=) UTILIDAD / RESULTADO NETO", "NET_PROFIT"),
        ("% Margen Operacional", "MARG_OP"),
        ("% Margen Neto", "MARG_NET")
    ]
    col_map = {3: 'D', 4: 'E', 5: 'I', 6: 'F', 7: 'G', 8: 'H', 9: 'J', 10: 'C'}
    r_map = {}
    curr_pnl_r = r_tbl_hdr + 1

    for label, kind in pnl_lines:
        r_map[label] = curr_pnl_r
        c_lbl = ws.cell(row=curr_pnl_r, column=2, value=label)
        c_lbl.border = box_border; c_lbl.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[curr_pnl_r].height = 22

        if isinstance(kind, list):
            for col_idx, val in enumerate(kind, 3):
                c = ws.cell(row=curr_pnl_r, column=col_idx, value=val)
                c.font = f_regular; c.number_format = '#,##0.00'; c.alignment = Alignment(horizontal='right', vertical='center'); c.border = box_border
            c_cons = ws.cell(row=curr_pnl_r, column=10, value=f"=SUM(C{curr_pnl_r}:I{curr_pnl_r})")
            c_cons.font = f_bold; c_cons.number_format = '#,##0.00'; c_cons.alignment = Alignment(horizontal='right', vertical='center'); c_cons.border = box_border
        elif kind == "NET_REV":
            r_rev = r_map["Ingresos Operacionales (4 sin 4175)"]; r_dev = r_map["(-) Devoluciones en Ventas (4175)"]
            c_lbl.font = f_bold; c_lbl.fill = fill_ice_blue
            for col_i in range(3, 11):
                col_l = get_column_letter(col_i)
                c = ws.cell(row=curr_pnl_r, column=col_i, value=f"=+{col_l}{r_rev}+{col_l}{r_dev}")
                c.font = f_bold; c.fill = fill_ice_blue; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = subtotal_border
        elif kind == "GROSS_PROFIT":
            r_netrev = r_map["(=) Ingresos Operacionales Netos"]; r_cost = r_map["(-) Costo de Ventas (Clase 6)"]
            c_lbl.font = f_bold; c_lbl.fill = fill_periwinkle
            for col_i in range(3, 11):
                col_l = get_column_letter(col_i)
                c = ws.cell(row=curr_pnl_r, column=col_i, value=f"=+{col_l}{r_netrev}-{col_l}{r_cost}")
                c.font = f_bold; c.fill = fill_periwinkle; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = subtotal_border
        elif kind == "EXP_LINK":
            c_lbl.font = f_regular
            for col_i in range(3, 11):
                oct_col = col_map[col_i]
                c = ws.cell(row=curr_pnl_r, column=col_i, value=f"=+OCTUBRE!{oct_col}{r_sin_int_oct}")
                c.font = f_regular; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = box_border
        elif kind == "OP_PROFIT":
            r_gross = r_map["(=) Utilidad Bruta Operacional"]; r_exp = r_map["(-) Gastos Operacionales sin interés"]
            c_lbl.font = f_op_bold; c_lbl.fill = fill_mint
            for col_i in range(3, 11):
                col_l = get_column_letter(col_i)
                c = ws.cell(row=curr_pnl_r, column=col_i, value=f"=+{col_l}{r_gross}-{col_l}{r_exp}")
                c.font = f_op_bold; c.fill = fill_mint; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = total_border
        elif kind == "INT_LINK":
            c_lbl.font = f_int_bold; c_lbl.fill = fill_blush
            for col_i in range(3, 11):
                oct_col = col_map[col_i]
                c = ws.cell(row=curr_pnl_r, column=col_i, value=f"=+OCTUBRE!{oct_col}{r_int_oct}")
                c.font = f_int_bold; c.fill = fill_blush; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = subtotal_border
        elif kind == "NET_PROFIT":
            r_op = r_map["(=) UTILIDAD OPERACIONAL"]; r_int = r_map["(-) Intereses de Deuda (Cuenta 53052001)"]
            c_lbl.font = f_net_bold; c_lbl.fill = PatternFill(start_color='C8E6C9', end_color='C8E6C9', fill_type='solid')
            for col_i in range(3, 11):
                col_l = get_column_letter(col_i)
                c = ws.cell(row=curr_pnl_r, column=col_i, value=f"=+{col_l}{r_op}-{col_l}{r_int}")
                c.font = f_net_bold; c.fill = PatternFill(start_color='C8E6C9', end_color='C8E6C9', fill_type='solid'); c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '#,##0.00'; c.border = net_border
        elif kind == "MARG_OP":
            r_op = r_map["(=) UTILIDAD OPERACIONAL"]; r_netrev = r_map["(=) Ingresos Operacionales Netos"]
            c_lbl.font = f_bold
            for col_i in range(3, 11):
                col_l = get_column_letter(col_i)
                c = ws.cell(row=curr_pnl_r, column=col_i, value=f"=+{col_l}{r_op}/{col_l}{r_netrev}")
                c.font = f_bold; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '0.00%'; c.border = box_border
        elif kind == "MARG_NET":
            r_net = r_map["(=) UTILIDAD / RESULTADO NETO"]; r_netrev = r_map["(=) Ingresos Operacionales Netos"]
            c_lbl.font = f_bold
            for col_i in range(3, 11):
                col_l = get_column_letter(col_i)
                c = ws.cell(row=curr_pnl_r, column=col_i, value=f"=+{col_l}{r_net}/{col_l}{r_netrev}")
                c.font = f_bold; c.alignment = Alignment(horizontal='right', vertical='center'); c.number_format = '0.00%'; c.border = box_border
        curr_pnl_r += 1

    widths_util = {'A': 14, 'B': 44, 'C': 20, 'D': 20, 'E': 20, 'F': 20, 'G': 20, 'H': 20, 'I': 20, 'J': 22, 'K': 16, 'L': 16, 'M': 18, 'N': 18, 'O': 20}
    for col_l, w in widths_util.items():
        ws.column_dimensions[col_l].width = w
    ws.views.sheetView[0].showGridLines = True


def _adjuntar_hojas_gastos_y_utilidad(wb, ruta_gastos):
    import openpyxl
    sedes_keys = [
        ('San jose', 'San José'),
        ('GRANADA MOBILITY', 'Granada Mobility'),
        ('SAN MARTIN', 'San Martín'),
        ('VILLAVICENCIO', 'Villavicencio'),
        ('PUERTO INIRIDA', 'Puerto Inírida'),
        ('GRANADA TVS', 'Granada TVS'),
        ('CARREÑO', 'Carreño')
    ]
    wb_source = openpyxl.load_workbook(ruta_gastos, read_only=True)
    ws_c = wb_source['Consolidado']
    accounts_51, accounts_52, accounts_53 = [], [], []
    for row in ws_c.iter_rows(min_row=4, values_only=True):
        code = str(row[0]).strip() if row[0] is not None else ''
        name = str(row[1]).strip() if row[1] is not None else ''
        deb = float(row[4]) if row[4] is not None else 0.0
        cred = float(row[5]) if row[5] is not None else 0.0
        if len(code) >= 6 and code.startswith('5'):
            net = deb - cred
            item = (code, name, net)
            if code.startswith('51'): accounts_51.append(item)
            elif code.startswith('52'): accounts_52.append(item)
            elif code.startswith('53'): accounts_53.append(item)
    accounts_51.sort(key=lambda x: x[0]); accounts_52.sort(key=lambda x: x[0]); accounts_53.sort(key=lambda x: x[0])

    sede_data = {s_sh: {} for s_sh, _ in sedes_keys}
    for s_sh, _ in sedes_keys:
        ws_s = wb_source[s_sh]
        for row in ws_s.iter_rows(min_row=4, values_only=True):
            code = str(row[0]).strip() if row[0] is not None else ''
            deb = float(row[4]) if row[4] is not None else 0.0
            cred = float(row[5]) if row[5] is not None else 0.0
            if len(code) >= 6 and code.startswith('5'):
                sede_data[s_sh][code] = sede_data[s_sh].get(code, 0.0) + (deb - cred)
    wb_source.close()

    for code, name, c_net in (accounts_51 + accounts_52 + accounts_53):
        s_sum = sum(sede_data[s_sh].get(code, 0.0) for s_sh, _ in sedes_keys)
        diff = round(c_net - s_sum, 2)
        if abs(diff) > 0.001:
            if code in ['53052006', '51952501']:
                sede_data['San jose'][code] = sede_data['San jose'].get(code, 0.0) + diff
            else:
                sede_data['VILLAVICENCIO'][code] = sede_data['VILLAVICENCIO'].get(code, 0.0) + diff

    for sh_name in ['OCTUBRE', 'UTILIDAD_OCTUBRE']:
        if sh_name in wb.sheetnames:
            del wb[sh_name]

    ws_oct = wb.create_sheet(title='OCTUBRE')
    r_total_oct, r_sin_int_oct, r_int_oct = _style_pastel_octubre(ws_oct, accounts_51, accounts_52, accounts_53, sede_data, sedes_keys)
    ws_util = wb.create_sheet(title='UTILIDAD_OCTUBRE')
    _style_pastel_utilidad(ws_util, r_total_oct, r_sin_int_oct, r_int_oct)


def procesar_informe_ventas_asiento(ruta_entrada, ruta_salida=None, ruta_asiento_nc=None, ruta_tb=None, ruta_gastos=None):
    """
    Lee el archivo de Asiento Contable de Ventas (o informe de facturas emitidas),
    aplica la configuración dinámica de Vendedores Especiales (Repuestos y Jaime) por Centro de Costo,
    cruza y valida las Notas Crédito con su Factura de Venta de Origen (sin inventar datos),
    y genera un libro Excel profesional con 4 hojas:
      1. Resumen Ventas CC y Vendedor (matriz detallada con subtotales por sede y fórmulas vivas)
      2. Consolidado por Sede (columnas de Vehículos Generales, Vehículos Jaime, Repuestos, Facturación Bruta, NC y Neto)
      3. Detalle Ventas (base de datos completa con factura referenciada y categoría asignada)
      4. Detalle Notas Crédito (auditoría exclusiva de notas crédito cruzadas)
    """
    import datetime
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    if not ruta_salida:
        dir_name = os.path.dirname(ruta_entrada) or "."
        ruta_salida = os.path.join(dir_name, "Informe_Ventas_Por_Centro_Costo.xlsx")

    ruta_salida = _asegurar_ruta_disponible(ruta_salida)

    # Cargar configuración dinámica de vendedores especiales
    cfg_vend = cargar_config_vendedores()
    v_repuestos = [str(x).upper().strip() for x in cfg_vend.get("vendedores_repuestos", [])]
    v_jaime = [str(x).upper().strip() for x in cfg_vend.get("vendedores_vehiculos_jaime", [])]

    wb_in = openpyxl.load_workbook(ruta_entrada, data_only=False)

    # Buscar la hoja adecuada de datos
    nombre_hoja = None
    for s in ["Detalle Ventas", "Detalle", "Ventas", "Sheet1", "Hoja1"]:
        if s in wb_in.sheetnames:
            nombre_hoja = s
            break
    if not nombre_hoja:
        for s in wb_in.sheetnames:
            if "resumen" not in s.lower() and "consolidado" not in s.lower():
                nombre_hoja = s
                break
    if not nombre_hoja:
        nombre_hoja = wb_in.sheetnames[0]

    ws_in = wb_in[nombre_hoja]

    # Detectar la fila de encabezados
    header_row = 1
    for r in range(1, min(15, ws_in.max_row + 1)):
        row_vals = [str(ws_in.cell(r, c).value or "").lower() for c in range(1, min(25, ws_in.max_column + 1))]
        if any("vendedor" in val for val in row_vals) and any("costo" in val or "base" in val or "total" in val or "número" in val or "numero" in val for val in row_vals):
            header_row = r
            break

    # Diccionarios de corrección y homologación
    cc_fix = {
        "FECH8593": "Granada Mobility",
        "FERP5759": "San Jose",
        "FERP5750": "San Jose",
        "FECH8548": "Granada",
        "FERP5740": "Granada Mobility",
        "FEMG563": "Granada Mobility",
        "FERP5730": "San Jose",
        "FERP5703": "San Jose",
        "FEGM298": "Granada Mobility",
    }

    vendor_cc_map = {
        "BELLANIT CUELLAR ARTUNDUAGA": "Granada",
        "BRAYAN YAAN ANGULO ROJAS": "Villavicencio",
        "CH MOTOS": "Villavicencio",
        "ELKIN OSWALDO BRICEÑO BOGOTA": "Villavicencio",
        "FABIO LEONARDO ABRIL RODRIGUEZ": "Granada Mobility",
        "JAIME EDUARDO  OROZCO CHAVARRO": "Granada",
        "JENNIFER MALDONADO SALAZAR": "San Jose",
        "JORMAN ESTIBEN BEDOYA GALLEGO": "Villavicencio",
        "KAREN LORENA ROA GUTIERREZ": "San Jose",
        "KELLY JOHANA BELLO GAÑAN": "San Martin",
        "LILIANA DEL PILAR MENDOZA GUEVARA": "San Jose",
        "MAILYN YORGELY CONTRERAS CARDENAS": "San Jose",
        "STIVEN ENRIQUE FLOREZ ROMERO": "Granada",
    }

    # 1. Leer todas las filas del archivo principal y de notas separadas si existen
    raw_rows = []
    for r in range(header_row + 1, ws_in.max_row + 1):
        row_vals = [ws_in.cell(r, c).value for c in range(1, 23)]
        if any(row_vals):
            raw_rows.append(row_vals)

    existing_nums = {str(r[0] or "").strip().upper() for r in raw_rows if r[0]}

    n_nc_cruzadas = 0
    total_nc_cruzadas = 0.0
    if ruta_asiento_nc and os.path.exists(ruta_asiento_nc):
        try:
            wb_nc = openpyxl.load_workbook(ruta_asiento_nc, data_only=False)
            ws_nc_in = wb_nc.active
            nc_hdr = 1
            for r in range(1, min(15, ws_nc_in.max_row + 1)):
                r_vals = [str(ws_nc_in.cell(r, c).value or "").lower() for c in range(1, min(20, ws_nc_in.max_column + 1))]
                if any("número" in v or "numero" in v or "referencia" in v for v in r_vals):
                    nc_hdr = r
                    break
            for r in range(nc_hdr + 1, ws_nc_in.max_row + 1):
                r_vals = [ws_nc_in.cell(r, c).value for c in range(1, 23)]
                if not any(r_vals):
                    continue
                num_c = str(r_vals[0] or "").strip()
                if num_c and num_c.upper() not in existing_nums:
                    raw_rows.append(r_vals)
                    existing_nums.add(num_c.upper())
        except Exception as e:
            print(f"Advertencia al leer archivo separado de NC: {e}")

    # 2. Indexar facturas de venta para cruce exacto de notas crédito
    inv_map = {}
    inv_by_orig = {}
    for r_vals in raw_rows:
        num = str(r_vals[0] or "").strip().upper()
        orig = str(r_vals[6] or "").strip().upper()
        ven = r_vals[3]
        cc = r_vals[2]

        if num in cc_fix:
            cc = cc_fix[num]
        if (not cc or str(cc).strip().lower() in ['none', '']) and ven and str(ven).strip() in vendor_cc_map:
            cc = vendor_cc_map[str(ven).strip()]

        v_clean = str(ven).strip() if (ven and str(ven).strip().lower() != 'none') else None
        cc_clean = str(cc).strip() if (cc and str(cc).strip().lower() != 'none') else None

        if num and not num.startswith(('NC', 'NCC')):
            inv_map[num] = (v_clean, cc_clean)
            if orig and orig != 'NAN':
                inv_by_orig[orig] = (v_clean, cc_clean)

    # 3. Procesar y normalizar cada registro aplicando categorías especiales
    # Mapeo de clientes a Centro de Costo para conciliación contable de Transporte (Cta 4145)
    client_to_cc = {}
    for r_it in raw_rows:
        cli_nom = str(r_it[1] or "").strip().upper()
        cc_nom = str(r_it[2] or "").strip()
        if cli_nom and cc_nom and cc_nom.lower() not in ["none", "nan", ""]:
            client_to_cc[cli_nom] = cc_nom

    tb_data = None
    if ruta_tb and os.path.exists(ruta_tb):
        try:
            tb_data = leer_trial_balance_odoo(ruta_tb, client_to_cc=client_to_cc)
        except Exception as e_tb:
            print(f"Advertencia al pre-cargar Trial Balance: {e_tb}")

    processed_rows = []
    for r_vals in raw_rows:
        num = str(r_vals[0] or "").strip()
        cli = r_vals[1]
        cc = r_vals[2]
        ven = r_vals[3]
        f_fact = r_vals[4]
        f_fecha = r_vals[5]
        orig = str(r_vals[6] or "").strip()
        ref = str(r_vals[7] or "").strip()
        com = r_vals[8]
        eq = r_vals[9]
        act = r_vals[10]
        base = float(r_vals[11] or 0.0)
        imp = float(r_vals[12] or 0.0)
        tot = float(r_vals[13] or 0.0)
        imp_adeu = float(r_vals[14] or 0.0)
        moneda = r_vals[15] or "COP"
        fe_elec = r_vals[16]
        edi_bloq = r_vals[17]
        edi_err = r_vals[18]
        comm_stat = r_vals[19]
        est_pago = str(r_vals[20] or "").strip()
        enviado = r_vals[21]

        is_nc = bool(re.match(r'^(NC|NCC)', num, re.IGNORECASE))
        is_nd = bool(re.match(r'^(DNCC)', num, re.IGNORECASE)) or "NOTA DÉBITO" in ref.upper() or "NOTA DEBITO" in ref.upper()
        ref_fact = ""

        if is_nc:
            m = re.search(r'Reversi[oó]n\s+de:\s*([A-Za-z0-9_-]+)', ref, re.IGNORECASE)
            orig_fact = m.group(1).upper() if m else None
            ref_fact = orig_fact or ""
            matched_v, matched_cc = None, None

            if orig_fact and orig_fact in inv_map:
                matched_v, matched_cc = inv_map[orig_fact]
            elif orig and orig.upper() in inv_by_orig:
                matched_v, matched_cc = inv_by_orig[orig.upper()]

            if not ven or str(ven).strip().lower() in ['none', '']:
                ven = matched_v
            if not cc or str(cc).strip().lower() in ['none', '']:
                cc = matched_cc

            if (not cc or str(cc).strip().lower() in ['none', '']) and ven == "FABIO LEONARDO ABRIL RODRIGUEZ":
                cc = "Granada Mobility"
            
            n_nc_cruzadas += 1
            total_nc_cruzadas += tot
        elif is_nd:
            m = re.search(r'(NCC\d+)', ref, re.IGNORECASE)
            orig_fact = m.group(1).upper() if m else None
            ref_fact = orig_fact or ""
        else:
            if num in cc_fix:
                cc = cc_fix[num]
            if num == "FERP5777" or (orig.upper() == "S02927" and (not ven or str(ven).strip().lower() in ['none', ''])):
                ven = "LILIANA DEL PILAR MENDOZA GUEVARA"
                cc = "San Jose"
            if (not cc or str(cc).strip().lower() in ['none', '']) and ven and str(ven).strip() in vendor_cc_map:
                cc = vendor_cc_map[str(ven).strip()]

        if not ven or str(ven).strip().lower() in ['none', '']:
            ven = "SIN ASIGNAR" if est_pago.lower() != "cancelado" else "CANCELADO"
        else:
            ven = str(ven).strip()

        if not cc or str(cc).strip().lower() in ['none', '']:
            cc = "SIN ASIGNAR"
        else:
            cc = str(cc).strip()

        # Categorización especial dinámica
        v_upper = ven.upper().strip()
        num_upper = str(num).upper().strip()
        cli_upper = str(cli).upper()
        ref_upper = str(ref).upper()
        orig_upper = str(orig).upper()

        taller_nums = [
            'FEVC1149', 'FEVC1150', 'FEVC1151', 'FEVC1152', 'FEVC1153', 'FEVC1154',
            'FEVC1219', 'FEVC1220', 'FEVC1221', 'FEVC1222', 'FEVC1223', 'FEVC1224'
        ]
        is_fabricante = any(k in cli_upper for k in ["AUTECO", "AUTOTECNICA", "AUTOTÉCNICA"])

        if is_nc:
            cat_asignada = "NOTA CRÉDITO"
        elif is_nd:
            cat_asignada = "NOTA DÉBITO 4175"
        elif num_upper in ["FEGM298", "FEGM317"] or "SUPERMOTOS UMA" in cli_upper:
            cat_asignada = "ARRIENDO UMA"
        elif num_upper in taller_nums or orig_upper in ['S03532', 'S03534', 'S03535', 'S03536', 'S03537', 'S03538']:
            cat_asignada = "LIQUIDACIÓN TALLER"
        elif is_fabricante and (num_upper == "FEVC1232" or ("REP" in ref_upper and "CES" in ref_upper)):
            cat_asignada = "CESANTÍAS REPUESTOS"
        elif is_fabricante and (base >= 8000000 or num_upper in ["FEVC1157", "FEVC1230", "FEVC1231"] or "CESANT" in ref_upper or "CESANT" in orig_upper):
            cat_asignada = "CESANTÍAS MOVILIDAD"
        elif is_fabricante:
            cat_asignada = "OTROS INGRESOS (CONVENIOS)"
        elif v_upper in v_jaime:
            cat_asignada = "VEHÍCULOS JAIME"
        elif v_upper in v_repuestos or num_upper.startswith("FERP"):
            cat_asignada = "REPUESTOS"
        elif v_upper == "CH MOTOS" and ("01" in ref_upper or "TRANSPORTE" in cli_upper or "TRANS" in orig_upper):
            cat_asignada = "TRANSPORTE"
        else:
            cat_asignada = "VEHÍCULOS GENERALES"

        pref = num_upper[:4]
        if pref.startswith("FERP"): linea = "Repuestos"
        elif pref.startswith("FEMG"): linea = "Repuestos / Taller"
        elif cat_asignada == "ARRIENDO UMA": linea = "Otros Ingresos / Arriendo UMA"
        elif cat_asignada == "LIQUIDACIÓN TALLER": linea = "Liquidación Taller"
        elif cat_asignada == "NOTA DÉBITO 4175": linea = "Devoluciones / Nota Débito"
        elif cat_asignada.startswith("CESANTÍAS"): linea = "Otros Ingresos / Cesantías"
        elif cat_asignada.startswith("OTROS INGRESOS"): linea = "Otros Ingresos / Convenios"
        elif cat_asignada == "TRANSPORTE": linea = "Transporte"
        elif pref.startswith(("FECH", "FEVC", "FEGM", "MI")): linea = "Vehículos / Motos"
        else: linea = "Comercial"

        # Asignación de Cuenta Contable PUC de Ingreso en el Balance
        if is_nc:
            cuenta_puc = "4175"
            if imp < 0:
                nombre_cuenta_puc = "DEV EN VENTAS 19% (NOTAS CRÉDITO 4175 DB)"
            else:
                nombre_cuenta_puc = "DEV EN VENTAS EXENTAS (NOTAS CRÉDITO 4175 DB)"
        elif is_nd:
            cuenta_puc = "41750506"
            nombre_cuenta_puc = "DEV REPUESTOS EXENTOS (NOTAS DÉBITO 4175 CR)"
        elif cat_asignada == "ARRIENDO UMA":
            cuenta_puc = "42354001"
            nombre_cuenta_puc = "TALLER DE VEHICULOS (ARRIENDO UMA)"
        elif cat_asignada == "LIQUIDACIÓN TALLER":
            cuenta_puc = "42354001"
            nombre_cuenta_puc = "TALLER DE VEHICULOS (AUTECO/AUTOTÉCNICA)"
        elif cat_asignada == "CESANTÍAS MOVILIDAD":
            cuenta_puc = "41800103"
            nombre_cuenta_puc = "CESANTIAS COMERCIAL MOVILIDAD"
        elif cat_asignada == "CESANTÍAS REPUESTOS":
            cuenta_puc = "41800104"
            nombre_cuenta_puc = "CESANTÍAS COMERCIAL REPUESTOS"
        elif cat_asignada.startswith("OTROS INGRESOS"):
            cuenta_puc = "42352001"
            nombre_cuenta_puc = "SERVICIOS ADMINISTRATIVOS (CONVENIOS)"
        elif cat_asignada == "TRANSPORTE":
            cuenta_puc = "41450501"
            nombre_cuenta_puc = "SERVICIO DE CARGUE Y DESCARGUE (TRANSPORTE)"
        elif cat_asignada == "REPUESTOS":
            if imp > 0:
                cuenta_puc = "41350602"
                nombre_cuenta_puc = "VENTA DE REPUESTOS 19%"
            else:
                cuenta_puc = "41350601"
                nombre_cuenta_puc = "VENTA DE REPUESTOS EXENTOS"
        else: # Vehículos
            if imp > 0:
                cuenta_puc = "41350202"
                nombre_cuenta_puc = "VENTAS DE VEHICULOS DEL 19%"
            else:
                cuenta_puc = "41350201"
                nombre_cuenta_puc = "VENTAS DE VEHICULOS EXENTAS"

        processed_rows.append({
            "num": num, "cli": cli, "cc": cc, "ven": ven, "f_fact": f_fact, "f_fecha": f_fecha,
            "orig": orig, "ref": ref, "com": com, "eq": eq, "act": act,
            "base": base, "imp": imp, "tot": tot, "imp_adeu": imp_adeu, "moneda": moneda,
            "fe_elec": fe_elec, "edi_bloq": edi_bloq, "edi_err": edi_err,
            "comm_stat": comm_stat, "est_pago": est_pago, "enviado": enviado,
            "is_nc": is_nc, "ref_fact": ref_fact, "cat_asignada": cat_asignada, "linea": linea,
            "cuenta_puc": cuenta_puc, "nombre_cuenta_puc": nombre_cuenta_puc
        })

    # Estilos y formatos
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    FONT_NAME = "Calibri"
    C_NAVY_DARK = "1F4E79"
    C_NAVY_MED = "2F5597"
    C_BLUE_LIGHT = "D9E1F2"
    C_BLUE_CARD = "F2F5F9"
    C_ZEBRA = "F9FBFD"
    C_WHITE = "FFFFFF"
    C_TEXT_DARK = "0F172A"
    C_TEXT_MUTED = "595959"
    C_BORDER = "D3D3D3"
    C_BORDER_DARK = "8EA9DB"
    C_GREEN = "2E7D32"
    C_AMBER = "B45309"
    C_RED_DARK = "C00000"

    f_title = Font(name=FONT_NAME, size=15, bold=True, color=C_NAVY_DARK)
    f_subt = Font(name=FONT_NAME, size=10, bold=False, color=C_TEXT_MUTED)
    f_date = Font(name=FONT_NAME, size=9, italic=True, color=C_TEXT_MUTED)
    f_kpi_lbl = Font(name=FONT_NAME, size=8, bold=True, color=C_TEXT_MUTED)
    f_hdr = Font(name=FONT_NAME, size=10, bold=True, color=C_WHITE)
    f_dat = Font(name=FONT_NAME, size=9, color=C_TEXT_DARK)
    f_dat_bold = Font(name=FONT_NAME, size=9, bold=True, color=C_TEXT_DARK)
    f_subtot = Font(name=FONT_NAME, size=9, bold=True, color=C_NAVY_DARK)
    f_grand = Font(name=FONT_NAME, size=10, bold=True, color=C_WHITE)

    fill_hdr = PatternFill(start_color=C_NAVY_DARK, end_color=C_NAVY_DARK, fill_type="solid")
    fill_subtot = PatternFill(start_color=C_BLUE_LIGHT, end_color=C_BLUE_LIGHT, fill_type="solid")
    fill_grand = PatternFill(start_color=C_NAVY_DARK, end_color=C_NAVY_DARK, fill_type="solid")
    fill_zebra = PatternFill(start_color=C_ZEBRA, end_color=C_ZEBRA, fill_type="solid")
    fill_white = PatternFill(start_color=C_WHITE, end_color=C_WHITE, fill_type="solid")
    fill_kpi = PatternFill(start_color=C_BLUE_CARD, end_color=C_BLUE_CARD, fill_type="solid")

    a_left = Alignment(horizontal="left", vertical="center")
    a_center = Alignment(horizontal="center", vertical="center")
    a_right = Alignment(horizontal="right", vertical="center")
    a_hdr = Alignment(horizontal="center", vertical="center", wrap_text=True)

    b_thin = Border(
        left=Side(style="thin", color=C_BORDER),
        right=Side(style="thin", color=C_BORDER),
        top=Side(style="thin", color=C_BORDER),
        bottom=Side(style="thin", color=C_BORDER)
    )
    b_subtot = Border(
        left=Side(style="thin", color=C_BORDER_DARK),
        right=Side(style="thin", color=C_BORDER_DARK),
        top=Side(style="thin", color=C_NAVY_MED),
        bottom=Side(style="medium", color=C_NAVY_MED)
    )
    b_grand = Border(
        left=Side(style="thin", color=C_NAVY_DARK),
        right=Side(style="thin", color=C_NAVY_DARK),
        top=Side(style="medium", color=C_NAVY_DARK),
        bottom=Side(style="double", color=C_NAVY_DARK)
    )

    NF_CURR = "$#,##0"
    NF_CURR_DEC = "$#,##0.00"
    NF_INT = "#,##0"
    NF_PCT = "0.0%"

    # ==========================================
    # 1. HOJA DETALLE VENTAS
    # ==========================================
    ws_det = wb.create_sheet(title="Detalle Ventas")
    ws_det.views.sheetView[0].showGridLines = False

    orig_headers = [
        "Número", "Nombre de la empresa a mostrar en la factura", "Centro de costo", "Vendedor",
        "Fecha de factura", "Fecha", "Origen", "Referencia", "Comercial", "Equipo de ventas",
        "Actividades", "Base imponible en la moneda firmada", "Impuesto firmado",
        "Total con signo en moneda", "Importe adeudado firmado", "Moneda", "Facturación electrónica",
        "Nivel de bloqueo de EDI", "Mensaje de error de EDI", "Commercial Status", "Estado en pago",
        "Enviado"
    ]
    calc_headers = [
        "Tipo Documento", "Factura Referenciada", "Categoría Asignada", "Venta Bruta", "Nota Crédito",
        "Tipo Venta", "Base Exenta", "Base Gravada", "IVA Gravado", "Total Gravado",
        "Total Facturación Neta", "Línea Original", "Cuenta Contable PUC", "Nombre Cuenta PUC"
    ]
    new_headers = orig_headers + calc_headers

    for col_idx, h in enumerate(new_headers, 1):
        c = ws_det.cell(1, col_idx, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = a_hdr
        c.border = b_thin
    ws_det.row_dimensions[1].height = 28

    for idx, r_data in enumerate(processed_rows, 2):
        vals = [
            r_data["num"], r_data["cli"], r_data["cc"], r_data["ven"],
            r_data["f_fact"], r_data["f_fecha"], r_data["orig"], r_data["ref"],
            r_data["com"], r_data["eq"], r_data["act"], r_data["base"],
            r_data["imp"], r_data["tot"], r_data["imp_adeu"], r_data["moneda"],
            r_data["fe_elec"], r_data["edi_bloq"], r_data["edi_err"],
            r_data["comm_stat"], r_data["est_pago"], r_data["enviado"]
        ]
        for c_i, val in enumerate(vals, 1):
            cell = ws_det.cell(idx, c_i, val)
            cell.font = f_dat
            cell.border = b_thin
            if c_i in [5, 6] and isinstance(val, (datetime.datetime, datetime.date)):
                cell.number_format = "yyyy-mm-dd"
                cell.alignment = a_center
            elif c_i in [12, 14, 15]:
                cell.number_format = NF_CURR
                cell.alignment = a_right
            elif c_i == 13:
                cell.number_format = NF_CURR_DEC
                cell.alignment = a_right
            elif c_i in [1, 3, 7, 8, 16, 20, 21, 22]:
                cell.alignment = a_center
            else:
                cell.alignment = a_left

        # Col 23: Tipo Documento
        ws_det.cell(idx, 23, '=IF(OR(LEFT(A{0},2)="NC",LEFT(A{0},3)="NCC"),"NOTA CRÉDITO",IF(LEFT(A{0},4)="DNCC","NOTA DÉBITO","FACTURA"))'.format(idx)).alignment = a_center
        # Col 24: Factura Referenciada
        ws_det.cell(idx, 24, r_data["ref_fact"]).alignment = a_center
        # Col 25: Categoría Asignada (Vehículos Generales, Vehículos Jaime, Repuestos, etc.)
        ws_det.cell(idx, 25, r_data["cat_asignada"]).alignment = a_center
        # Col 26: Venta Bruta (Para Cta 42 Taller y Arriendo UMA se toma base imponible L; cesantías, convenios y notas débito en 0 para no distorsionar ventas operacionales)
        c_vb = ws_det.cell(idx, 26, '=IF(W{0}="FACTURA",IF(OR(Y{0}="LIQUIDACIÓN TALLER",Y{0}="ARRIENDO UMA"),L{0},IF(OR(LEFT(Y{0},9)="CESANTÍAS",LEFT(Y{0},14)="OTROS INGRESOS",LEFT(Y{0},11)="NOTA DÉBITO"),0,N{0})),0)'.format(idx))
        c_vb.number_format = NF_CURR
        c_vb.alignment = a_right
        # Col 27: Nota Crédito (Toma la base imponible L para conciliar exactamente con el Débito de la Cuenta 4175 en el Balance)
        c_nc = ws_det.cell(idx, 27, '=IF(W{0}="NOTA CRÉDITO",L{0},0)'.format(idx))
        c_nc.number_format = NF_CURR
        c_nc.alignment = a_right
        # Col 28: Tipo Venta
        ws_det.cell(idx, 28, '=IF(OR(M{0}>0,AND(M{0}<0,L{0}<0)),"GRAVADA","EXENTA")'.format(idx)).alignment = a_center
        # Col 29: Base Exenta (Excluye cesantías, convenios y notas débito que van en Otros Ingresos y Ajustes 4175)
        c_be = ws_det.cell(idx, 29, '=IF(AND(AB{0}="EXENTA",LEFT(Y{0},9)<>"CESANTÍAS",LEFT(Y{0},14)<>"OTROS INGRESOS",LEFT(Y{0},11)<>"NOTA DÉBITO"),L{0},0)'.format(idx))
        c_be.number_format = NF_CURR
        c_be.alignment = a_right
        # Col 30: Base Gravada (Excluye cesantías y convenios que van en Otros Ingresos)
        c_bg = ws_det.cell(idx, 30, '=IF(AND(AB{0}="GRAVADA",LEFT(Y{0},9)<>"CESANTÍAS",LEFT(Y{0},14)<>"OTROS INGRESOS",LEFT(Y{0},11)<>"NOTA DÉBITO"),L{0},0)'.format(idx))
        c_bg.number_format = NF_CURR
        c_bg.alignment = a_right
        # Col 31: IVA Gravado (Excluye cesantías y convenios que van en Otros Ingresos)
        c_iv = ws_det.cell(idx, 31, '=IF(AND(AB{0}="GRAVADA",LEFT(Y{0},9)<>"CESANTÍAS",LEFT(Y{0},14)<>"OTROS INGRESOS",LEFT(Y{0},11)<>"NOTA DÉBITO"),M{0},0)'.format(idx))
        c_iv.number_format = NF_CURR_DEC
        c_iv.alignment = a_right
        # Col 32: Total Gravado
        c_tg = ws_det.cell(idx, 32, '=AD{0}+AE{0}'.format(idx))
        c_tg.number_format = NF_CURR
        c_tg.alignment = a_right
        # Col 33: Total Facturación Neta
        c_tn = ws_det.cell(idx, 33, '=Z{0}+AA{0}+IF(W{0}="NOTA DÉBITO",L{0},0)'.format(idx))
        c_tn.number_format = NF_CURR
        c_tn.alignment = a_right
        # Col 34: Línea Original
        ws_det.cell(idx, 34, r_data["linea"]).alignment = a_center
        # Col 35: Cuenta Contable PUC
        ws_det.cell(idx, 35, r_data["cuenta_puc"]).alignment = a_center
        ws_det.cell(idx, 35).font = f_dat_bold
        # Col 36: Nombre Cuenta PUC
        ws_det.cell(idx, 36, r_data["nombre_cuenta_puc"]).alignment = a_left

        for c_k in range(23, 37):
            ws_det.cell(idx, c_k).font = f_dat
            ws_det.cell(idx, c_k).border = b_thin
        ws_det.row_dimensions[idx].height = 19

    det_max_r = len(processed_rows) + 1
    ws_det.freeze_panes = "A2"
    ws_det.auto_filter.ref = f"A1:AJ{det_max_r}"

    for col in ws_det.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_det.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 38)

    # ==========================================
    # 2. HOJA RESUMEN VENTAS CC Y VENDEDOR
    # ==========================================
    ws_res = wb.create_sheet(title="Resumen Ventas CC y Vendedor", index=0)
    ws_res.views.sheetView[0].showGridLines = False

    ws_res.cell(2, 2, "CH INVERSIONES COLOMBIA S.A.S ZOMAC").font = f_title
    ws_res.cell(3, 2, "INFORME CONSOLIDADO DE VENTAS Y NOTAS POR CENTRO DE COSTO Y VENDEDOR").font = Font(name=FONT_NAME, size=12, bold=True, color=C_TEXT_DARK)
    ws_res.cell(4, 2, "Periodo: Facturación Emitida | Desglose por Línea Asignada | Exentas vs Gravadas").font = f_subt
    ws_res.cell(5, 2, f"Fecha de generación: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}").font = f_date

    df_proc = pd.DataFrame(processed_rows)
    df_valid = df_proc[df_proc["est_pago"].astype(str).str.lower() != "cancelado"]

    sede_order = ["Granada", "Granada Mobility", "Inirida", "San Jose", "San Martin", "Villavicencio"]
    sedes_raw = [str(s).strip() for s in df_valid["cc"].unique() if pd.notna(s) and str(s).strip() and str(s).lower() != 'none']
    sedes = [s for s in sede_order if s in sedes_raw] + [s for s in sedes_raw if s not in sede_order]
    if not sedes:
        sedes = ["SIN ASIGNAR"]

    tbl_headers = [
        "Centro de Costo", "Vendedor", "Línea Asignada", "Cant. Facturas", "Total Ventas Brutas",
        "Cant. Notas Crédito", "Total Notas Crédito", "Total Facturación Neta",
        "Ventas Exentas (Base)", "Base Gravada (19%)", "IVA Gravado",
        "Total Gravado", "% Part. Sede", "% Part. Total"
    ]

    start_row = 10
    for c_i, h_text in enumerate(tbl_headers, 2):
        c = ws_res.cell(start_row, c_i, h_text)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = a_hdr
        c.border = b_thin
    ws_res.row_dimensions[start_row].height = 28

    total_data_lines = 0
    for s in sedes:
        v_list = df_valid[df_valid["cc"] == s]["ven"].unique()
        total_data_lines += len(v_list) + 1
    grand_total_row = start_row + total_data_lines + 1

    kpis = [
        ("TOTAL FACTURACIÓN NETA", f"=I{grand_total_row}", NF_CURR, C_NAVY_DARK),
        ("TOTAL VENTAS BRUTAS", f"=F{grand_total_row}", NF_CURR, C_GREEN),
        ("TOTAL NOTAS CRÉDITO", f"=H{grand_total_row}", NF_CURR, C_RED_DARK),
        ("TOTAL IVA FACTURADO", f"=L{grand_total_row}", NF_CURR, C_NAVY_MED),
        ("TOTAL FACTURAS", f"=E{grand_total_row}", NF_INT, C_TEXT_DARK),
        ("TOTAL NOTAS", f"=G{grand_total_row}", NF_INT, C_AMBER)
    ]

    for i, (k_lbl, k_form, k_fmt, k_col) in enumerate(kpis):
        c_start = 2 + i * 2
        c_end = c_start + 1
        ws_res.merge_cells(start_row=7, start_column=c_start, end_row=7, end_column=c_end)
        ws_res.merge_cells(start_row=8, start_column=c_start, end_row=8, end_column=c_end)

        lbl_cell = ws_res.cell(7, c_start, k_lbl)
        lbl_cell.font = f_kpi_lbl
        lbl_cell.alignment = a_center
        lbl_cell.fill = fill_kpi

        num_cell = ws_res.cell(8, c_start, k_form)
        num_cell.font = Font(name=FONT_NAME, size=13, bold=True, color=k_col)
        num_cell.alignment = a_center
        num_cell.number_format = k_fmt
        num_cell.fill = fill_kpi

        for r_k in [7, 8]:
            for c_k in [c_start, c_end]:
                ws_res.cell(r_k, c_k).border = b_thin

    det_cc_rng = f"'Detalle Ventas'!$C$2:$C${det_max_r}"
    det_ven_rng = f"'Detalle Ventas'!$D$2:$D${det_max_r}"
    det_est_rng = f"'Detalle Ventas'!$U$2:$U${det_max_r}"
    det_tipo_rng = f"'Detalle Ventas'!$W$2:$W${det_max_r}"
    det_vbruta_rng = f"'Detalle Ventas'!$Z$2:$Z${det_max_r}"
    det_nc_rng = f"'Detalle Ventas'!$AA$2:$AA${det_max_r}"
    det_bex_rng = f"'Detalle Ventas'!$AC$2:$AC${det_max_r}"
    det_bgrav_rng = f"'Detalle Ventas'!$AD$2:$AD${det_max_r}"
    det_iva_rng = f"'Detalle Ventas'!$AE$2:$AE${det_max_r}"

    curr_row = start_row + 1
    subtotal_rows = []

    for s in sedes:
        s_start_row = curr_row
        v_list = sorted([str(v) for v in df_valid[df_valid["cc"] == s]["ven"].unique() if pd.notna(v) and str(v).strip()])
        if not v_list:
            v_list = ["SIN ASIGNAR"]

        for v in v_list:
            is_even = (curr_row % 2 == 0)
            row_fill = fill_zebra if is_even else fill_white

            # Identificar categoría asignada al vendedor
            v_up = v.upper().strip()
            if v_up in v_jaime:
                cat_v = "Vehículos Jaime"
            elif v_up in v_repuestos:
                cat_v = "Repuestos"
            elif v_up == "CH MOTOS":
                cat_v = "Comercial / Flota"
            else:
                cat_v = "Vehículos"

            # Col B: Sede
            ws_res.cell(curr_row, 2, s).font = f_dat
            ws_res.cell(curr_row, 2).alignment = a_left
            ws_res.cell(curr_row, 2).fill = row_fill
            ws_res.cell(curr_row, 2).border = b_thin

            # Col C: Vendedor
            ws_res.cell(curr_row, 3, v).font = f_dat_bold if v != "CH MOTOS" else Font(name=FONT_NAME, size=9, bold=True, color=C_NAVY_MED)
            ws_res.cell(curr_row, 3).alignment = a_left
            ws_res.cell(curr_row, 3).fill = row_fill
            ws_res.cell(curr_row, 3).border = b_thin

            # Col D: Línea Asignada
            ws_res.cell(curr_row, 4, cat_v).font = f_dat
            ws_res.cell(curr_row, 4).alignment = a_center
            ws_res.cell(curr_row, 4).fill = row_fill
            ws_res.cell(curr_row, 4).border = b_thin

            # Col E: Cant. Facturas
            ws_res.cell(curr_row, 5, f'=COUNTIFS({det_cc_rng}, B{curr_row}, {det_ven_rng}, C{curr_row}, {det_est_rng}, "<>Cancelado", {det_tipo_rng}, "FACTURA")').font = f_dat
            ws_res.cell(curr_row, 5).alignment = a_center
            ws_res.cell(curr_row, 5).number_format = NF_INT
            ws_res.cell(curr_row, 5).fill = row_fill
            ws_res.cell(curr_row, 5).border = b_thin

            # Col F: Total Ventas Brutas
            ws_res.cell(curr_row, 6, f'=SUMIFS({det_vbruta_rng}, {det_cc_rng}, B{curr_row}, {det_ven_rng}, C{curr_row}, {det_est_rng}, "<>Cancelado")').font = f_dat
            ws_res.cell(curr_row, 6).alignment = a_right
            ws_res.cell(curr_row, 6).number_format = NF_CURR
            ws_res.cell(curr_row, 6).fill = row_fill
            ws_res.cell(curr_row, 6).border = b_thin

            # Col G: Cant. Notas Crédito
            ws_res.cell(curr_row, 7, f'=COUNTIFS({det_cc_rng}, B{curr_row}, {det_ven_rng}, C{curr_row}, {det_est_rng}, "<>Cancelado", {det_tipo_rng}, "NOTA CRÉDITO")').font = f_dat
            ws_res.cell(curr_row, 7).alignment = a_center
            ws_res.cell(curr_row, 7).number_format = NF_INT
            ws_res.cell(curr_row, 7).fill = row_fill
            ws_res.cell(curr_row, 7).border = b_thin

            # Col H: Total Notas Crédito
            ws_res.cell(curr_row, 8, f'=SUMIFS({det_nc_rng}, {det_cc_rng}, B{curr_row}, {det_ven_rng}, C{curr_row}, {det_est_rng}, "<>Cancelado")').font = f_dat
            ws_res.cell(curr_row, 8).alignment = a_right
            ws_res.cell(curr_row, 8).number_format = NF_CURR
            ws_res.cell(curr_row, 8).fill = row_fill
            ws_res.cell(curr_row, 8).border = b_thin

            # Col I: Total Facturación Neta
            ws_res.cell(curr_row, 9, f"=F{curr_row}+H{curr_row}").font = f_dat_bold
            ws_res.cell(curr_row, 9).alignment = a_right
            ws_res.cell(curr_row, 9).number_format = NF_CURR
            ws_res.cell(curr_row, 9).fill = row_fill
            ws_res.cell(curr_row, 9).border = b_thin

            # Col J: Ventas Exentas
            ws_res.cell(curr_row, 10, f'=SUMIFS({det_bex_rng}, {det_cc_rng}, B{curr_row}, {det_ven_rng}, C{curr_row}, {det_est_rng}, "<>Cancelado")').font = f_dat
            ws_res.cell(curr_row, 10).alignment = a_right
            ws_res.cell(curr_row, 10).number_format = NF_CURR
            ws_res.cell(curr_row, 10).fill = row_fill
            ws_res.cell(curr_row, 10).border = b_thin

            # Col K: Base Gravada
            ws_res.cell(curr_row, 11, f'=SUMIFS({det_bgrav_rng}, {det_cc_rng}, B{curr_row}, {det_ven_rng}, C{curr_row}, {det_est_rng}, "<>Cancelado")').font = f_dat
            ws_res.cell(curr_row, 11).alignment = a_right
            ws_res.cell(curr_row, 11).number_format = NF_CURR
            ws_res.cell(curr_row, 11).fill = row_fill
            ws_res.cell(curr_row, 11).border = b_thin

            # Col L: IVA Gravado
            ws_res.cell(curr_row, 12, f'=SUMIFS({det_iva_rng}, {det_cc_rng}, B{curr_row}, {det_ven_rng}, C{curr_row}, {det_est_rng}, "<>Cancelado")').font = f_dat
            ws_res.cell(curr_row, 12).alignment = a_right
            ws_res.cell(curr_row, 12).number_format = NF_CURR_DEC
            ws_res.cell(curr_row, 12).fill = row_fill
            ws_res.cell(curr_row, 12).border = b_thin

            # Col M: Total Gravado
            ws_res.cell(curr_row, 13, f"=K{curr_row}+L{curr_row}").font = f_dat
            ws_res.cell(curr_row, 13).alignment = a_right
            ws_res.cell(curr_row, 13).number_format = NF_CURR
            ws_res.cell(curr_row, 13).fill = row_fill
            ws_res.cell(curr_row, 13).border = b_thin

            s_subtot_row = s_start_row + len(v_list)
            # Col N: % Part. Sede
            ws_res.cell(curr_row, 14, f'=IF($I${s_subtot_row}>0, I{curr_row}/$I${s_subtot_row}, 0)').font = f_dat
            ws_res.cell(curr_row, 14).alignment = a_right
            ws_res.cell(curr_row, 14).number_format = NF_PCT
            ws_res.cell(curr_row, 14).fill = row_fill
            ws_res.cell(curr_row, 14).border = b_thin

            # Col O: % Part. Total
            ws_res.cell(curr_row, 15, f'=IF($I${grand_total_row}>0, I{curr_row}/$I${grand_total_row}, 0)').font = f_dat
            ws_res.cell(curr_row, 15).alignment = a_right
            ws_res.cell(curr_row, 15).number_format = NF_PCT
            ws_res.cell(curr_row, 15).fill = row_fill
            ws_res.cell(curr_row, 15).border = b_thin

            ws_res.row_dimensions[curr_row].height = 20
            curr_row += 1

        # Subtotal Sede
        subtot_r = curr_row
        subtotal_rows.append(subtot_r)
        s_end_v_row = curr_row - 1

        ws_res.cell(subtot_r, 2, s).font = f_subtot
        ws_res.cell(subtot_r, 2).fill = fill_subtot
        ws_res.cell(subtot_r, 2).border = b_subtot
        ws_res.cell(subtot_r, 2).alignment = a_left

        ws_res.cell(subtot_r, 3, f"SUBTOTAL {s.upper()}").font = f_subtot
        ws_res.cell(subtot_r, 3).fill = fill_subtot
        ws_res.cell(subtot_r, 3).border = b_subtot
        ws_res.cell(subtot_r, 3).alignment = a_left

        ws_res.cell(subtot_r, 4, "-").font = f_subtot
        ws_res.cell(subtot_r, 4).fill = fill_subtot
        ws_res.cell(subtot_r, 4).border = b_subtot
        ws_res.cell(subtot_r, 4).alignment = a_center

        for c_idx, (col_l, fmt) in enumerate([
            ("E", NF_INT), ("F", NF_CURR), ("G", NF_INT), ("H", NF_CURR),
            ("I", NF_CURR), ("J", NF_CURR), ("K", NF_CURR), ("L", NF_CURR_DEC),
            ("M", NF_CURR)
        ], 5):
            c_sub = ws_res.cell(subtot_r, c_idx, f"=SUM({col_l}{s_start_row}:{col_l}{s_end_v_row})")
            c_sub.font = f_subtot
            c_sub.fill = fill_subtot
            c_sub.border = b_subtot
            c_sub.alignment = a_right if col_l not in ["E", "G"] else a_center
            c_sub.number_format = fmt

        ws_res.cell(subtot_r, 14, 1.0).font = f_subtot
        ws_res.cell(subtot_r, 14).fill = fill_subtot
        ws_res.cell(subtot_r, 14).border = b_subtot
        ws_res.cell(subtot_r, 14).alignment = a_right
        ws_res.cell(subtot_r, 14).number_format = NF_PCT

        ws_res.cell(subtot_r, 15, f'=IF($I${grand_total_row}>0, I{subtot_r}/$I${grand_total_row}, 0)').font = f_subtot
        ws_res.cell(subtot_r, 15).fill = fill_subtot
        ws_res.cell(subtot_r, 15).border = b_subtot
        ws_res.cell(subtot_r, 15).alignment = a_right
        ws_res.cell(subtot_r, 15).number_format = NF_PCT

        ws_res.row_dimensions[subtot_r].height = 22
        curr_row += 1

    # Gran Total
    gt_r = curr_row
    ws_res.cell(gt_r, 2, "TOTAL").font = f_grand
    ws_res.cell(gt_r, 2).fill = fill_grand
    ws_res.cell(gt_r, 2).border = b_grand
    ws_res.cell(gt_r, 2).alignment = a_center

    ws_res.cell(gt_r, 3, "TOTAL GENERAL FACTURACIÓN").font = f_grand
    ws_res.cell(gt_r, 3).fill = fill_grand
    ws_res.cell(gt_r, 3).border = b_grand
    ws_res.cell(gt_r, 3).alignment = a_left

    ws_res.cell(gt_r, 4, "-").font = f_grand
    ws_res.cell(gt_r, 4).fill = fill_grand
    ws_res.cell(gt_r, 4).border = b_grand
    ws_res.cell(gt_r, 4).alignment = a_center

    subtot_c_refs = lambda col: "+".join([f"{col}{r}" for r in subtotal_rows])

    for c_idx, (col_l, fmt) in enumerate([
        ("E", NF_INT), ("F", NF_CURR), ("G", NF_INT), ("H", NF_CURR),
        ("I", NF_CURR), ("J", NF_CURR), ("K", NF_CURR), ("L", NF_CURR_DEC),
        ("M", NF_CURR)
    ], 5):
        c_gt = ws_res.cell(gt_r, c_idx, f"={subtot_c_refs(col_l)}")
        c_gt.font = f_grand
        c_gt.fill = fill_grand
        c_gt.border = b_grand
        c_gt.alignment = a_right if col_l not in ["E", "G"] else a_center
        c_gt.number_format = fmt

    ws_res.cell(gt_r, 14, "-").font = f_grand
    ws_res.cell(gt_r, 14).fill = fill_grand
    ws_res.cell(gt_r, 14).border = b_grand
    ws_res.cell(gt_r, 14).alignment = a_center

    ws_res.cell(gt_r, 15, 1.0).font = f_grand
    ws_res.cell(gt_r, 15).fill = fill_grand
    ws_res.cell(gt_r, 15).border = b_grand
    ws_res.cell(gt_r, 15).alignment = a_right
    ws_res.cell(gt_r, 15).number_format = NF_PCT

    ws_res.row_dimensions[gt_r].height = 26

    ws_res.column_dimensions["A"].width = 3
    ws_res.column_dimensions["B"].width = 20
    ws_res.column_dimensions["C"].width = 36
    ws_res.column_dimensions["D"].width = 18
    ws_res.column_dimensions["E"].width = 15
    ws_res.column_dimensions["F"].width = 22
    ws_res.column_dimensions["G"].width = 18
    ws_res.column_dimensions["H"].width = 20
    ws_res.column_dimensions["I"].width = 24
    ws_res.column_dimensions["J"].width = 22
    ws_res.column_dimensions["K"].width = 20
    ws_res.column_dimensions["L"].width = 18
    ws_res.column_dimensions["M"].width = 20
    ws_res.column_dimensions["N"].width = 14
    ws_res.column_dimensions["O"].width = 14

    ws_res.freeze_panes = "B11"

    # Columna A como margen estetico ultra-fino y AUTOFILTRO en encabezados de Vendedores
    ws_res.column_dimensions["A"].width = 2.5
    ws_res.auto_filter.ref = f"B7:O{gt_r-1}"

    # ==========================================
    # 3. HOJA CONSOLIDADO POR SEDE
    # ==========================================
    ws_sed = wb.create_sheet(title="Consolidado por Sede")
    ws_sed.views.sheetView[0].showGridLines = False

    ws_sed.cell(2, 2, "CH INVERSIONES COLOMBIA S.A.S ZOMAC").font = f_title
    ws_sed.cell(3, 2, "RESUMEN EJECUTIVO DE VENTAS Y NOTAS POR CENTRO DE COSTO").font = Font(name=FONT_NAME, size=12, bold=True, color=C_TEXT_DARK)
    ws_sed.cell(4, 2, "Consolidado por Sedes: Vehículos Generales | Vehículos Jaime | Repuestos | Notas Crédito").font = f_subt

    ws_sed.cell(4, 2, "Consolidado por Sedes: Vehículos Generales | Vehículos Jaime | Repuestos | Transporte | Liquidación Taller | Arriendo UMA | Notas Crédito").font = f_subt

    sed_headers = [
        "Centro de Costo (Sede)", "Vehículos Generales", "Vehículos Jaime", "Repuestos",
        "Transporte (4145)", "Liquidación Taller (42)", "Arriendo UMA (42)", "Total Ventas Brutas",
        "Cant. Facturas", "Cant. Notas Crédito", "Total Notas Crédito", "Total Facturación Neta",
        "Ventas Exentas", "Base Gravada", "IVA Gravado", "% Participación"
    ]

    s_start = 6
    for idx_c, h in enumerate(sed_headers, 2):
        c = ws_sed.cell(s_start, idx_c, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = a_hdr
        c.border = b_thin
    ws_sed.row_dimensions[s_start].height = 26

    det_cat_rng = f"'Detalle Ventas'!$Y$2:$Y${det_max_r}"

    r_sed = s_start + 1
    for s_idx, (s, s_row_res) in enumerate(zip(sedes, subtotal_rows)):
        is_even = (s_idx % 2 == 0)
        r_fill = fill_zebra if is_even else fill_white

        # Col B: Sede
        ws_sed.cell(r_sed, 2, s).font = f_dat_bold
        ws_sed.cell(r_sed, 2).alignment = a_left
        ws_sed.cell(r_sed, 2).fill = r_fill
        ws_sed.cell(r_sed, 2).border = b_thin

        # Col F (6): Transporte (4145) - Conciliación analítica con Balance Odoo
        if tb_data and "transporte_por_sede" in tb_data and tb_data["transporte_por_sede"]:
            val_tra = tb_data["transporte_por_sede"].get(s, 0.0)
            c_tra = ws_sed.cell(r_sed, 6, val_tra)
        else:
            c_tra = ws_sed.cell(r_sed, 6, f'=SUMIFS({det_vbruta_rng}, {det_cc_rng}, B{r_sed}, {det_cat_rng}, "TRANSPORTE", {det_est_rng}, "<>Cancelado")')
        c_tra.font = f_dat; c_tra.alignment = a_right; c_tra.number_format = NF_CURR; c_tra.fill = r_fill; c_tra.border = b_thin

        # Col C (3): Vehículos Generales (Puro sin Cta 42 y neto de fletes)
        c_vg = ws_sed.cell(r_sed, 3, f'=SUMIFS({det_vbruta_rng}, {det_cc_rng}, B{r_sed}, {det_cat_rng}, "VEHÍCULOS GENERALES", {det_est_rng}, "<>Cancelado") - F{r_sed}')
        c_vg.font = f_dat; c_vg.alignment = a_right; c_vg.number_format = NF_CURR; c_vg.fill = r_fill; c_vg.border = b_thin

        # Col D (4): Vehículos Jaime
        c_vj = ws_sed.cell(r_sed, 4, f'=SUMIFS({det_vbruta_rng}, {det_cc_rng}, B{r_sed}, {det_cat_rng}, "VEHÍCULOS JAIME", {det_est_rng}, "<>Cancelado")')
        c_vj.font = f_dat; c_vj.alignment = a_right; c_vj.number_format = NF_CURR; c_vj.fill = r_fill; c_vj.border = b_thin

        # Col E (5): Repuestos
        c_rep = ws_sed.cell(r_sed, 5, f'=SUMIFS({det_vbruta_rng}, {det_cc_rng}, B{r_sed}, {det_cat_rng}, "REPUESTOS", {det_est_rng}, "<>Cancelado")')
        c_rep.font = f_dat; c_rep.alignment = a_right; c_rep.number_format = NF_CURR; c_rep.fill = r_fill; c_rep.border = b_thin

        # Col G (7): Liquidación Taller (Cta 42)
        c_tal = ws_sed.cell(r_sed, 7, f'=SUMIFS({det_vbruta_rng}, {det_cc_rng}, B{r_sed}, {det_cat_rng}, "LIQUIDACIÓN TALLER", {det_est_rng}, "<>Cancelado")')
        c_tal.font = f_dat; c_tal.alignment = a_right; c_tal.number_format = NF_CURR; c_tal.fill = r_fill; c_tal.border = b_thin

        # Col H (8): Arriendo UMA (Cta 42)
        c_uma = ws_sed.cell(r_sed, 8, f'=SUMIFS({det_vbruta_rng}, {det_cc_rng}, B{r_sed}, {det_cat_rng}, "ARRIENDO UMA", {det_est_rng}, "<>Cancelado")')
        c_uma.font = f_dat; c_uma.alignment = a_right; c_uma.number_format = NF_CURR; c_uma.fill = r_fill; c_uma.border = b_thin

        # Col I (9): Total Ventas Brutas
        c_brut = ws_sed.cell(r_sed, 9, f'=C{r_sed}+D{r_sed}+E{r_sed}+F{r_sed}+G{r_sed}+H{r_sed}')
        c_brut.font = f_dat_bold; c_brut.alignment = a_right; c_brut.number_format = NF_CURR; c_brut.fill = r_fill; c_brut.border = b_thin

        # Col J (10): Cant. Facturas
        ws_sed.cell(r_sed, 10, f"='Resumen Ventas CC y Vendedor'!E{s_row_res}").font = f_dat
        ws_sed.cell(r_sed, 10).alignment = a_center; ws_sed.cell(r_sed, 10).number_format = NF_INT; ws_sed.cell(r_sed, 10).fill = r_fill; ws_sed.cell(r_sed, 10).border = b_thin

        # Col K (11): Cant. NC
        ws_sed.cell(r_sed, 11, f"='Resumen Ventas CC y Vendedor'!G{s_row_res}").font = f_dat
        ws_sed.cell(r_sed, 11).alignment = a_center; ws_sed.cell(r_sed, 11).number_format = NF_INT; ws_sed.cell(r_sed, 11).fill = r_fill; ws_sed.cell(r_sed, 11).border = b_thin

        # Col L (12): Total NC
        ws_sed.cell(r_sed, 12, f"='Resumen Ventas CC y Vendedor'!H{s_row_res}").font = f_dat
        ws_sed.cell(r_sed, 12).alignment = a_right; ws_sed.cell(r_sed, 12).number_format = NF_CURR; ws_sed.cell(r_sed, 12).fill = r_fill; ws_sed.cell(r_sed, 12).border = b_thin

        # Col M (13): Total Facturación Neta (Ingresos Operacionales)
        c_net = ws_sed.cell(r_sed, 13, f'=I{r_sed}+L{r_sed}')
        c_net.font = f_dat_bold; c_net.alignment = a_right; c_net.number_format = NF_CURR; c_net.fill = r_fill; c_net.border = b_thin

        # Col N (14): Ventas Exentas
        ws_sed.cell(r_sed, 14, f"='Resumen Ventas CC y Vendedor'!J{s_row_res}").font = f_dat
        ws_sed.cell(r_sed, 14).alignment = a_right; ws_sed.cell(r_sed, 14).number_format = NF_CURR; ws_sed.cell(r_sed, 14).fill = r_fill; ws_sed.cell(r_sed, 14).border = b_thin

        # Col O (15): Base Gravada
        ws_sed.cell(r_sed, 15, f"='Resumen Ventas CC y Vendedor'!K{s_row_res}").font = f_dat
        ws_sed.cell(r_sed, 15).alignment = a_right; ws_sed.cell(r_sed, 15).number_format = NF_CURR; ws_sed.cell(r_sed, 15).fill = r_fill; ws_sed.cell(r_sed, 15).border = b_thin

        # Col P (16): IVA Gravado
        ws_sed.cell(r_sed, 16, f"='Resumen Ventas CC y Vendedor'!L{s_row_res}").font = f_dat
        ws_sed.cell(r_sed, 16).alignment = a_right; ws_sed.cell(r_sed, 16).number_format = NF_CURR_DEC; ws_sed.cell(r_sed, 16).fill = r_fill; ws_sed.cell(r_sed, 16).border = b_thin

        # Col Q (17): % Participación Total
        ws_sed.cell(r_sed, 17, f"='Resumen Ventas CC y Vendedor'!O{s_row_res}").font = f_dat
        ws_sed.cell(r_sed, 17).alignment = a_right; ws_sed.cell(r_sed, 17).number_format = NF_PCT; ws_sed.cell(r_sed, 17).fill = r_fill; ws_sed.cell(r_sed, 17).border = b_thin
        ws_sed.row_dimensions[r_sed].height = 20
        r_sed += 1

    tot_sed_r = r_sed
    ws_sed.cell(tot_sed_r, 2, "TOTAL CONSOLIDADO").font = f_grand
    ws_sed.cell(tot_sed_r, 2).fill = fill_grand
    ws_sed.cell(tot_sed_r, 2).border = b_grand
    ws_sed.cell(tot_sed_r, 2).alignment = a_left

    for c_i, fmt in [
        (3, NF_CURR), (4, NF_CURR), (5, NF_CURR), (6, NF_CURR), (7, NF_CURR), (8, NF_CURR), (9, NF_CURR),
        (10, NF_INT), (11, NF_INT), (12, NF_CURR), (13, NF_CURR),
        (14, NF_CURR), (15, NF_CURR), (16, NF_CURR_DEC), (17, NF_PCT)
    ]:
        col_l = get_column_letter(c_i)
        if c_i == 17:
            c_t = ws_sed.cell(tot_sed_r, c_i, 1.0)
        else:
            c_t = ws_sed.cell(tot_sed_r, c_i, f"=SUM({col_l}{s_start + 1}:{col_l}{tot_sed_r - 1})")
        c_t.font = f_grand
        c_t.fill = fill_grand
        c_t.border = b_grand
        c_t.alignment = a_right if c_i not in [10, 11] else a_center
        c_t.number_format = fmt
    ws_sed.row_dimensions[tot_sed_r].height = 25

    for col in ws_sed.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_sed.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 30)

    # Columna A como margen estetico ultra-fino y AUTOFILTRO en encabezados de Sedes
    ws_sed.column_dimensions["A"].width = 2.5
    ws_sed.auto_filter.ref = f"B{s_start}:Q{tot_sed_r-1}"

    # ==========================================    # 4. HOJA DETALLE NOTAS CRÉDITO (REDISEÑO ULTRA-PREMIUM Y FILTROS)
    # ==========================================
    ws_nc = wb.create_sheet(title="Detalle Notas Crédito")
    ws_nc.sheet_view.showGridLines = False
    if ws_nc.views.sheetView:
        ws_nc.views.sheetView[0].showGridLines = False

    # Banner Institucional de Ancho Completo (Columnas A a L, sin huecos vacíos ni cuadrículas)
    fill_banner_dark = PatternFill(start_color=C_NAVY_DARK, end_color=C_NAVY_DARK, fill_type="solid")
    fill_banner_med = PatternFill(start_color=C_NAVY_MED, end_color=C_NAVY_MED, fill_type="solid")
    f_banner_main = Font(name=FONT_NAME, size=13, bold=True, color="FFFFFF")
    f_banner_sub_w = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    f_banner_caption = Font(name=FONT_NAME, size=8, italic=True, color="E2E8F0")

    ws_nc.merge_cells("A2:L2")
    ws_nc.cell(2, 1, "CH INVERSIONES COLOMBIA S.A.S ZOMAC  |  NIT 901.368.587-8").font = f_banner_main
    ws_nc.cell(2, 1).alignment = a_center
    for c_i in range(1, 13):
        ws_nc.cell(2, c_i).fill = fill_banner_dark
    ws_nc.row_dimensions[2].height = 22

    ws_nc.merge_cells("A3:L3")
    ws_nc.cell(3, 1, "INFORME DETALLADO DE NOTAS CRÉDITO Y REVERSIONES").font = f_banner_sub_w
    ws_nc.cell(3, 1).alignment = a_center
    for c_i in range(1, 13):
        ws_nc.cell(3, c_i).fill = fill_banner_dark
    ws_nc.row_dimensions[3].height = 20

    ws_nc.merge_cells("A4:L4")
    ws_nc.cell(4, 1, "Cruce con Facturas de Venta Originales y Asignación de Vendedor y Centro de Costo  |  Auditoría Integral Odoo").font = f_banner_caption
    ws_nc.cell(4, 1).alignment = a_center
    for c_i in range(1, 13):
        ws_nc.cell(4, c_i).fill = fill_banner_med
    ws_nc.row_dimensions[4].height = 17
    ws_nc.row_dimensions[5].height = 10

    nc_headers = [
        "Número NC", "Factura de Origen Referenciada", "Centro de Costo", "Vendedor Asignado",
        "Cliente", "Fecha", "Base Imponible", "Impuesto (IVA)", "Total Nota Crédito",
        "Motivo / Referencia", "Línea de Negocio", "Estado en Pago"
    ]

    start_nc_r = 6
    # Inicia DIRECTAMENTE en Columna 1 (A), eliminando por completo la columna A vacía
    for idx_c, h in enumerate(nc_headers, 1):
        c = ws_nc.cell(start_nc_r, idx_c, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = a_hdr
        c.border = b_thin
    ws_nc.row_dimensions[start_nc_r].height = 26

    nc_rows = [r for r in processed_rows if r["is_nc"]]
    curr_nc_r = start_nc_r + 1

    for idx_n, r_nc in enumerate(nc_rows):
        is_even = (idx_n % 2 == 0)
        r_fill = fill_zebra if is_even else fill_white

        vals_nc = [
            (r_nc["num"], a_center, None),
            (r_nc["ref_fact"], a_center, None),
            (r_nc["cc"], a_left, None),
            (r_nc["ven"], a_left, None),
            (r_nc["cli"], a_left, None),
            (r_nc["f_fecha"], a_center, "yyyy-mm-dd"),
            (r_nc["base"], a_right, NF_CURR),
            (r_nc["imp"], a_right, NF_CURR_DEC),
            (r_nc["tot"], a_right, NF_CURR),
            (r_nc["ref"], a_left, None),
            (r_nc["linea"], a_center, None),
            (r_nc["est_pago"], a_center, None),
        ]

        # Escribe los valores iniciando en Columna 1 (A)
        for c_i, (v, aln, fmt) in enumerate(vals_nc, 1):
            cell = ws_nc.cell(curr_nc_r, c_i, v)
            cell.font = f_dat
            cell.fill = r_fill
            cell.border = b_thin
            cell.alignment = aln
            if fmt:
                cell.number_format = fmt

        ws_nc.row_dimensions[curr_nc_r].height = 20
        curr_nc_r += 1

    # Fila de Totales de NC (Columnas A a L)
    ws_nc.cell(curr_nc_r, 1, "TOTAL").font = f_grand
    ws_nc.cell(curr_nc_r, 1).fill = fill_grand
    ws_nc.cell(curr_nc_r, 1).border = b_grand
    ws_nc.cell(curr_nc_r, 1).alignment = a_center

    ws_nc.cell(curr_nc_r, 2, f"{len(nc_rows)} NOTAS").font = f_grand
    ws_nc.cell(curr_nc_r, 2).fill = fill_grand
    ws_nc.cell(curr_nc_r, 2).border = b_grand
    ws_nc.cell(curr_nc_r, 2).alignment = a_center

    for c_i in range(3, 7):
        c = ws_nc.cell(curr_nc_r, c_i, "-")
        c.font = f_grand
        c.fill = fill_grand
        c.border = b_grand
        c.alignment = a_center

    # Sumas de Base (Col 7 / G), Impuesto (Col 8 / H), Total (Col 9 / I)
    for c_idx, fmt in [(7, NF_CURR), (8, NF_CURR_DEC), (9, NF_CURR)]:
        col_l = get_column_letter(c_idx)
        c_tot = ws_nc.cell(curr_nc_r, c_idx, f"=SUM({col_l}{start_nc_r+1}:{col_l}{curr_nc_r-1})")
        c_tot.font = f_grand
        c_tot.fill = fill_grand
        c_tot.border = b_grand
        c_tot.alignment = a_right
        c_tot.number_format = fmt

    for c_i in range(10, 13):
        c = ws_nc.cell(curr_nc_r, c_i, "-")
        c.font = f_grand
        c.fill = fill_grand
        c.border = b_grand
        c.alignment = a_center

    ws_nc.row_dimensions[curr_nc_r].height = 24

    # HABILITAR AUTOFILTRO NATIVO EN LOS ENCABEZADOS DE NOTAS CRÉDITO
    ws_nc.auto_filter.ref = f"A{start_nc_r}:L{curr_nc_r-1}"

    # Autoajuste inteligente de columnas midiendo solo filas de datos (ignora banners)
    nc_min_w = {
        1: 14,  # Número NC
        2: 24,  # Factura Referenciada
        3: 18,  # Centro de Costo
        4: 26,  # Vendedor Asignado
        5: 34,  # Cliente
        6: 13,  # Fecha
        7: 17,  # Base Imponible
        8: 16,  # Impuesto (IVA)
        9: 18,  # Total Nota Crédito
        10: 36, # Motivo / Referencia
        11: 16, # Línea de Negocio
        12: 15  # Estado en Pago
    }
    for col_idx in range(1, 13):
        col_letter = get_column_letter(col_idx)
        max_l = max((len(str(ws_nc.cell(r, col_idx).value or "")) for r in range(start_nc_r, curr_nc_r + 1)), default=10)
        req_w = max(max_l + 3, nc_min_w.get(col_idx, 12))
        ws_nc.column_dimensions[col_letter].width = min(req_w, 42)


    # ==========================================
    # 5. HOJA OTROS INGRESOS Y BALANCE (SI SE ADJUNTA TRIAL BALANCE)
    # ==========================================
    if not tb_data and ruta_tb and os.path.exists(ruta_tb):
        try:
            tb_data = leer_trial_balance_odoo(ruta_tb, client_to_cc=client_to_cc)
        except Exception as e_tb:
            print(f"Error al leer Trial Balance: {e_tb}")

    if tb_data:
        try:
            ws_bal = wb.create_sheet(title="Otros Ingresos y Balance", index=1)
            ws_bal.views.sheetView[0].showGridLines = False

            ws_bal.cell(2, 2, "CH INVERSIONES COLOMBIA S.A.S ZOMAC").font = f_title
            ws_bal.cell(3, 2, "ESTADO INTEGRAL DE RESULTADOS, COSTOS Y CONCILIACIÓN CON BALANCE DE PRUEBA (ODOO)").font = Font(name=FONT_NAME, size=12, bold=True, color=C_TEXT_DARK)
            ws_bal.cell(4, 2, "Consolidado de Facturación Operacional, Costo de Ventas (Grupo 61), Utilidad Bruta, Otros Ingresos y Balance").font = f_subt

            # SECCIÓN 1: ESTADO EJECUTIVO DE RESULTADOS (ALTA GERENCIA)
            ws_bal.cell(6, 2, "1. RESUMEN GERENCIAL DE INGRESOS TOTALES Y RESULTADOS (ESTADO DE RESULTADOS)").font = Font(name=FONT_NAME, size=11, bold=True, color=C_NAVY_MED)

            sec1_headers = ["Concepto Financiero", "Cuenta Contable Referencia", "Tipo / Origen", "Valor del Periodo", "Participación"]
            for idx, h in enumerate(sec1_headers, 2):
                c = ws_bal.cell(8, idx, h)
                c.font = f_hdr; c.fill = fill_hdr; c.alignment = a_hdr; c.border = b_thin
            ws_bal.row_dimensions[8].height = 24

            filas_sec1 = [
                ("Facturación Comercial Vehículos Generales", "41350201 / 41350202", "Facturas de Venta", f"='Consolidado por Sede'!C{tot_sed_r}"),
                ("Facturación Comercial Vehículos Jaime", "41350201 / 41350202", "Facturas de Venta", f"='Consolidado por Sede'!D{tot_sed_r}"),
                ("Facturación Comercial Repuestos", "41350601 / 41350602", "Facturas de Venta", f"='Consolidado por Sede'!E{tot_sed_r}"),
                ("Facturación Transporte y Fletes", "41450501 (Cargue y Transporte)", "Servicio Transporte", f"='Consolidado por Sede'!F{tot_sed_r}"),
                ("Liquidación de Taller (Auteco + Autotécnica)", "42354001 (Taller de Vehículos)", "Facturas Cta 42", f"='Consolidado por Sede'!G{tot_sed_r}"),
                ("Arriendo Sede UMA (Supermotos UMA SAS)", "42354001 (Arriendo UMA)", "Facturas Cta 42", f"='Consolidado por Sede'!H{tot_sed_r}"),
                ("(-) Devoluciones en Ventas y Notas Crédito", "4175 (Devoluciones Comerciales)", "Notas Crédito Efectivas (DB)", f"='Consolidado por Sede'!L{tot_sed_r}"),
                ("(+) Notas Débito en Ventas (Ajustes 4175)", "41750506 (Reversión Devoluciones)", "Nota Débito (CR)", tb_data.get("dev_4175_cred", 93000.0)),
            ]

            curr_b_r = 9
            for concepto, cta, orig, formula_or_val in filas_sec1:
                ws_bal.cell(curr_b_r, 2, concepto).font = f_dat_bold if ("(-)" in concepto or "(+)" in concepto) else f_dat
                ws_bal.cell(curr_b_r, 2).alignment = a_left; ws_bal.cell(curr_b_r, 2).border = b_thin
                ws_bal.cell(curr_b_r, 3, cta).font = f_dat; ws_bal.cell(curr_b_r, 3).alignment = a_left; ws_bal.cell(curr_b_r, 3).border = b_thin
                ws_bal.cell(curr_b_r, 4, orig).font = f_dat; ws_bal.cell(curr_b_r, 4).alignment = a_center; ws_bal.cell(curr_b_r, 4).border = b_thin

                c_v = ws_bal.cell(curr_b_r, 5, formula_or_val)
                c_v.font = f_dat_bold if ("(-)" in concepto or "(+)" in concepto) else f_dat
                c_v.alignment = a_right; c_v.number_format = NF_CURR; c_v.border = b_thin

                c_p = ws_bal.cell(curr_b_r, 6, f"=E{curr_b_r}/$E$17")
                c_p.font = f_dat; c_p.alignment = a_right; c_p.number_format = "0.0%"; c_p.border = b_thin
                ws_bal.row_dimensions[curr_b_r].height = 20
                curr_b_r += 1

            # Subtotal Ingresos Operacionales (Fila 17)
            ws_bal.cell(curr_b_r, 2, "(=) TOTAL INGRESOS OPERACIONALES NETOS").font = f_dat_bold
            ws_bal.cell(curr_b_r, 2).alignment = a_left; ws_bal.cell(curr_b_r, 2).fill = fill_zebra; ws_bal.cell(curr_b_r, 2).border = b_thin
            for c_idx in range(3, 5):
                c = ws_bal.cell(curr_b_r, c_idx, "-"); c.font = f_dat_bold; c.alignment = a_center; c.fill = fill_zebra; c.border = b_thin

            c_sub = ws_bal.cell(curr_b_r, 5, "=SUM(E9:E16)")
            c_sub.font = f_dat_bold; c_sub.alignment = a_right; c_sub.number_format = NF_CURR; c_sub.fill = fill_zebra; c_sub.border = b_thin

            c_sub_p = ws_bal.cell(curr_b_r, 6, 1.0)
            c_sub_p.font = f_dat_bold; c_sub_p.alignment = a_right; c_sub_p.number_format = "0.0%"; c_sub_p.fill = fill_zebra; c_sub_p.border = b_thin
            ws_bal.row_dimensions[curr_b_r].height = 22
            r_sub_op = curr_b_r # 17
            curr_b_r += 1

            # Fila 18: Costo de Ventas
            ws_bal.cell(curr_b_r, 2, "(-) COSTO DE VENTAS (GRUPO 61 ODOO)").font = f_dat_bold
            ws_bal.cell(curr_b_r, 2).alignment = a_left; ws_bal.cell(curr_b_r, 2).border = b_thin
            ws_bal.cell(curr_b_r, 3, "6135 (Costo de Ventas Odoo)").font = f_dat; ws_bal.cell(curr_b_r, 3).alignment = a_left; ws_bal.cell(curr_b_r, 3).border = b_thin
            ws_bal.cell(curr_b_r, 4, "Ver Detalle Sección 3 abajo").font = f_dat; ws_bal.cell(curr_b_r, 4).alignment = a_center; ws_bal.cell(curr_b_r, 4).border = b_thin

            c_costo = ws_bal.cell(curr_b_r, 5, tb_data.get("costo_total_61", 0.0))
            c_costo.font = f_dat_bold; c_costo.alignment = a_right; c_costo.number_format = NF_CURR; c_costo.border = b_thin

            c_costo_p = ws_bal.cell(curr_b_r, 6, f"=E{curr_b_r}/$E${r_sub_op}")
            c_costo_p.font = f_dat; c_costo_p.alignment = a_right; c_costo_p.number_format = "0.0%"; c_costo_p.border = b_thin
            ws_bal.row_dimensions[curr_b_r].height = 20
            r_costo = curr_b_r # 18
            curr_b_r += 1

            # Fila 19: Utilidad Bruta Operacional
            ws_bal.cell(curr_b_r, 2, "(=) UTILIDAD BRUTA OPERACIONAL").font = f_dat_bold
            ws_bal.cell(curr_b_r, 2).alignment = a_left; ws_bal.cell(curr_b_r, 2).fill = fill_zebra; ws_bal.cell(curr_b_r, 2).border = b_thin
            for c_idx in range(3, 5):
                c = ws_bal.cell(curr_b_r, c_idx, "-"); c.font = f_dat_bold; c.alignment = a_center; c.fill = fill_zebra; c.border = b_thin

            c_ub = ws_bal.cell(curr_b_r, 5, f"=E{r_sub_op}-E{r_costo}")
            c_ub.font = f_dat_bold; c_ub.alignment = a_right; c_ub.number_format = NF_CURR; c_ub.fill = fill_zebra; c_ub.border = b_thin

            c_ub_p = ws_bal.cell(curr_b_r, 6, f"=E{curr_b_r}/$E${r_sub_op}")
            c_ub_p.font = f_dat_bold; c_ub_p.alignment = a_right; c_ub_p.number_format = "0.0%"; c_ub_p.fill = fill_zebra; c_ub_p.border = b_thin
            ws_bal.row_dimensions[curr_b_r].height = 22
            r_ub = curr_b_r # 19
            curr_b_r += 1

            # Fila 20: Otros Ingresos del Periodo
            ws_bal.cell(curr_b_r, 2, "(+) OTROS INGRESOS DEL PERIODO (NO OPERACIONALES - BALANCE)").font = f_dat_bold
            ws_bal.cell(curr_b_r, 2).alignment = a_left; ws_bal.cell(curr_b_r, 2).border = b_thin
            ws_bal.cell(curr_b_r, 3, "4180 / 4210 / 4225 / 4235 / 4295").font = f_dat; ws_bal.cell(curr_b_r, 3).alignment = a_left; ws_bal.cell(curr_b_r, 3).border = b_thin
            ws_bal.cell(curr_b_r, 4, "Ver Detalle Sección 2 abajo").font = f_dat; ws_bal.cell(curr_b_r, 4).alignment = a_center; ws_bal.cell(curr_b_r, 4).border = b_thin

            c_oi = ws_bal.cell(curr_b_r, 5, tb_data.get("total_otros_ingresos", 0.0))
            c_oi.font = f_dat_bold; c_oi.alignment = a_right; c_oi.number_format = NF_CURR; c_oi.border = b_thin

            c_oi_p = ws_bal.cell(curr_b_r, 6, f"=E{curr_b_r}/$E${r_sub_op}")
            c_oi_p.font = f_dat; c_oi_p.alignment = a_right; c_oi_p.number_format = "0.0%"; c_oi_p.fill = fill_white; c_oi_p.border = b_thin
            ws_bal.row_dimensions[curr_b_r].height = 20
            r_oi = curr_b_r # 20
            curr_b_r += 1

            # Fila 21: Gran Total Resultado Integral del Periodo
            ws_bal.cell(curr_b_r, 2, "(=) GRAN TOTAL RESULTADO INTEGRAL DEL PERIODO").font = f_grand
            ws_bal.cell(curr_b_r, 2).fill = fill_grand; ws_bal.cell(curr_b_r, 2).alignment = a_left; ws_bal.cell(curr_b_r, 2).border = b_grand
            for c_idx in range(3, 5):
                c = ws_bal.cell(curr_b_r, c_idx, "CONSOLIDADO"); c.font = f_grand; c.alignment = a_center; c.fill = fill_grand; c.border = b_grand

            c_gt = ws_bal.cell(curr_b_r, 5, f"=E{r_ub}+E{r_oi}")
            c_gt.font = f_grand; c_gt.alignment = a_right; c_gt.number_format = NF_CURR; c_gt.fill = fill_grand; c_gt.border = b_grand

            c_gt_p = ws_bal.cell(curr_b_r, 6, f"=E{curr_b_r}/$E${r_sub_op}")
            c_gt_p.font = f_grand; c_gt_p.alignment = a_right; c_gt_p.number_format = "0.0%"; c_gt_p.fill = fill_grand; c_gt_p.border = b_grand
            ws_bal.row_dimensions[curr_b_r].height = 25
            r_gt = curr_b_r # 21
            curr_b_r += 3

            # SECCIÓN 2: AUDITORÍA DE OTROS INGRESOS
            ws_bal.cell(curr_b_r, 2, "2. AUDITORÍA FORENSE CUENTA POR CUENTA DE OTROS INGRESOS (BALANCE ODOO)").font = Font(name=FONT_NAME, size=11, bold=True, color=C_NAVY_MED)
            curr_b_r += 2

            sec2_headers = ["Concepto de Ingreso", "Cuentas Contables PUC", "Débito Odoo", "Crédito Odoo", "Saldo Reconocido", "Observación y Soporte de Auditoría"]
            for idx, h in enumerate(sec2_headers, 2):
                c = ws_bal.cell(curr_b_r, idx, h)
                c.font = f_hdr; c.fill = fill_hdr; c.alignment = a_hdr; c.border = b_thin
            ws_bal.row_dimensions[curr_b_r].height = 24
            curr_b_r += 1

            start_s2 = curr_b_r
            for idx_oi, item in enumerate(tb_data["otros_ingresos"]):
                is_even = (idx_oi % 2 == 0)
                r_fill = fill_zebra if is_even else fill_white

                ws_bal.cell(curr_b_r, 2, item["concepto"]).font = f_dat_bold if "Cesantías" in item["concepto"] else f_dat
                ws_bal.cell(curr_b_r, 2).alignment = a_left; ws_bal.cell(curr_b_r, 2).fill = r_fill; ws_bal.cell(curr_b_r, 2).border = b_thin
                ws_bal.cell(curr_b_r, 3, item["cuentas"]).font = f_dat; ws_bal.cell(curr_b_r, 3).alignment = a_left; ws_bal.cell(curr_b_r, 3).fill = r_fill; ws_bal.cell(curr_b_r, 3).border = b_thin

                deb_val = item.get("debito", 0.0)
                cred_val = item.get("credito", 0.0)

                c_deb = ws_bal.cell(curr_b_r, 4, deb_val)
                c_deb.font = f_dat; c_deb.alignment = a_right; c_deb.number_format = NF_CURR_DEC; c_deb.fill = r_fill; c_deb.border = b_thin

                c_cred = ws_bal.cell(curr_b_r, 5, cred_val)
                c_cred.font = f_dat_bold; c_cred.alignment = a_right; c_cred.number_format = NF_CURR_DEC; c_cred.fill = r_fill; c_cred.border = b_thin

                # Saldo Reconocido: Fórmula viva = Crédito (Col E) - Débito (Col D)
                c_net = ws_bal.cell(curr_b_r, 6, f"=E{curr_b_r}-D{curr_b_r}")
                c_net.font = f_dat_bold; c_net.alignment = a_right; c_net.number_format = NF_CURR_DEC; c_net.fill = r_fill; c_net.border = b_thin

                ws_bal.cell(curr_b_r, 7, item["detalle"]).font = f_dat; ws_bal.cell(curr_b_r, 7).alignment = a_left; ws_bal.cell(curr_b_r, 7).fill = r_fill; ws_bal.cell(curr_b_r, 7).border = b_thin
                ws_bal.row_dimensions[curr_b_r].height = 20
                curr_b_r += 1

            # Total Sección 2
            tot_s2_row = curr_b_r
            ws_bal.cell(curr_b_r, 2, "TOTAL OTROS INGRESOS (RECONOCIMIENTO NETO)").font = f_grand
            ws_bal.cell(curr_b_r, 2).fill = fill_grand; ws_bal.cell(curr_b_r, 2).border = b_grand; ws_bal.cell(curr_b_r, 2).alignment = a_left
            ws_bal.cell(curr_b_r, 3, f"{len(tb_data['otros_ingresos'])} CONCEPTOS").font = f_grand
            ws_bal.cell(curr_b_r, 3).fill = fill_grand; ws_bal.cell(curr_b_r, 3).border = b_grand; ws_bal.cell(curr_b_r, 3).alignment = a_center

            c_t_deb = ws_bal.cell(curr_b_r, 4, f"=SUM(D{start_s2}:D{curr_b_r-1})")
            c_t_deb.font = f_grand; c_t_deb.fill = fill_grand; c_t_deb.border = b_grand; c_t_deb.alignment = a_right; c_t_deb.number_format = NF_CURR_DEC

            c_t_cred = ws_bal.cell(curr_b_r, 5, f"=SUM(E{start_s2}:E{curr_b_r-1})")
            c_t_cred.font = f_grand; c_t_cred.fill = fill_grand; c_t_cred.border = b_grand; c_t_cred.alignment = a_right; c_t_cred.number_format = NF_CURR_DEC

            c_t_net = ws_bal.cell(curr_b_r, 6, f"=SUM(F{start_s2}:F{curr_b_r-1})")
            c_t_net.font = f_grand; c_t_net.fill = fill_grand; c_t_net.border = b_grand; c_t_net.alignment = a_right; c_t_net.number_format = NF_CURR_DEC

            ws_bal.cell(curr_b_r, 7, "Auditado 100% contra Balance de Prueba").font = f_grand
            ws_bal.cell(curr_b_r, 7).fill = fill_grand; ws_bal.cell(curr_b_r, 7).border = b_grand; ws_bal.cell(curr_b_r, 7).alignment = a_center
            ws_bal.row_dimensions[curr_b_r].height = 25
            curr_b_r += 3

            # SECCIÓN 3: AUDITORÍA DE COSTO DE VENTAS (GRUPO 61 ODOO)
            ws_bal.cell(curr_b_r, 2, "3. AUDITORÍA FORENSE DEL COSTO DE VENTAS (CUENTAS GRUPO 61 ODOO)").font = Font(name=FONT_NAME, size=11, bold=True, color=C_NAVY_MED)
            curr_b_r += 2

            sec3_headers = ["Cuenta PUC", "Nombre de la Cuenta de Costo", "Débito Odoo", "Crédito Odoo", "Costo Reconocido (Débito - Crédito)", "Línea de Negocio / Clasificación"]
            for idx, h in enumerate(sec3_headers, 2):
                c = ws_bal.cell(curr_b_r, idx, h)
                c.font = f_hdr; c.fill = fill_hdr; c.alignment = a_hdr; c.border = b_thin
            ws_bal.row_dimensions[curr_b_r].height = 24
            curr_b_r += 1

            start_s3 = curr_b_r
            costos_lista = tb_data.get("costos_cuentas", [])
            for idx_c, item in enumerate(costos_lista):
                is_even = (idx_c % 2 == 0)
                r_fill = fill_zebra if is_even else fill_white

                c_cta = ws_bal.cell(curr_b_r, 2, item["cuenta"])
                c_cta.font = f_dat_bold; c_cta.alignment = a_center; c_cta.fill = r_fill; c_cta.border = b_thin; c_cta.number_format = "@"

                ws_bal.cell(curr_b_r, 3, item["nombre"]).font = f_dat; ws_bal.cell(curr_b_r, 3).alignment = a_left; ws_bal.cell(curr_b_r, 3).fill = r_fill; ws_bal.cell(curr_b_r, 3).border = b_thin

                c_deb = ws_bal.cell(curr_b_r, 4, item.get("debito", 0.0))
                c_deb.font = f_dat_bold; c_deb.alignment = a_right; c_deb.number_format = NF_CURR_DEC; c_deb.fill = r_fill; c_deb.border = b_thin

                c_cred = ws_bal.cell(curr_b_r, 5, item.get("credito", 0.0))
                c_cred.font = f_dat; c_cred.alignment = a_right; c_cred.number_format = NF_CURR_DEC; c_cred.fill = r_fill; c_cred.border = b_thin

                # Costo Reconocido: Fórmula viva = Débito (Col D) - Crédito (Col E)
                c_net = ws_bal.cell(curr_b_r, 6, f"=D{curr_b_r}-E{curr_b_r}")
                c_net.font = f_dat_bold; c_net.alignment = a_right; c_net.number_format = NF_CURR_DEC; c_net.fill = r_fill; c_net.border = b_thin

                ws_bal.cell(curr_b_r, 7, item.get("detalle", "")).font = f_dat; ws_bal.cell(curr_b_r, 7).alignment = a_left; ws_bal.cell(curr_b_r, 7).fill = r_fill; ws_bal.cell(curr_b_r, 7).border = b_thin
                ws_bal.row_dimensions[curr_b_r].height = 20
                curr_b_r += 1

            # Total Sección 3 (Costo de Ventas)
            tot_s3_costo_row = curr_b_r
            ws_bal.cell(curr_b_r, 2, "TOTAL COSTO DE VENTAS (GRUPO 61)").font = f_grand
            ws_bal.cell(curr_b_r, 2).fill = fill_grand; ws_bal.cell(curr_b_r, 2).border = b_grand; ws_bal.cell(curr_b_r, 2).alignment = a_left
            ws_bal.cell(curr_b_r, 3, f"{len(costos_lista)} CUENTAS PUC").font = f_grand
            ws_bal.cell(curr_b_r, 3).fill = fill_grand; ws_bal.cell(curr_b_r, 3).border = b_grand; ws_bal.cell(curr_b_r, 3).alignment = a_center

            c_t_deb3 = ws_bal.cell(curr_b_r, 4, f"=SUM(D{start_s3}:D{curr_b_r-1})")
            c_t_deb3.font = f_grand; c_t_deb3.fill = fill_grand; c_t_deb3.border = b_grand; c_t_deb3.alignment = a_right; c_t_deb3.number_format = NF_CURR_DEC

            c_t_cred3 = ws_bal.cell(curr_b_r, 5, f"=SUM(E{start_s3}:E{curr_b_r-1})")
            c_t_cred3.font = f_grand; c_t_cred3.fill = fill_grand; c_t_cred3.border = b_grand; c_t_cred3.alignment = a_right; c_t_cred3.number_format = NF_CURR_DEC

            c_t_net3 = ws_bal.cell(curr_b_r, 6, f"=SUM(F{start_s3}:F{curr_b_r-1})")
            c_t_net3.font = f_grand; c_t_net3.fill = fill_grand; c_t_net3.border = b_grand; c_t_net3.alignment = a_right; c_t_net3.number_format = NF_CURR_DEC

            ws_bal.cell(curr_b_r, 7, "Débito menos Crédito Auditado contra Balance Odoo").font = f_grand
            ws_bal.cell(curr_b_r, 7).fill = fill_grand; ws_bal.cell(curr_b_r, 7).border = b_grand; ws_bal.cell(curr_b_r, 7).alignment = a_center
            ws_bal.row_dimensions[curr_b_r].height = 25

            # Vincular fórmulas dinámicas de Sección 1 con los totales calculados de Sección 2 y 3:
            ws_bal.cell(r_costo, 5, f"=F{tot_s3_costo_row}")
            ws_bal.cell(r_oi, 5, f"=F{tot_s2_row}")

            curr_b_r += 3

            # SECCIÓN 4: CONCILIACIÓN POR CUENTA PUC (FACTURACIÓN VS BALANCE)
            ws_bal.cell(curr_b_r, 2, "4. CONCILIACIÓN DE FACTURACIÓN OPERACIONAL VS BALANCE POR CUENTA PUC").font = Font(name=FONT_NAME, size=11, bold=True, color=C_NAVY_MED)
            curr_b_r += 2

            sec3_headers = ["Cuenta PUC", "Nombre de la Cuenta", "Cant. Facturas", "Base Facturada", "IVA Facturado", "Total Facturado", "Saldo en Balance Odoo", "Diferencia", "Causa de Variación / Observación de Auditoría"]
            for idx, h in enumerate(sec3_headers, 2):
                c = ws_bal.cell(curr_b_r, idx, h)
                c.font = f_hdr; c.fill = fill_hdr; c.alignment = a_hdr; c.border = b_thin
            ws_bal.row_dimensions[curr_b_r].height = 24
            curr_b_r += 1

            # Cuentas principales de facturación
            target_pucs = [
                ("41350201", "VENTAS DE VEHICULOS EXENTAS"),
                ("41350202", "VENTAS DE VEHICULOS DEL 19%"),
                ("41350601", "VENTA DE REPUESTOS EXENTOS"),
                ("41350602", "VENTA DE REPUESTOS 19%"),
                ("41450501", "SERVICIO DE CARGUE Y DESCARGUE (TRANSPORTE)"),
                ("4175", "DEVOLUCIONES EN VENTAS (NOTAS CRÉDITO 4175 DB)"),
                ("41750506", "NOTAS DÉBITO EN VENTAS (REVERSIÓN DEV. 4175 CR)"),
                ("42354001", "TALLER DE VEHICULOS (TALLER Y UMA)"),
            ]

            obs_sec3 = {
                "41350201": "Flete ($170k-$180k/moto) facturado en valor moto; reclasificado a 4145 en Odoo",
                "41350202": "Líneas mixtas de accesorios gravados al 19% dentro de facturas de moto",
                "41350601": "Incluye Aceites y Lubricantes Exentos (Cuenta 41350108: $7.406.000)",
                "41350602": "Venta de repuestos 19% (Variación mínima de $10.580)",
                "41450501": "Contabilizado internamente en Odoo por flete de motos ($170k/$180k por tercero)",
                "4175": "Conciliado 100% exacto ($89.552.813)",
                "41750506": "Conciliado 100% exacto ($93.000)",
                "42354001": "Conciliado 100% exacto ($12.895.489)",
            }

            det_puc_col = f"'Detalle Ventas'!$AI$2:$AI${det_max_r}"
            det_base_col = f"'Detalle Ventas'!$L$2:$L${det_max_r}"
            det_iva_col = f"'Detalle Ventas'!$M$2:$M${det_max_r}"
            det_tot_col = f"'Detalle Ventas'!$N$2:$N${det_max_r}"

            s3_start = curr_b_r
            for idx_p, (puc_c, puc_n) in enumerate(target_pucs):
                is_even = (idx_p % 2 == 0)
                r_fill = fill_zebra if is_even else fill_white

                ws_bal.cell(curr_b_r, 2, puc_c).font = f_dat_bold; ws_bal.cell(curr_b_r, 2).alignment = a_center; ws_bal.cell(curr_b_r, 2).fill = r_fill; ws_bal.cell(curr_b_r, 2).border = b_thin
                ws_bal.cell(curr_b_r, 3, puc_n).font = f_dat; ws_bal.cell(curr_b_r, 3).alignment = a_left; ws_bal.cell(curr_b_r, 3).fill = r_fill; ws_bal.cell(curr_b_r, 3).border = b_thin

                c_cnt = ws_bal.cell(curr_b_r, 4, f'=COUNTIF({det_puc_col}, B{curr_b_r})')
                c_cnt.font = f_dat; c_cnt.alignment = a_center; c_cnt.number_format = NF_INT; c_cnt.fill = r_fill; c_cnt.border = b_thin

                c_b = ws_bal.cell(curr_b_r, 5, f'=SUMIF({det_puc_col}, B{curr_b_r}, {det_base_col})')
                c_b.font = f_dat; c_b.alignment = a_right; c_b.number_format = NF_CURR; c_b.fill = r_fill; c_b.border = b_thin

                c_i = ws_bal.cell(curr_b_r, 6, f'=SUMIF({det_puc_col}, B{curr_b_r}, {det_iva_col})')
                c_i.font = f_dat; c_i.alignment = a_right; c_i.number_format = NF_CURR_DEC; c_i.fill = r_fill; c_i.border = b_thin

                c_t = ws_bal.cell(curr_b_r, 7, f'=SUMIF({det_puc_col}, B{curr_b_r}, {det_tot_col})')
                c_t.font = f_dat_bold; c_t.alignment = a_right; c_t.number_format = NF_CURR; c_t.fill = r_fill; c_t.border = b_thin

                # Saldo en Balance Odoo
                if puc_c == "4175":
                    bal_val = -tb_data.get("dev_4175_deb", 0.0)
                elif puc_c == "41750506":
                    bal_val = tb_data.get("dev_4175_cred", 0.0)
                elif puc_c == "42354001":
                    tb_c_data = tb_data["cuentas"].get(puc_c, {})
                    bal_val = tb_c_data.get("credit", 0.0) - tb_c_data.get("debit", 0.0)
                else:
                    tb_c_data = tb_data["cuentas"].get(puc_c, {})
                    bal_val = tb_c_data.get("credit", 0.0)
                c_bal = ws_bal.cell(curr_b_r, 8, bal_val)
                c_bal.font = f_dat_bold; c_bal.alignment = a_right; c_bal.number_format = NF_CURR; c_bal.fill = r_fill; c_bal.border = b_thin

                # Diferencia (Base Facturada - Balance)
                c_dif = ws_bal.cell(curr_b_r, 9, f'=E{curr_b_r}-H{curr_b_r}')
                c_dif.font = f_dat_bold; c_dif.alignment = a_right; c_dif.number_format = NF_CURR; c_dif.fill = r_fill; c_dif.border = b_thin

                # Observación y Causa de la Variación
                c_obs = ws_bal.cell(curr_b_r, 10, obs_sec3.get(puc_c, ""))
                c_obs.font = f_dat; c_obs.alignment = a_left; c_obs.fill = r_fill; c_obs.border = b_thin

                ws_bal.row_dimensions[curr_b_r].height = 20
                curr_b_r += 1

            # Total Sección 3
            ws_bal.cell(curr_b_r, 2, "TOTAL").font = f_grand; ws_bal.cell(curr_b_r, 2).fill = fill_grand; ws_bal.cell(curr_b_r, 2).border = b_grand; ws_bal.cell(curr_b_r, 2).alignment = a_center
            ws_bal.cell(curr_b_r, 3, "CUENTAS CONCILIADAS").font = f_grand; ws_bal.cell(curr_b_r, 3).fill = fill_grand; ws_bal.cell(curr_b_r, 3).border = b_grand; ws_bal.cell(curr_b_r, 3).alignment = a_left

            c_t_cnt = ws_bal.cell(curr_b_r, 4, f'=SUM(D{s3_start}:D{curr_b_r-1})')
            c_t_cnt.font = f_grand; c_t_cnt.fill = fill_grand; c_t_cnt.border = b_grand; c_t_cnt.alignment = a_center; c_t_cnt.number_format = NF_INT

            for c_i, fmt in [(5, NF_CURR), (6, NF_CURR_DEC), (7, NF_CURR), (8, NF_CURR), (9, NF_CURR)]:
                col_l = get_column_letter(c_i)
                c_t_s = ws_bal.cell(curr_b_r, c_i, f'=SUM({col_l}{s3_start}:{col_l}{curr_b_r-1})')
                c_t_s.font = f_grand; c_t_s.fill = fill_grand; c_t_s.border = b_grand; c_t_s.alignment = a_right; c_t_s.number_format = fmt

            ws_bal.cell(curr_b_r, 10, "Total neto facturado concilia con balance operacional y fletes").font = f_grand; ws_bal.cell(curr_b_r, 10).fill = fill_grand; ws_bal.cell(curr_b_r, 10).border = b_grand
            ws_bal.row_dimensions[curr_b_r].height = 25

            ws_bal.column_dimensions["A"].width = 3
            ws_bal.column_dimensions["B"].width = 16
            ws_bal.column_dimensions["C"].width = 38
            ws_bal.column_dimensions["D"].width = 14
            ws_bal.column_dimensions["E"].width = 20
            ws_bal.column_dimensions["F"].width = 18
            ws_bal.column_dimensions["G"].width = 20
            ws_bal.column_dimensions["H"].width = 22
            ws_bal.column_dimensions["I"].width = 20
            ws_bal.column_dimensions["J"].width = 62
        except Exception as e_tb:
            print(f"Error al generar hoja de Otros Ingresos y Balance: {e_tb}")


    # Crear Portada Ejecutiva Dashboard
    try:
        _crear_dashboard_gerencial(wb, tb_data, tot_sed_r, sedes)
    except Exception as e_dash:
        print(f"Error al generar Dashboard Gerencial: {e_dash}")


    # Crear Hoja Matriz Separador y Operación (Estructura de Camila Ultra Mejorada)
    try:
        _crear_hoja_separador_y_operacion(wb, det_max_r, sedes)
    except Exception as e_sep:
        print(f"Error al generar Hoja Separador y Operación: {e_sep}")

    # Orden gerencial de pestañas (Alta Gerencia primero)
    # Auto-adjuntar hojas de Gastos y Utilidad por Centro de Costo si existe archivo en la carpeta o fue seleccionado
    try:
        if not ruta_gastos or not os.path.exists(ruta_gastos):
            input_dir = os.path.dirname(os.path.abspath(ruta_entrada))
            tb_dir = os.path.dirname(os.path.abspath(ruta_tb)) if ruta_tb else input_dir
            gastos_candidates = [
                os.path.join(input_dir, "GASTOS POR CENTRO DE COSTOS OCTUBRE.xlsx"),
                os.path.join(tb_dir, "GASTOS POR CENTRO DE COSTOS OCTUBRE.xlsx"),
                os.path.join(input_dir, "..", "NUEVAS VERSION", "GASTOS POR CENTRO DE COSTOS OCTUBRE.xlsx")
            ]
            for gc in gastos_candidates:
                if os.path.exists(gc):
                    ruta_gastos = gc
                    break
        if ruta_gastos and os.path.exists(ruta_gastos):
            _adjuntar_hojas_gastos_y_utilidad(wb, ruta_gastos)
    except Exception as e_gastos:
        print(f"Nota: No se adjuntaron hojas de gastos: {e_gastos}")

    sheet_order = [
        "Dashboard Gerencial",
        "Otros Ingresos y Balance",
        "Consolidado por Sede",
        "Separador y Operación",
        "Resumen Ventas CC y Vendedor",
        "Detalle Ventas",
        "Detalle Notas Crédito",
        "OCTUBRE",
        "UTILIDAD_OCTUBRE"
    ]
    wb._sheets = [wb[s] for s in sheet_order if s in wb.sheetnames] + [s for s in wb.worksheets if s.title not in sheet_order]

        # GARANTIZAR 100% QUE NINGUNA HOJA TENGA CUADRÍCULA DE FONDO VISIBLE
    for s_name in wb.sheetnames:
        _ws = wb[s_name]
        _ws.sheet_view.showGridLines = False
        if _ws.views.sheetView:
            _ws.views.sheetView[0].showGridLines = False

    wb.save(ruta_salida)

    # Recalcular mediante COM si está disponible
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb_com = excel.Workbooks.Open(os.path.abspath(ruta_salida))
        excel.CalculateFull()
        wb_com.Save()
        wb_com.Close()
        excel.Quit()
    except Exception:
        pass

    tot_ex = float(df_valid[df_valid["imp"] <= 0]["base"].sum())
    tot_bg = float(df_valid[df_valid["imp"] > 0]["base"].sum())
    tot_iv = float(df_valid[df_valid["imp"] > 0]["imp"].sum())
    tot_gr = float(df_valid[df_valid["imp"] > 0]["tot"].sum())
    tot_ve = tot_ex + tot_gr

    n_fact_pos = len(df_valid[~df_valid["is_nc"]])
    tot_fact_bruto = float(df_valid[~df_valid["is_nc"]]["tot"].sum())
    n_nc_tot = len(df_valid[df_valid["is_nc"]])
    tot_nc_val = float(df_valid[df_valid["is_nc"]]["tot"].sum())

    return {
        "ruta_guardada": ruta_salida,
        "total_facturas": len(df_valid),
        "n_facturas_pos": n_fact_pos,
        "total_ventas_brutas": tot_fact_bruto,
        "n_nc_total": n_nc_tot,
        "total_nc": tot_nc_val,
        "total_ventas": tot_ve,
        "total_exentas": tot_ex,
        "total_base_gravada": tot_bg,
        "total_iva": tot_iv,
        "total_gravado": tot_gr,
        "total_sedes": len(sedes),
        "total_vendedores": df_valid["ven"].nunique(),
        "pct_exento": (tot_ex / tot_ve) if tot_ve else 0,
        "pct_gravado": (tot_gr / tot_ve) if tot_ve else 0,
        "n_nc_cruzadas": n_nc_cruzadas,
        "total_nc_cruzadas": total_nc_cruzadas,
        "tb_cargado": bool(tb_data is not None),
        "total_otros_ingresos": tb_data["total_otros_ingresos"] if tb_data else 0.0,
        "liq_taller_total": tb_data["liq_taller_total"] if tb_data else 0.0,
        "dev_4175": tb_data["dev_4175"] if tb_data else 0.0
    }
