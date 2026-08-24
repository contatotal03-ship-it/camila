import os
import json

CONFIG_FILE = "configuracion_conciliador.json"

DEFAULT_CONFIG = {
    "compras": {
        "hojas_validas": [
            "COMPRAS",
            "FACTURA DE COMPRA",
            "DOCUMENTO EQUIVALENTE A FACTURA",
            "DOC. SOPORTE"
        ],
        "dian_incluir": ["soporte", "peaje"],
        "dian_grupo": "Recibido",
        "dian_tipos_validos": [
            "Factura electrónica de Venta",
            "Documento equivalente - Cobro de peajes",
            "Factura Soporte en adquisiciones efectuadas a sujetos no obligados a expedir factura o documento equivalente",
            "Nota de Crédito electrónica"
        ],
        "cuentas_iva": ["2408"],
        "cuentas_total": ["2205", "2365"]
    },
    "ventas": {
        "hojas_validas": [
            "FACTURACION SAN JOSE GUAVIARE",
            "Facturación Villavicencio",
            "FACTURACION INIRIDA",
            "Facturación Granada",
            "FACTURACION REPUESTOS GM",
            "FACTURACION CARRENO",
            "FACTURACION REPUESTOS SJ",
            "DIAN Notas Credito FEVC"
        ],
        "dian_excluir": ["soporte", "peaje"],
        "dian_grupo": "Emitido",
        "dian_tipos_validos": [
            "Factura electrónica de Venta",
            "Nota de Crédito electrónica"
        ],
        "cuentas_iva": ["2408"],
        "cuentas_excluir": ["6", "14"],
        "cuentas_transporte": ["413572"]
    }
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_CONFIG, f, indent=4, ensure_ascii=False)
        return DEFAULT_CONFIG
    
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        try:
            return json.load(f)
        except:
            return DEFAULT_CONFIG

def get_hojas_compras():
    c = load_config()
    return [h.upper().strip() for h in c.get('compras', {}).get('hojas_validas', DEFAULT_CONFIG['compras']['hojas_validas'])]

def get_hojas_ventas():
    c = load_config()
    return [h.upper().strip() for h in c.get('ventas', {}).get('hojas_validas', DEFAULT_CONFIG['ventas']['hojas_validas'])]

def get_dian_compras_incluir():
    c = load_config()
    return [t.lower().strip() for t in c.get('compras', {}).get('dian_incluir', DEFAULT_CONFIG['compras']['dian_incluir'])]

def get_dian_ventas_excluir():
    c = load_config()
    return [t.lower().strip() for t in c.get('ventas', {}).get('dian_excluir', DEFAULT_CONFIG['ventas']['dian_excluir'])]

def get_dian_grupo_compras():
    c = load_config()
    return c.get('compras', {}).get('dian_grupo', DEFAULT_CONFIG['compras']['dian_grupo'])

def get_dian_grupo_ventas():
    c = load_config()
    return c.get('ventas', {}).get('dian_grupo', DEFAULT_CONFIG['ventas']['dian_grupo'])

def get_dian_tipos_validos_compras():
    c = load_config()
    return [t.strip() for t in c.get('compras', {}).get('dian_tipos_validos', DEFAULT_CONFIG['compras']['dian_tipos_validos'])]

def get_dian_tipos_validos_ventas():
    c = load_config()
    return [t.strip() for t in c.get('ventas', {}).get('dian_tipos_validos', DEFAULT_CONFIG['ventas']['dian_tipos_validos'])]

def get_cuentas_iva_compras():
    c = load_config()
    return [t.strip() for t in c.get('compras', {}).get('cuentas_iva', DEFAULT_CONFIG['compras']['cuentas_iva'])]

def get_cuentas_total_compras():
    c = load_config()
    return [t.strip() for t in c.get('compras', {}).get('cuentas_total', DEFAULT_CONFIG['compras']['cuentas_total'])]

def get_cuentas_iva_ventas():
    c = load_config()
    return [t.strip() for t in c.get('ventas', {}).get('cuentas_iva', DEFAULT_CONFIG['ventas']['cuentas_iva'])]

def get_cuentas_excluir_ventas():
    c = load_config()
    return [t.strip() for t in c.get('ventas', {}).get('cuentas_excluir', DEFAULT_CONFIG['ventas']['cuentas_excluir'])]
def get_cuentas_transporte_ventas():
    c = load_config()
    return [t.strip() for t in c.get('ventas', {}).get('cuentas_transporte', ["413572"])]
