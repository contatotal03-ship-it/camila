import pandas as pd
import zipfile
import io
import os

def robust_read_excel(path_or_file, **kwargs):
    """
    Intenta leer un archivo Excel con pandas. Si falla debido a un error de
    hoja de estilos (común en exportaciones de Odoo), intenta reparar el archivo
    en memoria eliminando la hoja de estilos corrupta.
    """
    try:
        # Intento normal
        return pd.read_excel(path_or_file, **kwargs)
    except Exception as e:
        error_msg = str(e)
        # El error de openpyxl suele ser "could not read stylesheet"
        if "could not read stylesheet" in error_msg or "stylesheet" in error_msg.lower():
            print(f"Detectado error de estilos en Excel. Intentando auto-reparación...")
            try:
                # Si es una ruta (string), abrimos el archivo
                if isinstance(path_or_file, str) and os.path.exists(path_or_file):
                    with open(path_or_file, 'rb') as f:
                        file_data = f.read()
                elif hasattr(path_or_file, 'read'):
                    # Si es un objeto file-like, leemos su contenido
                    if hasattr(path_or_file, 'seek'):
                        path_or_file.seek(0)
                    file_data = path_or_file.read()
                else:
                    raise e # No sabemos cómo manejarlo, relanzamos original
                
                # Reparación en memoria
                in_buffer = io.BytesIO(file_data)
                out_buffer = io.BytesIO()
                
                with zipfile.ZipFile(in_buffer, 'r') as zin:
                    with zipfile.ZipFile(out_buffer, 'w') as zout:
                        for item in zin.infolist():
                            # Omitimos el archivo de estilos corrupto
                            if item.filename == 'xl/styles.xml':
                                continue
                            zout.writestr(item, zin.read(item.filename))
                
                out_buffer.seek(0)
                # Reintentamos con el archivo "reparado"
                # Forzamos engine='openpyxl' ya que sabemos que el problema es con esa librería
                if 'engine' not in kwargs:
                    kwargs['engine'] = 'openpyxl'
                return pd.read_excel(out_buffer, **kwargs)
            except Exception as repair_e:
                print(f"La auto-reparación falló: {repair_e}")
                raise e # Lanzamos el error original si la reparación falla
        else:
            raise e

def robust_ExcelFile(path_or_file, **kwargs):
    """
    Versión robusta de pd.ExcelFile.
    """
    try:
        return pd.ExcelFile(path_or_file, **kwargs)
    except Exception as e:
        error_msg = str(e)
        if "could not read stylesheet" in error_msg or "stylesheet" in error_msg.lower():
            print(f"Detectado error de estilos en Excel (ExcelFile). Intentando auto-reparación...")
            try:
                if isinstance(path_or_file, str) and os.path.exists(path_or_file):
                    with open(path_or_file, 'rb') as f:
                        file_data = f.read()
                elif hasattr(path_or_file, 'read'):
                    if hasattr(path_or_file, 'seek'):
                        path_or_file.seek(0)
                    file_data = path_or_file.read()
                else:
                    raise e
                
                in_buffer = io.BytesIO(file_data)
                out_buffer = io.BytesIO()
                
                with zipfile.ZipFile(in_buffer, 'r') as zin:
                    with zipfile.ZipFile(out_buffer, 'w') as zout:
                        for item in zin.infolist():
                            if item.filename == 'xl/styles.xml':
                                continue
                            zout.writestr(item, zin.read(item.filename))
                
                out_buffer.seek(0)
                if 'engine' not in kwargs:
                    kwargs['engine'] = 'openpyxl'
                return pd.ExcelFile(out_buffer, **kwargs)
            except Exception:
                raise e
        else:
            raise e
