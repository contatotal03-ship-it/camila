"""
MÓDULO DE SEGURIDAD Y LICENCIAMIENTO CLOUD
Camila Conciliador
"""

import subprocess
import hashlib
import os
import sys
import json
import base64
import time
import socket
import ssl
from datetime import datetime, timedelta
import urllib.request
import urllib.error

# Resiliencia SSL para Windows
try:
    if hasattr(ssl, "_create_unverified_context"):
        ssl._create_default_https_context = ssl._create_unverified_context
except Exception:
    pass

SECRET_KEY = b"FISCALYTIX_SECURE_CLOUD_KEY_2026_DIANFLOW_PRO_v3"
CLOUD_DB_RAW_URL = "https://raw.githubusercontent.com/contatotal03-ship-it/sistema-licencias/main/licencias_db.enc"
DIAS_GRACIA_OFFLINE = 15

_INFO_LICENCIA_SESION = {}


def get_hwid():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    h_file = os.path.join(base_dir, ".hwid_cache")
    if os.path.exists(h_file):
        try:
            with open(h_file, "r", encoding="utf-8") as f:
                c_val = f.read().strip()
            if c_val and len(c_val) == 64:
                return c_val
        except Exception:
            pass

    uuid_str = None
    try:
        cmd = [
            "powershell",
            "-NoProfile",
            "-Command",
            "(Get-CimInstance -Class Win32_ComputerSystemProduct).UUID",
        ]
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        output = subprocess.check_output(cmd, startupinfo=startupinfo, timeout=8).decode().strip()
        if output and "swbemobject" not in output.lower() and len(output) > 10:
            uuid_str = output
    except Exception:
        pass

    if not uuid_str:
        try:
            cmd = "wmic csproduct get uuid"
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            output = subprocess.check_output(cmd, shell=True, startupinfo=startupinfo, timeout=6).decode().strip()
            lines = [line.strip() for line in output.splitlines() if line.strip()]
            for line in lines:
                if line.lower() != "uuid" and len(line) > 10:
                    uuid_str = line
                    break
        except Exception:
            pass

    if not uuid_str:
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
                uuid_str, _ = winreg.QueryValueEx(key, "MachineGuid")
        except Exception:
            uuid_str = socket.gethostname()

    uuid_str = (uuid_str or "GENERIC_PC_ID").strip()
    hwid_final = hashlib.sha256(uuid_str.encode()).hexdigest()

    try:
        with open(h_file, "w", encoding="utf-8") as f:
            f.write(hwid_final)
    except Exception:
        pass

    return hwid_final


def get_pc_name():
    return os.environ.get("COMPUTERNAME", socket.gethostname() or "PC-Usuario")


def _xor_crypt(data: bytes, key: bytes) -> bytes:
    key_len = len(key)
    return bytes(b ^ key[i % key_len] for i, b in enumerate(data))


def _encrypt_payload(data_dict: dict) -> str:
    raw_json = json.dumps(data_dict, ensure_ascii=False).encode("utf-8")
    sig = hashlib.sha256(raw_json + SECRET_KEY).hexdigest()
    envelope = {
        "sig": sig,
        "payload": base64.b64encode(raw_json).decode("ascii")
    }
    crypted = _xor_crypt(json.dumps(envelope).encode("utf-8"), SECRET_KEY)
    return base64.b64encode(crypted).decode("ascii")


def _decrypt_payload(encrypted_str: str) -> dict:
    crypted = base64.b64decode(encrypted_str.encode("ascii"))
    decrypted_bytes = _xor_crypt(crypted, SECRET_KEY)
    envelope = json.loads(decrypted_bytes.decode("utf-8"))
    raw_json = base64.b64decode(envelope["payload"].encode("ascii"))
    expected_sig = hashlib.sha256(raw_json + SECRET_KEY).hexdigest()
    if envelope.get("sig") != expected_sig:
        raise ValueError("Firma criptográfica inválida.")
    return json.loads(raw_json.decode("utf-8"))


def _get_cache_path():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, ".licencia_camila.enc")


def _guardar_cache_local(info_licencia: dict):
    try:
        info_guardar = dict(info_licencia)
        info_guardar["_hwid_local"] = get_hwid()
        info_guardar["_pc_local"] = get_pc_name()
        info_guardar["_ultima_validacion_local"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        crypted_str = _encrypt_payload(info_guardar)
        with open(_get_cache_path(), "w", encoding="utf-8") as f:
            f.write(crypted_str)
    except Exception as e:
        print(f"⚠️ Error guardando caché local de licencia: {e}")


def _leer_cache_local() -> dict:
    cache_path = _get_cache_path()
    if not os.path.exists(cache_path):
        return {}
    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            crypted_str = f.read().strip()
        if not crypted_str:
            return {}
        info = _decrypt_payload(crypted_str)
        if info.get("_hwid_local") != get_hwid():
            return {}
        return info
    except Exception:
        return {}


def descargar_base_datos_nube(timeout=5) -> dict:
    try:
        url = f"{CLOUD_DB_RAW_URL}?t={int(time.time())}"
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": "CamilaConciliadorLicenseEngine/3.0",
                "Cache-Control": "no-cache"
            }
        )
        ctx = ssl._create_unverified_context() if hasattr(ssl, "_create_unverified_context") else None
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            content = resp.read().decode("utf-8").strip()
            if not content:
                return {}
            return _decrypt_payload(content)
    except Exception:
        return {}


def activar_serial_cloud(serial: str, programa="Camila Conciliador") -> tuple[bool, str, dict]:
    serial = (serial or "").strip().upper()
    if not serial:
        return False, "Por favor ingresa un Serial de activación.", {}

    db_nube = descargar_base_datos_nube(timeout=6)
    if not db_nube:
        cache = _leer_cache_local()
        if cache and cache.get("serial", "").upper() == serial:
            return True, "Activación local cargada correctamente.", cache
        return False, "No fue posible conectar con el servidor de licencias. Verifica tu conexión a internet.", {}

    lic_info = db_nube.get(serial)
    if not lic_info:
        return False, "El Serial ingresado no existe o es inválido.", {}

    prog_lic = (lic_info.get("programa") or "TODOS").upper()
    if prog_lic != "TODOS" and "CAROLINA" not in prog_lic and programa.upper() not in prog_lic:
        return False, f"Este Serial corresponde a '{lic_info.get('programa')}' y no a {programa}.", {}

    estado = (lic_info.get("estado") or "ACTIVO").upper()
    if estado != "ACTIVO":
        return False, f"Esta licencia se encuentra {estado}. Contacta al administrador.", {}

    vence = lic_info.get("vence", "PERMANENTE")
    if vence and vence.upper() != "PERMANENTE":
        try:
            fecha_vence = datetime.strptime(vence, "%Y-%m-%d").date()
            if datetime.now().date() > fecha_vence:
                return False, f"Esta licencia venció el {vence}. Contacta al administrador para renovar.", {}
        except Exception:
            pass

    hwid_actual = get_hwid()
    nombre_pc = get_pc_name()
    equipos = lic_info.get("equipos", [])
    max_equipos = int(lic_info.get("max_equipos", 1))

    hwid_encontrado = any(eq.get("hwid") == hwid_actual for eq in equipos)
    if not hwid_encontrado and len(equipos) >= max_equipos:
        return False, f"Límite de equipos alcanzado ({max_equipos} permitidos).", {}

    info_resultado = {
        "serial": serial,
        "cliente": lic_info.get("cliente", "Usuario Autorizado"),
        "programa": lic_info.get("programa", programa),
        "modulos": lic_info.get("modulos", "TODOS"),
        "vence": vence,
        "estado": "ACTIVO",
        "max_equipos": max_equipos,
        "fecha_activacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    _guardar_cache_local(info_resultado)
    global _INFO_LICENCIA_SESION
    _INFO_LICENCIA_SESION = dict(info_resultado)

    return True, f"¡Licencia activada con éxito para {info_resultado['cliente']}!", info_resultado


def validar_licencia(programa="Camila Conciliador") -> tuple[bool, str]:
    global _INFO_LICENCIA_SESION
    cache = _leer_cache_local()
    hwid_actual = get_hwid()

    if not cache:
        try:
            db_nube_auto = descargar_base_datos_nube(timeout=5)
            if db_nube_auto:
                for s_k, s_v in db_nube_auto.items():
                    if s_k.startswith("_"):
                        continue
                    eqs = s_v.get("equipos", [])
                    if any(eq.get("hwid") == hwid_actual for eq in eqs):
                        p_lic = (s_v.get("programa") or "TODOS").upper()
                        prog_req = programa.upper()
                        if p_lic == "TODOS" or prog_req in p_lic or "CAMILA" in p_lic or "DIANFLOW" in p_lic:
                            estado = (s_v.get("estado") or "ACTIVO").upper()
                            vence = s_v.get("vence", "PERMANENTE")
                            if estado == "ACTIVO":
                                if not vence or vence.upper() == "PERMANENTE" or datetime.now().date() <= datetime.strptime(vence, "%Y-%m-%d").date():
                                    cache = {
                                        "serial": s_k,
                                        "cliente": s_v.get("cliente", "Usuario Autorizado"),
                                        "programa": s_v.get("programa", programa),
                                        "modulos": s_v.get("modulos", "TODOS"),
                                        "vence": vence,
                                        "estado": "ACTIVO",
                                        "max_equipos": s_v.get("max_equipos", 1),
                                        "fecha_activacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    }
                                    _guardar_cache_local(cache)
                                    _INFO_LICENCIA_SESION = dict(cache)
                                    return True, f"Licencia Activa ({cache.get('cliente', 'Usuario')})"
        except Exception:
            pass
        return False, "La aplicación no cuenta con una licencia de activación activa."

    serial = cache.get("serial", "")
    db_nube = descargar_base_datos_nube(timeout=3)
    
    if db_nube and serial in db_nube:
        lic_cloud = db_nube[serial]
        estado = (lic_cloud.get("estado") or "ACTIVO").upper()
        if estado != "ACTIVO":
            cache["estado"] = estado
            _guardar_cache_local(cache)
            return False, f"Licencia {estado}."

        vence = lic_cloud.get("vence", "PERMANENTE")
        if vence and vence.upper() != "PERMANENTE":
            try:
                fecha_vence = datetime.strptime(vence, "%Y-%m-%d").date()
                if datetime.now().date() > fecha_vence:
                    cache["estado"] = "VENCIDO"
                    _guardar_cache_local(cache)
                    return False, f"La licencia ha expirado el {vence}."
            except Exception:
                pass

        cache["cliente"] = lic_cloud.get("cliente", cache.get("cliente"))
        cache["modulos"] = lic_cloud.get("modulos", cache.get("modulos"))
        cache["vence"] = vence
        cache["estado"] = "ACTIVO"
        _guardar_cache_local(cache)
        _INFO_LICENCIA_SESION = dict(cache)
        return True, f"Licencia Activa ({cache.get('cliente', 'Usuario')})"

    if cache:
        if cache.get("estado", "").upper() != "ACTIVO":
            return False, "Licencia suspendida o inactiva."

        vence = cache.get("vence", "PERMANENTE")
        if vence and vence.upper() != "PERMANENTE":
            try:
                if datetime.now().date() > datetime.strptime(vence, "%Y-%m-%d").date():
                    return False, f"La licencia expiró el {vence}."
            except Exception:
                pass

        _INFO_LICENCIA_SESION = dict(cache)
        return True, f"Licencia Activa (Modo Offline)"

    return False, "No se encontró una licencia válida."


def obtener_info_licencia() -> dict:
    global _INFO_LICENCIA_SESION
    if not _INFO_LICENCIA_SESION:
        _INFO_LICENCIA_SESION = _leer_cache_local()
    return _INFO_LICENCIA_SESION
