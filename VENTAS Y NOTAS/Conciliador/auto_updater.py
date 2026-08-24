import os
import sys
import json
import urllib.request
import urllib.error
import threading
import subprocess
import time
import zipfile
import shutil
import tempfile
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QProgressBar, QTextEdit, QMessageBox
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QObject
from PyQt6.QtGui import QFont, QIcon

APP_NAME = "Conciliador_DIAN_vs_Sistema"
GITHUB_REPO = "contatotal03-ship-it/camila"
_AUTH_TOKEN = "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"
DEFAULT_VERSION = "1.1.0"

def _get_headers(accept="application/vnd.github+json"):
    headers = {
        "User-Agent": "Camila-AutoUpdater",
        "Accept": accept
    }
    if _AUTH_TOKEN:
        headers["Authorization"] = f"Bearer {_AUTH_TOKEN}"
    return headers

def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

def get_executable_path():
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    return os.path.abspath(__file__)

def get_local_version_info():
    candidates = [
        os.path.join(get_app_dir(), "version.json"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "version.json")
    ]
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidates.insert(0, os.path.join(sys._MEIPASS, "version.json"))
        
    for p in candidates:
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
    return {}

def get_local_version():
    info = get_local_version_info()
    return info.get("version", DEFAULT_VERSION)

def parse_version_tuple(ver_str):
    clean = str(ver_str).strip().lower().lstrip("v")
    parts = []
    for p in clean.split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])

def is_newer_version(remote_ver, local_ver):
    try:
        return parse_version_tuple(remote_ver) > parse_version_tuple(local_ver)
    except Exception:
        return str(remote_ver).strip() != str(local_ver).strip()

def get_remote_version_info():
    # 1. Intentar consultar GitHub Releases API
    try:
        rel_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        req = urllib.request.Request(rel_url, headers=_get_headers())
        with urllib.request.urlopen(req, timeout=8) as resp:
            if resp.status == 200:
                data = json.loads(resp.read().decode("utf-8"))
                tag = data.get("tag_name", "").lstrip("v")
                body = data.get("body", "Mejoras y optimizaciones generales.")
                assets = data.get("assets", [])

                target_asset = None
                for asset in assets:
                    name = asset.get("name", "").lower()
                    if name.endswith(".zip"):
                        target_asset = asset
                        break
                if not target_asset:
                    for asset in assets:
                        name = asset.get("name", "").lower()
                        if name.endswith(".exe"):
                            target_asset = asset
                            break

                if target_asset:
                    return {
                        "version": tag,
                        "release_notes": body,
                        "download_url": target_asset.get("url"),
                        "asset_name": target_asset.get("name", "Conciliador_DIAN_vs_Sistema.exe"),
                        "size": target_asset.get("size", 0),
                        "is_api_asset": True
                    }
    except Exception:
        pass

    # 2. Fallback: Consultar version.json
    try:
        contents_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/VENTAS%20Y%20NOTAS/Conciliador/version.json"
        req = urllib.request.Request(contents_url, headers=_get_headers(accept="application/vnd.github.raw"))
        with urllib.request.urlopen(req, timeout=8) as resp:
            if resp.status == 200:
                data = json.loads(resp.read().decode("utf-8"))
                return {
                    "version": data.get("version", ""),
                    "release_notes": data.get("release_notes", "Actualización del sistema."),
                    "download_url": data.get("download_url", ""),
                    "asset_name": "Conciliador_DIAN_vs_Sistema.exe",
                    "size": 0,
                    "is_api_asset": False
                }
    except Exception as e:
        raise Exception(f"No se pudo consultar el servidor de actualizaciones ({e})")

    return None

class UpdateSignals(QObject):
    found_update = pyqtSignal(dict)
    no_update = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(float, int, int)
    finished = pyqtSignal(str)

_SIGNALS = None

def get_signals():
    global _SIGNALS
    if _SIGNALS is None:
        _SIGNALS = UpdateSignals()
    return _SIGNALS

def check_for_updates_threaded(parent_widget=None, silent=True):
    signals = get_signals()
    
    try:
        signals.found_update.disconnect()
    except Exception:
        pass
    try:
        signals.no_update.disconnect()
    except Exception:
        pass
    try:
        signals.error.disconnect()
    except Exception:
        pass

    signals.found_update.connect(lambda info: _prompt_update(parent_widget, info))
    if not silent:
        signals.no_update.connect(lambda ver: _notify_no_updates(parent_widget, ver))
        signals.error.connect(lambda err: QMessageBox.warning(parent_widget, "Actualizaciones", f"No se pudo verificar actualizaciones:\n{err}"))

    def _worker():
        local_ver = get_local_version()
        try:
            remote_info = get_remote_version_info()
            if not remote_info:
                signals.no_update.emit(local_ver)
                return
            if is_newer_version(remote_info.get("version", ""), local_ver):
                signals.found_update.emit(remote_info)
            else:
                signals.no_update.emit(local_ver)
        except Exception as e:
            signals.error.emit(str(e))

    threading.Thread(target=_worker, daemon=True).start()

def _notify_no_updates(parent_widget, local_ver):
    QMessageBox.information(
        parent_widget,
        "Sistema Actualizado",
        f"¡Tienes la versión más reciente del Conciliador (v{local_ver})!\n\nNo hay nuevas actualizaciones pendientes."
    )

def _prompt_update(parent_widget, remote_info):
    remote_ver = remote_info.get("version")
    local_ver = get_local_version()
    notes = remote_info.get("release_notes", "")

    dialog = QDialog(parent_widget)
    dialog.setWindowTitle("🔄 Actualización Disponible - Conciliador DIAN")
    dialog.setFixedSize(500, 380)
    dialog.setStyleSheet("""
        QDialog { background-color: #0d1821; color: #ffffff; }
        QLabel { color: #ffffff; }
        QTextEdit { background-color: #12212e; color: #e0e6ed; border: 1px solid #1e3a5f; border-radius: 6px; padding: 8px; }
        QPushButton#BtnUpdate { background-color: #00d4aa; color: #0a141d; font-weight: bold; font-size: 13px; border-radius: 6px; padding: 10px 18px; }
        QPushButton#BtnUpdate:hover { background-color: #00b894; }
        QPushButton#BtnLater { background-color: #192b3c; color: #94a3b8; font-size: 13px; border-radius: 6px; padding: 10px 18px; }
        QPushButton#BtnLater:hover { background-color: #233b52; color: #ffffff; }
    """)

    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(24, 20, 24, 20)
    layout.setSpacing(12)

    lbl_title = QLabel("🎉 ¡Nueva Versión Disponible!")
    lbl_title.setFont(QFont("Segoe UI", 15, QFont.Weight.Bold))
    lbl_title.setStyleSheet("color: #00d4aa;")
    layout.addWidget(lbl_title)

    lbl_ver = QLabel(f"Versión instalada: v{local_ver}   ➔   Nueva versión: v{remote_ver}")
    lbl_ver.setStyleSheet("color: #7faabb; font-weight: 600;")
    layout.addWidget(lbl_ver)

    lbl_notes = QLabel("Novedades:")
    layout.addWidget(lbl_notes)

    txt_notes = QTextEdit()
    txt_notes.setReadOnly(True)
    txt_notes.setText(notes)
    layout.addWidget(txt_notes)

    btn_layout = QHBoxLayout()
    btn_update = QPushButton("⬇️ Descargar e Instalar Ahora")
    btn_update.setObjectName("BtnUpdate")
    btn_later = QPushButton("Más Tarde")
    btn_later.setObjectName("BtnLater")

    btn_layout.addWidget(btn_update)
    btn_layout.addWidget(btn_later)
    layout.addLayout(btn_layout)

    def on_update():
        dialog.accept()
        _perform_update(parent_widget, remote_info)

    btn_update.clicked.connect(on_update)
    btn_later.clicked.connect(dialog.reject)

    dialog.exec()

def _perform_update(parent_widget, remote_info):
    if not getattr(sys, "frozen", False):
        QMessageBox.information(
            parent_widget,
            "Modo Desarrollo",
            "Estás ejecutando el código fuente (.py).\nPara actualizar ejecuta 'ACTUALIZAR_DESDE_GITHUB.bat'."
        )
        return

    download_url = remote_info.get("download_url")
    if not download_url:
        QMessageBox.critical(parent_widget, "Error", "No se encontró enlace de descarga.")
        return

    progress_dialog = QDialog(parent_widget)
    progress_dialog.setWindowTitle("Descargando Actualización...")
    progress_dialog.setFixedSize(420, 160)
    progress_dialog.setStyleSheet("""
        QDialog { background-color: #0d1821; color: #ffffff; }
        QLabel { color: #ffffff; }
        QProgressBar { background-color: #12212e; border: 1px solid #1e3a5f; border-radius: 5px; text-align: center; color: #ffffff; font-weight: bold; }
        QProgressBar::chunk { background-color: #00d4aa; border-radius: 4px; }
    """)

    p_layout = QVBoxLayout(progress_dialog)
    p_layout.setContentsMargins(24, 20, 24, 20)

    lbl_prog = QLabel(f"Descargando v{remote_info.get('version')}...")
    lbl_prog.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
    p_layout.addWidget(lbl_prog)

    bar = QProgressBar()
    bar.setRange(0, 100)
    bar.setValue(0)
    p_layout.addWidget(bar)

    lbl_det = QLabel("Conectando con el servidor...")
    lbl_det.setStyleSheet("color: #7faabb; font-size: 11px;")
    p_layout.addWidget(lbl_det)

    signals = get_signals()
    try:
        signals.progress.disconnect()
    except Exception:
        pass
    try:
        signals.finished.disconnect()
    except Exception:
        pass

    def on_progress(pct, downloaded, total):
        bar.setValue(int(pct * 100))
        mb_d = downloaded / (1024 * 1024)
        mb_t = total / (1024 * 1024)
        lbl_det.setText(f"{pct*100:.1f}% ({mb_d:.1f} MB de {mb_t:.1f} MB)")

    def on_finished(bat_path):
        progress_dialog.accept()
        try:
            subprocess.Popen(bat_path, shell=True, creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
            sys.exit(0)
        except Exception:
            sys.exit(0)

    signals.progress.connect(on_progress)
    signals.finished.connect(on_finished)

    def _download_thread():
        current_exe = get_executable_path()
        app_dir = os.path.dirname(current_exe)
        asset_name = remote_info.get("asset_name", "Conciliador_DIAN_vs_Sistema.exe")
        is_zip = asset_name.lower().endswith(".zip")

        temp_dir = tempfile.mkdtemp(prefix="camila_update_")
        temp_download_file = os.path.join(temp_dir, asset_name)
        updater_bat = os.path.join(app_dir, "updater_replace.bat")

        try:
            accept_hdr = "application/octet-stream" if ("api.github.com" in download_url or remote_info.get("is_api_asset")) else "*/*"
            req = urllib.request.Request(download_url, headers=_get_headers(accept=accept_hdr))

            with urllib.request.urlopen(req, timeout=90) as response:
                total_size = int(response.info().get("Content-Length", 0))
                if total_size == 0 and remote_info.get("size"):
                    total_size = int(remote_info.get("size"))
                
                downloaded = 0
                block_size = 65536

                with open(temp_download_file, "wb") as f_out:
                    while True:
                        buffer = response.read(block_size)
                        if not buffer:
                            break
                        f_out.write(buffer)
                        downloaded += len(buffer)
                        if total_size > 0:
                            signals.progress.emit(downloaded / total_size, downloaded, total_size)

            if not os.path.exists(temp_download_file) or os.path.getsize(temp_download_file) < 50000:
                raise Exception("El archivo descargado está incompleto.")

            staging_dir = os.path.join(temp_dir, "extracted")
            if is_zip:
                os.makedirs(staging_dir, exist_ok=True)
                with zipfile.ZipFile(temp_download_file, "r") as zf:
                    zf.extractall(staging_dir)
                
                bat_content = f"""@echo off
chcp 65001 > nul
echo Actualizando Conciliador...
timeout /t 1 /nobreak > nul
:loop
taskkill /F /IM "Conciliador_DIAN_vs_Sistema.exe" > nul 2>&1
timeout /t 1 /nobreak > nul
del "{current_exe}" > nul 2>&1
if exist "{current_exe}" (
    goto loop
)
xcopy /s /e /y /q "{staging_dir}\\*" "{app_dir}\\" > nul
start "" "{current_exe}"
rd /s /q "{temp_dir}" > nul 2>&1
del "%~f0" > nul
"""
            else:
                bat_content = f"""@echo off
chcp 65001 > nul
echo Actualizando Conciliador...
timeout /t 1 /nobreak > nul
:loop
taskkill /F /IM "Conciliador_DIAN_vs_Sistema.exe" > nul 2>&1
timeout /t 1 /nobreak > nul
del "{current_exe}" > nul 2>&1
if exist "{current_exe}" (
    goto loop
)
move /y "{temp_download_file}" "{current_exe}" > nul
start "" "{current_exe}"
rd /s /q "{temp_dir}" > nul 2>&1
del "%~f0" > nul
"""

            with open(updater_bat, "w", encoding="utf-8") as f_bat:
                f_bat.write(bat_content)

            signals.finished.emit(updater_bat)
        except Exception as ex:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            signals.error.emit(str(ex))

    threading.Thread(target=_download_thread, daemon=True).start()
    progress_dialog.exec()
