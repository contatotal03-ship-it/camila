# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('ui/styles.qss', 'ui/'), ('ui/logo.png', 'ui/'), ('version.json', '.'), ('vendedores_config.json', '.')]
binaries = [('C:\\Users\\Admin\\AppData\\Local\\Programs\\Python\\Python313\\vcruntime140.dll', '.'), ('C:\\Users\\Admin\\AppData\\Local\\Programs\\Python\\Python313\\vcruntime140_1.dll', '.'), ('C:\\Users\\Admin\\AppData\\Local\\Programs\\Python\\Python313\\python313.dll', '.'), ('C:\\Users\\Admin\\AppData\\Local\\Programs\\Python\\Python313\\python3.dll', '.'), ('C:\\Windows\\System32\\msvcp140.dll', '.'), ('C:\\Windows\\System32\\msvcp140_1.dll', '.'), ('C:\\Windows\\System32\\msvcp140_2.dll', '.'), ('C:\\Windows\\System32\\msvcp140_atomic_wait.dll', '.'), ('C:\\Windows\\System32\\msvcp140_codecvt_ids.dll', '.'), ('C:\\Windows\\System32\\vccorlib140.dll', '.')]
hiddenimports = []
tmp_ret = collect_all('certifi')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Conciliador_DIAN_vs_Sistema',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['ui\\logo.ico'],
)
