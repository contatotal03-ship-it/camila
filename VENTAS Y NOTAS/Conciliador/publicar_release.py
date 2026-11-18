import os
import sys
import json
import ssl
import urllib.request
import urllib.error
import subprocess
import shutil
from pathlib import Path

# Resiliencia SSL para Windows
try:
    if hasattr(ssl, "_create_unverified_context"):
        ssl._create_default_https_context = ssl._create_unverified_context
except Exception:
    pass

GITHUB_REPO = "contatotal03-ship-it/camila"
AUTH_TOKEN = "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"
BASE_DIR = Path(__file__).parent.resolve()

def get_headers(accept="application/vnd.github+json", content_type=None):
    headers = {
        "User-Agent": "Camila-ReleasePublisher",
        "Accept": accept,
        "Authorization": f"Bearer {AUTH_TOKEN}"
    }
    if content_type:
        headers["Content-Type"] = content_type
    return headers

def read_version_json():
    v_file = BASE_DIR / "version.json"
    if v_file.exists():
        with open(v_file, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "version": "1.1.0",
        "release_notes": "Mejoras y correcciones generales.",
        "download_url": ""
    }

def save_version_json(data):
    v_file = BASE_DIR / "version.json"
    with open(v_file, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def build_exe_and_zip(version):
    print("\n========================================================")
    print("  [1/4] COMPILANDO EJECUTABLE Y GENERANDO PAQUETE ZIP")
    print("========================================================")
    
    subprocess.run(["taskkill", "/F", "/IM", "Conciliador_DIAN_vs_Sistema.exe"], capture_output=True)
    
    build_dir = BASE_DIR / "build"
    dist_dir = BASE_DIR / "dist"
    rel_dir = BASE_DIR / "Conciliador_Release"
    zip_path = BASE_DIR / f"Conciliador_DIAN_v{version}.zip"
    
    if build_dir.exists():
        shutil.rmtree(build_dir, ignore_errors=True)
    if dist_dir.exists():
        shutil.rmtree(dist_dir, ignore_errors=True)
    if rel_dir.exists():
        shutil.rmtree(rel_dir, ignore_errors=True)
    if zip_path.exists():
        try:
            zip_path.unlink()
        except Exception:
            pass
        
    py_dir = Path(sys.executable).parent
    sys32 = Path("C:/Windows/System32")
    vc_redist = BASE_DIR / "VC_redist.x64.exe"

    candidate_dlls = [
        py_dir / "vcruntime140.dll",
        py_dir / "vcruntime140_1.dll",
        py_dir / "python313.dll",
        py_dir / "python3.dll",
        sys32 / "msvcp140.dll",
        sys32 / "msvcp140_1.dll",
        sys32 / "msvcp140_2.dll",
        sys32 / "msvcp140_atomic_wait.dll",
        sys32 / "msvcp140_codecvt_ids.dll",
        sys32 / "vccorlib140.dll",
    ]
    existing_dlls = [d for d in candidate_dlls if d.exists()]

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--noupx",
    ]
    for d in existing_dlls:
        cmd.append(f"--add-binary={d};.")

    cmd.extend([
        '--icon=ui/logo.ico' if (BASE_DIR / 'ui' / 'logo.ico').exists() else '--name=Conciliador_DIAN_vs_Sistema',
        "--name", "Conciliador_DIAN_vs_Sistema",
        "--add-data", "ui/styles.qss;ui/",
        "--add-data", "ui/logo.png;ui/",
        "--add-data", "version.json;.",
        "--add-data", "vendedores_config.json;.",
        "--collect-all", "certifi",
        "main.py"
    ])
    
    print(f"Ejecutando PyInstaller...")
    res = subprocess.run(cmd, cwd=str(BASE_DIR))
    
    exe_path = dist_dir / "Conciliador_DIAN_vs_Sistema.exe"
    if res.returncode != 0 or not exe_path.exists():
        print("\n[ERROR] La compilacion de Conciliador_DIAN_vs_Sistema.exe fallo.")
        return None, None
        
    size_mb = os.path.getsize(exe_path) / (1024 * 1024)
    print(f"\n[OK] Ejecutable generado: {exe_path} ({size_mb:.2f} MB)")
    
    # Sincronizar carpeta Release
    rel_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(exe_path, rel_dir / "Conciliador_DIAN_vs_Sistema.exe")
    
    for opt_file in ["version.json", "configuracion_conciliador.json", "vendedores_config.json"]:
        src = BASE_DIR / opt_file
        if src.exists():
            shutil.copy2(src, rel_dir / opt_file)

    # Copiar DLLs de runtime y dependencias de Windows en el paquete
    for dll_f in existing_dlls:
        shutil.copy2(dll_f, rel_dir / dll_f.name)

    if vc_redist.exists():
        shutil.copy2(vc_redist, rel_dir / vc_redist.name)

    # Crear script de instalación de dependencias VC++
    bat_vc = rel_dir / "INSTALAR_COMPONENTES_WINDOWS.bat"
    bat_vc.write_text("""@echo off
chcp 65001 > nul
echo ============================================================
echo   INSTALADOR DE COMPONENTES MICROSOFT VISUAL C++ (x64)
echo   Camila Conciliador Financiero
echo ============================================================
echo Instalando librerias de runtime para Windows...
"%~dp0VC_redist.x64.exe" /passive /norestart
echo [OK] Componentes listos. Ya puedes iniciar Conciliador_DIAN_vs_Sistema.exe
pause
""", encoding="utf-8")
            
    # Comprimir ZIP
    print(f"Generando archivo ZIP comprimido: {zip_path.name}...")
    import zipfile
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(rel_dir):
            for file in files:
                full_p = Path(root) / file
                rel_p = full_p.relative_to(rel_dir)
                zf.write(full_p, rel_p)
                
    zip_size_mb = os.path.getsize(zip_path) / (1024 * 1024)
    print(f"[OK] Archivo ZIP generado exitosamente: {zip_path} ({zip_size_mb:.2f} MB)")
    
    return exe_path, zip_path

def get_or_create_release(version, notes):
    print("\n========================================================")
    print(f"  [2/4] VERIFICANDO / CREANDO RELEASE EN GITHUB (v{version})")
    print("========================================================")
    
    tag = f"v{version.lstrip('v')}"
    
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/tags/{tag}"
        req = urllib.request.Request(url, headers=get_headers())
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            print(f"[INFO] El release {tag} ya existe en GitHub (ID: {data.get('id')}).")
            return data
    except urllib.error.HTTPError as e:
        if e.code != 404:
            print(f"[WARN] Error verificando release: {e}")
            
    print(f"Creando nuevo release {tag} en GitHub...")
    create_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases"
    payload = {
        "tag_name": tag,
        "target_commitish": "main",
        "name": f"Conciliador DIAN vs Sistema {tag}",
        "body": notes,
        "draft": False,
        "prerelease": False
    }
    
    req = urllib.request.Request(
        create_url,
        data=json.dumps(payload).encode("utf-8"),
        headers=get_headers(content_type="application/json"),
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        data = json.loads(resp.read().decode("utf-8"))
        print(f"[OK] Release {tag} creado exitosamente en GitHub (ID: {data.get('id')})")
        return data

def upload_single_asset(release_data, file_path, asset_name):
    for asset in release_data.get("assets", []):
        if asset.get("name") == asset_name:
            print(f"Eliminando asset existente '{asset_name}' (ID: {asset['id']})...")
            del_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/assets/{asset['id']}"
            del_req = urllib.request.Request(del_url, headers=get_headers(), method="DELETE")
            try:
                urllib.request.urlopen(del_req, timeout=10)
            except Exception as e:
                print(f"[WARN] No se pudo eliminar asset previo: {e}")
                
    upload_url = release_data["upload_url"].split("{")[0]
    upload_url = f"{upload_url}?name={asset_name}"
    
    file_size = os.path.getsize(file_path)
    print(f"Subiendo {asset_name} ({file_size / (1024*1024):.2f} MB) a GitHub Releases...")
    
    content_type = "application/zip" if asset_name.endswith(".zip") else "application/octet-stream"
    
    with open(file_path, "rb") as f:
        file_bytes = f.read()
        
    req = urllib.request.Request(
        upload_url,
        data=file_bytes,
        headers=get_headers(
            accept="application/vnd.github+json",
            content_type=content_type
        ),
        method="POST"
    )
    
    with urllib.request.urlopen(req, timeout=180) as resp:
        asset_info = json.loads(resp.read().decode("utf-8"))
        print(f"[OK] {asset_name} subido exitosamente!")
        return asset_info

def upload_release_assets(release_data, exe_path, zip_path):
    print("\n========================================================")
    print("  [3/4] SUBIENDO ARCHIVOS (ZIP Y EXE) A GITHUB RELEASES")
    print("========================================================")
    
    if zip_path and zip_path.exists():
        upload_single_asset(release_data, zip_path, zip_path.name)
        
    if exe_path and exe_path.exists():
        upload_single_asset(release_data, exe_path, "Conciliador_DIAN_vs_Sistema.exe")

def sync_git_and_version(version, notes):
    print("\n========================================================")
    print("  [4/4] SINCRONIZANDO VERSION.JSON CON GITHUB")
    print("========================================================")
    
    v_data = {
        "version": version,
        "release_notes": notes,
        "download_url": f"https://github.com/{GITHUB_REPO}/releases/latest/download/Conciliador_DIAN_vs_Sistema.exe"
    }
    save_version_json(v_data)
    
    repo_root = BASE_DIR.parent.parent
    subprocess.run(["git", "add", "."], cwd=str(repo_root))
    subprocess.run(["git", "commit", "-m", f"Release v{version} - {notes.splitlines()[0] if notes else 'Actualizacion'}"], cwd=str(repo_root))
    subprocess.run(["git", "push", "origin", "main"], cwd=str(repo_root))
    print("[OK] version.json y cambios sincronizados en GitHub.")

def main():
    print("================================================================")
    print("       PUBLICADOR DE RELEASES Y AUTO-UPDATER - CAMILA CONCILIADOR ")
    print("================================================================")
    
    current_data = read_version_json()
    curr_ver = current_data.get("version", "1.1.0")
    curr_notes = current_data.get("release_notes", "")
    
    print(f"\nVersión configurada actualmente: v{curr_ver}")
    if len(sys.argv) > 1 and sys.argv[1].strip():
        new_ver = sys.argv[1].strip().lstrip("v")
    else:
        new_ver = input(f"Ingresa la nueva versión (Enter para mantener {curr_ver}): ").strip()
        if not new_ver:
            new_ver = curr_ver
        new_ver = new_ver.lstrip("v")

    if len(sys.argv) > 2 and sys.argv[2].strip():
        new_notes = sys.argv[2].strip()
    else:
        print(f"\nNotas de versión actuales:\n{curr_notes}\n")
        new_notes = input("Escribe nuevas notas (Enter para mantener las anteriores): ").strip()
        if not new_notes:
            new_notes = curr_notes

    # Guardar version.json ANTES de compilar para que PyInstaller empaquete la versión correcta
    v_data = {
        "version": new_ver,
        "release_notes": new_notes,
        "download_url": f"https://github.com/{GITHUB_REPO}/releases/latest/download/Conciliador_DIAN_vs_Sistema.exe"
    }
    save_version_json(v_data)
            
    exe_path, zip_path = build_exe_and_zip(new_ver)
    if not exe_path:
        sys.exit(1)
        
    release_data = get_or_create_release(new_ver, new_notes)
    upload_release_assets(release_data, exe_path, zip_path)
    sync_git_and_version(new_ver, new_notes)
    
    print("\n================================================================")
    print(f"  [OK] PROCESO COMPLETADO! Conciliador v{new_ver} está publicado.")
    print(f"  Archivo ZIP: {zip_path.name}")
    print("  Los usuarios del EXE recibirán la actualización automáticamente.")
    print("================================================================\n")

if __name__ == "__main__":
    main()
