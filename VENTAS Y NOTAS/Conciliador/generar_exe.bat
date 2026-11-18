@echo off
title Compilador y Empaquetador Release - Conciliador DIAN
chcp 65001 > nul
cls

echo ===============================================================================
echo            COMPILADOR Y EMPAQUETADOR CONCILIADOR RELEASE (.ZIP / .EXE)
echo ===============================================================================
echo.

rem Cerrar procesos previos
taskkill /F /IM Conciliador_DIAN_vs_Sistema.exe > nul 2>&1

echo [*] Verificando entorno de Python...
set "PYTHON_EXE=python"
where %PYTHON_EXE% >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    if exist "C:\Users\Admin\AppData\Local\Programs\Python\Python313\python.exe" (
        set "PYTHON_EXE=C:\Users\Admin\AppData\Local\Programs\Python\Python313\python.exe"
    )
)

echo [*] Usando Python: %PYTHON_EXE%
echo.

echo [*] 1. Limpiando compilaciones anteriores (build / dist / Release)...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "Conciliador_Release" rmdir /s /q "Conciliador_Release"

echo.
echo [*] 2. Compilando Conciliador_DIAN_vs_Sistema.exe con PyInstaller...
"%PYTHON_EXE%" -m PyInstaller --noconfirm --clean --onefile --windowed --icon="ui/logo.ico" --name "Conciliador_DIAN_vs_Sistema" --add-data "ui/styles.qss;ui/" --add-data "ui/logo.png;ui/" --add-data "version.json;." --add-data "vendedores_config.json;." --collect-all certifi main.py

if not exist "dist\Conciliador_DIAN_vs_Sistema.exe" (
    echo.
    echo ===============================================================================
    echo    [ERROR] Ocurrio un problema al compilar el ejecutable.
    echo ===============================================================================
    pause
    exit /b 1
)

echo.
echo [*] 3. Sincronizando carpeta de distribucion Conciliador_Release...
if not exist "Conciliador_Release" mkdir "Conciliador_Release"

copy /y "dist\Conciliador_DIAN_vs_Sistema.exe" "Conciliador_Release\Conciliador_DIAN_vs_Sistema.exe" > nul
copy /y "dist\Conciliador_DIAN_vs_Sistema.exe" ".\Conciliador_DIAN_vs_Sistema.exe" > nul
if exist "version.json" copy /y "version.json" "Conciliador_Release\" > nul
if exist "configuracion_conciliador.json" copy /y "configuracion_conciliador.json" "Conciliador_Release\" > nul

echo.
echo ===============================================================================
echo   [OK] Carpeta Release lista en: %CD%\Conciliador_Release
echo ===============================================================================
echo.

echo [*] 4. Generando archivo ZIP para GitHub Releases y distribucion...

rem Leer version actual desde version.json
for /f "delims=" %%V in ('"%PYTHON_EXE%" get_version.py') do set APP_VERSION=%%V

set ZIP_NAME=Conciliador_DIAN_v%APP_VERSION%.zip

rem Eliminar ZIP previo con el mismo nombre si existe
if exist "%ZIP_NAME%" del /f /q "%ZIP_NAME%"

powershell -NoProfile -Command "Compress-Archive -Path 'Conciliador_Release\*' -DestinationPath '%ZIP_NAME%' -Force"

if exist "%ZIP_NAME%" (
    echo.
    echo ===============================================================================
    echo   [OK] ZIP GENERADO CON EXITO: %CD%\%ZIP_NAME%
    echo.
    echo   PASOS PARA PUBLICAR MANUALMENTE EN GITHUB RELEASES:
    echo   1. Ir a: https://github.com/contatotal03-ship-it/camila/releases/new
    echo   2. Tag / Etiqueta: v%APP_VERSION%
    echo   3. Titulo: Conciliador DIAN v%APP_VERSION%
    echo   4. Adjuntar el archivo ZIP: %ZIP_NAME% y/o Conciliador_DIAN_vs_Sistema.exe
    echo   5. Publicar el release
    echo ===============================================================================
) else (
    echo   [ADVERTENCIA] No se pudo generar el archivo ZIP comprimido.
)

echo.
set /p pub="Deseas que el sistema publique automaticamente el ZIP y EXE a GitHub ahora? (S/N): "
if /i "%pub%"=="S" (
    "%PYTHON_EXE%" publicar_release.py
) else (
    explorer .
)

echo.
pause
