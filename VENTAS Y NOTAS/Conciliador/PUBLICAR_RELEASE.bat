@echo off
title Publicador Automatico de Releases - Conciliador DIAN
chcp 65001 > nul
cls

set "PYTHON_EXE=python"
where %PYTHON_EXE% >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    if exist "C:\Users\Admin\AppData\Local\Programs\Python\Python313\python.exe" (
        set "PYTHON_EXE=C:\Users\Admin\AppData\Local\Programs\Python\Python313\python.exe"
    )
)

"%PYTHON_EXE%" publicar_release.py
pause
