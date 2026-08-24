@echo off
echo ==============================================
echo  CREANDO EJECUTABLE (.EXE) CON PYINSTALLER
echo ==============================================
echo.
echo Limpiando versiones anteriores...
rmdir /s /q build
rmdir /s /q dist
del /q *.spec

echo.
echo Compilando la aplicacion...
python -m PyInstaller --noconfirm --onefile --windowed --icon="ui/logo.ico" --name "Conciliador_DIAN_vs_Sistema" --add-data "ui/styles.qss;ui/" --add-data "ui/logo.png;ui/" main.py

echo.
echo ==============================================
echo  PROCESO FINALIZADO
echo  Moviendo el ejecutable a la carpeta principal...
echo ==============================================
move /y "dist\Conciliador_DIAN_vs_Sistema.exe" "."
rmdir /s /q build
rmdir /s /q dist
del /q *.spec
echo Listo!
