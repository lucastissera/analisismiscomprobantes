@echo off
setlocal
cd /d "%~dp0"
echo Instalando PyInstaller si hace falta...
python -m pip install --upgrade pip pyinstaller
python -m pip install -r requirements.txt
echo Compilando portable y sincronizando claves...
python tools\portable_build.py
if errorlevel 1 exit /b 1
echo.
echo Listo. Ejecutable y archivos en:
echo   %~dp0dist\MisComprobantesAnalisis\
echo Ejecutable principal:
echo   %~dp0dist\MisComprobantesAnalisis\MisComprobantesAnalisis.exe
if exist "%~dp0auth_users.json" (
  echo auth_users.json copiado junto al .exe.
) else (
  echo Aviso: no hay auth_users.json en la raiz; el portable usara el ejemplo empaquetado.
)
endlocal
