@echo off
setlocal
cd /d "%~dp0"
python -m pip install -q watchdog pyinstaller 2>nul
python -m pip install -q -r requirements.txt 2>nul
python tools\portable_watch.py %*
endlocal
