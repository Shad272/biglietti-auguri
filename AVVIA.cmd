@echo off
cd /d "%~dp0"
if not exist "%~dp0runtime\pythonw.exe" (
  echo File runtime mancanti. Estrai interamente lo ZIP in una cartella e riprova.
  pause
  exit /b 1
)
if not exist "%~dp0biglietti.py" (
  echo File del programma mancanti. Estrai interamente lo ZIP e riprova.
  pause
  exit /b 1
)
start "" "%~dp0runtime\pythonw.exe" "%~dp0biglietti.py"
