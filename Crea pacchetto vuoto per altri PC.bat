@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\crea_pacchetto_distribuzione.ps1" -EmptyDatabase
if errorlevel 1 (
  echo.
  echo Creazione pacchetto vuoto non completata.
  pause
  exit /b 1
)
echo.
echo Pacchetto vuoto creato correttamente.
pause
endlocal
