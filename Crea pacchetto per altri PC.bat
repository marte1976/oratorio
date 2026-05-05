@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\crea_pacchetto_distribuzione.ps1"
if errorlevel 1 (
  echo.
  echo Creazione pacchetto non completata.
  pause
  exit /b 1
)
echo.
echo Pacchetto creato correttamente.
pause
endlocal
