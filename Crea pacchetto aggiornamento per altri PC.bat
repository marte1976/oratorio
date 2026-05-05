@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\crea_pacchetto_distribuzione.ps1" -UpdateOnly -CreateZip
if errorlevel 1 (
  echo.
  echo Creazione pacchetto aggiornamento non completata.
  pause
  exit /b 1
)
echo.
echo Pacchetto aggiornamento ZIP creato correttamente.
pause
endlocal
