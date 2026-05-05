@echo off
setlocal
set "PROJECT_DIR=%~dp0"
set "LOCAL_PYTHONW=%PROJECT_DIR%runtime\python\pythonw.exe"
set "PYTHONW=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
set "APP_PATH=%PROJECT_DIR%app.py"

for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":8000" ^| find "LISTENING"') do taskkill /PID %%p /F >nul 2>&1

if exist "%LOCAL_PYTHONW%" (
    start "" "%LOCAL_PYTHONW%" "%APP_PATH%"
) else if exist "%PYTHONW%" (
    start "" "%PYTHONW%" "%APP_PATH%"
) else (
    start "" pythonw "%APP_PATH%"
)

timeout /t 3 /nobreak >nul
start "" "http://oratoriocarloacutis.don:8000/login"
endlocal
