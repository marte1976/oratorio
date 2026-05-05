param(
    [switch]$NoBrowser
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$healthcheckUrl = "http://127.0.0.1:8000/login"
$browserUrl = "http://oratoriocarloacutis.don:8000/login"
$appPath = Join-Path $projectRoot "app.py"
$localPythonw = Join-Path $projectRoot "runtime\python\pythonw.exe"
$bundledPythonw = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
$logDir = Join-Path $projectRoot "outputs"
$logPath = Join-Path $logDir "desktop-launcher.log"

function Write-LauncherLog {
    param([string]$Message)
    try {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -LiteralPath $logPath -Value "[$timestamp] $Message"
    } catch {
    }
}

function Test-ServerReady {
    try {
        $response = Invoke-WebRequest -UseBasicParsing -Uri $healthcheckUrl -TimeoutSec 4
        return ($response.StatusCode -ge 200 -and $response.StatusCode -lt 500)
    } catch {
        return $false
    }
}

function Resolve-Pythonw {
    if (Test-Path -LiteralPath $localPythonw) {
        return $localPythonw
    }
    if (Test-Path -LiteralPath $bundledPythonw) {
        return $bundledPythonw
    }
    return "pythonw"
}

if (-not (Test-ServerReady)) {
    $pythonw = Resolve-Pythonw
    Write-LauncherLog "Server non attivo. Avvio con $pythonw"
    Start-Process -FilePath $pythonw -ArgumentList ('"' + $appPath + '"') -WorkingDirectory $projectRoot -WindowStyle Hidden

    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline) {
        if (Test-ServerReady) {
            break
        }
        Start-Sleep -Milliseconds 500
    }
}

if ($NoBrowser) {
    if (Test-ServerReady) {
        Write-LauncherLog "Server pronto. Avvio in background senza browser."
    } else {
        Write-LauncherLog "Server non raggiungibile dopo l'avvio in background."
    }
    exit
}

if (Test-ServerReady) {
    Write-LauncherLog "Server pronto. Apro il browser."
} else {
    Write-LauncherLog "Server non raggiungibile, provo comunque ad aprire il browser."
}

Start-Process -FilePath "explorer.exe" -ArgumentList ('"' + $browserUrl + '"')
