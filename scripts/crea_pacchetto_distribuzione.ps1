param(
    [string]$OutputRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) "dist"),
    [string]$DatabaseSourcePath = "",
    [switch]$CreateZip,
    [switch]$EmptyDatabase,
    [switch]$UpdateOnly
)

$ErrorActionPreference = "Stop"
$OutputRoot = [System.IO.Path]::GetFullPath($OutputRoot)

$projectRoot = Split-Path -Parent $PSScriptRoot
$versionFilePath = Join-Path $projectRoot "VERSION.txt"
$installedDatabasePath = "C:\OratorioCarloAcutis\database\gestione_associazione.sqlite"
$fallbackDatabasePath = Join-Path $projectRoot "database\gestione_associazione.sqlite"
$pythonRuntimeSource = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python"
$nodeRuntimeSource = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\node"

if ($EmptyDatabase -and $UpdateOnly) {
    throw "Le opzioni -EmptyDatabase e -UpdateOnly non possono essere usate insieme."
}

if (-not $EmptyDatabase -and -not $UpdateOnly -and [string]::IsNullOrWhiteSpace($DatabaseSourcePath)) {
    if (Test-Path -LiteralPath $installedDatabasePath) {
        $DatabaseSourcePath = $installedDatabasePath
    } else {
        $DatabaseSourcePath = $fallbackDatabasePath
    }
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Copy-DirectoryTree {
    param(
        [string]$Source,
        [string]$Destination
    )

    if (-not (Test-Path -LiteralPath $Source)) {
        throw "Cartella sorgente non trovata: $Source"
    }

    Ensure-Directory -Path $Destination
    $arguments = @(
        $Source,
        $Destination,
        "/E",
        "/R:2",
        "/W:1",
        "/NFL",
        "/NDL",
        "/NJH",
        "/NJS",
        "/NP"
    )
    & robocopy @arguments | Out-Null
    if ($LASTEXITCODE -ge 8) {
        throw "Errore durante la copia di $Source verso $Destination."
    }
}

function Copy-FileSafe {
    param(
        [string]$Source,
        [string]$Destination
    )

    if (-not (Test-Path -LiteralPath $Source)) {
        throw "File sorgente non trovato: $Source"
    }

    Ensure-Directory -Path (Split-Path -Parent $Destination)
    Copy-Item -LiteralPath $Source -Destination $Destination -Force
}

function New-EmptyPackageDatabase {
    param(
        [string]$PayloadRootPath,
        [string]$PythonRuntimePath
    )

    $pythonExe = Join-Path $PythonRuntimePath "python.exe"
    $scriptPath = Join-Path $PayloadRootPath "scripts\crea_database.py"
    if (-not (Test-Path -LiteralPath $pythonExe)) {
        throw "Python runtime del pacchetto non trovato: $pythonExe"
    }
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        throw "Script di creazione database non trovato: $scriptPath"
    }

    & $pythonExe $scriptPath | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Creazione del database vuoto non riuscita."
    }
}

if (-not (Test-Path -LiteralPath $pythonRuntimeSource)) {
    throw "Runtime Python non trovato in $pythonRuntimeSource"
}

if (-not (Test-Path -LiteralPath $nodeRuntimeSource)) {
    throw "Runtime Node non trovato in $nodeRuntimeSource"
}

if (-not $EmptyDatabase -and -not $UpdateOnly -and -not (Test-Path -LiteralPath $DatabaseSourcePath)) {
    throw "Database sorgente non trovato: $DatabaseSourcePath"
}

$appVersion = if (Test-Path -LiteralPath $versionFilePath) {
    (Get-Content -LiteralPath $versionFilePath -TotalCount 1 | Out-String).Trim()
} else {
    "2026.05.02.1"
}
$safeVersion = ($appVersion -replace '[^0-9A-Za-z._-]', '-').Trim('-')
if ([string]::IsNullOrWhiteSpace($safeVersion)) {
    $safeVersion = "0.0.0"
}

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$packageLabel = if ($UpdateOnly) {
    "OratorioCarloAcutis-aggiornamento-v$safeVersion-$timestamp"
} elseif ($EmptyDatabase) {
    "OratorioCarloAcutis-pacchetto-vuoto-v$safeVersion-$timestamp"
} else {
    "OratorioCarloAcutis-pacchetto-v$safeVersion-$timestamp"
}
$packageName = $packageLabel
$packageRoot = Join-Path $OutputRoot $packageName
$payloadRoot = Join-Path $packageRoot "payload\OratorioCarloAcutis"
$rootScriptsPath = Join-Path $packageRoot "scripts"
$zipPath = Join-Path $OutputRoot ($packageName + ".zip")

Ensure-Directory -Path $OutputRoot
if (Test-Path -LiteralPath $packageRoot) {
    Remove-Item -LiteralPath $packageRoot -Recurse -Force
}
if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

$rootFiles = @(
    "app.py",
    "launcher.pyw",
    "README.md",
    "VERSION.txt",
    "Apri Oratorio Carlo Acutis.vbs",
    "Avvia server Oratorio Carlo Acutis.bat",
    "Installa avvio automatico.vbs",
    "Disattiva avvio automatico.vbs"
)

$scriptFiles = @(
    "apri_da_desktop.ps1",
    "avvia_gestionale.ps1",
    "crea_database.py",
    "export_report.mjs",
    "sincronizza_installazione.ps1",
    "installa_pacchetto_locale.ps1",
    "crea_pacchetto_distribuzione.ps1"
)

Ensure-Directory -Path $payloadRoot
Ensure-Directory -Path $rootScriptsPath
Ensure-Directory -Path (Join-Path $payloadRoot "database")
Ensure-Directory -Path (Join-Path $payloadRoot "outputs")

foreach ($fileName in $rootFiles) {
    Copy-FileSafe `
        -Source (Join-Path $projectRoot $fileName) `
        -Destination (Join-Path $payloadRoot $fileName)
}

foreach ($folderName in @("data", "static")) {
    Copy-DirectoryTree `
        -Source (Join-Path $projectRoot $folderName) `
        -Destination (Join-Path $payloadRoot $folderName)
}

foreach ($scriptName in $scriptFiles) {
    Copy-FileSafe `
        -Source (Join-Path $projectRoot "scripts\$scriptName") `
        -Destination (Join-Path $payloadRoot "scripts\$scriptName")
    Copy-FileSafe `
        -Source (Join-Path $projectRoot "scripts\$scriptName") `
        -Destination (Join-Path $rootScriptsPath $scriptName)
}

Copy-FileSafe `
    -Source (Join-Path $projectRoot "database\schema_associazione.sql") `
    -Destination (Join-Path $payloadRoot "database\schema_associazione.sql")
Copy-FileSafe `
    -Source (Join-Path $projectRoot "database\query_utili.sql") `
    -Destination (Join-Path $payloadRoot "database\query_utili.sql")

Copy-DirectoryTree `
    -Source $pythonRuntimeSource `
    -Destination (Join-Path $payloadRoot "runtime\python")
Copy-DirectoryTree `
    -Source $nodeRuntimeSource `
    -Destination (Join-Path $payloadRoot "runtime\node")

if ($UpdateOnly) {
    # Nessun database nel pacchetto: l'installer manterra quello locale del PC di destinazione.
} elseif ($EmptyDatabase) {
    New-EmptyPackageDatabase `
        -PayloadRootPath $payloadRoot `
        -PythonRuntimePath (Join-Path $payloadRoot "runtime\python")
} else {
    Copy-FileSafe `
        -Source $DatabaseSourcePath `
        -Destination (Join-Path $payloadRoot "database\gestione_associazione.sqlite")
}

$installBatch = @'
@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\installa_pacchetto_locale.ps1"
if errorlevel 1 (
  echo.
  echo Installazione non completata.
  pause
  exit /b 1
)
echo.
echo Installazione completata.
pause
endlocal
'@
$replaceBatch = @'
@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\installa_pacchetto_locale.ps1" -ReplaceDatabase
if errorlevel 1 (
  echo.
  echo Installazione non completata.
  pause
  exit /b 1
)
echo.
echo Installazione completata con sostituzione del database locale.
pause
endlocal
'@
$updateBatch = @'
@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\installa_pacchetto_locale.ps1"
if errorlevel 1 (
  echo.
  echo Aggiornamento non completato.
  pause
  exit /b 1
)
echo.
echo Aggiornamento completato mantenendo il database locale.
pause
endlocal
'@
$guidedUpdateBatch = @'
@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\installa_pacchetto_locale.ps1"
if errorlevel 1 (
  echo.
  echo Installazione aggiornamento non completata.
  pause
  exit /b 1
)
echo.
echo Installazione aggiornamento completata mantenendo il database locale.
pause
endlocal
'@
$packageReadme = if ($UpdateOnly) { @'
PACCHETTO AGGIORNAMENTO ORATORIO CARLO ACUTIS

Usa questa versione quando vuoi aggiornare il software su un altro PC
senza toccare i dati che sono gia stati registrati su quel PC.

1. Copia questa cartella o lo ZIP sul PC da aggiornare.
2. Se hai ricevuto lo ZIP, estrailo in una cartella normale.
3. Chiudi il gestionale se e aperto.
4. Fai doppio clic su "Installa aggiornamento guidato.bat".
5. Il software verra aggiornato e il database locale restera invariato.

Note importanti:
- Questo pacchetto NON contiene il database
- Serve solo per aggiornare programma, launcher, report, grafica e script
- Versione pacchetto: APP_VERSION_PLACEHOLDER
- Durante aggiornamento Windows potrebbe chiedere una conferma amministratore
  per configurare l'indirizzo locale oratoriocarloacutis.don:8000
- I dati del PC di destinazione restano quelli gia presenti in
  %LOCALAPPDATA%\OratorioCarloAcutis\database\gestione_associazione.sqlite
'@ } elseif ($EmptyDatabase) { @'
PACCHETTO ORATORIO CARLO ACUTIS - VERSIONE VUOTA

1. Copia questa cartella o lo ZIP su un altro PC Windows.
2. Se hai ricevuto lo ZIP, estrailo in una cartella normale.
3. Fai doppio clic su "Installa su questo PC.bat".
4. Alla fine troverai il collegamento "Gestionale Oratorio Carlo Acutis" sul Desktop.
5. Al primo accesso il gestionale ti chiedera di creare l'utente amministratore.

Note importanti:
- Il pacchetto installa una copia locale in %LOCALAPPDATA%\OratorioCarloAcutis
- Non serve installare Python o Node sul PC destinazione
- Durante installazione Windows potrebbe chiedere una conferma amministratore
  per configurare l'indirizzo locale oratoriocarloacutis.don:8000
- Questa versione contiene un database pulito, senza utenti e senza dati operativi
- Se il PC ha gia un database locale, "Installa su questo PC.bat" lo mantiene
- Se vuoi sostituire anche il database locale con quello vuoto del pacchetto, usa
  "Installa sostituendo database locale.bat"

Attenzione:
- Ogni PC installato in questo modo avra una copia locale autonoma del database
- Se vuoi dati condivisi in tempo reale tra piu PC, serve invece una soluzione
  centralizzata su un solo server o PC host
'@ } else { @'
PACCHETTO ORATORIO CARLO ACUTIS

1. Copia questa cartella o lo ZIP su un altro PC Windows.
2. Se hai ricevuto lo ZIP, estrailo in una cartella normale.
3. Fai doppio clic su "Installa su questo PC.bat".
4. Alla fine troverai il collegamento "Gestionale Oratorio Carlo Acutis" sul Desktop.

Note importanti:
- Il pacchetto installa una copia locale in %LOCALAPPDATA%\OratorioCarloAcutis
- Non serve installare Python o Node sul PC destinazione
- Durante installazione Windows potrebbe chiedere una conferma amministratore
  per configurare l'indirizzo locale oratoriocarloacutis.don:8000
- Se il PC ha gia un database locale, "Installa su questo PC.bat" lo mantiene
- Se vuoi sostituire anche il database locale con quello del pacchetto, usa
  "Installa sostituendo database locale.bat"

Attenzione:
- Ogni PC installato in questo modo avra una copia locale autonoma del database
- Se vuoi dati condivisi in tempo reale tra piu PC, serve invece una soluzione
  centralizzata su un solo server o PC host
'@ }

$packageReadme = $packageReadme.Replace("APP_VERSION_PLACEHOLDER", $appVersion)

Set-Content -LiteralPath (Join-Path $packageRoot "Installa su questo PC.bat") -Value $installBatch -Encoding ASCII
Set-Content -LiteralPath (Join-Path $packageRoot "Installa sostituendo database locale.bat") -Value $replaceBatch -Encoding ASCII
Set-Content -LiteralPath (Join-Path $packageRoot "Aggiorna questo PC.bat") -Value $updateBatch -Encoding ASCII
Set-Content -LiteralPath (Join-Path $packageRoot "Installa aggiornamento guidato.bat") -Value $guidedUpdateBatch -Encoding ASCII
Set-Content -LiteralPath (Join-Path $packageRoot "LEGGIMI_INSTALLAZIONE.txt") -Value $packageReadme -Encoding UTF8

if ($CreateZip) {
    $pythonZipExe = Join-Path $pythonRuntimeSource "python.exe"
    if (-not (Test-Path -LiteralPath $pythonZipExe)) {
        throw "Python runtime non trovato per la creazione dello ZIP: $pythonZipExe"
    }
    $zipScriptPath = Join-Path $env:TEMP ("oratorio_zip_" + [guid]::NewGuid().ToString("N") + ".py")
    $zipScript = @'
from pathlib import Path
import sys
import zipfile

source_root = Path(sys.argv[1]).resolve()
zip_path = Path(sys.argv[2]).resolve()
zip_path.parent.mkdir(parents=True, exist_ok=True)
if zip_path.exists():
    zip_path.unlink()

with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
    for item in source_root.rglob("*"):
        if item.is_file():
            archive.write(item, item.relative_to(source_root))
'@
    Set-Content -LiteralPath $zipScriptPath -Value $zipScript -Encoding UTF8
    try {
        & $pythonZipExe $zipScriptPath $packageRoot $zipPath | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "Creazione ZIP non riuscita."
        }
    } finally {
        if (Test-Path -LiteralPath $zipScriptPath) {
            Remove-Item -LiteralPath $zipScriptPath -Force
        }
    }
}

Write-Host ""
Write-Host "Pacchetto creato con successo." -ForegroundColor Green
Write-Host "Cartella: $packageRoot"
if ($CreateZip) {
    Write-Host "ZIP: $zipPath"
}
if ($UpdateOnly) {
    Write-Host "Database incluso: nessuno, pacchetto solo aggiornamento software"
} elseif ($EmptyDatabase) {
    Write-Host "Database incluso: versione vuota inizializzata da schema"
} else {
    Write-Host "Database incluso: $DatabaseSourcePath"
}
