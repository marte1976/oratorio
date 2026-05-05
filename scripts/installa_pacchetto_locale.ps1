param(
    [string]$TargetRoot = (Join-Path $env:LOCALAPPDATA "OratorioCarloAcutis"),
    [switch]$ReplaceDatabase,
    [switch]$NoLaunch
)

$ErrorActionPreference = "Stop"
$BrowserHostName = "oratoriocarloacutis.don"
$BrowserHostEntry = "127.0.0.1 $BrowserHostName"
$BrowserUrl = "http://$BrowserHostName`:8000/login"

$packageRoot = Split-Path -Parent $PSScriptRoot
$payloadRoot = Join-Path $packageRoot "payload\OratorioCarloAcutis"
$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "Gestionale Oratorio Carlo Acutis.lnk"
$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Oratorio Carlo Acutis"
$startMenuShortcut = Join-Path $startMenuDir "Gestionale Oratorio Carlo Acutis.lnk"
$startupDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Startup"
$startupShortcut = Join-Path $startupDir "Oratorio Carlo Acutis - avvio in background.lnk"

function Get-VersionString {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return ""
    }
    return ((Get-Content -LiteralPath $Path -TotalCount 1 | Out-String).Trim())
}

function Compare-VersionString {
    param(
        [string]$Left,
        [string]$Right
    )

    if ([string]::IsNullOrWhiteSpace($Left) -and [string]::IsNullOrWhiteSpace($Right)) {
        return 0
    }
    if ([string]::IsNullOrWhiteSpace($Left)) {
        return -1
    }
    if ([string]::IsNullOrWhiteSpace($Right)) {
        return 1
    }

    try {
        $leftVersion = [version]($Left -replace '[^0-9.]', '')
        $rightVersion = [version]($Right -replace '[^0-9.]', '')
        return $leftVersion.CompareTo($rightVersion)
    } catch {
        return [string]::Compare($Left, $Right, $true)
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
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

function Create-Shortcut {
    param(
        [string]$ShortcutPath,
        [string]$TargetPath,
        [string]$Arguments,
        [string]$WorkingDirectory,
        [string]$IconLocation
    )

    Ensure-Directory -Path (Split-Path -Parent $ShortcutPath)
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.Arguments = $Arguments
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.IconLocation = $IconLocation
    $shortcut.Save()
}

function Ensure-HostsAlias {
    param([string]$HostName)

    $hostsPath = Join-Path $env:SystemRoot "System32\drivers\etc\hosts"
    $lines = @()
    if (Test-Path -LiteralPath $hostsPath) {
        $lines = Get-Content -LiteralPath $hostsPath -ErrorAction Stop
    }

    $normalizedLines = [System.Collections.Generic.List[string]]::new()
    $found = $false
    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ($trimmed -match '^\s*#' -or [string]::IsNullOrWhiteSpace($trimmed)) {
            $normalizedLines.Add($line)
            continue
        }

        $tokens = $trimmed -split '\s+'
        if ($tokens.Count -ge 2 -and $tokens[1..($tokens.Count - 1)] -contains $HostName) {
            if (-not $found) {
                $normalizedLines.Add("127.0.0.1 $HostName")
                $found = $true
            }
            continue
        }

        $normalizedLines.Add($line)
    }

    if (-not $found) {
        if ($normalizedLines.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($normalizedLines[$normalizedLines.Count - 1])) {
            $normalizedLines.Add("")
        }
        $normalizedLines.Add("127.0.0.1 $HostName")
    }

    Set-Content -LiteralPath $hostsPath -Value $normalizedLines -Encoding ASCII
}

if (-not (Test-IsAdministrator)) {
    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", ('"' + $PSCommandPath + '"'),
        "-TargetRoot", ('"' + $TargetRoot + '"')
    )
    if ($ReplaceDatabase) {
        $arguments += "-ReplaceDatabase"
    }
    if ($NoLaunch) {
        $arguments += "-NoLaunch"
    }
    Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs | Out-Null
    exit
}

if (-not (Test-Path -LiteralPath $payloadRoot)) {
    throw "Pacchetto non valido: cartella payload non trovata."
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

$folderTrees = @(
    "data",
    "outputs",
    "runtime",
    "scripts",
    "static"
)

Ensure-Directory -Path $TargetRoot
Ensure-Directory -Path (Join-Path $TargetRoot "database")
Ensure-Directory -Path (Join-Path $TargetRoot "outputs")

$existingVersionPath = Join-Path $TargetRoot "VERSION.txt"
$packageVersionPath = Join-Path $payloadRoot "VERSION.txt"
$installedVersion = Get-VersionString -Path $existingVersionPath
$packageVersion = Get-VersionString -Path $packageVersionPath

foreach ($fileName in $rootFiles) {
    Copy-FileSafe `
        -Source (Join-Path $payloadRoot $fileName) `
        -Destination (Join-Path $TargetRoot $fileName)
}

foreach ($folderName in $folderTrees) {
    Copy-DirectoryTree `
        -Source (Join-Path $payloadRoot $folderName) `
        -Destination (Join-Path $TargetRoot $folderName)
}

Copy-FileSafe `
    -Source (Join-Path $payloadRoot "database\schema_associazione.sql") `
    -Destination (Join-Path $TargetRoot "database\schema_associazione.sql")
Copy-FileSafe `
    -Source (Join-Path $payloadRoot "database\query_utili.sql") `
    -Destination (Join-Path $TargetRoot "database\query_utili.sql")

$sourceDbPath = Join-Path $payloadRoot "database\gestione_associazione.sqlite"
$targetDbPath = Join-Path $TargetRoot "database\gestione_associazione.sqlite"
$hasDatabasePayload = Test-Path -LiteralPath $sourceDbPath
if (-not $hasDatabasePayload) {
    Write-Host "Pacchetto di aggiornamento: database locale mantenuto senza modifiche." -ForegroundColor Yellow
} elseif ($ReplaceDatabase -or -not (Test-Path -LiteralPath $targetDbPath)) {
    Copy-FileSafe -Source $sourceDbPath -Destination $targetDbPath
} else {
    Write-Host "Database locale gia presente: mantenuto senza sovrascrittura." -ForegroundColor Yellow
}

$powershellPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$wscriptPath = Join-Path $env:SystemRoot "System32\wscript.exe"
$launchVbsPath = Join-Path $TargetRoot "Apri Oratorio Carlo Acutis.vbs"
$iconPath = Join-Path $TargetRoot "static\logo-ca.ico"

Ensure-HostsAlias -HostName $BrowserHostName
ipconfig /flushdns | Out-Null

Create-Shortcut `
    -ShortcutPath $desktopShortcut `
    -TargetPath $wscriptPath `
    -Arguments ('"' + $launchVbsPath + '"') `
    -WorkingDirectory $TargetRoot `
    -IconLocation ($iconPath + ",0")

Create-Shortcut `
    -ShortcutPath $startMenuShortcut `
    -TargetPath $wscriptPath `
    -Arguments ('"' + $launchVbsPath + '"') `
    -WorkingDirectory $TargetRoot `
    -IconLocation ($iconPath + ",0")

Create-Shortcut `
    -ShortcutPath $startupShortcut `
    -TargetPath $wscriptPath `
    -Arguments ('"' + $launchVbsPath + '" --nobrowser') `
    -WorkingDirectory $TargetRoot `
    -IconLocation ($iconPath + ",0")

Write-Host ""
Write-Host "Installazione completata." -ForegroundColor Green
Write-Host "Percorso: $TargetRoot"
Write-Host "Collegamento Desktop: $desktopShortcut"
if (-not [string]::IsNullOrWhiteSpace($packageVersion)) {
    Write-Host "Versione pacchetto: $packageVersion"
}
if (-not [string]::IsNullOrWhiteSpace($installedVersion)) {
    Write-Host "Versione precedente installata: $installedVersion"
}
$compareResult = Compare-VersionString -Left $packageVersion -Right $installedVersion
if ($compareResult -gt 0) {
    Write-Host "Controllo versione: il pacchetto e piu recente della versione precedente." -ForegroundColor Green
} elseif ($compareResult -eq 0 -and -not [string]::IsNullOrWhiteSpace($packageVersion)) {
    Write-Host "Controllo versione: stai reinstallando la stessa versione del software." -ForegroundColor Yellow
} elseif ($compareResult -lt 0) {
    Write-Host "Controllo versione: il pacchetto risulta precedente rispetto alla versione gia installata." -ForegroundColor Yellow
}

if (-not $NoLaunch) {
    Start-Process -FilePath $wscriptPath -ArgumentList ('"' + $launchVbsPath + '"') -WindowStyle Hidden
}
