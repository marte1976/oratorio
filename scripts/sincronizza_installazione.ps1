$sourceRoot = $env:ORATORIO_SYNC_SOURCE
$targetRoot = Split-Path -Parent $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($sourceRoot)) {
    return
}

if (-not (Test-Path -LiteralPath $sourceRoot)) {
    return
}

$resolvedSourceRoot = (Resolve-Path -LiteralPath $sourceRoot).ProviderPath
$resolvedTargetRoot = (Resolve-Path -LiteralPath $targetRoot).ProviderPath

if ($resolvedSourceRoot -ieq $resolvedTargetRoot) {
    return
}

$filesToCopy = @(
    "app.py",
    "launcher.pyw",
    "README.md",
    "VERSION.txt",
    "Apri Oratorio Carlo Acutis.vbs",
    "Avvia server Oratorio Carlo Acutis.bat",
    "Installa avvio automatico.vbs",
    "Disattiva avvio automatico.vbs",
    "Crea pacchetto per altri PC.bat",
    "Crea pacchetto vuoto per altri PC.bat",
    "Crea pacchetto aggiornamento per altri PC.bat",
    "database\schema_associazione.sql",
    "database\query_utili.sql"
)

foreach ($relativePath in $filesToCopy) {
    $sourcePath = Join-Path $resolvedSourceRoot $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        continue
    }

    $destinationPath = Join-Path $resolvedTargetRoot $relativePath
    $destinationDir = Split-Path -Parent $destinationPath
    if (-not (Test-Path -LiteralPath $destinationDir)) {
        New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
    }

    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

$foldersToCopy = @("static", "data", "scripts")

foreach ($relativeFolder in $foldersToCopy) {
    $sourceFolder = Join-Path $resolvedSourceRoot $relativeFolder
    if (-not (Test-Path -LiteralPath $sourceFolder)) {
        continue
    }

    $destinationFolder = Join-Path $resolvedTargetRoot $relativeFolder
    if (-not (Test-Path -LiteralPath $destinationFolder)) {
        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
    }

    Get-ChildItem -LiteralPath $sourceFolder -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $destinationFolder -Recurse -Force
    }
}
