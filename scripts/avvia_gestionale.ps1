$projectRoot = Split-Path -Parent $PSScriptRoot
$bundledPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if (Test-Path $bundledPython) {
    $pythonExe = $bundledPython
} else {
    $pythonExe = "python"
}

& $pythonExe (Join-Path $projectRoot "app.py")
