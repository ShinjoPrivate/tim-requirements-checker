$ErrorActionPreference = "Stop"

$bundledPython = "$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if (Test-Path $bundledPython) {
    & $bundledPython "$PSScriptRoot\app.py"
} else {
    python "$PSScriptRoot\app.py"
}
