$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$taskPython = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
if (Test-Path -LiteralPath $taskPython) {
    & $taskPython (Join-Path $PSScriptRoot 'run_local.py') --server.address 127.0.0.1 --server.headless true --browser.gatherUsageStats false
} else {
    python (Join-Path $PSScriptRoot 'run_local.py') --server.address 127.0.0.1 --server.headless true --browser.gatherUsageStats false
}
