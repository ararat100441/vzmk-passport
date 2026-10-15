param([switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
$taskPython = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\pythonw.exe'
$taskConsolePython = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
if (-not (Test-Path -LiteralPath $taskPython)) {
    throw 'Не найден Python 3.12. Используется тот же runtime, что для Этапов 1–2.'
}
Push-Location -LiteralPath $PSScriptRoot
try {
    & $taskConsolePython -c "from passport_app.config import Settings; from passport_app.catalog import CatalogService; import passport_app.catalog_gui; CatalogService(Settings.load('settings.json'))"
    if ($LASTEXITCODE -ne 0) { throw 'Проверка запуска Этапа 3 не пройдена.' }
} finally {
    Pop-Location
}
if ($CheckOnly) { Write-Output 'Stage 3 launcher preflight: PASS'; exit 0 }
Start-Process -FilePath $taskPython -ArgumentList @('"' + (Join-Path $PSScriptRoot 'stage3.py') + '"') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden
