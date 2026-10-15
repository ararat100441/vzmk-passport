param([switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
$taskPython = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
$taskPythonw = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\pythonw.exe'
if (-not (Test-Path -LiteralPath $taskPython) -or -not (Test-Path -LiteralPath $taskPythonw)) {
    throw 'Не найден существующий Python 3.12. Запуск E2E остановлен.'
}
# Fresh, fixed E2E paths on every launch; no production config is used.
$taskSettings = @{
    documents_root = Join-Path $PSScriptRoot 'Документы'
    registry_path = Join-Path $PSScriptRoot 'ВЗМК.xlsx'
    projects = @{
        'Гараж Казахстан' = @{
            passport_root = 'Гараж казахстан/Паспорт'
            template = 'Гараж казахстан/Паспорт/1_2026 - Приложение № 1.xlsx'
        }
    }
}
$taskConfig = Join-Path $PSScriptRoot 'settings.e2e.json'
foreach ($taskPath in @($PSScriptRoot, (Join-Path $PSScriptRoot 'app'), $taskConfig, (Join-Path $PSScriptRoot 'app\settings.json'))) {
    if ((Get-Item -LiteralPath $taskPath -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
        throw 'Обнаружено перенаправление E2E-пути. Запуск остановлен до записи конфигурации.'
    }
}
$taskSettings | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $taskConfig -Encoding UTF8
$taskSettings | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'app\settings.json') -Encoding UTF8
& $taskPython (Join-Path $PSScriptRoot 'check_e2e.py')
if ($LASTEXITCODE -ne 0) { throw 'Проверка безопасности E2E не пройдена. GUI не запущен.' }
if ($CheckOnly) { return }
$taskMain = Join-Path $PSScriptRoot 'app\main.py'
Start-Process -FilePath $taskPythonw -ArgumentList @('"' + $taskMain + '"', '--config', '"' + $taskConfig + '"') -WorkingDirectory $PSScriptRoot -WindowStyle Normal
