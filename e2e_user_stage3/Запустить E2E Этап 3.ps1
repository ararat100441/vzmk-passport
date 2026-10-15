param([switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
$env:PYTHONIOENCODING = 'utf-8'
$taskPython = 'C:\Users\arara\AppData\Local\Programs\Python\Python312\python.exe'
$taskWrapper = Join-Path $PSScriptRoot 'launch_stage3_e2e.py'
$taskArguments = @('-I', '-B', $taskWrapper)
if ($CheckOnly) { $taskArguments += '--check-only' }
$taskExitCode = 1
try {
    if (-not (Test-Path -LiteralPath $taskPython -PathType Leaf)) { throw 'Python 3.12 не найден. Запуск остановлен.' }
    & $taskPython @taskArguments
    $taskExitCode = $LASTEXITCODE
} catch {
    Write-Host ('STOP: ' + $_.Exception.Message)
}
if (-not $CheckOnly) { Read-Host 'Нажмите Enter, чтобы закрыть E2E launcher' | Out-Null }
exit $taskExitCode
