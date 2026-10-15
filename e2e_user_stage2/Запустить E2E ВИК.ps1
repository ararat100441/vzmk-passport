param([switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
$pythonExe = 'C:\Users\arara\AppData\Local\Programs\Python\Python312\python.exe'
$wrapper = Join-Path $PSScriptRoot 'launch_vik_e2e.py'
$arguments = @('-I', '-B', $wrapper)
if ($CheckOnly) { $arguments += '--check-only' }
$exitCode = 1
try {
    if (!(Test-Path -LiteralPath $pythonExe -PathType Leaf)) { throw 'Python не найден. Ничего не установлено и не запущено.' }
    & $pythonExe @arguments
    $exitCode = $LASTEXITCODE
} catch {
    Write-Host ('STOP: ' + $_.Exception.Message)
}
if (!$CheckOnly) { Read-Host 'Нажмите Enter, чтобы закрыть тестовое окно' | Out-Null }
exit $exitCode
