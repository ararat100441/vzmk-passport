$ErrorActionPreference = 'Stop'
$taskPython = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\pythonw.exe'
if (-not (Test-Path -LiteralPath $taskPython)) {
    throw 'Не найден Python 3.12. Нужны Python, openpyxl и pywin32.'
}
Start-Process -FilePath $taskPython -ArgumentList @('"' + (Join-Path $PSScriptRoot 'main.py') + '"') -WorkingDirectory $PSScriptRoot -WindowStyle Normal
