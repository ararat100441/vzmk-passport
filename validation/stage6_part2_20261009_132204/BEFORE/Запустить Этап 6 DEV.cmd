@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$env:PYTHONUTF8='1'; $stage6Python=Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'; if(-not (Test-Path -LiteralPath $stage6Python)){Write-Host 'STOP: Python 3.12 runtime missing'; exit 2}; & $stage6Python -B stage6.py; exit $LASTEXITCODE"
if errorlevel 1 pause
endlocal
