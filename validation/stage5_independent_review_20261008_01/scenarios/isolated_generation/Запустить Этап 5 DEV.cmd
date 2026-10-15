@echo off
setlocal
cd /d "%~dp0"
set "STAGE5_COMMAND=gui"
if /I "%~1"=="CheckOnly" set "STAGE5_COMMAND=check"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$env:PYTHONUTF8='1'; $p=Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'; if(-not (Test-Path -LiteralPath $p)){Write-Host 'STOP: Python 3.12 runtime missing'; exit 2}; & $p -B stage5.py $env:STAGE5_COMMAND; exit $LASTEXITCODE"
if errorlevel 1 pause
endlocal
