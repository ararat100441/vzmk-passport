@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Запустить Этап 3.ps1"
set taskExitCode=%errorlevel%
if not "%taskExitCode%"=="0" pause
exit /b %taskExitCode%
