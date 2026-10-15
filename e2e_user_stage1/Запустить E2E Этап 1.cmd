@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Запустить E2E Этап 1.ps1" %*
if errorlevel 1 pause
