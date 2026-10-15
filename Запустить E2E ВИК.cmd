@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0e2e_user_stage2\Запустить E2E ВИК.ps1" %*
exit /b %errorlevel%
