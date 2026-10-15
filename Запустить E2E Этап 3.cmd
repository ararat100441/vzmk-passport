@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0e2e_user_stage3\Запустить E2E Этап 3.ps1" %*
exit /b %errorlevel%
