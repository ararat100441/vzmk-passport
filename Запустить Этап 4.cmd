@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" -B stage4.py gui
if errorlevel 1 pause
