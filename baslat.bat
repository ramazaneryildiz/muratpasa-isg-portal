@echo off
chcp 65001 > nul
title Muratpasa Ilce MEM ISG Portali
cls

cd /d "%~dp0"

set "PYTHON_EXE=C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe"

if not exist "%PYTHON_EXE%" (
    set "PYTHON_EXE=python"
)

:: Onceki muhtemel cakisan islemleri kapat
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8000 ^| findstr LISTENING 2^>nul') do (
    taskkill /F /PID %%a >nul 2>&1
)
taskkill /F /IM cloudflared.exe >nul 2>&1

:: Portali calistir
"%PYTHON_EXE%" "%~dp0run_portal.py"

pause
