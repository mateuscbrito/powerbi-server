@echo off
setlocal

:: Title
title SARA PowerBI Server (Source Mode)

:: Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found! Please install Python 3.10+ and add to PATH.
    pause
    exit /b 1
)

:: Create Virtual Environment if missing
if not exist "venv" (
    echo [INFO] Creating Virtual Environment...
    python -m venv venv
    
    echo [INFO] Installing Dependencies...
    call venv\Scripts\activate
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate
)

:: Run Server
echo [INFO] Starting SARA Server...
set PYTHONPATH=%~dp0src
python -u -m sara_powerbi.server

pause
