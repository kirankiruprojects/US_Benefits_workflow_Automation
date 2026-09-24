@echo off
title Workforce Junction Launcher
cd /d "%~dp0"

echo ===================================================
echo   Starting Workforce Junction Desktop Application
echo ===================================================
echo.

:: 1. Start the Vite web server in the background
echo [1/2] Starting Web UI Server (npm run dev)...
start "Workforce Junction - Web Server" /min cmd /c "npm run dev"

:: Wait 4 seconds for Vite server to boot up
timeout /t 4 /nobreak >nul

:: 2. Launch the Desktop Application Window
echo [2/2] Opening Desktop Application Window...
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" desktop\main.py
) else (
    python desktop\main.py
)

echo.
echo Application closed.
