@echo off
title Workforce Junction - Standalone App Builder
cd /d "%~dp0"

echo =================================================================
echo   Building Standalone Zero-Install Workforce Junction App
echo =================================================================
echo.

if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" build-standalone.py
) else (
    python build-standalone.py
)

echo.
pause
