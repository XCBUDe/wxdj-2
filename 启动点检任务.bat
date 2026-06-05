@echo off
cd /d "%~dp0"

echo ============================================================
echo   Autohome Chery Dealer Inspection Tool
echo   Put your two xlsx files in this folder, then run.
echo ============================================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.11+
    echo Download: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] Installing Python dependencies...
python -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies. Check network connection.
    pause
    exit /b 1
)

echo [2/3] Installing Chromium browser (first run only, ~150MB)...
python -m playwright install chromium
if errorlevel 1 (
    echo [ERROR] Failed to install Chromium. Check network connection.
    pause
    exit /b 1
)

echo [3/3] Starting crawl...
echo.
python task.py

echo.
echo ============================================================
echo   Done. Output files are in the output\ folder.
echo ============================================================
pause
