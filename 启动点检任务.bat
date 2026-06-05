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

echo [1/3] Installing Python dependencies (using China mirror)...
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo.
    echo [WARN] Mirror failed, trying official PyPI...
    python -m pip install -r requirements.txt
)

echo.
echo [2/3] Installing Chromium browser (first run only, ~150MB)...
set PLAYWRIGHT_DOWNLOAD_HOST=https://cdn.npmmirror.com/binaries/playwright
python -m playwright install chromium
if errorlevel 1 (
    echo.
    echo [WARN] Mirror failed, trying default download host...
    set PLAYWRIGHT_DOWNLOAD_HOST=
    python -m playwright install chromium
)

echo.
echo [3/3] Starting crawl...
echo.
python task.py

echo.
echo ============================================================
echo   Done. Output files are in the output\ folder.
echo ============================================================
pause
