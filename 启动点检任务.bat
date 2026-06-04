@echo off
cd /d "%~dp0"

echo ============================================================
echo   Autohome Chery Dealer Inspection Tool
echo   Before running: open task.py, set the two paths at top:
echo     dealer_list_xlsx  ^<-- dealer list .xlsx
echo     standard_xlsx     ^<-- pricing standard .xlsx
echo ============================================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.11+
    echo Download: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/2] Installing dependencies...
python -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies. Check network connection.
    pause
    exit /b 1
)

echo [2/2] Starting task...
echo.
python task.py

echo.
echo ============================================================
echo   Done. Output files are in the output\ folder.
echo ============================================================
pause
