@echo off
rem Launch the inspection control panel GUI
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python not found. Please install Python 3.10+ from python.org
  echo         and check "Add python.exe to PATH" during install.
  pause
  exit /b 1
)

python control_panel.py
if errorlevel 1 (
  echo.
  echo [ERROR] Panel exited abnormally. See message above.
  pause
)
