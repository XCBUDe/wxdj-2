@echo off
cd /d "%~dp0"
echo Running yiche_probe2.py ...
python yiche_probe2.py > yiche_probe2_output.txt 2>&1
echo.
echo Done. Output saved to yiche_probe2_output.txt
echo Please send that file to the agent.
pause
