@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Installing/updating dependencies...
py -m pip install -r requirements.txt
if errorlevel 1 goto :error
echo.
echo Starting MPSTATS collector...
py mpstats_competitors.py
echo.
pause
exit /b 0

:error
echo.
echo Failed to install dependencies.
pause
exit /b 1
