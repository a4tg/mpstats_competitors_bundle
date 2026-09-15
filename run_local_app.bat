@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -c "import openpyxl, selenium, requests" >nul 2>&1
if errorlevel 1 (
  py -m pip install -r requirements.txt
  if errorlevel 1 exit /b 1
)
py local_app\server.py
if errorlevel 1 pause
