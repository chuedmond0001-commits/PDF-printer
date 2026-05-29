@echo off
setlocal
cd /d "%~dp0"

echo Choose tool:
echo [1] Print files
echo [2] Merge PDFs/photos
echo.
set /p TOOL_CHOICE=Enter 1 or 2: 

set "SCRIPT=Print code.py"
if "%TOOL_CHOICE%"=="2" set "SCRIPT=Merge PDF Tool.py"

if exist "venv\Scripts\python.exe" (
    "venv\Scripts\python.exe" "%SCRIPT%"
) else (
    python "%SCRIPT%"
)

echo.
pause
