@echo off
REM COA Converter Web - Windows Start Script
setlocal

cd /d "%~dp0"

REM Prefer venv if present
if exist ".venv\Scripts\python.exe" (
    set "PYTHON=.venv\Scripts\python.exe"
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] Python not found. Install Python 3.9+ from https://python.org and re-run.
        pause
        exit /b 1
    )
    set "PYTHON=python"
)

REM Install dependencies on first run
if not exist ".deps_installed" (
    echo Installing dependencies...
    "%PYTHON%" -m pip install --upgrade pip
    "%PYTHON%" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] pip install failed.
        pause
        exit /b 1
    )
    echo. > .deps_installed
)

echo Starting COA Converter Web...
echo   Local: http://127.0.0.1:5050

REM Open browser after a short delay
start "" /b cmd /c "timeout /t 2 /nobreak >nul && start http://127.0.0.1:5050"

"%PYTHON%" app.py
endlocal
