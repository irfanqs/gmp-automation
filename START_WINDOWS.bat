@echo off
chcp 65001 >nul
echo ============================================================
echo   GMP Automation System - Production Start
echo ============================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed!
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

REM Check if poppler is available
where pdftoppm >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] Poppler is not installed!
    echo PDF processing requires Poppler. Please install it:
    echo   1. Download from: https://github.com/oschwartz10612/poppler-windows/releases
    echo   2. Extract to C:\poppler
    echo   3. Add C:\poppler\Library\bin to your system PATH
    echo.
)

REM Install Python dependencies
echo [1/2] Installing Python packages...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [WARNING] Some packages may have failed to install.
)

echo.
echo [2/2] Starting GMP Automation System with Waitress WSGI...
echo.
echo ============================================================
echo   Open your browser and go to: http://localhost:5002/offline
echo   Press Ctrl+C to stop the server.
echo ============================================================
echo.

waitress-serve --host=0.0.0.0 --port=5002 app:app
pause
