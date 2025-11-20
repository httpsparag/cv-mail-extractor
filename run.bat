@echo off
REM CV Email Extractor - Startup Script

echo.
echo ========================================
echo   CV Email Extractor - Web UI
echo ========================================
echo.

REM Check if virtual environment exists
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo Error: Failed to create virtual environment
        pause
        exit /b 1
    )
)

REM Activate virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate.bat

REM Install/update requirements
echo Installing dependencies...
pip install -r requirements.txt -q
if errorlevel 1 (
    echo Error: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Starting Application
echo ========================================
echo.
echo Opening http://localhost:5000 in your browser...
echo Press CTRL+C to stop the server
echo.

REM Start the application
python app.py

pause
