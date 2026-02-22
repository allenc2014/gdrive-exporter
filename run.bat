@echo off
REM gdrive-exporter Launcher for Windows
REM This script launches the main.py application

echo Starting gdrive-exporter...

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python and add it to your PATH
    pause
    exit /b 1
)

REM Check if main.py exists in gdrive-exporter folder
if not exist "gdrive-exporter\main.py" (
    echo Error: main.py not found in gdrive-exporter directory
    echo Please ensure the gdrive-exporter folder contains main.py
    pause
    exit /b 1
)

REM Run the main application
python gdrive-exporter\main.py %*

REM Keep window open if there was an error
if errorlevel 1 (
    echo.
    echo Application exited with an error. Press any key to close...
    pause >nul
)
