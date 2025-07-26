@echo off
cd /d "%~dp0"

python launcher.py 2>nul
if errorlevel 1 (
    python3 launcher.py 2>nul
    if errorlevel 1 (
        echo Python is not installed or not in PATH.
        echo Please install Python 3.6 or later from https://python.org
        echo.
        echo Make sure to check "Add Python to PATH" during installation.
        pause
    )
)