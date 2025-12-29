@echo off
REM Invoice Printer Launcher for Windows
cd /d "%~dp0"

REM Check if virtual environment exists, create if not
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    echo Installing dependencies...
    call venv\Scripts\activate.bat
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate.bat
)

REM Try to run GUI version, fallback to CLI if tkinter not available
python -c "import tkinter" 2>nul
if errorlevel 1 (
    echo Warning: tkinter not available. Using CLI version instead...
    python invoice_printer_cli.py
) else (
    python invoice_printer.py
)

if errorlevel 1 (
    echo.
    echo Error: Script failed
    echo Please make sure Python is installed and dependencies are installed
    echo.
    pause
)

