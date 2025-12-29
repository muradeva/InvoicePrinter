@echo off
REM Build script for Windows .exe
echo Building Invoice Printer executable...
echo.

REM Check if virtual environment exists
if exist "venv" (
    call venv\Scripts\activate.bat
) else (
    echo Creating virtual environment...
    python -m venv venv
    call venv\Scripts\activate.bat
    echo Installing dependencies...
    pip install -r requirements.txt
    pip install pyinstaller
)

REM Build the executable
python build_exe.py

if errorlevel 1 (
    echo.
    echo Build failed. Trying alternative method...
    pyinstaller --onefile --windowed --name=InvoicePrinter invoice_printer.py
)

echo.
echo Done! Check the 'dist' folder for InvoicePrinter.exe
pause

