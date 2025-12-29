@echo off
REM Simple build script - no virtual environment
REM Just installs dependencies and builds

echo Installing/updating dependencies...
pip install -r requirements.txt

echo.
echo Building executable...
REM Use python -m PyInstaller instead of pyinstaller command (more reliable)
python -m PyInstaller --onefile --windowed --name=InvoicePrinter invoice_printer.py

if errorlevel 1 (
    echo Build failed!
    pause
) else (
    echo.
    echo Success! Find InvoicePrinter.exe in the 'dist' folder
    pause
)

