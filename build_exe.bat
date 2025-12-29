@echo off
REM Build script for Windows .exe
echo Building Invoice Printer executable...
echo.

REM Check if dependencies are installed
python -c "import pypdf" 2>nul
if errorlevel 1 (
    echo Installing dependencies...
    pip install -r requirements.txt
) else (
    echo Dependencies already installed.
)

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

echo.
echo Building executable (this may take a few minutes)...
echo.

REM Build the executable directly
pyinstaller --onefile --windowed --name=InvoicePrinter invoice_printer.py

if errorlevel 1 (
    echo.
    echo Build failed. Please check the error messages above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build successful!
echo ========================================
echo.
echo Executable location: dist\InvoicePrinter.exe
echo.
echo You can now copy InvoicePrinter.exe to your invoice folder and run it.
echo.
pause

