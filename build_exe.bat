@echo off
REM Build script for Windows .exe
echo Building Invoice Printer executable...
echo.

REM Check if dependencies are installed
python -c "import pypdf" 2>nul
if errorlevel 1 (
    echo Installing dependencies...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        echo Please check your Python installation and internet connection.
        pause
        exit /b 1
    )
) else (
    echo Dependencies already installed.
)

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    echo (If this fails, try running as Administrator or check antivirus settings)
    pip install pyinstaller
    if errorlevel 1 (
        echo.
        echo WARNING: PyInstaller installation had issues.
        echo Trying to continue anyway...
        echo.
    )
)

echo.
echo Building executable (this may take a few minutes)...
echo.

REM Build the executable using python -m PyInstaller (more reliable than pyinstaller command)
python -m PyInstaller --onefile --windowed --name=InvoicePrinter invoice_printer.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo Build failed!
    echo ========================================
    echo.
    echo Troubleshooting steps:
    echo 1. Try running this script as Administrator (right-click -^> Run as administrator)
    echo 2. Temporarily disable antivirus and try again
    echo 3. Make sure no other programs are using Python files
    echo 4. Try: pip install --upgrade pyinstaller
    echo 5. Try: python -m pip install --upgrade pip
    echo.
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

