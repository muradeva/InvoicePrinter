# Building Windows Executable (.exe)

This guide explains how to create a standalone Windows executable (.exe) file for the Invoice Printer application.

## Prerequisites

- Windows operating system
- Python 3.7 or higher installed
- Internet connection (for downloading dependencies)

## Quick Build (Recommended)

1. **Double-click `build_exe.bat`**
   - This will automatically:
     - Create a virtual environment if needed
     - Install all dependencies
     - Build the executable

2. **Find your executable**
   - After building, look in the `dist` folder
   - You'll find `InvoicePrinter.exe`

3. **Copy to your invoice folder**
   - Copy `InvoicePrinter.exe` to the folder containing your invoice PDFs
   - Double-click to run - no installation needed!

## Manual Build

If you prefer to build manually:

```batch
# Create virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Build executable
pyinstaller --onefile --windowed --name=InvoicePrinter invoice_printer.py
```

## Using the Executable

1. **Place the .exe file** in the same folder as your invoice PDFs
2. **Double-click** `InvoicePrinter.exe` to run
3. **Enter your details**:
   - Prefix (optional): e.g., "C" for C300.pdf, or leave empty for 123.pdf
   - Start Invoice No: First invoice number
   - End Invoice No: Last invoice number
4. **Click "Print Invoices"**

The application will automatically:
- Use the current folder (where the .exe is located)
- Find all PDFs in the specified range
- Create triplicate copies of last pages
- Send to your default printer
- Clean up temporary files

## Notes

- **First run**: Windows may show a security warning. Click "More info" → "Run anyway" (if you trust the source)
- **File size**: The .exe will be around 15-20 MB (includes Python runtime)
- **Portable**: No installation required - just copy and run
- **Current directory**: The app uses the folder where the .exe is located, so place it in your invoice folder

## Troubleshooting

- **"Windows protected your PC"**: This is normal for unsigned executables. Click "More info" → "Run anyway"
- **Antivirus warning**: Some antivirus software may flag PyInstaller executables. This is a false positive
- **Missing DLL errors**: Make sure you're building on Windows, not Mac/Linux
- **Build fails**: Try running `pip install --upgrade pyinstaller`

## Distribution

You can distribute the `InvoicePrinter.exe` file to other Windows computers without requiring Python installation. Just copy the .exe file to the invoice folder and run it.

