# Invoice Printer - Triplicate Last Pages

A cross-platform application to automatically print invoices with triplicate copies of the last pages. Works on both Windows and Mac.

## Features

- **Batch Processing**: Print multiple invoices in a range (e.g., 123-150 or C300-C350)
- **Prefix Support**: Handle invoices with prefixes (e.g., C300.pdf, C301.pdf)
- **Smart Triplication**: Automatically triplicates last pages based on document length:
  - 3-5 pages: Last 1 page triplicated
  - 6-8 pages: Last 2 pages triplicated
  - 9-11 pages: Last 3 pages triplicated
  - And so on...
- **Cross-Platform**: Works on both Windows and Mac
- **User-Friendly GUI**: Simple interface for easy operation
- **Standalone Executable**: Create a Windows .exe file that can be placed directly in your invoice folder
- **Auto-Detect Folder**: Uses the current directory (where the app is located) - no need to browse for folders

## Requirements

- Python 3.7 or higher
- Required packages (install via `pip install -r requirements.txt`):
  - pypdf>=3.17.0

## Installation

1. Clone or download this repository
2. Install dependencies:

   **Option 1: Using virtual environment (Recommended)**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

   **Option 2: Global installation**
   ```bash
   pip install -r requirements.txt
   ```
   
   Note: On some systems (especially macOS with Homebrew Python), you may need to use `pip install --break-system-packages -r requirements.txt` or use a virtual environment.

## Quick Start (Windows .exe)

**Easiest way to use on Windows:**

1. **Build the executable** (one-time setup):
   - Double-click `build_exe.bat` on a Windows machine
   - Wait for the build to complete (takes 2-5 minutes)
   - Find `InvoicePrinter.exe` in the `dist` folder

2. **Use the executable**:
   - Copy `InvoicePrinter.exe` to your invoice folder (where your PDFs are)
   - Double-click `InvoicePrinter.exe` to run
   - Enter prefix (if any), start and end invoice numbers
   - Click "Print Invoices"

The app automatically uses the folder where the .exe is located - no need to browse for folders!

See [BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md) for detailed build instructions.

## Usage

### GUI Version (Recommended)

1. If using virtual environment, activate it first:
   ```bash
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

2. Run the application:
   ```bash
   python invoice_printer.py
   ```
   
   Or use the launcher script (automatically handles virtual environment):
   - **Mac/Linux**: `./run_invoice_printer.sh`
   - **Windows**: Double-click `run_invoice_printer.bat`

2. In the GUI:
   - **Current Directory**: Shows the folder where the app is running (automatically detected)
   - **Prefix (optional)**: Enter prefix if your invoices have one (e.g., "C" for C300.pdf, C301.pdf). Leave empty for invoices without prefix (e.g., 123.pdf, 124.pdf)
   - **Start Invoice No**: Enter the starting invoice number
   - **End Invoice No**: Enter the ending invoice number
   - Click **Print Invoices** to start the process

**Note**: The app automatically uses the folder where it's located. For the .exe version, place it in your invoice folder. For the Python script, it uses the script's directory.

3. The application will:
   - Find all PDFs in the specified range
   - Create temporary PDFs with triplicated last pages
   - Send them to your default printer
   - Clean up temporary files

## How It Works

1. **Directory Detection**: The app automatically uses the folder where it's located (current directory)
2. **File Finding**: The app searches for PDF files matching the pattern `{prefix}{number}.pdf` in the specified range
3. **Page Analysis**: For each PDF, it counts total pages and determines how many last pages need triplication
4. **PDF Creation**: Creates a temporary PDF with all original pages plus 2 additional copies of the last page(s) that need triplication
5. **Printing**: Sends the modified PDF to your system's default printer
6. **Cleanup**: Removes temporary files after printing

## Triplication Logic

The app uses the following logic to determine how many last pages to triplicate:
- PDFs with 1-2 pages: No triplication
- PDFs with 3-5 pages: Last 1 page triplicated (3 copies total)
- PDFs with 6-8 pages: Last 2 pages triplicated (3 copies each)
- PDFs with 9-11 pages: Last 3 pages triplicated (3 copies each)
- Pattern continues: For every 3 additional pages, one more last page is triplicated

## Notes

- Make sure your default printer is set up and ready
- The app will skip missing invoice numbers and continue with available ones
- Temporary files are automatically cleaned up after printing
- On Windows, the app uses the default system print dialog
- On Mac, the app uses the `lpr` command to print

### CLI Version (Alternative)

If tkinter is not available, you can use the command-line version:

```bash
# If using virtual environment
source venv/bin/activate  # On Windows: venv\Scripts\activate
python invoice_printer_cli.py
```

The CLI version will prompt you for:
- Invoice folder path
- Prefix (optional)
- Start invoice number
- End invoice number
- Confirmation before printing

## Troubleshooting

- **"No module named '_tkinter'"**: 
  - **Mac (Homebrew)**: Install tkinter support: `brew install python-tk`
  - **Mac (System Python)**: System Python usually has tkinter. Use `/usr/local/bin/python3` or install python-tk
  - **Windows**: tkinter is usually included with Python. If missing, reinstall Python with "tcl/tk" support
  - **Alternative**: Use the CLI version (`invoice_printer_cli.py`) which doesn't require tkinter

- **"No invoice files found"**: Check that your folder path is correct and invoice numbers match the naming pattern

- **Print errors**: Ensure your printer is connected and set as default

- **Permission errors**: Make sure you have read access to the invoice folder and write access for temporary files

## License

This project is provided as-is for personal/commercial use.

