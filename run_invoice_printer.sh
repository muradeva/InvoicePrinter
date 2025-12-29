#!/bin/bash
# Invoice Printer Launcher for Mac/Linux

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if virtual environment exists, create if not
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
    echo "Installing dependencies..."
    source venv/bin/activate
    pip install -r requirements.txt
else
    source venv/bin/activate
fi

# Try different Python versions that might have tkinter
if command -v python3.14 &> /dev/null; then
    PYTHON_CMD="python3.14"
elif command -v /usr/local/bin/python3 &> /dev/null; then
    PYTHON_CMD="/usr/local/bin/python3"
elif command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
else
    echo "Error: Python 3 not found"
    exit 1
fi

# Check if tkinter is available
if ! $PYTHON_CMD -c "import tkinter" 2>/dev/null; then
    echo "Warning: tkinter not available. Using CLI version instead..."
    python invoice_printer_cli.py
else
    python invoice_printer.py
fi

if [ $? -ne 0 ]; then
    echo ""
    echo "Error: Script failed"
    echo ""
    read -p "Press enter to continue..."
fi

