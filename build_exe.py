#!/usr/bin/env python3
"""
Build script to create Windows .exe file using PyInstaller
"""

import subprocess
import sys
import os

def build_exe():
    """Build the Windows executable"""
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    print("Building Windows executable...")
    print("This may take a few minutes...\n")
    
    # PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",  # Create a single executable file
        "--windowed",  # No console window (GUI only)
        "--name=InvoicePrinter",  # Name of the executable
        "--icon=NONE",  # No icon (can add .ico file later if needed)
        "--add-data=README.md;.",  # Include README (optional)
        "invoice_printer.py"
    ]
    
    try:
        subprocess.check_call(cmd)
        print("\n" + "="*60)
        print("Build successful!")
        print("="*60)
        print(f"\nExecutable location: dist/InvoicePrinter.exe")
        print("\nYou can now:")
        print("1. Copy InvoicePrinter.exe to your invoice folder")
        print("2. Double-click to run")
        print("3. The app will automatically use the current folder")
        print("\nNote: The first run may be slower as Windows verifies the executable.")
    except subprocess.CalledProcessError as e:
        print(f"\nBuild failed with error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    build_exe()

