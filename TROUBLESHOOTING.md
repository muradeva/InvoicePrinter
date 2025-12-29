# Troubleshooting Build Issues

## PyInstaller Installation Failed

If you see errors like:
```
ERROR: Could not install packages due to an OSError: [WinError 2] The system cannot find the file specified
'pyinstaller' is not recognized as an internal or external command
```

### Solutions (try in order):

1. **Run as Administrator**
   - Right-click `build_exe.bat`
   - Select "Run as administrator"
   - This fixes most permission issues

2. **Upgrade pip first**
   ```batch
   python -m pip install --upgrade pip
   ```
   Then try building again.

3. **Install PyInstaller manually**
   ```batch
   python -m pip install pyinstaller
   ```
   If this fails, try:
   ```batch
   python -m pip install --user pyinstaller
   ```

4. **Check Antivirus**
   - Temporarily disable antivirus software
   - Some antivirus programs block PyInstaller installation
   - Add Python folder to antivirus exclusions if needed

5. **Close other programs**
   - Close any IDEs, terminals, or programs using Python
   - Make sure no Python processes are running

6. **Use python -m PyInstaller**
   - The build script now uses `python -m PyInstaller` instead of `pyinstaller` command
   - This is more reliable and doesn't require the Scripts folder

7. **Check Python installation**
   ```batch
   python --version
   python -m pip --version
   ```
   Make sure both work correctly.

8. **Clean install**
   ```batch
   python -m pip uninstall pyinstaller
   python -m pip install pyinstaller
   ```

## Build Succeeds but .exe Doesn't Run

1. **Windows Security Warning**
   - First run: Click "More info" → "Run anyway"
   - This is normal for unsigned executables

2. **Missing DLL errors**
   - Make sure you built on Windows (not Mac/Linux)
   - Try building on the target Windows version

3. **Antivirus blocking**
   - Some antivirus software flags PyInstaller executables
   - This is usually a false positive
   - Add exception or use different antivirus

## Other Issues

### "Module not found" errors
- Make sure all dependencies are installed: `pip install -r requirements.txt`
- Check that you're using the correct Python version

### Build takes too long
- This is normal - PyInstaller bundles Python and all libraries
- First build: 2-5 minutes
- Subsequent builds: Usually faster

### Large file size
- Normal: 15-20 MB (includes Python runtime)
- This is expected for standalone executables

## Still Having Issues?

1. Check Python version: `python --version` (should be 3.7+)
2. Check pip version: `python -m pip --version`
3. Try building manually:
   ```batch
   python -m pip install pypdf pyinstaller
   python -m PyInstaller --onefile --windowed --name=InvoicePrinter invoice_printer.py
   ```

