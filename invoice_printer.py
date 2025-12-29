#!/usr/bin/env python3
"""
Invoice Printer Application
Automatically prints invoices with triplicate copies of last pages
Works on both Windows and Mac
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from pypdf import PdfReader, PdfWriter
import subprocess
import platform
import time


class InvoicePrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Printer - Triplicate Last Pages")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # Variables
        # Get the directory where the script/exe is located
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            self.base_path = Path(sys.executable).parent
        else:
            # Running as script
            self.base_path = Path(__file__).parent
        
        self.prefix = tk.StringVar()
        self.start_invoice = tk.StringVar()
        self.end_invoice = tk.StringVar()
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Current directory info
        folder_info = ttk.Label(main_frame, text=f"Current Directory: {self.base_path}", 
                               font=("Arial", 9), foreground="gray")
        folder_info.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Prefix
        ttk.Label(main_frame, text="Prefix (optional):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.prefix, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(main_frame, text="e.g., C or leave empty", font=("Arial", 8)).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Start Invoice
        ttk.Label(main_frame, text="Start Invoice No:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.start_invoice, width=50).grid(row=3, column=1, padx=5, pady=5)
        
        # End Invoice
        ttk.Label(main_frame, text="End Invoice No:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.end_invoice, width=50).grid(row=4, column=1, padx=5, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Print Invoices", command=self.print_invoices, width=20).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Exit", command=self.root.quit, width=20).pack(side=tk.LEFT, padx=10)
        
        # Status text
        self.status_text = tk.Text(main_frame, height=10, width=70, wrap=tk.WORD)
        self.status_text.grid(row=6, column=0, columnspan=3, pady=10)
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.grid(row=6, column=3, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
    def log_status(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update()
        
    def get_triplicate_pages(self, total_pages):
        """
        Determine how many last pages should be triplicated
        - 3-5 pages: last 1 page triplicate
        - 6-8 pages: last 2 pages triplicate
        - 9-11 pages: last 3 pages triplicate
        - Pattern: for pages in range (3n, 3n+2], last n pages triplicate
        """
        if total_pages < 3:
            return 0
        
        # Calculate: pages 3-5 -> 1, pages 6-8 -> 2, pages 9-11 -> 3, etc.
        # Formula: ceil((total_pages - 2) / 3)
        triplicate_count = ((total_pages - 2) + 2) // 3
        return min(triplicate_count, total_pages)  # Don't exceed total pages
        
    def create_triplicate_pdf(self, pdf_path):
        """
        Create a new PDF with triplicate copies of last pages
        """
        try:
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            total_pages = len(reader.pages)
            
            triplicate_count = self.get_triplicate_pages(total_pages)
            
            # Add all original pages
            for i in range(total_pages):
                writer.add_page(reader.pages[i])
            
            # Add triplicate copies of last pages
            if triplicate_count > 0:
                start_triplicate = total_pages - triplicate_count
                for _ in range(2):  # Add 2 more copies (making total 3)
                    for i in range(start_triplicate, total_pages):
                        writer.add_page(reader.pages[i])
            
            # Save to temporary file
            temp_path = pdf_path.replace('.pdf', '_temp_print.pdf')
            with open(temp_path, 'wb') as output_file:
                writer.write(output_file)
                
            return temp_path, total_pages, triplicate_count
            
        except Exception as e:
            raise Exception(f"Error processing PDF {pdf_path}: {str(e)}")
    
    def print_pdf(self, pdf_path):
        """
        Print PDF using system default printer
        Works on both Windows and Mac
        """
        system = platform.system()
        
        try:
            if system == "Windows":
                # Windows: Try multiple methods to print PDF
                # Escape backslashes for PowerShell
                pdf_path_escaped = pdf_path.replace("\\", "\\\\").replace('"', '\\"')
                
                # Method 1: Try using Adobe Reader if available (most reliable)
                try:
                    acroread_paths = [
                        r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                        r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                        r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                    ]
                    for acroread_path in acroread_paths:
                        if os.path.exists(acroread_path):
                            subprocess.run([
                                acroread_path, "/t", pdf_path
                            ], check=True, timeout=20, creationflags=subprocess.CREATE_NO_WINDOW)
                            time.sleep(2)
                            return
                except:
                    pass
                
                # Method 2: Try using Microsoft Edge (usually available on Windows 10+)
                try:
                    edge_paths = [
                        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                    ]
                    for edge_path in edge_paths:
                        if os.path.exists(edge_path):
                            # Edge can print PDFs directly
                            subprocess.run([
                                edge_path, f"file:///{pdf_path}", "--print"
                            ], check=True, timeout=20, creationflags=subprocess.CREATE_NO_WINDOW)
                            time.sleep(2)
                            return
                except:
                    pass
                
                # Method 3: Try PowerShell with Start-Process and Print verb
                try:
                    ps_command = f'Start-Process -FilePath "{pdf_path}" -Verb Print -WindowStyle Hidden'
                    result = subprocess.run([
                        "powershell", "-Command", ps_command
                    ], check=True, timeout=20, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
                    time.sleep(2)
                    return
                except:
                    pass
                
                # Method 4: Try os.startfile with print verb (requires file association)
                try:
                    os.startfile(pdf_path, "print")
                    time.sleep(2)
                    return
                except:
                    pass
                
                # If all methods fail, provide helpful error
                raise Exception(
                    "Could not print PDF. Please try one of these:\n"
                    "1. Install Adobe Reader (free) from adobe.com\n"
                    "2. Set a default PDF application: Right-click PDF -> Open with -> Choose default\n"
                    "3. Manually print the temporary PDF files from the invoice folder\n"
                    f"Temporary file location: {pdf_path}"
                )
                    
            elif system == "Darwin":  # macOS
                # macOS: use lpr command
                subprocess.run(["lpr", pdf_path], check=True, timeout=30)
            else:
                # Linux fallback
                subprocess.run(["lp", pdf_path], check=True, timeout=30)
                
            # Small delay to allow print job to be queued
            time.sleep(1)
                
        except subprocess.TimeoutExpired:
            raise Exception("Print command timed out")
        except Exception as e:
            raise Exception(f"Error printing PDF: {str(e)}")
    
    def find_invoice_files(self, folder_path, prefix, start_no, end_no):
        """
        Find all invoice PDF files in the given range
        """
        files = []
        folder = Path(folder_path)
        
        if not folder.exists():
            raise FileNotFoundError(f"Folder not found: {folder_path}")
        
        for invoice_no in range(start_no, end_no + 1):
            # Try with prefix
            if prefix:
                filename = f"{prefix}{invoice_no}.pdf"
            else:
                filename = f"{invoice_no}.pdf"
            
            file_path = folder / filename
            if file_path.exists():
                files.append(file_path)
            else:
                self.log_status(f"Warning: {filename} not found")
        
        return files
    
    def print_invoices(self):
        """
        Main function to process and print invoices
        """
        # Clear status
        self.status_text.delete(1.0, tk.END)
        
        # Use current directory (where exe/script is located)
        folder_path = str(self.base_path)
        prefix = self.prefix.get().strip()
        start_str = self.start_invoice.get().strip()
        end_str = self.end_invoice.get().strip()
        
        if not start_str or not end_str:
            messagebox.showerror("Error", "Please enter start and end invoice numbers")
            return
        
        try:
            start_no = int(start_str)
            end_no = int(end_str)
        except ValueError:
            messagebox.showerror("Error", "Invoice numbers must be integers")
            return
        
        if start_no > end_no:
            messagebox.showerror("Error", "Start invoice number must be less than or equal to end invoice number")
            return
        
        try:
            # Find invoice files
            self.log_status(f"Searching for invoices in: {folder_path}")
            self.log_status(f"Range: {start_no} to {end_no}...")
            invoice_files = self.find_invoice_files(folder_path, prefix, start_no, end_no)
            
            if not invoice_files:
                messagebox.showwarning("Warning", "No invoice files found in the specified range")
                return
            
            self.log_status(f"Found {len(invoice_files)} invoice file(s)")
            
            # Process each invoice
            temp_files = []
            for i, pdf_path in enumerate(invoice_files, 1):
                self.log_status(f"\nProcessing {pdf_path.name} ({i}/{len(invoice_files)})...")
                temp_path = None
                
                try:
                    # Create triplicate PDF
                    temp_path, total_pages, triplicate_count = self.create_triplicate_pdf(str(pdf_path))
                    temp_files.append(temp_path)
                    
                    self.log_status(f"  - Total pages: {total_pages}")
                    self.log_status(f"  - Triplicate pages: {triplicate_count}")
                    self.log_status(f"  - Sending to printer...")
                    
                    # Print
                    self.print_pdf(temp_path)
                    self.log_status(f"  - ✓ Printed successfully")
                    
                    # Small delay between print jobs to avoid overwhelming printer
                    if i < len(invoice_files):
                        time.sleep(2)
                    
                except Exception as e:
                    self.log_status(f"  - ✗ Error: {str(e)}")
                    # Still try to clean up temp file on error if it was created
                    if temp_path and os.path.exists(temp_path):
                        temp_files.append(temp_path)
                    continue
            
            # Cleanup temporary files
            self.log_status(f"\nCleaning up temporary files...")
            for temp_file in temp_files:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                except:
                    pass
            
            self.log_status(f"\n✓ Process completed!")
            messagebox.showinfo("Success", f"Processed {len(invoice_files)} invoice(s)")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.log_status(f"Error: {str(e)}")


def main():
    root = tk.Tk()
    app = InvoicePrinterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

