#!/usr/bin/env python3
"""
Invoice Printer CLI Application
Automatically prints invoices with triplicate copies of last pages
Works on both Windows and Mac
"""

import os
import sys
from pathlib import Path
from pypdf import PdfReader, PdfWriter
import subprocess
import platform
import time


class InvoicePrinterCLI:
    def __init__(self):
        pass
        
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
        Create a new PDF with ONLY the triplicate pages (last page(s))
        (No original pages, just the last page(s) printed once)
        """
        try:
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            total_pages = len(reader.pages)
            
            triplicate_count = self.get_triplicate_pages(total_pages)
            
            if triplicate_count == 0:
                raise Exception("PDF has less than 3 pages - no triplicate pages to print")
            
            # Add ONLY the triplicate pages (last page(s), printed once)
            start_triplicate = total_pages - triplicate_count
            # Add the last page(s) once
            for i in range(start_triplicate, total_pages):
                page = reader.pages[i]
                writer.add_page(page)
            
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
        errors = []
        
        try:
            if system == "Windows":
                # Windows: Try multiple methods to print PDF
                
                # Method 1: Try using Adobe Reader if available (most reliable)
                acroread_paths = [
                    r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                    r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                    r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                ]
                for acroread_path in acroread_paths:
                    if os.path.exists(acroread_path):
                        try:
                            subprocess.run([
                                acroread_path, "/t", pdf_path
                            ], check=True, timeout=20)
                            time.sleep(2)
                            return  # Success!
                        except Exception as e:
                            errors.append(f"Adobe Reader: {str(e)}")
                            continue
                
                # Method 2: Try PowerShell with Start-Process and Print verb
                try:
                    ps_command = f'Start-Process -FilePath "{pdf_path}" -Verb Print -WindowStyle Hidden'
                    result = subprocess.run([
                        "powershell", "-Command", ps_command
                    ], check=True, timeout=20, capture_output=True, text=True)
                    if result.returncode == 0:
                        time.sleep(2)
                        return  # Success!
                    else:
                        errors.append(f"PowerShell: {result.stderr}")
                except Exception as e:
                    errors.append(f"PowerShell: {str(e)}")
                
                # Method 3: Try os.startfile with print verb
                try:
                    os.startfile(pdf_path, "print")
                    time.sleep(2)
                    return  # Success!
                except Exception as e:
                    errors.append(f"os.startfile: {str(e)}")
                
                # If all methods fail, provide helpful error
                error_details = "\n".join(errors) if errors else "All print methods failed"
                raise Exception(
                    f"Could not print PDF.\n\n"
                    f"Errors: {error_details}\n\n"
                    f"Solutions:\n"
                    f"1. Install Adobe Reader (free) from adobe.com\n"
                    f"2. Set a default PDF application\n"
                    f"3. Manually print: {pdf_path}"
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
                print(f"Warning: {filename} not found")
        
        return files
    
    def print_invoices(self, folder_path, prefix, start_no, end_no):
        """
        Main function to process and print invoices
        """
        print(f"\n{'='*60}")
        print("Invoice Printer - Triplicate Last Pages")
        print(f"{'='*60}\n")
        
        # Find invoice files
        print(f"Searching for invoices from {start_no} to {end_no}...")
        invoice_files = self.find_invoice_files(folder_path, prefix, start_no, end_no)
        
        if not invoice_files:
            print("No invoice files found in the specified range")
            return False
        
        print(f"Found {len(invoice_files)} invoice file(s)\n")
        
        # Process each invoice
        temp_files = []
        success_count = 0
        
        for i, pdf_path in enumerate(invoice_files, 1):
            print(f"\n[{i}/{len(invoice_files)}] Processing {pdf_path.name}...")
            temp_path = None
            
            try:
                # Create triplicate PDF
                temp_path, total_pages, triplicate_count = self.create_triplicate_pdf(str(pdf_path))
                temp_files.append(temp_path)
                
                # Verify temp file
                verify_reader = PdfReader(temp_path)
                final_pages = len(verify_reader.pages)
                
                print(f"  - Original PDF pages: {total_pages}")
                print(f"  - Triplicate pages: last {triplicate_count} page(s)")
                print(f"  - Pages to print: {final_pages} (once)")
                print(f"  - Sending to printer...", end=" ", flush=True)
                
                # Print
                self.print_pdf(temp_path)
                print("✓")
                success_count += 1
                
                # Small delay between print jobs to avoid overwhelming printer
                if i < len(invoice_files):
                    time.sleep(2)
                
            except Exception as e:
                print(f"✗ Error: {str(e)}")
                # Still try to clean up temp file on error if it was created
                if temp_path and os.path.exists(temp_path):
                    temp_files.append(temp_path)
                continue
        
        # Cleanup temporary files
        print(f"\nCleaning up temporary files...")
        for temp_file in temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except:
                pass
        
        print(f"\n{'='*60}")
        print(f"Process completed! Successfully printed {success_count}/{len(invoice_files)} invoice(s)")
        print(f"{'='*60}\n")
        return True


def main():
    printer = InvoicePrinterCLI()
    
    print("\nInvoice Printer - CLI Version")
    print("=" * 60)
    
    # Get folder path
    folder_path = input("\nEnter invoice folder path: ").strip()
    if not folder_path:
        print("Error: Folder path is required")
        return
    
    if not os.path.exists(folder_path):
        print(f"Error: Folder not found: {folder_path}")
        return
    
    # Get prefix
    prefix = input("Enter prefix (optional, e.g., 'C' for C300.pdf, or press Enter for none): ").strip()
    
    # Get invoice range
    try:
        start_str = input("Enter start invoice number: ").strip()
        end_str = input("Enter end invoice number: ").strip()
        
        if not start_str or not end_str:
            print("Error: Start and end invoice numbers are required")
            return
        
        start_no = int(start_str)
        end_no = int(end_str)
        
        if start_no > end_no:
            print("Error: Start invoice number must be less than or equal to end invoice number")
            return
        
    except ValueError:
        print("Error: Invoice numbers must be integers")
        return
    
    # Confirm
    print(f"\nReady to print invoices:")
    print(f"  Folder: {folder_path}")
    print(f"  Prefix: {prefix if prefix else '(none)'}")
    print(f"  Range: {start_no} to {end_no}")
    
    confirm = input("\nProceed with printing? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Cancelled.")
        return
    
    # Print invoices
    printer.print_invoices(folder_path, prefix, start_no, end_no)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nCancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\nError: {str(e)}")
        sys.exit(1)

