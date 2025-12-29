#!/usr/bin/env python3
"""
Invoice Triplicate-Only Printer
Prints ONLY the last triplicate page(s) of each invoice
Skips any invoice with more than 11 pages
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from pypdf import PdfReader, PdfWriter
try:
    from pypdf.generic import Transformation
except ImportError:
    # For older pypdf versions
    from pypdf import Transformation
import platform
import subprocess
import time

class TriplicateOnlyPrinter:
    def __init__(self, root):
        self.root = root
        self.root.title("Triplicate Pages Only Printer (Skip >11 Pages)")
        self.root.geometry("700x550")
        self.root.resizable(False, False)

        self.folder_var = tk.StringVar()
        self.prefix_var = tk.StringVar()
        self.start_var = tk.StringVar()
        self.end_var = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Folder selection
        ttk.Label(main_frame, text="Invoice Folder:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.folder_var, width=60).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_folder).grid(row=0, column=2, padx=5)

        # Prefix
        ttk.Label(main_frame, text="Prefix (optional):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.prefix_var, width=60).grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(main_frame, text="e.g., 'C' for C300.pdf, or 'Invoice no Example - '", 
                  font=("Arial", 8), foreground="gray").grid(row=2, column=1, sticky=tk.W)

        # Start and End
        ttk.Label(main_frame, text="Start Invoice No:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.start_var, width=30).grid(row=3, column=1, sticky=tk.W, padx=5)

        ttk.Label(main_frame, text="End Invoice No:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.end_var, width=30).grid(row=4, column=1, sticky=tk.W, padx=5)

        # Info label about skip rule
        ttk.Label(main_frame, text="Note: Invoices with more than 11 pages will be SKIPPED", 
                  foreground="red", font=("Arial", 9, "bold")).grid(row=5, column=0, columnspan=3, pady=10)

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=6, column=0, columnspan=3, pady=20)
        ttk.Button(btn_frame, text="Print Triplicate Pages Only", command=self.start_printing).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Exit", command=self.root.quit).pack(side=tk.LEFT, padx=10)

        # Status log
        self.status_text = tk.Text(main_frame, height=15, width=80, wrap=tk.WORD)
        self.status_text.grid(row=7, column=0, columnspan=3, pady=10)
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.grid(row=7, column=3, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)

        # Default folder
        self.folder_var.set(os.getcwd())

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)

    def log(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update()

    def get_triplicate_count(self, total_pages):
        if total_pages < 3:
            return 0
        # 3-5 → 1, 6-8 → 2, 9-11 → 3
        return min(((total_pages - 2) + 2) // 3, 3)  # Max 3 triplicate pages

    def create_triplicate_pdf(self, pdf_path):
        reader = PdfReader(pdf_path)
        total_pages = len(reader.pages)

        if total_pages > 11:
            raise ValueError(f"Too many pages ({total_pages} > 11)")

        triplicate_count = self.get_triplicate_count(total_pages)
        if triplicate_count == 0:
            raise ValueError("Less than 3 pages: no triplicate pages")

        # A5 size in points: 148mm x 210mm = 419.53 x 595.28 points
        A5_WIDTH = 419.53
        A5_HEIGHT = 595.28

        writer = PdfWriter()
        # Add only the last N pages, resized to A5
        for i in range(total_pages - triplicate_count, total_pages):
            page = reader.pages[i]
            
            # Get original page dimensions
            original_width = float(page.mediabox.width)
            original_height = float(page.mediabox.height)
            
            # Calculate scale to fit A5 while maintaining aspect ratio
            scale_x = A5_WIDTH / original_width
            scale_y = A5_HEIGHT / original_height
            scale = min(scale_x, scale_y)  # Use smaller scale to fit both dimensions
            
            # Calculate new dimensions
            new_width = original_width * scale
            new_height = original_height * scale
            
            # Center the page on A5
            offset_x = (A5_WIDTH - new_width) / 2
            offset_y = (A5_HEIGHT - new_height) / 2
            
            # Create transformation: scale and translate
            transformation = Transformation().scale(scale, scale).translate(offset_x, offset_y)
            page.add_transformation(transformation)
            
            # Set mediabox to A5 size
            page.mediabox.lower_left = (0, 0)
            page.mediabox.upper_right = (A5_WIDTH, A5_HEIGHT)
            
            writer.add_page(page)

        temp_path = str(Path(pdf_path).with_name(Path(pdf_path).stem + "_triplicate_only.pdf"))
        with open(temp_path, "wb") as f:
            writer.write(f)

        return temp_path, triplicate_count, total_pages

    def print_pdf(self, pdf_path):
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"File not found: {pdf_path}")

        system = platform.system()

        if system == "Windows":
            # Method 1: os.startfile - simplest and often silent
            try:
                os.startfile(pdf_path, "print")
                time.sleep(2)
                return
            except Exception as e:
                self.log(f"  → os.startfile failed: {e}")

            # Method 2: PowerShell
            try:
                ps_cmd = f'Start-Process -FilePath "{pdf_path}" -Verb Print'
                subprocess.run(["powershell", "-Command", ps_cmd], check=True, timeout=15)
                time.sleep(2)
                return
            except Exception as e:
                self.log(f"  → PowerShell failed: {e}")

            # Method 3: Adobe Reader
            adobe_paths = [
                r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            ]
            for path in adobe_paths:
                if os.path.exists(path):
                    try:
                        subprocess.run([path, "/t", pdf_path], timeout=20)
                        time.sleep(2)
                        return
                    except:
                        pass

            raise Exception(
                "All print methods failed. Please ensure:\n"
                "1. A PDF viewer is installed (Edge, Chrome, or any PDF app)\n"
                "2. PDF files are associated with a default application\n"
                "3. Or install Adobe Reader for more reliable printing"
            )

        elif system == "Darwin":
            subprocess.run(["lpr", pdf_path], check=True, timeout=30)
        else:
            subprocess.run(["lp", pdf_path], check=True, timeout=30)

    def find_files(self):
        folder = Path(self.folder_var.get())
        if not folder.exists():
            raise FileNotFoundError("Folder does not exist")

        prefix = self.prefix_var.get().strip()
        try:
            start = int(self.start_var.get())
            end = int(self.end_var.get())
            if start > end:
                raise ValueError
        except ValueError:
            raise ValueError("Invalid start/end numbers")

        files = []
        for num in range(start, end + 1):
            filename = f"{prefix}{num}.pdf"
            filepath = folder / filename
            if filepath.exists():
                files.append(str(filepath))
            else:
                self.log(f"Missing: {filename}")

        return files

    def start_printing(self):
        self.status_text.delete(1.0, tk.END)

        try:
            files = self.find_files()
            if not files:
                messagebox.showwarning("No Files", "No matching invoice files found!")
                return

            self.log(f"Found {len(files)} invoice(s). Processing (skipping >11 pages)...\n")

            printed_count = 0
            skipped_count = 0

            for i, pdf in enumerate(files, 1):
                filename = Path(pdf).name
                self.log(f"Processing [{i}/{len(files)}]: {filename}")

                temp_pdf = None
                try:
                    # This will raise error if >11 pages
                    temp_pdf, count, total_pages = self.create_triplicate_pdf(pdf)

                    self.log(f"  → {total_pages} pages → printing last {count} triplicate page(s) (A5 size)")
                    self.log(f"  → Sending to printer...")
                    self.print_pdf(temp_pdf)
                    self.log(f"  → Printed successfully\n")
                    printed_count += 1
                    
                    # Delete temp file after successful print
                    try:
                        if temp_pdf and os.path.exists(temp_pdf):
                            os.remove(temp_pdf)
                    except Exception as e:
                        self.log(f"  → Warning: Could not delete temp file: {e}\n")
                    
                    time.sleep(1)

                except ValueError as ve:
                    msg = str(ve)
                    if "Too many pages" in msg:
                        self.log(f"  → ⚠ SKIPPED: {msg}\n")
                        skipped_count += 1
                    elif "Less than 3 pages" in msg:
                        self.log(f"  → ⚠ SKIPPED: {msg}\n")
                        skipped_count += 1
                    else:
                        self.log(f"  → Error: {msg}\n")

                except Exception as e:
                    self.log(f"  → ✗ Print error: {e}\n")
                    # Try to delete temp file even on error
                    try:
                        if temp_pdf and os.path.exists(temp_pdf):
                            os.remove(temp_pdf)
                    except:
                        pass

            summary = f"\nFinished!\nPrinted triplicate for {printed_count} invoice(s)"
            if skipped_count > 0:
                summary += f"\nSkipped {skipped_count} invoice(s) (>11 pages or <3 pages)"
            self.log(summary)

            messagebox.showinfo("Complete", summary)

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.log(f"Error: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TriplicateOnlyPrinter(root)
    root.mainloop()