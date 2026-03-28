#!/usr/bin/env python3
"""
Excel to draw.io Converter - Desktop GUI with Tkinter
Simple file picker, sheet selection, and conversion
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import json
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from converter import ExcelReader, DrawioWriter


class ExcelToDrawioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to draw.io Converter")
        self.root.geometry("700x600")
        
        self.input_file = None
        self.sheet_data = {}
        self.selected_sheets = []
        
        self.setup_ui()
        
    def setup_ui(self):
        # Title
        title = tk.Label(
            self.root, 
            text="Excel to draw.io Converter",
            font=("Arial", 18, "bold")
        )
        title.pack(pady=10)
        
        # File selection frame
        file_frame = tk.Frame(self.root, pady=10)
        file_frame.pack(fill=tk.X, padx=20)
        
        self.file_label = tk.Label(file_frame, text="No file selected", font=("Arial", 10))
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        browse_btn = tk.Button(
            file_frame,
            text="Browse Excel File",
            command=self.browse_file,
            bg="#667eea",
            fg="white",
            padx=15,
            pady=5
        )
        browse_btn.pack(side=tk.RIGHT)
        
        # Separator
        ttk.Separator(self.root, orient='horizontal').pack(fill=tk.X, padx=20, pady=10)
        
        # Sheets frame
        sheets_label = tk.Label(self.root, text="Select sheets to convert:", font=("Arial", 12, "bold"))
        sheets_label.pack(pady=5)
        
        # Checkboxes for sheets
        self.sheets_frame = tk.Frame(self.root)
        self.sheets_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        self.sheet_vars = {}
        self.sheet_checkboxes = {}
        
        # Options frame
        options_frame = tk.Frame(self.root, pady=10)
        options_frame.pack(fill=tk.X, padx=20)
        
        tk.Label(options_frame, text="Options:").pack(anchor=tk.W)
        
        # Include connectors option
        self.include_connectors = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame,
            text="Include connectors/lines",
            variable=self.include_connectors
        ).pack(anchor=tk.W)
        
        # Include cell colors option
        self.include_cell_colors = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame,
            text="Include cell background colors",
            variable=self.include_cell_colors
        ).pack(anchor=tk.W)
        
        # Separator
        ttk.Separator(self.root, orient='horizontal').pack(fill=tk.X, padx=20, pady=10)
        
        # Convert button
        self.convert_btn = tk.Button(
            self.root,
            text="Convert to draw.io",
            command=self.start_conversion,
            bg="#28a745",
            fg="white",
            font=("Arial", 14, "bold"),
            padx=30,
            pady=10,
            state=tk.DISABLED
        )
        self.convert_btn.pack(pady=10)
        
        # Progress
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=20, pady=5)
        
        # Status text
        self.status = tk.Label(self.root, text="", fg="blue")
        self.status.pack(pady=5)
        
        # Log area
        self.log_area = scrolledtext.ScrolledText(
            self.root,
            height=8,
            width=80,
            state='disabled',
            font=("Courier", 8)
        )
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
    def log(self, message):
        """Add message to log area"""
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')
        self.root.update_idletasks()
        
    def browse_file(self):
        """Open file dialog to select Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls *.xlsm"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            self.input_file = filename
            self.file_label.config(text=os.path.basename(filename))
            self.log(f"Selected: {filename}")
            self.load_sheets()
            
    def load_sheets(self):
        """Load sheets from selected Excel file"""
        if not self.input_file:
            return
            
        self.log("Loading sheets...")
        self.progress.start()
        self.convert_btn.config(state=tk.DISABLED)
        
        def do_load():
            try:
                reader = ExcelReader(self.input_file)
                self.sheet_data = reader.read_all()
                sheets = list(self.sheet_data.keys())
                
                # Clear previous checkboxes
                for checkbox in self.sheet_checkboxes.values():
                    checkbox.destroy()
                self.sheet_vars.clear()
                self.sheet_checkboxes.clear()
                
                # Create new checkboxes
                for sheet in sheets:
                    var = tk.BooleanVar(value=True)
                    self.sheet_vars[sheet] = var
                    cb = tk.Checkbutton(
                        self.sheets_frame,
                        text=sheet,
                        variable=var,
                        command=self.update_selected_sheets
                    )
                    cb.pack(anchor=tk.W)
                    self.sheet_checkboxes[sheet] = cb
                    
                self.log(f"Found {len(sheets)} sheets")
                self.update_selected_sheets()
                
            except Exception as e:
                self.log(f"Error loading file: {e}")
                messagebox.showerror("Error", f"Failed to load Excel file:\n{e}")
            finally:
                self.progress.stop()
                
        thread = threading.Thread(target=do_load)
        thread.start()
        
    def update_selected_sheets(self):
        """Update selected sheets list"""
        self.selected_sheets = [s for s, v in self.sheet_vars.items() if v.get()]
        
        if self.selected_sheets:
            self.convert_btn.config(state=tk.NORMAL)
            self.status.config(text=f"{len(self.selected_sheets)} sheet(s) selected")
        else:
            self.convert_btn.config(state=tk.DISABLED)
            self.status.config(text="No sheets selected")
            
    def start_conversion(self):
        """Start the conversion process"""
        if not self.selected_sheets:
            messagebox.showwarning("No Selection", "Please select at least one sheet")
            return
            
        self.log("Starting conversion...")
        self.convert_btn.config(state=tk.DISABLED)
        self.progress.start(10)
        
        def do_convert():
            try:
                reader = ExcelReader(self.input_file, sheet_names=self.selected_sheets)
                data = reader.read_all()
                
                self.log(f"Processing {len(data)} sheet(s)...")
                
                # Ask where to save
                default_name = os.path.splitext(os.path.basename(self.input_file))[0] + ".drawio"
                output_file = filedialog.asksaveasfilename(
                    title="Save draw.io file",
                    defaultextension=".drawio",
                    filetypes=[("draw.io files", "*.drawio"), ("All files", "*.*")],
                    initialvalue=default_name
                )
                
                if not output_file:
                    self.log("Conversion cancelled")
                    return
                    
                self.log(f"Writing to {output_file}...")
                
                writer = DrawioWriter(data)
                writer.write(output_file)
                
                size = os.path.getsize(output_file)
                self.log(f"Done! Output: {output_file} ({size/1024/1024:.1f} MB)")
                
                messagebox.showinfo(
                    "Success!",
                    f"Conversion complete!\n\nFile saved:\n{output_file}"
                )
                
            except Exception as e:
                self.log(f"Error: {e}")
                messagebox.showerror("Error", f"Conversion failed:\n{e}")
            finally:
                self.progress.stop()
                self.convert_btn.config(state=tk.NORMAL)
                
        thread = threading.Thread(target=do_convert)
        thread.start()


def main():
    root = tk.Tk()
    app = ExcelToDrawioApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
