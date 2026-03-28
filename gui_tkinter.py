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
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from converter import ExcelReader, convert_excel_to_drawio


class ScrollableFrame(ttk.Frame):
    """A frame with a vertical scrollbar for content that may overflow."""
    
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        
        # Create canvas and scrollbar
        self.canvas = tk.Canvas(self, highlightthickness=0, **kwargs)
        self.scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        
        # Configure canvas
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack scrollbar and canvas
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create inner frame for content
        self.inner_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner_frame, anchor=tk.NW)
        
        # Bind events
        self.inner_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # Bind mousewheel for cross-platform scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)
    
    def _on_frame_configure(self, event=None):
        """Update scrollregion when inner frame changes."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def _on_canvas_configure(self, event=None):
        """Resize inner frame to match canvas width."""
        canvas_width = event.width if event else self.canvas.winfo_width()
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)
    
    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling."""
        if sys.platform == "darwin":
            self.canvas.yview_scroll(-1 * event.delta, "units")
        else:
            self.canvas.yview_scroll(-1 * (event.delta // 120), "units")
    
    def get_inner(self):
        """Return the inner frame for adding widgets."""
        return self.inner_frame
    
    def destroy(self):
        """Clean up bindings on destroy."""
        self.canvas.unbind_all("<MouseWheel>")
        super().destroy()


class ExcelToDrawioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to draw.io Converter")
        self.root.geometry("900x700")
        self.root.minsize(700, 600)
        
        self.input_file = None
        self.sheet_data = {}
        self.selected_sheets = []
        self.output_folder = None
        
        self.setup_ui()
        
    def setup_ui(self):
        # Title
        title = tk.Label(
            self.root, 
            text="Excel to draw.io Converter",
            font=("Arial", 18, "bold")
        )
        title.pack(pady=10)
        
        # ===== MAIN PANED WINDOW =====
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        # Upper section (file + sheets + options)
        self.upper_paned = ttk.PanedWindow(self.main_paned, orient=tk.HORIZONTAL)
        self.main_paned.add(self.upper_paned, weight=3)
        
        # ===== FILE SELECTION SECTION =====
        file_frame = ttk.Frame(self.upper_paned)
        self.upper_paned.add(file_frame, weight=1)
        
        file_label_title = tk.Label(
            file_frame, 
            text="Step 1: Select Excel File", 
            font=("Arial", 11, "bold")
        )
        file_label_title.pack(pady=(5, 5))
        
        self.file_label = tk.Label(
            file_frame, 
            text="No file selected", 
            font=("Arial", 9),
            anchor=tk.W,
            wraplength=200
        )
        self.file_label.pack(fill=tk.X, padx=5, pady=2)
        
        browse_btn = tk.Button(
            file_frame,
            text="Browse Excel File",
            command=self.browse_file,
            bg="#667eea",
            fg="white",
            padx=15,
            pady=5
        )
        browse_btn.pack(pady=5)
        
        # ===== SHEET SELECTION SECTION =====
        sheets_container = ttk.Frame(self.upper_paned)
        self.upper_paned.add(sheets_container, weight=2)
        
        # Sheets label and select buttons
        sheets_header = tk.Frame(sheets_container)
        sheets_header.pack(fill=tk.X, pady=(5, 0))
        
        sheets_label = tk.Label(
            sheets_header, 
            text="Step 2: Select Sheets:", 
            font=("Arial", 11, "bold")
        )
        sheets_label.pack(side=tk.LEFT)
        
        select_all_btn = tk.Button(
            sheets_header,
            text="All",
            command=self.select_all_sheets,
            font=("Arial", 9),
            padx=8,
            pady=2
        )
        select_all_btn.pack(side=tk.RIGHT, padx=2)
        
        select_none_btn = tk.Button(
            sheets_header,
            text="None",
            command=self.deselect_all_sheets,
            font=("Arial", 9),
            padx=8,
            pady=2
        )
        select_none_btn.pack(side=tk.RIGHT)
        
        # Scrollable frame for sheet checkboxes
        self.sheets_scroll = ScrollableFrame(sheets_container)
        self.sheets_scroll.pack(fill=tk.BOTH, expand=True, padx=(0, 5), pady=5)
        
        self.sheet_vars = {}
        self.sheet_checkboxes = {}
        
        # ===== OPTIONS SECTION =====
        options_container = ttk.Frame(self.upper_paned)
        self.upper_paned.add(options_container, weight=1)
        
        options_label = tk.Label(
            options_container, 
            text="Step 3: Convert", 
            font=("Arial", 11, "bold")
        )
        options_label.pack(pady=(5, 10))
        
        # Convert button
        self.convert_btn = tk.Button(
            options_container,
            text="Convert to draw.io",
            command=self.start_conversion,
            bg="#28a745",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10,
            state=tk.DISABLED
        )
        self.convert_btn.pack(pady=20)
        
        # ===== LOG AREA (in paned window) =====
        log_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(log_frame, weight=1)
        
        log_label = tk.Label(log_frame, text="Log:", font=("Arial", 10, "bold"))
        log_label.pack(anchor=tk.W, pady=(5, 0))
        
        self.log_area = scrolledtext.ScrolledText(
            log_frame,
            height=8,
            width=80,
            state='disabled',
            font=("Courier", 8)
        )
        self.log_area.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # ===== OUTPUT SECTION =====
        output_frame = tk.Frame(self.root, pady=5)
        output_frame.pack(fill=tk.X, padx=20)
        
        self.output_path_label = tk.Label(
            output_frame, 
            text="", 
            fg="blue", 
            font=("Arial", 9), 
            anchor=tk.W
        )
        self.output_path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.open_folder_btn = tk.Button(
            output_frame, 
            text="Open Folder", 
            command=self.open_output_folder, 
            state=tk.DISABLED, 
            padx=10
        )
        self.open_folder_btn.pack(side=tk.RIGHT, padx=2)
        
        self.output_file = None
        
    def log(self, message):
        """Add message to log area"""
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')
        self.root.update_idletasks()
        
    def select_all_sheets(self):
        """Select all sheets."""
        for var in self.sheet_vars.values():
            var.set(True)
        self.update_selected_sheets()
        self.log("Selected all sheets")

    def deselect_all_sheets(self):
        """Deselect all sheets."""
        for var in self.sheet_vars.values():
            var.set(False)
        self.update_selected_sheets()
        self.log("Deselected all sheets")
        
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
        self.convert_btn.config(state=tk.DISABLED)
        
        def do_load():
            try:
                reader = ExcelReader(self.input_file)
                sheets = list(reader.wb.sheetnames)
                self.sheet_data = {sheet: {"title": sheet} for sheet in sheets}
                reader.close()
                
                # Clear previous checkboxes
                for checkbox in self.sheet_checkboxes.values():
                    checkbox.destroy()
                self.sheet_vars.clear()
                self.sheet_checkboxes.clear()
                
                # Get inner frame for adding checkboxes
                inner = self.sheets_scroll.get_inner()
                
                # Create new checkboxes
                for sheet in sheets:
                    var = tk.BooleanVar(value=True)
                    self.sheet_vars[sheet] = var
                    cb = tk.Checkbutton(
                        inner,
                        text=sheet,
                        variable=var,
                        command=self.update_selected_sheets
                    )
                    cb.pack(anchor=tk.W, pady=1)
                    self.sheet_checkboxes[sheet] = cb
                    
                self.log(f"Found {len(sheets)} sheets")
                self.update_selected_sheets()
                
            except Exception as e:
                self.log(f"Error loading file: {e}")
                messagebox.showerror("Error", f"Failed to load Excel file:\n{e}")
                
        thread = threading.Thread(target=do_load)
        thread.start()
        
    def update_selected_sheets(self):
        """Update selected sheets list"""
        self.selected_sheets = [s for s, v in self.sheet_vars.items() if v.get()]
        
        if self.selected_sheets:
            self.convert_btn.config(state=tk.NORMAL)
            self.log(f"{len(self.selected_sheets)} sheet(s) selected")
        else:
            self.convert_btn.config(state=tk.DISABLED)
            
    def start_conversion(self):
        """Start the conversion process"""
        if not self.selected_sheets:
            messagebox.showwarning("No Selection", "Please select at least one sheet")
            return

        # Ask for output file
        default_name = (
            Path(self.input_file).with_suffix(".drawio").name
            if self.input_file
            else "output.drawio"
        )
        output_path = filedialog.asksaveasfilename(
            title="Save draw.io file",
            initialdir=os.path.dirname(self.input_file) if self.input_file else None,
            initialfile=default_name,
            defaultextension=".drawio",
            filetypes=[("draw.io files", "*.drawio"), ("All files", "*.*")],
        )

        if not output_path:
            self.log("Conversion cancelled")
            return

        self.output_folder = os.path.dirname(output_path)
        self.log(f"Saving to: {output_path}")
        self.log("Starting conversion...")
        self.convert_btn.config(state=tk.DISABLED)

        def do_convert():
            try:
                self.log(f"Writing: {os.path.basename(output_path)}")
                result = convert_excel_to_drawio(
                    input_path=self.input_file,
                    output_path=output_path,
                    sheet_names=self.selected_sheets,
                )

                self.output_file = output_path
                self.log(f"Done! 1 file created ({len(result.sheet_names)} sheet pages)")

                self.root.after(0, lambda: messagebox.showinfo(
                    "Success!",
                    f"Conversion complete!\n\nSaved:\n{output_path}\n\nSheets: {len(result.sheet_names)}"
                ))

                self.root.after(0, lambda: self.output_path_label.config(text=output_path))
                self.root.after(0, lambda: self.open_folder_btn.config(state=tk.NORMAL))

            except Exception as e:
                self.log(f"Error: {e}")
                self.root.after(0, lambda: messagebox.showerror("Error", f"Conversion failed:\n{e}"))
            finally:
                self.root.after(0, lambda: self.convert_btn.config(state=tk.NORMAL))

        thread = threading.Thread(target=do_convert)
        thread.start()

    def open_output_folder(self):
        """Open the folder containing the output files."""
        if self.output_folder:
            if sys.platform == "win32":
                os.startfile(self.output_folder)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.run(["open", self.output_folder])
            else:
                import subprocess
                subprocess.run(["xdg-open", self.output_folder])

    def open_output_file(self):
        """Open the output file with default application."""
        if self.output_file and os.path.exists(self.output_file):
            if sys.platform == "win32":
                os.startfile(self.output_file)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.run(["open", self.output_file])
            else:
                import subprocess
                subprocess.run(["xdg-open", self.output_file])


def main():
    root = tk.Tk()
    app = ExcelToDrawioApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
