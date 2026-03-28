#!/usr/bin/env python3
"""
Excel to draw.io Converter - Desktop GUI with Tkinter
Simple file picker, sheet selection, and conversion

Features:
- Resizable window (default 900x700, min 700x600)
- Scrollable sheet selection list
- Paned window layout for flexible resizing
- Expandable components
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from converter import ExcelReader, DrawioWriter


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

    def destroy(self):
        """Clean up bindings on destroy."""
        self.canvas.unbind_all("<MouseWheel>")
        super().destroy()

    def get_inner(self):
        """Return the inner frame for adding widgets."""
        return self.inner_frame


class ExcelToDrawioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to draw.io Converter")

        # Window configuration - resizable with minimum size
        self.root.geometry("900x700")
        self.root.minsize(700, 600)
        self.root.resizable(True, True)

        self.input_file = None
        self.sheet_data = {}
        self.selected_sheets = []

        # Configure grid weights for resizing
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)  # Paned window area

        self.setup_ui()

    def setup_ui(self):
        # Title
        title = tk.Label(
            self.root,
            text="Excel to draw.io Converter",
            font=("Arial", 18, "bold")
        )
        title.grid(row=0, column=0, pady=10, sticky=tk.N)

        # Main container with paned window
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        self.main_paned.grid(row=1, column=0, sticky=tk.NSEW, padx=10, pady=5)

        # Upper section - Sheet selection and Options
        self.upper_paned = ttk.PanedWindow(self.main_paned, orient=tk.HORIZONTAL)
        self.main_paned.add(self.upper_paned, weight=3)

        # ===== SHEET SELECTION SECTION =====
        sheets_container = ttk.Frame(self.upper_paned)
        self.upper_paned.add(sheets_container, weight=3)

        # Sheets label and select buttons
        sheets_header = tk.Frame(sheets_container)
        sheets_header.pack(fill=tk.X, pady=(5, 0))

        sheets_label = tk.Label(
            sheets_header,
            text="Select sheets:",
            font=("Arial", 12, "bold")
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

        # Options label
        options_label = tk.Label(
            options_container,
            text="Options:",
            font=("Arial", 12, "bold")
        )
        options_label.pack(pady=(5, 10))

        # Include connectors option
        self.include_connectors = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_container,
            text="Include connectors/lines",
            variable=self.include_connectors
        ).pack(anchor=tk.W, padx=20, pady=2)

        # Include cell colors option
        self.include_cell_colors = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_container,
            text="Include cell background colors",
            variable=self.include_cell_colors
        ).pack(anchor=tk.W, padx=20, pady=2)

        # ===== FILE SELECTION =====
        # File selection frame
        file_frame = tk.Frame(self.root)
        file_frame.grid(row=2, column=0, sticky=tk.EW, padx=20, pady=5)

        self.file_label = tk.Label(
            file_frame,
            text="No file selected",
            font=("Arial", 10),
            anchor=tk.W
        )
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

        # ===== LOG AREA (in paned window) =====
        log_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(log_frame, weight=1)

        # Log area
        self.log_area = scrolledtext.ScrolledText(
            log_frame,
            height=6,
            state='disabled',
            font=("Courier", 8)
        )
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # ===== BOTTOM SECTION =====
        bottom_frame = tk.Frame(self.root)
        bottom_frame.grid(row=3, column=0, sticky=tk.EW, padx=20, pady=(0, 10))

        # Status text
        self.status = tk.Label(bottom_frame, text="No sheets selected", fg="blue", anchor=tk.W)
        self.status.pack(side=tk.LEFT, pady=5)

        # Progress
        self.progress = ttk.Progressbar(bottom_frame, mode='indeterminate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=5)

        # Convert button
        self.convert_btn = tk.Button(
            bottom_frame,
            text="Convert to draw.io",
            command=self.start_conversion,
            bg="#28a745",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=5,
            state=tk.DISABLED
        )
        self.convert_btn.pack(side=tk.RIGHT, pady=5)

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
                    initialdir=os.path.dirname(self.input_file) if self.input_file else None,
                    initialfile=default_name
                )

                if not output_file:
                    self.log("Conversion cancelled")
                    return

                self.log(f"Writing to {output_file}...")

                writer = DrawioWriter(data)
                writer.write(output_file)

                size = os.path.getsize(output_file)
                self.log(f"Done! Output: {output_file} ({size/1024/1024:.2f} MB)")

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


    def open_output_folder(self):
        if self.output_file:
            folder = os.path.dirname(self.output_file)
            if sys.platform == "win32":
                os.startfile(folder)
            elif sys.platform == "darwin":
                subprocess.run(["open", folder])
            else:
                subprocess.run(["xdg-open", folder])

    def open_output_file(self):
        if self.output_file and os.path.exists(self.output_file):
            if sys.platform == "win32":
                os.startfile(self.output_file)
            elif sys.platform == "darwin":
                subprocess.run(["open", self.output_file])
            else:
                subprocess.run(["xdg-open", self.output_file])

    def open_output_folder(self):
        if self.output_file:
            folder = os.path.dirname(self.output_file)
            if sys.platform == 'win32':
                os.startfile(folder)
            elif sys.platform == 'darwin':
                subprocess.run(['open', folder])
            else:
                subprocess.run(['xdg-open', folder])

    def open_output_file(self):
        if self.output_file and os.path.exists(self.output_file):
            if sys.platform == 'win32':
                os.startfile(self.output_file)
            elif sys.platform == 'darwin':
                subprocess.run(['open', self.output_file])
            else:
                subprocess.run(['xdg-open', self.output_file])

def main():
    root = tk.Tk()
    app = ExcelToDrawioApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
