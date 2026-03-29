import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import excel_to_drawio as etd


def supported_filetypes():
    return [
        ("Excel files", "*.xlsx *.xlsm"),
        ("Excel Workbook", "*.xlsx"),
        ("Excel Macro-Enabled Workbook", "*.xlsm"),
    ]


def format_success_message(input_path, sheet_names, output_path):
    if isinstance(sheet_names, str):
        names = [sheet_names]
    else:
        names = [str(name) for name in sheet_names if str(name).strip()]

    if len(names) <= 1:
        sheet_name = names[0] if names else ""
        return (
            f"Converted '{Path(input_path).name}'\n"
            f"Sheet: {sheet_name}\n"
            f"Output: {output_path}"
        )

    preview = ", ".join(names[:3])
    if len(names) > 3:
        preview += ", ..."
    return (
        f"Converted '{Path(input_path).name}'\n"
        f"Sheets: {len(names)}\n"
        f"Selection: {preview}\n"
        f"Output: {output_path}"
    )


class ExcelToDrawioApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel to Draw.io Converter")
        self.geometry("820x560")
        self.minsize(720, 420)
        self.resizable(True, True)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Select an Excel file to begin.")

        self._build_layout()

    def _build_layout(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(3, weight=1)

        ttk.Label(self, text="Excel File").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))
        ttk.Entry(self, textvariable=self.input_var).grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=(12, 6))
        ttk.Button(self, text="Browse...", command=self.on_browse).grid(row=0, column=2, sticky="ew", padx=(0, 12), pady=(12, 6))

        ttk.Label(self, text="Sheets").grid(row=1, column=0, sticky="nw", padx=12, pady=6)
        sheet_frame = ttk.Frame(self)
        sheet_frame.grid(row=1, column=1, columnspan=2, sticky="nsew", padx=(0, 12), pady=6)
        sheet_frame.columnconfigure(0, weight=1)
        sheet_frame.rowconfigure(0, weight=1)

        self.sheet_list = tk.Listbox(sheet_frame, exportselection=False, selectmode="extended")
        self.sheet_list.grid(row=0, column=0, sticky="nsew")
        self.sheet_list.bind("<<ListboxSelect>>", self.on_sheet_selected)

        sheet_scroll = ttk.Scrollbar(sheet_frame, orient="vertical", command=self.sheet_list.yview)
        sheet_scroll.grid(row=0, column=1, sticky="ns")
        self.sheet_list.configure(yscrollcommand=sheet_scroll.set)

        ttk.Label(self, text="Output").grid(row=2, column=0, sticky="w", padx=12, pady=6)
        ttk.Entry(self, textvariable=self.output_var).grid(row=2, column=1, sticky="ew", padx=(0, 6), pady=6)
        ttk.Button(self, text="Save As...", command=self.on_save_as).grid(row=2, column=2, sticky="ew", padx=(0, 12), pady=6)

        log_frame = ttk.Frame(self)
        log_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", padx=12, pady=6)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)

        ttk.Label(log_frame, textvariable=self.status_var).grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.log_text = tk.Text(log_frame, wrap="word", height=10, state="disabled")
        self.log_text.grid(row=1, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=1, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        action_frame = ttk.Frame(self)
        action_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 12))
        action_frame.columnconfigure(0, weight=1)
        self.convert_button = ttk.Button(action_frame, text="Convert", command=self.on_convert)
        self.convert_button.grid(row=0, column=1, sticky="e")

    def append_log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def selected_sheet_name(self):
        selection = self.sheet_list.curselection()
        if not selection:
            return None
        return self.sheet_list.get(selection[0])

    def selected_sheet_names(self):
        selection = self.sheet_list.curselection()
        return [self.sheet_list.get(index) for index in selection]

    def on_browse(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=supported_filetypes(),
        )
        if not path:
            return

        self.input_var.set(path)
        self.sheet_list.delete(0, "end")
        self.output_var.set("")

        try:
            sheets = etd.list_supported_sheets(path)
        except Exception as exc:
            self.status_var.set("Failed to load sheets.")
            self.append_log(f"Error loading sheets: {exc}")
            messagebox.showerror("Load Error", str(exc))
            return

        for sheet in sheets:
            self.sheet_list.insert("end", sheet)

        if sheets:
            self.sheet_list.selection_clear(0, "end")
            self.sheet_list.selection_set(0)
            self.on_sheet_selected()
            self.status_var.set(f"Loaded {len(sheets)} sheet(s).")
            self.append_log(f"Loaded workbook: {path}")
        else:
            self.status_var.set("No sheets found in workbook.")
            self.append_log(f"No sheets found in workbook: {path}")

    def on_sheet_selected(self, _event=None):
        input_path = self.input_var.get().strip()
        sheet_names = self.selected_sheet_names()
        if not input_path or not sheet_names:
            self.output_var.set("")
            return
        if len(sheet_names) == 1:
            self.output_var.set(etd.suggest_output_path(input_path, sheet_names[0]))
        else:
            self.output_var.set(etd.suggest_multi_output_path(input_path))

    def on_save_as(self):
        current = self.output_var.get().strip() or "output.drawio"
        current_path = Path(current)
        initial_dir = str(current_path.parent) if current_path.parent.exists() else None
        path = filedialog.asksaveasfilename(
            title="Save Draw.io file",
            defaultextension=".drawio",
            filetypes=[("Draw.io file", "*.drawio")],
            initialfile=current_path.name,
            initialdir=initial_dir,
        )
        if path:
            self.output_var.set(path)

    def on_convert(self):
        input_path = self.input_var.get().strip()
        sheet_names = self.selected_sheet_names()
        output_path = self.output_var.get().strip()

        if not input_path:
            messagebox.showerror("Missing File", "Please select an Excel file.")
            return
        if not sheet_names:
            messagebox.showerror("Missing Sheet", "Please select at least one sheet.")
            return
        if not output_path:
            messagebox.showerror("Missing Output", "Please choose an output path.")
            return

        self.convert_button.configure(state="disabled")
        self.status_var.set("Converting...")
        self.append_log(f"Converting {len(sheet_names)} sheet(s) from '{input_path}'...")
        self.update_idletasks()

        try:
            etd.convert_sheets_to_file(input_path, sheet_names, output_path)
        except Exception as exc:
            self.status_var.set("Conversion failed.")
            self.append_log(f"Conversion failed: {exc}")
            messagebox.showerror("Conversion Error", str(exc))
        else:
            msg = format_success_message(input_path, sheet_names, output_path)
            self.status_var.set("Conversion complete.")
            self.append_log(msg)
            messagebox.showinfo("Success", msg)
        finally:
            self.convert_button.configure(state="normal")


def main():
    app = ExcelToDrawioApp()
    app.mainloop()


if __name__ == "__main__":
    main()
