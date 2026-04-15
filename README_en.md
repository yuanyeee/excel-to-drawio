# Excel to Draw.io Converter

[🇯🇵 日本語](README.md) | [🇬🇧 English](README_en.md) | [🇨🇳 简体中文](README_zh.md)

A tool to convert the sheet contents of Excel files (.xlsx / .xlsm) into Draw.io (.drawio) format.
Available in both desktop GUI and command line (CLI) modes.

## Features

- **Simple Operation**: Just select a file, choose sheets, and click the convert button.
- **Multi-sheet Support**: Convert multiple sheets at once.
- **Shape Conversion**: Converts shapes in Excel (rectangles, ellipses, diamonds, connectors, etc.) to Draw.io format.
- **Cell Information Conversion**: Reproduces cell fill colors, borders, text labels, and merged cells.
- **Style Retention**: Retains fill colors, line colors, and font styles.

## Supported File Formats

| Format | Description | Supported |
|------|------|------|
| .xlsx | Excel 2007 and later | ✅ |
| .xlsm | Excel with macros | ✅ |

## Installation

```powershell
cd excel-to-drawio
pip install -r requirements.txt
```

Dependencies: `openpyxl`

## Usage

### Desktop GUI

Run the following inside the `ExcelToDrawIO/` folder.

```powershell
cd ExcelToDrawIO
python desktop_app.py
```

1. Click **"Browse..."** to select an Excel file.
2. **Select the target sheets** from the loaded sheet list (multiple selection allowed).
3. The output destination is automatically set in **"Output"** (can be changed with "Save As...").
4. Click **"Convert"** to execute the conversion.
5. Once completed, a `.drawio` file will be saved.

For detailed instructions, refer to [ExcelToDrawIO/docs/desktop_app_usage_en.md](ExcelToDrawIO/docs/desktop_app_usage_en.md).

### Command Line (CLI)

Run the following inside the `ExcelToDrawIO/` folder.

```bash
# Convert all sheets (output file name is automatically generated)
python excel_to_drawio.py input.xlsx

# Specify the output file
python excel_to_drawio.py input.xlsx -o output.drawio

# Convert specific sheets only
python excel_to_drawio.py input.xlsx -s "Sheet1" "Sheet2"

# Show sheet list
python excel_to_drawio.py input.xlsx -l
```

| Option | Description |
|---|---|
| `input` (Required) | Input Excel file (.xlsx / .xlsm) |
| `-o`, `--output` | Output file path (Default: `input_file_name.drawio`) |
| `-s`, `--sheets` | Specify sheet names to convert (Default: all sheets) |
| `-l`, `--list` | Show sheet list and exit |

## Project Structure

```
excel-to-drawio/
├── ExcelToDrawIO/
│   ├── excel_to_drawio.py   # Conversion engine core (can also be run as CLI)
│   ├── desktop_app.py       # Desktop GUI app (tkinter)
│   └── docs/
│       └── desktop_app_usage.md  # GUI operation manual
├── requirements.txt          # Dependencies
├── .gitignore
└── README.md
```

## Environment Requirements

- Python 3.8+
- Windows / Mac / Linux
- tkinter (Python standard library)

## Troubleshooting

### `tkinter` error occurs
Please verify that Python is installed properly. Install Python 3.8 or later from the official website.

### Cannot open Excel file
Please make sure the file is not opened in another program. Close it and try again.

### Conversion result cannot be opened in Draw.io
Please download the latest version from the [Draw.io Official Website](https://www.drawio.com/) and try again.

## License

MIT License
