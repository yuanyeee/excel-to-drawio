# Excel to Draw.io Desktop Tool Execution Guide

[🇯🇵 日本語](desktop_app_usage.md) | [🇬🇧 English](desktop_app_usage_en.md) | [🇨🇳 简体中文](desktop_app_usage_zh.md)

## Overview

This tool is a Python desktop application that converts specified sheets from Excel files into Draw.io format.

- Supported formats: `.xlsx`, `.xlsm`
- UI: `tkinter`
- Operation: Select file → Select sheets → Select options → Convert

## Prerequisites

- Windows
- Python 3 is installed
- The following files must be present in this folder:
  - `desktop_app.py`
  - `excel_to_drawio.py`

## How to Start

Run the following in the working directory:

```powershell
python .\desktop_app.py
```

## Usage

1. Click `Browse...` to select the Excel file you want to convert.
2. Select the target sheets from the loaded sheet list.
3. The output `.drawio` filename will be automatically populated in `Output`.
4. Change the save destination with `Save As...` if necessary.
5. Click `Convert`.
6. Upon successful completion, a `.drawio` file will be created at the destination.

## Key Options

- Include images: Enable/disable image embedding
- Include borders: Enable/disable border rendering
- Merge same-color fills: Render adjacent cells with the same fill color as a merged shape
- Skip hidden rows/cols: Exclude hidden rows/columns

## UI Components

- `Excel File`
  - Input file path
- `Sheets`
  - List of sheets in the workbook
  - Scrollable if there are many sheets
- `Output`
  - Output `.drawio` file path
- `Convert`
  - Execute conversion
- Bottom log area
  - Displays loading results, conversion results, and error details

## Notes

- `.xls` is not supported in this version.
- Conversion is executed synchronously. It may take some time for large files to complete.
- Depending on the sheet, complex Excel-specific shape representations may not perfectly match in Draw.io.
- If there are access permission issues with the input file or output destination, an error dialog will be displayed.

## Common Operations

### Automatic Output Destination
When you select a sheet, it will automatically be set to save in the same folder as `sheet_name.drawio`.

### Saving with a Different Name
Use `Save As...` to specify your preferred save destination and filename.

### When an Error Occurs
Check the log area at the bottom of the screen and the popup messages.
