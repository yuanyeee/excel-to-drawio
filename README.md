# Excel to draw.io Converter

Convert Excel shapes, diagrams, flowcharts, and ER diagrams to draw.io format.

## Features

- 🔄 Convert shapes from multiple Excel sheets to a single draw.io file
- 📑 Each Excel sheet becomes a separate tab/page in draw.io
- 📐 Supports various shape types: rectangles, ellipses, diamonds, connectors, etc.
- 🎨 Preserves basic styling: fill color, stroke color, text
- 📊 **Cell-based diagrams**: Merged cells with borders are converted to shapes
- 📝 **Cell content**: Text in cells is preserved
- 🖥️ **Two GUI Options**: Desktop GUI (Tkinter) or Web GUI
- 📊 **Progress bar**: Real-time progress updates for large file conversions
- 🎛️ **Conversion options**: Include/exclude connectors and cell colors
- 🔧 **Resizable window**: Desktop GUI adjusts to your screen size
- 📜 **Scrollable sheet list**: Easily navigate files with many sheets

## Installation

```bash
git clone https://github.com/yuanyeee/excel-to-drawio.git
cd excel-to-drawio
pip install -r requirements.txt
```

**Requirements:**
- Python 3.8+
- openpyxl >= 3.1.0
- click >= 8.1.0

## Usage

### Desktop GUI (Tkinter) - Recommended

The desktop GUI provides a simple interface with:
- File browser
- Sheet selection with scrollable list
- Conversion options
- Progress tracking
- Log output

```bash
python gui_tkinter.py
```

**Features:**
- Resizable window (default: 900x700, minimum: 700x600)
- Scrollable sheet list for files with many sheets
- Paned window layout for flexible organization
- Native file dialogs

### Web GUI

The web GUI provides a browser-based interface:
- Drag & drop file upload
- Real-time progress streaming
- Modern, responsive design

```bash
python serve.py 8765
```

Then open http://localhost:8765 in your browser.

## GUI Options

| Option | Description | Default |
|--------|-------------|---------|
| Include connectors/lines | Include connector and arrow shapes | ✅ On |
| Include cell background colors | Preserve cell fill colors | ✅ On |

## Command Line Usage

### Basic Usage

```bash
python main.py input.xlsx
```

### Specify Output File

```bash
python main.py input.xlsx -o output.drawio
```

### Convert Specific Sheets

```bash
python main.py input.xlsx --sheets "Sheet1" "Sheet2"
```

### Verbose Mode

```bash
python main.py input.xlsx -v
```

## Supported File Formats

| Format | Extension | Description | Support Level |
|--------|-----------|-------------|---------------|
| Excel Workbook | .xlsx | Default Excel format (2007+) | ✅ Full |
| Excel Workbook | .xls | Legacy Excel format (97-2003) | ✅ Full |
| Excel Macro-Enabled | .xlsm | Excel with VBA macros | ✅ Full (macros ignored) |

## Supported Elements

### Drawing Shapes ✅
| Excel Shape | draw.io Shape |
|------------|---------------|
| Rectangle | Rectangle |
| Rounded Rectangle | Rounded Rectangle |
| Ellipse | Ellipse |
| Diamond | Diamond |
| Text Box | Text |
| Line | Line |
| Arrow | Arrow |
| Connector | Connector |
| And more... | |

### Cell-Based Diagrams ✅
| Excel Feature | draw.io Output |
|--------------|-----------------|
| Merged cells with text | Rectangle with text |
| Cell fill color | Rectangle fill color |
| Cell font (size, bold, color) | Text styling |
| Cell borders (top/bottom/left/right) | Lines |

### ⚠️ Not Yet Supported
- Charts (planned)
- Images (planned)
- SmartArt (planned)

## How Cell-Based Diagrams Work

Excel often uses merged cells with borders to create diagram-like layouts:

```
Excel:
┌───────────┬───────────┐
│  Title    │          │  <- merged cell
├─────┬─────┼───────────┤
│  A  │  B  │     C     │  <- merged cells
└─────┴─────┴───────────┘

draw.io:
┌───────────┬───────────┐
│  Title    │          │
├─────┬─────┼───────────┤
│  A  │  B  │     C     │
└─────┴─────┴───────────┘
```

## Troubleshooting

### Common Issues

**"Failed to load Excel file" error**
- Ensure the file is a valid Excel file (.xlsx, .xls, .xlsm)
- Check that the file is not corrupted or password-protected

**"No sheets selected" but sheets are visible**
- Make sure at least one sheet checkbox is checked
- Click on the checkbox itself (not just the label)

**GUI window is too small/large**
- The window is resizable - drag the edges to adjust
- Default size is 900x700 pixels

**Sheet list doesn't scroll**
- The scrollable frame should automatically show a scrollbar when needed
- Try resizing the window horizontally

**Progress bar doesn't move**
- For large files, conversion may take time
- Check the log area for progress messages

### Getting Help

If you encounter issues:
1. Check the log area for error messages
2. Try running with the web GUI for more detailed error reporting
3. Open an issue on GitHub with the error message and file (if possible)

## Development

### Project Structure

```
excel-to-drawio/
├── main.py                  # CLI entry point
├── serve.py                 # GUI web server
├── gui_tkinter.py           # Desktop GUI (Tkinter)
├── run.bat                  # Windows launcher
├── run.sh                   # Mac/Linux launcher
├── converter/
│   ├── __init__.py
│   ├── excel_reader.py       # Excel shape extraction + cell extraction
│   ├── shape_mapper.py       # Shape type mapping
│   ├── drawio_writer.py      # draw.io XML generation
│   └── cell_border.py        # Cell border extraction
├── tests/
│   └── test_converter.py
├── requirements.txt
├── README.md
├── SPEC.md                  # Detailed specification
└── .gitignore
```

### Running Tests

```bash
python -m pytest tests/
```

## License

MIT License - see LICENSE file for details

## Contributing

Pull requests are welcome! Please read the SPEC.md for design guidelines.
