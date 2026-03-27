# Excel to draw.io Converter

Convert Excel shapes, diagrams, flowcharts, and ER diagrams to draw.io format.

## Features

- 🔄 Convert shapes from multiple Excel sheets to a single draw.io file
- 📑 Each Excel sheet becomes a separate tab/page in draw.io
- 📐 Supports various shape types: rectangles, ellipses, diamonds, connectors, etc.
- 🎨 Preserves basic styling: fill color, stroke color, text
- 📊 **Cell-based diagrams**: Merged cells with borders are converted to shapes
- 📝 **Cell content**: Text in cells is preserved
- 🖥️ **GUI Mode**: Easy-to-use web interface (no command line needed)
- 📊 **Progress bar**: Real-time progress updates for large file conversions
- 🎛️ **Conversion options**: Output format (draw.io/SVG), include/exclude connectors and cell colors
- 💾 **Settings persistence**: Options saved in localStorage between sessions

## Quick Start - GUI Mode (Recommended)

### Windows / Mac / Linux

1. Download or clone this repository
2. Double-click `run.bat` (Windows) or `run.sh` (Mac/Linux)
3. Open http://localhost:8765 in your browser
4. Drag & drop your Excel file
5. Select sheets to convert
6. Configure options (format, connectors, cell colors)
7. Click "Convert" and download

#### GUI Options

| Option | Description | Default |
|--------|-------------|---------|
| Output Format | draw.io (.drawio) or SVG (.svg) | draw.io |
| Include connectors/lines | Include connector and arrow shapes | ✅ On |
| Include cell background colors | Preserve cell fill colors | ✅ On |

## Command Line Usage

### Installation

```bash
pip install openpyxl click
```

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

## Development

### Setup

```bash
git clone https://github.com/yuanyeee/excel-to-drawio.git
cd excel-to-drawio
pip install -r requirements.txt
```

### Run GUI Server

```bash
python serve.py [port]
# Default port: 8765
# Open http://localhost:8765 in browser
```

> **Note**: The GUI uses Server-Sent Events (SSE) for real-time progress streaming. The `/convert-stream` endpoint provides live conversion progress, while `/convert` remains available for backward compatibility.

### Run CLI

```bash
python main.py input.xlsx
```

### Test

```bash
python -m pytest tests/
```

## Project Structure

```
excel-to-drawio/
├── main.py                  # CLI entry point
├── serve.py                 # GUI web server
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
└── .gitignore
```

## License

MIT License

## Contributing

Pull requests are welcome!
