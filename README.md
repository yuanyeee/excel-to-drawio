# Excel to draw.io Converter

Convert Excel shapes, diagrams, flowcharts, and ER diagrams to draw.io format.

## Features

- 🔄 Convert shapes from multiple Excel sheets to a single draw.io file
- 📑 Each Excel sheet becomes a separate tab/page in draw.io
- 📐 Supports various shape types: rectangles, ellipses, diamonds, connectors, etc.
- 🎨 Preserves basic styling: fill color, stroke color, text
- 📊 **Cell-based diagrams**: Merged cells with borders are converted to shapes
- 📝 **Cell content**: Text in cells is preserved
- ⚡ Simple CLI interface

## Installation

```bash
pip install openpyxl click
```

Or install from source:

```bash
git clone https://github.com/yuanyeee/excel-to-drawio.git
cd excel-to-drawio
pip install -e .
```

## Usage

### Basic Usage

```bash
python main.py input.xlsx
```

This will create `input.drawio` with all sheets converted.

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

The converter:
1. Detects merged cell ranges
2. Extracts cell position, size, and content
3. Converts each cell/merged-range to a draw.io rectangle
4. Preserves fill colors and text styling

## Development

### Setup

```bash
git clone https://github.com/yuanyeee/excel-to-drawio.git
cd excel-to-drawio
pip install -r requirements.txt
```

### Run

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
├── converter/
│   ├── __init__.py
│   ├── excel_reader.py       # Excel shape extraction + cell extraction
│   ├── shape_mapper.py       # Shape type mapping
│   ├── drawio_writer.py      # draw.io XML generation
│   └── cell_border.py        # Cell border extraction
├── tests/
│   └── test_converter.py
├── requirements.txt
├── setup.py
├── README.md
└── .gitignore
```

## License

MIT License

## Contributing

Pull requests are welcome!
