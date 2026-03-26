# Excel to draw.io Converter

Convert Excel shapes, diagrams, flowcharts, and ER diagrams to draw.io format.

## Features

- 🔄 Convert shapes from multiple Excel sheets to a single draw.io file
- 📑 Each Excel sheet becomes a separate tab/page in draw.io
- 📐 Supports various shape types: rectangles, ellipses, diamonds, connectors, etc.
- 🎨 Preserves basic styling: fill color, stroke color, text
- ⚡ Simple CLI interface

## Installation

```bash
pip install excel-to-drawio
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
excel-to-drawio input.xlsx
```

This will create `input.drawio` with all sheets converted.

### Specify Output File

```bash
excel-to-drawio input.xlsx -o output.drawio
```

### Convert Specific Sheets

```bash
excel-to-drawio input.xlsx --sheets "Sheet1" "Sheet2"
```

### Verbose Mode

```bash
excel-to-drawio input.xlsx -v
```

## Supported Shapes

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

## Supported Excel Elements

- ✅ Shapes (rectangles, circles, diamonds, etc.)
- ✅ Text boxes
- ✅ Connectors and arrows
- ✅ Basic styling (colors, line width)
- ⚠️ Charts (planned)
- ⚠️ Images (planned)
- ⚠️ SmartArt (planned)

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

## License

MIT License

## Contributing

Pull requests are welcome!
