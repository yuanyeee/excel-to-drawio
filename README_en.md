# Excel to Draw.io Converter

[🇯🇵 日本語](README.md) | [🇬🇧 English](README_en.md) | [🇨🇳 简体中文](README_zh.md)

A tool that converts Excel (.xlsx / .xlsm) sheets into Draw.io (.drawio) format.
Both a desktop GUI and a command-line interface (CLI) are provided.

## Features

- Simple workflow: pick a file, select sheets, press Convert
- Multi-sheet conversion
- Drawing shapes: rectangles, ellipses, diamonds, connectors, and more
- Cell fidelity: fills, borders, text labels, and merged cells
- Style preservation: fill color, line color, font style
- Image (pic) embedding via base64 data URIs
- Hyperlinks, hidden row/column skipping, text rotation
- Per-conversion color theme (no global mutable state)

## Requirements

- Python 3.8+
- No third-party dependencies (standard library only)
- tkinter (bundled with standard Python) for the GUI

## Install

```bash
pip install .
```

This exposes two console commands: `excel-to-drawio` (CLI) and
`excel-to-drawio-gui` (GUI). You can also run the package directly without
installing it (`python -m excel_to_drawio`).

## Usage

### Command line

```bash
# Convert every sheet (output name is derived automatically)
python -m excel_to_drawio input.xlsx

# Specify an output file
python -m excel_to_drawio input.xlsx -o output.drawio

# Convert specific sheets
python -m excel_to_drawio input.xlsx -s "Sheet1" "Sheet2"

# List sheets
python -m excel_to_drawio input.xlsx -l
```

| Option | Description |
|---|---|
| `input` (required) | Input Excel file (.xlsx / .xlsm) |
| `-o`, `--output` | Output path (default: `<input>.drawio`) |
| `-s`, `--sheets` | Sheet names to convert (default: all) |
| `-l`, `--list` | List sheets and exit |
| `--no-images` / `--no-borders` / `--no-fills` / `--no-labels` / `--no-shapes` | Disable a rendering pass |
| `--no-merge-fills` | Disable same-color fill merging |
| `--skip-hidden` | Skip hidden rows/columns |
| `--no-page-mode` | Emit `page="0"` instead of `page="1"` |
| `--scale` | Scale factor (default: 1.0) |

### Desktop GUI

```bash
python -m excel_to_drawio.desktop_app
# or:
excel-to-drawio-gui
```

## Project structure

```
excel-to-drawio/
├── excel_to_drawio/          # the Python package
│   ├── __init__.py           # public API
│   ├── __main__.py           # CLI entry point
│   ├── desktop_app.py        # tkinter GUI
│   ├── config.py             # ConvertConfig
│   ├── constants.py          # OOXML namespaces and lookup tables
│   ├── colors.py             # Theme + color resolution
│   ├── grid.py               # cell-coordinate helpers
│   ├── ooxml.py              # low-level OOXML reading
│   ├── geometry.py           # DrawingML geometry helpers
│   ├── builder.py            # Drawio XML builder
│   ├── styles.py             # cell styles / fills / borders / labels
│   ├── images.py             # image extraction
│   ├── connectors.py         # connector rendering
│   ├── shapes.py             # shape rendering
│   └── convert.py            # conversion orchestration
├── pyproject.toml            # packaging + console scripts
├── LICENSE                   # MIT
└── README.md                 # this file
```

## License

MIT — see [LICENSE](LICENSE).
