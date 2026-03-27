"""
Excel to draw.io Converter

Supports:
- Drawing shapes (rectangles, ellipses, diamonds, etc.)
- Cell-based diagrams (merged cells with borders)
- Connectors and arrows
- Basic styling (colors, fonts, alignment)
"""

from .excel_reader import ExcelReader, Shape, Connector
from .drawio_writer import DrawioWriter, SimpleDrawioWriter
from .shape_mapper import ShapeMapper
from .cell_border import CellBorder, CellGrid, extract_borders

__version__ = "0.2.0"
__all__ = [
    "ExcelReader",
    "Shape",
    "Connector",
    "DrawioWriter",
    "SimpleDrawioWriter",
    "ShapeMapper",
    "CellBorder",
    "CellGrid",
    "extract_borders",
]
