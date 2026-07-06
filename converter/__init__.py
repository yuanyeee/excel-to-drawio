"""
Excel to draw.io Converter

Supports:
- Drawing shapes (rectangles, ellipses, diamonds, flowchart shapes, etc.)
- Cell background fills (adjacent same-color cells merged into rectangles)
- Cell borders
- Cell text labels with font styling and number formatting
- Group shapes with coordinate transforms
- Connectors and arrows
- Theme / indexed / scheme color resolution
"""

from .excel_to_drawio import convert_excel_to_drawio, ConversionResult

__version__ = "0.3.0"
__all__ = [
    "convert_excel_to_drawio",
    "ConversionResult",
]
