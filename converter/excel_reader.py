"""
Excel Reader - Extract shapes from Excel worksheets
"""

from typing import Optional
from openpyxl import load_workbook
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing


class Shape:
    """Represents a shape extracted from Excel"""

    def __init__(
        self,
        shape_id: int,
        name: str,
        type: str,
        x: float,
        y: float,
        width: float,
        height: float,
        text: str = "",
        style: dict = None,
    ):
        self.id = shape_id
        self.name = name
        self.type = type
        self.x = x  # EMUs
        self.y = y  # EMUs
        self.width = width  # EMUs
        self.height = height  # EMUs
        self.text = text
        self.style = style or {}

    def to_pixels(self):
        """Convert EMUs to pixels (96 DPI)"""
        return {
            "x": self.x / 914400 * 96,
            "y": self.y / 914400 * 96,
            "width": self.width / 914400 * 96,
            "height": self.height / 914400 * 96,
        }


class Connector:
    """Represents a connector/arrow from Excel"""

    def __init__(
        self,
        shape_id: int,
        name: str,
        type: str,
        points: list = None,
        style: dict = None,
    ):
        self.id = shape_id
        self.name = name
        self.type = type
        self.points = points or []  # List of (x, y) tuples in EMUs
        self.style = style or {}


class ExcelReader:
    """Read Excel files and extract shapes"""

    def __init__(self, filepath: str, sheet_names: list = None):
        self.filepath = filepath
        self.sheet_names = sheet_names
        self.wb = load_workbook(filepath, data_only=True)

    def read_all(self):
        """Read all sheets and return their shapes"""
        sheets_data = {}

        target_sheets = (
            self.sheet_names
            if self.sheet_names
            else self.wb.sheetnames
        )

        for sheet_name in target_sheets:
            if sheet_name in self.wb.sheetnames:
                ws = self.wb[sheet_name]
                shapes, connectors = self._extract_shapes(ws)
                sheets_data[sheet_name] = {
                    "shapes": shapes,
                    "connectors": connectors,
                    "title": sheet_name,
                }

        return sheets_data

    def _extract_shapes(self, worksheet):
        """Extract shapes and connectors from a worksheet"""
        shapes = []
        connectors = []

        # Get drawing objects
        if worksheet._charts:
            # TODO: Handle charts (extract as images)
            pass

        # Iterate through all shapes in the worksheet
        for shape in worksheet._sheets:
            shape_obj = self._parse_shape(shape)
            if shape_obj:
                shapes.append(shape_obj)

        # Also check for shapes in the worksheet's drawing
        # (some shapes are stored differently)
        for drawing in worksheet._drawings:
            if isinstance(drawing, SpreadsheetDrawing):
                for shape in drawing.shapes:
                    shape_obj = self._parse_shape(shape)
                    if shape_obj:
                        shapes.append(shape_obj)

        return shapes, connectors

    def _parse_shape(self, shape) -> Optional[Shape]:
        """Parse a single shape object"""
        try:
            shape_id = shape.shape_id if hasattr(shape, "shape_id") else 0
            name = shape.name if hasattr(shape, "name") else ""
            type_ = shape.type if hasattr(shape, "type") else "rectangle"

            # Get position and size
            x = getattr(shape, "x", 0) or 0
            y = getattr(shape, "y", 0) or 0
            width = getattr(shape, "width", 0) or 0
            height = getattr(shape, "height", 0) or 0

            # Get text content
            text = ""
            if hasattr(shape, "text") and shape.text:
                text = str(shape.text)
            elif hasattr(shape, "value") and shape.value:
                text = str(shape.value)

            # Get style
            style = self._extract_style(shape)

            return Shape(
                shape_id=shape_id,
                name=name,
                type=type_,
                x=x,
                y=y,
                width=width,
                height=height,
                text=text,
                style=style,
            )
        except Exception:
            return None

    def _extract_style(self, shape) -> dict:
        """Extract style properties from a shape"""
        style = {}

        try:
            # Fill color
            if hasattr(shape, "fill") and shape.fill:
                fill = shape.fill
                if hasattr(fill, "fgColor") and fill.fgColor:
                    color = fill.fgColor
                    if hasattr(color, "rgb"):
                        style["fillColor"] = self._rgb_to_hex(color.rgb)
                    elif hasattr(color, "theme"):
                        style["fillColor"] = f"theme:{color.theme}"

            # Line/stroke
            if hasattr(shape, "line") and shape.line:
                line = shape.line
                if hasattr(line, "color") and line.color:
                    color = line.color
                    if hasattr(color, "rgb"):
                        style["strokeColor"] = self._rgb_to_hex(color.rgb)
                if hasattr(line, "width"):
                    style["strokeWidth"] = line.width

        except Exception:
            pass

        return style

    def _rgb_to_hex(self, rgb: str) -> str:
        """Convert RGB string to hex color"""
        if not rgb:
            return "#000000"
        # Handle ARGB format (with alpha)
        if len(rgb) == 8:
            rgb = rgb[2:]  # Remove alpha
        if len(rgb) == 6:
            return f"#{rgb.upper()}"
        return "#000000"

    def close(self):
        """Close the workbook"""
        self.wb.close()
