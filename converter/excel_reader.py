"""
Excel Reader - Extract shapes from Excel worksheets including cells and borders
"""

from typing import Optional, List, Dict, Tuple
from openpyxl import load_workbook
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .cell_border import extract_borders, CellGrid


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
        source: str = "shape",  # "shape" or "cell"
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
        self.source = source  # Track whether from drawing shape or cell

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
    """Read Excel files and extract shapes and cell-based diagrams"""

    def __init__(self, filepath: str, sheet_names: list = None, include_cells: bool = True):
        self.filepath = filepath
        self.sheet_names = sheet_names
        self.include_cells = include_cells
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
        shape_id_counter = 1

        # 1. Extract drawing shapes (inserted shapes)
        drawing_shapes, drawing_connectors = self._extract_drawing_shapes(worksheet, shape_id_counter)
        shapes.extend(drawing_shapes)
        connectors.extend(drawing_connectors)
        shape_id_counter += len(drawing_shapes)

        # 2. Extract cell-based shapes (merged cells with content)
        if self.include_cells:
            cell_shapes = self._extract_cell_shapes(worksheet, shape_id_counter)
            shapes.extend(cell_shapes)

        return shapes, connectors

    def _extract_drawing_shapes(self, worksheet, start_id: int = 1):
        """Extract shapes from worksheet drawings"""
        shapes = []
        connectors = []

        # Iterate through all shapes in the worksheet
        for idx, shape in enumerate(worksheet._sheets):
            shape_obj = self._parse_shape(shape, start_id + idx, source="shape")
            if shape_obj:
                shapes.append(shape_obj)

        # Also check for shapes in the worksheet's drawing
        for drawing in worksheet._drawings:
            if isinstance(drawing, SpreadsheetDrawing):
                for shape in drawing.shapes:
                    shape_obj = self._parse_shape(shape, start_id + len(shapes), source="shape")
                    if shape_obj:
                        shapes.append(shape_obj)

        return shapes, connectors

    def _extract_cell_shapes(self, worksheet: Worksheet, start_id: int = 1) -> List[Shape]:
        """Extract shapes from cells (merged cells with content or borders)"""
        shapes = []
        grid = CellGrid(worksheet)
        shape_id = start_id

        # Extract merged cells with content
        for merged_range in worksheet.merged_cells.ranges:
            cell = worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)

            # Check if cell has content
            if cell.value:
                x, y, width, height = grid.get_merged_cell_bounds(
                    merged_range.min_row,
                    merged_range.max_row,
                    merged_range.min_col,
                    merged_range.max_col,
                )

                text = str(cell.value) if cell.value else ""
                style = self._extract_cell_style(cell)

                shapes.append(
                    Shape(
                        shape_id=shape_id,
                        name=f"Cell_{get_column_letter(merged_range.min_col)}{merged_range.min_row}",
                        type="rectangle",
                        x=x,
                        y=y,
                        width=width,
                        height=height,
                        text=text,
                        style=style,
                        source="cell",
                    )
                )
                shape_id += 1

        # Extract cells with fill color
        for row in worksheet.iter_rows():
            for cell in row:
                if not cell.value:
                    continue

                # Skip cells that are part of merged ranges (except top-left)
                merged = grid.is_merged_cell(cell.row, cell.column)
                if merged:
                    min_row, max_row, min_col, max_col = merged
                    if cell.row != min_row or cell.column != min_col:
                        continue

                # Check if cell has fill
                has_fill = False
                if cell.fill and cell.fill.fill_type == "solid":
                    has_fill = True

                if has_fill or cell.value:
                    x, y, width, height = grid.get_cell_position(cell.row, cell.column)
                    text = str(cell.value) if cell.value else ""
                    style = self._extract_cell_style(cell)

                    shapes.append(
                        Shape(
                            shape_id=shape_id,
                            name=f"Cell_{get_column_letter(cell.column)}{cell.row}",
                            type="rectangle",
                            x=x,
                            y=y,
                            width=width,
                            height=height,
                            text=text,
                            style=style,
                            source="cell",
                        )
                    )
                    shape_id += 1

        return shapes

    def _parse_shape(self, shape, shape_id: int, source: str = "shape") -> Optional[Shape]:
        """Parse a single shape object"""
        try:
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
                source=source,
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

    def _extract_cell_style(self, cell) -> dict:
        """Extract style properties from a cell"""
        style = {}

        try:
            # Fill color
            if cell.fill and cell.fill.fill_type == "solid":
                fgColor = cell.fill.fgColor
                if hasattr(fgColor, "rgb") and fgColor.rgb:
                    rgb = fgColor.rgb
                    if len(rgb) == 8:
                        rgb = rgb[2:]  # Remove alpha
                    style["fillColor"] = f"#{rgb.upper()}"
                elif hasattr(fgColor, "theme"):
                    style["fillColor"] = f"theme:{fgColor.theme}"

            # Font
            if cell.font:
                font = cell.font
                if hasattr(font, "size") and font.size:
                    style["fontSize"] = int(font.size)
                if hasattr(font, "bold") and font.bold:
                    style["fontStyle"] = "bold"
                if hasattr(font, "color") and font.color:
                    color = font.color
                    if hasattr(color, "rgb") and color.rgb:
                        rgb = color.rgb
                        if len(rgb) == 8:
                            rgb = rgb[2:]
                        style["fontColor"] = f"#{rgb.upper()}"

            # Alignment
            if cell.alignment:
                align = cell.alignment
                if hasattr(align, "horizontal") and align.horizontal:
                    style["align"] = align.horizontal
                if hasattr(align, "vertical") and align.vertical:
                    style["verticalAlign"] = align.vertical

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
