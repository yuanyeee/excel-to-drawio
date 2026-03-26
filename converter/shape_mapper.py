"""
Shape Mapper - Map Excel shape types to draw.io shape types
"""

from typing import Optional


# Excel shape types to draw.io shape types
EXCEL_TO_DRAWIO_MAP = {
    # Basic shapes
    "rectangle": "rectangle",
    "roundRectangle": "roundRect",
    "ellipse": "ellipse",
    "diamond": "diamond",
    "parallelogram": "parallelogram",
    "trapezoid": "trapezoid",
    "hexagon": "hexagon",
    "octagon": "octagon",
    "triangle": "triangle",
    "pentagon": "pentagon",
    "cross": "cross",
    "star": "star",
    "heart": "heart",
    "lightningBolt": "lightning",
    "sun": "sun",
    "moon": "moon",
    "cloud": "cloud",
    "arc": "arc",
    "line": "line",
    "bentLine": "bentConnector",
    "straightConnector": "connector",
    "curvedConnector": "curvedConnector",
    "arrow": "arrow",
    # Text box
    "textBox": "text",
    # Group
    "group": "group",
}


class ShapeMapper:
    """Map Excel shapes to draw.io equivalents"""

    @staticmethod
    def map_type(excel_type: str) -> str:
        """
        Map Excel shape type to draw.io shape type
        """
        if not excel_type:
            return "rectangle"

        # Remove namespace prefix if present
        if "." in excel_type:
            excel_type = excel_type.split(".")[-1]

        # Direct lookup
        if excel_type in EXCEL_TO_DRAWIO_MAP:
            return EXCEL_TO_DRAWIO_MAP[excel_type]

        # Try case-insensitive matching
        lower_type = excel_type.lower()
        for key, value in EXCEL_TO_DRAWIO_MAP.items():
            if key.lower() == lower_type:
                return value

        # Default to rectangle
        return "rectangle"

    @staticmethod
    def map_style(excel_style: dict) -> dict:
        """
        Map Excel style to draw.io style
        """
        drawio_style = {}

        # Fill color
        if "fillColor" in excel_style:
            drawio_style["fillColor"] = excel_style["fillColor"]

        # Stroke color
        if "strokeColor" in excel_style:
            drawio_style["strokeColor"] = excel_style["strokeColor"]

        # Stroke width
        if "strokeWidth" in excel_style:
            width = excel_style["strokeWidth"]
            # Convert points to pixels (roughly)
            if isinstance(width, (int, float)):
                drawio_style["strokeWidth"] = width
            else:
                drawio_style["strokeWidth"] = 2

        return drawio_style

    @staticmethod
    def build_style_string(style: dict) -> str:
        """
        Build draw.io style string from dict
        e.g., "rounded=1;whiteSpace=wrap;html=1;fillColor=#4A90D9;strokeColor=#2E5C8A;fontSize=12;"
        """
        parts = []

        # Shape-specific
        shape_type = style.get("shape", "rectangle")
        if shape_type == "roundRect":
            parts.append("rounded=1")

        # Text wrapping
        parts.append("whiteSpace=wrap")

        # HTML support
        parts.append("html=1")

        # Fill
        if "fillColor" in style:
            parts.append(f'fillColor={style["fillColor"]}')

        # Stroke
        if "strokeColor" in style:
            parts.append(f'strokeColor={style["strokeColor"]}')
        if "strokeWidth" in style:
            parts.append(f'strokeWidth={style["strokeWidth"]}')

        # Font
        if "fontSize" in style:
            parts.append(f'fontSize={style["fontSize"]}')
        if "fontColor" in style:
            parts.append(f'fontColor={style["fontColor"]}')

        # Alignment
        if "align" in style:
            parts.append(f'align={style["align"]}')
        if "verticalAlign" in style:
            parts.append(f'verticalAlign={style["verticalAlign"]}')

        return ";".join(parts)
