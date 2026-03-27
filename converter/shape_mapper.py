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
    "straightConnector1": "straightConnector",
    "straightConnector": "straightConnector",
    "curvedConnector": "curvedConnector3",
    "arrow": "arrow",
    # Flowchart shapes
    "flowChartProcess": "process",
    "flowChartDecision": "decision",
    "flowChartTerminator": "terminator",
    "flowChartOffpageConnector": "offPageConnector",
    "offpageConnector": "offPageConnector",
    "offPageConnector": "offPageConnector",
    # Text
    "textBox": "text",
    "text": "text",
    # Notes
    "note": "note",
    # Groups
    "group": "group",
    "grpSp": "group",
    # Custom / path-based
    "custom": "custom",
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
            color = excel_style["fillColor"]
            # Handle scheme colors (keep as-is, they'll be resolved later)
            if color:
                drawio_style["fillColor"] = color

        # Stroke color
        if "strokeColor" in excel_style:
            drawio_style["strokeColor"] = excel_style["strokeColor"]

        # Stroke width
        if "strokeWidth" in excel_style:
            width = excel_style["strokeWidth"]
            if isinstance(width, (int, float)):
                drawio_style["strokeWidth"] = width
            else:
                drawio_style["strokeWidth"] = 2

        # Font size
        if "fontSize" in excel_style:
            drawio_style["fontSize"] = excel_style["fontSize"]

        # Font color
        if "fontColor" in excel_style:
            drawio_style["fontColor"] = excel_style["fontColor"]

        # Font style
        if "fontStyle" in excel_style:
            style = excel_style["fontStyle"]
            if style == "bold":
                drawio_style["fontStyle"] = "1"

        # Alignment
        if "align" in excel_style:
            drawio_style["align"] = excel_style["align"]
        if "verticalAlign" in excel_style:
            drawio_style["verticalAlign"] = excel_style["verticalAlign"]

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
        elif shape_type == "rectangle":
            if style.get("rounded") == "0":
                parts.append("rounded=0")
            elif "rounded" in style:
                parts.append(f"rounded={style['rounded']}")

        # Special handling for off-page connectors
        if shape_type == "offPageConnector":
            parts.append("verticalLabelPosition=bottom")
            parts.append("labelBackgroundColor=#FFFFFF")
            parts.append("fontColor=#FF0000")

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
        if "fontStyle" in style:
            parts.append(f'fontStyle={style["fontStyle"]}')

        # Alignment
        if "align" in style:
            parts.append(f'align={style["align"]}')
        if "verticalAlign" in style:
            parts.append(f'verticalAlign={style["verticalAlign"]}')

        return ";".join(parts)

    @staticmethod
    def get_geometry_aspect(shape_type: str, width: float, height: float) -> tuple:
        """
        Return (geometry_name, aspect_ratio) for custom geometry handling.
        For shapes with special aspect ratios.
        """
        return (shape_type, width / height if height != 0 else 1)
