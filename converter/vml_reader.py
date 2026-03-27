"""
VML Reader - Parse VML (Vector Markup Language) shapes from Excel drawings
Handles legacy VML format used alongside DrawingML in Excel worksheets
"""

import re
from typing import List, Dict, Optional, Tuple
from xml.etree import ElementTree as ET

# VML namespaces
VML_NS = {
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "x": "urn:schemas-microsoft-com:office:excel",
    "pvml": "urn:schemas-microsoft-com:office:powerpoint",
}


class VMLShape:
    """Represents a shape parsed from VML"""

    def __init__(
        self,
        shape_id: str,
        name: str,
        type: str,
        left: float,
        top: float,
        width: float,
        height: float,
        fill_color: str = None,
        stroke_color: str = None,
        stroke_width: float = None,
        text: str = "",
        style: dict = None,
        anchor: dict = None,  # Cell anchor info
        object_type: str = None,  # ClientData ObjectType
    ):
        self.shape_id = shape_id
        self.name = name
        self.type = type
        self.left = left
        self.top = top
        self.width = width
        self.height = height
        self.fill_color = fill_color
        self.stroke_color = stroke_color
        self.stroke_width = stroke_width
        self.text = text
        self.style = style or {}
        self.anchor = anchor  # {col, colOff, row, rowOff}
        self.object_type = object_type


class VMLReader:
    """Parse VML shapes from Excel VML files"""

    # VML shape type constants
    SHAPE_TYPES = {
        "_x0000_t202": "note",  # Text box (common for notes)
        "_x0000_t204": "textBox",
        "_x0000_t2": "rectangle",
        "_x0000_t1": "rectangle",
    }

    def __init__(self):
        self.shapes = []

    def read_file(self, filepath: str) -> List[VMLShape]:
        """Read and parse a VML file"""
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()
            return self._parse_vml_root(root)
        except Exception as e:
            # Return empty list on parse failure
            return []

    def read_string(self, content: str) -> List[VMLShape]:
        """Parse VML from string content"""
        try:
            # Handle the double-XML wrapper issue
            content = content.strip()
            if content.startswith('<?xml'):
                # Remove XML declaration and find the actual root
                content = re.sub(r'<\?xml[^?]*\?>', '', content, count=1)
            
            # Fix double-encoded entities
            content = content.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
            
            root = ET.fromstring(content)
            return self._parse_vml_root(root)
        except Exception as e:
            return []

    def _parse_vml_root(self, root: ET.Element) -> List[VMLShape]:
        """Parse the VML root element"""
        shapes = []

        # Find all v:shape elements
        for shape_elem in root.iter():
            if shape_elem.tag.endswith("}shape") or shape_elem.tag == "shape":
                shape = self._parse_shape(shape_elem)
                if shape:
                    shapes.append(shape)

        # Also check for shapetype definitions that might affect parsing
        return shapes

    def _parse_shape(self, elem: ET.Element) -> Optional[VMLShape]:
        """Parse a single VML shape element"""
        try:
            # Get shape ID
            shape_id = elem.get("id", "")
            if not shape_id:
                return None

            # Get shape type
            type_attr = elem.get("type", "")
            shape_type = self.SHAPE_TYPES.get(type_attr, "shape")

            # Get name from o:title or alt
            name = elem.get("title", "") or elem.get("alt", "") or shape_id

            # Parse style attribute
            style_str = elem.get("style", "")
            style_props = self._parse_style(style_str)

            # Get position and size from style
            left = self._parse_style_value(style_props.get("left", "0"), 0)
            top = self._parse_style_value(style_props.get("top", "0"), 0)
            width = self._parse_style_value(style_props.get("width", "0"), 0)
            height = self._parse_style_value(style_props.get("height", "0"), 0)

            # Get fill color
            fill_color = elem.get("fillcolor", None)
            if not fill_color:
                fill_elem = elem.find(".//{urn:schemas-microsoft-com:vml}fill")
                if fill_elem is not None:
                    fill_color = fill_elem.get("color", None)
                    color2 = fill_elem.get("color2", None)
                    if color2:
                        fill_color = color2  # Use color2 for gradient fills

            # Get stroke
            stroke_color = elem.get("strokecolor", None)
            stroke_width = None
            stroke_elem = elem.find(".//{urn:schemas-microsoft-com:vml}stroke")
            if stroke_elem is not None:
                if not stroke_color:
                    stroke_color = stroke_elem.get("color", None)
                stroke_width_str = stroke_elem.get("weight", None)
                if stroke_width_str:
                    stroke_width = self._parse_stroke_weight(stroke_width_str)

            # Get text content from v:textbox
            text = ""
            textbox = elem.find(".//{urn:schemas-microsoft-com:vml}textbox")
            if textbox is not None:
                for t in textbox.iter():
                    if t.text:
                        text += t.text
                    for child in t:
                        if child.text:
                            text += child.text

            # Get ClientData for cell anchoring
            anchor = None
            object_type = None
            client_data = elem.find(".//{urn:schemas-microsoft-com:office:excel}ClientData")
            if client_data is not None:
                object_type = client_data.get("ObjectType", None)
                
                # Parse anchor info
                anchor_elem = client_data.find("{urn:schemas-microsoft-com:office:excel}Anchor")
                if anchor_elem is not None:
                    try:
                        anchor = {
                            "col1": int(anchor_elem.text.split(",")[0].strip()) if anchor_elem.text else 0,
                        }
                        # Format: "col1, colOff1, row1, rowOff1, col2, colOff2, row2, rowOff2"
                        parts = [p.strip() for p in anchor_elem.text.split(",")]
                        if len(parts) >= 8:
                            anchor = {
                                "col1": int(parts[0]),
                                "colOff1": int(parts[1]),
                                "row1": int(parts[2]),
                                "rowOff1": int(parts[3]),
                                "col2": int(parts[4]),
                                "colOff2": int(parts[5]),
                                "row2": int(parts[6]),
                                "rowOff2": int(parts[7]),
                            }
                    except (ValueError, IndexError):
                        anchor = None

            # Determine actual type based on ClientData ObjectType
            if object_type == "Note":
                shape_type = "note"
            elif object_type:
                shape_type = object_type.lower()

            return VMLShape(
                shape_id=shape_id,
                name=name,
                type=shape_type,
                left=left,
                top=top,
                width=width,
                height=height,
                fill_color=fill_color,
                stroke_color=stroke_color,
                stroke_width=stroke_width,
                text=text.strip(),
                style=style_props,
                anchor=anchor,
                object_type=object_type,
            )

        except Exception:
            return None

    def _parse_style(self, style_str: str) -> Dict[str, str]:
        """Parse VML style string into dict"""
        props = {}
        if not style_str:
            return props

        for part in style_str.split(";"):
            part = part.strip()
            if ":" in part:
                key, value = part.split(":", 1)
                props[key.strip()] = value.strip()

        return props

    def _parse_style_value(self, value: str, default: float) -> float:
        """Parse a style value (e.g., '100px', '50pt') to EMUs or pixels"""
        if not value:
            return default

        value = value.strip().lower()

        # Handle percentage (relative to something)
        if value.endswith("%"):
            return default

        # Try to extract numeric value
        match = re.match(r"([0-9.]+)(px|pt|em|in)?", value)
        if match:
            num = float(match.group(1))
            unit = match.group(2) or "px"

            # Convert to pixels (standard for VML)
            if unit == "pt":
                return num * 1.333  # 1pt = 1.333px
            elif unit == "em":
                return num * 16  # approximate
            elif unit == "in":
                return num * 96  # 1in = 96px
            else:
                return num  # px

        return default

    def _parse_stroke_weight(self, weight_str: str) -> Optional[float]:
        """Parse stroke weight (e.g., '1pt', '1.5pt')"""
        if not weight_str:
            return None

        weight_str = weight_str.strip().lower()

        match = re.match(r"([0-9.]+)(pt|px)?", weight_str)
        if match:
            num = float(match.group(1))
            unit = match.group(2) or "pt"
            if unit == "pt":
                return num * 1.333  # Convert to px
            return num

        return None


def convert_vml_shapes_to_shapes(vml_shapes: List[VMLShape], start_id: int = 1) -> List:
    """Convert VMLShape objects to the Shape format used by drawio_writer"""
    from .excel_reader import Shape

    shapes = []
    for idx, vml_shape in enumerate(vml_shapes):
        shape = Shape(
            shape_id=start_id + idx,
            name=vml_shape.name,
            type=vml_shape.type,
            x=vml_shape.left * 914400 / 96,  # px to EMU
            y=vml_shape.top * 914400 / 96,
            width=vml_shape.width * 914400 / 96,
            height=vml_shape.height * 914400 / 96,
            text=vml_shape.text,
            style={
                "fillColor": vml_shape.fill_color,
                "strokeColor": vml_shape.stroke_color,
                "strokeWidth": vml_shape.stroke_width,
            } if vml_shape.fill_color or vml_shape.stroke_color else {},
            source="vml",
        )
        shapes.append(shape)

    return shapes
