"""
draw.io Writer - Generate draw.io XML format with support for cell borders,
custom geometry, deep group shapes, and off-page connectors
"""

import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import Dict, List, Tuple

from .shape_mapper import ShapeMapper


class DrawioWriter:
    """Generate draw.io XML files"""

    def __init__(self, sheets_data: Dict, include_borders: bool = True):
        self.sheets_data = sheets_data
        self.mapper = ShapeMapper()
        self.shape_counter = 0
        self.include_borders = include_borders

    def write(self, output_path: str):
        """Write all sheets to a single draw.io file with multiple pages"""
        mxfile = self._create_mxfile()

        for idx, (sheet_name, data) in enumerate(self.sheets_data.items(), start=1):
            self._add_page(mxfile, sheet_name, data, idx)

        xml_str = self._prettify(mxfile)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(xml_str)

    def _create_mxfile(self) -> ET.Element:
        """Create the mxfile root element (Claude script compatible defaults)."""
        mxfile = ET.Element("mxfile")
        mxfile.set("host", "Claude")
        mxfile.set("version", "24.7.5")
        mxfile.set("type", "device")
        return mxfile

    def _calc_page_size(self, shapes: List, connectors: List) -> Tuple[int, int]:
        max_x = 0.0
        max_y = 0.0

        for shape in shapes:
            sx = shape.x / 914400 * 96
            sy = shape.y / 914400 * 96
            sw = shape.width / 914400 * 96
            sh = shape.height / 914400 * 96
            max_x = max(max_x, sx + sw)
            max_y = max(max_y, sy + sh)

        for conn in connectors:
            for px, py in getattr(conn, "points", []) or []:
                cx = px / 914400 * 96
                cy = py / 914400 * 96
                max_x = max(max_x, cx)
                max_y = max(max_y, cy)

        # fromClaude script behavior: at least 2000 with a 10% margin
        page_w = max(2000, int(max_x * 1.10))
        page_h = max(2000, int(max_y * 1.10))
        return page_w, page_h

    def _add_page(self, parent: ET.Element, sheet_name: str, data: dict, page_idx: int):
        """Add a page (sheet) to the diagram"""
        page = ET.SubElement(parent, "diagram")
        page.set("id", f"d{page_idx}")
        page.set("name", sheet_name)

        shapes = data.get("shapes", [])
        connectors = data.get("connectors", [])
        page_w, page_h = self._calc_page_size(shapes, connectors)

        model = ET.SubElement(page, "mxGraphModel")
        model.set("grid", "0")
        model.set("guides", "1")
        model.set("tooltips", "1")
        model.set("connect", "1")
        model.set("arrows", "1")
        model.set("fold", "1")
        model.set("page", "1")
        model.set("pageScale", "1")
        model.set("pageWidth", str(page_w))
        model.set("pageHeight", str(page_h))
        model.set("math", "0")
        model.set("shadow", "0")

        root = ET.SubElement(model, "root")
        ET.SubElement(root, "mxCell", id="0")
        ET.SubElement(root, "mxCell", id="1", parent="0")

        self._write_shapes(root, shapes)
        self._write_connectors(root, connectors)

    def _write_shapes(self, parent: ET.Element, shapes: List):
        """Write shapes to the XML"""
        cell_id = 2

        for shape in shapes:
            self.shape_counter += 1
            cell = ET.SubElement(parent, "mxCell")
            cell.set("id", str(cell_id))
            cell.set("parent", "1")
            cell.set("vertex", "1")

            geo = ET.SubElement(cell, "mxGeometry")
            geo.set("x", str(shape.x / 914400 * 96))
            geo.set("y", str(shape.y / 914400 * 96))
            geo.set("width", str(shape.width / 914400 * 96))
            geo.set("height", str(shape.height / 914400 * 96))
            geo.set("as", "geometry")

            shape_type = self.mapper.map_type(shape.type)

            style_dict = self.mapper.map_style(shape.style)
            style_dict["shape"] = shape_type

            if shape.source == "cell":
                style_dict["rounded"] = "0"

            if not shape.text and "fontSize" not in style_dict:
                style_dict["fontSize"] = "12"

            if shape.type in ("offpageConnector", "offPageConnector"):
                style_dict["shape"] = "offPageConnector"
                style_dict["verticalLabelPosition"] = "bottom"
                if shape.text:
                    style_dict["labelBackgroundColor"] = "#FFFFFF"

            if shape.type == "custom" and shape.geometry == "path" and shape.path_data:
                style_dict["shape"] = "custom"

            style_str = self.mapper.build_style_string(style_dict)
            cell.set("style", style_str)

            if shape.text:
                cell.set("value", shape.text)

            cell_id += 1

    def _write_connectors(self, parent: ET.Element, connectors: List):
        """Write connectors/arrows to the XML"""
        cell_id = self.shape_counter + 10

        for conn in connectors:
            self.shape_counter += 1
            cell = ET.SubElement(parent, "mxCell")
            cell.set("id", str(cell_id))
            cell.set("parent", "1")
            cell.set("edge", "1")

            if hasattr(conn, "source_id") and conn.source_id:
                cell.set("source", str(conn.source_id))
            if hasattr(conn, "target_id") and conn.target_id:
                cell.set("target", str(conn.target_id))

            style_dict = self.mapper.map_style(conn.style)
            style_str = self.mapper.build_style_string(style_dict)
            cell.set("style", style_str)

            if conn.points and len(conn.points) >= 2:
                geo = ET.SubElement(cell, "mxGeometry")
                geo.set("as", "geometry")
                geo.set("relative", "1")

                source_point = ET.SubElement(geo, "mxPoint")
                source_point.set("x", str(conn.points[0][0] / 914400 * 96))
                source_point.set("y", str(conn.points[0][1] / 914400 * 96))
                source_point.set("as", "sourcePoint")

                target_point = ET.SubElement(geo, "mxPoint")
                target_point.set("x", str(conn.points[-1][0] / 914400 * 96))
                target_point.set("y", str(conn.points[-1][1] / 914400 * 96))
                target_point.set("as", "targetPoint")

                if len(conn.points) > 2:
                    points_array = ET.SubElement(geo, "Array")
                    points_array.set("as", "points")
                    for point in conn.points[1:-1]:
                        mx_point = ET.SubElement(points_array, "mxPoint")
                        mx_point.set("x", str(point[0] / 914400 * 96))
                        mx_point.set("y", str(point[1] / 914400 * 96))

            cell_id += 1

    def _prettify(self, elem: ET.Element) -> str:
        """Return a pretty-printed XML string"""
        rough_string = ET.tostring(elem, encoding="utf-8").decode("utf-8")
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")


class SimpleDrawioWriter:
    """Simplified draw.io writer for single-sheet output"""

    def __init__(self, shapes: List, title: str = "Sheet"):
        self.shapes = shapes
        self.title = title
        self.mapper = ShapeMapper()

    def write(self, output_path: str):
        """Write to a draw.io file"""
        root = ET.Element("mxfile")
        root.set("host", "excel-to-drawio")
        root.set("version", "24.0.0")

        diagram = ET.SubElement(root, "diagram")
        diagram.set("name", self.title)
        diagram.set("id", "0")

        model = ET.SubElement(diagram, "mxGraphModel")
        model.set("dx", "1200")
        model.set("dy", "800")
        model.set("grid", "1")
        model.set("gridSize", "10")
        model.set("guides", "1")
        model.set("math", "0")

        root_cell = ET.SubElement(model, "root")
        ET.SubElement(root_cell, "mxCell", id="0")
        ET.SubElement(root_cell, "mxCell", id="1", parent="0")

        cell_id = 2
        for shape in self.shapes:
            cell = ET.SubElement(root_cell, "mxCell")
            cell.set("id", str(cell_id))
            cell.set("parent", "1")
            cell.set("vertex", "1")

            geo = ET.SubElement(cell, "mxGeometry")
            geo.set("x", str(shape.x / 914400 * 96))
            geo.set("y", str(shape.y / 914400 * 96))
            geo.set("width", str(shape.width / 914400 * 96))
            geo.set("height", str(shape.height / 914400 * 96))
            geo.set("as", "geometry")

            shape_type = self.mapper.map_type(shape.type)
            style_dict = self.mapper.map_style(shape.style)
            style_dict["shape"] = shape_type
            cell.set("style", self.mapper.build_style_string(style_dict))

            if shape.text:
                cell.set("value", shape.text)

            cell_id += 1

        xml_str = minidom.parseString(ET.tostring(root, encoding="utf-8")).toprettyxml(indent="  ")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(xml_str)
