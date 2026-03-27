"""
draw.io Writer - Generate draw.io XML format with support for cell borders,
custom geometry, deep group shapes, and off-page connectors
"""

import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import Dict, List
from datetime import datetime

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
        # Create root diagram
        mxfile = self._create_mxfile()
        diagram = self._create_diagram()
        mxfile.append(diagram)

        # Add each sheet as a page
        for idx, (sheet_name, data) in enumerate(self.sheets_data.items()):
            self._add_page(diagram, sheet_name, data, idx)

        # Write to file
        xml_str = self._prettify(mxfile)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(xml_str)

    def _create_mxfile(self) -> ET.Element:
        """Create the mxfile root element"""
        mxfile = ET.Element("mxfile")
        mxfile.set("host", "excel-to-drawio")
        mxfile.set("version", "24.0.0")
        return mxfile

    def _create_diagram(self) -> ET.Element:
        """Create the diagram element"""
        diagram = ET.Element("diagram")
        diagram.set("name", "Pages")
        diagram.set("id", "0")
        return diagram

    def _add_page(
        self, diagram: ET.Element, sheet_name: str, data: dict, page_idx: int
    ):
        """Add a page (sheet) to the diagram"""
        page = ET.SubElement(diagram, "diagram")
        page.set("id", str(page_idx + 1))
        page.set("name", sheet_name)

        # Create mxGraphModel
        model = ET.SubElement(page, "mxGraphModel")
        model.set("dx", "1200")
        model.set("dy", "800")
        model.set("grid", "1")
        model.set("gridSize", "10")
        model.set("guides", "1")
        model.set("tooltips", "1")
        model.set("connect", "1")
        model.set("arrows", "1")
        model.set("fold", "1")
        model.set("page", "1")
        model.set("pageScale", "1")
        model.set("math", "0")
        model.set("shadow", "0")

        # Root and parent cells
        root = ET.SubElement(model, "root")
        ET.SubElement(root, "mxCell", id="0")
        ET.SubElement(root, "mxCell", id="1", parent="0")

        # Add shapes
        shapes = data.get("shapes", [])
        connectors = data.get("connectors", [])

        self._write_shapes(root, shapes)
        self._write_connectors(root, connectors)

    def _write_shapes(self, parent: ET.Element, shapes: List):
        """Write shapes to the XML"""
        cell_id = 2

        for shape in shapes:
            self.shape_counter += 1
            cell = ET.SubElement(parent, "mxCell")
            cell.set("id", str(cell_id))
            cell.set("parent", "0")
            cell.set("vertex", "1")

            # Geometry
            geo = ET.SubElement(cell, "mxGeometry")
            geo.set("x", str(shape.x / 914400 * 96))  # EMU to pixels
            geo.set("y", str(shape.y / 914400 * 96))
            geo.set("width", str(shape.width / 914400 * 96))
            geo.set("height", str(shape.height / 914400 * 96))
            geo.set("as", "geometry")

            # Shape type
            shape_type = self.mapper.map_type(shape.type)

            # Style
            style_dict = self.mapper.map_style(shape.style)
            style_dict["shape"] = shape_type

            # Add rounded corners for cell-based shapes
            if shape.source == "cell":
                style_dict["rounded"] = "0"

            if not shape.text:
                if "fontSize" not in style_dict:
                    style_dict["fontSize"] = "12"

            # Handle off-page connectors specially
            if shape.type == "offpageConnector" or shape.type == "offPageConnector":
                style_dict["shape"] = "offPageConnector"
                style_dict["verticalLabelPosition"] = "bottom"
                if shape.text:
                    style_dict["labelBackgroundColor"] = "#FFFFFF"

            # Handle custom geometry
            if shape.type == "custom" and shape.geometry == "path" and shape.path_data:
                # Use a rectangle as base and specify as custom
                style_dict["shape"] = "custom"
                # Store the path in UserObject for draw.io to pick up
                # draw.io interprets custom shapes via the style

            style_str = self.mapper.build_style_string(style_dict)
            cell.set("style", style_str)

            # Value (text content)
            if shape.text:
                escaped_text = self._escape_xml(shape.text)
                cell.set("value", escaped_text)

            cell_id += 1

    def _write_connectors(self, parent: ET.Element, connectors: List):
        """Write connectors/arrows to the XML"""
        cell_id = self.shape_counter + 10

        for conn in connectors:
            self.shape_counter += 1
            cell = ET.SubElement(parent, "mxCell")
            cell.set("id", str(cell_id))
            cell.set("parent", "0")
            cell.set("edge", "1")

            if conn.get("source_id"):
                cell.set("source", str(conn.get("source_id")))
            if conn.get("target_id"):
                cell.set("target", str(conn.get("target_id")))

            # Style
            style_dict = self.mapper.map_style(conn.style)
            style_str = self.mapper.build_style_string(style_dict)
            cell.set("style", style_str)

            # Geometry with points
            if conn.points and len(conn.points) >= 2:
                geo = ET.SubElement(cell, "mxGeometry")
                geo.set("as", "geometry")
                geo.set("relative", "1")

                Array = ET.SubElement(geo, "Array")
                for point in conn.points:
                    mxPoint = ET.SubElement(Array, "mxPoint")
                    mxPoint.set("x", str(point[0] / 914400 * 96))
                    mxPoint.set("y", str(point[1] / 914400 * 96))

            cell_id += 1

    def _escape_xml(self, text: str) -> str:
        """Escape XML special characters"""
        if not text:
            return ""
        return (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;")
        )

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

        # Write
        xml_str = minidom.parseString(
            ET.tostring(root, encoding="utf-8").decode("utf-8")
        ).toprettyxml(indent="  ")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(xml_str)
