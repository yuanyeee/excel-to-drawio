"""
High-level Excel -> draw.io conversion entrypoint.

This module intentionally implements conversion core logic directly
(workbook XML + drawing XML parsing + draw.io XML generation)
to reduce dependency on ExcelReader/DrawioWriter extraction heuristics.
"""

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.dom import minidom


SS_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

EMU_PER_PX = 9525  # 914400 / 96
DEFAULT_COL_EMU = 922880
DEFAULT_ROW_EMU = 255780

SKIP_FILL_COLORS = {
    "FFFFFF", "FFFFFE", "F2F2F2", "F3F3F3", "EBEBEB", "E7E6E6", "EEECE1",
    "D9D9D9", "BFBFBF", "000000", "0D0D0D",
}


@dataclass
class DrawioShape:
    x: float
    y: float
    width: float
    height: float
    text: str
    style: Dict[str, str]
    shape_type: str = "rectangle"


@dataclass
class DrawioConnector:
    points: List[Tuple[float, float]]
    style: Dict[str, str]


@dataclass
class ConversionResult:
    """Summary of a conversion run."""

    input_path: Path
    output_path: Path
    sheet_names: List[str]
    sheets_data: Dict


def _local(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag


def _to_px(emu: float) -> float:
    return emu / EMU_PER_PX


def _normalize_hex_color(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    v = str(value).strip().upper()
    if v.startswith("#"):
        v = v[1:]
    if len(v) == 8:
        v = v[2:]
    if len(v) != 6 or not re.fullmatch(r"[0-9A-F]{6}", v):
        return None
    return v


def _extract_solid_color(fill_or_line_elem: ET.Element) -> Optional[str]:
    srgb = fill_or_line_elem.find(f"{{{A_NS}}}srgbClr")
    if srgb is None:
        return None
    return _normalize_hex_color(srgb.get("val"))


def _extract_sp_style(sp_pr: Optional[ET.Element]) -> Dict[str, str]:
    style: Dict[str, str] = {}
    if sp_pr is None:
        return style

    solid_fill = sp_pr.find(f"{{{A_NS}}}solidFill")
    if solid_fill is not None:
        fill = _extract_solid_color(solid_fill)
        if fill and fill not in SKIP_FILL_COLORS:
            style["fillColor"] = f"#{fill}"

    ln = sp_pr.find(f"{{{A_NS}}}ln")
    if ln is not None:
        line_fill = ln.find(f"{{{A_NS}}}solidFill")
        if line_fill is not None:
            stroke = _extract_solid_color(line_fill)
            if stroke:
                style["strokeColor"] = f"#{stroke}"
        w = ln.get("w")
        if w:
            try:
                style["strokeWidth"] = str(round(int(w) / 12700, 2))
            except ValueError:
                pass

        tail_end = ln.find(f"{{{A_NS}}}tailEnd")
        if tail_end is not None and tail_end.get("type"):
            style["endArrow"] = tail_end.get("type")

    return style


def _extract_text(sp_elem: ET.Element) -> str:
    tx_body = sp_elem.find(f"{{{XDR_NS}}}txBody")
    if tx_body is None:
        return ""
    return "".join((t.text or "") for t in tx_body.iter(f"{{{A_NS}}}t")).strip()


def _parse_anchor_origin(anchor: ET.Element) -> Tuple[float, float]:
    from_elem = anchor.find(f"{{{XDR_NS}}}from")
    if from_elem is None:
        return 0.0, 0.0
    col = int((from_elem.findtext(f"{{{XDR_NS}}}col") or "0"))
    col_off = int((from_elem.findtext(f"{{{XDR_NS}}}colOff") or "0"))
    row = int((from_elem.findtext(f"{{{XDR_NS}}}row") or "0"))
    row_off = int((from_elem.findtext(f"{{{XDR_NS}}}rowOff") or "0"))
    x = col * DEFAULT_COL_EMU + col_off
    y = row * DEFAULT_ROW_EMU + row_off
    return float(x), float(y)


def _parse_xfrm(sp_pr: Optional[ET.Element], anchor_x: float, anchor_y: float) -> Tuple[float, float, float, float]:
    if sp_pr is None:
        return anchor_x, anchor_y, 100 * EMU_PER_PX, 40 * EMU_PER_PX
    xfrm = sp_pr.find(f"{{{A_NS}}}xfrm")
    if xfrm is None:
        return anchor_x, anchor_y, 100 * EMU_PER_PX, 40 * EMU_PER_PX

    off = xfrm.find(f"{{{A_NS}}}off")
    ext = xfrm.find(f"{{{A_NS}}}ext")
    x = anchor_x + int(off.get("x", "0")) if off is not None else anchor_x
    y = anchor_y + int(off.get("y", "0")) if off is not None else anchor_y
    w = int(ext.get("cx", str(100 * EMU_PER_PX))) if ext is not None else 100 * EMU_PER_PX
    h = int(ext.get("cy", str(40 * EMU_PER_PX))) if ext is not None else 40 * EMU_PER_PX
    return float(x), float(y), float(max(w, 1)), float(max(h, 1))


def _parse_shape(sp_elem: ET.Element, anchor_x: float, anchor_y: float) -> Optional[DrawioShape]:
    sp_pr = sp_elem.find(f"{{{XDR_NS}}}spPr")
    x, y, w, h = _parse_xfrm(sp_pr, anchor_x, anchor_y)
    style = _extract_sp_style(sp_pr)
    text = _extract_text(sp_elem)

    # Keep only meaningful shapes.
    has_fill = bool(style.get("fillColor"))
    has_stroke = bool(style.get("strokeColor"))
    if not text and not has_fill and not has_stroke:
        return None
    if not text and min(_to_px(w), _to_px(h)) < 2:
        return None

    return DrawioShape(
        x=x,
        y=y,
        width=w,
        height=h,
        text=text,
        style=style,
    )


def _parse_connector(cxn_elem: ET.Element, anchor_x: float, anchor_y: float) -> Optional[DrawioConnector]:
    sp_pr = cxn_elem.find(f"{{{XDR_NS}}}spPr")
    x, y, w, h = _parse_xfrm(sp_pr, anchor_x, anchor_y)
    p1 = (x, y)
    p2 = (x + w, y + h)
    style = _extract_sp_style(sp_pr)

    dx = abs(_to_px(p2[0] - p1[0]))
    dy = abs(_to_px(p2[1] - p1[1]))
    length = (dx ** 2 + dy ** 2) ** 0.5
    if length < 8 and style.get("endArrow", "").lower() in ("", "none"):
        return None
    return DrawioConnector(points=[p1, p2], style=style)


def _parse_drawing_xml(content: bytes) -> Tuple[List[DrawioShape], List[DrawioConnector]]:
    root = ET.fromstring(content)
    shapes: List[DrawioShape] = []
    connectors: List[DrawioConnector] = []

    for anchor in root:
        if _local(anchor.tag) not in ("oneCellAnchor", "twoCellAnchor"):
            continue
        anchor_x, anchor_y = _parse_anchor_origin(anchor)
        for child in anchor:
            name = _local(child.tag)
            if name == "sp":
                shape = _parse_shape(child, anchor_x, anchor_y)
                if shape:
                    shapes.append(shape)
            elif name == "cxnSp":
                conn = _parse_connector(child, anchor_x, anchor_y)
                if conn:
                    connectors.append(conn)
            # grpSp/pic are intentionally skipped to avoid noisy artifacts.

    return shapes, connectors


def _find_sheet_targets(zf: zipfile.ZipFile, sheet_names: Optional[List[str]]) -> List[Tuple[str, Optional[str]]]:
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {r.get("Id"): r.get("Target") for r in rels_root}

    selected = set(sheet_names) if sheet_names else None
    targets: List[Tuple[str, Optional[str]]] = []
    for sheet in wb_root.findall(f".//{{{SS_NS}}}sheet"):
        name = sheet.get("name", "")
        if selected is not None and name not in selected:
            continue
        rid = sheet.get(f"{{{REL_NS}}}id")
        target = rel_map.get(rid)
        if target:
            targets.append((name, f"xl/{target.lstrip('/')}"))
    return targets


def _find_drawing_for_sheet(zf: zipfile.ZipFile, sheet_xml_path: str) -> Optional[str]:
    m = re.search(r"sheet(\d+)\.xml$", sheet_xml_path)
    if not m:
        return None
    sheet_no = m.group(1)
    rels_path = f"xl/worksheets/_rels/sheet{sheet_no}.xml.rels"
    if rels_path not in zf.namelist():
        return None

    rels_root = ET.fromstring(zf.read(rels_path))
    for rel in rels_root.findall(f".//{{{PKG_REL_NS}}}Relationship"):
        rel_type = rel.get("Type", "")
        if "drawing" in rel_type and "vml" not in rel_type:
            target = rel.get("Target", "")
            return f"xl/{target.lstrip('/').replace('../', '')}"
    return None


def _build_style(style: Dict[str, str], shape_type: str = "rectangle") -> str:
    pairs = {"shape": shape_type, "whiteSpace": "wrap", "html": "1"}
    pairs.update(style or {})
    return ";".join(f"{k}={v}" for k, v in pairs.items()) + ";"


def _build_drawio_xml(sheets_data: Dict[str, Dict[str, List]]) -> str:
    mxfile = ET.Element("mxfile", host="excel-to-drawio", version="24.0.0")

    for idx, (sheet_name, data) in enumerate(sheets_data.items(), start=1):
        diagram = ET.SubElement(mxfile, "diagram", id=str(idx + 1), name=sheet_name)
        model = ET.SubElement(
            diagram,
            "mxGraphModel",
            dx="1200",
            dy="800",
            grid="1",
            gridSize="10",
            guides="1",
            tooltips="1",
            connect="1",
            arrows="1",
            fold="1",
            page="1",
            pageScale="1",
            math="0",
            shadow="0",
        )
        root = ET.SubElement(model, "root")
        ET.SubElement(root, "mxCell", id="0")
        ET.SubElement(root, "mxCell", id="1", parent="0")

        cell_id = 2
        for shape in data.get("shapes", []):
            cell = ET.SubElement(root, "mxCell", id=str(cell_id), parent="0", vertex="1")
            cell.set("style", _build_style(shape.style, shape.shape_type))
            if shape.text:
                cell.set("value", shape.text)
            geo = ET.SubElement(cell, "mxGeometry")
            geo.set("x", str(_to_px(shape.x)))
            geo.set("y", str(_to_px(shape.y)))
            geo.set("width", str(_to_px(shape.width)))
            geo.set("height", str(_to_px(shape.height)))
            geo.set("as", "geometry")
            cell_id += 1

        for conn in data.get("connectors", []):
            cell = ET.SubElement(root, "mxCell", id=str(cell_id), parent="0", edge="1")
            style = {"endArrow": "none", "html": "1"}
            style.update(conn.style or {})
            cell.set("style", ";".join(f"{k}={v}" for k, v in style.items()) + ";")
            geo = ET.SubElement(cell, "mxGeometry", relative="1")
            geo.set("as", "geometry")
            source = ET.SubElement(geo, "mxPoint")
            source.set("as", "sourcePoint")
            source.set("x", str(_to_px(conn.points[0][0])))
            source.set("y", str(_to_px(conn.points[0][1])))
            target = ET.SubElement(geo, "mxPoint")
            target.set("as", "targetPoint")
            target.set("x", str(_to_px(conn.points[-1][0])))
            target.set("y", str(_to_px(conn.points[-1][1])))
            cell_id += 1

    rough = ET.tostring(mxfile, encoding="utf-8")
    return minidom.parseString(rough).toprettyxml(indent="  ")


def convert_excel_to_drawio(
    input_path: str,
    output_path: str,
    sheet_names: Optional[List[str]] = None,
    include_cells: bool = True,  # kept for API compatibility; not used in direct-core mode.
) -> ConversionResult:
    input_file = Path(input_path)
    output_file = Path(output_path)

    sheets_data: Dict[str, Dict[str, List]] = {}
    with zipfile.ZipFile(input_file, "r") as zf:
        for name, sheet_xml_path in _find_sheet_targets(zf, sheet_names):
            drawing_path = _find_drawing_for_sheet(zf, sheet_xml_path)
            if drawing_path and drawing_path in zf.namelist():
                shapes, connectors = _parse_drawing_xml(zf.read(drawing_path))
            else:
                shapes, connectors = [], []
            sheets_data[name] = {"shapes": shapes, "connectors": connectors, "title": name}

    xml = _build_drawio_xml(sheets_data)
    output_file.write_text(xml, encoding="utf-8")

    return ConversionResult(
        input_path=input_file,
        output_path=output_file,
        sheet_names=list(sheets_data.keys()),
        sheets_data=sheets_data,
    )
