"""
High-level Excel -> draw.io conversion entrypoint.

This module provides a single function that can be shared by both CLI and GUI.
"""

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

from .excel_reader import ExcelReader
from .drawio_writer import DrawioWriter


@dataclass
class ConversionResult:
    """Summary of a conversion run."""

    input_path: Path
    output_path: Path
    sheet_names: List[str]
    sheets_data: Dict


_DEFAULT_FILL_COLORS = {
    "", "none", "#FFFFFF", "#FFFFFE", "#F2F2F2", "#F3F3F3", "#EBEBEB", "#E7E6E6", "#EEECE1",
}
_LIGHT_STROKE_COLORS = {
    "", "none", "#D9D9D9", "#CCCCCC", "#BFBFBF", "#C0C0C0", "#E0E0E0", "#E6E6E6", "#F0F0F0",
}


def _normalize_style_color(value: Optional[str]) -> str:
    if not value:
        return ""
    return str(value).strip().upper()


def _is_noise_shape(shape) -> bool:
    """Heuristic filter for tiny/blank scaffold shapes and grid-like artifacts."""
    text = (shape.text or "").strip()
    if text:
        return False

    width_px = shape.width / 914400 * 96 if shape.width else 0
    height_px = shape.height / 914400 * 96 if shape.height else 0
    area_px = width_px * height_px

    style = shape.style or {}
    fill = _normalize_style_color(style.get("fillColor"))
    stroke = _normalize_style_color(style.get("strokeColor"))
    is_default_fill = fill in _DEFAULT_FILL_COLORS or fill.startswith("SCHEME:")
    is_light_stroke = stroke in _LIGHT_STROKE_COLORS

    # Very thin or tiny empty boxes/lines are usually noise.
    if min(width_px, height_px) <= 2:
        return True
    if area_px < 40 and is_default_fill:
        return True
    if is_default_fill and is_light_stroke and (width_px < 24 or height_px < 24):
        return True

    return False


def _is_noise_connector(connector) -> bool:
    """Heuristic filter for short/light connectors that look like artifacts."""
    points = connector.points or []
    if len(points) < 2:
        return True

    # Compute simple end-to-end length in px.
    x1, y1 = points[0]
    x2, y2 = points[-1]
    dx_px = abs((x2 - x1) / 914400 * 96)
    dy_px = abs((y2 - y1) / 914400 * 96)
    length_px = (dx_px ** 2 + dy_px ** 2) ** 0.5

    style = connector.style or {}
    end_arrow = str(style.get("endArrow", "")).strip().lower()
    stroke = _normalize_style_color(style.get("strokeColor"))
    is_light_stroke = stroke in _LIGHT_STROKE_COLORS

    if length_px < 8 and end_arrow in ("", "none"):
        return True
    if length_px < 16 and is_light_stroke and end_arrow in ("", "none"):
        return True
    return False


def _filter_noise(sheets_data: Dict) -> Dict:
    filtered = {}
    for sheet_name, data in sheets_data.items():
        shapes = [s for s in data.get("shapes", []) if not _is_noise_shape(s)]
        connectors = [c for c in data.get("connectors", []) if not _is_noise_connector(c)]
        filtered[sheet_name] = {
            **data,
            "shapes": shapes,
            "connectors": connectors,
        }
    return filtered


def convert_excel_to_drawio(
    input_path: str,
    output_path: str,
    sheet_names: Optional[List[str]] = None,
    include_cells: bool = False,
) -> ConversionResult:
    """
    Convert an Excel workbook into a single multi-page draw.io file.
    """
    input_file = Path(input_path)
    output_file = Path(output_path)

    reader = ExcelReader(str(input_file), sheet_names=sheet_names, include_cells=include_cells)
    try:
        sheets_data = reader.read_all()
    finally:
        reader.close()

    sheets_data = _filter_noise(sheets_data)

    writer = DrawioWriter(sheets_data)
    writer.write(str(output_file))

    return ConversionResult(
        input_path=input_file,
        output_path=output_file,
        sheet_names=list(sheets_data.keys()),
        sheets_data=sheets_data,
    )
