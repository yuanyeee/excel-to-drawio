"""
High-level Excel -> draw.io conversion entrypoint.

This module orchestrates the existing reader/writer pipeline so that:
- input Excel file and target sheet names are supplied via parameters,
- cell-derived objects are included when requested,
- all selected sheets are exported into a single multi-page draw.io file.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

from .drawio_writer import DrawioWriter
from .excel_reader import ExcelReader


@dataclass
class ConversionResult:
    """Summary of a conversion run."""

    input_path: Path
    output_path: Path
    sheet_names: List[str]
    sheets_data: Dict


def convert_excel_to_drawio(
    input_path: str,
    output_path: str,
    sheet_names: Optional[List[str]] = None,
    include_cells: bool = True,
) -> ConversionResult:
    """
    Convert an Excel workbook to a draw.io file.

    Args:
        input_path: Path to source Excel file.
        output_path: Path to output .drawio file.
        sheet_names: Optional list of target sheet names. If omitted, all sheets.
        include_cells: Whether to include cell-based objects (fills/borders/labels).

    Returns:
        ConversionResult with selected sheet names and extracted data.
    """

    input_file = Path(input_path)
    output_file = Path(output_path)

    reader = ExcelReader(
        filepath=str(input_file),
        sheet_names=sheet_names,
        include_cells=include_cells,
    )
    sheets_data = reader.read_all()

    writer = DrawioWriter(sheets_data)
    writer.write(str(output_file))

    return ConversionResult(
        input_path=input_file,
        output_path=output_file,
        sheet_names=list(sheets_data.keys()),
        sheets_data=sheets_data,
    )
