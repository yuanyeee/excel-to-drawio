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

    writer = DrawioWriter(sheets_data)
    writer.write(str(output_file))

    return ConversionResult(
        input_path=input_file,
        output_path=output_file,
        sheet_names=list(sheets_data.keys()),
        sheets_data=sheets_data,
    )
