#!/usr/bin/env python3
"""
Excel to draw.io Converter
Convert Excel shapes/diagrams to draw.io format
"""

import sys
import argparse
from pathlib import Path

from converter.excel_to_drawio import convert_excel_to_drawio


def main():
    parser = argparse.ArgumentParser(
        description="Convert Excel shapes/diagrams to draw.io format"
    )
    parser.add_argument("input", help="Input Excel file (.xlsx)")
    parser.add_argument(
        "-o", "--output", help="Output draw.io file (default: input.xlsx -> input.drawio)"
    )
    parser.add_argument(
        "-s",
        "--sheets",
        nargs="+",
        help="Specific sheets to convert (default: all sheets)",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Verbose output"
    )
    parser.add_argument(
        "--include-cells",
        action="store_true",
        help="Include cell-based shapes (can add many grid-like objects)",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: File '{input_path}' not found", file=sys.stderr)
        sys.exit(1)

    output_path = Path(args.output) if args.output else input_path.with_suffix(".drawio")

    if args.verbose:
        print(f"Input:  {input_path}")
        print(f"Output: {output_path}")
        print(f"Sheets: {args.sheets or 'all'}")
        print(f"Include cells: {args.include_cells}")

    result = convert_excel_to_drawio(
        input_path=str(input_path),
        output_path=str(output_path),
        sheet_names=args.sheets,
    )

    print(f"\n✅ Success! Created: {output_path}")
    print(f"   {len(result.sheet_names)} sheet(s) converted")


if __name__ == "__main__":
    main()
