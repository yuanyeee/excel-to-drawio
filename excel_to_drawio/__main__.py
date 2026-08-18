"""Command-line entry point for excel-to-drawio.

Run with:  python -m excel_to_drawio   (or the installed excel-to-drawio
console script).
"""
import argparse
import sys

from . import (
    ConvertConfig,
    convert_sheets_to_file,
    list_supported_sheets,
    suggest_multi_output_path,
)


def build_parser():
    parser = argparse.ArgumentParser(
        prog='excel-to-drawio',
        description='Excel (.xlsx/.xlsm) to Draw.io (.drawio) converter',
    )
    parser.add_argument('input', help='Input Excel file (.xlsx / .xlsm)')
    parser.add_argument('-o', '--output', default=None, help='Output file path')
    parser.add_argument('-s', '--sheets', nargs='+', default=None, help='Sheet names to convert')
    parser.add_argument('-l', '--list', action='store_true', dest='list_sheets', help='List sheets and exit')
    parser.add_argument('--no-images', action='store_true', help='Disable image embedding')
    parser.add_argument('--no-borders', action='store_true', help='Disable border rendering')
    parser.add_argument('--no-fills', action='store_true', help='Disable fill rendering')
    parser.add_argument('--no-labels', action='store_true', help='Disable label rendering')
    parser.add_argument('--no-shapes', action='store_true', help='Disable shape rendering')
    parser.add_argument('--no-merge-fills', action='store_true', help='Disable fill merging')
    parser.add_argument('--skip-hidden', action='store_true', help='Skip hidden rows/columns')
    parser.add_argument('--no-page-mode', action='store_true', help='Emit page="0" instead of page="1"')
    parser.add_argument('--scale', type=float, default=1.0, help='Scale factor (default: 1.0)')
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)

    if args.list_sheets:
        for name in list_supported_sheets(args.input):
            print(name)
        return 0

    cfg = ConvertConfig(
        scale=args.scale,
        embed_images=not args.no_images,
        render_images=not args.no_images,
        render_borders=not args.no_borders,
        render_fills=not args.no_fills,
        render_labels=not args.no_labels,
        render_shapes=not args.no_shapes,
        merge_fills=not args.no_merge_fills,
        skip_hidden=args.skip_hidden,
        page_mode=not args.no_page_mode,
    )
    sheets = args.sheets or list_supported_sheets(args.input)
    output = args.output or suggest_multi_output_path(args.input)
    convert_sheets_to_file(args.input, sheets, output, cfg=cfg)
    return 0


if __name__ == '__main__':
    sys.exit(main())
