"""Top-level conversion orchestration and public API."""
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from .builder import DrawioBuilder
from .colors import load_theme
from .config import ConvertConfig
from .constants import SS
from .grid import _auto_detect_bounds, _build_grid
from .ooxml import (
    _find_paths,
    _load_shared_strings,
    _validate_workbook_suffix,
)
from .shapes import _add_drawing_shapes
from .styles import (
    _add_cell_borders,
    _add_cell_fills,
    _add_cell_labels,
    _parse_cell_borders,
    _parse_cell_number_formats,
    _parse_cell_styles,
    _parse_cell_text_styles,
    _parse_hyperlinks,
)

def _log(msg):
    sys.stdout.buffer.write((msg + '\n').encode('utf-8', errors='replace'))


def list_supported_sheets(input_path):
    _validate_workbook_suffix(input_path)
    with zipfile.ZipFile(input_path, 'r') as z:
        wb = ET.fromstring(z.read('xl/workbook.xml').decode('utf-8'))
        return [sh.attrib.get('name')
                for sh in wb.findall('.//{%s}sheet' % SS)
                if sh.attrib.get('name')]


def suggest_output_path(input_path, sheet_name):
    _validate_workbook_suffix(input_path)
    safe = re.sub(r'[\\/:*?"<>|]+', '_', str(sheet_name)).strip() or 'output'
    return str(Path(input_path).with_name(f'{safe}.drawio'))


def suggest_multi_output_path(input_path):
    _validate_workbook_suffix(input_path)
    return str(Path(input_path).with_suffix('.drawio'))


def _prepare_resources(z, theme, log):
    # The theme is loaded (fresh) by the caller before any color resolution
    # runs, so scheme references like accent2 resolve against this workbook's
    # actual theme instead of the built-in Office defaults.
    return {
        'shared': _load_shared_strings(z),
        'xf_fills': _parse_cell_styles(z, theme, log),
        'xf_borders': _parse_cell_borders(z, theme, log),
        'xf_text_styles': _parse_cell_text_styles(z, theme, log),
        'xf_numfmts': _parse_cell_number_formats(z, log),
    }


def _build_sheet_xml(z, sheet_name, diagram_id, resources, cfg, theme, log):
    sf, drw_path = _find_paths(z, sheet_name)
    log(f"Sheet XML: {sf}")
    log(f"Drawing:   {drw_path or '(none)'}")
    sh_root = ET.fromstring(z.read(sf).decode('utf-8'))
    col_x, row_y, col_w, row_h = _build_grid(sh_root, cfg)
    bounds = _auto_detect_bounds(sh_root)
    log(f"Bounds: rows {bounds[0]}-{bounds[1]}, cols {bounds[2]}-{bounds[3]}")
    hyperlinks = _parse_hyperlinks(z, sf)
    log(f"Hyperlinks: {len(hyperlinks)}")
    bld = DrawioBuilder(diagram_name=sheet_name, page_mode=cfg.page_mode)
    if cfg.render_fills:
        log("Processing fills...")
        fc = _add_cell_fills(sh_root, col_x, row_y, col_w, row_h,
                             resources['xf_fills'], bld, cfg, bounds, log)
        log(f"  Fill rects: {fc}")
    if cfg.render_borders:
        log("Processing borders...")
        bc = _add_cell_borders(sh_root, col_x, row_y, col_w, row_h,
                               resources['xf_borders'], resources['xf_fills'],
                               bld, cfg, bounds)
        log(f"  Border segments: {bc}")
    if drw_path and cfg.render_shapes:
        before = bld._next
        _add_drawing_shapes(z, drw_path, col_x, row_y, bld, cfg, theme)
        log(f"Drawing shapes: {bld._next - before}")
    if cfg.render_labels:
        before = bld._next
        _add_cell_labels(sh_root, col_x, row_y, col_w, row_h,
                         resources['shared'], resources['xf_text_styles'],
                         resources['xf_numfmts'], resources['xf_fills'],
                         bld, cfg, bounds, hyperlinks)
        log(f"Cell labels: {bld._next - before}")
    log(f"Total shapes: {bld._next - 2}")
    return bld.diagram_xml(diagram_id=diagram_id)


def convert_sheets_to_file(input_path, sheet_names, output_path, cfg=None, log_func=None):
    """Convert one or more sheets to a single .drawio file."""
    _validate_workbook_suffix(input_path)
    if cfg is None:
        cfg = ConvertConfig()
    if isinstance(sheet_names, str):
        sheet_names = [sheet_names]
    names = [str(n) for n in sheet_names if str(n).strip()]
    if not names:
        raise ValueError('At least one sheet must be selected')
    log = log_func or _log
    log(f"Opening '{input_path}' ...")
    with zipfile.ZipFile(input_path, 'r') as z:
        theme = load_theme(z)
        resources = _prepare_resources(z, theme, log)
        diagrams = []
        for idx, sn in enumerate(names, start=1):
            log(f"Processing sheet '{sn}' ...")
            diagrams.append(_build_sheet_xml(z, sn, f'd{idx}', resources, cfg, theme, log))
    xml_out = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<mxfile host="ExcelToDrawIO" version="1.0" type="device">\n'
        + ''.join(diagrams)
        + '</mxfile>\n'
    )
    # ``newline='\n'`` keeps the output byte-identical across platforms
    # (Windows text mode would otherwise write CRLF and break golden tests).
    with open(output_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write(xml_out)
    log(f"Written '{output_path}' ({len(xml_out):,} chars)")


def convert(xlsm, sheet=None, out=None, cfg=None, log_func=None):
    """Convert Excel file to Draw.io format."""
    if out is None:
        out = suggest_output_path(xlsm, sheet) if sheet else suggest_multi_output_path(xlsm)
    sheets = [sheet] if sheet else list_supported_sheets(xlsm)
    convert_sheets_to_file(xlsm, sheets, out, cfg=cfg, log_func=log_func)

