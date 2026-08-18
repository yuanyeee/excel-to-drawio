"""Grid / cell-coordinate helpers."""
import re
from collections import defaultdict

from .constants import (
    FONT_ALIASES,
    SS,
)

def _emu_px(emu, cfg):
    return emu / cfg.emu_per_px / cfg.scale


def _chars_px(c, cfg):
    """Convert Excel column width (characters) to pixels.

    Uses the OOXML spec formula:
        pixels = Truncate(((256 * width + Truncate(128 / MDW)) / 256) * MDW)
    where MDW is the maximum-digit-width of the workbook's default font in
    pixels (cfg.char_width, defaults to 7 for Calibri 11).
    """
    if c <= 0:
        return 0
    mdw = max(1, cfg.char_width)
    return max(1, int((256 * c + int(128 / mdw)) / 256 * mdw))


def _pts_px(pts, cfg):
    return round(pts * cfg.point_to_px)


def _col_letter_to_idx(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def _cell_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref)
    if not m:
        raise ValueError(f'Invalid cell ref: {ref}')
    return _col_letter_to_idx(m.group(1)), int(m.group(2)) - 1


def _normalize_font_name(name):
    if not name:
        return None
    return FONT_ALIASES.get(name, name)


def _parse_range_ref(ref):
    if ':' not in ref:
        c, r = _cell_ref(ref)
        return c, r, c, r
    m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', ref)
    if not m:
        raise ValueError(ref)
    return (_col_letter_to_idx(m.group(1)), int(m.group(2)) - 1,
            _col_letter_to_idx(m.group(3)), int(m.group(4)) - 1)


def _build_grid(sh_root, cfg):
    """Build pixel coordinate arrays from column widths and row heights.

    The grid is sized dynamically from the workbook's actual extent so that
    sheets with more than 500 rows (or columns) are not silently collapsed
    into the last slot of a fixed-size array.
    """
    ns = {'x': SS}
    # Resolve default column width and row height. Excel stores both on
    # sheetFormatPr; when absent, default to ~9.14 (64 px via the OOXML formula
    # at MDW=7, matching Calibri 11) and 15 pt (the Excel default row height).
    # Honoring defaultRowHeight is critical when a workbook overrides it
    # (e.g. with an alternate font / DPI), otherwise every implicit row drifts
    # by a few pixels and stacks up over hundreds of rows.
    default_col_w = 9.14
    default_row_h = 15.0
    fmt_pr = sh_root.find(f'{{{SS}}}sheetFormatPr')
    if fmt_pr is not None:
        dcw = fmt_pr.attrib.get('defaultColWidth')
        if dcw:
            try:
                default_col_w = float(dcw)
            except ValueError:
                pass
        drh = fmt_pr.attrib.get('defaultRowHeight')
        if drh:
            try:
                default_row_h = float(drh)
            except ValueError:
                pass
    col_w = defaultdict(lambda: default_col_w)
    max_col_seen = 0
    for col_el in sh_root.findall('.//x:col', ns):
        mn = int(col_el.attrib.get('min', 1))
        mx = int(col_el.attrib.get('max', 1))
        w = float(col_el.attrib.get('width', 8))
        hidden = col_el.attrib.get('hidden') == '1'
        for c in range(mn - 1, mx):
            col_w[c] = 0.0 if (hidden and cfg.skip_hidden) else w
        if mx > max_col_seen:
            max_col_seen = mx

    row_h = defaultdict(lambda: default_row_h)
    max_row_seen = 0
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1))
        if r > max_row_seen:
            max_row_seen = r
        ht = row_el.attrib.get('ht')
        hidden = row_el.attrib.get('hidden') == '1'
        if hidden and cfg.skip_hidden:
            row_h[r - 1] = 0.0
        elif ht:
            row_h[r - 1] = float(ht)
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except (ValueError, TypeError):
                continue
            if c + 1 > max_col_seen:
                max_col_seen = c + 1

    MAX = max(500, max_row_seen + 50, max_col_seen + 50)
    col_x = [0] * (MAX + 1)
    for i in range(MAX):
        col_x[i + 1] = col_x[i] + _chars_px(col_w[i], cfg)
    row_y = [0] * (MAX + 1)
    for i in range(MAX):
        row_y[i + 1] = row_y[i] + _pts_px(row_h[i], cfg)
    return col_x, row_y, col_w, row_h


def _auto_detect_bounds(sh_root):
    """Scan actual cell data to find min/max row and column indices."""
    ns = {'x': SS}
    min_r, max_r, min_c, max_c = 9999, 0, 9999, 0
    found = False
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except (ValueError, TypeError):
                continue
            found = True
            min_r, max_r = min(min_r, r), max(max_r, r)
            min_c, max_c = min(min_c, c), max(max_c, c)
    if not found:
        return 0, 0, 0, 0
    return min_r, max_r, min_c, max_c

