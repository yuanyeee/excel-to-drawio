"""
Excel to Draw.io Converter Plus - General Purpose Edition

Excel (.xlsx / .xlsm) to Draw.io (.drawio) converter.
Improvements over the original ExcelToDrawIO:
  - Image (pic) embedding via base64 data URI
  - Auto-detected sheet dimensions (no hardcoded limits)
  - Full border styles (dashed, dotted, double, hair)
  - Hyperlink support on cells
  - Hidden row/column skipping option
  - Text rotation, underline, strikethrough support
  - Config dataclass (no global variables)
  - No flowchart-specific heuristics (general purpose)
"""

import base64
import dataclasses
import html
import mimetypes
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from math import ceil
from pathlib import Path

# ======================================================================
#  XML Namespaces
# ======================================================================
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ASVG = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
A14 = 'http://schemas.microsoft.com/office/drawing/2010/main'

# Image formats that browsers (and therefore drawio) can render as a data URI.
RENDERABLE_IMG_EXTS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'webp'}
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

# DrawingML preset dash patterns -> drawio dashPattern strings.
PRST_DASH_MAP = {
    'solid': None,
    'dash': '8 4',
    'dot': '1 4',
    'dashDot': '8 4 1 4',
    'lgDash': '16 4',
    'sysDash': '5 2',
    'sysDot': '1 2',
    'lgDashDot': '16 4 1 4',
    'lgDashDotDot': '16 4 1 4 1 4',
    'dashDotDot': '8 4 1 4 1 4',
    'sysDashDot': '5 2 1 2',
    'sysDashDotDot': '5 2 1 2 1 2',
}

# DrawingML arrow head type -> drawio arrow style.
ARROW_MAP = {
    'triangle': 'classic',
    'arrow': 'classic',
    'stealth': 'classicThin',
    'diamond': 'diamondThin',
    'oval': 'oval',
    'none': 'none',
}

# ======================================================================
#  Color Tables
# ======================================================================
# DrawingML scheme colors. Keys MUST match the values used in OOXML
# ``<a:schemeClr val="..."/>`` (e.g. ``accent1``). Legacy short aliases
# (``acc1`` etc.) are kept for backwards compatibility with older callers.
# ``_load_theme_colors`` mutates this dict in place when an .xlsx workbook
# ships its own ``xl/theme/theme1.xml`` so per-workbook color schemes are
# honored.
SCHEME_COLORS = {
    'dk1': '000000', 'lt1': 'FFFFFF', 'dk2': '44546A', 'lt2': 'E7E6E6',
    'accent1': '4472C4', 'accent2': 'ED7D31', 'accent3': 'A5A5A5',
    'accent4': 'FFC000', 'accent5': '5B9BD5', 'accent6': '70AD47',
    'hlink': '0563C1', 'folHlink': '954F72',
    'bg1': 'FFFFFF', 'bg2': 'E7E6E6', 'tx1': '000000', 'tx2': '44546A',
    'phClr': 'FFFFFF',
    # Legacy short aliases — keep for backwards compatibility.
    'acc1': '4472C4', 'acc2': 'ED7D31', 'acc3': 'A5A5A5',
    'acc4': 'FFC000', 'acc5': '5B9BD5', 'acc6': '70AD47',
}

# Excel theme color index order (used by <fgColor theme="N"/>):
# 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4-9=accent1..6, 10=hlink, 11=folHlink.
# (Note: Excel swaps lt1<->dk1 and lt2<->dk2 vs. the natural OOXML order.)
THEME_INDEX_NAMES = [
    'lt1', 'dk1', 'lt2', 'dk2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink', 'folHlink',
]

THEME_FILL_COLORS = [SCHEME_COLORS[n] for n in THEME_INDEX_NAMES]

INDEXED_COLORS = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
    'FFFFFF', 'FFFFFF',
]

SKIP_COLORS = {
    'FFFFFF', 'FFFFFE', 'F2F2F2', 'F3F3F3', 'EBEBEB', 'E7E6E6', 'EEECE1',
    'D9D9D9', 'BFBFBF', '000000', '0D0D0D',
}

# ======================================================================
#  Shape -> DrawIO Style Mapping (extended)
# ======================================================================
GEOM_STYLES = {
    'rect': '',
    'roundRect': 'rounded=1;arcSize=10;',
    'ellipse': 'ellipse;',
    'diamond': 'rhombus;',
    'triangle': 'triangle;',
    'parallelogram': 'parallelogram;',
    'trapezoid': 'trapezoid;',
    'hexagon': 'hexagon;',
    'octagon': 'octagon;',
    'star5': 'shape=mxgraph.basic.star;',
    'cloud': 'shape=cloud;',
    'heart': 'shape=mxgraph.basic.heart;',
    'can': 'shape=cylinder3;',
    'cube': 'shape=cube;',
    'bevel': 'shape=mxgraph.basic.rounded_frame;',
    'donut': 'shape=mxgraph.basic.donut;',
    'noSmoking': 'shape=mxgraph.basic.no_symbol;',
    'blockArc': 'shape=mxgraph.basic.arc;',
    'foldedCorner': 'shape=note;',
    'frame': 'shape=mxgraph.basic.frame;',
    'plaque': 'shape=mxgraph.basic.plaque;',
    # Flowchart
    'flowChartProcess': 'shape=mxgraph.flowchart.process;',
    'flowChartDecision': 'shape=mxgraph.flowchart.decision;',
    'flowChartTerminator': 'shape=mxgraph.flowchart.terminator;',
    'flowChartManualInput': 'shape=mxgraph.flowchart.manual_input;',
    'flowChartDocument': 'shape=mxgraph.flowchart.document;',
    'flowChartPredefinedProcess': 'shape=mxgraph.flowchart.predefined_process;',
    'flowChartConnector': 'ellipse;',
    'flowChartOffpageConnector': 'shape=offPageConnector;',
    'flowChartPunchedTape': 'shape=mxgraph.flowchart.punched_tape;',
    'flowChartSort': 'shape=mxgraph.flowchart.sort;',
    'flowChartPreparation': 'shape=mxgraph.flowchart.preparation;',
    'flowChartManualOperation': 'shape=mxgraph.flowchart.manual_operation;',
    'flowChartMerge': 'shape=mxgraph.flowchart.merge;',
    'flowChartInternalStorage': 'shape=mxgraph.flowchart.internal_storage;',
    'flowChartDelay': 'shape=mxgraph.flowchart.delay;',
    'flowChartAlternateProcess': 'rounded=1;',
    'flowChartMultidocument': 'shape=mxgraph.flowchart.multi-document;',
    'flowChartDisplay': 'shape=mxgraph.flowchart.display;',
    # Pentagon / HomePlate
    'homePlate': 'shape=offPageConnector;',
    'pentagon': 'shape=offPageConnector;',
    # Callouts
    'wedgeRoundRectCallout': 'shape=callout;rounded=1;',
    'wedgeRectCallout': 'shape=callout;',
    'cloudCallout': 'shape=callout;rounded=1;',
    # Arrows
    'bentArrow': 'shape=mxgraph.arrows2.bent_arrow;',
    'chevron': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=20;notch=0;',
    'rightArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=east;',
    'leftArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=west;',
    'upArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=north;',
    'downArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=south;',
    'leftRightArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;',
    'upDownArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;',
    'notchedRightArrow': 'shape=mxgraph.arrows2.notched_arrow;',
    'stripedRightArrow': 'shape=mxgraph.arrows2.striped_arrow;',
}

FONT_ALIASES = {
    '\uff2d\uff33 \u30b4\u30b7\u30c3\u30af': 'MS PGothic',
    '\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af': 'MS PGothic',
    'MS Gothic': 'MS PGothic',
    'MS PGothic': 'MS PGothic',
    '\uff2d\uff33 \u660e\u671d': 'MS PMincho',
    '\uff2d\uff33 \uff30\u660e\u671d': 'MS PMincho',
    '\u6e38\u30b4\u30b7\u30c3\u30af': 'Yu Gothic',
    '\u6e38\u30b4\u30b7\u30c3\u30af Light': 'Yu Gothic Light',
    '\u6e38\u660e\u671d': 'Yu Mincho',
    '\u30e1\u30a4\u30ea\u30aa': 'Meiryo',
    'Meiryo': 'Meiryo',
}

# Excel border style -> (DrawIO strokeWidth, dashPattern or None)
BORDER_STYLE_MAP = {
    'thin': (1, None),
    'medium': (2, None),
    'thick': (3, None),
    'hair': (0.5, None),
    'dashed': (1, '8 8'),
    'mediumDashed': (2, '8 8'),
    'dotted': (1, '2 2'),
    'dashDot': (1, '8 4 2 4'),
    'mediumDashDot': (2, '8 4 2 4'),
    'dashDotDot': (1, '8 4 2 4 2 4'),
    'mediumDashDotDot': (2, '8 4 2 4 2 4'),
    'slantDashDot': (2, '8 4 2 4'),
    'double': (1, None),  # rendered as 2 lines in add_cell_borders
}

# ======================================================================
#  Configuration
# ======================================================================
@dataclasses.dataclass
class ConvertConfig:
    """Conversion settings. All fields have sensible defaults."""
    scale: float = 1.0
    char_width: int = 7
    point_to_px: float = 96 / 72
    emu_per_px: int = 9525
    embed_images: bool = True
    skip_hidden: bool = False
    merge_fills: bool = True
    render_borders: bool = True
    render_fills: bool = True
    render_labels: bool = True
    render_shapes: bool = True
    render_images: bool = True

# ======================================================================
#  Utilities
# ======================================================================
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

def _apply_tint(hex6, tint):
    """Apply Excel cell-color ``tint`` attribute (-1.0 to 1.0).

    >0: blend toward white. <0: blend toward black. Used by xf cell colors
    (``<color theme="1" tint="-0.25"/>``).
    """
    try:
        r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
        t = float(tint)
        if t > 0:
            r, g, b = int(r + (255 - r) * t), int(g + (255 - g) * t), int(b + (255 - b) * t)
        else:
            r, g, b = int(r * (1 + t)), int(g * (1 + t)), int(b * (1 + t))
        r, g, b = max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b))
        return f'{r:02X}{g:02X}{b:02X}'
    except Exception:
        return hex6


def _rgb_to_hsl(r, g, b):
    """Convert 0-255 RGB to (H, S, L) in [0, 1]."""
    rf, gf, bf = r / 255.0, g / 255.0, b / 255.0
    mx, mn = max(rf, gf, bf), min(rf, gf, bf)
    l = (mx + mn) / 2.0
    if mx == mn:
        return 0.0, 0.0, l
    d = mx - mn
    s = d / (2.0 - mx - mn) if l > 0.5 else d / (mx + mn)
    if mx == rf:
        h = (gf - bf) / d + (6.0 if gf < bf else 0.0)
    elif mx == gf:
        h = (bf - rf) / d + 2.0
    else:
        h = (rf - gf) / d + 4.0
    return h / 6.0, s, l


def _hsl_to_rgb(h, s, l):
    """Convert (H, S, L) in [0, 1] back to 0-255 RGB tuple."""
    if s == 0:
        v = int(round(l * 255))
        return v, v, v

    def _hue(p, q, t):
        if t < 0:
            t += 1
        if t > 1:
            t -= 1
        if t < 1 / 6:
            return p + (q - p) * 6 * t
        if t < 1 / 2:
            return q
        if t < 2 / 3:
            return p + (q - p) * (2 / 3 - t) * 6
        return p

    q = l * (1 + s) if l < 0.5 else l + s - l * s
    p = 2 * l - q
    r = _hue(p, q, h + 1 / 3)
    g = _hue(p, q, h)
    b = _hue(p, q, h - 1 / 3)
    return (int(round(r * 255)),
            int(round(g * 255)),
            int(round(b * 255)))


def _apply_lum_mod_off(hex6, lum_mod, lum_off):
    """Apply DrawingML ``lumMod``/``lumOff`` via HSL luminance scaling.

    OOXML defines: ``L_new = L_old * lumMod + lumOff`` where lumMod / lumOff
    are 0..1 (read as the raw int / 100000). This is the correct algorithm —
    the older ``_apply_tint``-based shortcut produced visibly wrong results
    (e.g. accent2 with lumMod=75000/lumOff=0 came out gray instead of darker
    orange).
    """
    try:
        r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
        h, s, l = _rgb_to_hsl(r, g, b)
        l = max(0.0, min(1.0, l * lum_mod + lum_off))
        r, g, b = _hsl_to_rgb(h, s, l)
        return f'{r:02X}{g:02X}{b:02X}'
    except Exception:
        return hex6


def _apply_color_modifiers(hex6, parent):
    """Apply DrawingML color child modifiers (lumMod/lumOff/tint/shade)."""
    if parent is None:
        return hex6
    lm_el = parent.find(f'{{{A}}}lumMod')
    lo_el = parent.find(f'{{{A}}}lumOff')
    if lm_el is not None or lo_el is not None:
        try:
            lm = int(lm_el.attrib.get('val', '100000')) / 100000.0 if lm_el is not None else 1.0
            lo = int(lo_el.attrib.get('val', '0')) / 100000.0 if lo_el is not None else 0.0
            hex6 = _apply_lum_mod_off(hex6, lm, lo)
        except (TypeError, ValueError):
            pass
    tint_el = parent.find(f'{{{A}}}tint')
    if tint_el is not None:
        try:
            t = int(tint_el.attrib.get('val', '0')) / 100000.0
            # DrawingML tint is positive (0..1) and lightens toward white.
            hex6 = _apply_tint(hex6, t)
        except (TypeError, ValueError):
            pass
    shade_el = parent.find(f'{{{A}}}shade')
    if shade_el is not None:
        try:
            sval = int(shade_el.attrib.get('val', '100000')) / 100000.0
            r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
            r = max(0, min(255, int(r * sval)))
            g = max(0, min(255, int(g * sval)))
            b = max(0, min(255, int(b * sval)))
            hex6 = f'{r:02X}{g:02X}{b:02X}'
        except (TypeError, ValueError):
            pass
    return hex6


def _load_theme_colors(z):
    """Populate ``SCHEME_COLORS`` / ``THEME_FILL_COLORS`` from ``xl/theme/theme1.xml``.

    DrawingML stores 12 scheme colors in ``<a:clrScheme>``: dk1, lt1, dk2, lt2,
    accent1..accent6, hlink, folHlink. Each child wraps either ``<a:srgbClr/>``
    or ``<a:sysClr lastClr="..."/>``. We mutate the module-level dicts in place
    so every subsequent ``_parse_drawing_color`` / ``_parse_color_el`` call uses
    the workbook's actual scheme. Failing this step silently leaves the defaults
    untouched.
    """
    candidates = [n for n in z.namelist()
                  if n.startswith('xl/theme/theme') and n.endswith('.xml')]
    if not candidates:
        return
    try:
        root = ET.fromstring(z.read(sorted(candidates)[0]).decode('utf-8'))
    except Exception:
        return
    scheme = root.find(f'.//{{{A}}}clrScheme')
    if scheme is None:
        return
    for child in scheme:
        name = child.tag.split('}')[-1]
        srgb = child.find(f'{{{A}}}srgbClr')
        sys_el = child.find(f'{{{A}}}sysClr')
        hex6 = None
        if srgb is not None:
            hex6 = srgb.attrib.get('val', '').upper()
        elif sys_el is not None:
            hex6 = (sys_el.attrib.get('lastClr') or '').upper()
        if hex6 and len(hex6) == 6:
            SCHEME_COLORS[name] = hex6
            if name.startswith('accent') and name[-1].isdigit():
                SCHEME_COLORS['acc' + name[-1]] = hex6
    # Refresh aliases that mirror the primary names.
    SCHEME_COLORS['bg1'] = SCHEME_COLORS.get('lt1', SCHEME_COLORS['bg1'])
    SCHEME_COLORS['bg2'] = SCHEME_COLORS.get('lt2', SCHEME_COLORS['bg2'])
    SCHEME_COLORS['tx1'] = SCHEME_COLORS.get('dk1', SCHEME_COLORS['tx1'])
    SCHEME_COLORS['tx2'] = SCHEME_COLORS.get('dk2', SCHEME_COLORS['tx2'])
    # Refresh positional theme list used by cell fills.
    THEME_FILL_COLORS[:] = [SCHEME_COLORS[n] for n in THEME_INDEX_NAMES]

def _parse_color_el(color_el, default='#000000'):
    """Parse fgColor/bgColor/color element to '#RRGGBB' with tint correction."""
    if color_el is None:
        return default
    rgb = color_el.attrib.get('rgb', '')
    if rgb:
        h6 = (rgb[2:] if len(rgb) == 8 else rgb[:6]).upper()
        tint = color_el.attrib.get('tint', '')
        if tint:
            h6 = _apply_tint(h6, tint)
        return '#' + h6
    theme = color_el.attrib.get('theme', '')
    if theme:
        idx = int(theme)
        base = THEME_FILL_COLORS[idx] if idx < len(THEME_FILL_COLORS) else None
        if base:
            tint = color_el.attrib.get('tint', '')
            if tint:
                base = _apply_tint(base, tint)
            return '#' + base
    indexed = color_el.attrib.get('indexed', '')
    if indexed:
        idx = int(indexed)
        if idx == 64:
            return default
        icolor = INDEXED_COLORS[idx] if idx < len(INDEXED_COLORS) else None
        if icolor:
            return '#' + icolor
    return default

def _should_skip_fill(hex6):
    return hex6.upper().lstrip('#') in SKIP_COLORS

def _parse_range_ref(ref):
    if ':' not in ref:
        c, r = _cell_ref(ref)
        return c, r, c, r
    m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', ref)
    if not m:
        raise ValueError(ref)
    return (_col_letter_to_idx(m.group(1)), int(m.group(2)) - 1,
            _col_letter_to_idx(m.group(3)), int(m.group(4)) - 1)

def _log(msg):
    sys.stdout.buffer.write((msg + '\n').encode('utf-8', errors='replace'))


# ======================================================================
#  Grid Builder
# ======================================================================
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
            except Exception:
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
            except Exception:
                continue
            found = True
            min_r, max_r = min(min_r, r), max(max_r, r)
            min_c, max_c = min(min_c, c), max(max_c, c)
    if not found:
        return 0, 0, 0, 0
    return min_r, max_r, min_c, max_c


# ======================================================================
#  DrawIO XML Builder
# ======================================================================
class DrawioBuilder:
    def __init__(self, diagram_name='Sheet1'):
        self._cells = []
        self._next = 2
        self._seen = set()
        self._max_x = 0
        self._max_y = 0
        self._diagram_name = diagram_name

    def add(self, text, x, y, w, h, style, force=False):
        x, y = round(x), round(y)
        w, h = round(max(w, 1)), round(max(h, 1))
        key = (x, y, w, h, style[:60])
        if key in self._seen and not force:
            return
        self._seen.add(key)
        self._max_x = max(self._max_x, x + w)
        self._max_y = max(self._max_y, y + h)
        cid = self._next
        self._next += 1
        esc = html.escape(str(text))
        self._cells.append(
            f'    <mxCell id="{cid}" value="{esc}" style="{style}" vertex="1" parent="1">'
            f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/>'
            f'</mxCell>'
        )

    def add_image(self, x, y, w, h, data_uri, extra_style=None):
        """Add an embedded image as a DrawIO image shape.

        ``extra_style`` may be a string of already-joined style fragments (no
        leading/trailing ``;``) such as ``"rotation=90;flipH=1"`` that will be
        appended after the default image style.
        """
        x, y = round(x), round(y)
        w, h = round(max(w, 1)), round(max(h, 1))
        self._max_x = max(self._max_x, x + w)
        self._max_y = max(self._max_y, y + h)
        cid = self._next
        self._next += 1
        style = (f'shape=image;verticalLabelPosition=bottom;labelBackgroundColor=default;'
                 f'verticalAlign=top;aspect=fixed;imageAspect=0;'
                 f'image={data_uri};')
        if extra_style:
            style += extra_style.strip(';') + ';'
        self._cells.append(
            f'    <mxCell id="{cid}" value="" style="{style}" vertex="1" parent="1">'
            f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/>'
            f'</mxCell>'
        )

    def add_edge(self, x1, y1, x2, y2, style, points=None):
        """Add a drawio edge (line) between two explicit points.

        Used for connector shapes (``xdr:cxnSp``) so they render as real lines
        instead of collapsed vertex rectangles. The edge has no source/target
        cell — drawio uses the explicit ``sourcePoint``/``targetPoint`` mxPoints
        in the geometry.
        """
        x1, y1 = round(x1), round(y1)
        x2, y2 = round(x2), round(y2)
        self._max_x = max(self._max_x, x1, x2)
        self._max_y = max(self._max_y, y1, y2)
        cid = self._next
        self._next += 1
        points_xml = ''
        if points:
            pts = []
            for px, py in points:
                pts.append(f'<mxPoint x="{round(px)}" y="{round(py)}"/>')
                self._max_x = max(self._max_x, round(px))
                self._max_y = max(self._max_y, round(py))
            points_xml = f'<Array as="points">{"".join(pts)}</Array>'
        self._cells.append(
            f'    <mxCell id="{cid}" value="" style="{style}" edge="1" parent="1">'
            f'<mxGeometry relative="1" as="geometry">'
            f'<mxPoint x="{x1}" y="{y1}" as="sourcePoint"/>'
            f'<mxPoint x="{x2}" y="{y2}" as="targetPoint"/>'
            f'{points_xml}'
            f'</mxGeometry>'
            f'</mxCell>'
        )

    def diagram_xml(self, diagram_id='d1'):
        page_w = max(2000, int(self._max_x * 1.10))
        page_h = max(2000, int(self._max_y * 1.10))
        hdr = (
            f'  <diagram id="{diagram_id}" name="{html.escape(str(self._diagram_name))}">\n'
            '    <mxGraphModel grid="0" guides="1" tooltips="1" connect="1" arrows="1"\n'
            f'                  fold="1" page="1" pageScale="1" pageWidth="{page_w}"\n'
            f'                  pageHeight="{page_h}" math="0" shadow="0">\n'
            '      <root>\n'
            '        <mxCell id="0"/>\n'
            '        <mxCell id="1" parent="0"/>\n'
        )
        ftr = '      </root>\n    </mxGraphModel>\n  </diagram>\n'
        return hdr + '\n'.join(self._cells) + '\n' + ftr


# ======================================================================
#  Styles Parsers
# ======================================================================
def _parse_cell_styles(z):
    """Parse styles.xml: xf_index -> fill color '#RRGGBB'."""
    xf_fills = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except Exception:
        return xf_fills
    ns = {'x': SS}
    fills = []
    for fill_el in root.findall('.//x:fills/x:fill', ns):
        color = None
        pf = fill_el.find(f'{{{SS}}}patternFill')
        if pf is not None and pf.attrib.get('patternType', 'none') != 'none':
            fg = pf.find(f'{{{SS}}}fgColor')
            if fg is not None:
                c = _parse_color_el(fg, default=None)
                if c and not _should_skip_fill(c):
                    color = c
        fills.append(color)
    for i, xf in enumerate(root.findall('.//x:cellXfs/x:xf', ns)):
        fill_id = int(xf.attrib.get('fillId', '0'))
        if fill_id < len(fills) and fills[fill_id]:
            xf_fills[i] = fills[fill_id]
    return xf_fills


def _parse_cell_borders(z):
    """Parse styles.xml: xf_index -> {side: (color, width, dash_pattern)}."""
    xf_borders = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except Exception:
        return xf_borders
    ns = {'x': SS}
    border_defs = []
    for bel in root.findall('.//x:borders/x:border', ns):
        sides = {}
        for side in ('left', 'right', 'top', 'bottom'):
            sel = bel.find(f'{{{SS}}}{side}')
            if sel is None:
                continue
            sname = sel.attrib.get('style')
            if not sname:
                continue
            color = _parse_color_el(sel.find(f'{{{SS}}}color'))
            bw, dash = BORDER_STYLE_MAP.get(sname, (1, None))
            sides[side] = (color, bw, dash, sname)
        border_defs.append(sides)
    for i, xf in enumerate(root.findall('.//x:cellXfs/x:xf', ns)):
        bid = int(xf.attrib.get('borderId', '0'))
        if 0 <= bid < len(border_defs) and border_defs[bid]:
            xf_borders[i] = border_defs[bid]
    return xf_borders


def _parse_cell_text_styles(z):
    """Parse styles.xml: xf_index -> text style dict (font, alignment, underline, strike)."""
    xf_text_styles = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except Exception:
        return xf_text_styles
    ns = {'x': SS}
    fonts = []
    for font_el in root.findall('.//x:fonts/x:font', ns):
        name_el = font_el.find(f'{{{SS}}}name')
        size_el = font_el.find(f'{{{SS}}}sz')
        color_el = font_el.find(f'{{{SS}}}color')
        bold = font_el.find(f'{{{SS}}}b') is not None
        italic = font_el.find(f'{{{SS}}}i') is not None
        underline = font_el.find(f'{{{SS}}}u') is not None
        strike = font_el.find(f'{{{SS}}}strike') is not None
        fonts.append({
            'fontFamily': _normalize_font_name(name_el.attrib.get('val')) if name_el is not None else None,
            'fontSize': max(6, round(float(size_el.attrib.get('val', '11')))) if size_el is not None else 11,
            'fontColor': _parse_color_el(color_el, default='#000000'),
            'bold': bold,
            'italic': italic,
            'underline': underline,
            'strike': strike,
        })
    for i, xf in enumerate(root.findall('.//x:cellXfs/x:xf', ns)):
        style = {}
        fid = int(xf.attrib.get('fontId', '0'))
        if 0 <= fid < len(fonts):
            f = fonts[fid]
            if f.get('fontFamily'):
                style['fontFamily'] = str(f['fontFamily']).replace('"', '')
            if f.get('fontSize'):
                style['fontSize'] = f['fontSize']
            if f.get('fontColor') and f['fontColor'] != '#000000':
                style['fontColor'] = f['fontColor']
            fs = 0
            if f.get('bold'):
                fs |= 1
            if f.get('italic'):
                fs |= 2
            if f.get('underline'):
                fs |= 4
            if f.get('strike'):
                style['textDecoration'] = 'line-through'
            if fs:
                style['fontStyle'] = fs
        al = xf.find(f'{{{SS}}}alignment')
        if al is not None:
            h = al.attrib.get('horizontal')
            v = al.attrib.get('vertical')
            if h in ('left', 'center', 'right'):
                style['align'] = h
            if v in ('top', 'center', 'bottom'):
                style['verticalAlign'] = {'center': 'middle'}.get(v, v)
            if al.attrib.get('wrapText') == '1':
                style['wrapText'] = True
            rot = al.attrib.get('textRotation')
            if rot:
                style['rotation'] = int(rot)
        xf_text_styles[i] = style
    return xf_text_styles


def _parse_cell_number_formats(z):
    """Parse styles.xml: xf_index -> (numFmtId, formatCode)."""
    xf_numfmts = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except Exception:
        return xf_numfmts
    ns = {'x': SS}
    custom = {
        int(el.attrib.get('numFmtId', '0')): el.attrib.get('formatCode', '')
        for el in root.findall('.//x:numFmts/x:numFmt', ns)
    }
    for i, xf in enumerate(root.findall('.//x:cellXfs/x:xf', ns)):
        nid = int(xf.attrib.get('numFmtId', '0'))
        xf_numfmts[i] = (nid, custom.get(nid, ''))
    return xf_numfmts


# ======================================================================
#  Helper Functions
# ======================================================================
def _build_merged_cell_maps(sh_root):
    ns = {'x': SS}
    merged_topleft = {}
    merged_children = set()
    for mc in sh_root.findall('.//x:mergeCell', ns):
        ref = mc.attrib.get('ref', '')
        if not ref:
            continue
        try:
            c1, r1, c2, r2 = _parse_range_ref(ref)
        except Exception:
            continue
        merged_topleft[(r1, c1)] = (r2, c2)
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                if rr != r1 or cc != c1:
                    merged_children.add((rr, cc))
    return merged_topleft, merged_children


def _build_merge_owner_map(sh_root):
    """Map each merged-cell coordinate to the merge block top-left cell."""
    merged_topleft, _ = _build_merged_cell_maps(sh_root)
    owner = {}
    for (r1, c1), (r2, c2) in merged_topleft.items():
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                owner[(rr, cc)] = (r1, c1)
    return owner


def _read_cell_raw_text(cell, shared_strings):
    ns = {'x': SS}
    cell_type = cell.attrib.get('t', '')
    if cell_type == 'inlineStr':
        inline = cell.find('x:is', ns)
        if inline is None:
            return ''
        return ''.join(t.text for t in inline.iter(f'{{{SS}}}t') if t.text)
    v_el = cell.find('x:v', ns)
    if v_el is None or v_el.text is None:
        return ''
    if cell_type == 's':
        idx = int(v_el.text)
        return shared_strings[idx] if idx < len(shared_strings) else ''
    return v_el.text


def _build_cell_value_map(sh_root, shared_strings):
    ns = {'x': SS}
    value_map = {}
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            value_map[(r, c)] = _read_cell_raw_text(cell, shared_strings)
    return value_map


def _build_fill_grid(sh_root, xf_fills):
    ns = {'x': SS}
    grid = {}
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            s_attr = int(cell.attrib.get('s', 0))
            fc = xf_fills.get(s_attr)
            if fc:
                grid[(r, c)] = fc
    return grid


def _parse_hyperlinks(z, sheet_path):
    """Parse hyperlinks from sheet XML and its relationships."""
    hyperlinks = {}
    try:
        sh_root = ET.fromstring(z.read(sheet_path).decode('utf-8'))
    except Exception:
        return hyperlinks
    ns = {'x': SS}
    # Load sheet rels for external hyperlinks
    num = sheet_path.rsplit('/', 1)[-1].replace('sheet', '').replace('.xml', '')
    rels_path = f'xl/worksheets/_rels/sheet{num}.xml.rels'
    ext_links = {}
    if rels_path in z.namelist():
        try:
            rels_root = ET.fromstring(z.read(rels_path).decode('utf-8'))
            for rel in rels_root:
                if 'hyperlink' in rel.attrib.get('Type', '').lower():
                    rid = rel.attrib.get('Id', '')
                    target = rel.attrib.get('Target', '')
                    if rid and target:
                        ext_links[rid] = target
        except Exception:
            pass
    for hl in sh_root.findall('.//x:hyperlinks/x:hyperlink', ns):
        ref = hl.attrib.get('ref', '')
        if not ref:
            continue
        rid = hl.attrib.get(f'{{{R}}}id', '')
        location = hl.attrib.get('location', '')
        url = ext_links.get(rid, '') or location
        if url:
            try:
                c, r = _cell_ref(ref.split(':')[0])
                hyperlinks[(r, c)] = url
            except Exception:
                pass
    return hyperlinks


def _format_excel_time(value):
    total_minutes = int(round(value * 24 * 60))
    return f'{total_minutes // 60}:{total_minutes % 60:02d}'


def _format_numeric_value(raw, style_numfmt):
    try:
        fv = float(raw)
    except ValueError:
        return raw
    num_fmt_id, fmt_code = style_numfmt
    fmt = (fmt_code or '').lower()
    is_time = (num_fmt_id in {18, 19, 20, 21, 22, 45, 46, 47}
               or ('h' in fmt and 'm' in fmt))
    if is_time:
        return _format_excel_time(fv)
    return str(int(fv)) if fv.is_integer() else raw


# ======================================================================
#  Cell Fill Rendering
# ======================================================================
def _add_cell_fills(sh_root, col_x, row_y, col_w, row_h, xf_fills, bld, cfg, bounds, log):
    """Render cell background fills. Optionally merge adjacent same-color cells."""
    ns = {'x': SS}
    min_r, max_r, min_c, max_c = bounds
    color_grid = {}
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        if r < min_r or r > max_r:
            continue
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            if c < min_c or c > max_c:
                continue
            s_attr = int(cell.attrib.get('s', 0))
            fc = xf_fills.get(s_attr)
            if fc:
                color_grid[(r, c)] = fc
    # Propagate merged cell colors
    merged_topleft, _ = _build_merged_cell_maps(sh_root)
    for (r1, c1), (r2, c2) in merged_topleft.items():
        color = color_grid.get((r1, c1))
        if color:
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    if (rr, cc) not in color_grid:
                        color_grid[(rr, cc)] = color
    log(f"  Color grid cells: {len(color_grid)}")
    if not color_grid:
        return 0
    CX_LAST = len(col_x) - 1
    RY_LAST = len(row_y) - 1
    if not cfg.merge_fills:
        count = 0
        for (r, c), color in sorted(color_grid.items()):
            px = col_x[min(c, CX_LAST)] / cfg.scale
            py = row_y[min(r, RY_LAST)] / cfg.scale
            pw = max(2.0, col_x[min(c + 1, CX_LAST)] / cfg.scale - px)
            ph = max(2.0, row_y[min(r + 1, RY_LAST)] / cfg.scale - py)
            style = f'whiteSpace=wrap;html=1;fillColor={color};strokeColor=none;'
            bld.add('', px, py, pw, ph, style)
            count += 1
        return count
    # Merge adjacent same-color cells into rectangles
    processed = set()
    count = 0
    for (r, c) in sorted(color_grid.keys()):
        if (r, c) in processed:
            continue
        color = color_grid[(r, c)]
        c_end = c
        while color_grid.get((r, c_end + 1)) == color and (r, c_end + 1) not in processed:
            c_end += 1
        r_end = r
        while True:
            nr = r_end + 1
            if all(color_grid.get((nr, cc)) == color and (nr, cc) not in processed
                   for cc in range(c, c_end + 1)):
                r_end = nr
            else:
                break
        for rr in range(r, r_end + 1):
            for cc in range(c, c_end + 1):
                processed.add((rr, cc))
        px = col_x[min(c, CX_LAST)] / cfg.scale
        py = row_y[min(r, RY_LAST)] / cfg.scale
        px_end = col_x[min(c_end + 1, CX_LAST)] / cfg.scale
        py_end = row_y[min(r_end + 1, RY_LAST)] / cfg.scale
        w = max(2.0, px_end - px)
        h = max(2.0, py_end - py)
        style = f'whiteSpace=wrap;html=1;fillColor={color};strokeColor=none;'
        bld.add('', px, py, w, h, style)
        count += 1
    return count


# ======================================================================
#  Cell Border Rendering
# ======================================================================
def _add_cell_borders(sh_root, col_x, row_y, col_w, row_h, xf_borders, xf_fills, bld, cfg, bounds):
    """Render cell borders with full dash pattern support.

    Skips internal left/right borders between two adjacent same-fill cells so
    that a horizontal run of filled cells (e.g. a wide yellow label row) does
    not show phantom vertical dividers that Excel itself does not render.
    The outer left/right borders of the filled region (where the neighbor is
    unfilled or a different color) are preserved.
    """
    ns = {'x': SS}
    min_r, max_r, min_c, max_c = bounds

    # Pre-scan: record fill color per (r, c) to drive internal-border suppression.
    fill_positions = {}
    merge_owner = _build_merge_owner_map(sh_root)
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            s_attr = int(cell.attrib.get('s', 0))
            fc = xf_fills.get(s_attr)
            if fc:
                fill_positions[(r, c)] = fc

    CX_LAST = len(col_x) - 1
    RY_LAST = len(row_y) - 1
    count = 0
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        if r < min_r or r > max_r:
            continue
        cy = row_y[min(r, RY_LAST)] / cfg.scale
        ch = max(1.0, _pts_px(row_h[r], cfg) / cfg.scale)
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            if c < min_c or c > max_c:
                continue
            s_attr = int(cell.attrib.get('s', 0))
            border_info = xf_borders.get(s_attr)
            if not border_info:
                continue
            cx = col_x[min(c, CX_LAST)] / cfg.scale
            cw = max(1.0, _chars_px(col_w[c], cfg) / cfg.scale)
            own_fill = fill_positions.get((r, c))
            own_merge = merge_owner.get((r, c))
            for side, (color, width_px, dash, _sname) in border_info.items():
                # Suppress internal vertical/horizontal dividers between same-fill cells.
                if own_fill:
                    if side == 'left' and fill_positions.get((r, c - 1)) == own_fill:
                        continue
                    if side == 'right' and fill_positions.get((r, c + 1)) == own_fill:
                        continue
                    if side == 'top' and fill_positions.get((r - 1, c)) == own_fill:
                        continue
                    if side == 'bottom' and fill_positions.get((r + 1, c)) == own_fill:
                        continue
                # Suppress borders inside a merged-cell block.
                if own_merge:
                    if side == 'left' and merge_owner.get((r, c - 1)) == own_merge:
                        continue
                    if side == 'right' and merge_owner.get((r, c + 1)) == own_merge:
                        continue
                    if side == 'top' and merge_owner.get((r - 1, c)) == own_merge:
                        continue
                    if side == 'bottom' and merge_owner.get((r + 1, c)) == own_merge:
                        continue
                dash_style = f'dashPattern={dash};' if dash else ''
                style = (f'whiteSpace=wrap;html=1;fillColor={color};strokeColor={color};'
                         f'strokeWidth={width_px};{dash_style}')
                if side == 'top':
                    bld.add('', cx, cy, cw, max(width_px, 1), style)
                elif side == 'bottom':
                    bld.add('', cx, cy + ch - max(width_px, 1), cw, max(width_px, 1), style)
                elif side == 'left':
                    bld.add('', cx, cy, max(width_px, 1), ch, style)
                elif side == 'right':
                    bld.add('', cx + cw - max(width_px, 1), cy, max(width_px, 1), ch, style)
                count += 1
    return count


# ======================================================================
#  Cell Text Style & Label Rendering
# ======================================================================
def _estimate_text_units(text):
    """Estimate text display width units (narrow=0.35, ASCII=0.6, CJK=1.0)."""
    units = 0.0
    for ch in text:
        code = ord(ch)
        if ch in 'ilI1.:;| ':
            units += 0.35
        elif code < 128:
            units += 0.6
        else:
            units += 1.0
    return max(units, 1.0)


def _fit_font_size(text, width, height, base_font_size):
    """Shrink font size until text fits in width x height."""
    font_size = max(6, base_font_size)
    while font_size > 6:
        line_cap = max(1.0, (width - 2) / max(font_size * 0.95, 1))
        req_lines = ceil(_estimate_text_units(text) / line_cap)
        max_lines = max(1, int(height / max(font_size * 1.15, 1)))
        if req_lines <= max_lines:
            break
        font_size -= 1
    return font_size


def _is_compact_label(text):
    """Short labels like '12:34' or '42' get compact center alignment."""
    s = str(text).strip()
    if re.fullmatch(r'\d{1,2}[\uff1a:]\d{2}', s):
        return True
    if re.fullmatch(r'\d+', s) and len(s) <= 2:
        return True
    return False


def _make_cell_text_style(style_info, text, width, height, compact=False):
    """Build DrawIO style string for a cell text label with font auto-fit."""
    eff = dict(style_info)
    if compact:
        eff['align'] = 'center'
        eff['verticalAlign'] = 'middle'
    fsz = _fit_font_size(text, width, height, eff.get('fontSize', 10))
    parts = [
        'text', 'html=1', 'strokeColor=none', 'fillColor=none',
        'whiteSpace=wrap',
        f'overflow={"hidden" if compact else "fill"}',
        f'align={eff.get("align", "left")}',
        f'verticalAlign={eff.get("verticalAlign", "middle")}',
        f'fontSize={fsz}',
    ]
    if eff.get('fontFamily'):
        parts.append(f'fontFamily={eff["fontFamily"]}')
    if eff.get('fontColor'):
        parts.append(f'fontColor={eff["fontColor"]}')
    if eff.get('fontStyle'):
        parts.append(f'fontStyle={eff["fontStyle"]}')
    if eff.get('textDecoration'):
        parts.append(f'textDecoration={eff["textDecoration"]}')
    if eff.get('rotation'):
        parts.append(f'rotation={-eff["rotation"]}')
    parts.append('spacingTop=1' if compact else 'spacingTop=3')
    if not compact and eff.get('align', 'left') == 'left':
        parts.append('spacingLeft=5')
    return ';'.join(parts) + ';'


def _add_cell_labels(sh_root, col_x, row_y, col_w, row_h, shared_strings,
                     xf_text_styles, xf_numfmts, xf_fills, bld, cfg, bounds, hyperlinks):
    """Render cell text labels with hyperlink, rotation and text extension support."""
    ns = {'x': SS}
    min_r, max_r, min_c, max_c = bounds
    CX_LAST = len(col_x) - 1
    RY_LAST = len(row_y) - 1
    merged_topleft, merged_children = _build_merged_cell_maps(sh_root)
    value_map = _build_cell_value_map(sh_root, shared_strings)
    fill_grid = _build_fill_grid(sh_root, xf_fills)
    count = 0
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        if r < min_r or r > max_r:
            continue
        ry = row_y[min(r, RY_LAST)] / cfg.scale
        rh = max(1.0, _pts_px(row_h[r], cfg) / cfg.scale)
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            if (r, c) in merged_children:
                continue
            if c < min_c or c > max_c:
                continue
            t = cell.attrib.get('t', '')
            raw_value = _read_cell_raw_text(cell, shared_strings)
            if raw_value == '' or raw_value.strip() == '':
                continue
            if t in {'s', 'str', 'inlineStr'}:
                val = raw_value
            else:
                s_attr = int(cell.attrib.get('s', 0))
                val = _format_numeric_value(raw_value, xf_numfmts.get(s_attr, (0, '')))
            if not val:
                continue
            cx = col_x[min(c, CX_LAST)] / cfg.scale
            s_attr = int(cell.attrib.get('s', 0))
            style_info = xf_text_styles.get(s_attr, {})
            compact = _is_compact_label(val)
            if (r, c) in merged_topleft:
                r_end, c_end = merged_topleft[(r, c)]
                cw = max(1.0, (col_x[min(c_end + 1, CX_LAST)] - col_x[min(c, CX_LAST)]) / cfg.scale)
                ch = max(1.0, (row_y[min(r_end + 1, RY_LAST)] - row_y[min(r, RY_LAST)]) / cfg.scale)
                text_x, text_y, text_w, text_h = cx, ry, cw, ch
            else:
                # Non-merged: try to extend text into adjacent empty cells on the right
                base_w = max(1.0, _chars_px(col_w[c], cfg) / cfg.scale)
                ch = rh
                c_end = c
                if not compact:
                    base_font = style_info.get('fontSize', 10)
                    needed_px = _estimate_text_units(val) * base_font * 0.72 + 10
                    own_fill = fill_grid.get((r, c))
                    acc_w = base_w
                    nc = c + 1
                    while nc <= max_c and nc < CX_LAST:
                        # Stop if next cell has any value
                        next_val = value_map.get((r, nc), '')
                        if next_val and next_val.strip():
                            break
                        next_fill = fill_grid.get((r, nc))
                        if own_fill:
                            # Extend only while adjacent cells share the same fill color
                            if next_fill != own_fill:
                                break
                        else:
                            # No fill: stop if adjacent cell has a fill
                            if next_fill:
                                break
                            # Stop once accumulated width covers needed width
                            if acc_w >= needed_px:
                                break
                        nc_w = max(1.0, _chars_px(col_w[nc], cfg) / cfg.scale)
                        acc_w += nc_w
                        c_end = nc
                        nc += 1
                if c_end > c:
                    cw = max(1.0, (col_x[min(c_end + 1, CX_LAST)] - col_x[min(c, CX_LAST)]) / cfg.scale)
                else:
                    cw = base_w
                text_x, text_y, text_w, text_h = cx, ry, cw, ch
                # Padding for left-aligned non-compact labels
                if not compact and style_info.get('align', 'left') == 'left':
                    text_x += 2
                    text_w = max(1.0, text_w - 2)
                    text_y += 2
                    text_h = max(1.0, text_h - 2)
            # Attach hyperlink if present
            link = hyperlinks.get((r, c), '')
            display_val = val
            if link:
                display_val = f'<a href="{html.escape(link)}">{html.escape(val)}</a>'
            cell_style = _make_cell_text_style(style_info, val, text_w, text_h, compact=compact)
            bld.add(display_val, text_x, text_y, text_w, text_h, cell_style, force=True)
            count += 1
    return count


# ======================================================================
#  Drawing Shape Helpers
# ======================================================================
def _parse_drawing_color(el):
    if el is None:
        return None
    sf = el.find(f'{{{A}}}solidFill') or el
    s = sf.find(f'{{{A}}}srgbClr')
    if s is not None:
        hex6 = s.attrib.get('val', '000000').upper()
        hex6 = _apply_color_modifiers(hex6, s)
        return '#' + hex6
    sc = sf.find(f'{{{A}}}schemeClr')
    if sc is not None:
        name = sc.attrib.get('val', 'dk1')
        base = SCHEME_COLORS.get(name, '808080')
        base = _apply_color_modifiers(base, sc)
        return '#' + base.upper()
    sy = sf.find(f'{{{A}}}sysClr')
    if sy is not None:
        last = sy.attrib.get('lastClr')
        if last:
            hex6 = last.upper()
            hex6 = _apply_color_modifiers(hex6, sy)
            return '#' + hex6
    return None


def _sp_fill(sp_pr):
    if sp_pr.find(f'{{{A}}}noFill') is not None:
        return 'none'
    for fill_tag in (f'{{{A}}}solidFill', f'{{{A}}}gradFill', f'{{{A}}}pattFill'):
        fe = sp_pr.find(fill_tag)
        if fe is not None:
            if fill_tag.endswith('solidFill'):
                c = _parse_drawing_color(fe)
            elif fill_tag.endswith('gradFill'):
                gs = fe.find(f'.//{{{A}}}gs')
                c = _parse_drawing_color(gs) if gs is not None else None
            else:
                bg = fe.find(f'{{{A}}}bgClr')
                c = _parse_drawing_color(bg) if bg is not None else None
            if c:
                return c
    return '#FFFFFF'


def _sp_line(sp_pr):
    ln = sp_pr.find(f'{{{A}}}ln')
    if ln is None:
        return '#000000', 1
    if ln.find(f'{{{A}}}noFill') is not None:
        return 'none', 0
    sf = ln.find(f'{{{A}}}solidFill')
    color = _parse_drawing_color(sf) if sf is not None else '#000000'
    if color is None:
        color = '#000000'
    w_emu = int(ln.attrib.get('w', '12700'))
    return color, max(1, round(w_emu / 12700))


def _sp_geom(sp_pr):
    g = sp_pr.find(f'{{{A}}}prstGeom')
    return g.attrib.get('prst', 'rect') if g is not None else 'rect'


def _sp_fontsize(txb):
    if txb is None:
        return 9
    for tag in (f'{{{A}}}rPr', f'{{{A}}}endParaRPr'):
        e = txb.find(f'.//{tag}')
        if e is not None:
            sz = e.attrib.get('sz')
            if sz:
                return max(7, round(int(sz) / 100))
    return 9


def _sp_font_style(txb):
    if txb is None:
        return {}, None
    rpr = txb.find(f'.//{{{A}}}rPr')
    if rpr is None:
        rpr = txb.find(f'.//{{{A}}}endParaRPr')
    if rpr is None:
        return {}, None
    extra = {}
    solid = rpr.find(f'{{{A}}}solidFill')
    if solid is not None:
        fc = _parse_drawing_color(solid)
        if fc and fc not in ('#000000', '#FFFFFF'):
            extra['fontColor'] = fc
    fs = 0
    if rpr.attrib.get('b') == '1':
        fs |= 1
    if rpr.attrib.get('i') == '1':
        fs |= 2
    if fs:
        extra['fontStyle'] = fs
    return extra, None


def _make_shape_style(prst, fill, lc, lw, fsz, font_extra=None,
                      shape_override=None, extra_parts=None):
    parts = ['whiteSpace=wrap', 'html=1']
    if shape_override:
        parts.append(shape_override.rstrip(';'))
    else:
        extra = GEOM_STYLES.get(prst, '')
        if extra:
            parts.append(extra.rstrip(';'))
    parts.append(f'fillColor={fill}' if fill != 'none' else 'fillColor=none')
    parts.append(f'strokeColor={lc}' if lc != 'none' else 'strokeColor=none')
    if lw > 1:
        parts.append(f'strokeWidth={lw}')
    if fsz != 9:
        parts.append(f'fontSize={fsz}')
    if font_extra:
        if 'fontColor' in font_extra:
            parts.append(f'fontColor={font_extra["fontColor"]}')
        if 'fontStyle' in font_extra:
            parts.append(f'fontStyle={font_extra["fontStyle"]}')
    if extra_parts:
        parts.extend(p for p in extra_parts if p)
    return ';'.join(parts) + ';'


def _get_text(el):
    return ''.join(t.text for t in el.iter(f'{{{A}}}t') if t.text)


def _get_xfrm(xfrm):
    def iv(el, attr, default=0):
        return int(el.attrib.get(attr, default)) if el is not None else default
    off = xfrm.find(f'{{{A}}}off')
    ext = xfrm.find(f'{{{A}}}ext')
    choff = xfrm.find(f'{{{A}}}chOff')
    chext = xfrm.find(f'{{{A}}}chExt')
    ox, oy = iv(off, 'x'), iv(off, 'y')
    ecx, ecy = iv(ext, 'cx'), iv(ext, 'cy')
    chox, choy = iv(choff, 'x', ox), iv(choff, 'y', oy)
    chcx, chcy = iv(chext, 'cx', ecx), iv(chext, 'cy', ecy)
    return ox, oy, ecx, ecy, chox, choy, chcx, chcy


def _descend_alt_content(parent):
    """Return a flat list of effective children, unwrapping mc:AlternateContent.

    Modern Office documents wrap shapes in
    ``<mc:AlternateContent><mc:Choice…/><mc:Fallback…/></mc:AlternateContent>``
    where ``mc:Choice`` uses features we do not implement (a14 icons, imgProps,
    etc). Preferring ``mc:Fallback`` lets us pick up PNG/JPG previews even for
    Office Icons and OLE embedded objects.
    """
    out = []
    for child in parent:
        tag = child.tag
        if tag.startswith('{' + MC + '}') and tag.endswith('}AlternateContent'):
            fb = child.find(f'{{{MC}}}Fallback')
            if fb is not None:
                out.extend(list(fb))
                continue
            ch = child.find(f'{{{MC}}}Choice')
            if ch is not None:
                out.extend(list(ch))
                continue
        else:
            out.append(child)
    return out


def _xfrm_transform(xfrm):
    """Read (rotation_deg, flipH, flipV) from an ``a:xfrm`` element."""
    if xfrm is None:
        return 0.0, 0, 0
    rot = 0.0
    try:
        rot_raw = xfrm.attrib.get('rot')
        if rot_raw is not None:
            rot = int(rot_raw) / 60000.0
    except (TypeError, ValueError):
        rot = 0.0
    fh = 1 if xfrm.attrib.get('flipH') == '1' else 0
    fv = 1 if xfrm.attrib.get('flipV') == '1' else 0
    return rot, fh, fv


def _rotate_point(px, py, cx, cy, deg):
    """Rotate point around center by ``deg`` degrees in screen coordinates."""
    if not deg:
        return px, py
    from math import cos, radians, sin
    t = radians(deg)
    dx, dy = px - cx, py - cy
    rx = dx * cos(t) - dy * sin(t)
    ry = dx * sin(t) + dy * cos(t)
    return cx + rx, cy + ry


def _append_transform_style(parts, rot, fh, fv):
    """Append ``rotation``/``flipH``/``flipV`` fragments when non-zero."""
    if rot:
        parts.append(f'rotation={round(rot, 2)}')
    if fh:
        parts.append('flipH=1')
    if fv:
        parts.append('flipV=1')


def _ln_style_parts(ln):
    """Extract drawio style fragments (dash/arrows) from an ``a:ln`` element."""
    parts = []
    if ln is None:
        return parts, False, False
    head = ln.find(f'{{{A}}}headEnd')
    tail = ln.find(f'{{{A}}}tailEnd')
    has_head = head is not None
    has_tail = tail is not None
    if has_head:
        htype = head.attrib.get('type', 'none')
        parts.append(f'startArrow={ARROW_MAP.get(htype, "classic")}')
    if has_tail:
        ttype = tail.attrib.get('type', 'none')
        parts.append(f'endArrow={ARROW_MAP.get(ttype, "classic")}')
    prst = ln.find(f'{{{A}}}prstDash')
    if prst is not None:
        pval = prst.attrib.get('val', 'solid')
        dp = PRST_DASH_MAP.get(pval)
        if dp:
            parts.append('dashed=1')
            parts.append(f'dashPattern={dp}')
    return parts, has_head, has_tail


def _extract_txbody_html(txBody):
    """Walk an ``xdr:txBody`` and return (html, has_rich).

    When any run has non-default formatting (color/size/bold/italic/underline),
    returns an HTML-escaped label with inline ``<font>``/``<b>``/``<i>``/``<u>``
    tags. When all runs are plain, returns the plain text with ``has_rich=False``
    so callers can keep the legacy path.
    """
    if txBody is None:
        return '', False
    paragraphs = []
    has_rich = False
    plain_parts = []
    for p in txBody.findall(f'{{{A}}}p'):
        runs_html = []
        for r in p.findall(f'{{{A}}}r'):
            t_el = r.find(f'{{{A}}}t')
            text = t_el.text if (t_el is not None and t_el.text) else ''
            if not text:
                continue
            plain_parts.append(text)
            esc = html.escape(text)
            rpr = r.find(f'{{{A}}}rPr')
            style_bits = []
            color = None
            bold = italic = underline = False
            if rpr is not None:
                sz = rpr.attrib.get('sz')
                if sz:
                    try:
                        style_bits.append(f'font-size:{int(sz) // 100}px')
                    except ValueError:
                        pass
                solid = rpr.find(f'{{{A}}}solidFill')
                if solid is not None:
                    color = _parse_drawing_color(solid)
                bold = rpr.attrib.get('b') == '1'
                italic = rpr.attrib.get('i') == '1'
                underline = rpr.attrib.get('u', 'none') not in ('none', '')
            if color and color not in ('#000000',):
                open_tag = f'<font color="{color}"'
                if style_bits:
                    open_tag += f' style="{";".join(style_bits)}"'
                open_tag += '>'
                close_tag = '</font>'
            elif style_bits:
                open_tag = f'<font style="{";".join(style_bits)}">'
                close_tag = '</font>'
            else:
                open_tag = ''
                close_tag = ''
            chunk = esc
            if bold:
                chunk = f'<b>{chunk}</b>'
                has_rich = True
            if italic:
                chunk = f'<i>{chunk}</i>'
                has_rich = True
            if underline:
                chunk = f'<u>{chunk}</u>'
                has_rich = True
            if open_tag:
                chunk = f'{open_tag}{chunk}{close_tag}'
                has_rich = True
            runs_html.append(chunk)
        if runs_html:
            paragraphs.append(''.join(runs_html))
    if not paragraphs:
        return '', False
    if has_rich:
        return '<br>'.join(paragraphs), True
    return '\n'.join(plain_parts), False


def _custgeom_to_stencil(path_elem):
    """Convert an ``a:path`` under ``a:pathLst`` to a drawio stencil string.

    Returns ``'stencil(<base64>)'`` or ``None`` when the path is empty or
    malformed. Bezier and line segments are mapped 1:1 to drawio stencil
    primitives. Quadratic Bezier curves are promoted to cubic using the
    standard formula ``C1 = P0 + 2/3 (P1 - P0)``, ``C2 = P2 + 2/3 (P1 - P2)``.
    """
    if path_elem is None:
        return None
    try:
        w = int(path_elem.attrib.get('w', '0'))
        h = int(path_elem.attrib.get('h', '0'))
    except ValueError:
        return None
    if w <= 0 or h <= 0:
        return None
    commands = []
    cursor = (0.0, 0.0)

    def _pts(node):
        out = []
        for pt in node.findall(f'{{{A}}}pt'):
            try:
                out.append((float(pt.attrib.get('x', '0')),
                            float(pt.attrib.get('y', '0'))))
            except ValueError:
                return []
        return out

    for child in path_elem:
        tag = child.tag.split('}')[-1]
        if tag == 'moveTo':
            pts = _pts(child)
            if pts:
                x, y = pts[0]
                commands.append(f'<move x="{x:.2f}" y="{y:.2f}"/>')
                cursor = (x, y)
        elif tag == 'lnTo':
            pts = _pts(child)
            if pts:
                x, y = pts[0]
                commands.append(f'<line x="{x:.2f}" y="{y:.2f}"/>')
                cursor = (x, y)
        elif tag == 'cubicBezTo':
            pts = _pts(child)
            if len(pts) == 3:
                (x1, y1), (x2, y2), (x3, y3) = pts
                commands.append(
                    f'<curve x1="{x1:.2f}" y1="{y1:.2f}" '
                    f'x2="{x2:.2f}" y2="{y2:.2f}" '
                    f'x3="{x3:.2f}" y3="{y3:.2f}"/>'
                )
                cursor = (x3, y3)
        elif tag == 'quadBezTo':
            pts = _pts(child)
            if len(pts) == 2:
                (qx, qy), (px, py) = pts
                x0, y0 = cursor
                c1x = x0 + 2 / 3 * (qx - x0)
                c1y = y0 + 2 / 3 * (qy - y0)
                c2x = px + 2 / 3 * (qx - px)
                c2y = py + 2 / 3 * (qy - py)
                commands.append(
                    f'<curve x1="{c1x:.2f}" y1="{c1y:.2f}" '
                    f'x2="{c2x:.2f}" y2="{c2y:.2f}" '
                    f'x3="{px:.2f}" y3="{py:.2f}"/>'
                )
                cursor = (px, py)
        elif tag == 'close':
            commands.append('<close/>')
    if not commands:
        return None
    stencil_xml = (
        f'<shape h="{h}" w="{w}" aspect="variable" strokewidth="inherit">'
        '<foreground><path>'
        + ''.join(commands)
        + '</path><fillstroke/></foreground></shape>'
    )
    b64 = base64.b64encode(stencil_xml.encode('utf-8')).decode('ascii')
    return f'stencil({b64})'


# ======================================================================
#  Image Extraction
# ======================================================================
def _extract_images(z, drawing_path):
    """Extract images referenced by drawing XML.

    Returns {rId: data_uri_or_url_or_None}. ``None`` marks a relationship that
    exists but points at an image format the browser (and drawio) cannot render
    directly. Callers should draw a placeholder rectangle for those instead of
    emitting a broken <img>. External ``TargetMode="External"`` relationships
    with an ``http(s)://`` URL are passed through verbatim so drawio can load
    the image at render time.

    For EMF/WMF/TIFF entries, if a same-stem PNG/JPG/SVG sibling exists in
    ``xl/media/`` (Office frequently ships both the vector and a raster
    fallback), the fallback is used instead so the icon still shows up.
    """
    images = {}
    num = drawing_path.rsplit('/', 1)[-1].replace('drawing', '').replace('.xml', '')
    rels_path = f'xl/drawings/_rels/drawing{num}.xml.rels'
    if rels_path not in z.namelist():
        return images
    try:
        rels_root = ET.fromstring(z.read(rels_path).decode('utf-8'))
    except Exception:
        return images

    mime_map = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
                'gif': 'image/gif', 'bmp': 'image/bmp', 'svg': 'image/svg+xml',
                'webp': 'image/webp'}

    zip_names = set(z.namelist())

    for rel in rels_root:
        rtype = rel.attrib.get('Type', '')
        if 'image' not in rtype.lower():
            continue
        rid = rel.attrib.get('Id', '')
        target = rel.attrib.get('Target', '')
        if not rid or not target:
            continue
        # External (linked) image — pass URL through when renderable by browser.
        if rel.attrib.get('TargetMode') == 'External':
            if target.startswith(('http://', 'https://')):
                images[rid] = target
            else:
                images[rid] = None
            continue
        img_path = 'xl/drawings/' + target if not target.startswith('/') else target.lstrip('/')
        img_path = img_path.replace('/../', '/').replace('/drawings/media/', '/media/')
        # Normalize: ../media/image1.png -> xl/media/image1.png
        if '../media/' in target:
            img_path = 'xl/media/' + target.split('../media/')[-1]

        if img_path not in zip_names:
            images[rid] = None
            continue

        ext = img_path.rsplit('.', 1)[-1].lower()

        # Non-renderable (EMF/WMF/TIFF/…): try to find a same-stem raster/SVG
        # fallback that Office may have saved alongside the original.
        if ext not in RENDERABLE_IMG_EXTS:
            stem_dir = img_path.rsplit('/', 1)[0]
            stem_name = img_path.rsplit('/', 1)[-1].rsplit('.', 1)[0]
            fallback = None
            for cand_ext in ('png', 'jpg', 'jpeg', 'svg', 'gif'):
                cand = f'{stem_dir}/{stem_name}.{cand_ext}'
                if cand in zip_names:
                    fallback = cand
                    break
            if fallback is None:
                images[rid] = None
                continue
            img_path = fallback
            ext = cand_ext

        mime = mime_map.get(ext)
        if not mime:
            images[rid] = None
            continue
        try:
            img_data = z.read(img_path)
        except Exception:
            images[rid] = None
            continue
        b64 = base64.b64encode(img_data).decode('ascii')
        images[rid] = f'data:{mime};base64,{b64}'
    return images


# ======================================================================
#  Shape / Connector / Picture Emitters
# ======================================================================
def _render_sp(sp, ax, ay, w, h, bld):
    """Render a shape at a resolved pixel rect, applying transform/line/geom/text."""
    spr = sp.find(f'{{{XDR}}}spPr')
    if spr is None:
        return
    if w < 1 or h < 1:
        return
    fill = _sp_fill(spr)
    lc, lw = _sp_line(spr)
    prst = _sp_geom(spr)
    txb = sp.find(f'{{{XDR}}}txBody')
    html_text, has_rich = _extract_txbody_html(txb)
    plain_text = _get_text(sp)
    text_value = html_text if has_rich else plain_text
    if not text_value and fill in ('#FFFFFF', 'none') and lc == 'none':
        return
    fsz = _sp_fontsize(txb)
    fe, _ = _sp_font_style(txb)

    # Optional custGeom → drawio stencil
    shape_override = None
    custgeom = spr.find(f'{{{A}}}custGeom')
    if custgeom is not None:
        path = custgeom.find(f'{{{A}}}pathLst/{{{A}}}path')
        stencil = _custgeom_to_stencil(path)
        if stencil:
            shape_override = f'shape={stencil}'

    extra = []
    ln = spr.find(f'{{{A}}}ln')
    ln_parts, _, _ = _ln_style_parts(ln)
    extra.extend(ln_parts)
    xfrm = spr.find(f'{{{A}}}xfrm')
    rot, fh, fv = _xfrm_transform(xfrm)
    _append_transform_style(extra, rot, fh, fv)

    style = _make_shape_style(prst, fill, lc, lw, fsz, fe,
                              shape_override=shape_override, extra_parts=extra)
    bld.add(text_value, ax, ay, w, h, style, force=bool(text_value))


def _emit_sp(sp, pax, pay, sx, sy, bld):
    spr = sp.find(f'{{{XDR}}}spPr')
    if spr is None:
        return
    xfrm = spr.find(f'{{{A}}}xfrm')
    if xfrm is None:
        return
    off = xfrm.find(f'{{{A}}}off')
    ext = xfrm.find(f'{{{A}}}ext')
    if off is None or ext is None:
        return
    ax = pax + int(off.attrib.get('x', 0)) * sx
    ay = pay + int(off.attrib.get('y', 0)) * sy
    w = int(ext.attrib.get('cx', 0)) * sx
    h = int(ext.attrib.get('cy', 0)) * sy
    _render_sp(sp, ax, ay, w, h, bld)


def _render_cxnsp_at_rect(cxn, ax, ay, w, h, bld):
    """Emit a connector as a drawio edge for a pre-resolved bbox rect.

    ``ax/ay/w/h`` is the connector's bounding box in drawio pixels. The line
    runs along one of the bbox diagonals, selected by ``flipH``/``flipV`` on
    the underlying ``a:xfrm``. Used by both the anchor-level path (which
    derives the rect from ``_anchor_rect``, so it shares pixel math with
    shapes and cell labels) and the group-level path in ``_emit_cxnsp``.
    """
    spr = cxn.find(f'{{{XDR}}}spPr')
    if spr is None:
        return
    prst_el = spr.find(f'{{{A}}}prstGeom')
    prst_name = prst_el.attrib.get('prst', '') if prst_el is not None else ''
    xfrm = spr.find(f'{{{A}}}xfrm')
    rot, fh, fv = _xfrm_transform(xfrm)
    edge_points = None
    if prst_name.startswith('bentConnector'):
        # OOXML bent connectors are orthogonal/elbow polylines.
        # When rendered as free drawio edges (no source/target cells), passing
        # only sourcePoint/targetPoint often collapses into a straight segment.
        # Keep opposite-corner endpoints and provide explicit waypoint(s) so the
        # elbow remains visible.
        if not fh and not fv:
            x1, y1, x2, y2 = ax, ay, ax + w, ay + h
        elif fh and not fv:
            x1, y1, x2, y2 = ax + w, ay, ax, ay + h
        elif fv and not fh:
            x1, y1, x2, y2 = ax, ay + h, ax + w, ay
        else:
            x1, y1, x2, y2 = ax + w, ay + h, ax, ay
        m = re.search(r'(\d+)$', prst_name or '')
        idx = int(m.group(1)) if m else 2
        # bentConnector2/4 and 3/5 are mirrored variants.
        # Swap bend corner selection so the elbow direction rotates 180°
        # from the previous mapping (matches Excel orientation better).
        first_corner = (x1, y2) if (idx % 2 == 0) else (x2, y1)
        edge_points = [first_corner]
    else:
        # Non-elbow connectors: center-line endpoints along the major axis.
        if w >= h:
            y = ay + (h / 2.0)
            x1, y1, x2, y2 = ax, y, ax + w, y
            if fh:
                x1, y1, x2, y2 = x2, y2, x1, y1
        else:
            x = ax + (w / 2.0)
            x1, y1, x2, y2 = x, ay, x, ay + h
            if fv:
                x1, y1, x2, y2 = x2, y2, x1, y1
    eff_rot = rot
    # Elbow connectors are usually quarter-turn oriented; snap near-right-angle
    # rotations to avoid mirrored routing caused by tiny float/import noise.
    if prst_name.startswith('bentConnector') and rot:
        q = round(rot / 90.0)
        snapped = q * 90.0
        if abs(rot - snapped) <= 1.0:
            eff_rot = snapped
    if eff_rot:
        cx, cy = ax + (w / 2.0), ay + (h / 2.0)
        # OOXML a:xfrm rot uses opposite sign vs drawio screen coordinates here.
        # Use negative rotation to avoid mirrored connector direction.
        x1, y1 = _rotate_point(x1, y1, cx, cy, -eff_rot)
        x2, y2 = _rotate_point(x2, y2, cx, cy, -eff_rot)
        if edge_points:
            edge_points = [
                _rotate_point(px, py, cx, cy, -eff_rot)
                for px, py in edge_points
            ]

    # Line appearance
    ln = spr.find(f'{{{A}}}ln')
    if ln is not None and ln.find(f'{{{A}}}noFill') is not None:
        return
    if ln is not None:
        sf = ln.find(f'{{{A}}}solidFill')
        color = _parse_drawing_color(sf) if sf is not None else '#000000'
        if color is None:
            color = '#000000'
    else:
        color = '#000000'
    lw_emu = int(ln.attrib.get('w', '12700')) if ln is not None else 12700
    lw_px = max(1, round(lw_emu / 12700))
    ln_parts, has_head, has_tail = _ln_style_parts(ln)

    # Preset connector geometry -> drawio edge routing hint.
    parts = ['html=1', 'rounded=0', 'jumpStyle=none']
    if prst_name.startswith('bentConnector'):
        # orthogonalEdgeStyle can over-route wildly without source/target IDs.
        # elbowEdgeStyle is more stable for free-floating endpoints.
        parts.append('edgeStyle=elbowEdgeStyle')
        if abs(x2 - x1) >= abs(y2 - y1):
            parts.append('elbow=horizontal')
        else:
            parts.append('elbow=vertical')
    elif prst_name.startswith('curvedConnector'):
        parts.append('edgeStyle=none')
        parts.append('curved=1')
    elif prst_name.startswith('straightConnector'):
        parts.append('edgeStyle=none')
    parts.append(f'strokeColor={color}')
    if lw_px > 1:
        parts.append(f'strokeWidth={lw_px}')
    # Suppress drawio's default classic arrow when OOXML has no markers.
    if not has_head:
        parts.append('startArrow=none')
    if not has_tail:
        parts.append('endArrow=none')
    parts.extend(ln_parts)
    style = ';'.join(parts) + ';'
    bld.add_edge(x1, y1, x2, y2, style, points=edge_points)


def _emit_cxnsp(cxn, pax, pay, sx, sy, bld):
    """Emit a connector shape whose bbox is stored in its own ``a:xfrm``.

    Used by the grpSp walker where the connector's xfrm is in group-local
    coordinates. Anchor-level connectors should use
    ``_render_cxnsp_at_rect`` directly with the resolved anchor rect so they
    share pixel math with shapes and cell labels.
    """
    spr = cxn.find(f'{{{XDR}}}spPr')
    if spr is None:
        return
    xfrm = spr.find(f'{{{A}}}xfrm')
    if xfrm is None:
        return
    off = xfrm.find(f'{{{A}}}off')
    ext = xfrm.find(f'{{{A}}}ext')
    if off is None or ext is None:
        return
    ax = pax + int(off.attrib.get('x', 0)) * sx
    ay = pay + int(off.attrib.get('y', 0)) * sy
    w = int(ext.attrib.get('cx', 0)) * sx
    h = int(ext.attrib.get('cy', 0)) * sy
    _render_cxnsp_at_rect(cxn, ax, ay, w, h, bld)


def _emit_pic(pic, images, pax, pay, sx, sy, bld):
    """Emit a picture element as an embedded image in DrawIO.

    Prefers the SVG alternate (``a14:svgBlip``) when Office shipped one next
    to a raster primary (the modern "Insert > Icons" path). If the resolved
    image format is not renderable by the browser, falls back to a dashed
    placeholder rectangle so the layout stays intact.
    """
    blip_fill = pic.find(f'{{{XDR}}}blipFill')
    if blip_fill is None:
        return
    blip = blip_fill.find(f'{{{A}}}blip')
    if blip is None:
        return
    primary_rid = blip.attrib.get(f'{{{R}}}embed', '')

    # Prefer svgBlip extension when available and resolvable.
    chosen_rid = primary_rid
    ext_lst = blip.find(f'{{{A}}}extLst')
    if ext_lst is not None:
        svg_blip = ext_lst.find(f'.//{{{ASVG}}}svgBlip')
        if svg_blip is not None:
            svg_rid = svg_blip.attrib.get(f'{{{R}}}embed', '')
            if svg_rid and images.get(svg_rid):
                chosen_rid = svg_rid

    data_uri = images.get(chosen_rid)
    # If the SVG extension didn't help but primary is renderable, use primary.
    if not data_uri and primary_rid and images.get(primary_rid):
        data_uri = images[primary_rid]

    spr = pic.find(f'{{{XDR}}}spPr')
    if spr is None:
        return
    xfrm = spr.find(f'{{{A}}}xfrm')
    if xfrm is None:
        return
    off = xfrm.find(f'{{{A}}}off')
    ext = xfrm.find(f'{{{A}}}ext')
    if off is None or ext is None:
        return
    ax = pax + int(off.attrib.get('x', 0)) * sx
    ay = pay + int(off.attrib.get('y', 0)) * sy
    w = int(ext.attrib.get('cx', 0)) * sx
    h = int(ext.attrib.get('cy', 0)) * sy
    _render_pic_at_rect(ax, ay, w, h, data_uri,
                        (primary_rid or chosen_rid), xfrm, bld)


def _render_pic_at_rect(ax, ay, w, h, data_uri, has_ref, xfrm, bld):
    """Render a picture at a resolved pixel rect, honoring rotation/flip.

    When ``data_uri`` is falsy but ``has_ref`` is true, draw a dashed
    placeholder so the layout is preserved for unsupported formats.
    """
    if w < 1 or h < 1:
        return
    rot, fh, fv = _xfrm_transform(xfrm)
    extras = []
    _append_transform_style(extras, rot, fh, fv)
    extra_style = ';'.join(extras) if extras else None
    if data_uri:
        bld.add_image(ax, ay, w, h, data_uri, extra_style=extra_style)
        return
    if has_ref:
        placeholder_parts = [
            'whiteSpace=wrap', 'html=1', 'fillColor=#F5F5F5',
            'strokeColor=#BDBDBD', 'dashed=1', 'align=center',
            'verticalAlign=middle', 'fontSize=8', 'fontColor=#757575',
        ]
        if extras:
            placeholder_parts.extend(extras)
        bld.add('[image]', ax, ay, w, h, ';'.join(placeholder_parts) + ';')


def _walk_group(grp, pax, pay, sx, sy, bld, images, depth=0):
    if depth > 25:
        return
    grp_pr = grp.find(f'{{{XDR}}}grpSpPr')
    if grp_pr is None:
        return
    xfrm = grp_pr.find(f'{{{A}}}xfrm')
    if xfrm is None:
        return
    ox, oy, ecx, ecy, chox, choy, chcx, chcy = _get_xfrm(xfrm)
    gax, gay = pax + ox * sx, pay + oy * sy
    gw, gh = ecx * sx, ecy * sy
    csx = (gw / chcx) if chcx else sx
    csy = (gh / chcy) if chcy else sy
    cox = gax - chox * csx
    coy = gay - choy * csy
    for child in _descend_alt_content(grp):
        ct = child.tag.split('}')[-1]
        if ct == 'sp':
            _emit_sp(child, cox, coy, csx, csy, bld)
        elif ct == 'cxnSp':
            _emit_cxnsp(child, cox, coy, csx, csy, bld)
        elif ct == 'grpSp':
            _walk_group(child, cox, coy, csx, csy, bld, images, depth + 1)
        elif ct == 'pic':
            _emit_pic(child, images, cox, coy, csx, csy, bld)


def _anchor_rect(anchor, col_x, row_y, cfg):
    from_el = anchor.find(f'{{{XDR}}}from')
    if from_el is None:
        return None
    cx_last = len(col_x) - 1
    ry_last = len(row_y) - 1
    fc = int(from_el.findtext(f'{{{XDR}}}col', '0') or '0')
    fco = int(from_el.findtext(f'{{{XDR}}}colOff', '0') or '0')
    fr = int(from_el.findtext(f'{{{XDR}}}row', '0') or '0')
    fro = int(from_el.findtext(f'{{{XDR}}}rowOff', '0') or '0')
    anc_x = col_x[min(fc, cx_last)] / cfg.scale + _emu_px(fco, cfg)
    anc_y = row_y[min(fr, ry_last)] / cfg.scale + _emu_px(fro, cfg)
    to_el = anchor.find(f'{{{XDR}}}to')
    ext_el = anchor.find(f'{{{XDR}}}ext')
    if to_el is not None:
        tc = int(to_el.findtext(f'{{{XDR}}}col', '0') or '0')
        tco = int(to_el.findtext(f'{{{XDR}}}colOff', '0') or '0')
        tr = int(to_el.findtext(f'{{{XDR}}}row', '0') or '0')
        tro = int(to_el.findtext(f'{{{XDR}}}rowOff', '0') or '0')
        anc_w = max(2.0, col_x[min(tc, cx_last)] / cfg.scale + _emu_px(tco, cfg) - anc_x)
        anc_h = max(2.0, row_y[min(tr, ry_last)] / cfg.scale + _emu_px(tro, cfg) - anc_y)
    elif ext_el is not None:
        anc_w = max(2.0, _emu_px(int(ext_el.attrib.get('cx', '9525')), cfg))
        anc_h = max(2.0, _emu_px(int(ext_el.attrib.get('cy', '9525')), cfg))
    else:
        anc_w, anc_h = 80.0, 24.0
    return anc_x, anc_y, anc_w, anc_h


# ======================================================================
#  Drawing Shapes (main entry)
# ======================================================================
def _add_drawing_shapes(z, drawing_path, col_x, row_y, bld, cfg):
    """Parse drawing XML and emit shapes, connectors, and images."""
    dr = ET.fromstring(z.read(drawing_path).decode('utf-8'))
    sc = 1.0 / cfg.emu_per_px / cfg.scale
    images = _extract_images(z, drawing_path) if cfg.render_images else {}
    for anchor in dr:
        tag = anchor.tag.split('}')[-1]
        if tag not in ('oneCellAnchor', 'twoCellAnchor'):
            continue
        rect = _anchor_rect(anchor, col_x, row_y, cfg)
        if rect is None:
            continue
        anc_x, anc_y, anc_w, anc_h = rect
        for child in _descend_alt_content(anchor):
            ct = child.tag.split('}')[-1]
            if ct == 'sp':
                _render_sp(child, anc_x, anc_y, anc_w, anc_h, bld)
            elif ct == 'grpSp':
                grp_pr = child.find(f'{{{XDR}}}grpSpPr')
                if grp_pr is None:
                    continue
                xfrm = grp_pr.find(f'{{{A}}}xfrm')
                if xfrm is None:
                    continue
                _, _, ecx, ecy, chox, choy, chcx, chcy = _get_xfrm(xfrm)
                csx = (anc_w / chcx) if chcx else sc
                csy = (anc_h / chcy) if chcy else sc
                cox = anc_x - chox * csx
                coy = anc_y - choy * csy
                for gc in _descend_alt_content(child):
                    gct = gc.tag.split('}')[-1]
                    if gct == 'sp':
                        _emit_sp(gc, cox, coy, csx, csy, bld)
                    elif gct == 'grpSp':
                        _walk_group(gc, cox, coy, csx, csy, bld, images)
                    elif gct == 'cxnSp':
                        _emit_cxnsp(gc, cox, coy, csx, csy, bld)
                    elif gct == 'pic':
                        _emit_pic(gc, images, cox, coy, csx, csy, bld)
            elif ct == 'cxnSp':
                # Anchor-level connector: use the _anchor_rect bbox so the
                # line endpoints align with the same pixel math the cell
                # labels and shapes use. Bypassing this and reading raw EMU
                # introduces a small but visible rightward/downward drift
                # that accumulates across columns.
                _render_cxnsp_at_rect(child, anc_x, anc_y, anc_w, anc_h, bld)
            elif ct == 'pic':
                # Top-level picture in anchor: resolve the primary/SVG alternate
                # rid and render at the anchor rect.
                blip_fill = child.find(f'{{{XDR}}}blipFill')
                if blip_fill is None:
                    continue
                blip = blip_fill.find(f'{{{A}}}blip')
                if blip is None:
                    continue
                primary_rid = blip.attrib.get(f'{{{R}}}embed', '')
                chosen_rid = primary_rid
                ext_lst = blip.find(f'{{{A}}}extLst')
                if ext_lst is not None:
                    svg_blip = ext_lst.find(f'.//{{{ASVG}}}svgBlip')
                    if svg_blip is not None:
                        svg_rid = svg_blip.attrib.get(f'{{{R}}}embed', '')
                        if svg_rid and images.get(svg_rid):
                            chosen_rid = svg_rid
                data_uri = images.get(chosen_rid)
                if not data_uri and primary_rid and images.get(primary_rid):
                    data_uri = images[primary_rid]
                spr = child.find(f'{{{XDR}}}spPr')
                pic_xfrm = spr.find(f'{{{A}}}xfrm') if spr is not None else None
                _render_pic_at_rect(anc_x, anc_y, anc_w, anc_h, data_uri,
                                    bool(primary_rid or chosen_rid),
                                    pic_xfrm, bld)


# ======================================================================
#  Path Resolution
# ======================================================================
def _find_paths(z, sheet_name):
    wb = ET.fromstring(z.read('xl/workbook.xml').decode('utf-8'))
    rid = next((sh.attrib.get(f'{{{R}}}id')
                for sh in wb.findall('.//{%s}sheet' % SS)
                if sh.attrib.get('name') == sheet_name), None)
    if not rid:
        available = [s.attrib.get('name') for s in wb.findall('.//{%s}sheet' % SS)]
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {available}")
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels').decode('utf-8'))
    sf = next(('xl/' + r.attrib['Target'].lstrip('/')
               for r in rels if r.attrib.get('Id') == rid), None)
    num = sf.rsplit('/', 1)[-1].replace('sheet', '').replace('.xml', '')
    rels_path = f'xl/worksheets/_rels/sheet{num}.xml.rels'
    if rels_path not in z.namelist():
        return sf, None
    sr = ET.fromstring(z.read(rels_path).decode('utf-8'))
    drw = next(('xl/' + r.attrib['Target'].lstrip('../')
                for r in sr
                if 'drawing' in r.attrib.get('Type', '')
                and 'vml' not in r.attrib.get('Type', '')), None)
    return sf, drw


def _load_shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    ss_root = ET.fromstring(z.read('xl/sharedStrings.xml').decode('utf-8'))
    return [
        ''.join(t.text for t in si.iter(f'{{{SS}}}t') if t.text)
        for si in ss_root.findall(f'{{{SS}}}si')
    ]


def _validate_workbook_suffix(input_path):
    suffix = Path(input_path).suffix.lower()
    if suffix not in {'.xlsx', '.xlsm'}:
        raise ValueError('Supported file types are .xlsx and .xlsm')


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


# ======================================================================
#  Main Converter
# ======================================================================
def _prepare_resources(z):
    # Populate SCHEME_COLORS / THEME_FILL_COLORS from the workbook's own
    # theme1.xml before any color resolution runs. Otherwise scheme references
    # like accent2 fall back to the built-in Office defaults, which usually
    # mismatch the colors the user actually saved in Excel.
    _load_theme_colors(z)
    return {
        'shared': _load_shared_strings(z),
        'xf_fills': _parse_cell_styles(z),
        'xf_borders': _parse_cell_borders(z),
        'xf_text_styles': _parse_cell_text_styles(z),
        'xf_numfmts': _parse_cell_number_formats(z),
    }


def _build_sheet_xml(z, sheet_name, diagram_id, resources, cfg, log):
    sf, drw_path = _find_paths(z, sheet_name)
    log(f"Sheet XML: {sf}")
    log(f"Drawing:   {drw_path or '(none)'}")
    sh_root = ET.fromstring(z.read(sf).decode('utf-8'))
    col_x, row_y, col_w, row_h = _build_grid(sh_root, cfg)
    bounds = _auto_detect_bounds(sh_root)
    log(f"Bounds: rows {bounds[0]}-{bounds[1]}, cols {bounds[2]}-{bounds[3]}")
    hyperlinks = _parse_hyperlinks(z, sf)
    log(f"Hyperlinks: {len(hyperlinks)}")
    bld = DrawioBuilder(diagram_name=sheet_name)
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
        _add_drawing_shapes(z, drw_path, col_x, row_y, bld, cfg)
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
        resources = _prepare_resources(z)
        diagrams = []
        for idx, sn in enumerate(names, start=1):
            log(f"Processing sheet '{sn}' ...")
            diagrams.append(_build_sheet_xml(z, sn, f'd{idx}', resources, cfg, log))
    xml_out = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<mxfile host="ExcelToDrawIOPlus" version="1.0" type="device">\n'
        + ''.join(diagrams)
        + '</mxfile>\n'
    )
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(xml_out)
    log(f"Written '{output_path}' ({len(xml_out):,} chars)")


def convert(xlsm, sheet=None, out=None, cfg=None, log_func=None):
    """Convert Excel file to Draw.io format."""
    if out is None:
        out = suggest_output_path(xlsm, sheet) if sheet else suggest_multi_output_path(xlsm)
    sheets = [sheet] if sheet else list_supported_sheets(xlsm)
    convert_sheets_to_file(xlsm, sheets, out, cfg=cfg, log_func=log_func)


# ======================================================================
#  CLI Entry Point
# ======================================================================
if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(
        description='Excel (.xlsx/.xlsm) to Draw.io (.drawio) General-Purpose Converter',
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
    parser.add_argument('--scale', type=float, default=1.0, help='Scale factor (default: 1.0)')
    args = parser.parse_args()

    if args.list_sheets:
        for name in list_supported_sheets(args.input):
            print(name)
        sys.exit(0)

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
    )
    sheets = args.sheets or list_supported_sheets(args.input)
    output = args.output or suggest_multi_output_path(args.input)
    convert_sheets_to_file(args.input, sheets, output, cfg=cfg)
