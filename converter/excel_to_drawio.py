"""
High-level Excel -> draw.io conversion entrypoint.

Self-contained OOXML parser that extracts:
  - Cell background fills (adjacent same-color cells merged into rectangles)
  - Cell borders
  - Drawing shapes (sp / grpSp / cxnSp) with full coordinate transforms
  - Cell text labels with font styling and number formatting

Ported from tests/excel_to_drawio-fromClaude.py (v10) and generalized
for multi-sheet / CLI / GUI usage.
"""

from __future__ import annotations

import html as html_mod
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field
from math import ceil
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


# ═══════════════════════════════════════════════════════════════════════════════
#  XML namespaces
# ═══════════════════════════════════════════════════════════════════════════════

SS_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

_NS_X = {"x": SS_NS}

# ═══════════════════════════════════════════════════════════════════════════════
#  Constants
# ═══════════════════════════════════════════════════════════════════════════════

EMU_PER_PX = 9525
CHAR_WIDTH = 7
POINT_TO_PX = 96 / 72
CELL_BOX_LEFT_PAD = 2
FILLED_TEXT_TOP_PAD = 2
GRID_MAX = 300

SKIP_FILL_COLORS = {
    "FFFFFF", "FFFFFE", "F2F2F2", "F3F3F3", "EBEBEB", "E7E6E6", "EEECE1",
    "D9D9D9", "BFBFBF", "000000", "0D0D0D",
}

# ═══════════════════════════════════════════════════════════════════════════════
#  Color tables
# ═══════════════════════════════════════════════════════════════════════════════

SCHEME_COLORS = {
    "dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6",
    "acc1": "4472C4", "acc2": "ED7D31", "acc3": "A9D18E", "acc4": "FFC000",
    "acc5": "5B9BD5", "acc6": "70AD47", "hlink": "0563C1", "folHlink": "954F72",
    "bg1": "FFFFFF", "bg2": "E7E6E6", "tx1": "000000", "tx2": "44546A",
    "phClr": "FFFFFF",
}

THEME_FILL_COLORS = [
    "FFFFFF", "000000", "EEECE1", "1F497D",
    "4BACC6", "4472C4", "9BBB59", "F79646",
    "FFFF00", "A9D18E", "5B9BD5", "70AD47",
]

INDEXED_COLORS = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
    "FFFFFF", "FFFFFF",
]

# ═══════════════════════════════════════════════════════════════════════════════
#  Shape style mapping  (Excel preset geometry → draw.io style)
# ═══════════════════════════════════════════════════════════════════════════════

GEOM_STYLES = {
    "roundRect":                  "rounded=1;arcSize=10;",
    "ellipse":                    "ellipse;",
    "diamond":                    "rhombus;",
    "triangle":                   "triangle;",
    "parallelogram":              "parallelogram;",
    "trapezoid":                  "trapezoid;",
    "hexagon":                    "hexagon;",
    "octagon":                    "octagon;",
    "flowChartOffpageConnector":  "shape=offPageConnector;",
    "flowChartProcess":           "shape=mxgraph.flowchart.process;",
    "flowChartDecision":          "shape=mxgraph.flowchart.decision;",
    "flowChartTerminator":        "shape=mxgraph.flowchart.terminator;",
    "flowChartManualInput":       "shape=mxgraph.flowchart.manual_input;",
    "flowChartDocument":          "shape=mxgraph.flowchart.document;",
    "flowChartPredefinedProcess": "shape=mxgraph.flowchart.predefined_process;",
    "flowChartConnector":         "ellipse;",
    "flowChartPunchedTape":       "shape=mxgraph.flowchart.punched_tape;",
    "flowChartSort":              "shape=mxgraph.flowchart.sort;",
    "homePlate":                  "shape=offPageConnector;",
    "pentagon":                   "shape=offPageConnector;",
    "wedgeRoundRectCallout":      "shape=callout;rounded=1;",
    "wedgeRectCallout":           "shape=callout;",
    "cloudCallout":               "shape=callout;rounded=1;",
    "bentArrow":                  "shape=mxgraph.arrows2.bent_arrow;",
    "chevron":                    "shape=mxgraph.arrows2.arrow;dy=0.6;dx=20;notch=0;",
    "rightArrow":                 "shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=east;",
    "leftArrow":                  "shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=west;",
    "upArrow":                    "shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=north;",
    "downArrow":                  "shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=south;",
}

FONT_ALIASES = {
    "ＭＳ ゴシック": "MS PGothic",
    "ＭＳ Ｐゴシック": "MS PGothic",
    "MS Gothic": "MS PGothic",
    "MS PGothic": "MS PGothic",
    "ＭＳ 明朝": "MS PMincho",
    "ＭＳ Ｐ明朝": "MS PMincho",
    "游ゴシック": "Yu Gothic",
    "游ゴシック Light": "Yu Gothic Light",
    "游明朝": "Yu Mincho",
    "メイリオ": "Meiryo",
    "Meiryo": "Meiryo",
}

OFFPAGE_LABEL_RE = re.compile(r"[A-Z]{1,2}\d?|\d{1,2}")


# ═══════════════════════════════════════════════════════════════════════════════
#  Result dataclass (public API)
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class ConversionResult:
    """Summary of a conversion run."""
    input_path: Path
    output_path: Path
    sheet_names: List[str]
    sheets_data: Dict[str, Any] = field(default_factory=dict)


# ═══════════════════════════════════════════════════════════════════════════════
#  Low-level utilities
# ═══════════════════════════════════════════════════════════════════════════════

def _emu_px(emu: int) -> float:
    return emu / EMU_PER_PX


def _chars_px(c: float) -> int:
    return max(1, int(c * CHAR_WIDTH + 0.5))


def _pts_px(pts: float) -> int:
    return round(pts * POINT_TO_PX)


def _col_letter_to_idx(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def _cell_ref(ref: str) -> Tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref)
    if not m:
        raise ValueError(f"Invalid cell ref: {ref}")
    return _col_letter_to_idx(m.group(1)), int(m.group(2)) - 1


def _parse_range_ref(ref: str) -> Tuple[int, int, int, int]:
    if ":" not in ref:
        c, r = _cell_ref(ref)
        return c, r, c, r
    m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
    if not m:
        raise ValueError(ref)
    return (
        _col_letter_to_idx(m.group(1)), int(m.group(2)) - 1,
        _col_letter_to_idx(m.group(3)), int(m.group(4)) - 1,
    )


def _is_filler(text: str) -> bool:
    return len(text.strip()) == 0


def _normalize_font_name(name: Optional[str]) -> Optional[str]:
    if not name:
        return None
    return FONT_ALIASES.get(name, name)


def _apply_tint(hex6: str, tint: float) -> str:
    try:
        r = int(hex6[0:2], 16)
        g = int(hex6[2:4], 16)
        b = int(hex6[4:6], 16)
        t = float(tint)
        if t > 0:
            r = int(r + (255 - r) * t)
            g = int(g + (255 - g) * t)
            b = int(b + (255 - b) * t)
        else:
            r = int(r * (1 + t))
            g = int(g * (1 + t))
            b = int(b * (1 + t))
        r, g, b = max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b))
        return f"{r:02X}{g:02X}{b:02X}"
    except Exception:
        return hex6


def _should_skip_fill(hex6: str) -> bool:
    return hex6.upper().lstrip("#") in SKIP_FILL_COLORS


def _log(msg: str, verbose: bool = True) -> None:
    if verbose:
        sys.stdout.buffer.write((msg + "\n").encode("utf-8", errors="replace"))


# ═══════════════════════════════════════════════════════════════════════════════
#  Color parsing  (DrawingML + Cell styles)
# ═══════════════════════════════════════════════════════════════════════════════

def _parse_color_element(color_el: Optional[ET.Element], default: Optional[str] = "#000000") -> Optional[str]:
    """Parse fgColor / bgColor / color element → '#RRGGBB' (with tint correction)."""
    if color_el is None:
        return default

    rgb = color_el.get("rgb", "")
    if rgb:
        h6 = (rgb[2:] if len(rgb) == 8 else rgb[:6]).upper()
        tint = color_el.get("tint", "")
        if tint:
            h6 = _apply_tint(h6, float(tint))
        return "#" + h6

    theme = color_el.get("theme", "")
    if theme:
        idx = int(theme)
        base = THEME_FILL_COLORS[idx] if idx < len(THEME_FILL_COLORS) else None
        if base:
            tint = color_el.get("tint", "")
            if tint:
                base = _apply_tint(base, float(tint))
            return "#" + base

    indexed = color_el.get("indexed", "")
    if indexed:
        idx = int(indexed)
        if idx == 64:
            return default
        if idx < len(INDEXED_COLORS):
            return "#" + INDEXED_COLORS[idx]

    return default


def _parse_drawml_color(el: Optional[ET.Element]) -> Optional[str]:
    """Parse DrawingML color (srgbClr / schemeClr / sysClr)."""
    if el is None:
        return None
    sf = el.find(f"{{{A_NS}}}solidFill") or el
    s = sf.find(f"{{{A_NS}}}srgbClr")
    if s is not None:
        return "#" + s.get("val", "000000").upper()
    sc = sf.find(f"{{{A_NS}}}schemeClr")
    if sc is not None:
        base = SCHEME_COLORS.get(sc.get("val", "dk1"), "808080")
        lum_mod = sc.find(f"{{{A_NS}}}lumMod")
        lum_off = sc.find(f"{{{A_NS}}}lumOff")
        if lum_mod is not None or lum_off is not None:
            mod = int(lum_mod.get("val", "100000")) / 100000 if lum_mod is not None else 1.0
            off = int(lum_off.get("val", "0")) / 100000 if lum_off is not None else 0.0
            base = _apply_tint(base, mod - 1 + off)
        return "#" + base.upper()
    sy = sf.find(f"{{{A_NS}}}sysClr")
    if sy is not None:
        last = sy.get("lastClr")
        if last:
            return "#" + last.upper()
    return None


# ═══════════════════════════════════════════════════════════════════════════════
#  Drawing shape style helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _sp_fill(sp_pr: ET.Element) -> str:
    if sp_pr.find(f"{{{A_NS}}}noFill") is not None:
        return "none"
    for fill_tag in (f"{{{A_NS}}}solidFill", f"{{{A_NS}}}gradFill", f"{{{A_NS}}}pattFill"):
        fe = sp_pr.find(fill_tag)
        if fe is not None:
            if fill_tag.endswith("solidFill"):
                c = _parse_drawml_color(fe)
            elif fill_tag.endswith("gradFill"):
                gs = fe.find(f".//{{{A_NS}}}gs")
                c = _parse_drawml_color(gs) if gs is not None else None
            else:
                bg = fe.find(f"{{{A_NS}}}bgClr")
                c = _parse_drawml_color(bg) if bg is not None else None
            if c:
                return c
    return "#FFFFFF"


def _sp_line(sp_pr: ET.Element) -> Tuple[str, int]:
    ln = sp_pr.find(f"{{{A_NS}}}ln")
    if ln is None:
        return "#000000", 1
    if ln.find(f"{{{A_NS}}}noFill") is not None:
        return "none", 0
    sf = ln.find(f"{{{A_NS}}}solidFill")
    color = _parse_drawml_color(sf) if sf is not None else "#000000"
    if color is None:
        color = "#000000"
    w_emu = int(ln.get("w", "12700"))
    return color, max(1, round(w_emu / 12700))


def _sp_geom(sp_pr: ET.Element) -> str:
    g = sp_pr.find(f"{{{A_NS}}}prstGeom")
    return g.get("prst", "rect") if g is not None else "rect"


def _sp_fontsize(txb: Optional[ET.Element]) -> int:
    if txb is None:
        return 9
    for tag in (f"{{{A_NS}}}rPr", f"{{{A_NS}}}endParaRPr"):
        e = txb.find(f".//{tag}")
        if e is not None:
            sz = e.get("sz")
            if sz:
                return max(7, round(int(sz) / 100))
    return 9


def _sp_font_style(txb: Optional[ET.Element]) -> Dict[str, Any]:
    if txb is None:
        return {}
    rpr = txb.find(f".//{{{A_NS}}}rPr") or txb.find(f".//{{{A_NS}}}endParaRPr")
    if rpr is None:
        return {}
    extra: Dict[str, Any] = {}
    solid = rpr.find(f"{{{A_NS}}}solidFill")
    if solid is not None:
        fc = _parse_drawml_color(solid)
        if fc and fc not in ("#000000", "#FFFFFF"):
            extra["fontColor"] = fc
    fs = 0
    if rpr.get("b") == "1":
        fs |= 1
    if rpr.get("i") == "1":
        fs |= 2
    if fs:
        extra["fontStyle"] = fs
    return extra


def _get_text(el: ET.Element) -> str:
    return "".join(t.text for t in el.iter(f"{{{A_NS}}}t") if t.text)


def _make_shape_style(prst: str, fill: str, lc: str, lw: int, fsz: int,
                      font_extra: Optional[Dict] = None) -> str:
    parts = ["whiteSpace=wrap", "html=1"]
    extra = GEOM_STYLES.get(prst, "")
    if extra:
        parts.append(extra.rstrip(";"))
    parts.append(f"fillColor={fill}" if fill != "none" else "fillColor=none")
    parts.append(f"strokeColor={lc}" if lc != "none" else "strokeColor=none")
    if lw > 1:
        parts.append(f"strokeWidth={lw}")
    if fsz != 9:
        parts.append(f"fontSize={fsz}")
    if font_extra:
        if "fontColor" in font_extra:
            parts.append(f"fontColor={font_extra['fontColor']}")
        if "fontStyle" in font_extra:
            parts.append(f"fontStyle={font_extra['fontStyle']}")
    return ";".join(parts) + ";"


# ═══════════════════════════════════════════════════════════════════════════════
#  Grid builder  (column widths / row heights → pixel coordinates)
# ═══════════════════════════════════════════════════════════════════════════════

def _build_grid(sh_root: ET.Element) -> Tuple[List[int], List[int], dict, dict]:
    col_w: dict = defaultdict(lambda: 8.0)
    for col_el in sh_root.findall(".//x:col", _NS_X):
        mn = int(col_el.get("min", "1"))
        mx = int(col_el.get("max", "1"))
        w = float(col_el.get("width", "8"))
        for c in range(mn - 1, mx):
            col_w[c] = w

    row_h: dict = defaultdict(lambda: 15.0)
    for row_el in sh_root.findall(".//x:row", _NS_X):
        r = int(row_el.get("r", "1"))
        ht = row_el.get("ht")
        if ht:
            row_h[r - 1] = float(ht)

    col_x = [0] * (GRID_MAX + 1)
    for i in range(GRID_MAX):
        col_x[i + 1] = col_x[i] + _chars_px(col_w[i])

    row_y = [0] * (GRID_MAX + 1)
    for i in range(GRID_MAX):
        row_y[i + 1] = row_y[i] + _pts_px(row_h[i])

    return col_x, row_y, col_w, row_h


# ═══════════════════════════════════════════════════════════════════════════════
#  DrawIO XML builder
# ═══════════════════════════════════════════════════════════════════════════════

class _DrawioBuilder:
    def __init__(self) -> None:
        self._cells: List[str] = []
        self._next: int = 2
        self._seen: set = set()
        self._max_x: int = 0
        self._max_y: int = 0

    def add(self, text: str, x: float, y: float, w: float, h: float,
            style: str, force: bool = False) -> None:
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
        esc = html_mod.escape(str(text))
        self._cells.append(
            f'    <mxCell id="{cid}" value="{esc}" style="{style}" '
            f'vertex="1" parent="1">'
            f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" '
            f'as="geometry"/></mxCell>'
        )

    @property
    def count(self) -> int:
        return self._next - 2

    def xml(self, sheet_name: str) -> str:
        page_w = max(2000, int(self._max_x * 1.10))
        page_h = max(2000, int(self._max_y * 1.10))
        hdr = (
            f'  <diagram id="d1" name="{html_mod.escape(sheet_name)}">\n'
            f'    <mxGraphModel grid="0" guides="1" tooltips="1" connect="1" '
            f'arrows="1"\n'
            f'                  fold="1" page="1" pageScale="1" '
            f'pageWidth="{page_w}"\n'
            f'                  pageHeight="{page_h}" math="0" shadow="0">\n'
            f"      <root>\n"
            f'        <mxCell id="0"/>\n'
            f'        <mxCell id="1" parent="0"/>\n'
        )
        ftr = "      </root>\n    </mxGraphModel>\n  </diagram>\n"
        return hdr + "\n".join(self._cells) + "\n" + ftr


# ═══════════════════════════════════════════════════════════════════════════════
#  styles.xml parsers
# ═══════════════════════════════════════════════════════════════════════════════

def _parse_cell_styles(zf: zipfile.ZipFile) -> Dict[int, str]:
    """xf_index → fill color '#RRGGBB'."""
    xf_fills: Dict[int, str] = {}
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except Exception:
        return xf_fills

    fills: List[Optional[str]] = []
    for fill_el in root.findall(".//x:fills/x:fill", _NS_X):
        color = None
        pf = fill_el.find(f"{{{SS_NS}}}patternFill")
        if pf is not None and pf.get("patternType", "none") != "none":
            fg = pf.find(f"{{{SS_NS}}}fgColor")
            if fg is not None:
                c = _parse_color_element(fg, default=None)
                if c and not _should_skip_fill(c):
                    color = c
        fills.append(color)

    for i, xf in enumerate(root.findall(".//x:cellXfs/x:xf", _NS_X)):
        fill_id = int(xf.get("fillId", "0"))
        if fill_id < len(fills) and fills[fill_id]:
            xf_fills[i] = fills[fill_id]

    return xf_fills


def _parse_cell_borders(zf: zipfile.ZipFile) -> Dict[int, Dict[str, Tuple[str, int]]]:
    """xf_index → {side: (color, width_px)}."""
    xf_borders: Dict[int, Dict[str, Tuple[str, int]]] = {}
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except Exception:
        return xf_borders

    def _bw(style_name: str) -> int:
        if style_name in ("medium", "mediumDashed", "mediumDashDot",
                          "mediumDashDotDot", "slantDashDot"):
            return 2
        if style_name == "thick":
            return 3
        return 1

    border_defs: List[Dict] = []
    for bel in root.findall(".//x:borders/x:border", _NS_X):
        sides: Dict[str, Tuple[str, int]] = {}
        for side in ("left", "right", "top", "bottom"):
            sel = bel.find(f"{{{SS_NS}}}{side}")
            if sel is None:
                continue
            sname = sel.get("style")
            if not sname:
                continue
            color = _parse_color_element(sel.find(f"{{{SS_NS}}}color"))
            sides[side] = (color, _bw(sname))
        border_defs.append(sides)

    for i, xf in enumerate(root.findall(".//x:cellXfs/x:xf", _NS_X)):
        bid = int(xf.get("borderId", "0"))
        if 0 <= bid < len(border_defs) and border_defs[bid]:
            xf_borders[i] = border_defs[bid]

    return xf_borders


def _parse_cell_text_styles(zf: zipfile.ZipFile) -> Dict[int, Dict]:
    """xf_index → text style dict."""
    xf_text_styles: Dict[int, Dict] = {}
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except Exception:
        return xf_text_styles

    fonts: List[Dict] = []
    for font_el in root.findall(".//x:fonts/x:font", _NS_X):
        name_el = font_el.find(f"{{{SS_NS}}}name")
        size_el = font_el.find(f"{{{SS_NS}}}sz")
        color_el = font_el.find(f"{{{SS_NS}}}color")
        bold = font_el.find(f"{{{SS_NS}}}b") is not None
        italic = font_el.find(f"{{{SS_NS}}}i") is not None
        fonts.append({
            "fontFamily": _normalize_font_name(name_el.get("val")) if name_el is not None else None,
            "fontSize": max(6, round(float(size_el.get("val", "11")))) if size_el is not None else 11,
            "fontColor": _parse_color_element(color_el, default="#000000"),
            "bold": bold,
            "italic": italic,
        })

    for i, xf in enumerate(root.findall(".//x:cellXfs/x:xf", _NS_X)):
        style: Dict = {}
        fid = int(xf.get("fontId", "0"))
        if 0 <= fid < len(fonts):
            f = fonts[fid]
            if f.get("fontFamily"):
                style["fontFamily"] = f["fontFamily"]
            if f.get("fontSize"):
                style["fontSize"] = f["fontSize"]
            if f.get("fontColor") and f["fontColor"] != "#000000":
                style["fontColor"] = f["fontColor"]
            fs = 0
            if f.get("bold"):
                fs |= 1
            if f.get("italic"):
                fs |= 2
            if fs:
                style["fontStyle"] = fs

        al = xf.find(f"{{{SS_NS}}}alignment")
        if al is not None:
            h = al.get("horizontal")
            v = al.get("vertical")
            if h in ("left", "center", "right"):
                style["align"] = h
            if v in ("top", "center", "bottom"):
                style["verticalAlign"] = {"center": "middle"}.get(v, v)
            if al.get("wrapText") == "1":
                style["wrapText"] = True

        xf_text_styles[i] = style

    return xf_text_styles


def _parse_cell_number_formats(zf: zipfile.ZipFile) -> Dict[int, Tuple[int, str]]:
    """xf_index → (numFmtId, formatCode)."""
    xf_numfmts: Dict[int, Tuple[int, str]] = {}
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except Exception:
        return xf_numfmts

    custom = {
        int(el.get("numFmtId", "0")): el.get("formatCode", "")
        for el in root.findall(".//x:numFmts/x:numFmt", _NS_X)
    }
    for i, xf in enumerate(root.findall(".//x:cellXfs/x:xf", _NS_X)):
        nid = int(xf.get("numFmtId", "0"))
        xf_numfmts[i] = (nid, custom.get(nid, ""))

    return xf_numfmts


def _parse_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except Exception:
        return []
    return [
        "".join(t.text for t in si.iter(f"{{{SS_NS}}}t") if t.text)
        for si in root.findall(f"{{{SS_NS}}}si")
    ]


# ═══════════════════════════════════════════════════════════════════════════════
#  Merged cell maps
# ═══════════════════════════════════════════════════════════════════════════════

def _build_merged_cell_maps(sh_root: ET.Element) -> Tuple[Dict, set]:
    merged_topleft: Dict[Tuple[int, int], Tuple[int, int]] = {}
    merged_children: set = set()
    for mc in sh_root.findall(".//x:mergeCell", _NS_X):
        ref = mc.get("ref", "")
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


# ═══════════════════════════════════════════════════════════════════════════════
#  Cell fills  (merge adjacent same-color cells into rectangles)
# ═══════════════════════════════════════════════════════════════════════════════

def _add_cell_fills_merged(sh_root: ET.Element, col_x: List[int], row_y: List[int],
                           xf_fills: Dict[int, str], bld: _DrawioBuilder) -> int:
    color_grid: Dict[Tuple[int, int], str] = {}

    for row_el in sh_root.findall(".//x:row", _NS_X):
        r = int(row_el.get("r", "1")) - 1
        for cell in row_el.findall("x:c", _NS_X):
            ref = cell.get("r", "")
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            s_attr = int(cell.get("s", "0"))
            fc = xf_fills.get(s_attr)
            if fc:
                color_grid[(r, c)] = fc

    merged_topleft, _ = _build_merged_cell_maps(sh_root)
    for (r1, c1), (r2, c2) in merged_topleft.items():
        color = color_grid.get((r1, c1))
        if color:
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    if (rr, cc) not in color_grid:
                        color_grid[(rr, cc)] = color

    if not color_grid:
        return 0

    processed: set = set()
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
            if all(
                color_grid.get((nr, cc)) == color and (nr, cc) not in processed
                for cc in range(c, c_end + 1)
            ):
                r_end = nr
            else:
                break

        for rr in range(r, r_end + 1):
            for cc in range(c, c_end + 1):
                processed.add((rr, cc))

        px = max(0.0, col_x[min(c, GRID_MAX)] - CELL_BOX_LEFT_PAD)
        py = row_y[min(r, GRID_MAX)]
        px_end = col_x[min(c_end + 1, GRID_MAX)]
        py_end = row_y[min(r_end + 1, GRID_MAX)]
        w = max(2.0, px_end - px)
        h = max(2.0, py_end - py)

        style = f"whiteSpace=wrap;html=1;fillColor={color};strokeColor=none;"
        bld.add("", px, py, w, h, style)
        count += 1

    return count


# ═══════════════════════════════════════════════════════════════════════════════
#  Cell borders
# ═══════════════════════════════════════════════════════════════════════════════

def _add_cell_borders(sh_root: ET.Element, col_x: List[int], row_y: List[int],
                      col_w: dict, row_h: dict,
                      xf_borders: Dict, xf_fills: Dict, bld: _DrawioBuilder) -> int:
    count = 0
    BORDER_MARGIN = 10

    row_active: Dict[int, Tuple[int, int]] = {}
    filled_positions: set = set()
    for row_el in sh_root.findall(".//x:row", _NS_X):
        r = int(row_el.get("r", "1")) - 1
        cols: List[int] = []
        for cell in row_el.findall("x:c", _NS_X):
            ref = cell.get("r", "")
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            s_attr = int(cell.get("s", "0"))
            v_el = cell.find("x:v", _NS_X)
            if (v_el is not None and v_el.text is not None) or xf_fills.get(s_attr):
                cols.append(c)
            if xf_fills.get(s_attr):
                filled_positions.add((r, c))
        if cols:
            row_active[r] = (min(cols), max(cols) + BORDER_MARGIN)

    for row_el in sh_root.findall(".//x:row", _NS_X):
        r = int(row_el.get("r", "1")) - 1
        cy = row_y[min(r, GRID_MAX - 1)]
        ch = max(1.0, _pts_px(row_h[r]))

        for cell in row_el.findall("x:c", _NS_X):
            ref = cell.get("r", "")
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue

            s_attr = int(cell.get("s", "0"))
            border_info = xf_borders.get(s_attr)
            if not border_info:
                continue

            v_el = cell.find("x:v", _NS_X)
            has_value = v_el is not None and v_el.text is not None
            has_fill = xf_fills.get(s_attr) is not None
            if not has_value and not has_fill:
                active = row_active.get(r)
                if active is None or c < active[0] or c > active[1]:
                    continue

            cx = col_x[min(c, GRID_MAX - 1)]
            cw = max(1.0, _chars_px(col_w[c]))
            bx = max(0.0, cx - CELL_BOX_LEFT_PAD)
            bw = cw + min(CELL_BOX_LEFT_PAD, cx)

            cell_fill_color = xf_fills.get(s_attr)
            for side, (color, width_px) in border_info.items():
                if side == "left" and cell_fill_color and (r, c - 1) in filled_positions:
                    continue
                if side == "right" and cell_fill_color and (r, c + 1) in filled_positions:
                    continue
                style = f"whiteSpace=wrap;html=1;fillColor={color};strokeColor={color};"
                if side == "top":
                    bld.add("", bx, cy, bw, width_px, style)
                elif side == "bottom":
                    bld.add("", bx, cy + ch - width_px, bw, width_px, style)
                elif side == "left":
                    bld.add("", bx, cy, width_px, ch, style)
                elif side == "right":
                    bld.add("", cx + cw - width_px, cy, width_px, ch, style)
                count += 1

    return count


# ═══════════════════════════════════════════════════════════════════════════════
#  Drawing shapes  (DrawingML sp / grpSp / cxnSp)
# ═══════════════════════════════════════════════════════════════════════════════

def _get_xfrm(xfrm: ET.Element) -> Tuple[int, int, int, int, int, int, int, int]:
    def iv(el: Optional[ET.Element], attr: str, default: int = 0) -> int:
        return int(el.get(attr, str(default))) if el is not None else default

    off = xfrm.find(f"{{{A_NS}}}off")
    ext = xfrm.find(f"{{{A_NS}}}ext")
    choff = xfrm.find(f"{{{A_NS}}}chOff")
    chext = xfrm.find(f"{{{A_NS}}}chExt")
    ox, oy = iv(off, "x"), iv(off, "y")
    ecx, ecy = iv(ext, "cx"), iv(ext, "cy")
    chox, choy = iv(choff, "x", ox), iv(choff, "y", oy)
    chcx, chcy = iv(chext, "cx", ecx), iv(chext, "cy", ecy)
    return ox, oy, ecx, ecy, chox, choy, chcx, chcy


def _emit_sp(sp: ET.Element, pax: float, pay: float,
             sx: float, sy: float, bld: _DrawioBuilder) -> None:
    spr = sp.find(f"{{{XDR_NS}}}spPr")
    if spr is None:
        return
    xfrm = spr.find(f"{{{A_NS}}}xfrm")
    if xfrm is None:
        return
    off = xfrm.find(f"{{{A_NS}}}off")
    ext = xfrm.find(f"{{{A_NS}}}ext")
    if off is None or ext is None:
        return

    ax = pax + int(off.get("x", "0")) * sx
    ay = pay + int(off.get("y", "0")) * sy
    w = int(ext.get("cx", "0")) * sx
    h = int(ext.get("cy", "0")) * sy

    if w < 1 or h < 1:
        return

    text = _get_text(sp)
    fill = _sp_fill(spr)
    lc, lw = _sp_line(spr)
    prst = _sp_geom(spr)
    txb = sp.find(f"{{{XDR_NS}}}txBody")
    fsz = _sp_fontsize(txb)
    fe = _sp_font_style(txb)

    if not text and fill in ("#FFFFFF", "none") and lc == "none":
        return

    text_s = text.strip()
    if text_s and w < 80 and h < 80 and OFFPAGE_LABEL_RE.fullmatch(text_s):
        prst = "homePlate"
        if fill == "none":
            fill = "#FFFFFF"
        if lc == "none":
            lc, lw = "#000000", 1

    style = _make_shape_style(prst, fill, lc, lw, fsz, fe)
    bld.add(text, ax, ay, w, h, style, force=bool(text))


def _emit_cxnsp(cxn: ET.Element, pax: float, pay: float,
                sx: float, sy: float, bld: _DrawioBuilder) -> None:
    spr = cxn.find(f"{{{XDR_NS}}}spPr")
    if spr is None:
        return
    xfrm = spr.find(f"{{{A_NS}}}xfrm")
    if xfrm is None:
        return
    off = xfrm.find(f"{{{A_NS}}}off")
    ext = xfrm.find(f"{{{A_NS}}}ext")
    if off is None or ext is None:
        return

    ax = pax + int(off.get("x", "0")) * sx
    ay = pay + int(off.get("y", "0")) * sy
    raw_w = int(ext.get("cx", "0")) * sx
    raw_h = int(ext.get("cy", "0")) * sy

    w = raw_w if raw_w >= 1 else 2
    h = raw_h if raw_h >= 1 else 2

    ln = spr.find(f"{{{A_NS}}}ln")
    if ln is not None and ln.find(f"{{{A_NS}}}noFill") is not None:
        return

    if ln is not None:
        sf = ln.find(f"{{{A_NS}}}solidFill")
        color = _parse_drawml_color(sf) if sf is not None else "#000000"
        if color is None:
            color = "#000000"
    else:
        color = "#000000"

    lw_emu = int(ln.get("w", "12700")) if ln is not None else 12700
    lw_px = max(1, round(lw_emu / 12700))

    if raw_w < 1 or raw_h < 1:
        style = (f"whiteSpace=wrap;html=1;fillColor={color};"
                 f"strokeColor={color};strokeWidth={lw_px};")
    else:
        style = (f"whiteSpace=wrap;html=1;fillColor=none;"
                 f"strokeColor={color};strokeWidth={lw_px};")

    bld.add("", ax, ay, w, h, style)


def _walk_group(grp: ET.Element, pax: float, pay: float,
                sx: float, sy: float, bld: _DrawioBuilder, depth: int = 0) -> None:
    if depth > 25:
        return
    grp_pr = grp.find(f"{{{XDR_NS}}}grpSpPr")
    if grp_pr is None:
        return
    xfrm = grp_pr.find(f"{{{A_NS}}}xfrm")
    if xfrm is None:
        return

    ox, oy, ecx, ecy, chox, choy, chcx, chcy = _get_xfrm(xfrm)
    gax, gay = pax + ox * sx, pay + oy * sy
    gw, gh = ecx * sx, ecy * sy

    csx = (gw / chcx) if chcx else sx
    csy = (gh / chcy) if chcy else sy
    cox = gax - chox * csx
    coy = gay - choy * csy

    for child in grp:
        ct = child.tag.split("}")[-1]
        if ct == "sp":
            _emit_sp(child, cox, coy, csx, csy, bld)
        elif ct == "cxnSp":
            _emit_cxnsp(child, cox, coy, csx, csy, bld)
        elif ct == "grpSp":
            _walk_group(child, cox, coy, csx, csy, bld, depth + 1)


def _anchor_rect(anchor: ET.Element, col_x: List[int], row_y: List[int]) -> Optional[Tuple[float, float, float, float]]:
    from_el = anchor.find(f"{{{XDR_NS}}}from")
    if from_el is None:
        return None

    fc = int(from_el.findtext(f"{{{XDR_NS}}}col", "0") or "0")
    fco = int(from_el.findtext(f"{{{XDR_NS}}}colOff", "0") or "0")
    fr = int(from_el.findtext(f"{{{XDR_NS}}}row", "0") or "0")
    fro = int(from_el.findtext(f"{{{XDR_NS}}}rowOff", "0") or "0")

    anc_x = col_x[min(fc, GRID_MAX - 1)] + _emu_px(fco)
    anc_y = row_y[min(fr, GRID_MAX - 1)] + _emu_px(fro)

    to_el = anchor.find(f"{{{XDR_NS}}}to")
    ext_el = anchor.find(f"{{{XDR_NS}}}ext")

    if to_el is not None:
        tc = int(to_el.findtext(f"{{{XDR_NS}}}col", "0") or "0")
        tco = int(to_el.findtext(f"{{{XDR_NS}}}colOff", "0") or "0")
        tr = int(to_el.findtext(f"{{{XDR_NS}}}row", "0") or "0")
        tro = int(to_el.findtext(f"{{{XDR_NS}}}rowOff", "0") or "0")
        anc_w = max(2.0, col_x[min(tc, GRID_MAX - 1)] + _emu_px(tco) - anc_x)
        anc_h = max(2.0, row_y[min(tr, GRID_MAX - 1)] + _emu_px(tro) - anc_y)
    elif ext_el is not None:
        anc_w = max(2.0, _emu_px(int(ext_el.get("cx", "9525"))))
        anc_h = max(2.0, _emu_px(int(ext_el.get("cy", "9525"))))
    else:
        anc_w, anc_h = 80.0, 24.0

    return anc_x, anc_y, anc_w, anc_h


def _add_drawing_shapes(zf: zipfile.ZipFile, drawing_path: str,
                        col_x: List[int], row_y: List[int],
                        bld: _DrawioBuilder) -> None:
    dr = ET.fromstring(zf.read(drawing_path))
    sc = 1.0 / EMU_PER_PX

    for anchor in dr:
        tag = anchor.tag.split("}")[-1]
        if tag not in ("oneCellAnchor", "twoCellAnchor"):
            continue

        rect = _anchor_rect(anchor, col_x, row_y)
        if rect is None:
            continue
        anc_x, anc_y, anc_w, anc_h = rect

        for child in anchor:
            ct = child.tag.split("}")[-1]

            if ct == "sp":
                spr = child.find(f"{{{XDR_NS}}}spPr")
                if spr is None:
                    continue
                text = _get_text(child)
                fill = _sp_fill(spr)
                lc, lw = _sp_line(spr)
                prst = _sp_geom(spr)
                txb = child.find(f"{{{XDR_NS}}}txBody")
                fsz = _sp_fontsize(txb)
                fe = _sp_font_style(txb)
                if not text and fill in ("#FFFFFF", "none") and lc == "none":
                    continue
                text_s = text.strip()
                if text_s and anc_w < 80 and anc_h < 80 and OFFPAGE_LABEL_RE.fullmatch(text_s):
                    prst = "homePlate"
                    if fill == "none":
                        fill = "#FFFFFF"
                    if lc == "none":
                        lc, lw = "#000000", 1
                style = _make_shape_style(prst, fill, lc, lw, fsz, fe)
                bld.add(text, anc_x, anc_y, anc_w, anc_h, style, force=bool(text))

            elif ct == "grpSp":
                grp_pr = child.find(f"{{{XDR_NS}}}grpSpPr")
                if grp_pr is None:
                    continue
                xfrm = grp_pr.find(f"{{{A_NS}}}xfrm")
                if xfrm is None:
                    continue
                _, _, ecx, ecy, chox, choy, chcx, chcy = _get_xfrm(xfrm)
                csx = (anc_w / chcx) if chcx else sc
                csy = (anc_h / chcy) if chcy else sc
                cox = anc_x - chox * csx
                coy = anc_y - choy * csy
                for grandchild in child:
                    gct = grandchild.tag.split("}")[-1]
                    if gct == "sp":
                        _emit_sp(grandchild, cox, coy, csx, csy, bld)
                    elif gct == "grpSp":
                        _walk_group(grandchild, cox, coy, csx, csy, bld)
                    elif gct == "cxnSp":
                        _emit_cxnsp(grandchild, cox, coy, csx, csy, bld)

            elif ct == "cxnSp":
                _emit_cxnsp(child, 0, 0, sc, sc, bld)


# ═══════════════════════════════════════════════════════════════════════════════
#  Cell text labels
# ═══════════════════════════════════════════════════════════════════════════════

def _build_cell_value_map(sh_root: ET.Element, shared_strings: List[str]) -> Dict[Tuple[int, int], str]:
    value_map: Dict[Tuple[int, int], str] = {}
    for row_el in sh_root.findall(".//x:row", _NS_X):
        r = int(row_el.get("r", "1")) - 1
        for cell in row_el.findall("x:c", _NS_X):
            ref = cell.get("r", "")
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            t = cell.get("t", "")
            v_el = cell.find("x:v", _NS_X)
            if v_el is None or v_el.text is None:
                value_map[(r, c)] = ""
                continue
            if t == "s":
                idx = int(v_el.text)
                value_map[(r, c)] = shared_strings[idx] if idx < len(shared_strings) else ""
            else:
                value_map[(r, c)] = v_el.text
    return value_map


def _build_fill_grid(sh_root: ET.Element, xf_fills: Dict[int, str]) -> Dict[Tuple[int, int], str]:
    grid: Dict[Tuple[int, int], str] = {}
    for row_el in sh_root.findall(".//x:row", _NS_X):
        r = int(row_el.get("r", "1")) - 1
        for cell in row_el.findall("x:c", _NS_X):
            ref = cell.get("r", "")
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            s_attr = int(cell.get("s", "0"))
            fc = xf_fills.get(s_attr)
            if fc:
                grid[(r, c)] = fc
    return grid


def _format_excel_time(value: float) -> str:
    total_minutes = int(round(value * 24 * 60))
    return f"{total_minutes // 60}:{total_minutes % 60:02d}"


def _format_numeric_value(raw: str, style_numfmt: Tuple[int, str]) -> str:
    try:
        fv = float(raw)
    except ValueError:
        return raw
    num_fmt_id, fmt_code = style_numfmt
    fmt = (fmt_code or "").lower()
    is_time = (num_fmt_id in {18, 19, 20, 21, 22, 45, 46, 47}
               or ("h" in fmt and "m" in fmt))
    if is_time:
        return _format_excel_time(fv)
    return str(int(fv)) if fv.is_integer() else raw


def _estimate_text_units(text: str) -> float:
    units = 0.0
    for ch in text:
        code = ord(ch)
        if ch in "ilI1.:;| ":
            units += 0.35
        elif code < 128:
            units += 0.6
        else:
            units += 1.0
    return max(units, 1.0)


def _fit_font_size(text: str, width: float, height: float, base_font_size: int) -> int:
    font_size = max(6, base_font_size)
    while font_size > 6:
        line_cap = max(1.0, (width - 2) / max(font_size * 0.95, 1))
        req_lines = ceil(_estimate_text_units(text) / line_cap)
        max_lines = max(1, int(height / max(font_size * 1.15, 1)))
        if req_lines <= max_lines:
            break
        font_size -= 1
    return font_size


def _is_compact_label(text: str) -> bool:
    s = str(text).strip()
    if re.fullmatch(r"\d{1,2}[:：]\d{2}", s):
        return True
    if re.fullmatch(r"\d+", s) and len(s) <= 2:
        return True
    return False


def _make_cell_text_style(style_info: Dict, text: str, width: float,
                          height: float, compact: bool = False) -> str:
    eff = dict(style_info)
    if compact:
        eff["align"] = "center"
        eff["verticalAlign"] = "middle"
    fsz = _fit_font_size(text, width, height, eff.get("fontSize", 10))
    parts = [
        "text", "html=1", "strokeColor=none", "fillColor=none",
        "whiteSpace=wrap",
        f"overflow={'hidden' if compact else 'fill'}",
        f"align={eff.get('align', 'left')}",
        f"verticalAlign={eff.get('verticalAlign', 'middle')}",
        f"fontSize={fsz}",
    ]
    if eff.get("fontFamily"):
        parts.append(f"fontFamily={eff['fontFamily']}")
    if eff.get("fontColor"):
        parts.append(f"fontColor={eff['fontColor']}")
    if eff.get("fontStyle"):
        parts.append(f"fontStyle={eff['fontStyle']}")
    parts.append("spacingTop=1" if compact else "spacingTop=3")
    if not compact and eff.get("align", "left") == "left":
        parts.append("spacingLeft=5")
    return ";".join(parts) + ";"


def _add_cell_labels(sh_root: ET.Element, col_x: List[int], row_y: List[int],
                     col_w: dict, row_h: dict, shared_strings: List[str],
                     xf_text_styles: Dict, xf_numfmts: Dict,
                     xf_fills: Dict, bld: _DrawioBuilder) -> None:
    merged_topleft, merged_children = _build_merged_cell_maps(sh_root)
    value_map = _build_cell_value_map(sh_root, shared_strings)
    fill_grid = _build_fill_grid(sh_root, xf_fills)

    for row_el in sh_root.findall(".//x:row", _NS_X):
        r = int(row_el.get("r", "1")) - 1
        ry = row_y[min(r, GRID_MAX - 1)]
        rh = max(1.0, _pts_px(row_h[r]))

        for cell in row_el.findall("x:c", _NS_X):
            ref = cell.get("r", "")
            if not ref:
                continue
            try:
                c, _ = _cell_ref(ref)
            except Exception:
                continue
            if (r, c) in merged_children:
                continue

            t = cell.get("t", "")
            v_el = cell.find("x:v", _NS_X)
            if v_el is None or v_el.text is None:
                continue

            if t == "s":
                idx = int(v_el.text)
                val = shared_strings[idx] if idx < len(shared_strings) else ""
            elif t == "str":
                val = v_el.text
            else:
                s_attr = int(cell.get("s", "0"))
                val = _format_numeric_value(
                    v_el.text, xf_numfmts.get(s_attr, (0, ""))
                )

            if not val or _is_filler(val):
                continue

            cx = col_x[min(c, GRID_MAX - 1)]
            s_attr = int(cell.get("s", "0"))
            style_info = xf_text_styles.get(s_attr, {})
            compact = _is_compact_label(val)

            if (r, c) in merged_topleft:
                r_end, c_end = merged_topleft[(r, c)]
                cw = max(1.0, col_x[min(c_end + 1, GRID_MAX)] - col_x[min(c, GRID_MAX)])
                ch = max(1.0, row_y[min(r_end + 1, GRID_MAX)] - row_y[min(r, GRID_MAX)])
            else:
                c_end = c
                fill_color = fill_grid.get((r, c))
                if not compact:
                    fsz_est = style_info.get("fontSize", 10)
                    approx_px_needed = _estimate_text_units(str(val)) * fsz_est * 0.72 + 10
                    px_accumulated = _chars_px(col_w[c])
                    while c_end + 1 <= GRID_MAX:
                        next_val = value_map.get((r, c_end + 1), "")
                        next_fill = fill_grid.get((r, c_end + 1))
                        if next_val:
                            break
                        if fill_color and next_fill != fill_color:
                            break
                        if not fill_color and next_fill:
                            break
                        if not fill_color and px_accumulated >= approx_px_needed:
                            break
                        c_end += 1
                        px_accumulated += _chars_px(col_w[c_end])
                cw = max(1.0, col_x[min(c_end + 1, GRID_MAX)] - col_x[min(c, GRID_MAX)])
                ch = rh

            text_x = cx
            if not compact and style_info.get("align", "left") == "left":
                text_x += 2
                cw = max(1.0, cw - 2)

            text_y = ry
            text_h = ch
            if not compact and style_info.get("align", "left") == "left":
                text_y += FILLED_TEXT_TOP_PAD
                text_h = max(1.0, ch - FILLED_TEXT_TOP_PAD)

            cell_style = _make_cell_text_style(style_info, val, cw, text_h, compact=compact)
            bld.add(val, text_x, text_y, cw, text_h, cell_style, force=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  Sheet / drawing path resolution
# ═══════════════════════════════════════════════════════════════════════════════

def _resolve_xl_target(base_path: str, target: str) -> str:
    base_dir = Path(base_path).parent
    joined = (base_dir / target).as_posix()
    parts: List[str] = []
    for p in joined.split("/"):
        if p in ("", "."):
            continue
        if p == "..":
            if parts:
                parts.pop()
            continue
        parts.append(p)
    return "/".join(parts)


def _find_sheet_targets(zf: zipfile.ZipFile,
                        sheet_names: Optional[List[str]]) -> List[Tuple[str, str]]:
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {r.get("Id"): r.get("Target") for r in rels_root}

    selected = set(sheet_names) if sheet_names else None
    targets: List[Tuple[str, str]] = []
    for sheet in wb_root.findall(f".//{{{SS_NS}}}sheet"):
        name = sheet.get("name", "")
        if selected is not None and name not in selected:
            continue
        rid = sheet.get(f"{{{REL_NS}}}id")
        target = rel_map.get(rid)
        if target:
            targets.append((name, _resolve_xl_target("xl/workbook.xml", target)))
    return targets


def _find_drawing_for_sheet(zf: zipfile.ZipFile, sheet_xml_path: str) -> Optional[str]:
    sheet_file = Path(sheet_xml_path).name
    rels_path = f"xl/worksheets/_rels/{sheet_file}.rels"
    if rels_path not in zf.namelist():
        return None

    rels_root = ET.fromstring(zf.read(rels_path))
    for rel in rels_root.findall(f".//{{{PKG_REL_NS}}}Relationship"):
        rel_type = rel.get("Type", "")
        if "drawing" in rel_type and "vml" not in rel_type:
            target = rel.get("Target", "")
            return _resolve_xl_target(sheet_xml_path, target)

    return None


# ═══════════════════════════════════════════════════════════════════════════════
#  Per-sheet conversion
# ═══════════════════════════════════════════════════════════════════════════════

def _convert_sheet(zf: zipfile.ZipFile, sheet_name: str, sheet_xml_path: str,
                   shared_strings: List[str],
                   xf_fills: Dict, xf_borders: Dict,
                   xf_text_styles: Dict, xf_numfmts: Dict,
                   include_cells: bool, verbose: bool) -> str:
    sh_root = ET.fromstring(zf.read(sheet_xml_path))
    col_x, row_y, col_w, row_h = _build_grid(sh_root)

    bld = _DrawioBuilder()

    if include_cells:
        fc = _add_cell_fills_merged(sh_root, col_x, row_y, xf_fills, bld)
        _log(f"  [{sheet_name}] Fill rectangles: {fc}", verbose)

        bc = _add_cell_borders(sh_root, col_x, row_y, col_w, row_h,
                               xf_borders, xf_fills, bld)
        _log(f"  [{sheet_name}] Border segments: {bc}", verbose)

    drawing_path = _find_drawing_for_sheet(zf, sheet_xml_path)
    if drawing_path and drawing_path in zf.namelist():
        before = bld.count
        _add_drawing_shapes(zf, drawing_path, col_x, row_y, bld)
        _log(f"  [{sheet_name}] Drawing shapes: {bld.count - before}", verbose)
    else:
        _log(f"  [{sheet_name}] No drawing found.", verbose)

    if include_cells:
        before = bld.count
        _add_cell_labels(sh_root, col_x, row_y, col_w, row_h,
                         shared_strings, xf_text_styles, xf_numfmts, xf_fills, bld)
        _log(f"  [{sheet_name}] Cell labels: {bld.count - before}", verbose)

    _log(f"  [{sheet_name}] Total elements: {bld.count}", verbose)
    return bld.xml(sheet_name)


# ═══════════════════════════════════════════════════════════════════════════════
#  Public API
# ═══════════════════════════════════════════════════════════════════════════════

def convert_excel_to_drawio(
    input_path: str,
    output_path: str,
    sheet_names: Optional[List[str]] = None,
    include_cells: bool = True,
    verbose: bool = False,
) -> ConversionResult:
    """Convert an Excel file to draw.io XML.

    Args:
        input_path:   Path to the input .xlsx / .xlsm file.
        output_path:  Path for the output .drawio file.
        sheet_names:  Sheets to convert (None = all sheets).
        include_cells: Include cell fills, borders, and text labels.
        verbose:      Print progress to stdout.

    Returns:
        ConversionResult with conversion metadata.
    """
    input_file = Path(input_path)
    output_file = Path(output_path)

    sheets_data: Dict[str, Any] = {}

    with zipfile.ZipFile(input_file, "r") as zf:
        shared_strings = _parse_shared_strings(zf)
        xf_fills = _parse_cell_styles(zf)
        xf_borders = _parse_cell_borders(zf)
        xf_text_styles = _parse_cell_text_styles(zf)
        xf_numfmts = _parse_cell_number_formats(zf)

        _log(f"Fill styles:   {len(xf_fills)} xf indices", verbose)
        _log(f"Border styles: {len(xf_borders)} xf indices", verbose)
        _log(f"Text styles:   {len(xf_text_styles)} xf indices", verbose)

        diagram_fragments: List[str] = []

        for name, sheet_xml_path in _find_sheet_targets(zf, sheet_names):
            _log(f"Processing sheet: {name}", verbose)
            fragment = _convert_sheet(
                zf, name, sheet_xml_path,
                shared_strings, xf_fills, xf_borders,
                xf_text_styles, xf_numfmts,
                include_cells, verbose,
            )
            diagram_fragments.append(fragment)
            sheets_data[name] = {"title": name}

    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<mxfile host="excel-to-drawio" version="24.7.5" type="device">\n'
        + "".join(diagram_fragments)
        + "</mxfile>\n"
    )

    output_file.parent.mkdir(parents=True, exist_ok=True)
    output_file.write_text(xml, encoding="utf-8")

    return ConversionResult(
        input_path=input_file,
        output_path=output_file,
        sheet_names=list(sheets_data.keys()),
        sheets_data=sheets_data,
    )
