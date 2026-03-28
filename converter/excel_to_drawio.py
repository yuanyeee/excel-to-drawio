"""
High-level Excel -> draw.io conversion entrypoint.

Default behavior uses the maintained ExcelReader/DrawioWriter pipeline.
For backward-compatibility and reference, the previous self-contained OOXML
parser is also kept in this file as a legacy conversion path.
"""


from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Literal
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.dom import minidom

from .drawio_writer import DrawioWriter
from .excel_reader import ExcelReader


@dataclass
class ConversionResult:
    """Summary of a conversion run."""

    input_path: Path
    output_path: Path
    sheet_names: List[str]
    sheets_data: Dict


# === Legacy self-contained parser (kept intentionally) =======================
SS_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

EMU_PER_PX = 9525  # 914400 / 96
DEFAULT_COL_EMU = 922880
DEFAULT_ROW_EMU = 255780

LEGACY_SKIP_FILL_COLORS = {
    "FFFFFF", "FFFFFE", "F2F2F2", "F3F3F3", "EBEBEB", "E7E6E6", "EEECE1",
    "D9D9D9", "BFBFBF", "000000", "0D0D0D",
}

import zipfile, html, sys, re
import xml.etree.ElementTree as ET
from collections import defaultdict
from math import ceil

@dataclass
class _LegacyDrawioShape:
    x: float
    y: float
    width: float
    height: float
    text: str
    style: Dict[str, str]
    shape_type: str = "rectangle"

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

@dataclass
class _LegacyDrawioConnector:
    points: List[Tuple[float, float]]
    style: Dict[str, str]


def _legacy_local(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag

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

def _legacy_normalize_hex_color(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    v = str(value).strip().upper()
    if v.startswith("#"):
        v = v[1:]
    if len(v) == 8:  # AARRGGBB -> RRGGBB
        v = v[2:]
    if len(v) != 6 or not re.fullmatch(r"[0-9A-F]{6}", v):
        return None
    return v


def _legacy_extract_solid_color(fill_or_line_elem: ET.Element) -> Optional[str]:
    srgb = fill_or_line_elem.find(f"{{{A_NS}}}srgbClr")
    if srgb is None:
        return None
    return _legacy_normalize_hex_color(srgb.get("val"))

# ══════════════════════════════════════════════════════════════════════════════
#  設定ブロック（ここを変更して調整）
# ══════════════════════════════════════════════════════════════════════════════
TARGET_SHEET = 'ＮＥＴフロー(夜)_20250508'
INPUT_FILE   = 'ネットワークフロー図（H6）.xlsm'
OUTPUT_FILE  = 'NET_Flow_Night_20250508.drawio'

SCALE             = 1.0   # 全体スケール（1.0 = 実寸）
CHAR_WIDTH        = 7     # 1文字あたりのピクセル幅（標準幅）
POINT_TO_PX       = 96 / 72
EMU_PER_PX        = 9525

# テキスト配置の微調整
CELL_BOX_LEFT_PAD   = 2   # 塗りつぶし矩形の左オフセット（ピクセル）
FILLED_TEXT_TOP_PAD = 2   # 塗りつぶしセルテキストの上オフセット

# スキャン範囲（Excelの行/列インデックス、0始まり）
MIN_ROW, MAX_ROW = 0, 270
MIN_COL, MAX_COL = 0, 230

# ══════════════════════════════════════════════════════════════════════════════
#  Namespaces
# ══════════════════════════════════════════════════════════════════════════════
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
SS  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

# ══════════════════════════════════════════════════════════════════════════════
#  カラーテーブル
# ══════════════════════════════════════════════════════════════════════════════
# Drawing Office テーマカラー
SCHEME_COLORS = {
    'dk1':'000000','lt1':'FFFFFF','dk2':'44546A','lt2':'E7E6E6',
    'acc1':'4472C4','acc2':'ED7D31','acc3':'A9D18E','acc4':'FFC000',
    'acc5':'5B9BD5','acc6':'70AD47','hlink':'0563C1','folHlink':'954F72',
    'bg1':'FFFFFF','bg2':'E7E6E6','tx1':'000000','tx2':'44546A','phClr':'FFFFFF',
}

def _legacy_extract_sp_style(sp_pr: Optional[ET.Element]) -> Dict[str, str]:
    style: Dict[str, str] = {}
    if sp_pr is None:
        return style

    solid_fill = sp_pr.find(f"{{{A_NS}}}solidFill")
    if solid_fill is not None:
        fill = _legacy_extract_solid_color(solid_fill)
        if fill and fill not in LEGACY_SKIP_FILL_COLORS:
            style["fillColor"] = f"#{fill}"

    ln = sp_pr.find(f"{{{A_NS}}}ln")
    if ln is not None:
        line_fill = ln.find(f"{{{A_NS}}}solidFill")
        if line_fill is not None:
            stroke = _legacy_extract_solid_color(line_fill)
            if stroke:
                style["strokeColor"] = f"#{stroke}"

        w = ln.get("w")
        if w:
            try:
                style["strokeWidth"] = str(round(int(w) / 12700, 2))
            except ValueError:
                pass

# 先導/後続コネクター形状セット
OFFPAGE_CONNECTOR_PRSTS = {'flowChartOffpageConnector', 'homePlate', 'pentagon'}

# 先導/後続ラベル検出パターン（"2", "D1", "DA", "ZZ" 等）
# 小さな描画シェイプのテキストがこのパターンにマッチ → Off Page Connector形状に統一
OFFPAGE_LABEL_RE = re.compile(r'[A-Z]{1,2}\d?|\d{1,2}')


def _legacy_extract_text(sp_elem: ET.Element) -> str:
    tx_body = sp_elem.find(f"{{{XDR_NS}}}txBody")
    if tx_body is None:
        return ""
    return "".join((t.text or "") for t in tx_body.iter(f"{{{A_NS}}}t")).strip()

def emu_px(emu):
    return emu / EMU_PER_PX / SCALE

def _legacy_parse_anchor_origin(anchor: ET.Element) -> Tuple[float, float]:
    from_elem = anchor.find(f"{{{XDR_NS}}}from")
    if from_elem is None:
        return 0.0, 0.0




def _legacy_parse_xfrm(sp_pr: Optional[ET.Element], anchor_x: float, anchor_y: float) -> Tuple[float, float, float, float]:
    if sp_pr is None:
        return anchor_x, anchor_y, 100 * EMU_PER_PX, 40 * EMU_PER_PX

def col_letter_to_idx(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1

def cell_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref)
    if not m:
        raise ValueError(f'Invalid cell ref: {ref}')
    return col_letter_to_idx(m.group(1)), int(m.group(2)) - 1

def is_filler(text):
    # ハイフン（-／－）・アスタリスク（*／＊）ともに既存Excelの区切り表現として出力する
    # 空白のみのセルだけをスキップ
    return len(text.strip()) == 0

def normalize_font_name(name):
    if not name:
        return None
    return FONT_ALIASES.get(name, name)

def apply_tint(hex6, tint):
    """
    DrawML の tint 属性（-1.0〜1.0）を近似適用する。
    tint > 0: 白に近づける  /  tint < 0: 黒に近づける
    """
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
        return f'{r:02X}{g:02X}{b:02X}'
    except Exception:
        return hex6


# ══════════════════════════════════════════════════════════════════════════════
#  シートのグリッド（列幅・行高 → ピクセル座標）
# ══════════════════════════════════════════════════════════════════════════════

def build_grid(sh_root):
    col_w = defaultdict(lambda: 8.0)
    for col_el in sh_root.findall('.//x:col', {'x': SS}):
        mn = int(col_el.attrib.get('min', 1))
        mx = int(col_el.attrib.get('max', 1))
        w  = float(col_el.attrib.get('width', 8))
        for c in range(mn - 1, mx):
            col_w[c] = w

    row_h = defaultdict(lambda: 15.0)
    for row_el in sh_root.findall('.//x:row', {'x': SS}):
        r  = int(row_el.attrib.get('r', 1))
        ht = row_el.attrib.get('ht')
        if ht:
            row_h[r - 1] = float(ht)

    MAX = 300
    col_x = [0] * (MAX + 1)
    for i in range(MAX):
        col_x[i + 1] = col_x[i] + chars_px(col_w[i])

    row_y = [0] * (MAX + 1)
    for i in range(MAX):
        row_y[i + 1] = row_y[i] + pts_px(row_h[i])

    return col_x, row_y, col_w, row_h


# ══════════════════════════════════════════════════════════════════════════════
#  DrawIO XML ビルダー
# ══════════════════════════════════════════════════════════════════════════════

class DrawioBuilder:
    def __init__(self):
        self._cells   = []
        self._next    = 2
        self._seen    = set()
        self._max_x   = 0
        self._max_y   = 0

    def add(self, text, x, y, w, h, style, force=False):
        x, y = round(x), round(y)
        w, h = round(max(w, 1)), round(max(h, 1))
        key  = (x, y, w, h, style[:60])
        if key in self._seen and not force:
            return
        self._seen.add(key)
        self._max_x = max(self._max_x, x + w)
        self._max_y = max(self._max_y, y + h)
        cid = self._next
        self._next += 1
        esc = html.escape(str(text))
        # ★ v6 重要修正: width / height（DrawIOで正しく描画される）
        self._cells.append(
            f'    <mxCell id="{cid}" value="{esc}" style="{style}" vertex="1" parent="1">'
            f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/>'
            f'</mxCell>'
        )

def _legacy_parse_shape(sp_elem: ET.Element, anchor_x: float, anchor_y: float) -> Optional[_LegacyDrawioShape]:
    sp_pr = sp_elem.find(f"{{{XDR_NS}}}spPr")
    x, y, w, h = _legacy_parse_xfrm(sp_pr, anchor_x, anchor_y)
    style = _legacy_extract_sp_style(sp_pr)
    text = _legacy_extract_text(sp_elem)

    has_fill = bool(style.get("fillColor"))
    has_stroke = bool(style.get("strokeColor"))
    if not text and not has_fill and not has_stroke:
        return None
    if not text and min(_to_px(w), _to_px(h)) < 2:
        return None
    sf = el.find(f'{{{A}}}solidFill') or el
    s  = sf.find(f'{{{A}}}srgbClr')
    if s is not None:
        return '#' + s.attrib.get('val', '000000').upper()
    sc = sf.find(f'{{{A}}}schemeClr')
    if sc is not None:
        base = SCHEME_COLORS.get(sc.attrib.get('val', 'dk1'), '808080')
        lum_mod = sc.find(f'{{{A}}}lumMod')
        lum_off = sc.find(f'{{{A}}}lumOff')
        if lum_mod is not None or lum_off is not None:
            mod = int(lum_mod.attrib.get('val', '100000')) / 100000 if lum_mod is not None else 1.0
            off = int(lum_off.attrib.get('val', '0'))     / 100000 if lum_off is not None else 0.0
            base = apply_tint(base, (mod - 1 + off))
        return '#' + base.upper()
    sy = sf.find(f'{{{A}}}sysClr')
    if sy is not None:
        last = sy.attrib.get('lastClr')
        if last:
            return '#' + last.upper()
    return None

    return _LegacyDrawioShape(x=x, y=y, width=w, height=h, text=text, style=style)

def sp_fill(sp_pr):
    if sp_pr.find(f'{{{A}}}noFill') is not None:
        return 'none'
    for fill_tag in (f'{{{A}}}solidFill', f'{{{A}}}gradFill', f'{{{A}}}pattFill'):
        fe = sp_pr.find(fill_tag)
        if fe is not None:
            if fill_tag.endswith('solidFill'):
                c = parse_color(fe)
            elif fill_tag.endswith('gradFill'):
                gs = fe.find(f'.//{{{A}}}gs')
                c  = parse_color(gs) if gs is not None else None
            else:
                bg = fe.find(f'{{{A}}}bgClr')
                c  = parse_color(bg) if bg is not None else None
            if c:
                return c
    return '#FFFFFF'


def sp_line(sp_pr):
    ln = sp_pr.find(f'{{{A}}}ln')
    if ln is None:
        return '#000000', 1
    if ln.find(f'{{{A}}}noFill') is not None:
        return 'none', 0
    sf = ln.find(f'{{{A}}}solidFill')
    color = parse_color(sf) if sf is not None else '#000000'
    if color is None:
        color = '#000000'
    w_emu = int(ln.attrib.get('w', '12700'))
    return color, max(1, round(w_emu / 12700))


def sp_geom(sp_pr):
    g = sp_pr.find(f'{{{A}}}prstGeom')
    return g.attrib.get('prst', 'rect') if g is not None else 'rect'


def sp_fontsize(txb):
    if txb is None:
        return 9
    for tag in (f'{{{A}}}rPr', f'{{{A}}}endParaRPr'):
        e = txb.find(f'.//{tag}')
        if e is not None:
            sz = e.attrib.get('sz')
            if sz:
                return max(7, round(int(sz) / 100))
    return 9


def sp_font_style(txb):
    """DrawML テキスト要素から fontColor, fontStyle（bold/italic）を取得"""
    if txb is None:
        return {}, None
    rpr = txb.find(f'.//{{{A}}}rPr') or txb.find(f'.//{{{A}}}endParaRPr')
    if rpr is None:
        return {}, None
    extra = {}
    solid = rpr.find(f'{{{A}}}solidFill')
    if solid is not None:
        fc = parse_color(solid)
        if fc and fc not in ('#000000', '#FFFFFF'):
            extra['fontColor'] = fc
    fs = 0
    if rpr.attrib.get('b') == '1':  fs |= 1
    if rpr.attrib.get('i') == '1':  fs |= 2
    if fs:
        extra['fontStyle'] = fs
    return extra, None


def make_style(prst, fill, lc, lw, fsz, font_extra=None):
    parts = ['whiteSpace=wrap', 'html=1']
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
    return ';'.join(parts) + ';'


# ══════════════════════════════════════════════════════════════════════════════
#  グループ図形の座標変換
# ══════════════════════════════════════════════════════════════════════════════

def get_xfrm(xfrm):
    def iv(el, attr, default=0):
        return int(el.attrib.get(attr, default)) if el is not None else default
    off   = xfrm.find(f'{{{A}}}off')
    ext   = xfrm.find(f'{{{A}}}ext')
    choff = xfrm.find(f'{{{A}}}chOff')
    chext = xfrm.find(f'{{{A}}}chExt')
    ox, oy     = iv(off,   'x'), iv(off,   'y')
    ecx, ecy   = iv(ext,   'cx'), iv(ext,   'cy')
    chox, choy = iv(choff, 'x', ox), iv(choff, 'y', oy)
    chcx, chcy = iv(chext, 'cx', ecx), iv(chext, 'cy', ecy)
    return ox, oy, ecx, ecy, chox, choy, chcx, chcy


def get_text(el):
    return ''.join(t.text for t in el.iter(f'{{{A}}}t') if t.text)


def emit_sp(sp, pax, pay, sx, sy, bld):
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
    w  = int(ext.attrib.get('cx', 0)) * sx
    h  = int(ext.attrib.get('cy', 0)) * sy

    if w < 1 or h < 1:
        return

    text   = get_text(sp)
    fill   = sp_fill(spr)
    lc, lw = sp_line(spr)
    prst   = sp_geom(spr)
    txb    = sp.find(f'{{{XDR}}}txBody')
    fsz    = sp_fontsize(txb)
    fe, _  = sp_font_style(txb)

    if not text and fill in ('#FFFFFF', 'none') and lc == 'none':
        return

    # 先導/後続ラベル検出: 小さなシェイプの短いテキスト → Off Page Connector形状に統一
    text_s = text.strip()
    if text_s and w < 80 and h < 80 and OFFPAGE_LABEL_RE.fullmatch(text_s):
        prst = 'homePlate'
        if fill in ('none',):
            fill = '#FFFFFF'
        if lc == 'none':
            lc, lw = '#000000', 1

    style = make_style(prst, fill, lc, lw, fsz, fe)
    bld.add(text, ax, ay, w, h, style, force=bool(text))


def emit_cxnsp(cxn, pax, pay, sx, sy, bld):
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

def _legacy_parse_connector(cxn_elem: ET.Element, anchor_x: float, anchor_y: float) -> Optional[_LegacyDrawioConnector]:
    sp_pr = cxn_elem.find(f"{{{XDR_NS}}}spPr")
    x, y, w, h = _legacy_parse_xfrm(sp_pr, anchor_x, anchor_y)
    p1 = (x, y)
    p2 = (x + w, y + h)
    style = _legacy_extract_sp_style(sp_pr)

    w = raw_w if raw_w >= 1 else 2
    h = raw_h if raw_h >= 1 else 2

    return _LegacyDrawioConnector(points=[p1, p2], style=style)

    if ln is not None:
        sf    = ln.find(f'{{{A}}}solidFill')
        color = parse_color(sf) if sf is not None else '#000000'
        if color is None:
            color = '#000000'
    else:
        color = '#000000'

    lw_emu = int(ln.attrib.get('w', '12700')) if ln is not None else 12700
    lw_px  = max(1, round(lw_emu / 12700))

    if raw_w < 1 or raw_h < 1:
        style = (f'whiteSpace=wrap;html=1;fillColor={color};'
                 f'strokeColor={color};strokeWidth={lw_px};')
    else:
        style = (f'whiteSpace=wrap;html=1;fillColor=none;'
                 f'strokeColor={color};strokeWidth={lw_px};')

    bld.add('', ax, ay, w, h, style)


def walk_group(grp, pax, pay, sx, sy, bld, depth=0):
    if depth > 25:
        return
    grp_pr = grp.find(f'{{{XDR}}}grpSpPr')
    if grp_pr is None:
        return
    xfrm = grp_pr.find(f'{{{A}}}xfrm')
    if xfrm is None:
        return

    ox, oy, ecx, ecy, chox, choy, chcx, chcy = get_xfrm(xfrm)
    gax, gay = pax + ox * sx, pay + oy * sy
    gw, gh   = ecx * sx, ecy * sy

    csx = (gw / chcx) if chcx else sx
    csy = (gh / chcy) if chcy else sy
    cox = gax - chox * csx
    coy = gay - choy * csy

    for child in grp:
        ct = child.tag.split('}')[-1]
        if   ct == 'sp':    emit_sp(child,    cox, coy, csx, csy, bld)
        elif ct == 'cxnSp': emit_cxnsp(child, cox, coy, csx, csy, bld)
        elif ct == 'grpSp': walk_group(child, cox, coy, csx, csy, bld, depth + 1)
        # pic はスキップ


def anchor_rect(anchor, col_x, row_y):
    """
    アンカーのセル参照からピクセル矩形 (x, y, w, h) を返す。
    findtext() を使用して XML 要素の真偽値問題を回避する（v6方式）。
    """
    from_el = anchor.find(f'{{{XDR}}}from')
    if from_el is None:
        return None

def _legacy_parse_drawing_xml(content: bytes) -> Tuple[List[_LegacyDrawioShape], List[_LegacyDrawioConnector]]:
    root = ET.fromstring(content)
    shapes: List[_LegacyDrawioShape] = []
    connectors: List[_LegacyDrawioConnector] = []

    for anchor in root:
        if _legacy_local(anchor.tag) not in ("oneCellAnchor", "twoCellAnchor"):
            continue
        anchor_x, anchor_y = _legacy_parse_anchor_origin(anchor)
        for child in anchor:
            name = _legacy_local(child.tag)
            if name == "sp":
                shape = _legacy_parse_shape(child, anchor_x, anchor_y)
                if shape:
                    shapes.append(shape)
            elif name == "cxnSp":
                conn = _legacy_parse_connector(child, anchor_x, anchor_y)
                if conn:
                    connectors.append(conn)

    return shapes, connectors


def _legacy_resolve_xl_target(base_path: str, target: str) -> str:
    base_dir = Path(base_path).parent
    joined = (base_dir / target).as_posix()
    parts: List[str] = []
    for p in joined.split("/"):
        if p in ("", "."):
            continue
        try:
            c1, r1, c2, r2 = parse_range_ref(ref)
        except Exception:
            continue
        parts.append(p)
    return "/".join(parts)


def _legacy_find_sheet_targets(zf: zipfile.ZipFile, sheet_names: Optional[List[str]]) -> List[Tuple[str, Optional[str]]]:
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
            targets.append((name, _legacy_resolve_xl_target("xl/workbook.xml", target)))
    return targets


def _legacy_find_drawing_for_sheet(zf: zipfile.ZipFile, sheet_xml_path: str) -> Optional[str]:
    sheet_file = Path(sheet_xml_path).name
    rels_path = f"xl/worksheets/_rels/{sheet_file}.rels"
    if rels_path not in zf.namelist():
        return None

    rels_root = ET.fromstring(zf.read(rels_path))
    for rel in rels_root.findall(f".//{{{PKG_REL_NS}}}Relationship"):
        rel_type = rel.get("Type", "")
        if "drawing" in rel_type and "vml" not in rel_type:
            target = rel.get("Target", "")
            return _legacy_resolve_xl_target(sheet_xml_path, target)

    return None


def _legacy_build_style(style: Dict[str, str], shape_type: str = "rectangle") -> str:
    pairs = {"shape": shape_type, "whiteSpace": "wrap", "html": "1"}
    pairs.update(style or {})
    return ";".join(f"{k}={v}" for k, v in pairs.items()) + ";"


def _legacy_build_drawio_xml(sheets_data: Dict[str, Dict[str, List]]) -> str:
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
            cell.set("style", _legacy_build_style(shape.style, shape.shape_type))
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


def convert_excel_to_drawio_legacy(
    input_path: str,
    output_path: str,
    sheet_names: Optional[List[str]] = None,
    include_cells: bool = True,
) -> ConversionResult:
    """Legacy direct-OOXML converter kept for compatibility/reference."""
    input_file = Path(input_path)
    output_file = Path(output_path)

    sheets_data: Dict[str, Dict[str, List]] = {}
    with zipfile.ZipFile(input_file, "r") as zf:
        for name, sheet_xml_path in _legacy_find_sheet_targets(zf, sheet_names):
            drawing_path = _legacy_find_drawing_for_sheet(zf, sheet_xml_path)
            if drawing_path and drawing_path in zf.namelist():
                shapes, connectors = _legacy_parse_drawing_xml(zf.read(drawing_path))
            else:
                shapes, connectors = [], []
            sheets_data[name] = {
                "shapes": shapes,
                "connectors": connectors,
                "title": name,
            }

    xml = _legacy_build_drawio_xml(sheets_data)
    output_file.write_text(xml, encoding="utf-8")

    return ConversionResult(
        input_path=input_file,
        output_path=output_file,
        sheet_names=list(sheets_data.keys()),
        sheets_data=sheets_data,
    )


# === Preferred maintained path =================================================
def convert_excel_to_drawio(
    input_path: str,
    output_path: str,
    sheet_names: Optional[List[str]] = None,
    include_cells: bool = True,
    engine: Literal["pipeline", "legacy"] = "pipeline",
) -> ConversionResult:
    """
    Convert an Excel workbook to a draw.io file.

    Args:
        input_path: Path to source Excel file.
        output_path: Path to output .drawio file.
        sheet_names: Optional list of target sheet names. If omitted, all sheets.
        include_cells: Whether to include cell-based objects (fills/borders/labels).
        engine: "pipeline" (default) or "legacy".

    Returns:
        ConversionResult with selected sheet names and extracted data.
    """
    if engine == "legacy":
        return convert_excel_to_drawio_legacy(
            input_path=input_path,
            output_path=output_path,
            sheet_names=sheet_names,
            include_cells=include_cells,
        )

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
