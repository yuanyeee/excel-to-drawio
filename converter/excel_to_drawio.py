"""
High-level Excel -> draw.io conversion entrypoint.

This module orchestrates the existing reader/writer pipeline so that:
- input Excel file and target sheet names are supplied via parameters,
- cell-derived objects are included when requested,
- all selected sheets are exported into a single multi-page draw.io file.
"""


from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

from .drawio_writer import DrawioWriter
from .excel_reader import ExcelReader

from __future__ import annotations

import zipfile, html, sys, re
import xml.etree.ElementTree as ET
from collections import defaultdict
from math import ceil


from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

#from .drawio_writer import DrawioWriter
#from .excel_reader import ExcelReader


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

    
    input_file = Path(input_path)
    output_file = Path(output_path)
    
    """
    reader = ExcelReader(
        filepath=str(input_file),
        sheet_names=sheet_names,
        include_cells=include_cells,
    )
    sheets_data = reader.read_all()

    writer = DrawioWriter(sheets_data)
    writer.write(str(output_file))

    
    
    writer = DrawioWriter(sheets_data)
    writer.write(str(output_file))
    """
    for sheet_name in sheet_names:
        convert(input_file,sheet_names,output_file)
    
    return ConversionResult(
        input_path=input_file,
        output_path=output_file,
        sheet_names=list(sheets_data.keys()),
        sheets_data=sheets_data,
    )



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

# Office 2013 テーマカラー（cellXfs の fgColor/@theme 用）
THEME_FILL_COLORS = [
    'FFFFFF','000000','EEECE1','1F497D',
    '4BACC6','4472C4','9BBB59','F79646',
    'FFFF00','A9D18E','5B9BD5','70AD47',
]

# Office indexed color palette
INDEXED_COLORS = [
    '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
    '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
    '800000','008000','000080','808000','800080','008080','C0C0C0','808080',
    '9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF',
    '000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF',
    '00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99',
    '3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696',
    '003366','339966','003300','333300','993300','993366','333399','333333',
    'FFFFFF','FFFFFF',  # 64, 65 = system colors
]

# 描画不要な背景色（白系・黒系）
SKIP_COLORS = {
    'FFFFFF','FFFFFE','F2F2F2','F3F3F3','EBEBEB','E7E6E6','EEECE1',
    'D9D9D9','BFBFBF','000000','0D0D0D'
}

# 形状 → DrawIO スタイルマッピング
GEOM_STYLES = {
    'roundRect':                 'rounded=1;arcSize=10;',
    'ellipse':                   'ellipse;',
    'diamond':                   'rhombus;',
    'triangle':                  'triangle;',
    'parallelogram':             'parallelogram;',
    'trapezoid':                 'trapezoid;',
    'hexagon':                   'hexagon;',
    'octagon':                   'octagon;',
    # ── フローチャート系 ──────────────────────────────────────────────────────
    'flowChartOffpageConnector': 'shape=offPageConnector;',
    'flowChartProcess':          'shape=mxgraph.flowchart.process;',
    'flowChartDecision':         'shape=mxgraph.flowchart.decision;',
    'flowChartTerminator':       'shape=mxgraph.flowchart.terminator;',
    'flowChartManualInput':      'shape=mxgraph.flowchart.manual_input;',
    'flowChartDocument':         'shape=mxgraph.flowchart.document;',
    'flowChartPredefinedProcess':'shape=mxgraph.flowchart.predefined_process;',
    'flowChartConnector':        'ellipse;',
    'flowChartPunchedTape':      'shape=mxgraph.flowchart.punched_tape;',
    'flowChartSort':             'shape=mxgraph.flowchart.sort;',
    # ── 五角形（先導/後続ボックス）────────────────────────────────────────────
    # homePlate = Excel の「ホームプレート」= 五角形（下向き）
    # DrawIO の Off Page Connector と同じ形状
    'homePlate':                 'shape=offPageConnector;',
    'pentagon':                  'shape=offPageConnector;',
    # ── 吹き出し系 ───────────────────────────────────────────────────────────
    'wedgeRoundRectCallout':     'shape=callout;rounded=1;',
    'wedgeRectCallout':          'shape=callout;',
    'cloudCallout':              'shape=callout;rounded=1;',
    # ── 矢印系 ───────────────────────────────────────────────────────────────
    'bentArrow':                 'shape=mxgraph.arrows2.bent_arrow;',
    'chevron':                   'shape=mxgraph.arrows2.arrow;dy=0.6;dx=20;notch=0;',
    'rightArrow':                'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=east;',
    'leftArrow':                 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=west;',
    'upArrow':                   'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=north;',
    'downArrow':                 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=south;',
}

# フォント名の正規化マッピング
FONT_ALIASES = {
    'ＭＳ ゴシック':      'MS PGothic',
    'ＭＳ Ｐゴシック':    'MS PGothic',
    'MS Gothic':           'MS PGothic',
    'MS PGothic':          'MS PGothic',
    'ＭＳ 明朝':           'MS PMincho',
    'ＭＳ Ｐ明朝':         'MS PMincho',
    '游ゴシック':          'Yu Gothic',
    '游ゴシック Light':   'Yu Gothic Light',
    '游明朝':              'Yu Mincho',
    'メイリオ':            'Meiryo',
    'Meiryo':              'Meiryo',
}

# 先導/後続コネクター形状セット
OFFPAGE_CONNECTOR_PRSTS = {'flowChartOffpageConnector', 'homePlate', 'pentagon'}

# 先導/後続ラベル検出パターン（"2", "D1", "DA", "ZZ" 等）
# 小さな描画シェイプのテキストがこのパターンにマッチ → Off Page Connector形状に統一
OFFPAGE_LABEL_RE = re.compile(r'[A-Z]{1,2}\d?|\d{1,2}')


# ══════════════════════════════════════════════════════════════════════════════
#  ユーティリティ
# ══════════════════════════════════════════════════════════════════════════════

def emu_px(emu):
    return emu / EMU_PER_PX / SCALE

def chars_px(c):
    """Excel 文字幅単位 → ピクセル（過剰パディングなし・正確な計算）"""
    return max(1, int(c * CHAR_WIDTH + 0.5))




def pts_px(pts):
    return round(pts * POINT_TO_PX)

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

    def xml(self):
        page_w = max(2000, int(self._max_x * 1.10))
        page_h = max(2000, int(self._max_y * 1.10))
        hdr = (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<mxfile host="Claude" version="24.7.5" type="device">\n'
            f'  <diagram id="d1" name="{TARGET_SHEET}">\n'
            '    <mxGraphModel grid="0" guides="1" tooltips="1" connect="1" arrows="1"\n'
            f'                  fold="1" page="1" pageScale="1" pageWidth="{page_w}"\n'
            f'                  pageHeight="{page_h}" math="0" shadow="0">\n'
            '      <root>\n'
            '        <mxCell id="0"/>\n'
            '        <mxCell id="1" parent="0"/>\n'
        )
        ftr = '      </root>\n    </mxGraphModel>\n  </diagram>\n</mxfile>\n'
        return hdr + '\n'.join(self._cells) + '\n' + ftr


# ══════════════════════════════════════════════════════════════════════════════
#  styles.xml パーサー
# ══════════════════════════════════════════════════════════════════════════════

def _parse_color_el(color_el, default='#000000'):
    """fgColor/bgColor/color 要素から '#RRGGBB' を返す（tint 補正付き）"""
    if color_el is None:
        return default

    rgb = color_el.attrib.get('rgb', '')
    if rgb:
        h6 = (rgb[2:] if len(rgb) == 8 else rgb[:6]).upper()
        tint = color_el.attrib.get('tint', '')
        if tint:
            h6 = apply_tint(h6, tint)
        return '#' + h6

    theme = color_el.attrib.get('theme', '')
    if theme:
        idx = int(theme)
        base = THEME_FILL_COLORS[idx] if idx < len(THEME_FILL_COLORS) else None
        if base:
            tint = color_el.attrib.get('tint', '')
            if tint:
                base = apply_tint(base, tint)
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


def parse_cell_styles(z):
    """xf_index → 塗りつぶし色 '#RRGGBB' のマップ"""
    xf_fills = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except Exception as e:
        sys.stderr.write(f'styles.xml parse error: {e}\n')
        return xf_fills

    NS = {'x': SS}
    fills = []
    for fill_el in root.findall('.//x:fills/x:fill', NS):
        color = None
        pf = fill_el.find(f'{{{SS}}}patternFill')
        if pf is not None and pf.attrib.get('patternType', 'none') != 'none':
            fg = pf.find(f'{{{SS}}}fgColor')
            if fg is not None:
                c = _parse_color_el(fg, default=None)
                if c and not _should_skip_fill(c):
                    color = c
        fills.append(color)

    for i, xf in enumerate(root.findall('.//x:cellXfs/x:xf', NS)):
        fill_id = int(xf.attrib.get('fillId', '0'))
        if fill_id < len(fills) and fills[fill_id]:
            xf_fills[i] = fills[fill_id]

    return xf_fills


def parse_cell_borders(z):
    """xf_index → 罫線情報 {side: (color, width_px)} のマップ"""
    xf_borders = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except Exception as e:
        sys.stderr.write(f'styles.xml parse error: {e}\n')
        return xf_borders

    NS = {'x': SS}

    def _bw(style_name):
        if style_name in ('medium', 'mediumDashed', 'mediumDashDot',
                           'mediumDashDotDot', 'slantDashDot'):
            return 2
        if style_name == 'thick':
            return 3
        return 1

    border_defs = []
    for bel in root.findall('.//x:borders/x:border', NS):
        sides = {}
        for side in ('left', 'right', 'top', 'bottom'):
            sel = bel.find(f'{{{SS}}}{side}')
            if sel is None:
                continue
            sname = sel.attrib.get('style')
            if not sname:
                continue
            color = _parse_color_el(sel.find(f'{{{SS}}}color'))
            sides[side] = (color, _bw(sname))
        border_defs.append(sides)

    for i, xf in enumerate(root.findall('.//x:cellXfs/x:xf', NS)):
        bid = int(xf.attrib.get('borderId', '0'))
        if 0 <= bid < len(border_defs) and border_defs[bid]:
            xf_borders[i] = border_defs[bid]

    return xf_borders


def parse_cell_text_styles(z):
    """xf_index → テキストスタイル辞書 のマップ"""
    xf_text_styles = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except Exception as e:
        sys.stderr.write(f'styles.xml parse error: {e}\n')
        return xf_text_styles

    NS = {'x': SS}
    fonts = []
    for font_el in root.findall('.//x:fonts/x:font', NS):
        name_el  = font_el.find(f'{{{SS}}}name')
        size_el  = font_el.find(f'{{{SS}}}sz')
        color_el = font_el.find(f'{{{SS}}}color')
        bold     = font_el.find(f'{{{SS}}}b') is not None
        italic   = font_el.find(f'{{{SS}}}i') is not None
        fonts.append({
            'fontFamily': normalize_font_name(name_el.attrib.get('val')) if name_el is not None else None,
            'fontSize':   max(6, round(float(size_el.attrib.get('val', '11')))) if size_el is not None else 11,
            'fontColor':  _parse_color_el(color_el, default='#000000'),
            'bold':       bold,
            'italic':     italic,
        })

    for i, xf in enumerate(root.findall('.//x:cellXfs/x:xf', NS)):
        style = {}
        fid = int(xf.attrib.get('fontId', '0'))
        if 0 <= fid < len(fonts):
            f = fonts[fid]
            if f.get('fontFamily'):
                style['fontFamily'] = f['fontFamily']
            if f.get('fontSize'):
                style['fontSize'] = f['fontSize']
            if f.get('fontColor') and f['fontColor'] != '#000000':
                style['fontColor'] = f['fontColor']
            fs = 0
            if f.get('bold'):   fs |= 1
            if f.get('italic'): fs |= 2
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

        xf_text_styles[i] = style

    return xf_text_styles


def parse_cell_number_formats(z):
    """xf_index → (numFmtId, formatCode) のマップ"""
    xf_numfmts = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except Exception as e:
        sys.stderr.write(f'styles.xml parse error: {e}\n')
        return xf_numfmts

    NS = {'x': SS}
    custom = {
        int(el.attrib.get('numFmtId', '0')): el.attrib.get('formatCode', '')
        for el in root.findall('.//x:numFmts/x:numFmt', NS)
    }
    for i, xf in enumerate(root.findall('.//x:cellXfs/x:xf', NS)):
        nid = int(xf.attrib.get('numFmtId', '0'))
        xf_numfmts[i] = (nid, custom.get(nid, ''))

    return xf_numfmts


# ══════════════════════════════════════════════════════════════════════════════
#  セル塗りつぶし: 隣接同色セルを結合して矩形化
# ══════════════════════════════════════════════════════════════════════════════

def add_cell_fills_merged(sh_root, col_x, row_y, col_w_dict, row_h_dict, xf_fills, bld):
    """
    隣接する同色セルをまとめて大きな矩形として描画する（バッチジョブボックス単位）。
    """
    NS_X = {'x': SS}
    color_grid = {}

    for row_el in sh_root.findall('.//x:row', NS_X):
        r = int(row_el.attrib.get('r', 1)) - 1
        if r < MIN_ROW or r > MAX_ROW:
            continue
        for cell in row_el.findall('x:c', NS_X):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = cell_ref(ref)
            except Exception:
                continue
            if c < MIN_COL or c > MAX_COL:
                continue
            s_attr = int(cell.attrib.get('s', 0))
            fc = xf_fills.get(s_attr)
            if fc:
                color_grid[(r, c)] = fc

    # 結合セルの色を子セル全体に伝播する
    # Excelでは結合セルの塗りは左上セルにのみ定義され、子セルはXMLに出現しない
    merged_topleft, _ = build_merged_cell_maps(sh_root)
    for (r1, c1), (r2, c2) in merged_topleft.items():
        color = color_grid.get((r1, c1))
        if color:
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    if (rr, cc) not in color_grid:
                        color_grid[(rr, cc)] = color

    _log(f"  Color grid cells (after merge propagation): {len(color_grid)}")
    if not color_grid:
        return 0

    processed = set()
    count = 0

    for (r, c) in sorted(color_grid.keys()):
        if (r, c) in processed:
            continue
        color = color_grid[(r, c)]

        # 右方向の連続範囲
        c_end = c
        while color_grid.get((r, c_end + 1)) == color and (r, c_end + 1) not in processed:
            c_end += 1

        # 下方向の連続範囲（同じ列スパンが全て同色・未処理）
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

        px     = max(0.0, col_x[min(c, 300)] / SCALE - CELL_BOX_LEFT_PAD)
        py     = row_y[min(r, 300)] / SCALE
        px_end = col_x[min(c_end + 1, 300)] / SCALE
        py_end = row_y[min(r_end + 1, 300)] / SCALE
        w      = max(2.0, px_end - px)
        h      = max(2.0, py_end - py)

        style = f'whiteSpace=wrap;html=1;fillColor={color};strokeColor=none;'
        bld.add('', px, py, w, h, style)
        count += 1

    return count


# ══════════════════════════════════════════════════════════════════════════════
#  セル罫線描画
# ══════════════════════════════════════════════════════════════════════════════

def add_cell_borders(sh_root, col_x, row_y, col_w_dict, row_h_dict, xf_borders, xf_fills, bld):
    ns = {'x': SS}
    count = 0

    # ── 事前スキャン: 行ごとにコンテンツ（値または塗り）が存在する列範囲を計算
    # 値なし+塗りなしのセルでも、同行のコンテンツ列範囲内（右端+マージン）なら罫線を描画する
    # （ボックスの右端セルは空でも右罫線が定義されているため、スキップしてはいけない）
    BORDER_MARGIN = 10  # コンテンツ右端から右に許容する追加列数
    row_active = {}  # row → (min_col, max_col + BORDER_MARGIN)
    filled_positions = set()  # 塗りありセルの位置集合（縦罫線の内側/外側判定に使用）
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        cols = []
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = cell_ref(ref)
            except Exception:
                continue
            s_attr = int(cell.attrib.get('s', 0))
            v_el = cell.find('x:v', ns)
            if (v_el is not None and v_el.text is not None) or xf_fills.get(s_attr):
                cols.append(c)
            if xf_fills.get(s_attr):
                filled_positions.add((r, c))
        if cols:
            row_active[r] = (min(cols), max(cols) + BORDER_MARGIN)

    # ── 罫線描画
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        if r < MIN_ROW or r > MAX_ROW:
            continue
        cy = row_y[min(r, 299)] / SCALE
        ch = max(1.0, pts_px(row_h_dict[r]) / SCALE)

        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = cell_ref(ref)
            except Exception:
                continue
            if c < MIN_COL or c > MAX_COL:
                continue

            s_attr = int(cell.attrib.get('s', 0))
            border_info = xf_borders.get(s_attr)
            if not border_info:
                continue

            # v8改: 値なし＋塗りなしのセルは、同行のコンテンツ列範囲外のみスキップ
            # ボックス右端の空セル（右罫線が定義されている）を正しく描画するため
            v_el      = cell.find('x:v', ns)
            has_value = (v_el is not None and v_el.text is not None)
            has_fill  = xf_fills.get(s_attr) is not None
            if not has_value and not has_fill:
                active = row_active.get(r)
                if active is None or c < active[0] or c > active[1]:
                    continue

            cx = col_x[min(c, 299)] / SCALE
            cw = max(1.0, chars_px(col_w_dict[c]) / SCALE)
            bx = max(0.0, cx - CELL_BOX_LEFT_PAD)
            bw = cw + min(CELL_BOX_LEFT_PAD, cx)

            # 塗りありセルの縦罫線：隙接セルも塗りありなら「内側」→スキップ、塗りなしセルが隣なら「外側（ボックス左右端）」→残す
            cell_fill_color = xf_fills.get(s_attr)
            for side, (color, width_px) in border_info.items():
                if side == 'left' and cell_fill_color and (r, c - 1) in filled_positions:
                    continue
                if side == 'right' and cell_fill_color and (r, c + 1) in filled_positions:
                    continue
                style = f'whiteSpace=wrap;html=1;fillColor={color};strokeColor={color};'
                if side == 'top':
                    bld.add('', bx, cy, bw, width_px, style)
                elif side == 'bottom':
                    bld.add('', bx, cy + ch - width_px, bw, width_px, style)
                elif side == 'left':
                    bld.add('', bx, cy, width_px, ch, style)
                elif side == 'right':
                    bld.add('', cx + cw - width_px, cy, width_px, ch, style)
                count += 1

    return count


# ══════════════════════════════════════════════════════════════════════════════
#  描画図形（Drawing XML）のスタイルヘルパー
# ══════════════════════════════════════════════════════════════════════════════

def parse_color(el):
    """DrawML のカラー要素から '#RRGGBB' を返す"""
    if el is None:
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

    ax    = pax + int(off.attrib.get('x', 0)) * sx
    ay    = pay + int(off.attrib.get('y', 0)) * sy
    raw_w = int(ext.attrib.get('cx', 0)) * sx
    raw_h = int(ext.attrib.get('cy', 0)) * sy

    w = raw_w if raw_w >= 1 else 2
    h = raw_h if raw_h >= 1 else 2

    ln = spr.find(f'{{{A}}}ln')
    if ln is not None and ln.find(f'{{{A}}}noFill') is not None:
        return

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

    fc  = int(from_el.findtext(f'{{{XDR}}}col',    '0') or '0')
    fco = int(from_el.findtext(f'{{{XDR}}}colOff', '0') or '0')
    fr  = int(from_el.findtext(f'{{{XDR}}}row',    '0') or '0')
    fro = int(from_el.findtext(f'{{{XDR}}}rowOff', '0') or '0')

    anc_x = col_x[min(fc, 299)] / SCALE + emu_px(fco)
    anc_y = row_y[min(fr, 299)] / SCALE + emu_px(fro)

    to_el  = anchor.find(f'{{{XDR}}}to')
    ext_el = anchor.find(f'{{{XDR}}}ext')

    if to_el is not None:
        tc  = int(to_el.findtext(f'{{{XDR}}}col',    '0') or '0')
        tco = int(to_el.findtext(f'{{{XDR}}}colOff', '0') or '0')
        tr  = int(to_el.findtext(f'{{{XDR}}}row',    '0') or '0')
        tro = int(to_el.findtext(f'{{{XDR}}}rowOff', '0') or '0')
        anc_w = max(2.0, col_x[min(tc, 299)] / SCALE + emu_px(tco) - anc_x)
        anc_h = max(2.0, row_y[min(tr, 299)] / SCALE + emu_px(tro) - anc_y)
    elif ext_el is not None:
        anc_w = max(2.0, emu_px(int(ext_el.attrib.get('cx', '9525'))))
        anc_h = max(2.0, emu_px(int(ext_el.attrib.get('cy', '9525'))))
    else:
        anc_w, anc_h = 80.0, 24.0

    return anc_x, anc_y, anc_w, anc_h


def add_drawing_shapes(z, drawing_path, col_x, row_y, bld):
    """
    v9b アンカー基準方式（v6ロジック復元）:
    - sp: アンカーのセル参照位置を使用（xfrm.offは無視）
    - grpSp: アンカー位置をグループ原点として使用、chOff/chExtで内部座標変換
    - cxnSp: EMU絶対座標を使用
    - homePlate/pentagon/flowChartOffpageConnector → DrawIO標準シェイプ
    """
    dr = ET.fromstring(z.read(drawing_path).decode('utf-8'))
    sc = 1.0 / EMU_PER_PX / SCALE  # EMU → ピクセル変換スケール

    for anchor in dr:
        tag = anchor.tag.split('}')[-1]
        if tag not in ('oneCellAnchor', 'twoCellAnchor'):
            continue

        rect = anchor_rect(anchor, col_x, row_y)
        if rect is None:
            continue
        anc_x, anc_y, anc_w, anc_h = rect

        for child in anchor:
            ct = child.tag.split('}')[-1]

            if ct == 'sp':
                # spはアンカー位置・サイズで配置（xfrm.offはExcelが必ずしも同期しない）
                spr = child.find(f'{{{XDR}}}spPr')
                if spr is None:
                    continue
                text   = get_text(child)
                fill   = sp_fill(spr)
                lc, lw = sp_line(spr)
                prst   = sp_geom(spr)
                txb    = child.find(f'{{{XDR}}}txBody')
                fsz    = sp_fontsize(txb)
                fe, _  = sp_font_style(txb)
                if not text and fill in ('#FFFFFF', 'none') and lc == 'none':
                    continue
                # 先導/後続ラベル検出: Off Page Connector形状に統一
                text_s = text.strip()
                if text_s and anc_w < 80 and anc_h < 80 and OFFPAGE_LABEL_RE.fullmatch(text_s):
                    prst = 'homePlate'
                    if fill in ('none',):
                        fill = '#FFFFFF'
                    if lc == 'none':
                        lc, lw = '#000000', 1
                style = make_style(prst, fill, lc, lw, fsz, fe)
                bld.add(text, anc_x, anc_y, anc_w, anc_h, style, force=bool(text))

            elif ct == 'grpSp':
                # grpSpはアンカー位置を原点とし、chOff/chExtで内部座標系を構築
                grp_pr = child.find(f'{{{XDR}}}grpSpPr')
                if grp_pr is None:
                    continue
                xfrm = grp_pr.find(f'{{{A}}}xfrm')
                if xfrm is None:
                    continue
                _, _, ecx, ecy, chox, choy, chcx, chcy = get_xfrm(xfrm)
                # アンカーサイズ / 子座標系サイズ で変換スケールを決定
                csx = (anc_w / chcx) if chcx else sc
                csy = (anc_h / chcy) if chcy else sc
                cox = anc_x - chox * csx
                coy = anc_y - choy * csy
                for grandchild in child:
                    gct = grandchild.tag.split('}')[-1]
                    if   gct == 'sp':    emit_sp(grandchild,    cox, coy, csx, csy, bld)
                    elif gct == 'grpSp': walk_group(grandchild, cox, coy, csx, csy, bld)
                    elif gct == 'cxnSp': emit_cxnsp(grandchild, cox, coy, csx, csy, bld)
                    # pic はスキップ

            elif ct == 'cxnSp':
                # cxnSpはEMU絶対座標（アンカー内でも絶対位置が使われる）
                emit_cxnsp(child, 0, 0, sc, sc, bld)

            # 'pic' (画像) は DrawIO では表現困難なためスキップ


# ══════════════════════════════════════════════════════════════════════════════
#  セルテキストラベル
# ══════════════════════════════════════════════════════════════════════════════

def parse_range_ref(ref):
    if ':' not in ref:
        c, r = cell_ref(ref)
        return c, r, c, r
    m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', ref)
    if not m:
        raise ValueError(ref)
    return (col_letter_to_idx(m.group(1)), int(m.group(2)) - 1,
            col_letter_to_idx(m.group(3)), int(m.group(4)) - 1)


def build_merged_cell_maps(sh_root):
    ns = {'x': SS}
    merged_topleft  = {}
    merged_children = set()
    for mc in sh_root.findall('.//x:mergeCell', ns):
        ref = mc.attrib.get('ref', '')
        if not ref:
            continue
        try:
            c1, r1, c2, r2 = parse_range_ref(ref)
        except Exception:
            continue
        merged_topleft[(r1, c1)] = (r2, c2)
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                if rr != r1 or cc != c1:
                    merged_children.add((rr, cc))
    return merged_topleft, merged_children


def build_cell_value_map(sh_root, shared_strings):
    ns = {'x': SS}
    value_map = {}
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = cell_ref(ref)
            except Exception:
                continue
            t    = cell.attrib.get('t', '')
            v_el = cell.find('x:v', ns)
            if v_el is None or v_el.text is None:
                value_map[(r, c)] = ''
                continue
            if t == 's':
                idx = int(v_el.text)
                value_map[(r, c)] = shared_strings[idx] if idx < len(shared_strings) else ''
            else:
                value_map[(r, c)] = v_el.text
    return value_map


def build_fill_grid(sh_root, xf_fills):
    ns = {'x': SS}
    grid = {}
    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = cell_ref(ref)
            except Exception:
                continue
            s_attr = int(cell.attrib.get('s', 0))
            fc = xf_fills.get(s_attr)
            if fc:
                grid[(r, c)] = fc
    return grid


def format_excel_time(value):
    total_minutes = int(round(value * 24 * 60))
    return f'{total_minutes // 60}:{total_minutes % 60:02d}'


def format_numeric_value(raw, style_numfmt):
    try:
        fv = float(raw)
    except ValueError:
        return raw
    num_fmt_id, fmt_code = style_numfmt
    fmt = (fmt_code or '').lower()
    is_time = (num_fmt_id in {18, 19, 20, 21, 22, 45, 46, 47}
               or ('h' in fmt and 'm' in fmt))
    if is_time:
        return format_excel_time(fv)
    return str(int(fv)) if fv.is_integer() else raw


def estimate_text_units(text):
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


def fit_font_size(text, width, height, base_font_size):
    font_size = max(6, base_font_size)
    while font_size > 6:
        line_cap  = max(1.0, (width - 2) / max(font_size * 0.95, 1))
        req_lines = ceil(estimate_text_units(text) / line_cap)
        max_lines = max(1, int(height / max(font_size * 1.15, 1)))
        if req_lines <= max_lines:
            break
        font_size -= 1
    return font_size


def is_compact_label(text):
    s = str(text).strip()
    if re.fullmatch(r'\d{1,2}[:\uff1a]\d{2}', s):
        return True
    if re.fullmatch(r'\d+', s) and len(s) <= 2:
        return True
    return False


def is_offpage_marker_label(text):
    s = str(text).strip()
    return bool(re.fullmatch(r'[A-Z]{1,2}\d?', s))


def make_cell_text_style(style_info, text, width, height, compact=False):
    eff = dict(style_info)
    if compact:
        eff['align'] = 'center'
        eff['verticalAlign'] = 'middle'
    fsz = fit_font_size(text, width, height, eff.get('fontSize', 10))
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
    parts.append('spacingTop=1' if compact else 'spacingTop=3')
    if not compact and eff.get('align', 'left') == 'left':
        parts.append('spacingLeft=5')
    return ';'.join(parts) + ';'


def add_cell_labels(sh_root, col_x, row_y, col_w_dict, row_h_dict, shared_strings,
                    xf_text_styles, xf_numfmts, xf_fills, bld):
    ns = {'x': SS}
    merged_topleft, merged_children = build_merged_cell_maps(sh_root)
    value_map = build_cell_value_map(sh_root, shared_strings)
    fill_grid = build_fill_grid(sh_root, xf_fills)

    for row_el in sh_root.findall('.//x:row', ns):
        r = int(row_el.attrib.get('r', 1)) - 1
        if r < MIN_ROW or r > MAX_ROW:
            continue
        ry = row_y[min(r, 299)] / SCALE
        rh = max(1.0, pts_px(row_h_dict[r]) / SCALE)

        for cell in row_el.findall('x:c', ns):
            ref = cell.attrib.get('r', '')
            if not ref:
                continue
            try:
                c, _ = cell_ref(ref)
            except Exception:
                continue
            if (r, c) in merged_children:
                continue
            if c < MIN_COL or c > MAX_COL:
                continue

            t    = cell.attrib.get('t', '')
            v_el = cell.find('x:v', ns)
            if v_el is None or v_el.text is None:
                continue

            if t == 's':
                idx = int(v_el.text)
                val = shared_strings[idx] if idx < len(shared_strings) else ''
            elif t == 'str':
                val = v_el.text
            else:
                s_attr = int(cell.attrib.get('s', 0))
                val = format_numeric_value(
                    v_el.text,
                    xf_numfmts.get(s_attr, (0, ''))
                )

            if not val or is_filler(val):
                continue

            cx     = col_x[min(c, 299)] / SCALE
            s_attr = int(cell.attrib.get('s', 0))
            style_info = xf_text_styles.get(s_attr, {})
            compact    = is_compact_label(val)

            if (r, c) in merged_topleft:
                r_end, c_end = merged_topleft[(r, c)]
                cw = max(1.0, (col_x[min(c_end + 1, 300)] - col_x[min(c, 300)]) / SCALE)
                ch = max(1.0, (row_y[min(r_end + 1, 300)] - row_y[min(r, 300)]) / SCALE)
            else:
                c_end      = c
                fill_color = fill_grid.get((r, c))
                if not compact:
                    # 塗りなしセルはテキスト表示に必要な幅を推定し、その幅で延伸を制限する
                    # 塗りありセルは従来通り（同色が続く限り延伸）
                    fsz_est = style_info.get('fontSize', 10)
                    approx_px_needed = estimate_text_units(str(val)) * fsz_est * 0.72 + 10
                    px_accumulated = chars_px(col_w_dict[c])
                    while c_end + 1 <= MAX_COL:
                        next_val  = value_map.get((r, c_end + 1), '')
                        next_fill = fill_grid.get((r, c_end + 1))
                        if next_val:
                            break
                        if fill_color and next_fill != fill_color:
                            break
                        if not fill_color and next_fill:
                            break
                        # 塗りなしのとき: テキスト表示幅に達したら延伸停止
                        if not fill_color and px_accumulated >= approx_px_needed:
                            break
                        c_end += 1
                        px_accumulated += chars_px(col_w_dict[c_end])
                cw = max(1.0, (col_x[min(c_end + 1, 300)] - col_x[min(c, 300)]) / SCALE)
                ch = rh

            # v8: offpage_markerブロックを删除。
            # 先導/後続のSVGはadd_drawing_shapesのみで処理する。

            text_x = cx
            if not compact and style_info.get('align', 'left') == 'left':
                text_x += 2
                cw = max(1.0, cw - 2)

            text_y = ry
            text_h = ch
            if not compact and style_info.get('align', 'left') == 'left':
                text_y += FILLED_TEXT_TOP_PAD
                text_h = max(1.0, ch - FILLED_TEXT_TOP_PAD)

            cell_style = make_cell_text_style(style_info, val, cw, text_h, compact=compact)
            bld.add(val, text_x, text_y, cw, text_h, cell_style, force=True)


# ══════════════════════════════════════════════════════════════════════════════
#  パス解決（シート→描画XML）
# ══════════════════════════════════════════════════════════════════════════════

def find_paths(z, sheet_name):
    wb  = ET.fromstring(z.read('xl/workbook.xml').decode('utf-8'))
    rid = next((sh.attrib.get(f'{{{R}}}id')
                for sh in wb.findall('.//{%s}sheet' % SS)
                if sh.attrib.get('name') == sheet_name), None)
    if not rid:
        available = [s.attrib.get('name') for s in wb.findall('.//{%s}sheet' % SS)]
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {available}")

    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels').decode('utf-8'))
    sf   = next(('xl/' + r.attrib['Target'].lstrip('/')
                 for r in rels if r.attrib.get('Id') == rid), None)

    num       = sf.rsplit('/', 1)[-1].replace('sheet', '').replace('.xml', '')
    rels_path = f'xl/worksheets/_rels/sheet{num}.xml.rels'
    if rels_path not in z.namelist():
        return sf, None

    sr  = ET.fromstring(z.read(rels_path).decode('utf-8'))
    drw = next(('xl/' + r.attrib['Target'].lstrip('../')
                for r in sr
                if 'drawing' in r.attrib.get('Type', '')
                and 'vml' not in r.attrib.get('Type', '')), None)
    return sf, drw


# ══════════════════════════════════════════════════════════════════════════════
#  ログユーティリティ
# ══════════════════════════════════════════════════════════════════════════════

def _log(msg):
    sys.stdout.buffer.write((msg + '\n').encode('utf-8', errors='replace'))


# ══════════════════════════════════════════════════════════════════════════════
#  メインコンバーター
# ══════════════════════════════════════════════════════════════════════════════

def convert(xlsm=INPUT_FILE, sheet=TARGET_SHEET, out=OUTPUT_FILE):
    _log(f"Opening '{xlsm}' ...")

    with zipfile.ZipFile(xlsm, 'r') as z:
        sf, drw_path = find_paths(z, sheet)
        _log(f"Sheet XML: {sf}")
        _log(f"Drawing:   {drw_path or '(none)'}")

        sh_root = ET.fromstring(z.read(sf).decode('utf-8'))
        col_x, row_y, col_w_dict, row_h_dict = build_grid(sh_root)

        # 共有文字列
        ss_root = ET.fromstring(z.read('xl/sharedStrings.xml').decode('utf-8'))
        shared  = [
            ''.join(t.text for t in si.iter(f'{{{SS}}}t') if t.text)
            for si in ss_root.findall(f'{{{SS}}}si')
        ]

        # スタイル解析
        xf_fills        = parse_cell_styles(z)
        xf_borders      = parse_cell_borders(z)
        xf_text_styles  = parse_cell_text_styles(z)
        xf_numfmts      = parse_cell_number_formats(z)
        _log(f"Fill styles:   {len(xf_fills)} xf indices")
        _log(f"Border styles: {len(xf_borders)} xf indices")
        _log(f"Text styles:   {len(xf_text_styles)} xf indices")
        _log(f"Num formats:   {len(xf_numfmts)} xf indices")

        bld = DrawioBuilder()

        # ① セル塗りつぶし（隣接同色を結合）
        _log("Processing cell fills (merging adjacent cells)...")
        fc = add_cell_fills_merged(sh_root, col_x, row_y, col_w_dict, row_h_dict,
                                   xf_fills, bld)
        _log(f"  Merged fill rectangles: {fc}")

        # ② セル罫線
        _log("Processing cell borders...")
        bc = add_cell_borders(sh_root, col_x, row_y, col_w_dict, row_h_dict,
                              xf_borders, xf_fills, bld)
        _log(f"  Border segments: {bc}")

        # ③ 描画図形（Drawing XML）
        if drw_path:
            before = bld._next
            add_drawing_shapes(z, drw_path, col_x, row_y, bld)
            _log(f"Drawing shapes: {bld._next - before}")
        else:
            _log("No drawing found.")

        # ④ セルテキストラベル
        before = bld._next
        add_cell_labels(sh_root, col_x, row_y, col_w_dict, row_h_dict,
                        shared, xf_text_styles, xf_numfmts, xf_fills, bld)
        _log(f"Cell labels: {bld._next - before}")

        _log(f"Total shapes: {bld._next - 2}")

    xml_out = bld.xml()
    with open(out, 'w', encoding='utf-8') as f:
        f.write(xml_out)
    _log(f"Written '{out}' ({len(xml_out):,} chars)")
