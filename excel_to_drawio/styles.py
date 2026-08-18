"""Cell style parsing and rendering (fills, borders, labels)."""
import html
import re
import xml.etree.ElementTree as ET
from math import ceil

from .colors import _parse_color_el, _should_skip_fill
from .constants import (
    BORDER_STYLE_MAP,
    R,
    SS,
)
from .grid import (
    _cell_ref,
    _chars_px,
    _normalize_font_name,
    _parse_range_ref,
    _pts_px,
)
from .ooxml import _read_cell_raw_text

def _parse_cell_styles(z, theme, log):
    """Parse styles.xml: xf_index -> fill color '#RRGGBB'."""
    xf_fills = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except (ET.ParseError, UnicodeDecodeError, KeyError) as exc:
        log(f'warning: could not parse styles.xml (fills): {exc}')
        return xf_fills
    ns = {'x': SS}
    fills = []
    for fill_el in root.findall('.//x:fills/x:fill', ns):
        color = None
        pf = fill_el.find(f'{{{SS}}}patternFill')
        if pf is not None and pf.attrib.get('patternType', 'none') != 'none':
            fg = pf.find(f'{{{SS}}}fgColor')
            if fg is not None:
                c = _parse_color_el(fg, theme, default=None)
                if c and not _should_skip_fill(c):
                    color = c
        fills.append(color)
    for i, xf in enumerate(root.findall('.//x:cellXfs/x:xf', ns)):
        fill_id = int(xf.attrib.get('fillId', '0'))
        if fill_id < len(fills) and fills[fill_id]:
            xf_fills[i] = fills[fill_id]
    return xf_fills


def _parse_cell_borders(z, theme, log):
    """Parse styles.xml: xf_index -> {side: (color, width, dash_pattern)}."""
    xf_borders = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except (ET.ParseError, UnicodeDecodeError, KeyError) as exc:
        log(f'warning: could not parse styles.xml (borders): {exc}')
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
            color = _parse_color_el(sel.find(f'{{{SS}}}color'), theme)
            bw, dash = BORDER_STYLE_MAP.get(sname, (1, None))
            sides[side] = (color, bw, dash, sname)
        border_defs.append(sides)
    for i, xf in enumerate(root.findall('.//x:cellXfs/x:xf', ns)):
        bid = int(xf.attrib.get('borderId', '0'))
        if 0 <= bid < len(border_defs) and border_defs[bid]:
            xf_borders[i] = border_defs[bid]
    return xf_borders


def _parse_cell_text_styles(z, theme, log):
    """Parse styles.xml: xf_index -> text style dict (font, alignment, underline, strike)."""
    xf_text_styles = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except (ET.ParseError, UnicodeDecodeError, KeyError) as exc:
        log(f'warning: could not parse styles.xml (text styles): {exc}')
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
            'fontColor': _parse_color_el(color_el, theme, default='#000000'),
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


def _parse_cell_number_formats(z, log):
    """Parse styles.xml: xf_index -> (numFmtId, formatCode)."""
    xf_numfmts = {}
    try:
        root = ET.fromstring(z.read('xl/styles.xml').decode('utf-8'))
    except (ET.ParseError, UnicodeDecodeError, KeyError) as exc:
        log(f'warning: could not parse styles.xml (number formats): {exc}')
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
        except (ValueError, TypeError):
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
            except (ValueError, TypeError):
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
            except (ValueError, TypeError):
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
    except (ET.ParseError, UnicodeDecodeError, KeyError):
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
        except (ET.ParseError, UnicodeDecodeError, KeyError):
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
            except (ValueError, TypeError):
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
            except (ValueError, TypeError):
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
            except (ValueError, TypeError):
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
            except (ValueError, TypeError):
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
                style = (f'shape=line;html=1;strokeColor={color};'
                         f'strokeWidth={width_px};{dash_style}')
                # Borders are drawn purely as lines around the cell perimeter.
                pass_w = max(width_px, 1)
                if side == 'top':
                    bld.add('', cx, cy, cw, 1, style)
                elif side == 'bottom':
                    # To align outer edge roughly with Excel logic, draw on bottom boundary
                    bld.add('', cx, cy + ch - 1, cw, 1, style)
                elif side == 'left':
                    bld.add('', cx, cy, 1, ch, style + 'direction=south;')
                elif side == 'right':
                    bld.add('', cx + cw - 1, cy, 1, ch, style + 'direction=south;')
                count += 1
    return count


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
            except (ValueError, TypeError):
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

