"""Drawing shape (sp / grpSp) rendering."""
import html
import xml.etree.ElementTree as ET

from .colors import _parse_drawing_color
from .connectors import _emit_cxnsp, _render_cxnsp_at_rect
from .constants import (
    A,
    GEOM_STYLES,
    XDR,
)
from .geometry import (
    _append_transform_style,
    _custgeom_to_stencil,
    _descend_alt_content,
    _get_text,
    _get_xfrm,
    _ln_style_parts,
    _xfrm_transform,
)
from .grid import _emu_px
from .images import _emit_pic, _extract_images

def _sp_fill(sp_pr, theme):
    if sp_pr.find(f'{{{A}}}noFill') is not None:
        return 'none'
    for fill_tag in (f'{{{A}}}solidFill', f'{{{A}}}gradFill', f'{{{A}}}pattFill'):
        fe = sp_pr.find(fill_tag)
        if fe is not None:
            if fill_tag.endswith('solidFill'):
                c = _parse_drawing_color(fe, theme)
            elif fill_tag.endswith('gradFill'):
                gs = fe.find(f'.//{{{A}}}gs')
                c = _parse_drawing_color(gs, theme) if gs is not None else None
            else:
                bg = fe.find(f'{{{A}}}bgClr')
                c = _parse_drawing_color(bg, theme) if bg is not None else None
            if c:
                return c
    return '#FFFFFF'


def _sp_line(sp_pr, theme):
    ln = sp_pr.find(f'{{{A}}}ln')
    if ln is None:
        return '#000000', 1
    if ln.find(f'{{{A}}}noFill') is not None:
        return 'none', 0
    sf = ln.find(f'{{{A}}}solidFill')
    color = _parse_drawing_color(sf, theme) if sf is not None else '#000000'
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


def _sp_font_style(txb, theme):
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
        fc = _parse_drawing_color(solid, theme)
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


def _sp_text_align(txb):
    """Read horizontal/vertical text alignment from a shape's txBody.

    Returns (align, verticalAlign) using drawio style values, or None when
    OOXML did not specify the attribute (so the caller can keep drawio's
    defaults). Maps:
      a:pPr@algn:    l -> left, ctr -> center, r -> right, just -> justify
      a:bodyPr@anchor: t -> top, ctr -> middle, b -> bottom
    """
    if txb is None:
        return None, None
    align = None
    valign = None
    body_pr = txb.find(f'{{{A}}}bodyPr')
    if body_pr is not None:
        anc = body_pr.attrib.get('anchor')
        valign = {'t': 'top', 'ctr': 'middle', 'b': 'bottom'}.get(anc)
    p = txb.find(f'{{{A}}}p')
    if p is not None:
        ppr = p.find(f'{{{A}}}pPr')
        if ppr is not None:
            algn = ppr.attrib.get('algn')
            align = {'l': 'left', 'ctr': 'center',
                     'r': 'right', 'just': 'justify'}.get(algn)
    return align, valign


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


def _extract_txbody_html(txBody, theme):
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
                    color = _parse_drawing_color(solid, theme)
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


def _render_sp(sp, ax, ay, w, h, bld, theme):
    """Render a shape at a resolved pixel rect, applying transform/line/geom/text."""
    spr = sp.find(f'{{{XDR}}}spPr')
    if spr is None:
        return
    if w < 1 or h < 1:
        return
    fill = _sp_fill(spr, theme)
    lc, lw = _sp_line(spr, theme)
    prst = _sp_geom(spr)
    txb = sp.find(f'{{{XDR}}}txBody')
    html_text, has_rich = _extract_txbody_html(txb, theme)
    plain_text = _get_text(sp)
    text_value = html_text if has_rich else plain_text
    if not text_value and fill in ('#FFFFFF', 'none') and lc == 'none':
        return
    fsz = _sp_fontsize(txb)
    fe, _ = _sp_font_style(txb, theme)

    nv = sp.find(f'{{{XDR}}}nvSpPr/{{{XDR}}}cNvPr')
    sp_id = nv.attrib.get('id') if nv is not None else None

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
    text_align, text_valign = _sp_text_align(txb)
    if text_align:
        extra.append(f'align={text_align}')
    if text_valign:
        extra.append(f'verticalAlign={text_valign}')

    style = _make_shape_style(prst, fill, lc, lw, fsz, fe,
                              shape_override=shape_override, extra_parts=extra)
    bld.add(text_value, ax, ay, w, h, style, force=bool(text_value), sp_id=sp_id)


def _emit_sp(sp, pax, pay, sx, sy, bld, theme):
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
    _render_sp(sp, ax, ay, w, h, bld, theme)


def _walk_group(grp, pax, pay, sx, sy, bld, images, theme, depth=0):
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
            _emit_sp(child, cox, coy, csx, csy, bld, theme)
        elif ct == 'cxnSp':
            _emit_cxnsp(child, cox, coy, csx, csy, bld, theme)
        elif ct == 'grpSp':
            _walk_group(child, cox, coy, csx, csy, bld, images, theme, depth + 1)
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


def _add_drawing_shapes(z, drawing_path, col_x, row_y, bld, cfg, theme):
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
                _render_sp(child, anc_x, anc_y, anc_w, anc_h, bld, theme)
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
                        _emit_sp(gc, cox, coy, csx, csy, bld, theme)
                    elif gct == 'grpSp':
                        _walk_group(gc, cox, coy, csx, csy, bld, images, theme)
                    elif gct == 'cxnSp':
                        _emit_cxnsp(gc, cox, coy, csx, csy, bld, theme)
                    elif gct == 'pic':
                        _emit_pic(gc, images, cox, coy, csx, csy, bld)
            elif ct == 'cxnSp':
                # Anchor-level connector: use the _anchor_rect bbox so the
                # line endpoints align with the same pixel math the cell
                # labels and shapes use.
                # Also pass the exact from/to pixel corners so _render_cxnsp_at_rect
                # can skip xfrm rot/flip recomputation (the anchor already encodes
                # the final on-sheet position).
                cx_last = len(col_x) - 1
                ry_last = len(row_y) - 1
                from_el_c = anchor.find(f'{{{XDR}}}from')
                to_el_c   = anchor.find(f'{{{XDR}}}to')
                from_corner_c = None
                to_corner_c   = None
                if from_el_c is not None:
                    ffc  = int(from_el_c.findtext(f'{{{XDR}}}col',     '0') or '0')
                    ffco = int(from_el_c.findtext(f'{{{XDR}}}colOff', '0') or '0')
                    ffr  = int(from_el_c.findtext(f'{{{XDR}}}row',     '0') or '0')
                    ffro = int(from_el_c.findtext(f'{{{XDR}}}rowOff', '0') or '0')
                    from_corner_c = (
                        col_x[min(ffc, cx_last)] / cfg.scale + _emu_px(ffco, cfg),
                        row_y[min(ffr, ry_last)] / cfg.scale + _emu_px(ffro, cfg),
                    )
                if to_el_c is not None:
                    ftc  = int(to_el_c.findtext(f'{{{XDR}}}col',     '0') or '0')
                    ftco = int(to_el_c.findtext(f'{{{XDR}}}colOff', '0') or '0')
                    ftr  = int(to_el_c.findtext(f'{{{XDR}}}row',     '0') or '0')
                    ftro = int(to_el_c.findtext(f'{{{XDR}}}rowOff', '0') or '0')
                    to_corner_c = (
                        col_x[min(ftc, cx_last)] / cfg.scale + _emu_px(ftco, cfg),
                        row_y[min(ftr, ry_last)] / cfg.scale + _emu_px(ftro, cfg),
                    )
                _render_cxnsp_at_rect(child, anc_x, anc_y, anc_w, anc_h, bld, theme,
                                      from_corner=from_corner_c,
                                      to_corner=to_corner_c)
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

