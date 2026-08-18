"""DrawingML geometry / transform helpers."""
import base64

from .constants import (
    A,
    ARROW_MAP,
    MC,
    PRST_DASH_MAP,
)

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
    # OOXML headEnd/tailEnd semantics appear reversed relative to drawio's
    # startArrow/endArrow in our coordinate path emission. Map head->end and
    # tail->start so arrowheads land on the expected visual side.
    if has_head:
        htype = head.attrib.get('type', 'none')
        parts.append(f'endArrow={ARROW_MAP.get(htype, "classic")}')
    if has_tail:
        ttype = tail.attrib.get('type', 'none')
        parts.append(f'startArrow={ARROW_MAP.get(ttype, "classic")}')
    prst = ln.find(f'{{{A}}}prstDash')
    if prst is not None:
        pval = prst.attrib.get('val', 'solid')
        dp = PRST_DASH_MAP.get(pval)
        if dp:
            parts.append('dashed=1')
            parts.append(f'dashPattern={dp}')
    return parts, has_head, has_tail


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

