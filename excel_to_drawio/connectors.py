"""Connector (cxnSp) rendering."""
import re

from .colors import _parse_drawing_color
from .constants import (
    A,
    XDR,
)
from .geometry import _ln_style_parts, _rotate_point, _xfrm_transform

def _conn_side(preset, idx):
    """Map a stCxn/endCxn connection-point index to a drawio side.

    Most presets number the four cardinal points 0=top, 1=left, 2=bottom,
    3=right. The "can" (cylinder) preset adds a leading point, shifting the
    cardinal points to 1=top, 2=left, 3=bottom, 4=right.
    """
    try:
        idx = int(idx)
    except (TypeError, ValueError):
        return None
    if preset == 'can':
        return {0: 'top', 1: 'top', 2: 'left', 3: 'bottom', 4: 'right'}.get(idx)
    return {0: 'top', 1: 'left', 2: 'bottom', 3: 'right'}.get(idx)


def _render_cxnsp_at_rect(cxn, ax, ay, w, h, bld, theme, from_corner=None, to_corner=None, shape_preset=None):
    """Emit a connector as a drawio edge for a pre-resolved bbox rect.

    ``ax/ay/w/h`` is the connector's bounding box in drawio pixels.
    ``from_corner`` and ``to_corner`` are optional (x, y) absolute pixel
    coordinates for the true start/end of the connector (derived from the
    ``xdr:from`` / ``xdr:to`` anchor elements).  When provided they replace
    the flip-based endpoint heuristic so that anchor-level connectors always
    have exactly the right terminal positions.  The elbow waypoint is still
    derived from the anchor bbox corners.
    Used by both the anchor-level path (which derives the rect from
    ``_anchor_rect``, so it shares pixel math with shapes and cell labels)
    and the group-level path in ``_emit_cxnsp``.
    """
    spr = cxn.find(f'{{{XDR}}}spPr')
    if spr is None:
        return
    prst_el = spr.find(f'{{{A}}}prstGeom')
    prst_name = prst_el.attrib.get('prst', '') if prst_el is not None else ''
    nv_sp = cxn.find(f'{{{XDR}}}nvCxnSpPr/{{{XDR}}}cNvPr')
    cxn_id = nv_sp.attrib.get('id') if nv_sp is not None else None

    cnv = cxn.find(f'{{{XDR}}}nvCxnSpPr/{{{XDR}}}cNvCxnSpPr')
    has_bound_end = False
    src_id = None
    tgt_id = None
    src_idx = None
    tgt_idx = None
    if cnv is not None:
        st = cnv.find(f'{{{A}}}stCxn')
        ed = cnv.find(f'{{{A}}}endCxn')
        has_bound_end = (st is not None or ed is not None)
        if st is not None:
            src_id = st.attrib.get('id')
            src_idx = st.attrib.get('idx')
        if ed is not None:
            tgt_id = ed.attrib.get('id')
            tgt_idx = ed.attrib.get('idx')
    xfrm = spr.find(f'{{{A}}}xfrm')
    rot, fh, fv = _xfrm_transform(xfrm)
    edge_points = None

    # When the caller supplies exact from/to corners (anchor-level connectors)
    # we use them directly and skip all flip/rot heuristics.  The anchor bbox
    # already encodes the on-sheet position; applying xfrm rot/flip on top of
    # it re-rotates coordinates that have already been placed correctly and
    # produces wrong elbow positions or negative coordinates.
    if from_corner is not None and to_corner is not None:
        x1, y1 = from_corner
        x2, y2 = to_corner
        if prst_name.startswith('bentConnector'):
            m = re.search(r'(\d+)$', prst_name or '')
            idx = int(m.group(1)) if m else 2
            adj = None
            avlst = prst_el.find(f'{{{A}}}avLst') if prst_el is not None else None
            if avlst is not None:
                for gd in avlst.findall(f'{{{A}}}gd'):
                    if gd.attrib.get('name') != 'adj1':
                        continue
                    raw = gd.attrib.get('fmla', '')
                    if raw.startswith('val '):
                        raw = raw.split()[-1]
                    if not raw:
                        raw = gd.attrib.get('val', '')
                    try:
                        # Keep adj1 in a sane [-2, 2] window: legitimate
                        # elbow detours run a little outside [0,1] (e.g.
                        # -0.29, 1.03 for same-side routes), but corrupted
                        # values like -1109 / +21139 would blow up the
                        # waypoint and inflate the page to ~1M px.
                        adj = max(-2.0, min(2.0, int(raw) / 100000.0))
                    except (ValueError, TypeError):
                        adj = None
                    break
            # OOXML bent-connector presets are 3-segment routes with TWO
            # elbows, adjustable via adj1 (default 0.5):
            #   bentConnector2: right -> down -> right  (horizontal first)
            #       elbows at (xb, y1) and (xb, y2),  xb = x1 + (x2-x1)*adj1
            #   bentConnector3: down -> right -> down   (vertical first)
            #       elbows at (x1, yb) and (x2, yb),  yb = y1 + (y2-y1)*adj1
            # Route so that BOTH ends leave/enter perpendicular to their
            # shapes, otherwise the line hugs the boundary or wraps around.
            # The OOXML preset name alone doesn't reliably encode the axis
            # (bentConnector3 is used for both vertical and horizontal
            # exits). bentConnector4/5 are the fixed 2-segment variants.
            # adj1 may lie outside [0,1] (e.g. -0.29, 1.03): a negative value
            # pushes the elbow past the head so the route turns AWAY from the
            # exit side instead of hugging the shape boundary (same-side links).
            exit_side = _conn_side(shape_preset.get(src_id) if shape_preset else None, src_idx)
            entry_side = _conn_side(shape_preset.get(tgt_id) if shape_preset else None, tgt_idx)
            if exit_side is not None and entry_side is not None:
                exit_h = exit_side in ('left', 'right')
                entry_h = entry_side in ('left', 'right')
                if exit_h != entry_h:
                    # Mixed axes -> L-shape (2 segments) so the entry is
                    # perpendicular too.
                    edge_points = [(x2, y1)] if exit_h else [(x1, y2)]
                else:
                    horizontal_first = exit_h
            else:
                # One or both sides unknown: use whichever side is known,
                # else fall back to the preset.
                known = exit_side if exit_side is not None else entry_side
                if known in ('left', 'right'):
                    horizontal_first = True
                elif known in ('top', 'bottom'):
                    horizontal_first = False
                else:
                    horizontal_first = (idx in (2, 4))
            if edge_points is None:
                if idx in (2, 3):
                    if adj is None:
                        adj = 0.5
                    if horizontal_first:
                        # right/left -> down/up -> right/left
                        xb = x1 + (x2 - x1) * adj
                        edge_points = [(xb, y1), (xb, y2)]
                    else:
                        # down/up -> right/left -> down/up
                        yb = y1 + (y2 - y1) * adj
                        edge_points = [(x1, yb), (x2, yb)]
                elif horizontal_first:
                    edge_points = [(x2, y1)]
                else:
                    edge_points = [(x1, y2)]
    else:
        if prst_name.startswith('bentConnector'):
            # OOXML bent connectors are orthogonal/elbow polylines.
            # Keep opposite-corner endpoints and provide explicit waypoint(s).
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
            adj = None
            avlst = prst_el.find(f'{{{A}}}avLst') if prst_el is not None else None
            if avlst is not None:
                for gd in avlst.findall(f'{{{A}}}gd'):
                    if gd.attrib.get('name') != 'adj1':
                        continue
                    raw = gd.attrib.get('fmla', '')
                    if raw.startswith('val '):
                        raw = raw.split()[-1]
                    if not raw:
                        raw = gd.attrib.get('val', '')
                    try:
                        # Keep adj1 in a sane [-2, 2] window: legitimate
                        # elbow detours run a little outside [0,1] (e.g.
                        # -0.29, 1.03 for same-side routes), but corrupted
                        # values like -1109 / +21139 would blow up the
                        # waypoint and inflate the page to ~1M px.
                        adj = max(-2.0, min(2.0, int(raw) / 100000.0))
                    except (ValueError, TypeError):
                        adj = None
                    break
            if has_bound_end and adj is None:
                edge_points = None
            else:
                # Even idx (2, 4) -> horizontal first; odd idx (3, 5) -> vertical first.
                if idx in (3, 5):
                    if adj is not None:
                        yb = y1 + (y2 - y1) * adj
                        edge_points = [(x1, yb)]
                    else:
                        elbow_x = ax + w if fh else ax
                        elbow_y = ay if fv else ay + h
                        edge_points = [(elbow_x, elbow_y)]
                elif idx in (2, 4):
                    if adj is not None:
                        xb = x1 + (x2 - x1) * adj
                        edge_points = [(xb, y1)]
                    else:
                        elbow_x = ax if fh else ax + w
                        elbow_y = ay + h if fv else ay
                        edge_points = [(elbow_x, elbow_y)]
                else:
                    if adj is None:
                        elbow_x = ax + w if fh else ax
                        elbow_y = ay if fv else ay + h
                        edge_points = [(elbow_x, elbow_y)]
                    else:
                        edge_points = [(x1, y1 + (y2 - y1) * adj)]
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
        if prst_name.startswith('bentConnector') and rot:
            q = round(rot / 90.0)
            snapped = q * 90.0
            if abs(rot - snapped) <= 1.0:
                eff_rot = snapped
        if eff_rot:
            cx, cy = ax + (w / 2.0), ay + (h / 2.0)
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
        color = _parse_drawing_color(sf, theme) if sf is not None else '#000000'
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
        # Free connectors are stabilized by explicit points; shape-bound
        # connectors are better left on orthogonal routing.
        if has_bound_end:
            parts.append('edgeStyle=orthogonalEdgeStyle')
        else:
            parts.append('edgeStyle=none')
    elif prst_name.startswith('curvedConnector'):
        parts.append('edgeStyle=none')
        parts.append('curved=1')
    elif prst_name.startswith('straightConnector'):
        parts.append('edgeStyle=none')
        
    src_side = _conn_side(shape_preset.get(src_id) if shape_preset else None, src_idx)
    tgt_side = _conn_side(shape_preset.get(tgt_id) if shape_preset else None, tgt_idx)
    if src_side == 'top':
        parts.append('exitX=0.5;exitY=0;exitDx=0;exitDy=0')
    elif src_side == 'left':
        parts.append('exitX=0;exitY=0.5;exitDx=0;exitDy=0')
    elif src_side == 'bottom':
        parts.append('exitX=0.5;exitY=1;exitDx=0;exitDy=0')
    elif src_side == 'right':
        parts.append('exitX=1;exitY=0.5;exitDx=0;exitDy=0')
    if tgt_side == 'top':
        parts.append('entryX=0.5;entryY=0;entryDx=0;entryDy=0')
    elif tgt_side == 'left':
        parts.append('entryX=0;entryY=0.5;entryDx=0;entryDy=0')
    elif tgt_side == 'bottom':
        parts.append('entryX=0.5;entryY=1;entryDx=0;entryDy=0')
    elif tgt_side == 'right':
        parts.append('entryX=1;entryY=0.5;entryDx=0;entryDy=0')

    parts.append(f'strokeColor={color}')
    if lw_px > 1:
        parts.append(f'strokeWidth={lw_px}')
    # OOXML tailEnd is at the line's geometric end (the "pointing" side); it
    # maps to drawio's endArrow. headEnd maps to startArrow. _ln_style_parts
    # emits the inverted mapping for legacy reasons, so swap it back here for
    # all connector paths (anchor-level, bound, and free) since the coordinate
    # emission above already places drawio's start at OOXML's head and drawio's
    # end at OOXML's tail under both flip-aware and corner-supplied routing.
    if not has_tail:
        parts.append('endArrow=none')
    if not has_head:
        parts.append('startArrow=none')
    remapped = []
    for p in ln_parts:
        if p.startswith('startArrow='):
            remapped.append('endArrow=' + p[len('startArrow='):])
        elif p.startswith('endArrow='):
            remapped.append('startArrow=' + p[len('endArrow='):])
        else:
            remapped.append(p)
    parts.extend(remapped)
    style = ';'.join(parts) + ';'
    bld.add_edge(x1, y1, x2, y2, style, points=edge_points, src_id=src_id, tgt_id=tgt_id, edge_id=cxn_id)


def _emit_cxnsp(cxn, pax, pay, sx, sy, bld, theme, shape_preset=None):
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
    _render_cxnsp_at_rect(cxn, ax, ay, w, h, bld, theme, shape_preset=shape_preset)

