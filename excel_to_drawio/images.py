"""Drawing image extraction and picture emission."""
import base64
import xml.etree.ElementTree as ET

from .constants import (
    A,
    ASVG,
    R,
    RENDERABLE_IMG_EXTS,
    XDR,
)
from .geometry import _append_transform_style, _xfrm_transform

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
    except (ET.ParseError, UnicodeDecodeError):
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
        except (KeyError, zipfile.BadZipFile):
            images[rid] = None
            continue
        b64 = base64.b64encode(img_data).decode('ascii')
        images[rid] = f'data:{mime};base64,{b64}'
    return images


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
    nv = pic.find(f'{{{XDR}}}nvPicPr/{{{XDR}}}cNvPr')
    sp_id = nv.attrib.get('id') if nv is not None else None

    _render_pic_at_rect(ax, ay, w, h, data_uri,
                        (primary_rid or chosen_rid), xfrm, bld, sp_id=sp_id)


def _render_pic_at_rect(ax, ay, w, h, data_uri, has_ref, xfrm, bld, sp_id=None):
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
        bld.add_image(ax, ay, w, h, data_uri, extra_style=extra_style, sp_id=sp_id)
        return
    if has_ref:
        placeholder_parts = [
            'whiteSpace=wrap', 'html=1', 'fillColor=#F5F5F5',
            'strokeColor=#BDBDBD', 'dashed=1', 'align=center',
            'verticalAlign=middle', 'fontSize=8', 'fontColor=#757575',
        ]
        if extras:
            placeholder_parts.extend(extras)
        bld.add('[image]', ax, ay, w, h, ';'.join(placeholder_parts) + ';', sp_id=sp_id)

