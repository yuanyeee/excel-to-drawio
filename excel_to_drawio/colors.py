"""Color resolution and the per-conversion Theme."""
import dataclasses
import xml.etree.ElementTree as ET

from .constants import (
    A,
    INDEXED_COLORS,
    SCHEME_COLORS,
    SKIP_COLORS,
    THEME_FILL_COLORS,
    THEME_INDEX_NAMES,
)

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
    except (ValueError, TypeError):
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
    except (ValueError, TypeError):
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


@dataclasses.dataclass
class Theme:
    """Per-conversion color theme (no global mutable state).

    Holds the workbook's scheme colors, the positional theme-fill list used
    for cell fills, and the indexed color table. A fresh instance is built for
    every conversion so a workbook with a custom theme can never leak its
    colors into a subsequent conversion of a default-themed workbook.
    """
    scheme: dict = dataclasses.field(default_factory=lambda: dict(SCHEME_COLORS))
    fill_colors: list = dataclasses.field(default_factory=lambda: list(THEME_FILL_COLORS))
    indexed: list = dataclasses.field(default_factory=lambda: list(INDEXED_COLORS))


def load_theme(z):
    """Build a Theme from the workbook's xl/theme/theme1.xml.

    DrawingML stores 12 scheme colors in a:clrScheme: dk1, lt1, dk2, lt2,
    accent1..accent6, hlink, folHlink. Each child wraps either an srgbClr or a
    sysClr lastClr. A missing or unparseable theme leaves the Office defaults
    untouched (the Theme is built fresh from the defaults first, so no state
    carries over from a previous conversion).
    """
    theme = Theme()
    candidates = [n for n in z.namelist()
                  if n.startswith('xl/theme/theme') and n.endswith('.xml')]
    if not candidates:
        return theme
    try:
        root = ET.fromstring(z.read(sorted(candidates)[0]).decode('utf-8'))
    except (ET.ParseError, UnicodeDecodeError, KeyError):
        return theme
    scheme = root.find(f'.//{{{A}}}clrScheme')
    if scheme is None:
        return theme
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
            theme.scheme[name] = hex6
            if name.startswith('accent') and name[-1].isdigit():
                theme.scheme['acc' + name[-1]] = hex6
    # Refresh aliases that mirror the primary names.
    theme.scheme['bg1'] = theme.scheme.get('lt1', theme.scheme['bg1'])
    theme.scheme['bg2'] = theme.scheme.get('lt2', theme.scheme['bg2'])
    theme.scheme['tx1'] = theme.scheme.get('dk1', theme.scheme['tx1'])
    theme.scheme['tx2'] = theme.scheme.get('dk2', theme.scheme['tx2'])
    # Refresh positional theme list used by cell fills.
    theme.fill_colors = [theme.scheme[n] for n in THEME_INDEX_NAMES]
    return theme


def _parse_color_el(color_el, theme, default='#000000'):
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
    theme_idx = color_el.attrib.get('theme', '')
    if theme_idx:
        idx = int(theme_idx)
        base = theme.fill_colors[idx] if idx < len(theme.fill_colors) else None
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
        icolor = theme.indexed[idx] if idx < len(theme.indexed) else None
        if icolor:
            return '#' + icolor
    return default


def _should_skip_fill(hex6):
    return hex6.upper().lstrip('#') in SKIP_COLORS


def _parse_drawing_color(el, theme):
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
        base = theme.scheme.get(name, '808080')
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

