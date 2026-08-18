"""Low-level OOXML reading helpers."""
import xml.etree.ElementTree as ET
from pathlib import Path

from .constants import (
    R,
    SS,
)

def _extract_string_item_text(node):
    # Concatenate visible text from <si>/<is>, excluding <rPh> furigana (ruby).
    parts = []
    direct_t = node.find(f'{{{SS}}}t')
    if direct_t is not None and direct_t.text:
        parts.append(direct_t.text)
    for r in node.findall(f'{{{SS}}}r'):
        rt = r.find(f'{{{SS}}}t')
        if rt is not None and rt.text:
            parts.append(rt.text)
    return ''.join(parts)


def _read_cell_raw_text(cell, shared_strings):
    ns = {'x': SS}
    cell_type = cell.attrib.get('t', '')
    if cell_type == 'inlineStr':
        inline = cell.find('x:is', ns)
        if inline is None:
            return ''
        return _extract_string_item_text(inline)
    v_el = cell.find('x:v', ns)
    if v_el is None or v_el.text is None:
        return ''
    if cell_type == 's':
        idx = int(v_el.text)
        return shared_strings[idx] if idx < len(shared_strings) else ''
    return v_el.text


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
        _extract_string_item_text(si)
        for si in ss_root.findall(f'{{{SS}}}si')
    ]


def _validate_workbook_suffix(input_path):
    suffix = Path(input_path).suffix.lower()
    if suffix not in {'.xlsx', '.xlsm'}:
        raise ValueError('Supported file types are .xlsx and .xlsm')

