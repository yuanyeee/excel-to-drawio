"""Drawio XML builder."""
import html

class DrawioBuilder:
    def __init__(self, diagram_name='Sheet1', page_mode=True):
        self._cells = []
        self._next = 2
        self._seen = {}  # key -> cid
        self._emitted_cids = set()
        self._max_x = 0
        self._max_y = 0
        self._diagram_name = diagram_name
        self._page_mode = page_mode
        self._ooxml_map = {}

    def get_solid_cid(self, ooxml_id):
        if not ooxml_id:
            cid = self._next
            self._next += 1
            return cid
        if ooxml_id not in self._ooxml_map:
            self._ooxml_map[ooxml_id] = self._next
            self._next += 1
        return self._ooxml_map[ooxml_id]

    def add(self, text, x, y, w, h, style, force=False, sp_id=None):
        x, y = round(x), round(y)
        w, h = round(max(w, 1)), round(max(h, 1))
        # Ensure distinct instances are kept distinct if they have an ID
        key = (x, y, w, h, style[:60], sp_id)
        if key in self._seen and not force and sp_id is None:
            # If skipping, an ooxml_id wasn't provided, so we don't need to remap it
            return self._seen[key]
        
        cid = self.get_solid_cid(sp_id)
        if cid in self._emitted_cids:
            return cid

        self._seen[key] = cid
        self._emitted_cids.add(cid)

        self._max_x = max(self._max_x, x + w)
        self._max_y = max(self._max_y, y + h)
        esc = html.escape(str(text))
        self._cells.append(
            f'    <mxCell id="{cid}" value="{esc}" style="{style}" vertex="1" parent="1">'
            f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/>'
            f'</mxCell>'
        )

    def add_image(self, x, y, w, h, data_uri, extra_style=None, sp_id=None):
        """Add an embedded image as a DrawIO image shape.

        ``extra_style`` may be a string of already-joined style fragments (no
        leading/trailing ``;``) such as ``"rotation=90;flipH=1"`` that will be
        appended after the default image style.
        """
        x, y = round(x), round(y)
        w, h = round(max(w, 1)), round(max(h, 1))

        cid = self.get_solid_cid(sp_id)
        if cid in self._emitted_cids:
            return cid
        self._emitted_cids.add(cid)

        self._max_x = max(self._max_x, x + w)
        self._max_y = max(self._max_y, y + h)
        style = (f'shape=image;verticalLabelPosition=bottom;labelBackgroundColor=default;'
                 f'verticalAlign=top;aspect=fixed;imageAspect=0;'
                 f'image={data_uri};')
        if extra_style:
            style += extra_style.strip(';') + ';'
        self._cells.append(
            f'    <mxCell id="{cid}" value="" style="{style}" vertex="1" parent="1">'
            f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/>'
            f'</mxCell>'
        )

    def add_edge(self, x1, y1, x2, y2, style, points=None, src_id=None, tgt_id=None, edge_id=None):
        """Add a drawio edge (line) between two explicit points.

        Used for connector shapes (``xdr:cxnSp``) so they render as real lines
        instead of collapsed vertex rectangles. If src_id/tgt_id are provided,
        they bind the edge to existing shapes allowing orthogonal routing heuristics.
        """
        x1, y1 = round(x1), round(y1)
        x2, y2 = round(x2), round(y2)

        cid = self.get_solid_cid(edge_id)
        if cid in self._emitted_cids:
            return cid
        self._emitted_cids.add(cid)

        self._max_x = max(self._max_x, x1, x2)
        self._max_y = max(self._max_y, y1, y2)

        points_xml = ''
        if points:
            pts = []
            for px, py in points:
                pts.append(f'<mxPoint x="{round(px)}" y="{round(py)}"/>')
                self._max_x = max(self._max_x, round(px))
                self._max_y = max(self._max_y, round(py))
            points_xml = f'<Array as="points">{"".join(pts)}</Array>'
        src_attr = f' source="{self.get_solid_cid(src_id)}"' if src_id else ''
        tgt_attr = f' target="{self.get_solid_cid(tgt_id)}"' if tgt_id else ''
        self._cells.append(
            f'    <mxCell id="{cid}" value="" style="{style}" edge="1" parent="1"{src_attr}{tgt_attr}>'
            f'<mxGeometry relative="1" as="geometry">'
            f'<mxPoint x="{x1}" y="{y1}" as="sourcePoint"/>'
            f'<mxPoint x="{x2}" y="{y2}" as="targetPoint"/>'
            f'{points_xml}'
            f'</mxGeometry>'
            f'</mxCell>'
        )

    def diagram_xml(self, diagram_id='d1'):
        page_w = max(2000, int(self._max_x * 1.10))
        page_h = max(2000, int(self._max_y * 1.10))
        hdr = (
            f'  <diagram id="{diagram_id}" name="{html.escape(str(self._diagram_name))}">\n'
            '    <mxGraphModel grid="0" guides="1" tooltips="1" connect="1" arrows="1"\n'
            f'                  fold="1" page="{"1" if self._page_mode else "0"}" pageScale="1" pageWidth="{page_w}"\n'
            f'                  pageHeight="{page_h}" math="0" shadow="0">\n'
            '      <root>\n'
            '        <mxCell id="0"/>\n'
            '        <mxCell id="1" parent="0"/>\n'
        )
        ftr = '      </root>\n    </mxGraphModel>\n  </diagram>\n'
        return hdr + '\n'.join(self._cells) + '\n' + ftr

