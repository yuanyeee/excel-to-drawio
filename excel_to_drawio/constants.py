"""OOXML namespaces and static lookup tables."""

XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'


A = 'http://schemas.openxmlformats.org/drawingml/2006/main'


R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


ASVG = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'


MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'


A14 = 'http://schemas.microsoft.com/office/drawing/2010/main'


RENDERABLE_IMG_EXTS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'webp'}


REL = 'http://schemas.openxmlformats.org/package/2006/relationships'


PRST_DASH_MAP = {
    'solid': None,
    'dash': '8 4',
    'dot': '1 4',
    'dashDot': '8 4 1 4',
    'lgDash': '16 4',
    'sysDash': '5 2',
    'sysDot': '1 2',
    'lgDashDot': '16 4 1 4',
    'lgDashDotDot': '16 4 1 4 1 4',
    'dashDotDot': '8 4 1 4 1 4',
    'sysDashDot': '5 2 1 2',
    'sysDashDotDot': '5 2 1 2 1 2',
}


ARROW_MAP = {
    'triangle': 'classic',
    'arrow': 'classic',
    'stealth': 'classicThin',
    'diamond': 'diamondThin',
    'oval': 'oval',
    'none': 'none',
}


SCHEME_COLORS = {
    'dk1': '000000', 'lt1': 'FFFFFF', 'dk2': '44546A', 'lt2': 'E7E6E6',
    'accent1': '4472C4', 'accent2': 'ED7D31', 'accent3': 'A5A5A5',
    'accent4': 'FFC000', 'accent5': '5B9BD5', 'accent6': '70AD47',
    'hlink': '0563C1', 'folHlink': '954F72',
    'bg1': 'FFFFFF', 'bg2': 'E7E6E6', 'tx1': '000000', 'tx2': '44546A',
    'phClr': 'FFFFFF',
    # Legacy short aliases — keep for backwards compatibility.
    'acc1': '4472C4', 'acc2': 'ED7D31', 'acc3': 'A5A5A5',
    'acc4': 'FFC000', 'acc5': '5B9BD5', 'acc6': '70AD47',
}


THEME_INDEX_NAMES = [
    'lt1', 'dk1', 'lt2', 'dk2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink', 'folHlink',
]


THEME_FILL_COLORS = [SCHEME_COLORS[n] for n in THEME_INDEX_NAMES]


INDEXED_COLORS = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
    'FFFFFF', 'FFFFFF',
]


SKIP_COLORS = {
    'FFFFFF', 'FFFFFE', 'F2F2F2', 'F3F3F3', 'EBEBEB', 'E7E6E6', 'EEECE1',
    'D9D9D9', 'BFBFBF', '000000', '0D0D0D',
}


GEOM_STYLES = {
    'rect': '',
    'roundRect': 'rounded=1;arcSize=10;',
    'ellipse': 'ellipse;',
    'diamond': 'rhombus;',
    'triangle': 'triangle;',
    'parallelogram': 'parallelogram;',
    'trapezoid': 'trapezoid;',
    'hexagon': 'hexagon;',
    'octagon': 'octagon;',
    'star5': 'shape=mxgraph.basic.star;',
    'cloud': 'shape=cloud;',
    'heart': 'shape=mxgraph.basic.heart;',
    'can': 'shape=cylinder3;',
    'cube': 'shape=cube;',
    'bevel': 'shape=mxgraph.basic.rounded_frame;',
    'donut': 'shape=mxgraph.basic.donut;',
    'noSmoking': 'shape=mxgraph.basic.no_symbol;',
    'blockArc': 'shape=mxgraph.basic.arc;',
    'foldedCorner': 'shape=note;',
    'frame': 'shape=mxgraph.basic.frame;',
    'plaque': 'shape=mxgraph.basic.plaque;',
    # Flowchart
    'flowChartProcess': 'shape=mxgraph.flowchart.process;',
    'flowChartDecision': 'shape=mxgraph.flowchart.decision;',
    'flowChartTerminator': 'shape=mxgraph.flowchart.terminator;',
    'flowChartManualInput': 'shape=mxgraph.flowchart.manual_input;',
    'flowChartDocument': 'shape=mxgraph.flowchart.document;',
    'flowChartPredefinedProcess': 'shape=mxgraph.flowchart.predefined_process;',
    'flowChartConnector': 'ellipse;',
    'flowChartOffpageConnector': 'shape=offPageConnector;',
    'flowChartPunchedTape': 'shape=mxgraph.flowchart.punched_tape;',
    'flowChartSort': 'shape=mxgraph.flowchart.sort;',
    'flowChartPreparation': 'shape=mxgraph.flowchart.preparation;',
    'flowChartManualOperation': 'shape=mxgraph.flowchart.manual_operation;',
    'flowChartMerge': 'shape=mxgraph.flowchart.merge;',
    'flowChartInternalStorage': 'shape=mxgraph.flowchart.internal_storage;',
    'flowChartDelay': 'shape=mxgraph.flowchart.delay;',
    'flowChartAlternateProcess': 'rounded=1;',
    'flowChartMultidocument': 'shape=mxgraph.flowchart.multi-document;',
    'flowChartDisplay': 'shape=mxgraph.flowchart.display;',
    # Pentagon / HomePlate
    'homePlate': 'shape=offPageConnector;',
    'pentagon': 'shape=offPageConnector;',
    # Callouts
    'wedgeRoundRectCallout': 'shape=callout;rounded=1;',
    'wedgeRectCallout': 'shape=callout;',
    'cloudCallout': 'shape=callout;rounded=1;',
    # Arrows
    'bentArrow': 'shape=mxgraph.arrows2.bent_arrow;',
    'chevron': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=20;notch=0;',
    'rightArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=east;',
    'leftArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=west;',
    'upArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=north;',
    'downArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;direction=south;',
    'leftRightArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;',
    'upDownArrow': 'shape=mxgraph.arrows2.arrow;dy=0.6;dx=40;',
    'notchedRightArrow': 'shape=mxgraph.arrows2.notched_arrow;',
    'stripedRightArrow': 'shape=mxgraph.arrows2.striped_arrow;',
}


FONT_ALIASES = {
    '\uff2d\uff33 \u30b4\u30b7\u30c3\u30af': 'MS PGothic',
    '\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af': 'MS PGothic',
    'MS Gothic': 'MS PGothic',
    'MS PGothic': 'MS PGothic',
    '\uff2d\uff33 \u660e\u671d': 'MS PMincho',
    '\uff2d\uff33 \uff30\u660e\u671d': 'MS PMincho',
    '\u6e38\u30b4\u30b7\u30c3\u30af': 'Yu Gothic',
    '\u6e38\u30b4\u30b7\u30c3\u30af Light': 'Yu Gothic Light',
    '\u6e38\u660e\u671d': 'Yu Mincho',
    '\u30e1\u30a4\u30ea\u30aa': 'Meiryo',
    'Meiryo': 'Meiryo',
}


BORDER_STYLE_MAP = {
    'thin': (1, None),
    'medium': (2, None),
    'thick': (3, None),
    'hair': (0.5, None),
    'dashed': (1, '8 8'),
    'mediumDashed': (2, '8 8'),
    'dotted': (1, '2 2'),
    'dashDot': (1, '8 4 2 4'),
    'mediumDashDot': (2, '8 4 2 4'),
    'dashDotDot': (1, '8 4 2 4 2 4'),
    'mediumDashDotDot': (2, '8 4 2 4 2 4'),
    'slantDashDot': (2, '8 4 2 4'),
    'double': (1, None),  # rendered as 2 lines in add_cell_borders
}

