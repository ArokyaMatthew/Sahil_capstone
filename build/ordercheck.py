"""Checks child-element order of common DrawingML containers against the OOXML schema sequence."""
import sys, zipfile, re
from lxml import etree
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
FILL = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']
SEQ = {
 'rPr': [['ln'], FILL, ['effectLst', 'effectDag'], ['highlight'], ['uLnTx', 'uLn'], ['uFillTx', 'uFill'], ['latin'], ['ea'], ['cs'], ['sym'], ['hlinkClick'], ['hlinkMouseOver'], ['rtl'], ['extLst']],
 'pPr': [['lnSpc'], ['spcBef'], ['spcAft'], ['buClrTx', 'buClr'], ['buSzTx', 'buSzPct', 'buSzPts'], ['buFontTx', 'buFont'], ['buNone', 'buAutoNum', 'buChar', 'buBlip'], ['tabLst'], ['defRPr'], ['extLst']],
 'bodyPr': [['prstTxWarp'], ['noAutofit', 'normAutofit', 'spAutoFit'], ['scene3d'], ['sp3d'], ['flatTx'], ['extLst']],
 'spPr': [['xfrm'], ['custGeom', 'prstGeom'], FILL, ['ln'], ['effectLst', 'effectDag'], ['scene3d'], ['sp3d'], ['extLst']],
 'tcPr': [['lnL'], ['lnR'], ['lnT'], ['lnB'], ['lnTlToBr'], ['lnBlToTr'], ['cell3D'], FILL + ['grpFill'], ['headers'], ['extLst']],
 'ln': [FILL, ['prstDash', 'custDash'], ['round', 'bevel', 'miter'], ['headEnd'], ['tailEnd'], ['extLst']],
}
alias = {'defRPr': 'rPr', 'endParaRPr': 'rPr', 'rPr': 'rPr', 'pPr': 'pPr', 'lvl1pPr': 'pPr', 'bodyPr': 'bodyPr', 'spPr': 'spPr', 'tcPr': 'tcPr', 'ln': 'ln', 'lnL': 'ln', 'lnR': 'ln', 'lnT': 'ln', 'lnB': 'ln'}
bad = 0
for f in sys.argv[1:]:
    z = zipfile.ZipFile(f)
    for n in z.namelist():
        if not re.match(r'ppt/(slides|slideLayouts|slideMasters)/[^/]+\.xml$', n):
            continue
        root = etree.fromstring(z.read(n))
        for el in root.iter():
            if not isinstance(el.tag, str):
                continue
            local = el.tag.split('}')[1]
            if local not in alias:
                continue
            seq = SEQ[alias[local]]
            pos = -1
            for ch in el:
                if not isinstance(ch.tag, str):
                    continue
                cl = ch.tag.split('}')[1]
                idx = next((i for i, g in enumerate(seq) if cl in g), None)
                if idx is None:
                    continue
                if idx < pos:
                    bad += 1
                    if bad < 15:
                        print(f, n, local, 'child', cl, 'out of order:', [c.tag.split('}')[1] for c in el])
                pos = max(pos, idx)
print('violations:', bad)
