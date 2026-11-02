"""Checks WordprocessingML child order for pPr, rPr, tblPr, trPr-less containers, tcPr, sectPr against the schema."""
import sys, zipfile
from collections import Counter
from lxml import etree
SEQ = {
 'pPr': 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
 'rPr': 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath',
 'tblPr': 'tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription',
 'tcPr': 'cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark',
 'sectPr': 'headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange',
}
SEQ = {k: v.split() for k, v in SEQ.items()}
bad = Counter()
ex = {}
for f in sys.argv[1:]:
    z = zipfile.ZipFile(f)
    for n in z.namelist():
        if not n.startswith('word/') or not n.endswith('.xml'):
            continue
        root = etree.fromstring(z.read(n))
        for el in root.iter():
            if not isinstance(el.tag, str) or '}' not in el.tag:
                continue
            local = el.tag.split('}')[1]
            if local not in SEQ:
                continue
            seq = SEQ[local]
            pos = -1
            for ch in el:
                if not isinstance(ch.tag, str):
                    continue
                cl = ch.tag.split('}')[1]
                if cl not in seq:
                    continue
                i = seq.index(cl)
                if i < pos:
                    key = (n, local, cl)
                    bad[key] += 1
                    ex.setdefault(key, [c.tag.split('}')[1] for c in el])
                pos = max(pos, i)
for k, v in bad.items():
    print(v, k, ex[k])
print('violations:', sum(bad.values()))
