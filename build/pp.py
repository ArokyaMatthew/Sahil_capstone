"""Shared python-pptx helpers for the Review-2 deck and the poster."""
import os
import re
from lxml import etree
from PIL import ImageFont
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- palette
G = '007367'        # GITAM / template structure green
G_DARK = '00564D'
FOOT_L, FOOT_M, FOOT_R = '198176', '66ABA4', '99C7C2'
BT_BG, BT_TX, BB_BG = 'CCE5CC', '008000', 'E6F2E6'   # beamer exampleblock
FADE = 'B7D7D3'
INK = '1F2D2B'
MUTED = '4A5A58'
BROWN = 'A57C52'    # GITAM accent (poster template)
WHITE = 'FFFFFF'

ALIGN = {'l': PP_ALIGN.LEFT, 'c': PP_ALIGN.CENTER, 'r': PP_ALIGN.RIGHT, 'j': PP_ALIGN.JUSTIFY}
ANCHOR = {'t': MSO_ANCHOR.TOP, 'm': MSO_ANCHOR.MIDDLE, 'b': MSO_ANCHOR.BOTTOM}


def rgb(h):
    return RGBColor.from_string(h)


# ---------------------------------------------------------------- text metrics
_WIN_FONTS = os.path.exists('C:/Windows/Fonts/cambria.ttc')
_FONT_FILES = {
    # Windows originals, or their metric-compatible Linux twins (Caladea = Cambria, Liberation Sans = Arial)
    ('Cambria', False): 'C:/Windows/Fonts/cambria.ttc' if _WIN_FONTS else '/usr/share/fonts/truetype/crosextra/Caladea-Regular.ttf',
    ('Cambria', True): 'C:/Windows/Fonts/cambriab.ttf' if _WIN_FONTS else '/usr/share/fonts/truetype/crosextra/Caladea-Bold.ttf',
    ('Arial', False): 'C:/Windows/Fonts/arial.ttf' if _WIN_FONTS else '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
    ('Arial', True): 'C:/Windows/Fonts/arialbd.ttf' if _WIN_FONTS else '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf',
}
_FONTS = {}


def _font(name, bold):
    key = (name, bold)
    if key not in _FONTS:
        _FONTS[key] = ImageFont.truetype(_FONT_FILES[key], 200)
    return _FONTS[key]


def text_w(s, size, name='Cambria', bold=False):
    """Rendered width of s in points."""
    return _font(name, bold).getlength(s) * size / 200.0


def n_lines(text, width_in, size, name='Cambria', bold=False, slack=1.04):
    avail = width_in * 72
    total = 0
    for para in strip_md(text).split('\n'):
        line, n = '', 1
        for w in para.split(' '):
            cand = (line + ' ' + w).strip()
            if text_w(cand, size, name, bold) * slack <= avail:
                line = cand
            else:
                n += 1 if line else 0
                line = w
        total += n
    return total


def strip_md(s):
    s = re.sub(r'\{\{c:[0-9A-Fa-f]{6}\|(.*?)\}\}', r'\1', s)
    return s.replace('**', '')


# ---------------------------------------------------------------- shape basics
def _drop_style(shape):
    st = shape._element.find(qn('p:style'))
    if st is not None:
        shape._element.remove(st)


def fill(shape, color):
    if color is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(color)


def line(shape, color=None, width=1.0, dash=None):
    if color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(color)
        shape.line.width = Pt(width)
        if dash:
            shape.line.dash_style = dash


def shadow(shape, alpha=35, dist=2.5, blur=4, angle=45):
    spPr = shape._element.spPr
    old = spPr.find(qn('a:effectLst'))
    if old is not None:
        spPr.remove(old)
    eff = etree.SubElement(spPr, qn('a:effectLst'))
    sh = etree.SubElement(eff, qn('a:outerShdw'), blurRad=str(int(blur * 12700)),
                          dist=str(int(dist * 12700)), dir=str(angle * 60000), algn='tl', rotWithShape='0')
    c = etree.SubElement(sh, qn('a:srgbClr'), val='000000')
    etree.SubElement(c, qn('a:alpha'), val=str(alpha * 1000))


def box(shapes, x, y, w, h, color=None, border=None, bw=1.0, kind=MSO_SHAPE.RECTANGLE,
        radius=None, name=None, dash=None, shade=False):
    s = shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    _drop_style(s)
    fill(s, color)
    line(s, border, bw, dash)
    if radius is not None and kind in (MSO_SHAPE.ROUNDED_RECTANGLE,):
        s.adjustments[0] = min(0.5, radius / min(w, h))
    if name:
        s.name = name
    if shade:
        shadow(s)
    return s


def top_round(shapes, x, y, w, h, color, radius=0.08, name=None):
    s = shapes.add_shape(MSO_SHAPE.ROUND_2_SAME_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    _drop_style(s)
    fill(s, color)
    line(s, None)
    s.adjustments[0] = min(0.5, radius / min(w, h))
    s.adjustments[1] = 0.0
    if name:
        s.name = name
    return s


# ---------------------------------------------------------------- paragraphs
def _runs(p, text, size, color, bold, italic, font):
    """Add runs to paragraph p; **x** toggles bold, {{c:RRGGBB|x}} colours a run."""
    tokens = re.split(r'(\*\*|\{\{c:[0-9A-Fa-f]{6}\|.*?\}\})', text)
    b = bold
    for tok in tokens:
        if tok == '':
            continue
        if tok == '**':
            b = not b
            continue
        col = color
        m = re.match(r'\{\{c:([0-9A-Fa-f]{6})\|(.*?)\}\}', tok)
        if m:
            col, tok = m.group(1), m.group(2)
        r = p.add_run()
        r.text = tok
        f = r.font
        f.size = Pt(size)
        f.name = font
        f.bold = b
        f.italic = italic
        f.color.rgb = rgb(col)


def bullet(p, char='\u25CF', color='0B5D0B', indent=0.24, pct=80, font='Arial'):
    pPr = p._p.get_or_add_pPr()
    pPr.set('marL', str(int(Inches(indent))))
    pPr.set('indent', str(-int(Inches(indent))))
    c = etree.SubElement(pPr, qn('a:buClr'))
    etree.SubElement(c, qn('a:srgbClr'), val=color)
    etree.SubElement(pPr, qn('a:buSzPct'), val=str(pct * 1000))
    etree.SubElement(pPr, qn('a:buFont'), typeface=font)
    etree.SubElement(pPr, qn('a:buChar'), char=char)


def no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    etree.SubElement(pPr, qn('a:buNone'))


def fill_tf(tf, paras, size=15, color=INK, bold=False, italic=False, align='l', font='Cambria',
            space_after=0, line_spacing=None, bullet_color='0B5D0B'):
    """paras: list of str or dict(text, bullet, size, color, bold, italic, align, space_after, space_before, indent)."""
    first = True
    for item in paras:
        if isinstance(item, str):
            item = {'text': item}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = ALIGN[item.get('align', align)]
        sa = item.get('space_after', space_after)
        if sa:
            p.space_after = Pt(sa)
        sb = item.get('space_before', 0)
        if sb:
            p.space_before = Pt(sb)
        ls = item.get('line_spacing', line_spacing)
        if ls:
            p.line_spacing = ls
        bl = item.get('bullet')
        if bl:
            bullet(p, char=bl if isinstance(bl, str) else '\u25CF',
                   color=item.get('bullet_color', bullet_color), indent=item.get('indent', 0.24))
        else:
            no_bullet(p)
            if item.get('indent'):
                p._p.get_or_add_pPr().set('marL', str(int(Inches(item['indent']))))
        _runs(p, item.get('text', ''), item.get('size', size), item.get('color', color),
              item.get('bold', bold), item.get('italic', italic), item.get('font', font))


def text(shapes, x, y, w, h, paras, size=15, color=INK, bold=False, italic=False, align='l', anchor='t',
         font='Cambria', margin=(0.04, 0.02, 0.04, 0.02), space_after=0, line_spacing=None, name=None,
         bullet_color='0B5D0B'):
    tb = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Inches(m) for m in margin]
    tf.vertical_anchor = ANCHOR[anchor]
    if isinstance(paras, str):
        paras = [paras]
    fill_tf(tf, paras, size, color, bold, italic, align, font, space_after, line_spacing, bullet_color)
    if name:
        tb.name = name
    return tb


def shape_text(shape, paras, size=12, color=INK, bold=False, align='c', anchor='m', font='Cambria',
               margin=(0.05, 0.03, 0.05, 0.03), italic=False, space_after=0):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Inches(m) for m in margin]
    tf.vertical_anchor = ANCHOR[anchor]
    if isinstance(paras, str):
        paras = [paras]
    fill_tf(tf, paras, size, color, bold, italic, align, font, space_after)


# ---------------------------------------------------------------- lines & arrows
def _arrowhead(ln, kind='triangle', end='tail', size='med'):
    tag = qn('a:tailEnd') if end == 'tail' else qn('a:headEnd')
    for old in ln.findall(tag):
        ln.remove(old)
    el = etree.SubElement(ln, tag, type=kind, w=size, len=size)
    # schema order inside a:ln: ... headEnd, tailEnd
    if end == 'head':
        tail = ln.find(qn('a:tailEnd'))
        if tail is not None:
            ln.remove(el)
            tail.addprevious(el)


def arrow(shapes, x1, y1, x2, y2, color=INK, width=1.5, head='triangle', start_head=None, dash=None, name=None):
    c = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    _drop_style(c)
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(width)
    if dash:
        c.line.dash_style = dash
    ln = c.line._get_or_add_ln()
    if start_head:
        _arrowhead(ln, start_head, 'head')
    if head:
        _arrowhead(ln, head, 'tail')
    if name:
        c.name = name
    return c


def polyline(shapes, pts, color=INK, width=1.5, head='triangle', dash=None, name=None):
    fb = shapes.build_freeform(int(Inches(pts[0][0])), int(Inches(pts[0][1])), scale=1.0)
    fb.add_line_segments([(int(Inches(x)), int(Inches(y))) for x, y in pts[1:]], close=False)
    s = fb.convert_to_shape()
    _drop_style(s)
    s.fill.background()
    s.line.color.rgb = rgb(color)
    s.line.width = Pt(width)
    if dash:
        s.line.dash_style = dash
    if head:
        _arrowhead(s.line._get_or_add_ln(), head, 'tail')
    if name:
        s.name = name
    return s


def icon(shapes, path, x, y, d):
    return shapes.add_picture(path, Inches(x), Inches(y), Inches(d), Inches(d))


def icon_disc(shapes, path, cx, cy, d, color=G, pad=0.22, border=None):
    c = box(shapes, cx - d / 2, cy - d / 2, d, d, color, border, 1.5, kind=MSO_SHAPE.OVAL)
    i = d * (1 - 2 * pad)
    icon(shapes, path, cx - i / 2, cy - i / 2, i)
    return c


# ---------------------------------------------------------------- table helpers
def cell_style(cell, fill_color=None, border='9FB8B5', bw=0.75, margin=(0.06, 0.03, 0.06, 0.03), anchor='m'):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
        for old in tcPr.findall(qn(tag)):
            tcPr.remove(old)
    els = []
    for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
        ln = etree.Element(qn(tag), w=str(int(Pt(bw))), cap='flat', cmpd='sng', algn='ctr')
        sf = etree.SubElement(ln, qn('a:solidFill'))
        etree.SubElement(sf, qn('a:srgbClr'), val=border)
        etree.SubElement(ln, qn('a:prstDash'), val='solid')
        els.append(ln)
    for i, el in enumerate(els):
        tcPr.insert(i, el)
    if fill_color:
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(fill_color)
    cell.margin_left, cell.margin_top, cell.margin_right, cell.margin_bottom = [Inches(m) for m in margin]
    cell.vertical_anchor = ANCHOR[anchor]


def cell_text(cell, paras, size=12, color=INK, bold=False, align='l', font='Cambria', italic=False):
    tf = cell.text_frame
    tf.word_wrap = True
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    for r in list(tf.paragraphs[0].runs):
        r._r.getparent().remove(r._r)
    if isinstance(paras, str):
        paras = [paras]
    fill_tf(tf, paras, size, color, bold, italic, align, font)


def plain_table(gf):
    tbl = gf.table._tbl
    tblPr = tbl.tblPr
    for a in ('firstRow', 'bandRow', 'firstCol', 'lastRow', 'lastCol', 'bandCol'):
        if tblPr.get(a) is not None:
            tblPr.set(a, '0')
    sid = tblPr.find(qn('a:tableStyleId'))
    if sid is not None:
        sid.text = '{5940675A-B579-460E-94D1-54222C63F5DA}'  # "No Style, Table Grid"
