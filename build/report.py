"""Fills the GITAM Capstone report template for the brain-tumor MRI project."""
import copy
import re
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

SP, OUT = sys.argv[1], sys.argv[2]
FIG = f'{SP}/figs'
d = Document(f'{SP}/inputs/Project Report_Capstone Introduction (1).docx')
TNR = 'Times New Roman'
TITLE = 'A Deep Learning Approach for Brain Tumor Classification Using MRI Images'
TEAM = [('B Chakardhar', '2023002241'), ('Sahil Purkayastha', '2023005003'),
        ('P. Nikhil Mohan', '2023003536'), ('G Yaswanth', '2023006345')]
GUIDE = 'Dr. P. Sreekanth'
DEPT = 'Department of Artificial Intelligence and Data Science'
DEGREE = 'B.Tech. in Computer Science and Engineering (Data Science)'

body = d.element.body
els = list(body.iterchildren())


# ===================================================================== helpers
def set_text(p_el, text, **fmt):
    """Keep the first run's formatting, replace all text of paragraph element p_el."""
    runs = p_el.findall(qn('w:r'))
    first = runs[0] if runs else None
    for r in runs[1:]:
        p_el.remove(r)
    if first is None:
        first = OxmlElement('w:r')
        p_el.append(first)
    for t in first.findall(qn('w:t')):
        first.remove(t)
    t = OxmlElement('w:t')
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    first.append(t)
    return first


def font(run, size=12, bold=None, italic=None, sub=False, sup=False):
    run.font.name = TNR
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), TNR)
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if sub:
        run.font.subscript = True
    if sup:
        run.font.superscript = True
    run.font.color.rgb = RGBColor(0, 0, 0)


# ---------------------------------------------------------------- citations
BIB = {
    'aamir22': 'M. Aamir *et al.*, “A deep learning approach for brain tumor classification using MRI images,” *Computers and Electrical Engineering*, vol. **101**, art. no. **108105**, (**2022**), doi: 10.1016/j.compeleceng.2022.108105.',
    'diaz21': 'F. J. Díaz-Pernas, M. Martínez-Zarzuela, M. Antón-Rodríguez, and D. González-Ortega, “A deep learning approach for brain tumor classification and segmentation using a multiscale convolutional neural network,” *Healthcare*, vol. **9**, no. 2, art. no. **153**, (**2021**), doi: 10.3390/healthcare9020153.',
    'aamir25': 'M. Aamir, Z. Rahman, U. A. Bhatti, W. A. Abro, J. A. Bhutto, and Z. He, “An automated deep learning framework for brain tumor classification using MRI imagery,” *Scientific Reports*, vol. **15**, art. no. **17593**, (**2025**), doi: 10.1038/s41598-025-02209-2.',
    'nick21': 'M. Nickparvar, “Brain tumor MRI dataset,” *Kaggle*, (**2021**), doi: 10.34740/KAGGLE/DSV/2645886.',
    'cheng15': 'J. Cheng *et al.*, “Enhanced performance of brain tumor classification via tumor region augmentation and partition,” *PLoS ONE*, vol. **10**, no. 10, art. no. **e0140381**, (**2015**), doi: 10.1371/journal.pone.0140381.',
    'sadr25': 'H. Sadr, M. Nazari, S. Yousefzadeh-Chabok, H. Emami, R. Rabiei, and A. Ashraf, “Enhancing brain tumor classification in MRI images: A deep learning-based approach for accurate diagnosis,” *Image and Vision Computing*, vol. **159**, art. no. **105555**, (**2025**), doi: 10.1016/j.imavis.2025.105555.',
    'vadde25': 'V. Vadde and W. Bukaita, “Comparative evaluation of CNN and ResNet18 architectures for MRI-based brain tumor classification using deep learning,” *Medical Research Archives*, vol. **13**, no. 12, (**2025**), doi: 10.18103/mra.v13i12.7100.',
    'rahman24': 'M. A. Rahman, P. Majumdar, A. Ghosh, and S. I. Tonni, “A proposed deep learning approach for brain tumor classification from MRI images,” in *Proc. 27th Int. Conf. Computer and Information Technology (ICCIT)*, Cox’s Bazar, Bangladesh, (**2024**), doi: 10.1109/ICCIT64611.2024.11022552.',
    'wong25': 'Y. Wong, E. L. M. Su, C. F. Yeong, W. Holderbaum, and C. Yang, “Brain tumor classification using MRI images and deep learning techniques,” *PLOS ONE*, vol. **20**, no. 5, art. no. **e0322624**, (**2025**), doi: 10.1371/journal.pone.0322624.',
    'omran25': 'H. M. Omran, K. Ibrahim, G. T. Abdel-Jaber, and A.-N. Sharkawy, “Brain tumor classification from MRI images using hybrid deep learning approaches: VGG19 with SoftMax and SVM classifiers,” *International Journal of Robotics and Control Systems*, vol. **6**, no. 1, pp. **16–35**, (**2025**), doi: 10.31763/ijrcs.v6i1.2260.',
    'zailan22': 'Z. N. I. Zailan, S. A. Mostafa, A. I. Abdulmaged, Z. Baharum, M. M. Jaber, and R. Hidayat, “Deep learning approach for prediction of brain tumor from small number of MRI images,” *JOIV: International Journal on Informatics Visualization*, vol. **6**, no. 2-2, pp. **581–586**, (**2022**), doi: 10.30630/joiv.6.2.987.',
    'nimma25': 'R. Nimmagadda and P. Kalpana Devi, “A deep learning approach for brain tumour classification and detection in MRI images using YOLOv7,” *Frontiers in Oncology*, vol. **15**, art. no. **1508326**, (**2025**), doi: 10.3389/fonc.2025.1508326.',
    'gomes26': 'E. F. Gomes and R. S. Barbosa, “Deep learning approaches for brain tumor classification in MRI scans: An analysis of model interpretability,” *Applied Sciences*, vol. **16**, no. 2, art. no. **831**, (**2026**), doi: 10.3390/app16020831.',
    'bouhafra25': 'S. Bouhafra and H. El Bahi, “Deep learning approaches for brain tumor detection and classification using MRI images (2020 to 2024): A systematic review,” *Journal of Imaging Informatics in Medicine*, vol. **38**, no. 3, pp. **1403–1433**, (**2025**), doi: 10.1007/s10278-024-01283-8.',
    'kaab25': 'A. Kaab, Z. Hachkar, and Y. Chihab, “Deep learning approaches to brain tumor detection: A review,” *Moroccan Journal of Health and Innovation*, vol. **1**, no. 2, (**2025**).',
    'deng09': 'J. Deng, W. Dong, R. Socher, L.-J. Li, K. Li, and L. Fei-Fei, “ImageNet: A large-scale hierarchical image database,” in *Proc. IEEE Conf. Computer Vision and Pattern Recognition (CVPR)*, pp. **248–255**, (**2009**), doi: 10.1109/CVPR.2009.5206848.',
    'tan19': 'M. Tan and Q. V. Le, “EfficientNet: Rethinking model scaling for convolutional neural networks,” in *Proc. 36th Int. Conf. Machine Learning (ICML)*, pp. **6105–6114**, (**2019**).',
    'he16': 'K. He, X. Zhang, S. Ren, and J. Sun, “Deep residual learning for image recognition,” in *Proc. IEEE Conf. Computer Vision and Pattern Recognition (CVPR)*, pp. **770–778**, (**2016**), doi: 10.1109/CVPR.2016.90.',
    'selvaraju17': 'R. R. Selvaraju, M. Cogswell, A. Das, R. Vedantam, D. Parikh, and D. Batra, “Grad-CAM: Visual explanations from deep networks via gradient-based localization,” in *Proc. IEEE Int. Conf. Computer Vision (ICCV)*, pp. **618–626**, (**2017**), doi: 10.1109/ICCV.2017.74.',
    'loshchilov19': 'I. Loshchilov and F. Hutter, “Decoupled weight decay regularization,” in *Proc. Int. Conf. Learning Representations (ICLR)*, (**2019**).',
    'loshchilov17': 'I. Loshchilov and F. Hutter, “SGDR: Stochastic gradient descent with warm restarts,” in *Proc. Int. Conf. Learning Representations (ICLR)*, (**2017**).',
    'szegedy16': 'C. Szegedy, V. Vanhoucke, S. Ioffe, J. Shlens, and Z. Wojna, “Rethinking the Inception architecture for computer vision,” in *Proc. IEEE Conf. Computer Vision and Pattern Recognition (CVPR)*, pp. **2818–2826**, (**2016**), doi: 10.1109/CVPR.2016.308.',
    'paszke19': 'A. Paszke *et al.*, “PyTorch: An imperative style, high-performance deep learning library,” in *Advances in Neural Information Processing Systems (NeurIPS)*, vol. **32**, pp. **8024–8035**, (**2019**).',
    'pedregosa11': 'F. Pedregosa *et al.*, “Scikit-learn: Machine learning in Python,” *Journal of Machine Learning Research*, vol. **12**, pp. **2825–2830**, (**2011**).',
    'buslaev20': 'A. Buslaev, V. I. Iglovikov, E. Khvedchenya, A. Parinov, M. Druzhinin, and A. A. Kalinin, “Albumentations: Fast and flexible image augmentations,” *Information*, vol. **11**, no. 2, art. no. **125**, (**2020**), doi: 10.3390/info11020125.',
}
ORDER = []


def cite(keys):
    nums = []
    for k in keys.split(','):
        k = k.strip()
        assert k in BIB, k
        if k not in ORDER:
            ORDER.append(k)
        nums.append(ORDER.index(k) + 1)
    nums = sorted(set(nums))
    parts, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        if j - i >= 2:
            parts.append(f'[{nums[i]}]–[{nums[j]}]')
        else:
            parts.extend(f'[{n}]' for n in nums[i:j + 1])
        i = j + 1
    return ', '.join(parts)


TOKEN = re.compile(r'(\*\*|\*|_\{[^}]*\}|\^\{[^}]*\}|\{[a-z0-9, ]+\})')


def add_runs(p, text, size=12, bold=False, italic=False):
    b, it = bold, italic
    for tok in TOKEN.split(text):
        if not tok:
            continue
        if tok == '**':
            b = not b
            continue
        if tok == '*':
            it = not it
            continue
        sub = sup = False
        if tok.startswith('_{'):
            tok, sub = tok[2:-1], True
        elif tok.startswith('^{'):
            tok, sup = tok[2:-1], True
        elif tok.startswith('{') and tok.endswith('}'):
            tok = cite(tok[1:-1])
        r = p.add_run(tok)
        font(r, size, b, it, sub, sup)


def fmt_para(p, align='j', line=1.5, before=0, after=12, keep_next=False, indent_left=None, first=None):
    pf = p.paragraph_format
    p.alignment = {'j': WD_ALIGN_PARAGRAPH.JUSTIFY, 'c': WD_ALIGN_PARAGRAPH.CENTER,
                   'l': WD_ALIGN_PARAGRAPH.LEFT, 'r': WD_ALIGN_PARAGRAPH.RIGHT}[align]
    pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.keep_with_next = keep_next
    if indent_left is not None:
        pf.left_indent = Cm(indent_left)
    if first is not None:
        pf.first_line_indent = Cm(first)


def P(text, align='j', size=12, after=12, line=1.5, bold=False, italic=False, keep_next=False):
    p = d.add_paragraph()
    fmt_para(p, align, line, 0, after, keep_next)
    add_runs(p, text, size, bold, italic)
    return p


def outline(p, level):
    pPr = p._p.get_or_add_pPr()
    o = OxmlElement('w:outlineLvl')
    o.set(qn('w:val'), str(level))
    pPr.append(o)


CH = [0]
SEC = [0]
SUB = [0]


def chapter(title):
    CH[0] += 1
    SEC[0] = 0
    p = d.add_paragraph()
    fmt_para(p, 'c', 1.5, 0, 12, True)
    p.paragraph_format.page_break_before = True
    add_runs(p, f'{CH[0]}. {title.upper()}', 14, True)
    outline(p, 0)
    return p


def section(title):
    SEC[0] += 1
    SUB[0] = 0
    p = d.add_paragraph()
    fmt_para(p, 'l', 1.5, 6, 6, True)
    add_runs(p, f'{CH[0]}.{SEC[0]} {title.upper()}', 12, True)
    outline(p, 1)


def subsection(title):
    SUB[0] += 1
    p = d.add_paragraph()
    fmt_para(p, 'l', 1.5, 4, 4, True)
    add_runs(p, f'{CH[0]}.{SEC[0]}.{SUB[0]} {title}', 12, True)
    outline(p, 2)


# bullets via a dedicated numbering definition
numbering = d.part.numbering_part.element
abs_id, num_id = '900', '900'
an = OxmlElement('w:abstractNum')
an.set(qn('w:abstractNumId'), abs_id)
lvl = OxmlElement('w:lvl')
lvl.set(qn('w:ilvl'), '0')
for tag, val in (('w:start', '1'), ('w:numFmt', 'bullet'), ('w:lvlText', '•'), ('w:lvlJc', 'left')):
    e = OxmlElement(tag)
    e.set(qn('w:val'), val)
    lvl.append(e)
ppr = OxmlElement('w:pPr')
ind = OxmlElement('w:ind')
ind.set(qn('w:left'), '567')
ind.set(qn('w:hanging'), '283')
ppr.append(ind)
lvl.append(ppr)
rpr = OxmlElement('w:rPr')
rf = OxmlElement('w:rFonts')
for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
    rf.set(qn(a), TNR)
rpr.append(rf)
lvl.append(rpr)
an.append(lvl)
first_num = numbering.find(qn('w:num'))
if first_num is not None:
    first_num.addprevious(an)
else:
    numbering.append(an)
num = OxmlElement('w:num')
num.set(qn('w:numId'), num_id)
ani = OxmlElement('w:abstractNumId')
ani.set(qn('w:val'), abs_id)
num.append(ani)
numbering.append(num)


def bullets(items, after_last=12):
    for k, it in enumerate(items):
        p = d.add_paragraph()
        fmt_para(p, 'j', 1.5, 0, after_last if k == len(items) - 1 else 2)
        pPr = p._p.get_or_add_pPr()
        numPr = OxmlElement('w:numPr')
        il = OxmlElement('w:ilvl')
        il.set(qn('w:val'), '0')
        ni = OxmlElement('w:numId')
        ni.set(qn('w:val'), num_id)
        numPr.append(il)
        numPr.append(ni)
        # schema order: numPr follows pStyle/keepNext/keepLines/pageBreakBefore/framePr/widowControl
        after = [c for c in pPr if c.tag in {qn('w:pStyle'), qn('w:keepNext'), qn('w:keepLines'), qn('w:pageBreakBefore'),
                                             qn('w:framePr'), qn('w:widowControl')}]
        after[-1].addnext(numPr) if after else pPr.insert(0, numPr)
        add_runs(p, it)


def equation(expr, label):
    p = d.add_paragraph()
    fmt_para(p, 'l', 1.5, 0, 8)
    ts = p.paragraph_format.tab_stops
    ts.add_tab_stop(Cm(8.5), WD_TAB_ALIGNMENT.CENTER)
    ts.add_tab_stop(Cm(17.0), WD_TAB_ALIGNMENT.RIGHT)
    add_runs(p, '\t' + expr + '\t' + f'({label})')


FIGN = {}
TABN = {}


def figure(path, caption, width_cm=16.5, key=None):
    n = sum(1 for k in FIGN.values() if k.startswith(f'{CH[0]}.')) + 1
    label = f'{CH[0]}.{n}'
    if key:
        FIGN[key] = label
    else:
        FIGN[f'_f{len(FIGN)}'] = label
    p = d.add_paragraph()
    fmt_para(p, 'c', 1.0, 6, 4, True)
    p.add_run().add_picture(path, width=Cm(width_cm))
    c = d.add_paragraph()
    fmt_para(c, 'c', 1.5, 0, 12)
    add_runs(c, f'Figure {label}: {caption}', 12, True)
    return label


def border_table(t, color='000000', sz='4'):
    tblPr = t._tbl.tblPr
    b = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), sz)
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), color)
        b.append(e)
    # schema order: tblBorders precedes shd/tblLayout/tblCellMar/tblLook
    nxt = [c for c in tblPr if c.tag in {qn('w:shd'), qn('w:tblLayout'), qn('w:tblCellMar'), qn('w:tblLook')}]
    nxt[0].addprevious(b) if nxt else tblPr.append(b)


def shade(cell, fillhex):
    tcPr = cell._tc.get_or_add_tcPr()
    s = OxmlElement('w:shd')
    s.set(qn('w:val'), 'clear')
    s.set(qn('w:color'), 'auto')
    s.set(qn('w:fill'), fillhex)
    tcPr.append(s)


def table(caption, header, rows, widths_cm, size=11, key=None, align=None):
    n = sum(1 for k in TABN.values() if k.startswith(f'{CH[0]}.')) + 1
    label = f'{CH[0]}.{n}'
    TABN[key or f'_t{len(TABN)}'] = label
    c = d.add_paragraph()
    fmt_para(c, 'c', 1.5, 6, 6, True)
    add_runs(c, f'Table {label}: {caption}', 12, True)
    t = d.add_table(rows=1 + len(rows), cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    border_table(t)
    for ci, w in enumerate(widths_cm):
        t.columns[ci].width = Cm(w)
    for ri, row in enumerate([header] + rows):
        tr = t.rows[ri]._tr
        if ri == 0:
            trPr = tr.get_or_add_trPr()
            h = OxmlElement('w:tblHeader')
            h.set(qn('w:val'), 'true')
            trPr.append(h)
        trPr_ = tr.get_or_add_trPr()
        trPr_.append(OxmlElement('w:cantSplit'))
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            cell.width = Cm(widths_cm[ci])
            if ri == 0:
                shade(cell, 'D9D9D9')
            para = cell.paragraphs[0]
            a = 'c' if ri == 0 else ((align or {}).get(ci, 'l'))
            fmt_para(para, a, 1.0, 1, 1)
            # short tables are kept on one page (spec 2.4: small tables sit with the text, never split)
            para.paragraph_format.keep_with_next = len(rows) <= 8 and ri < len(rows)
            add_runs(para, val, size, ri == 0)
    sp = d.add_paragraph()
    fmt_para(sp, 'l', 1.0, 0, 6)
    return label


# ===================================================================== front matter
P_TITLE = els[0]
set_text(P_TITLE, f'“{TITLE}”')
set_text(els[5], 'COMPUTER SCIENCE AND ENGINEERING')
set_text(els[8], '(Specialization in Data Science)')
for k in (6, 7):
    body.remove(els[k])
for i, (n, r) in enumerate(TEAM):
    set_text(els[12 + i], f'{n} ({r})')
set_text(els[19], GUIDE)
set_text(els[20], DEPT)
set_text(els[35], DEPT)

# declaration: its crest must open page 2 whatever the title-page length
Paragraph(els[40], d._body).paragraph_format.page_break_before = True
body.remove(els[39])

# declaration
set_text(els[48], f'We hereby declare that the project report entitled “{TITLE}” is an original work done in the {DEPT}, '
         'GITAM School of CSE, GITAM (Deemed to be University), submitted in partial fulfillment of the requirements for the '
         f'award of the degree of {DEGREE}. The work has not been submitted to any other college or University for the award '
         'of any degree or diploma.')
decl_tbl = d.add_table(rows=1 + len(TEAM), cols=3)
decl_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for ri, row in enumerate([('Registration No.', 'Name', 'Signature')] + [(r, n, '') for n, r in TEAM]):
    for ci, val in enumerate(row):
        cell = decl_tbl.cell(ri, ci)
        cell.width = Cm([5, 6, 5][ci])
        para = cell.paragraphs[0]
        fmt_para(para, 'l', 1.5, 6, 6)
        r_ = para.add_run(val)
        font(r_, 14, ri == 0)
els[53].addnext(decl_tbl._tbl)
body.remove(els[53])

# certificate
set_text(els[55], DEPT.upper())
names = ', '.join(f'{n} ({r})' for n, r in TEAM[:-1]) + f' and {TEAM[-1][0]} ({TEAM[-1][1]})'
set_text(els[67], f'This is to certify that the project report entitled “{TITLE}” is a bonafide record of work carried out by '
         f'{names}, students submitted in partial fulfillment of requirement for the award of degree of Bachelor of Technology '
         'in Computer Science and Engineering (Data Science).')
guide_line = copy.deepcopy(els[71])
set_text(guide_line, f'({GUIDE})')
els[71].addnext(guide_line)

# acknowledgement
ack = [
    f'We thank our project guide, {GUIDE}, for the guidance, regular reviews and candid feedback that shaped this work from problem '
    'identification to system design.',
    f'We thank the Head of the {DEPT} and the members of the Capstone Project review committee, whose comments at Review 1 led to '
    'sharper objectives, a traceable reference list and a fixed evaluation protocol.',
    'We are grateful to GITAM School of CSE, GITAM (Deemed to be University), Visakhapatnam, for the facilities and computing '
    'resources made available to us.',
    'We also thank the authors who released the public MRI datasets used in this work, and the developers of the open-source '
    'libraries on which the project depends.',
    'Finally, we thank our families and friends for their support throughout the project.',
]
anchor = els[73]
for txt in ack:
    p = d.add_paragraph()
    fmt_para(p, 'j', 1.5, 0, 12)
    add_runs(p, txt)
    anchor.addnext(p._p)
    anchor = p._p
sig = d.add_paragraph()
fmt_para(sig, 'r', 1.5, 24, 0)
for k, (n, r) in enumerate(TEAM):
    if k:
        sig.add_run().add_break()
    rr = sig.add_run(f'{n} ({r})')
    font(rr, 12, True)
anchor.addnext(sig._p)

# table of contents: drop empty rows, then write the measured page numbers (toc_pages.json, made by toc_pages.py)
toc = Table(els[77], d._body)  # the template's TOC table (d.tables[0] is now the declaration table)
for row in list(toc.rows):
    if not any(c.text.strip() for c in row.cells):
        row._tr.getparent().remove(row._tr)
import json, os  # noqa: E401,E402
_tp = f'{SP}/toc_pages.json'
if os.path.exists(_tp):
    pages = json.load(open(_tp))
    for row in list(toc.rows)[1:]:
        num = row.cells[0].text.strip().rstrip('.')
        if num in pages:
            cpar = row.cells[2].paragraphs[0]._p
            set_text(cpar, str(pages[num]))

# drop the discussion notes and documentation specifications (everything after the front-matter section break)
for el in els[79:-1]:
    body.remove(el)

# ===================================================================== page numbering
front, main = d.sections[0], d.sections[-1]


def page_field(par):
    r = par.add_run()
    fb = OxmlElement('w:fldChar')
    fb.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = ' PAGE '
    fe = OxmlElement('w:fldChar')
    fe.set(qn('w:fldCharType'), 'end')
    r._r.append(fb)
    r._r.append(it)
    r._r.append(fe)
    font(r, 12)


def set_footer(footer, with_number):
    for p in footer.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    p = footer.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if with_number:
        page_field(p)


def pg_fmt(sec, fmt=None, start=None):
    sp = sec._sectPr
    pn = sp.find(qn('w:pgNumType'))
    if pn is None:
        pn = OxmlElement('w:pgNumType')
        cols = sp.find(qn('w:cols'))
        cols.addprevious(pn)
    if fmt:
        pn.set(qn('w:fmt'), fmt)
    if start:
        pn.set(qn('w:start'), str(start))


front.different_first_page_header_footer = True
set_footer(front.footer, True)
set_footer(front.first_page_footer, False)
for h in (front.header, front.first_page_header):
    set_footer(h, False)
pg_fmt(front, 'lowerRoman', 1)

main.left_margin, main.right_margin = Cm(2.5), Cm(1.5)
main.top_margin, main.bottom_margin = Cm(1.5), Cm(1.5)
main.header_distance, main.footer_distance = Cm(0.8), Cm(0.7)
main.footer.is_linked_to_previous = False
main.header.is_linked_to_previous = False
set_footer(main.header, False)
set_footer(main.footer, True)
pg_fmt(main, 'decimal', 1)

# ===================================================================== 1. ABSTRACT
chapter('Abstract')
P('Deep-learning classifiers for brain-tumor MRI now report accuracies between 97 % and 99.7 %, but nearly every one of these '
  'figures is measured on the same public benchmark and summarised as a single overall accuracy. Such a figure says little about '
  'how a model behaves on images from another source, which tumor class it misses, or what it costs to run. This report presents '
  'the design of a four-class classifier — glioma, meningioma, pituitary tumor and no tumor — built to answer those three '
  'questions.')
P('The proposed system fine-tunes an ImageNet-pretrained EfficientNet-B0 and compares it with ResNet-50 and with a compact '
  'convolutional network trained from scratch. All three share one pipeline: perceptual-hash de-duplication, frozen and '
  'patient-level data splits, brain-region cropping, augmentation of the training split only, and tuning on validation data only. '
  'The models are evaluated on two public datasets — a 7,023-image four-class collection and the Figshare T1-weighted '
  'contrast-enhanced set of 3,064 slices from 233 patients — and are reported per class with precision, recall, F1-score, '
  'specificity, confusion matrices and ROC curves, together with parameter count and inference time. A Grad-CAM heatmap '
  'accompanies every prediction so that a reader can check where the model looked.')
P('The targets fixed before training are a macro-F1 of at least 0.95 on the held-out test set, a recall of at least 0.90 for every '
  'class, and fewer than 10 million parameters. This report covers the literature review, problem identification, requirement '
  'analysis, the proposed system, its architecture (workflow, data flow and use case diagrams) and the tools selected. '
  'Implementation and results follow in the next phase of the project.')
P('**Keywords:** brain tumor classification, magnetic resonance imaging, transfer learning, EfficientNet, Grad-CAM, '
  'cross-dataset evaluation.', align='l')

# ===================================================================== 2. INTRODUCTION
chapter('Introduction')
section('Background')
P('A brain tumor is an abnormal growth of cells inside the skull. Gliomas arise from the glial cells that support neurons and are '
  'often aggressive; meningiomas grow from the membranes that cover the brain and are usually slow-growing; pituitary tumors '
  'form in the pituitary gland at the base of the brain. Treatment, surgical planning and prognosis differ between these types, '
  'so naming the type correctly matters as much as detecting the tumor at all.')
P('Magnetic resonance imaging (MRI) is the standard imaging modality for the brain because it separates soft tissues well '
  'without ionising radiation. T1-weighted contrast-enhanced (T1-CE) sequences show tumor boundaries clearly and are widely used '
  'for this task. A radiologist reads a study slice by slice, which is slow, depends on scarce expertise and varies from one '
  'reader to another, while the number of studies to be read keeps growing.')
section('Deep Learning for MRI Classification')
P('Convolutional neural networks (CNNs) learn image features directly from pixels, so they do not need hand-designed texture or '
  'shape descriptors for each tumor type. Transfer learning reuses features learned on large natural-image collections, which '
  'matters here because labelled MRI data are limited. Published systems report accuracies of 98.95 % {aamir22}, 97.3 % '
  '{diaz21} and 99.67 % {aamir25} for brain-tumor classification. All three figures, however, are measured on the same '
  'Figshare T1-CE collection, and each is an aggregate over classes.')
section('Motivation and Justification of the Title')
P(f'The title, *{TITLE}*, names the three parts of the work: the method (deep learning, specifically transfer learning with '
  'CNNs), the task (assigning each image to one of four classes) and the input (two-dimensional MRI slices). What sets the '
  'project apart is how the approach is evaluated: on two public datasets {nick21, cheng15} under one reproducible protocol, '
  'per class rather than as one number, and with its computational cost and Grad-CAM explanations reported alongside its '
  'accuracy.')
section('Scope of the Project')
P('The following are within the scope of this project:', after=4)
bullets(['Slice-level, four-class classification of two-dimensional MRI images (glioma, meningioma, pituitary tumor, no tumor).',
         'Two public datasets, three candidate models and one shared training and evaluation pipeline.',
         'Per-class metrics, model cost (parameters and inference time) and Grad-CAM explanations.',
         'A browser-based research prototype that returns a class, its confidence and a heatmap for one uploaded slice.'])
P('Tumor segmentation, three-dimensional volumes, new data acquisition and any clinical deployment or diagnostic claim are '
  'outside the scope.')
section('Organisation of the Report')
P('Chapter 3 reviews thirteen publications on deep learning for brain-tumor MRI. Chapter 4 derives the research gaps, the problem '
  'statement and the objectives. Chapter 5 describes the existing approach and its drawbacks, and Chapter 6 the proposed system, '
  'including the requirement analysis. Chapter 7 presents the system architecture through block, data flow and use case '
  'diagrams. Chapter 8 lists the tools and technologies, Chapter 9 concludes with expected outcomes and the plan of work, and '
  'Chapter 10 lists the references.')

# ===================================================================== 3. LITERATURE REVIEW
chapter('Literature Review')
section('Review Method')
P('Thirteen publications were reviewed: eleven primary studies and two review articles, published between 2021 and 2026. A paper '
  'was included if it was a peer-reviewed journal article or IEEE conference paper on deep learning for brain-tumor '
  'classification or detection from MRI. Bibliographic details were taken from the publisher record. The studies are grouped '
  'below by what they contribute, and Table 3.1 summarises all thirteen.')
section('Studies on the Figshare Benchmark')
P('Aamir *et al.* {aamir22} enhance MRI contrast with an illumination boost and non-linear stretching, extract features with '
  'EfficientNet and ResNet-50, fuse them using partial least squares, generate tumor proposals by agglomerative clustering and '
  'classify them after ROI pooling to a fixed 7 × 7 vector. They report 98.95 % accuracy on the Figshare T1-CE set.')
P('Díaz-Pernas *et al.* {diaz21} use a three-pathway multiscale CNN (11 × 11, 7 × 7 and 3 × 3 kernels) on '
  '65 × 65 sliding windows with elastic-transform augmentation, and perform classification and segmentation together on the '
  'same Figshare set of 3,064 slices from 233 patients. Classification accuracy is 0.973, but the glioma Dice score is only '
  '0.779, and the authors record false positives in non-cerebral regions such as the sphenoidal sinuses.')
P('Aamir *et al.* {aamir25} build an end-to-end framework with adaptive guided filtering, U-Net segmentation, attention-based '
  'multiscale features (SE-ResNet101 and GoogLeNet) and an ensemble of random forest, SVM and k-NN classifiers. It reaches '
  '99.67 % on Figshare, the highest figure in the reviewed set, at the price of heavy pre-processing and an ensemble that the '
  'authors describe as computationally expensive and hard to interpret.')
P('Sadr *et al.* {sadr25} compare a 26-layer CNN with fine-tuned VGG16 and Xception models using global average pooling. They '
  'report 97.27 % on Figshare, 98.57 % on Br35H (Xception) and 100 % on the small Harvard Medical set, which shows how strongly '
  'the reported accuracy depends on the dataset chosen.')
section('Transfer Learning and Model Comparisons')
P('Vadde and Bukaita {vadde25} compare a custom CNN with an ImageNet-pretrained ResNet18 and find that the pretrained network '
  'performs better, at around 99.5 % accuracy, on a single-source 2D dataset. Rahman *et al.* {rahman24} benchmark DenseNet201, '
  'VGG19, MobileNetV2, ResNet50 and a custom CNN on a merged multi-class corpus, but publish few comparable performance and '
  'validation details. Wong *et al.* {wong25} fine-tune VGG16 with additional layers on a merged and balanced four-class '
  'dataset and add a graphical interface for prediction. Omran *et al.* {omran25} classify VGG19 features with an RBF-SVM and with '
  'a softmax layer and analyse misclassifications per class; parts of their experimental setup could not be verified from the '
  'publication. Zailan *et al.* {zailan22} train VGG16, InceptionV3 and MobileNetV2 on only 253 images and reach 86.00 %, far '
  'below the figures reported on Figshare.')
section('Detection and Explainability')
P('Nimmagadda and Kalpana Devi {nimma25} adapt YOLOv7 with CBAM attention and an SPPF+ block to detect and classify tumors in one '
  'pass, and use Grad-CAM to inspect the detections. Their data contain 2,870 images with 217 labelled boxes; they report a '
  'false-positive rate of 0.35 and the weakest precision–recall curve for the no-tumor class. Gomes and Barbosa {gomes26} '
  'compare VGG19, ResNet50, EfficientNetB3, Xception, MobileNetV2, DenseNet201, InceptionV3, a Vision Transformer and an '
  'ensemble on a 7,023-image dataset. Accuracies range from 89.47 % to 98.17 % (Vision Transformer). Their Grad-CAM, LIME and '
  'occlusion-sensitivity analysis shows that some models attend to regions outside the tumor even when the prediction is correct.')
section('Review Articles')
P('Bouhafra and El Bahi {bouhafra25} review 60 studies published between 2020 and January 2024 and group them into transfer '
  'learning, custom CNNs, autoencoders, transformers, attention mechanisms and GAN-based methods. They identify small, imbalanced '
  'and heterogeneous datasets as the main limitation and link acquisition artefacts directly to false positives. Kaab *et al.* '
  '{kaab25} review CNN-based detection and classification from 2020 to 2024 and conclude that no single technique has been '
  'established for clinical use, partly because studies differ too much in data and protocol to be compared.')
section('Summary of the Literature')
lit_rows = [
    (cite('aamir22'), 'EfficientNet + ResNet-50 features, PLS fusion, tumor proposals', 'Figshare T1-CE', '98.95 % accuracy', 'Single dataset; complex pipeline'),
    (cite('diaz21'), 'Three-pathway multiscale CNN, classification + segmentation', 'Figshare (3,064 slices, 233 patients)', '0.973 accuracy; glioma Dice 0.779', 'False positives outside the brain'),
    (cite('aamir25'), 'U-Net, SE-ResNet101 + GoogLeNet features, RF/SVM/k-NN ensemble', 'Figshare T1-CE', '99.67 % accuracy', 'Costly; hard to interpret'),
    (cite('sadr25'), '26-layer CNN; fine-tuned VGG16, Xception', 'Figshare, Br35H, Harvard', '97.27 %, 98.57 %, 100 %', 'Results depend on dataset size'),
    (cite('vadde25'), 'Custom CNN vs. pretrained ResNet18', 'Single-source 2D set', 'ResNet18 ≈ 99.5 %', 'One source; limited interpretability'),
    (cite('rahman24'), 'DenseNet201, VGG19, MobileNetV2, ResNet50, custom CNN', 'Merged multi-class corpus', 'Not comparably reported', 'Sparse validation detail'),
    (cite('wong25'), 'Fine-tuned VGG16 with extra layers, GUI', 'Merged, balanced 4-class set', 'High accuracy (4 classes)', 'Wider validation needed'),
    (cite('omran25'), 'VGG19 features + RBF-SVM / softmax', 'Public MRI set', 'Class-wise analysis', 'Setup not fully verifiable'),
    (cite('zailan22'), 'VGG16, InceptionV3, MobileNetV2', '253 images', '86.00 % accuracy', 'Very small dataset'),
    (cite('nimma25'), 'YOLOv7 + CBAM + SPPF+, Grad-CAM', '2,870 images, 217 boxes', 'False-positive rate 0.35', 'Weak on no-tumor class'),
    (cite('gomes26'), 'Eight CNNs, ViT, ensemble; Grad-CAM, LIME', '7,023 images, 4 classes', '89.47–98.17 %', 'Some models attend off-tumor'),
    (cite('bouhafra25'), 'Systematic review of 60 studies', '2020–2024 literature', 'Taxonomy of methods', 'Data heterogeneity unresolved'),
    (cite('kaab25'), 'Review of CNN-based detection', '2020–2024 literature', 'Survey of algorithms', 'No clinically established method'),
]
table('Summary of the reviewed literature', ['Ref.', 'Approach', 'Data', 'Reported result', 'Limitation'],
      [list(r) for r in lit_rows], [1.2, 5.0, 3.6, 3.3, 3.6], size=10, align={0: 'c'})

# ===================================================================== 4. PROBLEM IDENTIFICATION & OBJECTIVES
chapter('Problem Identification & Objectives')
section('Observations from the Literature')
P('Three observations stand out from Chapter 3. First, the highest accuracies — 98.95 % {aamir22}, 0.973 {diaz21} and '
  '99.67 % {aamir25} — are all measured on the same Figshare T1-CE collection, so the gains reported over one another are gains '
  'on the same images, not evidence that the models generalise. Second, overall accuracy is reported even where per-class '
  'behaviour is visibly uneven. Third, the studies differ in dataset version, number of classes, pre-processing, augmentation and '
  'metrics, which makes their numbers hard to compare {bouhafra25, kaab25}.')
section('Research Gaps')
P('**Gap 1 — No evidence of cross-dataset generalisation.** Near-ceiling accuracy on one benchmark does not transfer: on a '
  'different and much smaller dataset accuracy falls to 86.00 % {zailan22}, and other studies use a single-source 2D set '
  '{vadde25} or only 2,870 images {nimma25}. How a model performs on a second dataset is still unknown.')
P('**Gap 2 — Class-dependent failure is hidden by overall accuracy.** A glioma Dice score of 0.779 and false positives in the '
  'sphenoidal sinuses {diaz21}, a false-positive rate of 0.35 with the weakest precision–recall curve for the no-tumor class '
  '{nimma25}, and false positives linked to acquisition artefacts {bouhafra25} are all invisible in a single accuracy figure.')
P('**Gap 3 — No standard protocol, and cost and interpretability are ignored.** The most accurate approach relies on an '
  'ensemble described as expensive and hard to interpret {aamir25}, and explainability studies show models attending to '
  'irrelevant regions {gomes26}. Training cost, inference time and explanations are rarely reported with the accuracy.')
section('Problem Statement')
P('Deep-learning models for MRI brain-tumor classification report very high accuracy, but those results come almost entirely from '
  'a single public benchmark and are summarised by one overall figure that conceals which class fails. For four-class '
  'classification — glioma, meningioma, pituitary tumor and no tumor — it is not known how well such models hold up on a '
  'second dataset, how much accuracy is lost on the hardest class, or what the approach costs in computation and '
  'interpretability. This project develops a transfer-learning classifier and evaluates it on more than one public dataset under '
  'a documented, reproducible protocol, reporting results per class together with cost and visual explanations.')
section('Objectives')
P('Two objectives were fixed, each with a measurable target:', after=4)
bullets(['**O1:** To build a transfer-learning classifier based on EfficientNet-B0, benchmarked against ResNet-50 and a CNN trained '
         'from scratch, that reaches a macro-F1 of at least 0.95 on the untouched test set of the 7,023-image MRI dataset with '
         'fewer than 10 million parameters.',
         '**O2:** To evaluate the classifier under one documented protocol on a second public dataset (Figshare T1-CE, 3,064 slices, '
         '233 patients, patient-level splits) and to report per-class precision, recall, F1-score, specificity, confusion matrix '
         'and Grad-CAM maps, with a recall of at least 0.90 for every class.'])
table('Objectives, targets and the gaps they address', ['Objective', 'Measurable target', 'Gap addressed'],
      [['O1 — build and benchmark the classifier', 'Macro-F1 ≥ 0.95 on the D1 test set; < 10 M parameters; '
        '≤ 50 ms per image on a T4 GPU', 'Gap 3 (cost)'],
       ['O2 — evaluate on a second dataset, per class', 'Recall ≥ 0.90 for every class; per-class metrics, confusion '
        'matrix and Grad-CAM on D1 and D2', 'Gap 1, Gap 2']],
      [5.0, 8.2, 3.5], size=11)

# ===================================================================== 5. EXISTING SYSTEM
chapter('Existing System')
section('Manual Radiological Reading')
P('In current practice a radiologist reviews each MRI study slice by slice, identifies a suspected lesion and judges its type from '
  'its location, shape, enhancement pattern and the patient’s history. The process is accurate in expert hands but has '
  'well-known drawbacks:', after=4)
bullets(['It is slow: a single study contains many slices across several sequences.',
         'It depends on trained neuroradiologists, who are scarce outside large hospitals.',
         'Readings vary between observers, especially for small or atypical lesions.',
         'The backlog grows with imaging volume, while the number of readers does not.'])
section('Existing Deep-Learning Systems')
P('The automated systems in the literature follow a common pattern. A public dataset — most often Figshare T1-CE '
  '{cheng15} — is split at random into training and test images, a pretrained CNN or an ensemble is fine-tuned, and a single '
  'overall accuracy is reported {aamir22, diaz21, aamir25}. Some studies add segmentation {diaz21}, detection {nimma25} or '
  'explanation maps {gomes26}.')
section('Drawbacks of the Existing System')
bullets(['**Single-benchmark evaluation.** Almost all headline results come from one dataset, so their behaviour on images from '
         'another source is unknown (Gap 1).',
         '**Possible leakage between splits.** When slices are split at random rather than by patient, slices of one patient can '
         'appear in both the training and test sets, which inflates measured accuracy. Merged collections can also contain the '
         'same or near-identical image more than once.',
         '**Aggregate metrics only.** One accuracy figure hides the class that fails and the false positives it produces (Gap 2).',
         '**Cost and interpretability left out.** Ensembles and heavy pre-processing raise accuracy slightly but make the model '
         'slow and opaque, and these costs are seldom reported (Gap 3).',
         '**No common protocol.** Differences in classes, splits, augmentation and metrics make published numbers hard to compare.'])

# ===================================================================== 6. PROPOSED SYSTEM
chapter('Proposed System')
section('Overview')
P('The proposed system is a four-class MRI classifier built and evaluated under one reproducible protocol. Its implementation is '
  'organised in five phases, shown in Figure 6.1: data and environment, pipeline and baselines, primary model and tuning, '
  'evaluation and explainability, and prototype and documentation. A quality gate after the evaluation phase sends the work '
  'back to tuning if either target is missed, so the project moves on only when both targets hold.')
figure(f'{FIG}/fig_workflow.png', 'Proposed workflow: implementation phases and quality gate')
section('Requirement Analysis')
subsection('Users and Their Needs')
bullets(['**Clinician or radiologist (end user):** a quick second opinion on a single slice, with a visible reason in the form of a '
         'heatmap rather than a bare label.',
         '**Research team (developers):** one configuration-driven pipeline to retrain, compare models and reproduce every '
         'reported number.',
         '**Guide and reviewers (evaluators):** per-class and second-dataset results that can be checked against published work.'])
subsection('Functional Requirements')
table('Functional requirements', ['ID', 'Requirement', 'Description'], [
    ['FR1', 'Data ingestion and curation', 'Load both datasets with labels, check file integrity and remove exact and near-duplicate images.'],
    ['FR2', 'Pre-processing', 'Crop to the brain region, resize to 224 × 224, normalise intensities; augment the training split only.'],
    ['FR3', 'Model training', 'Train EfficientNet-B0, ResNet-50 and a baseline CNN from one configuration file with fixed random seeds; save checkpoints.'],
    ['FR4', 'Evaluation', 'Compute per-class precision, recall, F1-score, specificity, confusion matrix and ROC-AUC on both datasets.'],
    ['FR5', 'Prediction and explanation', 'Accept one MRI slice and return the predicted class, its confidence and a Grad-CAM heatmap.'],
], [1.3, 4.2, 11.2], size=11, align={0: 'c'})
subsection('Non-Functional Requirements')
table('Non-functional requirements', ['ID', 'Attribute', 'Measurable criterion'], [
    ['NFR1', 'Accuracy', 'Macro-F1 ≥ 0.95 on the held-out test set; recall ≥ 0.90 for every class.'],
    ['NFR2', 'Reproducibility', 'Fixed seeds, frozen split files and logged configurations; a rerun reproduces macro-F1 within ±0.5 points.'],
    ['NFR3', 'Efficiency', 'Fewer than 10 M parameters; ≤ 50 ms per image on an NVIDIA T4 GPU and ≤ 1 s on a laptop CPU.'],
    ['NFR4', 'Interpretability', 'Every prediction is accompanied by a heatmap showing the region that drove it.'],
    ['NFR5', 'Ethics and privacy', 'Only public, de-identified data are used; every output is labelled “research use, not a diagnosis”.'],
], [1.5, 3.5, 11.7], size=11, align={0: 'c'})
section('Datasets')
P('Two public datasets are used (Table 6.3). D1 is the primary dataset for training and testing all models. D2 provides patient '
  'identifiers, which makes patient-level splitting possible, and serves as the second dataset required by Objective O2.')
table('Datasets used in the project', ['', 'D1: Brain Tumor MRI Dataset ' + cite('nick21'), 'D2: Figshare T1-CE ' + cite('cheng15')], [
    ['Images', '7,023', '3,064 slices from 233 patients'],
    ['Classes', 'Glioma 1,621; meningioma 1,645; pituitary 1,757; no tumor 2,000', 'Meningioma 708; glioma 1,426; pituitary 930'],
    ['Format', 'JPEG, varying sizes', '512 × 512 T1-CE slices'],
    ['Split used', 'Official test set (1,311) frozen; training set (5,712) split 85 / 15 into training and validation, stratified',
     'Five-fold, patient-level cross-validation'],
], [2.6, 7.2, 6.9], size=11)
section('Data Curation and Splitting')
P('Before any split is made, every file is opened and checked, and labels are audited against folder names. Duplicate and '
  'near-duplicate images are found with a perceptual hash: two images whose hashes differ by only a few bits are treated as the '
  'same image, and only one copy is kept. The check runs within each dataset and across D1 and D2, because merged collections can '
  'contain images from the same original sources. The resulting split files are frozen and versioned, so every experiment uses '
  'exactly the same images. On D1 the official test set is never used for training or tuning. On D2 all slices of a patient fall '
  'in the same fold.')
section('Pre-processing and Augmentation')
P('Each image is converted to grayscale, thresholded, and cropped to the bounding box of the largest contour, which removes empty '
  'background around the skull. The crop is resized to 224 × 224 pixels, replicated to three channels and normalised with the '
  'ImageNet mean and standard deviation expected by the pretrained backbones. During training only, Albumentations {buslaev20} '
  'applies horizontal flips, rotations of up to ±15°, brightness and contrast changes, and random resized crops covering 85 '
  'to 100 % of the image. Validation and test images are only cropped, resized and normalised.')
section('Models')
P('Three models are trained with the same pipeline so that differences in the results come from the model alone (Table 6.4).')
table('Candidate models', ['Model', 'Parameters', 'GMACs (224 × 224)', 'Initialisation', 'Role'], [
    ['EfficientNet-B0 ' + cite('tan19'), '5.3 M', '0.39', 'ImageNet ' + cite('deng09'), 'Primary model'],
    ['ResNet-50 ' + cite('he16'), '25.6 M', '4.1', 'ImageNet', 'Transfer-learning baseline'],
    ['Custom CNN', '≈ 0.4 M', '≈ 0.7', 'Random', 'Ablation: value of pre-training'],
], [4.2, 2.3, 3.0, 3.2, 4.0], size=11, align={1: 'c', 2: 'c', 3: 'c'})
P('**EfficientNet-B0** {tan19} is built from mobile inverted-bottleneck blocks with squeeze-and-excitation and is scaled jointly in '
  'depth, width and resolution, which gives it the best accuracy per operation of the three; with 5.3 M parameters it also meets '
  'NFR3. **ResNet-50** {he16} uses residual connections that keep gradients stable in deep networks and is the backbone on which '
  'many of the reviewed studies build, so it serves as the reference. The **custom CNN** has four convolutional blocks with '
  '32, 64, 128 and 256 filters (3 × 3 convolution, batch normalisation, ReLU and max pooling), followed by global average '
  'pooling, dropout of 0.3 and a four-way output layer. It is trained from random weights to measure how much ImageNet '
  'pre-training adds on scarce MRI data, the question studied in {vadde25}.')
section('Training Procedure')
P('The pretrained models are fine-tuned in two stages. In the first stage, only the new classification head is trained for five '
  'epochs while the backbone is frozen. In the second stage, the whole network is trained at a learning rate ten times lower. '
  'Optimisation uses AdamW {loshchilov19} with a weight decay of 1 × 10^{-4}, a batch size of 32 and mixed-precision '
  'arithmetic. The learning rate follows a cosine schedule {loshchilov17} after a one-epoch warm-up. The loss is class-weighted '
  'cross-entropy with label smoothing of 0.1 {szegedy16}, so that the smallest class still carries weight. Training stops when '
  'the validation macro-F1 has not improved for seven epochs, and the checkpoint with the best validation macro-F1 is kept. All '
  'model selection and tuning use validation data only; the test split is opened once, for the final numbers.')
section('Evaluation Protocol')
P('For every class *c*, with TP_{c}, FP_{c}, FN_{c} and TN_{c} counted one-versus-rest from the confusion matrix, the following '
  'metrics are computed:', after=6)
equation('Precision_{c} = TP_{c} / (TP_{c} + FP_{c})', '6.1')
equation('Recall_{c} = TP_{c} / (TP_{c} + FN_{c})', '6.2')
equation('F1_{c} = 2 · Precision_{c} · Recall_{c} / (Precision_{c} + Recall_{c})', '6.3')
equation('Specificity_{c} = TN_{c} / (TN_{c} + FP_{c})', '6.4')
equation('Macro-F1 = (1 / K) Σ_{c} F1_{c},  K = 4', '6.5')
P('Accuracy, the confusion matrix and one-versus-rest ROC-AUC are reported alongside these metrics using scikit-learn '
  '{pedregosa11}. On D1, the frozen test set is evaluated once. On D2, results are given as the mean and standard deviation over '
  'the five patient-level folds. In a further cross-dataset check, the model trained on D1 is applied to D2, but only after D1 '
  'images that duplicate D2 images have been removed. Cost is reported as parameter count, training time and inference time '
  'per image. Grad-CAM {selvaraju17} heatmaps from the last convolutional block are reviewed for correctly and incorrectly '
  'classified images of every class.')
section('Improvements over the Existing System')
table('Existing system compared with the proposed system', ['Aspect', 'Existing system', 'Proposed system'], [
    ['Datasets', 'One public benchmark', 'Two public datasets under one protocol'],
    ['Data splits', 'Random, slice-level', 'Frozen splits; patient-level on D2; duplicates removed'],
    ['Metrics', 'Overall accuracy', 'Per-class precision, recall, F1, specificity, confusion matrix, ROC-AUC'],
    ['Cost', 'Rarely reported', 'Parameters, training time and inference time per image'],
    ['Explanation', 'Optional or absent', 'Grad-CAM heatmap for every prediction'],
    ['Reproducibility', 'Partial', 'Fixed seeds, versioned splits and logged configurations'],
], [3.3, 5.6, 7.8], size=11)

# ===================================================================== 7. SYSTEM ARCHITECTURE
chapter('System Architecture')
section('Block Diagram')
P('Figure 7.1 shows the complete path from input to output. Each block produces a defined artefact that the next block consumes, '
  'which makes every stage testable on its own.')
figure(f'{FIG}/fig_method.png', 'System architecture: block diagram from MRI input to explained prediction')
bullets(['**Input:** labelled slices from D1 and D2 for training, or one uploaded slice for inference.',
         '**Curation:** integrity checks, perceptual-hash de-duplication and frozen, patient-level splits.',
         '**Pre-processing:** brain-region crop, resizing, normalisation and training-only augmentation, producing 224 × 224 '
         'tensors.',
         '**Training:** the three models with the shared recipe of Section 6.7, producing one checkpoint per model.',
         '**Evaluation:** per-class metrics, confusion matrices, ROC curves, cost figures and Grad-CAM review. Tuning returns to '
         'the training block using validation data only.',
         '**Output:** the predicted class, its confidence and a heatmap overlay, together with the evaluation report.'])
section('Data Flow Diagrams')
subsection('Level-0 DFD')
P('The context diagram (Figure 7.2) treats the system as a single process with three external entities. The public datasets '
  'supply labelled images; the research team supplies configurations and receives the evaluation report; the end user sends one '
  'MRI slice and receives the class, its confidence and a heatmap.')
figure(f'{FIG}/fig_dfd0.png', 'Level-0 DFD (context diagram) in Gane–Sarson notation', width_cm=15.5)
subsection('Level-1 DFD')
P('The Level-1 DFD (Figure 7.3) opens the single process into five. Process 1.0 ingests and curates the raw images into data '
  'store D1. Process 2.0 pre-processes and splits them into store D2. Process 3.0 trains the models with the team’s '
  'configuration and writes the best checkpoints to the model registry D3. Process 4.0 reads the frozen test set from D2 and the '
  'trained model from D3 and writes metrics and heatmaps to D4, from which the team receives the evaluation report. Process 5.0 '
  'reads the trained model from D3 to answer the end user. The flows to and from the external entities match those of the '
  'Level-0 diagram, so the two levels are balanced.')
figure(f'{FIG}/fig_dfd1.png', 'Level-1 DFD of the proposed system')
section('Use Case Diagram')
P('Figure 7.4 shows the three actors and their use cases. The end user uploads a slice, views the class and confidence, may open '
  'the Grad-CAM heatmap (an extension of viewing the result) and can download a report. The research team curates and splits '
  'the datasets, trains and tunes the models, and evaluates them per class and across datasets. The guide and reviewers read '
  'the evaluation report. Both uploading a slice and curating a dataset include the same validation and pre-processing step, so '
  'training and inference share one routine.')
figure(f'{FIG}/fig_usecase.png', 'Use case diagram (UML) of the proposed system', width_cm=15.5)
section('Module Description')
table('Software modules of the proposed system', ['Module', 'Responsibility'], [
    ['data/', 'Dataset download scripts, integrity checks, perceptual-hash de-duplication and split files'],
    ['preprocess/', 'Brain-region cropping, resizing, normalisation and augmentation pipelines'],
    ['models/', 'EfficientNet-B0, ResNet-50 and custom CNN definitions with a common interface'],
    ['train.py', 'Two-stage training loop, logging, checkpointing and early stopping'],
    ['evaluate.py', 'Per-class metrics, confusion matrices, ROC curves, cost measurement'],
    ['explain.py', 'Grad-CAM heatmaps and overlays'],
    ['app.py', 'Streamlit prototype: upload a slice, view class, confidence and heatmap'],
    ['configs/', 'One YAML file per experiment: data, model, optimiser, schedule and seed'],
], [3.5, 13.2], size=11)

# ===================================================================== 8. TOOLS
chapter('Tools/Technologies Used')
section('Software Tools')
table('Software tools and their role in the project', ['Tool', 'Purpose', 'Why it was chosen'], [
    ['Python 3.11', 'Implementation language', 'Standard for deep learning; all required libraries are available'],
    ['PyTorch 2.x and torchvision ' + cite('paszke19'), 'Model definition, training, inference', 'ImageNet-pretrained backbones; mixed-precision GPU training'],
    ['scikit-learn ' + cite('pedregosa11'), 'Metrics and cross-validation', 'Tested implementations of per-class metrics, ROC-AUC and grouped folds'],
    ['OpenCV', 'Image reading and brain-region cropping', 'Fast thresholding and contour detection'],
    ['Albumentations ' + cite('buslaev20'), 'Training-time augmentation', 'Fast, composable image transforms'],
    ['imagehash', 'Perceptual hashing', 'Detects near-duplicate images across datasets'],
    ['pytorch-grad-cam', 'Grad-CAM heatmaps', 'Works with both CNN backbones without code changes'],
    ['NumPy, pandas, Matplotlib', 'Data handling and plots', 'Tables of results, confusion matrices and curves'],
    ['Streamlit', 'Research prototype', 'Browser interface in a few lines of Python'],
    ['Git, GitHub, TensorBoard', 'Version control and experiment tracking', 'Every result traceable to code, configuration and curves'],
], [4.6, 5.0, 7.1], size=11)
section('Technologies')
P('**Transfer learning.** Networks pretrained on ImageNet {deng09} already encode edges, textures and shapes. Fine-tuning them '
  'on MRI needs far fewer labelled images than training from scratch, which suits the dataset sizes available here.')
P('**Explainable AI.** Grad-CAM {selvaraju17} weights the feature maps of the last convolutional layer by the gradient of the '
  'predicted class and produces a coarse heatmap of the regions that drove the decision. It needs no change to the network and '
  'lets a reader see whether the model looked at the tumor, the issue raised in {gomes26}.')
P('**Cloud GPU computing.** Training runs on NVIDIA T4 GPUs on Google Colab or Kaggle, which are free, need no local hardware and '
  'give all four team members an identical environment.')
section('Hardware Requirements')
table('Hardware requirements', ['Component', 'Specification'], [
    ['GPU', 'NVIDIA T4, 16 GB (Google Colab / Kaggle) or a local GPU with at least 8 GB'],
    ['CPU and memory', '4 or more cores, 16 GB RAM'],
    ['Storage', 'About 10 GB for datasets, split files, checkpoints and logs'],
    ['Operating system', 'Windows 11 or Ubuntu 22.04'],
], [4.5, 12.2], size=11)

# ===================================================================== 9. CONCLUSION
chapter('Conclusion')
section('Summary')
P('This report set out the design of a four-class MRI brain-tumor classifier that is judged by more than its headline accuracy. '
  'The literature review showed that the best published figures come from one benchmark and one aggregate number, which led to '
  'three gaps: no evidence of cross-dataset generalisation, class-level failures hidden by overall accuracy, and no common '
  'protocol that also reports cost and interpretability. The proposed system answers them with an EfficientNet-B0 classifier, '
  'benchmarked against ResNet-50 and a CNN trained from scratch, trained and evaluated through one de-duplicated, '
  'patient-aware pipeline on two public datasets, and reported per class with Grad-CAM explanations and cost figures.')
section('Expected Outcomes')
bullets(['A classifier that reaches a macro-F1 of at least 0.95 on the untouched D1 test set with fewer than 10 M parameters.',
         'A recall of at least 0.90 for every class, so that no weak class is hidden in the average.',
         'Per-class results on a second dataset with patient-level splits, and a cross-dataset check from D1 to D2.',
         'A Grad-CAM heatmap for every prediction and a browser prototype for demonstration.',
         'Code, configurations and split files released so that every number can be reproduced.'])
P('These targets were fixed before training; they are pass/fail criteria for the next phase, not results.')
section('Plan of Work')
table('Workplan and timeline of Phase 1', ['Task', 'Task owner', '% completed', 'Month'], [
    ['Problem identification', 'Chakardhar', '100 %', 'September'],
    ['Literature survey', 'Nikhil', '100 %', 'September'],
    ['Project objectives', 'Sahil', '100 %', 'September'],
    ['Workplan', 'Yaswanth', '100 %', 'September'],
    ['Requirement analysis', 'Sahil', '100 %', 'October'],
    ['Identification of tools / technology / algorithms', 'Nikhil', '100 %', 'October'],
    ['System design (DFD and UML)', 'Yaswanth', '100 %', 'October'],
    ['Proposed methodology', 'Chakardhar', '100 %', 'October'],
    ['Closure of Phase 1 / project exhibition', 'All members', 'Upcoming', 'November'],
], [7.0, 3.4, 2.8, 3.5], size=11, align={2: 'c', 3: 'c'})
P('In the next phase the five implementation phases of Figure 6.1 will be carried out in order, beginning with data download, '
  'de-duplication and the frozen splits.')

# ===================================================================== 10. REFERENCES
chapter('References')
for k, key in enumerate(ORDER):
    p = d.add_paragraph()
    fmt_para(p, 'j', 1.5, 0, 12)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    add_runs(p, f'[{k + 1}]\t' + BIB[key], size=11)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(1.0))
unused = [k for k in BIB if k not in ORDER]
assert not unused, unused
assert ORDER.index('nick21') == 3 and ORDER.index('cheng15') == 4, ORDER[:6]

d.core_properties.title = TITLE
d.core_properties.author = 'Team DS-A-T6, GITAM School of CSE'
d.core_properties.last_modified_by = 'Team DS-A-T6'
d.save(OUT)
print('saved', OUT, 'refs', len(ORDER), 'figs', FIGN, 'tables', TABN)
