"""Builds the Review-2 deck on a replica of the GITAM Beamer (Boadilla) template."""
import os
import re
import sys
import uuid

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pp import *  # noqa: E402,F401
from lxml import etree  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.chart.data import CategoryChartData  # noqa: E402
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION  # noqa: E402
from pptx.enum.shapes import PP_PLACEHOLDER  # noqa: E402
from pptx.opc.constants import RELATIONSHIP_TYPE as RT  # noqa: E402
from pptx.shapes.shapetree import SlideShapes  # noqa: E402

SP = sys.argv[1]
OUT = sys.argv[2]
LOGO = f'{SP}/latex_zip/images/logo/school_logo.png'


def IC(n, c='w'):
    return f'{SP}/icons/{n}_{c}.png'


TOTAL = 28
DATE = 'November 2026'
TEAM = 'Team DS-A-T6 (GSCSE, GITAM)'
SHORT_TITLE = 'Capstone Project Introduction'
PROJECT = 'A Deep Learning Approach for Brain Tumor Classification Using MRI Images'
SECTIONS = ['Recap of Review 1', 'Problem Statement and Objectives', 'Requirement Analysis',
            'Identification of tools/technology/algorithms', 'Proposed Methodology', 'Project Workflow',
            'Data Flow or Use Case Diagram', 'Conclusion', 'Action Plan', 'Disclosure Statements']
ROUND = MSO_SHAPE.ROUNDED_RECTANGLE

prs = Presentation()
prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)


# =================================================================== theme + layouts
def patch_theme():
    tpart = prs.slide_master.part.part_related_by(RT.THEME)
    xml = tpart.blob.decode('utf-8')
    xml = re.sub(r'(<a:majorFont>\s*<a:latin typeface=")[^"]*(")', r'\1Cambria\2', xml)
    xml = re.sub(r'(<a:minorFont>\s*<a:latin typeface=")[^"]*(")', r'\1Cambria\2', xml)
    for k, v in {'dk2': '1F2D2B', 'lt2': 'E6F2E6', 'accent1': '007367', 'accent2': 'A57C52',
                 'accent3': '198176', 'accent4': '66ABA4', 'accent5': '99C7C2', 'accent6': '008000'}.items():
        xml = re.sub(rf'(<a:{k}>\s*<a:srgbClr val=")[0-9A-Fa-f]{{6}}(")', rf'\g<1>{v}\2', xml)
    xml = re.sub(r'<a:clrScheme name="[^"]*"', '<a:clrScheme name="GITAM Capstone"', xml)
    if hasattr(tpart, '_element'):
        tpart._element = etree.fromstring(xml.encode('utf-8'))
    else:
        tpart._blob = xml.encode('utf-8')


def add_footer(layout, with_logo):
    sh = SlideShapes(layout.shapes._spTree, layout)
    seg, y, h = 13.333 / 3, 7.22, 0.28
    for i, (col, txt, tcol) in enumerate([(FOOT_L, TEAM, WHITE), (FOOT_M, SHORT_TITLE, WHITE), (FOOT_R, None, '000000')]):
        b = box(sh, i * seg, y, seg + (0.01 if i < 2 else 0), h, col, name=f'Footer band {i + 1}')
        if txt:
            shape_text(b, txt, size=11, color=tcol, margin=(0.05, 0, 0.05, 0))
    text(sh, 2 * seg, y, 2.9, h, DATE, size=11, color='000000', align='c', anchor='m', margin=(0.05, 0, 0.05, 0),
         name='Footer date')
    tb = text(sh, 2 * seg + 2.9, y, seg - 3.02, h, '', size=11, color='000000', align='r', anchor='m',
              margin=(0, 0, 0, 0), name='Footer page number')
    p = tb.text_frame.paragraphs[0]
    fld = etree.SubElement(p._p, qn('a:fld'), id='{' + str(uuid.uuid4()).upper() + '}', type='slidenum')
    rPr = etree.SubElement(fld, qn('a:rPr'), lang='en-US', sz='1100')
    sf = etree.SubElement(rPr, qn('a:solidFill'))
    etree.SubElement(sf, qn('a:srgbClr'), val='000000')
    etree.SubElement(rPr, qn('a:latin'), typeface='Cambria')
    etree.SubElement(fld, qn('a:t')).text = '‹#›'
    r = p.add_run()
    r.text = f' / {TOTAL}'
    r.font.size, r.font.name, r.font.color.rgb = Pt(11), 'Cambria', rgb('000000')
    if with_logo:
        pic = sh.add_picture(LOGO, Inches(10.58), Inches(6.6), Inches(2.55))
        pic.name = 'GITAM logo'


def setup_layouts():
    content, blank = prs.slide_layouts[5], prs.slide_layouts[6]
    for lay in (content, blank):
        for ph in list(lay.placeholders):
            if ph.placeholder_format.type != PP_PLACEHOLDER.TITLE:
                ph._element.getparent().remove(ph._element)
    content._element.cSld.set('name', 'GITAM Content')
    blank._element.cSld.set('name', 'GITAM Title')
    ph = content.placeholders[0]
    ph.left, ph.top, ph.width, ph.height = Inches(0.3), Inches(0.16), Inches(12.7), Inches(0.7)
    txBody = ph._element.txBody
    bodyPr = txBody.find(qn('a:bodyPr'))
    bodyPr.set('anchor', 'ctr')
    for k in ('lIns', 'rIns'):
        bodyPr.set(k, '45720')
    lst = txBody.find(qn('a:lstStyle'))
    new = etree.fromstring(
        '<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:lvl1pPr algn="l">'
        '<a:defRPr sz="2800" b="0"><a:solidFill><a:srgbClr val="007367"/></a:solidFill>'
        '<a:latin typeface="Cambria"/><a:ea typeface="Cambria"/><a:cs typeface="Cambria"/></a:defRPr>'
        '</a:lvl1pPr></a:lstStyle>')
    lst.addprevious(new)
    lst.getparent().remove(lst)
    add_footer(content, True)
    add_footer(blank, False)
    return content, blank


patch_theme()
L_CONTENT, L_TITLE = setup_layouts()
SECTION_STARTS = []  # (section name, slide index)


def new_slide(title=None, notes=None, layout='content'):
    s = prs.slides.add_slide(L_CONTENT if layout == 'content' else L_TITLE)
    if title:
        s.shapes.title.text = title
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def block(s, x, y, w, h, title=None, paras=None, size=15, title_size=18, title_h=0.42, pad=0.16,
          space_after=4, anchor='t'):
    box(s.shapes, x, y, w, h, BB_BG, kind=ROUND, radius=0.07, shade=True, name='Block body')
    ty = y
    if title is not None:
        band = top_round(s.shapes, x, y, w, title_h, BT_BG, radius=0.07, name='Block title')
        shape_text(band, title, size=title_size, color=BT_TX, align='l', anchor='m', margin=(0.13, 0.02, 0.1, 0.02))
        ty = y + title_h
    inner = (x + pad, ty + 0.08, w - 2 * pad, y + h - ty - 0.14)
    if paras:
        text(s.shapes, *inner, paras, size=size, space_after=space_after, anchor=anchor)
    return inner


def caption(s, x, y, w, label, body, size=12, align='c'):
    text(s.shapes, x, y, w, 0.3, f'{{{{c:{G}|{label}}}}} {body}', size=size, color=INK, align=align, anchor='m')


def outline(cur):
    SECTION_STARTS.append((SECTIONS[cur - 1], len(prs.slides._sldIdLst)))
    s = new_slide('Outline', notes=f'Section {cur} of 10: {SECTIONS[cur - 1]}.')
    for i, name in enumerate(SECTIONS):
        y = 1.02 + i * 0.53
        on = (i + 1 == cur)
        d = box(s.shapes, 0.45, y, 0.34, 0.34, G if on else FADE, kind=MSO_SHAPE.OVAL)
        shape_text(d, str(i + 1), size=12, color=WHITE, bold=True, margin=(0, 0, 0, 0))
        text(s.shapes, 0.92, y - 0.07, 9.0, 0.48, name, size=19, color=G if on else FADE, anchor='m')
    return s


def stick(s, cx, top, label, color=INK):
    box(s.shapes, cx - 0.18, top, 0.36, 0.36, WHITE, color, 1.75, kind=MSO_SHAPE.OVAL)
    for (a, b, c, d) in [(cx, top + 0.36, cx, top + 0.92), (cx - 0.3, top + 0.55, cx + 0.3, top + 0.55),
                         (cx, top + 0.92, cx - 0.26, top + 1.3), (cx, top + 0.92, cx + 0.26, top + 1.3)]:
        arrow(s.shapes, a, b, c, d, color=color, width=1.75, head=None)
    text(s.shapes, cx - 0.9, top + 1.34, 1.8, 0.5, label.split('|'), size=12, bold=True, align='c', color=INK)


# =================================================================== 1. title
s = new_slide(layout='title', notes=(
    'Opening. Introduce the team and guide, then state the one-line aim: a four-class MRI brain-tumor '
    'classifier that is judged per class, on two public datasets, under one reproducible protocol. '
    'This review covers requirements, tools and algorithms, the methodology and the system design.'))
text(s.shapes, 0.5, 0.52, 12.33, 0.45, 'Review 2', size=24, color=G, align='c', anchor='m')
text(s.shapes, 0.5, 0.96, 12.33, 0.6, 'PROJ2999 Capstone Project Introduction', size=31, color=G, align='c', anchor='m')
text(s.shapes, 0.5, 1.62, 12.33, 0.55, PROJECT, size=21, color=G, bold=True, align='c', anchor='m')
text(s.shapes, 0.5, 2.5, 12.33, 0.8, [
    {'text': f'**B Chakardhar** {{{{c:{MUTED}|(2023002241)}}}}  **Sahil Purkayastha** {{{{c:{MUTED}|(2023005003)}}}}', 'space_after': 4},
    f'**P. Nikhil Mohan** {{{{c:{MUTED}|(2023003536)}}}}  **G Yaswanth** {{{{c:{MUTED}|(2023006345)}}}}'],
    size=16, align='c')
text(s.shapes, 0.5, 3.33, 12.33, 0.4, 'Guide: **Dr. P. Sreekanth**', size=16, align='c', anchor='m')
text(s.shapes, 0.5, 3.98, 12.33, 0.85, ['Department of Artificial Intelligence and Data Science', 'GITAM School of CSE',
                                         'GITAM Deemed to be University, Visakhapatnam'], size=13, align='c')
text(s.shapes, 0.5, 4.98, 12.33, 0.4, DATE, size=16, align='c', anchor='m')
s.shapes.add_picture(LOGO, Inches((13.333 - 3.4) / 2), Inches(5.72), Inches(3.4))

# =================================================================== section 1: recap
outline(1)
s = new_slide('Recap of Review 1', notes=(
    'Review 1 in one line: the reviewed papers report near-perfect accuracy, but almost entirely on one '
    'public benchmark and as a single number. That gave three gaps and the problem we carry into this review. '
    'The strip at the bottom shows where Review 2 sits in Phase 1.'))
cards = [
    ('brain', 'Domain', ['Four-class MRI classification: glioma, meningioma, pituitary tumor and no tumor.']),
    ('books', 'Literature', ['13 publications reviewed: 11 research papers and 2 surveys, published 2021–2026.']),
    ('warn', 'Research gaps', [{'text': '**G1** Accuracy shown on one benchmark', 'space_after': 3},
                               {'text': '**G2** Weak classes hidden in the average', 'space_after': 3},
                               '**G3** No common protocol; cost and clarity ignored']),
    ('target', 'Problem', ['How well do these models hold up on a second dataset, for each class, and at what cost?']),
    ('flag', 'Objectives', ['Train a transfer-learning classifier, then test it per class on two datasets under one protocol.']),
]
for i, (ic, head, body) in enumerate(cards):
    x = 0.42 + i * 2.52
    box(s.shapes, x, 1.38, 2.3, 3.0, BB_BG, kind=ROUND, radius=0.1, shade=True)
    icon_disc(s.shapes, IC(ic), x + 1.15, 1.38, 0.66, G)
    text(s.shapes, x + 0.1, 1.8, 2.1, 0.4, head, size=16, bold=True, color=G, align='c', anchor='m')
    text(s.shapes, x + 0.16, 2.24, 1.98, 2.08, body, size=13.5, align='l', space_after=2)
    if i < 4:
        c = box(s.shapes, x + 2.33, 2.65, 0.16, 0.3, FOOT_M, kind=MSO_SHAPE.CHEVRON)
# progress strip
MS = [1.95, 6.67, 11.0]
arrow(s.shapes, MS[0], 5.05, MS[1], 5.05, color=G, width=3, head=None)
arrow(s.shapes, MS[1], 5.05, MS[2], 5.05, color=FADE, width=3, head=None, dash=MSO_LINE_DASH_STYLE.DASH)
for cx, col, ring, head, sub in [
        (MS[0], G, False, 'Review 1 · 19 Sep 2026', 'Problem, literature, gaps, objectives'),
        (MS[1], G, True, 'Review 2 · this review', 'Requirements, tools, methodology, design'),
        (MS[2], FADE, False, 'Phase 1 closure · Nov 2026', 'Report, poster, project exhibition')]:
    if ring:
        box(s.shapes, cx - 0.2, 4.85, 0.4, 0.4, WHITE, G, 2.5, kind=MSO_SHAPE.OVAL)
        box(s.shapes, cx - 0.1, 4.95, 0.2, 0.2, G, kind=MSO_SHAPE.OVAL)
    else:
        box(s.shapes, cx - 0.14, 4.91, 0.28, 0.28, col, kind=MSO_SHAPE.OVAL)
    text(s.shapes, cx - 1.75, 5.32, 3.5, 0.75, [{'text': head, 'bold': True, 'color': G if col == G else MUTED},
                                                {'text': sub, 'size': 13, 'color': MUTED}], size=14.5, align='c')

s = new_slide('Feedback from Guide/Reviewers and actions taken after Review–1', notes=(
    'Each action answers the feedback item with the same number. Point to the slide that shows the evidence: '
    'objectives on slide 6, the rebuilt reference list on slides 25 and 26, and the fixed datasets, models '
    'and metrics on slides 8 to 14.'))
L = lambda t: f'**{{{{c:{G}|{t}}}}}** '  # noqa: E731
block(s, 0.45, 1.0, 12.43, 2.02, 'Feedback Received in Review 1', [
    {'text': L('F1') + 'Three objectives were too broad to verify — reduce them to one or two precise, measurable objectives.', 'bullet': True},
    {'text': L('F2') + 'Paper numbers on the literature slides did not match the reference list, so citations could not be traced.', 'bullet': True},
    {'text': L('F3') + 'Datasets, models and the evaluation protocol were still open — fix them before starting system design.', 'bullet': True},
], size=15.5, space_after=8, anchor='m')
SUB = {'size': 13, 'color': MUTED, 'indent': 0.62, 'space_after': 9}
ix, iy, iw, ih = block(s, 0.45, 3.27, 12.43, 3.18, 'Action taken')
text(s.shapes, ix, iy, iw, ih, [
    {'text': L('A1') + 'Objectives merged into two, each with a numeric target and a pass/fail rule.  {{c:007367|→ Slide 6}}', 'bullet': True, 'space_after': 1},
    dict(SUB, text='O1: macro-F1 ≥ 0.95 with fewer than 10 M parameters  ·  O2: recall ≥ 0.90 for every class, on a second dataset'),
    {'text': L('A2') + 'Reference list rebuilt in IEEE order; every bracketed number now resolves to its entry.  {{c:007367|→ Slides 25–26}}', 'bullet': True, 'space_after': 1},
    dict(SUB, text='18 references numbered by first citation, with the dataset and algorithm sources added'),
    {'text': L('A3') + 'Two public datasets, three candidate models and a per-class metric set fixed and documented.  {{c:007367|→ Slides 8–14}}', 'bullet': True, 'space_after': 1},
    dict(SUB, text='D1: 7,023 images, 4 classes  ·  D2: 3,064 slices from 233 patients  ·  EfficientNet-B0, ResNet-50, custom CNN'),
], size=15.5, anchor='m')

# =================================================================== section 2: aim
outline(2)
s = new_slide('Aim of the Project', notes=(
    'The problem statement is unchanged in substance from Review 1 and was approved by the guide. '
    'The objectives are now two, and each has a number we can check: O1 is about building the model and '
    'its cost, O2 about evaluating it honestly on a second dataset and per class.'))
block(s, 0.45, 1.0, 12.43, 2.08, 'Problem Statement', [
    'Deep-learning models for brain-tumor MRI report 97–99 % accuracy, but nearly all of it is measured on '
    'a single public benchmark and summarised as one overall number [1]–[3]. For four-class classification '
    '— glioma, meningioma, pituitary tumor and no tumor — it is not known how such models hold up on a '
    'second dataset, which class fails most, or what they cost to run and to explain. This project builds and '
    'evaluates a transfer-learning classifier under a documented, reproducible protocol that answers all three.'],
    size=15.5, anchor='m')
ix, iy, iw, ih = block(s, 0.45, 3.3, 12.43, 3.15, 'Objectives addressing the identified research gaps')
objs = [
    ('O1', 'Build a transfer-learning classifier (EfficientNet-B0, benchmarked against ResNet-50 and a CNN trained '
           'from scratch) that reaches **macro-F1 ≥ 0.95** on the untouched test set of the 7,023-image MRI dataset [4] '
           'with **fewer than 10 M parameters**.', 'Addresses G3'),
    ('O2', 'Evaluate it under one documented protocol on a second public dataset [5] (3,064 slices, 233 patients, '
           'patient-level splits) and report per-class precision, recall, F1, specificity, confusion matrix and '
           'Grad-CAM maps, with **recall ≥ 0.90 for every class**.', 'Addresses G1, G2'),
]
for k, (tag, body, gaps) in enumerate(objs):
    y = iy + 0.06 + k * 1.3
    d = box(s.shapes, ix, y + 0.27, 0.62, 0.62, G, kind=MSO_SHAPE.OVAL)
    shape_text(d, tag, size=16, bold=True, color=WHITE, margin=(0, 0, 0, 0))
    text(s.shapes, ix + 0.8, y, 9.45, 1.16, body, size=15.5, anchor='m')
    c = box(s.shapes, ix + 10.4, y + 0.4, 1.72, 0.36, WHITE, G, 1, kind=ROUND, radius=0.18)
    shape_text(c, gaps, size=12.5, color=G, bold=True, margin=(0.03, 0, 0.03, 0))

# =================================================================== section 3: requirements
outline(3)
s = new_slide('Requirement Analysis', notes=(
    'Functional requirements follow the data from ingestion to a single explained prediction. '
    'Non-functional requirements carry the numbers from the objectives, so each one can be tested.'))
FR = [('FR1', 'Data ingestion & curation', 'load both datasets with labels, check file integrity, remove exact and near-duplicate images.'),
      ('FR2', 'Pre-processing', 'crop to the brain region, resize to 224 × 224, normalise; augment the training split only.'),
      ('FR3', 'Model training', 'train EfficientNet-B0, ResNet-50 and a baseline CNN from one config file with fixed seeds.'),
      ('FR4', 'Evaluation', 'per-class precision, recall, F1, specificity, confusion matrix and ROC-AUC on both datasets.'),
      ('FR5', 'Prediction & explanation', 'accept one MRI slice; return the class, its confidence and a Grad-CAM heatmap.')]
NFR = [('NFR1', 'Accuracy', 'macro-F1 ≥ 0.95 on the held-out test set; recall ≥ 0.90 for every class.'),
       ('NFR2', 'Reproducibility', 'fixed seeds, frozen split files, logged configs; a rerun lands within ±0.5 points of macro-F1.'),
       ('NFR3', 'Efficiency', 'under 10 M parameters; ≤ 50 ms per image on a T4 GPU and ≤ 1 s on a laptop CPU.'),
       ('NFR4', 'Interpretability', 'every prediction carries a heatmap, so a reader can see which region drove it.'),
       ('NFR5', 'Ethics & privacy', 'public, de-identified data only; every output is labelled “research use, not a diagnosis”.')]
for (ttl, rows, y, h) in [('Functional Requirements', FR, 1.0, 2.68), ('Non-functional Requirements', NFR, 3.82, 2.68)]:
    block(s, 0.45, y, 12.43, h, ttl, [
        {'text': f'**{{{{c:{G}|{c}}}}}** **{n}** — {d}', 'bullet': True} for c, n, d in rows],
        size=15, space_after=8)

s = new_slide('Requirement Analysis (contd.)', notes=(
    'Three groups of users drive the requirements. The clinician needs a fast answer with a visible reason; '
    'the team needs a pipeline that reproduces every number; the guide and reviewers need per-class, '
    'second-dataset results. The right side lists the hardware and software the project needs.'))
ix, iy, iw, ih = block(s, 0.45, 1.0, 6.1, 5.45, 'Users and their needs')
users = [('doctor', 'Clinician / radiologist', 'end user',
          'A second opinion on a single slice in seconds, with a visible reason (heatmap) instead of a bare label.'),
         ('users', 'Research team', 'developers',
          'One config-driven pipeline to retrain, compare models and reproduce every reported number.'),
         ('grad', 'Guide & reviewers', 'evaluators',
          'Per-class and second-dataset results that can be checked against the published work.')]
for k, (ic, who, role, need) in enumerate(users):
    y = iy + 0.12 + k * 1.58
    icon_disc(s.shapes, IC(ic), ix + 0.45, y + 0.55, 0.78, G)
    text(s.shapes, ix + 1.05, y, iw - 1.1, 1.45, [
        {'text': f'**{who}**  {{{{c:{MUTED}|({role})}}}}', 'size': 15.5, 'space_after': 3},
        {'text': need, 'size': 14}], anchor='m')
ix, iy, iw, ih = block(s, 6.8, 1.0, 6.08, 5.45, 'Hardware and software requirements')
rows = [('GPU', 'NVIDIA T4 16 GB (Google Colab / Kaggle) or a local GPU ≥ 8 GB'),
        ('CPU / RAM', '4+ cores, 16 GB RAM'),
        ('Storage', '≈ 10 GB for datasets, splits, checkpoints and logs'),
        ('Language', 'Python 3.11'),
        ('Deep learning', 'PyTorch 2.x, torchvision'),
        ('Data & metrics', 'NumPy, pandas, scikit-learn'),
        ('Imaging', 'OpenCV, Albumentations, imagehash'),
        ('XAI & interface', 'pytorch-grad-cam, Streamlit'),
        ('Tooling', 'Git + GitHub, TensorBoard; Windows 11 / Ubuntu 22.04')]
gf = s.shapes.add_table(len(rows), 2, Inches(ix), Inches(iy + 0.02), Inches(iw), Inches(0.5 * len(rows)))
plain_table(gf)
gf.table.columns[0].width = Inches(1.6)
gf.table.columns[1].width = Inches(iw - 1.6)
for r, (k, v) in enumerate(rows):
    gf.table.rows[r].height = Inches(0.5)
    for c, val in enumerate((k, v)):
        cell = gf.table.cell(r, c)
        cell_style(cell, WHITE if r % 2 == 0 else 'F3F9F3', border='C2DCC2')
        cell_text(cell, val, size=13, bold=(c == 0), color=G if c == 0 else INK)

# =================================================================== section 4: tools
outline(4)
s = new_slide('Identification of Tools/Technology/Algorithms', notes=(
    'Every tool is here for a reason we can name. PyTorch gives us pretrained backbones; scikit-learn gives '
    'audited per-class metrics; Albumentations and OpenCV handle augmentation and cropping; pytorch-grad-cam '
    'produces the heatmaps. The algorithms are justified in detail on the next slide.'))


def grid(s, ix, iy, iw, items, cols, row_h, size=13.5, ic_color='g'):
    cw = iw / cols
    for k, (ic, name, desc) in enumerate(items):
        x = ix + (k % cols) * cw
        y = iy + (k // cols) * row_h
        icon(s.shapes, IC(ic, ic_color), x, y + 0.06, 0.34)
        text(s.shapes, x + 0.45, y, cw - 0.55, row_h, [
            {'text': name, 'bold': True, 'size': size + 0.5, 'space_after': 1},
            {'text': desc, 'size': size, 'color': MUTED}])


ix, iy, iw, ih = block(s, 0.45, 0.98, 12.43, 2.36, 'Tools')
grid(s, ix, iy, iw, [
    ('python', 'PyTorch 2 + torchvision [6]', 'ImageNet-pretrained backbones, mixed-precision GPU training'),
    ('chart', 'scikit-learn [7]', 'audited per-class metrics, confusion matrix, ROC-AUC'),
    ('crop', 'OpenCV + Albumentations [8]', 'brain-region cropping and train-time augmentation'),
    ('fire', 'pytorch-grad-cam [9]', 'class-activation heatmap for every prediction'),
    ('laptop', 'Streamlit', 'browser prototype: upload a slice, see class + heatmap'),
    ('github', 'Git / GitHub + TensorBoard', 'versioned code and configs, training curves'),
], cols=2, row_h=0.6)
ix, iy, iw, ih = block(s, 0.45, 3.5, 12.43, 1.42, 'Technology')
grid(s, ix, iy, iw, [
    ('layers', 'Transfer learning', 'reuses ImageNet features [10]; labelled MRI is scarce'),
    ('eye', 'Explainable AI (Grad-CAM)', 'shows the region behind each decision, as in [11]'),
    ('cloud', 'Cloud GPU (Colab / Kaggle T4)', 'free, identical environment for all four members'),
], cols=3, row_h=0.8, size=13)
ix, iy, iw, ih = block(s, 0.45, 5.08, 12.43, 1.42, 'Algorithms [12], [13]')
grid(s, ix, iy, iw, [
    ('target', 'EfficientNet-B0 (primary)', '5.3 M params; best accuracy per FLOP [12]'),
    ('sitemap', 'ResNet-50 (baseline)', 'residual backbone used widely in the literature [13]'),
    ('puzzle', 'Custom 4-block CNN', 'trained from scratch: what does pre-training add?'),
], cols=3, row_h=0.8, size=13)

s = new_slide('Identification of Algorithms (contd.) — Rationale', notes=(
    'Why three models: EfficientNet-B0 is the candidate we expect to win on accuracy per parameter; ResNet-50 '
    'is the reference most reviewed papers use; the scratch CNN tells us how much pre-training is worth on '
    'small MRI data. The training recipe below is the same for all three, so differences come from the model.'))
hdr = ['Model', 'Params', 'GMACs', 'Init', 'Role in the study']
rows = [('EfficientNet-B0 [12]', '5.3 M', '0.39', 'ImageNet', 'Primary — compound scaling gives the best accuracy per FLOP; meets NFR3'),
        ('ResNet-50 [13]', '25.6 M', '4.1', 'ImageNet', 'Baseline — the backbone most reviewed studies build on'),
        ('Custom CNN', '≈ 0.4 M', '≈ 0.7', 'Random', 'Ablation — measures what pre-training adds on scarce MRI data [17]')]
widths = [1.85, 0.85, 0.8, 0.95, 3.35]
gf = s.shapes.add_table(4, 5, Inches(0.45), Inches(1.05), Inches(sum(widths)), Inches(2.3))
plain_table(gf)
for c, w in enumerate(widths):
    gf.table.columns[c].width = Inches(w)
for r in range(4):
    gf.table.rows[r].height = Inches(0.44 if r == 0 else 0.62)
    for c in range(5):
        cell = gf.table.cell(r, c)
        if r == 0:
            cell_style(cell, G, border='FFFFFF')
            cell_text(cell, hdr[c], size=13, bold=True, color=WHITE, align='l' if c in (0, 4) else 'c')
        else:
            val = rows[r - 1][c]
            cell_style(cell, 'E6F2E6' if r == 1 else WHITE, border='C2DCC2')
            cell_text(cell, val, size=12.5 if c == 4 else 13, bold=(c == 0), color=G if c == 0 else INK,
                      align='l' if c in (0, 4) else 'c')
cd = CategoryChartData()
cd.categories = ['Custom CNN', 'EfficientNet-B0', 'ResNet-50']
cd.add_series('Parameters (millions)', (0.4, 5.3, 25.6))
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(8.55), Inches(0.98), Inches(4.35), Inches(2.45), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = True
ch.chart_title.text_frame.text = 'Parameters (millions)'
tp = ch.chart_title.text_frame.paragraphs[0]
tp.runs[0].font.size, tp.runs[0].font.bold, tp.runs[0].font.color.rgb, tp.runs[0].font.name = Pt(13), False, rgb(INK), 'Cambria'
ch.font.name, ch.font.size = 'Cambria', Pt(12)
plot = ch.plots[0]
plot.gap_width = 55
plot.has_data_labels = True
dl = plot.data_labels
dl.number_format, dl.number_format_is_linked = '0.0" M"', False
dl.position = XL_LABEL_POSITION.OUTSIDE_END
dl.font.size, dl.font.color.rgb = Pt(12), rgb(INK)
for i, col in enumerate([FOOT_R, G, FOOT_R]):
    pt = plot.series[0].points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = rgb(col)
va = ch.value_axis
va.visible = False
va.has_major_gridlines = False
va.maximum_scale = 33
va.minimum_scale = 0
ca = ch.category_axis
ca.tick_labels.font.size = Pt(12)
ca.tick_labels.font.color.rgb = rgb(INK)
ca.format.line.color.rgb = rgb('9FB8B5')
ca.has_major_gridlines = False
ix, iy, iw, ih = block(s, 0.45, 3.68, 12.43, 2.8, 'Common training recipe (identical for all three models)')
recipe = [('Two-stage fine-tuning', '5 epochs on the new head only, then the whole network at a 10× lower learning rate.'),
          ('Optimiser', 'AdamW [14], weight decay 1e-4, batch size 32, mixed precision.'),
          ('Loss', 'class-weighted cross-entropy with label smoothing 0.1 [15], so the smallest class still counts.'),
          ('Schedule', 'cosine learning-rate decay with 1-epoch warm-up [16]; early stopping on validation macro-F1.'),
          ('Model selection', 'on validation data only; the test split is opened once, for the final numbers.'),
          ('Explainability', 'Grad-CAM [9] on the last convolutional block of each model.')]
for k, (lab, body) in enumerate(recipe):
    x = ix + (k % 2) * (iw / 2)
    y = iy + (k // 2) * 0.78
    text(s.shapes, x, y, iw / 2 - 0.25, 0.75, [{'text': f'**{{{{c:{G}|{lab}}}}}** — {body}', 'bullet': True}], size=14)

# =================================================================== section 5: methodology
outline(5)
s = new_slide('Proposed Methodology', notes=(
    'Read the diagram left to right. Images enter from two datasets, are de-duplicated and split before anything '
    'else, so no test image leaks into training. After cropping, resizing and train-only augmentation, the same '
    'pipeline trains three models. Evaluation is per class, and tuning loops back on validation data only. '
    'The output is a class, a confidence and a heatmap.'))
cols = [
    ('INPUT', 'brain', ['D1 · Brain Tumor MRI [4]: 7,023 images, 4 classes',
                        'D2 · Figshare T1-CE [5]: 3,064 slices, 233 patients',
                        'Inference: one uploaded MRI slice']),
    ('CURATION', 'filter', ['File and label integrity check',
                            'Perceptual-hash de-duplication, D1 ↔ D2',
                            'D1: test set (1,311) frozen; rest 85 / 15 train / val',
                            'D2: patient-level 5-fold split']),
    ('PREPROCESS', 'crop', ['Crop to the brain region (largest contour)',
                                'Resize 224 × 224, 3 channels, ImageNet normalisation',
                                'Train-only augmentation: flip, ±15° rotation, brightness / contrast, resized crop']),
    ('TRAINING', 'layers', ['EfficientNet-B0 (ImageNet) — primary',
                            'ResNet-50 (ImageNet) — baseline',
                            'Custom CNN (scratch) — ablation',
                            'AdamW, cosine LR, weighted CE, early stop']),
    ('EVALUATION', 'chart', ['Accuracy, macro-F1',
                             'Per-class precision, recall, F1, specificity',
                             'Confusion matrix, ROC-AUC',
                             'Cost: parameters, ms / image',
                             'Grad-CAM review']),
    ('OUTPUT', 'check', ['Predicted class + confidence',
                         'Grad-CAM heatmap overlay',
                         'Per-class report and model comparison']),
]
PRODUCES = ['labelled MRI slices', 'clean, frozen split lists', '224 × 224 normalised tensors',
            'three trained checkpoints', 'metrics, matrices, heatmaps', 'an explained prediction']
x0, cw, gap, top, bot = 0.32, 1.88, 0.26, 1.0, 5.85
FS = 12
for i, (name, ic, items) in enumerate(cols):
    x = x0 + i * (cw + gap)
    box(s.shapes, x, top, cw, bot - top, BB_BG, kind=ROUND, radius=0.08, shade=True)
    top_round(s.shapes, x, top, cw, 0.8, G_DARK if i in (0, 5) else G, radius=0.08)
    icon(s.shapes, IC(ic), x + 0.12, top + 0.2, 0.4)
    text(s.shapes, x + 0.6, top + 0.06, cw - 0.62, 0.68, [{'text': f'STEP {i + 1}', 'size': 10, 'color': 'BFE3DE'},
                                                         {'text': name, 'size': 12, 'bold': True, 'color': WHITE}],
         anchor='m', margin=(0, 0, 0, 0))
    y = top + 0.93
    for it in items:
        h = n_lines(it, cw - 0.34, FS) * FS * 1.2 / 72 + 0.12
        b = box(s.shapes, x + 0.1, y, cw - 0.2, h, WHITE, 'B5D3CF', 0.75, kind=ROUND, radius=0.06)
        shape_text(b, it, size=FS, align='l', anchor='m', margin=(0.07, 0.03, 0.05, 0.03))
        y += h + 0.08
    assert y < bot - 0.72, (name, y)
    pc = box(s.shapes, x + 0.1, bot - 0.66, cw - 0.2, 0.56, 'FBF6F0', BROWN, 1, kind=ROUND, radius=0.08)
    shape_text(pc, [{'text': 'PRODUCES', 'size': 9, 'bold': True, 'color': BROWN},
                    {'text': PRODUCES[i], 'size': 11.5, 'color': INK}], align='c', margin=(0.04, 0.02, 0.04, 0.02))
    if i < 5:
        box(s.shapes, x + cw + 0.05, top + 0.25, 0.16, 0.3, BROWN, kind=MSO_SHAPE.CHEVRON)
xe, xt = x0 + 4 * (cw + gap) + cw / 2, x0 + 3 * (cw + gap) + cw / 2
polyline(s.shapes, [(xe, bot), (xe, 6.15), (xt, 6.15), (xt, bot)], color=BROWN, width=2)
text(s.shapes, xt - 0.3, 6.18, xe - xt + 0.6, 0.3, 'tuning loop — validation data only', size=11, italic=True,
     color=BROWN, align='c', anchor='m')
caption(s, 0.32, 6.06, 6.9, 'Figure:', 'Proposed methodology — from MRI input to an explained prediction', size=12)

# =================================================================== section 6: workflow
outline(6)
s = new_slide('Proposed Workflow', notes=(
    'Five implementation phases, each ending in a concrete deliverable. After phase 4 there is a quality gate: '
    'if the targets are not met we go back to phase 3 with error analysis rather than moving on. Only when both '
    'targets hold do we build the prototype and write up.'))
phases = [
    ('P1', 'Data & environment', ['Download D1 and D2; record licences and class counts',
                                  'EDA: class balance, image sizes, duplicates',
                                  'Freeze split files; set up repo, configs, seeds'], 'frozen splits + EDA notebook'),
    ('P2', 'Pipeline & baselines', ['Pre-processing and augmentation module',
                                    'Training loop with logging and checkpoints',
                                    'Train the custom CNN and ResNet-50'], 'baseline validation scores'),
    ('P3', 'Primary model & tuning', ['Two-stage fine-tuning of EfficientNet-B0',
                                      'Tune LR, augmentation, class weights on validation',
                                      'Ablation: with vs. without pre-training'], 'best checkpoint per model'),
    ('P4', 'Evaluation & explainability', ['One-time test on D1; patient-level 5-fold on D2',
                                           'Per-class metrics, confusion matrices, ROC',
                                           'Grad-CAM review; latency and size profiling'], 'evaluation report'),
    ('P5', 'Prototype & documentation', ['Streamlit demo: upload → class + heatmap',
                                         'Final report, poster and paper draft',
                                         'Release code, configs and weights on GitHub'], 'demo app + final report'),
]
shades = ['00564D', '00665B', '007367', '1A8176', '2B8C82']
cxs = []
for i, (pid, name, acts, deliv) in enumerate(phases):
    x = 0.35 + i * 2.55
    cxs.append(x + 1.175)
    kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
    ch_ = box(s.shapes, x - (0 if i == 0 else 0.12), 0.98, 2.5 if i == 0 else 2.62, 0.72, shades[i], kind=kind)
    ch_.adjustments[0] = 0.32
    shape_text(ch_, [{'text': pid, 'size': 11, 'color': 'BFE3DE', 'bold': True},
                     {'text': name, 'size': 12.5, 'bold': True, 'color': WHITE}],
               align='c', margin=(0.3 if i else 0.1, 0.02, 0.25, 0.02))
    box(s.shapes, x, 1.86, 2.35, 3.14, BB_BG, kind=ROUND, radius=0.08, shade=True)
    text(s.shapes, x + 0.1, 1.98, 2.17, 2.3, [{'text': a, 'bullet': True, 'space_after': 7, 'indent': 0.2} for a in acts],
         size=13)
    chip = box(s.shapes, x + 0.12, 4.32, 2.11, 0.58, WHITE, G, 1, kind=ROUND, radius=0.1)
    shape_text(chip, [{'text': 'Deliverable', 'size': 10, 'bold': True, 'color': BROWN},
                      {'text': deliv, 'size': 11.5, 'color': INK}], align='c', margin=(0.04, 0.02, 0.04, 0.02))
dg = box(s.shapes, cxs[3] - 1.2, 5.27, 2.4, 1.08, WHITE, BROWN, 1.75, kind=MSO_SHAPE.DIAMOND)
shape_text(dg, [{'text': 'Targets met?', 'bold': True, 'size': 12}, {'text': 'macro-F1 ≥ 0.95, recall ≥ 0.90', 'size': 9.5}],
           margin=(0.02, 0, 0.02, 0))
arrow(s.shapes, cxs[3], 5.0, cxs[3], 5.27, color=BROWN, width=1.75)
polyline(s.shapes, [(cxs[3] + 1.2, 5.81), (cxs[4], 5.81), (cxs[4], 5.0)], color=G, width=1.75)
polyline(s.shapes, [(cxs[3] - 1.2, 5.81), (cxs[2], 5.81), (cxs[2], 5.0)], color=BROWN, width=1.75, dash=MSO_LINE_DASH_STYLE.DASH)
text(s.shapes, cxs[3] + 1.25, 5.5, 1.2, 0.3, '**Yes** → P5', size=12, color=G, anchor='m')
text(s.shapes, cxs[2] + 0.08, 5.86, 1.3, 0.45, ['**No** → back to P3', 'with error analysis'], size=11, color=BROWN, anchor='t')
caption(s, 0.35, 6.08, 5.4, 'Figure:', 'Key implementation phases and the quality gate', size=12, align='l')

# =================================================================== section 7: DFD / use case
outline(7)


def entity(s, cx, cy, w, h, name, dup=False):
    b = box(s.shapes, cx - w / 2, cy - h / 2, w, h, FOOT_L, kind=MSO_SHAPE.RECTANGLE, shade=True)
    parts = name.split('|')
    shape_text(b, [parts[0]] + [{'text': p, 'size': 11, 'bold': False} for p in parts[1:]], size=12, bold=True,
               color=WHITE, margin=(0.06, 0.02, 0.06, 0.02))
    if dup:
        arrow(s.shapes, cx + w / 2 - 0.2, cy + h / 2, cx + w / 2, cy + h / 2 - 0.2, color=WHITE, width=1.5, head=None)
    return b


def process(s, cx, cy, w, h, num, name, size=12):
    x, y = cx - w / 2, cy - h / 2
    box(s.shapes, x, y, w, h, WHITE, G, 1.75, kind=ROUND, radius=0.12, shade=True)
    band = top_round(s.shapes, x, y, w, 0.3, G, radius=0.12)
    shape_text(band, num, size=11, bold=True, color=WHITE, margin=(0, 0, 0, 0))
    text(s.shapes, x + 0.05, y + 0.3, w - 0.1, h - 0.32, name, size=size, bold=True, align='c', anchor='m', color=INK)


def store(s, cx, cy, w, h, sid, name):
    x, y = cx - w / 2, cy - h / 2
    box(s.shapes, x, y, w, h, 'F1F7F6')
    for a in [(x, y, x + w, y), (x, y + h, x + w, y + h), (x, y, x, y + h), (x + 0.42, y, x + 0.42, y + h)]:
        arrow(s.shapes, *a, color=INK, width=1.5, head=None)
    text(s.shapes, x, y, 0.42, h, sid, size=12, bold=True, color=G, align='c', anchor='m', margin=(0, 0, 0, 0))
    text(s.shapes, x + 0.46, y, w - 0.48, h, name, size=11.5, align='l', anchor='m', margin=(0.02, 0, 0.02, 0))


def flow_label(s, x, y, w, h, t, align='c', size=10.5):
    text(s.shapes, x, y, w, h, t, size=size, italic=True, color=MUTED, align=align, anchor='m', margin=(0, 0, 0, 0))


s = new_slide('Data Flow Diagram (DFD) — Level 0 (Context Diagram)', notes=(
    'The context diagram treats the whole system as one process. Three external entities exchange data with it: '
    'the public datasets supply labelled images, the research team supplies configurations and receives the '
    'evaluation report, and the end user sends one MRI slice and receives the class, confidence and heatmap.'))
entity(s, 6.67, 1.38, 3.0, 0.75, 'Public MRI datasets|D1 Kaggle · D2 Figshare')
process(s, 6.67, 3.85, 2.9, 1.9, '0', 'Brain-Tumor MRI Classification System', size=15)
entity(s, 1.65, 3.85, 2.3, 0.9, 'Clinician / End user')
entity(s, 11.68, 3.85, 2.3, 0.9, 'Research team')
arrow(s.shapes, 6.67, 1.755, 6.67, 2.9, width=1.75)
flow_label(s, 6.8, 2.1, 2.6, 0.5, 'labelled MRI images', align='l', size=11.5)
arrow(s.shapes, 2.8, 3.6, 5.22, 3.6, width=1.75)
flow_label(s, 2.85, 3.18, 2.32, 0.38, 'MRI slice (PNG / JPG)', size=11.5)
arrow(s.shapes, 5.22, 4.1, 2.8, 4.1, width=1.75)
flow_label(s, 2.85, 4.15, 2.32, 0.55, 'class, confidence, Grad-CAM heatmap', size=11.5)
arrow(s.shapes, 10.53, 3.6, 8.12, 3.6, width=1.75)
flow_label(s, 8.17, 3.18, 2.32, 0.38, 'config, hyper-parameters', size=11.5)
arrow(s.shapes, 8.12, 4.1, 10.53, 4.1, width=1.75)
flow_label(s, 8.17, 4.15, 2.32, 0.55, 'evaluation report (metrics, heatmaps)', size=11.5)
# legend
lx, ly = 0.6, 5.55
box(s.shapes, lx, ly, 0.42, 0.3, FOOT_L)
text(s.shapes, lx + 0.5, ly - 0.03, 1.5, 0.36, 'External entity', size=11.5, anchor='m')
box(s.shapes, lx + 2.15, ly, 0.42, 0.3, WHITE, G, 1.5, kind=ROUND, radius=0.06)
text(s.shapes, lx + 2.65, ly - 0.03, 1.1, 0.36, 'Process', size=11.5, anchor='m')
store(s, lx + 4.25, ly + 0.15, 0.9, 0.3, '', '')
text(s.shapes, lx + 4.8, ly - 0.03, 1.2, 0.36, 'Data store', size=11.5, anchor='m')
arrow(s.shapes, lx + 6.1, ly + 0.15, lx + 6.65, ly + 0.15, width=1.5)
text(s.shapes, lx + 6.75, ly - 0.03, 1.1, 0.36, 'Data flow', size=11.5, anchor='m')
text(s.shapes, lx + 7.95, ly - 0.03, 2.0, 0.36, '(Gane–Sarson notation)', size=11, italic=True, color=MUTED, anchor='m')
caption(s, 0.45, 6.15, 9.8, 'Figure:', 'Level-0 DFD of the proposed system', size=12)

s = new_slide('Data Flow Diagram (DFD) — Level 1', notes=(
    'Level 1 opens the single process into five. The top row is the training pipeline: ingest and curate, '
    'pre-process and split, then train. Evaluation reads the frozen test split and the saved model and writes '
    'results that the team receives as a report. Prediction reads the same saved model to answer the end user. '
    'Flows match the context diagram, so the two levels are balanced.'))
C = [1.2, 3.75, 6.3, 8.85, 11.4]
R = [1.55, 3.4, 5.3]
EW, PW, SW = 1.75, 1.75, 1.75
entity(s, C[0], R[0], EW, 0.8, 'Public MRI datasets')
process(s, C[1], R[0], PW, 1.0, '1.0', 'Ingest & curate')
store(s, C[2], R[0], SW, 0.55, 'D1', 'Curated images')
process(s, C[3], R[0], PW, 1.0, '2.0', 'Pre-process & split')
store(s, C[4], R[0], SW, 0.55, 'D2', 'Split sets')
entity(s, C[0], R[1], EW, 0.8, 'Research team*', dup=True)
store(s, C[1], R[1], SW, 0.55, 'D4', 'Results & logs')
process(s, C[2], R[1], PW, 1.0, '4.0', 'Evaluate & explain')
entity(s, C[3], R[1], EW, 0.8, 'Research team')
process(s, C[4], R[1], PW, 1.0, '3.0', 'Train models')
entity(s, C[0], R[2], EW, 0.8, 'End user')
process(s, C[1], R[2], PW, 1.0, '5.0', 'Predict & explain')
store(s, C[4], R[2], SW, 0.55, 'D3', 'Model registry')
half = 0.875
for a, b, lab in [(0, 1, 'images + labels'), (1, 2, 'clean images'), (2, 3, 'images'), (3, 4, 'split files')]:
    arrow(s.shapes, C[a] + half, R[0], C[b] - half, R[0], width=1.5)
    flow_label(s, C[a] + half + 0.02, R[0] - 0.52, C[b] - C[a] - 2 * half - 0.04, 0.48, lab, size=10.5)
arrow(s.shapes, C[4], R[0] + 0.275, C[4], R[1] - 0.5, width=1.5)
flow_label(s, C[4] + 0.08, 2.1, 1.0, 0.4, 'train + val sets', align='l')
polyline(s.shapes, [(C[4] - 0.55, R[0] + 0.275), (C[4] - 0.55, 2.42), (C[2], 2.42), (C[2], R[1] - 0.5)], width=1.5)
flow_label(s, 7.2, 2.12, 1.4, 0.28, 'test set')
arrow(s.shapes, C[3] + half, R[1], C[4] - half, R[1], width=1.5)
flow_label(s, C[3] + half + 0.02, R[1] - 0.5, 0.76, 0.46, 'config', size=10.5)
arrow(s.shapes, C[2] - half, R[1], C[1] + half, R[1], width=1.5)
flow_label(s, C[1] + half + 0.02, R[1] - 0.52, 0.76, 0.48, 'metrics, maps', size=10.5)
arrow(s.shapes, C[1] - half, R[1], C[0] + half, R[1], width=1.5)
flow_label(s, C[0] + half + 0.02, R[1] - 0.5, 0.76, 0.46, 'report', size=10.5)
arrow(s.shapes, C[4], R[1] + 0.5, C[4], R[2] - 0.275, width=1.5)
flow_label(s, C[4] + 0.08, 4.62, 1.1, 0.3, 'checkpoint', align='l')
polyline(s.shapes, [(C[4] - 0.55, R[2] - 0.275), (C[4] - 0.55, 4.42), (C[2], 4.42), (C[2], R[1] + 0.5)], width=1.5)
flow_label(s, 7.2, 4.12, 1.4, 0.28, 'trained model')
arrow(s.shapes, C[4] - half, R[2], C[1] + half, R[2], width=1.5)
flow_label(s, 6.2, R[2] - 0.32, 2.6, 0.28, 'trained model (read-only)')
arrow(s.shapes, C[0] + half, R[2] - 0.2, C[1] - half, R[2] - 0.2, width=1.5)
flow_label(s, C[0] + half + 0.02, R[2] - 0.62, 0.76, 0.4, 'MRI slice', size=10.5)
arrow(s.shapes, C[1] - half, R[2] + 0.2, C[0] + half, R[2] + 0.2, width=1.5)
flow_label(s, C[0] + half + 0.02, R[2] + 0.24, 0.76, 0.46, 'class + heatmap', size=10.5)
text(s.shapes, 5.6, 6.05, 4.8, 0.28, '* duplicate of the same entity, drawn twice to avoid crossing flows', size=10.5,
     italic=True, color=MUTED, anchor='m')
caption(s, 0.45, 6.05, 5.0, 'Figure:', 'Level-1 DFD of the proposed system', size=12, align='l')

s = new_slide('Use Case Diagram', notes=(
    'Three actors use the system. The end user uploads a slice, sees the class and confidence, can open the '
    'heatmap and download a report. The research team curates data, trains and evaluates. The guide and '
    'reviewers read the evaluation report. Upload and curation both include the same validation and '
    'pre-processing step, so a single routine serves training and inference.'))
box(s.shapes, 2.85, 0.98, 7.4, 5.42, 'FBFDFC', G, 1.5, kind=ROUND, radius=0.1)
text(s.shapes, 2.95, 1.02, 7.2, 0.36, 'MRI Brain-Tumor Classification System', size=14, bold=True, color=G, align='c', anchor='m')


def uc(cx, cy, label, w=2.5, h=0.66):
    e = box(s.shapes, cx - w / 2, cy - h / 2, w, h, WHITE, G, 1.5, kind=MSO_SHAPE.OVAL, shade=False)
    shape_text(e, label, size=12, margin=(0.12, 0.02, 0.12, 0.02))


LX, RX, MX = 4.35, 8.7, 7.4
LY = [1.85, 2.8, 3.75, 4.7]
RY = [2.8, 3.75, 4.7, 5.65]
for y, t in zip(LY, ['Upload MRI slice', 'View class & confidence', 'View Grad-CAM heatmap', 'Download prediction report']):
    uc(LX, y, t)
uc(MX, 1.85, 'Validate & pre-process image', w=2.35)
for y, t in zip(RY, ['Curate & split datasets', 'Train & tune models', 'Evaluate per class & cross-dataset', 'Review evaluation report']):
    uc(RX, y, t)
stick(s, 1.45, 2.55, 'Clinician /|End user')
stick(s, 11.65, 2.45, 'Research team')
stick(s, 11.65, 4.62, 'Guide / Reviewer')
for y in LY:
    arrow(s.shapes, 1.8, 3.12, LX - 1.25, y, color=INK, width=1.25, head=None)
for y in RY[:3]:
    arrow(s.shapes, 11.3, 3.02, RX + 1.25, y, color=INK, width=1.25, head=None)
arrow(s.shapes, 11.3, 5.2, RX + 1.25, RY[3], color=INK, width=1.25, head=None)
D = MSO_LINE_DASH_STYLE.DASH
arrow(s.shapes, LX + 1.25, 1.85, MX - 1.175, 1.85, color=BROWN, width=1.25, head='arrow', dash=D)
text(s.shapes, 5.55, 1.5, 0.8, 0.28, '«include»', size=10, italic=True, color=BROWN, align='c', anchor='m', margin=(0, 0, 0, 0))
arrow(s.shapes, RX - 0.5, RY[0] - 0.33, MX + 0.25, 2.18, color=BROWN, width=1.25, head='arrow', dash=D)
text(s.shapes, 8.4, 2.18, 0.9, 0.28, '«include»', size=10, italic=True, color=BROWN, align='l', anchor='m', margin=(0, 0, 0, 0))
arrow(s.shapes, LX, LY[2] - 0.33, LX, LY[1] + 0.33, color=BROWN, width=1.25, head='arrow', dash=D)
text(s.shapes, LX + 0.08, 3.13, 0.9, 0.28, '«extend»', size=10, italic=True, color=BROWN, align='l', anchor='m', margin=(0, 0, 0, 0))
caption(s, 0.45, 6.5, 9.8, 'Figure:', 'Use case diagram (UML) of the proposed system', size=12)

# =================================================================== section 8: conclusion
outline(8)
s = new_slide('Conclusion', notes=(
    'Summarise: one pipeline, three models, two datasets, per-class reporting and heatmaps. The four numbers '
    'are the targets we fixed before training; they are pass/fail criteria, not results.'))
block(s, 0.45, 1.0, 12.43, 2.42, None, [
    {'text': 'We will build a four-class MRI classifier around **EfficientNet-B0**, fine-tuned from ImageNet and benchmarked '
             'against **ResNet-50** and a CNN trained from scratch. All three pass through the same pipeline: '
             'de-duplicated data, frozen splits, train-only augmentation and validation-only tuning.', 'space_after': 8},
    {'text': 'Each model is scored **per class on two public datasets**, with Grad-CAM showing where it looked and with '
             'its size and latency recorded. The expected outcome is a reproducible account of how well such a classifier '
             'works, for which tumor class, and at what cost — the three questions the reviewed literature leaves open.'}],
      size=16, pad=0.25)
stats = [('≥ 0.95', 'macro-F1 on the untouched D1 test set'),
         ('≥ 0.90', 'recall for every class — no hidden weak class'),
         ('< 10 M', 'parameters; ≤ 50 ms per image on a T4 GPU'),
         ('2 datasets', 'one documented protocol, a heatmap for every prediction')]
for k, (big, lab) in enumerate(stats):
    x = 0.45 + k * 3.18
    box(s.shapes, x, 3.72, 2.95, 1.82, WHITE, G, 1.25, kind=ROUND, radius=0.1, shade=True)
    text(s.shapes, x + 0.1, 3.82, 2.75, 0.85, big, size=34, bold=True, color=G, align='c', anchor='m')
    text(s.shapes, x + 0.2, 4.67, 2.55, 0.8, lab, size=14, color=INK, align='c', anchor='t')
text(s.shapes, 0.45, 5.78, 9.9, 0.32, 'Targets fixed before training — they are pass/fail criteria, not results.',
     size=12.5, italic=True, color=MUTED, anchor='m')

# =================================================================== section 9: action plan
outline(9)
s = new_slide('Workplan and Timeline', notes=(
    'All Review 1 tasks are closed, and the four Review 2 tasks shown in blue are complete with this presentation. '
    'Phase 1 closes in November with the project exhibition.'))
caption(s, 0.45, 0.98, 12.43, 'Table:', 'Proposed timeline of the project', size=14)
hdr = ['Task', 'Task Owner', '% of Task Completed', 'Sept', 'Oct', 'Nov']
plan = [('Problem Identification', 'Chakardhar', '100%', 0), ('Literature Survey', 'Nikhil', '100%', 0),
        ('Project Objectives', 'Sahil', '100%', 0), ('Workplan', 'Yaswanth', '100%', 0),
        ('Requirement Analysis', 'Sahil', '100%', 1),
        ('Identification of Tools/Technology/Algorithms', 'Nikhil', '100%', 1),
        ('System Design (DFD or UMLs)', 'Yaswanth', '100%', 1), ('Proposed Methodology', 'Chakardhar', '100%', 1),
        ('Closure of Phase 1/ Project Exhibition', 'All members', 'Upcoming', 2)]
MONTH = ['34A853', '2F5FD0', 'F4AFC4']
widths = [4.75, 2.0, 2.25, 1.15, 1.15, 1.13]
gf = s.shapes.add_table(10, 6, Inches(0.45), Inches(1.38), Inches(sum(widths)), Inches(4.7))
plain_table(gf)
for c, w in enumerate(widths):
    gf.table.columns[c].width = Inches(w)
for r in range(10):
    gf.table.rows[r].height = Inches(0.47)
    for c in range(6):
        cell = gf.table.cell(r, c)
        if r == 0:
            cell_style(cell, G, border='FFFFFF')
            cell_text(cell, hdr[c], size=13.5, bold=True, color=WHITE, align='l' if c == 0 else 'c')
            continue
        task, owner, pct, m = plan[r - 1]
        if c >= 3:
            cell_style(cell, MONTH[m] if c - 3 == m else WHITE, border='9FB8B5')
            cell_text(cell, '', size=12)
        else:
            cell_style(cell, WHITE, border='9FB8B5')
            cell_text(cell, (task, owner, pct)[c], size=13.5, align='l' if c == 0 else 'c',
                      bold=(c == 2 and pct == '100%'), color=G if (c == 2 and pct == '100%') else INK)
for k, (col, lab) in enumerate(zip(MONTH, ['Sept — Review 1 scope', 'Oct — Review 2 scope', 'Nov — Phase 1 closure'])):
    x = 0.45 + k * 2.85
    box(s.shapes, x, 6.24, 0.24, 0.24, col)
    text(s.shapes, x + 0.32, 6.19, 2.4, 0.34, lab, size=12, color=MUTED, anchor='m')

REFS = [
    'M. Aamir et al., “A deep learning approach for brain tumor classification using MRI images,” *Computers and Electrical Engineering*, vol. 101, art. no. 108105, 2022, doi: 10.1016/j.compeleceng.2022.108105.',
    'F. J. Díaz-Pernas, M. Martínez-Zarzuela, M. Antón-Rodríguez, and D. González-Ortega, “A deep learning approach for brain tumor classification and segmentation using a multiscale convolutional neural network,” *Healthcare*, vol. 9, no. 2, art. no. 153, 2021, doi: 10.3390/healthcare9020153.',
    'M. Aamir, Z. Rahman, U. A. Bhatti, W. A. Abro, J. A. Bhutto, and Z. He, “An automated deep learning framework for brain tumor classification using MRI imagery,” *Scientific Reports*, vol. 15, art. no. 17593, 2025, doi: 10.1038/s41598-025-02209-2.',
    'M. Nickparvar, “Brain tumor MRI dataset,” Kaggle, 2021, doi: 10.34740/KAGGLE/DSV/2645886.',
    'J. Cheng et al., “Enhanced performance of brain tumor classification via tumor region augmentation and partition,” *PLoS ONE*, vol. 10, no. 10, art. no. e0140381, 2015, doi: 10.1371/journal.pone.0140381.',
    'A. Paszke et al., “PyTorch: An imperative style, high-performance deep learning library,” in *Proc. Adv. Neural Inf. Process. Syst. (NeurIPS)*, vol. 32, 2019, pp. 8024–8035.',
    'F. Pedregosa et al., “Scikit-learn: Machine learning in Python,” *J. Mach. Learn. Res.*, vol. 12, pp. 2825–2830, 2011.',
    'A. Buslaev, V. I. Iglovikov, E. Khvedchenya, A. Parinov, M. Druzhinin, and A. A. Kalinin, “Albumentations: Fast and flexible image augmentations,” *Information*, vol. 11, no. 2, art. no. 125, 2020, doi: 10.3390/info11020125.',
    'R. R. Selvaraju, M. Cogswell, A. Das, R. Vedantam, D. Parikh, and D. Batra, “Grad-CAM: Visual explanations from deep networks via gradient-based localization,” in *Proc. IEEE Int. Conf. Comput. Vis. (ICCV)*, 2017, pp. 618–626, doi: 10.1109/ICCV.2017.74.',
    'J. Deng, W. Dong, R. Socher, L.-J. Li, K. Li, and L. Fei-Fei, “ImageNet: A large-scale hierarchical image database,” in *Proc. IEEE Conf. Comput. Vis. Pattern Recognit. (CVPR)*, 2009, pp. 248–255, doi: 10.1109/CVPR.2009.5206848.',
    'E. F. Gomes and R. S. Barbosa, “Deep learning approaches for brain tumor classification in MRI scans: An analysis of model interpretability,” *Applied Sciences*, vol. 16, no. 2, art. no. 831, 2026, doi: 10.3390/app16020831.',
    'M. Tan and Q. V. Le, “EfficientNet: Rethinking model scaling for convolutional neural networks,” in *Proc. 36th Int. Conf. Mach. Learn. (ICML)*, 2019, pp. 6105–6114.',
    'K. He, X. Zhang, S. Ren, and J. Sun, “Deep residual learning for image recognition,” in *Proc. IEEE Conf. Comput. Vis. Pattern Recognit. (CVPR)*, 2016, pp. 770–778, doi: 10.1109/CVPR.2016.90.',
    'I. Loshchilov and F. Hutter, “Decoupled weight decay regularization,” in *Proc. Int. Conf. Learn. Represent. (ICLR)*, 2019.',
    'C. Szegedy, V. Vanhoucke, S. Ioffe, J. Shlens, and Z. Wojna, “Rethinking the Inception architecture for computer vision,” in *Proc. IEEE Conf. Comput. Vis. Pattern Recognit. (CVPR)*, 2016, pp. 2818–2826, doi: 10.1109/CVPR.2016.308.',
    'I. Loshchilov and F. Hutter, “SGDR: Stochastic gradient descent with warm restarts,” in *Proc. Int. Conf. Learn. Represent. (ICLR)*, 2017.',
    'V. Vadde and W. Bukaita, “Comparative evaluation of CNN and ResNet18 architectures for MRI-based brain tumor classification using deep learning,” *Medical Research Archives*, vol. 13, no. 12, 2025, doi: 10.18103/mra.v13i12.7100.',
    'S. Bouhafra and H. El Bahi, “Deep learning approaches for brain tumor detection and classification using MRI images (2020 to 2024): A systematic review,” *J. Imaging Inform. Med.*, vol. 38, no. 3, pp. 1403–1433, 2025, doi: 10.1007/s10278-024-01283-8.',
]


def ref_slide(title, start, items, notes):
    s = new_slide(title, notes=notes)
    text(s.shapes, 0.45, 0.95, 12.4, 0.35, 'References in the IEEE format as follows:', size=14, anchor='m')
    tb = s.shapes.add_textbox(Inches(0.45), Inches(1.38), Inches(12.43), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    for k, r in enumerate(items):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(8)
        pPr = p._p.get_or_add_pPr()
        pPr.set('marL', str(int(Inches(0.5))))
        pPr.set('indent', str(-int(Inches(0.5))))
        no_bullet(p)
        parts = re.split(r'(\*[^*]+\*)', f'[{start + k}]\t' + r)
        for part in parts:
            if not part:
                continue
            run = p.add_run()
            ital = part.startswith('*') and part.endswith('*')
            run.text = part[1:-1] if ital else part
            run.font.size, run.font.name, run.font.italic = Pt(13), 'Cambria', ital
            run.font.color.rgb = rgb(INK if not ital else '1F4F4A')
    return s


ref_slide('References (1/2)', 1, REFS[:9], 'IEEE-format list, numbered in order of first citation in this deck.')
ref_slide('References (2/2)', 10, REFS[9:], 'Continuation of the IEEE-format reference list.')

# =================================================================== section 10: disclosure
outline(10)
s = new_slide('Disclosure Statements', notes='Read the three statements as given in the GITAM template.')
text(s.shapes, 0.45, 1.6, 12.2, 3.6, [
    {'text': 'I will maintain academic honesty by doing my own work and avoiding plagiarism, cheating, unauthorized collaboration, and falsifying information, use of AI Content.', 'bullet': True, 'space_after': 12},
    {'text': 'I will properly acknowledge and cite all sources I use, including information from the internet, books, software, and AI tools.', 'bullet': True, 'space_after': 12},
    {'text': 'I will use institutional resources responsibly and only for authorized academic or research activities. I understand that I am responsible for my own actions and decisions.', 'bullet': True}],
    size=19, anchor='m', bullet_color='0B5D0B')

# =================================================================== sections + save
assert len(prs.slides) == TOTAL, len(prs.slides)
P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
ids = [sid.id for sid in prs.slides._sldIdLst]
groups = [('Title', 0)] + SECTION_STARTS
pres = prs.part._element
extLst = pres.find(qn('p:extLst'))
if extLst is None:
    extLst = etree.SubElement(pres, qn('p:extLst'))
ext = etree.SubElement(extLst, qn('p:ext'), uri='{521415D9-36F7-43E2-AB2F-B90AF26B5E84}')
secLst = etree.SubElement(ext, '{%s}sectionLst' % P14, nsmap={'p14': P14})
for gi, (name, start) in enumerate(groups):
    end = groups[gi + 1][1] if gi + 1 < len(groups) else len(ids)
    sec = etree.SubElement(secLst, '{%s}section' % P14, name=name, id='{' + str(uuid.uuid4()).upper() + '}')
    lst = etree.SubElement(sec, '{%s}sldIdLst' % P14)
    for i in range(start, end):
        etree.SubElement(lst, '{%s}sldId' % P14, id=str(ids[i]))
prs.core_properties.title = 'Review 2 — ' + PROJECT
prs.core_properties.author = 'Team DS-A-T6, GITAM School of CSE'
prs.core_properties.last_modified_by = 'Team DS-A-T6'
prs.core_properties.subject = 'PROJ2999 Capstone Project Introduction'
prs.save(OUT)
print('saved', OUT, len(prs.slides), 'slides')
