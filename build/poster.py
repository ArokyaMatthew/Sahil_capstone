"""Fills the GITAM poster template (10.5 x 14 in, portrait) for the brain-tumor MRI capstone."""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pp import *  # noqa: E402,F401
from pptx import Presentation  # noqa: E402

SP, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(f'{SP}/inputs/Poster Template.pptx')
s = prs.slides[0]
GP, GD, BR = '007069', '004740', 'A58255'      # poster template palette
EMU = 914400
F = 'Arial'
BODY = 12


def shape(name, sid=None):
    for sh in s.shapes:
        if sh.name == name and (sid is None or sh.shape_id == sid):
            return sh
    raise KeyError(name)


def place(sh, x, y, w, h=None):
    sh.left, sh.top, sh.width = Inches(x), Inches(y), Inches(w)
    if h is not None:
        sh.height = Inches(h)


def refill(sh, paras, size=BODY, color='000000', bold=False, align='l', space_after=0, anchor=None):
    tf = sh.text_frame
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]._p
    for child in list(p0):
        p0.remove(child)
    if anchor:
        tf.vertical_anchor = ANCHOR[anchor]
    fill_tf(tf, paras if isinstance(paras, list) else [paras], size, color, bold, False, align, F, space_after,
            bullet_color=BR)


def fit(paras, w, size=BODY, gap_pt=0, slack=1.04):
    """Estimated height (in) of paragraphs in a box of width w (in) incl. default insets."""
    lines = sum(n_lines(p if isinstance(p, str) else p['text'], w - 0.2 - (0.22 if isinstance(p, dict) and p.get('bullet') else 0),
                        p.get('size', size) if isinstance(p, dict) else size, F, slack=slack) for p in paras)
    return lines * size * 1.17 / 72 + 0.12 + gap_pt * len(paras) / 72


# ------------------------------------------------------------------ header
title = shape('Google Shape;36;p1')
place(title, 0.53, 0.47, 7.3, 0.92)
refill(title, [{'text': 'A Deep Learning Approach for', 'color': GP},
               {'text': 'Brain Tumor Classification Using MRI Images', 'color': BR}], size=23, bold=True)
grp = shape('Google Shape;37;p1')
xfrm = grp._element.grpSpPr.find(qn('a:xfrm'))
off, ext, choff, chext = (xfrm.find(qn(t)) for t in ('a:off', 'a:ext', 'a:chOff', 'a:chExt'))
sy = int(ext.get('cy')) / int(chext.get('cy'))
bottom_rect = grp.shapes[1]._element.spPr.find(qn('a:xfrm')).find(qn('a:off'))
new_child_y = int(choff.get('y')) + int((1.43 * EMU - int(off.get('y'))) / sy)
bottom_rect.set('y', str(new_child_y))
new_chcy = new_child_y + 45719 - int(choff.get('y'))
chext.set('cy', str(new_chcy))
ext.set('cy', str(int(new_chcy * sy)))

names = shape('Google Shape;40;p1', 40)
place(names, 0.53, 1.56, 9.6)
refill(names, 'B Chakardhar  ·  Sahil Purkayastha  ·  P. Nikhil Mohan  ·  G Yaswanth', size=16, color=GD)
guide = shape('Google Shape;40;p1', 2)
place(guide, 0.53, 1.9, 9.6)
refill(guide, f'Guide: **Dr. P. Sreekanth** {{{{c:7F7F7F|Team DS-A-T6  ·  Review 2, PROJ2999 Capstone Project Introduction}}}}',
       size=14, color=GD)
foot = shape('Google Shape;10;p6')
place(foot, 0.53, 13.6, 9.6)
refill(foot, 'Department of Artificial Intelligence and Data Science  ·  GITAM School of CSE  ·  GITAM (Deemed to be University), Visakhapatnam',
       size=11, color='7F7F7F')

# ------------------------------------------------------------------ section frames
LX, RX, W = 0.53, 5.36, 4.72
HEAD_H = 0.31
TOP = 2.34


def section(head_name, body_name, label, x, y, body_h):
    hd = shape(head_name)
    place(hd, x, y, W)
    refill(hd, label, size=14, color=WHITE, bold=True)
    bd = shape(body_name)
    place(bd, x, y + HEAD_H, W, body_h)
    refill(bd, '')
    bd.text_frame.word_wrap = True
    return bd, (x, y + HEAD_H, W, body_h)


# --- left column
intro = [
    {'text': 'Early, accurate typing of a brain tumor shapes its treatment. MRI is the main imaging tool, but reading '
             'slices by hand is slow, needs scarce expertise and varies between readers.', 'space_after': 5},
    {'text': 'Deep CNNs report 97–99 % accuracy on public MRI benchmarks [1], [2], yet nearly every figure comes '
             'from one dataset and is given as one overall number.', 'space_after': 5},
    {'text': f'**{{{{c:{GP}|Problem statement:}}}}** for four-class MRI classification (glioma, meningioma, pituitary, '
             'no tumor), how well do such models hold up on a second dataset, which class fails most, and at what cost? '
             'This project answers all three under one documented, reproducible protocol.'},
]
gaps = [
    {'text': f'**{{{{c:{BR}|G1}}}}**  Accuracy proven on one benchmark only: 98.95 % [1] and 99.67 % [2] on Figshare, '
             'vs 86 % on 253 other images [4].', 'space_after': 4},
    {'text': f'**{{{{c:{BR}|G2}}}}**  A single overall accuracy hides weak classes and false positives.', 'space_after': 4},
    {'text': f'**{{{{c:{BR}|G3}}}}**  No common evaluation protocol; model cost and interpretability are rarely reported [2], [3].'},
]
objs = [
    {'text': f'**{{{{c:{GP}|O1}}}}**  Build an EfficientNet-B0 transfer-learning classifier [5], benchmarked against ResNet-50 '
             'and a scratch CNN, reaching **macro-F1 ≥ 0.95** on the 7,023-image test set with **< 10 M parameters** (G3).',
     'space_after': 6},
    {'text': f'**{{{{c:{GP}|O2}}}}**  Evaluate it on a second dataset (Figshare T1-CE: 3,064 slices, 233 patients, '
             'patient-level splits); report per-class precision, recall, F1, specificity and Grad-CAM maps, with '
             '**recall ≥ 0.90 for every class** (G1, G2).'},
]
h_intro = fit(intro, W, gap_pt=5)
h_gaps = fit(gaps, W, gap_pt=4)
h_objs = fit(objs, W, gap_pt=6)
LIT_H = 3.26
y = TOP
bd, _ = section('TextBox 3', 'TextBox 12', 'Introduction', LX, y, h_intro)
refill(bd, intro)
y += HEAD_H + h_intro + 0.12
bd_lit, (lx, ly, lw, lh) = section('TextBox 4', 'TextBox 13', 'Literature Review', LX, y, LIT_H)
y += HEAD_H + LIT_H + 0.12
bd, _ = section('TextBox 5', 'TextBox 14', 'Research Gaps', LX, y, h_gaps)
refill(bd, gaps)
y += HEAD_H + h_gaps + 0.12
obj_h = 13.48 - (y + HEAD_H)
assert obj_h >= h_objs, ('objectives overflow', obj_h, h_objs)
bd, _ = section('TextBox 2', 'TextBox 11', 'Project Objectives', LX, y, obj_h)
refill(bd, objs, anchor='m')
print('left column: intro %.2f gaps %.2f objs %.2f/%.2f' % (h_intro, h_gaps, h_objs, obj_h))

# literature table inside its frame
text(s.shapes, lx + 0.08, ly + 0.06, lw - 0.16, 0.26, '13 papers reviewed (2021–2026). Representative results:',
     size=BODY, font=F, anchor='m', margin=(0, 0, 0, 0))
rows = [('Ref.', 'Approach', 'Reported', 'Limitation'),
        ('[1]', 'EfficientNet + ResNet-50 feature fusion', '98.95 %|Figshare', 'one dataset only'),
        ('[2]', 'SE-ResNet101 + ensemble classifier', '99.67 %|Figshare', 'costly, hard to interpret'),
        ('[3]', 'Eight CNNs + ViT, Grad-CAM, LIME', '98.17 %|7,023 images', 'some models attend to irrelevant regions'),
        ('[4]', 'VGG16, InceptionV3, MobileNetV2', '86.00 %|253 images', 'small data: accuracy drops')]
cw_ = [0.44, 1.6, 1.08, 1.44]
row_h = []
for r, row in enumerate(rows):
    hmax = 0
    for c, val in enumerate(row):
        nl = sum(n_lines(v, cw_[c] - 0.1, BODY, F, bold=(r == 0 or c == 0)) for v in val.split('|'))
        hmax = max(hmax, nl * BODY * 1.17 / 72 + 0.08)
    row_h.append(hmax)
gf = s.shapes.add_table(len(rows), 4, Inches(lx + 0.08), Inches(ly + 0.36), Inches(sum(cw_)), Inches(sum(row_h)))
plain_table(gf)
for c, w_ in enumerate(cw_):
    gf.table.columns[c].width = Inches(w_)
for r, row in enumerate(rows):
    gf.table.rows[r].height = Inches(row_h[r])
    for c, val in enumerate(row):
        cell = gf.table.cell(r, c)
        cell_style(cell, BR if r == 0 else (WHITE if r % 2 else 'F6F0EA'), border='D9C6B2', margin=(0.04, 0.03, 0.04, 0.03))
        cell_text(cell, val.split('|'), size=BODY, font=F, bold=(r == 0 or c == 0), color=WHITE if r == 0 else ('000000' if c else GP),
                  align='c' if c in (0, 2) else 'l')
print('lit table height %.2f, frame %.2f' % (0.36 + sum(row_h) + 0.08, lh))
assert 0.36 + sum(row_h) + 0.06 <= lh

# --- right column
tools = [
    {'text': f'**{{{{c:{GP}|Tools:}}}}** Python, PyTorch + torchvision, scikit-learn, OpenCV, Albumentations, '
             'pytorch-grad-cam, Streamlit, GitHub.', 'space_after': 4},
    {'text': f'**{{{{c:{GP}|Technology:}}}}** ImageNet transfer learning; Grad-CAM explainability; Colab / Kaggle T4 GPU.',
     'space_after': 4},
    {'text': f'**{{{{c:{GP}|Algorithms:}}}}** EfficientNet-B0 [5] (5.3 M params, primary), ResNet-50 (25.6 M, baseline), '
             'scratch CNN (ablation); AdamW, cosine LR, class-weighted cross-entropy.'},
]
refs = [
    'M. Aamir et al., “A deep learning approach for brain tumor classification using MRI images,” *Comput. Electr. Eng.*, vol. 101, 2022, Art. no. 108105.',
    'M. Aamir et al., “An automated deep learning framework for brain tumor classification using MRI imagery,” *Sci. Rep.*, vol. 15, 2025, Art. no. 17593.',
    'E. F. Gomes and R. S. Barbosa, “Deep learning approaches for brain tumor classification in MRI scans: An analysis of model interpretability,” *Appl. Sci.*, vol. 16, no. 2, 2026, Art. no. 831.',
    'Z. N. I. Zailan et al., “Deep learning approach for prediction of brain tumor from small number of MRI images,” *JOIV*, vol. 6, no. 2-2, pp. 581–586, 2022.',
    'M. Tan and Q. V. Le, “EfficientNet: Rethinking model scaling for convolutional neural networks,” in *Proc. ICML*, 2019, pp. 6105–6114.',
]
h_tools = fit(tools, W, gap_pt=4, slack=1.0)
h_refs = sum(n_lines(f'[{i + 1}] ' + r.replace('*', ''), W - 0.2 - 0.32, BODY, F) for i, r in enumerate(refs)) * BODY * 1.17 / 72 + 0.12 + 4 * 3 / 72
y = TOP
meth_h = 13.48 - TOP - 3 * HEAD_H - 2 * 0.12 - h_tools - h_refs
bd_m, (mx, my, mw, mh) = section('TextBox 6', 'TextBox 15', 'Proposed Methodology', RX, y, meth_h)
y += HEAD_H + meth_h + 0.12
bd, _ = section('TextBox 7', 'TextBox 16', 'Tools/ Technology / Algorithms', RX, y, h_tools)
refill(bd, tools)
y += HEAD_H + h_tools + 0.12
bd_r, _ = section('TextBox 8', 'TextBox 17', 'References', RX, y, h_refs)
tf = bd_r.text_frame
for k, r in enumerate(refs):
    p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
    if k:
        p.space_before = Pt(3)
    pPr = p._p.get_or_add_pPr()
    pPr.set('marL', str(int(Inches(0.32))))
    pPr.set('indent', str(-int(Inches(0.32))))
    for part in re.split(r'(\*[^*]+\*)', f'[{k + 1}]\t' + r):
        if not part:
            continue
        run = p.add_run()
        it = part.startswith('*')
        run.text = part.strip('*') if it else part
        run.font.size, run.font.name, run.font.italic = Pt(BODY), F, it
print('right column: meth %.2f tools %.2f refs %.2f' % (meth_h, h_tools, h_refs))

# methodology: numbered pipeline + level-0 DFD
steps = [('Input', '7,023 images (4 classes) + 3,064 slices'),
         ('Curation', 'pHash de-duplication, frozen splits'),
         ('Preprocess', 'brain crop, 224 × 224, augmentation'),
         ('Training', 'EfficientNet-B0 vs ResNet-50 vs CNN'),
         ('Evaluation', 'per-class P / R / F1, ROC, cost'),
         ('Output', 'class, confidence, Grad-CAM heatmap')]
cx0 = mx + 0.3
ty = my + 0.12
row_tops = []
for k, (head, body_) in enumerate(steps):
    t = f'**{head}** — {body_}'
    h = max(0.36, n_lines(t, mw - 0.82, BODY, F, slack=1.0) * BODY * 1.17 / 72 + 0.06)
    row_tops.append((ty, h))
    ty += h + 0.07
arrow(s.shapes, cx0, row_tops[0][0] + 0.18, cx0, row_tops[-1][0] + 0.18, color='C9B49C', width=2.5, head=None)
for k, ((head, body_), (ty, h)) in enumerate(zip(steps, row_tops)):
    d = box(s.shapes, cx0 - 0.16, ty + 0.01, 0.32, 0.32, GP if k not in (0, 5) else BR, kind=MSO_SHAPE.OVAL)
    shape_text(d, str(k + 1), size=BODY, bold=True, color=WHITE, font=F, margin=(0, 0, 0, 0))
    text(s.shapes, cx0 + 0.3, ty, mw - 0.72, h, f'**{{{{c:{GP}|{head}}}}}** — {body_}', size=BODY, font=F,
         margin=(0, 0.02, 0, 0), anchor='t')
cap_y = ty + h + 0.04
text(s.shapes, mx + 0.1, cap_y, mw - 0.2, 0.24, 'Fig. 1: Proposed methodology, input to explained output', size=BODY,
     italic=True, color='555555', font=F, align='c', anchor='m', margin=(0, 0, 0, 0))
# DFD level 0
dy = cap_y + 0.34
sep = arrow(s.shapes, mx + 0.25, dy - 0.1, mx + mw - 0.25, dy - 0.1, color='D9C6B2', width=1, head=None)
cxm = mx + mw / 2
ent_w, ent_h, pw, ph = 0.98, 0.6, 1.3, 0.78
dsb = box(s.shapes, cxm - 0.95, dy + 0.02, 1.9, 0.42, BR, shade=False)
shape_text(dsb, 'Public MRI datasets', size=BODY, bold=True, color=WHITE, font=F, margin=(0.02, 0, 0.02, 0))
py = dy + 0.74
pr = box(s.shapes, cxm - pw / 2, py, pw, ph, WHITE, GP, 1.75, kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
band = top_round(s.shapes, cxm - pw / 2, py, pw, 0.24, GP, radius=0.1)
shape_text(band, '0', size=BODY, bold=True, color=WHITE, font=F, margin=(0, 0, 0, 0))
text(s.shapes, cxm - pw / 2, py + 0.24, pw, ph - 0.26, 'Classifier system', size=BODY, bold=True, font=F, align='c',
     anchor='m', margin=(0.03, 0, 0.03, 0))
arrow(s.shapes, cxm, dy + 0.44, cxm, py, color='333333', width=1.5)
text(s.shapes, cxm + 0.06, dy + 0.44, 1.2, 0.26, 'images', size=BODY, italic=True, color='555555', font=F, anchor='m',
     margin=(0, 0, 0, 0))
ey = py + ph / 2
for side, label, l_in, l_out in [(-1, 'End user', 'slice', 'result'), (1, 'Research team', 'config', 'report')]:
    ex = mx + 0.12 if side < 0 else mx + mw - 0.12 - ent_w
    e = box(s.shapes, ex, ey - ent_h / 2, ent_w, ent_h, BR)
    shape_text(e, label, size=BODY, bold=True, color=WHITE, font=F, margin=(0.02, 0, 0.02, 0))
    near = ex + ent_w if side < 0 else ex
    far = cxm - pw / 2 if side < 0 else cxm + pw / 2
    arrow(s.shapes, near, ey - 0.15, far, ey - 0.15, color='333333', width=1.5)
    arrow(s.shapes, far, ey + 0.15, near, ey + 0.15, color='333333', width=1.5)
    lx0, lw0 = min(near, far), abs(far - near)
    text(s.shapes, lx0, ey - 0.42, lw0, 0.24, l_in, size=BODY, italic=True, color='555555', font=F, align='c', anchor='m', margin=(0, 0, 0, 0))
    text(s.shapes, lx0, ey + 0.18, lw0, 0.24, l_out, size=BODY, italic=True, color='555555', font=F, align='c', anchor='m', margin=(0, 0, 0, 0))
cap2 = py + ph + 0.08
text(s.shapes, mx + 0.1, cap2, mw - 0.2, 0.24, 'Fig. 2: Level-0 DFD (Gane–Sarson notation)', size=BODY, italic=True,
     color='555555', font=F, align='c', anchor='m', margin=(0, 0, 0, 0))
print('methodology content ends at %.2f, frame ends at %.2f' % (cap2 + 0.24, my + mh))
assert cap2 + 0.24 <= my + mh + 0.02, "methodology overflow"

prs.core_properties.title = 'Poster — A Deep Learning Approach for Brain Tumor Classification Using MRI Images'
prs.core_properties.author = 'Team DS-A-T6, GITAM School of CSE'
prs.core_properties.last_modified_by = 'Team DS-A-T6'
prs.save(OUT)
print('saved', OUT)
