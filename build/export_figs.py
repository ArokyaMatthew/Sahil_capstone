"""Exports the five design diagrams (slides 14, 16, 18, 19, 20 of the Review-2 deck) as 300-dpi PNGs for the report."""
import os
import subprocess
import sys

from PIL import Image
from pptx import Presentation

SP = sys.argv[1]
deck = f'{SP}/out/review2.pptx'
fig = f'{SP}/figs'
os.makedirs(fig, exist_ok=True)
KEEP = [14, 16, 18, 19, 20]
CROPS = [('fig_method', 0.22, 0.92, 13.1, 6.52), ('fig_workflow', 0.25, 0.9, 13.1, 6.42),
         ('fig_dfd0', 0.4, 0.92, 12.95, 6.0), ('fig_dfd1', 0.25, 0.95, 12.45, 6.38),
         ('fig_usecase', 0.4, 0.92, 12.6, 6.45)]

p = Presentation(deck)
for n, sl in enumerate(p.slides, 1):
    if n in KEEP:  # the report has its own numbered captions, so drop the slide captions
        for sh in list(sl.shapes):
            if sh.has_text_frame and sh.text_frame.text.startswith('Figure:'):
                sh._element.getparent().remove(sh._element)
ids = p.slides._sldIdLst
for n, sid in reversed(list(enumerate(list(ids), 1))):
    if n not in KEEP:
        p.part.drop_rel(sid.rId)
        ids.remove(sid)
src = f'{fig}/figs_src.pptx'
p.save(src)
subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', src, '--outdir', fig], check=True, capture_output=True)
subprocess.run(['pdftoppm', '-png', '-r', '300', f'{fig}/figs_src.pdf', f'{fig}/sl'], check=True)
for k, (name, x0, y0, x1, y1) in enumerate(CROPS, 1):
    im = Image.open(f'{fig}/sl-{k}.png').convert('RGB')
    im.crop((int(x0 * 300), int(y0 * 300), int(x1 * 300), int(y1 * 300))).save(f'{fig}/{name}.png', dpi=(300, 300))
print('figures written to', fig)
