"""Re-creates the GITAM 'Poster Template.pptx' (10.5 x 14 in portrait) from the geometry recorded in the
previous session (shape names, ids, offsets, fills, fonts), so poster.py can fill it exactly as before."""
import sys
from io import BytesIO

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.util import Emu

SP, OUT = sys.argv[1], sys.argv[2]
NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
IN = 914400
BR, GP, GD, GREY = 'A58255', '007069', '004740', '7F7F7F'

prs = Presentation()
prs.slide_width, prs.slide_height = Emu(int(10.5 * IN)), Emu(14 * IN)
s = prs.slides.add_slide(prs.slide_layouts[6])
tree = s.shapes._spTree


def add(xml):
    tree.append(etree.fromstring(xml))


def inter_run(text, size, color, bold=False):
    b = ' b="1"' if bold else ''
    return (f'<a:r><a:rPr lang="en-US" sz="{size}"{b} dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Inter"/><a:ea typeface="Inter"/><a:cs typeface="Inter"/><a:sym typeface="Inter"/>'
            f'</a:rPr><a:t>{text}</a:t></a:r>')


def gshape(sid, name, x, y, cx, cy, runs, txbox=True):
    tb = ' txBox="1"' if txbox else ''
    add(f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr{tb}/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr spcFirstLastPara="1" wrap="square" lIns="71997" tIns="35989" rIns="71997" bIns="35989" '
        f'anchor="t" anchorCtr="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p>{runs}'
        f'<a:endParaRPr sz="1103" dirty="0"><a:latin typeface="Inter"/></a:endParaRPr></a:p></p:txBody></p:sp>')


ARIAL = ('<a:latin typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>'
         '<a:cs typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>')


def head(sid, name, x_in, y_in, label):
    add(f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(x_in * IN)}" y="{int(y_in * IN)}"/><a:ext cx="4320000" cy="307777"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{BR}"/></a:solidFill>'
        f'<a:ln><a:solidFill><a:srgbClr val="{BR}"/></a:solidFill></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr rtl="0">'
        f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr>'
        f'<a:r><a:rPr lang="en-IN" sz="1400" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill>{ARIAL}'
        f'</a:rPr><a:t>{label}</a:t></a:r></a:p></p:txBody></p:sp>')


def body(sid, name, x_in, y_in, h_in, lines):
    paras = ''.join(f'<a:p><a:r><a:rPr lang="en-US" sz="1200" dirty="0">{ARIAL}</a:rPr><a:t>{t}</a:t></a:r></a:p>'
                    for t in lines)
    add(f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(x_in * IN)}" y="{int(y_in * IN)}"/><a:ext cx="4320000" cy="{int(h_in * IN)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
        f'<a:ln><a:solidFill><a:srgbClr val="{BR}"/></a:solidFill></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>{paras}</p:txBody></p:sp>')


# header: title, two brown rules (group), presenter + guide names, footer, logo
gshape(36, 'Google Shape;36;p1', 757128, 443273, 6670714, 654379,
       inter_run('Title of the ', 3780, GP, True) + inter_run('Project', 3780, BR, True), txbox=False)
rule = ('<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Google Shape;{i};p1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
        '<a:xfrm><a:off x="894442" y="{y}"/><a:ext cx="7570108" cy="45719"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/>'
        '</a:prstGeom><a:solidFill><a:srgbClr val="A58255"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>')
add(f'<p:grpSp {NS}><p:nvGrpSpPr><p:cNvPr id="37" name="Google Shape;37;p1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    f'<p:grpSpPr><a:xfrm><a:off x="480598" y="356798"/><a:ext cx="6540488" cy="740854"/><a:chOff x="894442" y="2675335"/>'
    f'<a:chExt cx="7570108" cy="940767"/></a:xfrm></p:grpSpPr>{rule.format(i=38, y=2675335)}{rule.format(i=39, y=3570383)}'
    f'</p:grpSp>')
gshape(40, 'Google Shape;40;p1', 480598, 1467476, 3194752, 363530, inter_run('Presenter Names', 1890, GD))
gshape(2, 'Google Shape;40;p1', 480598, int(2.0 * IN), 3194752, 363530, inter_run('Guide Name', 1890, GD))
gshape(10, 'Google Shape;10;p6', 320511, 12444771, 3771974, 290882, inter_run('School or Dept. Name here', 1418, GREY))

# logo: GITAM crest + wordmark, cut from the official school logo shipped with the Review-2 template
logo = Image.open(f'{SP}/latex_zip/images/logo/school_logo.png').convert('RGBA')
crest = logo.crop((0, 0, 1125, logo.height))
W_, H_ = crest.width, int(round(crest.width * 708750 / 1653751))
canvas = Image.new('RGBA', (W_, H_), (0, 0, 0, 0))
canvas.paste(crest, (0, (H_ - crest.height) // 2))
buf = BytesIO()
canvas.save(buf, 'PNG')
buf.seek(0)
pic = s.shapes.add_picture(buf, Emu(7535012), Emu(356829), Emu(1653751), Emu(708750))
pic._element.nvPicPr.cNvPr.set('id', '11')
pic._element.nvPicPr.cNvPr.set('name', 'Google Shape;11;p6')

# section heads and bodies (names / ids exactly as in the GITAM template)
head(4, 'TextBox 3', 0.53, 2.61, 'Introduction')
head(5, 'TextBox 4', 0.53, 6.22, 'Literature Review')
head(6, 'TextBox 5', 0.53, 9.84, 'Research Gaps')
head(7, 'TextBox 6', 5.36, 2.61, 'Proposed Methodology')
head(8, 'TextBox 7', 5.36, 7.68, 'Tools/ Technology / Algorithms')
head(9, 'TextBox 8', 5.36, 11.27, 'References')
head(3, 'TextBox 2', 0.53, 11.6, 'Project Objectives')
body(13, 'TextBox 12', 0.53, 2.95, 3.07, ['Introduction to the project and problem statement here…'])
body(14, 'TextBox 13', 0.53, 6.56, 3.06, ['Comprehensive review of the related work.', 'Summary table or reference wise explanation'])
body(15, 'TextBox 14', 0.53, 10.17, 1.27, ['List of research gaps identified from the literature survey'])
body(16, 'TextBox 15', 5.36, 2.92, 4.59, ['Design of a well researched methodology framework for achieving the objectives.',
                                         'Include DFD/ER diagrams also.'])
body(17, 'TextBox 16', 5.36, 8.02, 3.04, ['Detailed study of the existing tools, technology and algorithms for the proposed solution/method.'])
body(18, 'TextBox 17', 5.36, 11.6, 1.92, ['Quality references used in the study in IEEE format.'])
body(12, 'TextBox 11', 0.53, 11.93, 1.59, ['Clear, specific and achievable objectives aligned with problem (1 or 2 objectives)'])

prs.save(OUT)
print('saved', OUT, len(s.shapes), 'shapes')
