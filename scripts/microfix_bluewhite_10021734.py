from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor

ROOT = Path('决赛发布材料/1002PPT决赛完善版')
SRC = ROOT / '智烤管家_10021728_创客25逐图学习蓝白标题终版.pptx'
OUT = ROOT / '智烤管家_10021734_创客25逐图学习蓝白视觉复核终稿.pptx'
prs = Presentation(SRC)

# S14: keep ≈190万 as a single, clean KPI line.
s14 = prs.slides[13]
for sh in s14.shapes:
    if getattr(sh, 'has_text_frame', False) and sh.text.strip() == '≈190万':
        sh.text_frame.margin_left = Inches(0)
        sh.text_frame.margin_right = Inches(0)
        sh.width = Inches(1.28)
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(17.0)
                r.font.color.rgb = RGBColor(255,255,255)

# S15: four closing tags stay on one line; remove overlapping duplicate '谢谢大家'.
s15 = prs.slides[14]
for sh in list(s15.shapes):
    if not getattr(sh, 'has_text_frame', False):
        continue
    txt = sh.text.strip()
    if txt in {'一库清底','双模决策','三智落地','四队长效'}:
        sh.width = Inches(0.82)
        sh.text_frame.margin_left = Inches(0)
        sh.text_frame.margin_right = Inches(0)
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(9.4)
                r.font.color.rgb = RGBColor(244,251,255)
    elif txt == '谢谢大家':
        el = sh._element
        el.getparent().remove(el)

prs.save(OUT)
print(OUT)
