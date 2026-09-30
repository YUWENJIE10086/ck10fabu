from docx import Document
from pathlib import Path
p=Path("读取资料/3之前项目官方叙述过的/附件2：宜昌市局+智烤管家-基于健康评估的烤房用投管一体化解决方案+数创成果报告.docx")
d=Document(p)
for i,s in enumerate(d.sections):
 print(i, "top", s.top_margin.cm, "bottom", s.bottom_margin.cm, "left", s.left_margin.cm, "right", s.right_margin.cm, "header", s.header_distance.cm, "footer", s.footer_distance.cm)
for i,p in enumerate(d.paragraphs[:30]):
 if p.text.strip():
  pf=p.paragraph_format
  print("P",i,repr(p.text[:50]),"align",p.alignment,"left",pf.left_indent.cm if pf.left_indent else None,"right",pf.right_indent.cm if pf.right_indent else None,"first",pf.first_line_indent.cm if pf.first_line_indent else None,"line",pf.line_spacing)
