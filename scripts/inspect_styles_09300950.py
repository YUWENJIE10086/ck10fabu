from docx import Document
p="读取资料/3之前项目官方叙述过的/附件2：宜昌市局+智烤管家-基于健康评估的烤房用投管一体化解决方案+数创成果报告.docx"
d=Document(p)
for nm in ["Normal","Heading 2","List Paragraph"]:
 s=d.styles[nm]
 print(nm,"font",s.font.name,"size",s.font.size.pt if s.font.size else None,"bold",s.font.bold)
 pf=s.paragraph_format
 print(" pf line",pf.line_spacing,"after",pf.space_after.pt if pf.space_after else None,"first",pf.first_line_indent.cm if pf.first_line_indent else None)
