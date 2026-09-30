from docx import Document
p="读取资料/3之前项目官方叙述过的/附件2：宜昌市局+智烤管家-基于健康评估的烤房用投管一体化解决方案+数创成果报告.docx"
d=Document(p)
for i,p in enumerate(d.paragraphs):
    if p.text.strip() and i>25:
        vals=[]
        for r in p.runs[:8]:
            if r.text:
                vals.append((r.text[:20],r.font.name,r.font.size.pt if r.font.size else None,r.bold))
        print("P",i,repr(p.text[:60]),"style",p.style.name,"runs",vals)
        if i>45: break
