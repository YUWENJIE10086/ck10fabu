from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from docx import Document
from docx.shared import Pt as DPt, Inches as DInches, RGBColor as DRGB
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement as DOxmlElement
from docx.oxml.ns import qn as dqn
from PIL import Image
import numpy as np, cv2, io

STAMP='10031742'
ROOT=Path('.')
DEST=ROOT/'决赛发布材料/1002PPT决赛完善版/1003'
DEST.mkdir(parents=True,exist_ok=True)
COVER=DEST/'智烤管家_142106_1比1蓝白山海光路抠图还原版_副本3.pptx'
PPT=DEST/f'智烤管家_{STAMP}_8分钟决赛蓝白黄标全组件可编辑重设计版.pptx'
DOC=DEST/f'智烤管家_{STAMP}_8分钟逐页演讲稿与PPT设计说明.docx'
TMP=Path('/tmp')/f'zhikao_{STAMP}'; TMP.mkdir(exist_ok=True)
NAVY='061A33'; NAVY2='0A2747'; BLUE='1C7CC4'; CYAN='42D2FF'; ICE='DFF8FF'; WHITE='F7FBFF'; MUTED='A9C5D9'; YELLOW='FFD34E'; FONT_CN='Microsoft YaHei'; FONT_EN='Arial'; W=13.333; H=7.5

def rgb(h): h=h.lstrip('#'); return RGBColor(int(h[:2],16),int(h[2:4],16),int(h[4:6],16))
def set_trans(fill,t=0):
    sf=fill._xPr.find(qn('a:solidFill'))
    if sf is not None and len(sf):
        c=sf[0]
        for a in list(c.findall(qn('a:alpha'))): c.remove(a)
        a=OxmlElement('a:alpha'); a.set('val',str(int((100-t)*1000))); c.append(a)
def set_ea(r):
    r.font.name=FONT_EN; p=r._r.get_or_add_rPr(); ea=p.find(qn('a:ea'))
    if ea is None: ea=OxmlElement('a:ea'); p.append(ea)
    ea.set('typeface',FONT_CN); la=p.find(qn('a:latin'))
    if la is None: la=OxmlElement('a:latin'); p.append(la)
    la.set('typeface',FONT_EN)
def grad(r):
    p=r._r.get_or_add_rPr()
    for tag in ('a:solidFill','a:gradFill','a:noFill','a:pattFill'):
        for e in list(p.findall(qn(tag))): p.remove(e)
    g=OxmlElement('a:gradFill'); g.set('rotWithShape','1'); l=OxmlElement('a:gsLst')
    for pos,col in [(0,WHITE),(52000,'CFF7FF'),(100000,CYAN)]:
        gs=OxmlElement('a:gs'); gs.set('pos',str(pos)); c=OxmlElement('a:srgbClr'); c.set('val',col); gs.append(c); l.append(gs)
    g.append(l); ln=OxmlElement('a:lin'); ln.set('ang','5400000'); ln.set('scaled','1'); g.append(ln); p.insert(0,g)
def solid(r,col):
    p=r._r.get_or_add_rPr()
    for tag in ('a:solidFill','a:gradFill','a:noFill','a:pattFill'):
        for e in list(p.findall(qn(tag))): p.remove(e)
    sf=OxmlElement('a:solidFill'); c=OxmlElement('a:srgbClr'); c.set('val',col); sf.append(c); p.insert(0,sf)
def text(slide,t,x,y,w,h,sz=16,col=WHITE,b=False,al=PP_ALIGN.LEFT,g=False,name=None):
    s=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));
    if name:s.name=name
    f=s.text_frame; f.clear(); f.word_wrap=True; f.vertical_anchor=MSO_ANCHOR.MIDDLE; f.margin_left=f.margin_right=Inches(.02); f.margin_top=f.margin_bottom=0
    p=f.paragraphs[0]; p.alignment=al; r=p.add_run(); r.text=t; set_ea(r); r.font.size=Pt(sz); r.font.bold=b; grad(r) if g else solid(r,col); return s
def rich(slide,parts,x,y,w,h,sz=15,al=PP_ALIGN.CENTER):
    s=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); f=s.text_frame; f.clear(); f.vertical_anchor=MSO_ANCHOR.MIDDLE; p=f.paragraphs[0]; p.alignment=al
    for tx,co,bo in parts: r=p.add_run(); r.text=tx; set_ea(r); r.font.size=Pt(sz); r.font.bold=bo; solid(r,co)
    return s
def rect(slide,x,y,w,h,fill=NAVY2,line=BLUE,rad=True,tr=0,lw=1.1):
    typ=MSO_SHAPE.ROUNDED_RECTANGLE if rad else MSO_SHAPE.RECTANGLE; s=slide.shapes.add_shape(typ,Inches(x),Inches(y),Inches(w),Inches(h)); s.fill.solid(); s.fill.fore_color.rgb=rgb(fill); set_trans(s.fill,tr)
    if line:s.line.color.rgb=rgb(line); s.line.width=Pt(lw)
    else:s.line.fill.background()
    return s
def line(slide,x1,y1,x2,y2,col=BLUE,w=1.0):
    s=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x1),Inches(y1),Inches(x2),Inches(y2)); s.line.color.rgb=rgb(col); s.line.width=Pt(w); return s
def circle(slide,x,y,d,fill=NAVY2,linec=CYAN,tr=0,lw=1):
    s=slide.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(d),Inches(d)); s.fill.solid(); s.fill.fore_color.rgb=rgb(fill); set_trans(s.fill,tr)
    if linec:s.line.color.rgb=rgb(linec); s.line.width=Pt(lw)
    else:s.line.fill.background()
    return s
def tag(slide,t,x,y,w,fill=YELLOW,tc=NAVY): rect(slide,x,y,w,.36,fill,None,True,0); text(slide,t,x+.04,y,w-.08,.35,11.2,tc,True,PP_ALIGN.CENTER)
def arrow(slide,x1,y1,x2,y2): s=line(slide,x1,y1,x2,y2,CYAN,1.7); s.line.end_arrowhead=True
def bg(slide):
    rect(slide,0,0,W,H,NAVY,None,False,0); rect(slide,0,0,W,1.36,'09233F',None,False,8)
    for x in [0.9,2.4,3.9,5.4,6.9,8.4,9.9,11.4,12.9]: line(slide,x,1.48,x,7.15,'153F61',.35)
    for y in [2,3,4,5,6,7]: line(slide,.55,y,12.8,y,'153F61',.3)
def title(slide,t,sub,ch):
    tag(slide,ch,.55,.34,.95); text(slide,t,2.12,.28,10.45,.62,24,WHITE,True,PP_ALIGN.LEFT,True,'页标题_可编辑'); text(slide,sub,.74,.92,11.9,.38,12.2,MUTED,False); line(slide,.72,1.3,12.62,1.3,'2A6D99',.8)
def kpi(slide,n,l,x,y,w=2.5): text(slide,n,x,y,w,.66,29,YELLOW,True,PP_ALIGN.CENTER); text(slide,l,x,y+.62,w,.4,12,WHITE,True,PP_ALIGN.CENTER)
def home(slide,cx,cy,s=.62):
    r=slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,Inches(cx-s*.5),Inches(cy-s*.52),Inches(s),Inches(s*.52)); r.fill.solid(); r.fill.fore_color.rgb=rgb(CYAN); r.line.fill.background(); rect(slide,cx-s*.38,cy-s*.10,s*.76,s*.55,'0F466E',CYAN,False,0,1)
def wrench(slide,cx,cy,s=.62): line(slide,cx-s*.3,cy+s*.28,cx+s*.24,cy-s*.26,CYAN,5); circle(slide,cx+s*.14,cy-s*.38,s*.30,NAVY,CYAN,0,4); circle(slide,cx-s*.39,cy+s*.17,s*.19,NAVY,CYAN,0,3)
def bars(slide,cx,cy,s=.65):
    for i,v in enumerate([.35,.62,.92]): rect(slide,cx-s*.45+i*s*.3,cy+s*.35-v*s*.62,s*.18,v*s*.62,CYAN,None,False)

def clean_cover():
    p=Presentation(COVER); sl=p.slides[0]; im=Image.open(io.BytesIO(sl.shapes[0].image.blob)).convert('RGB'); a=np.array(im); Hh,Ww=a.shape[:2]; sx=Ww/(p.slide_width/914400); sy=Hh/(p.slide_height/914400); m=np.zeros((Hh,Ww),np.uint8)
    for sh in list(sl.shapes)[1:9]:
        x=int(sh.left/914400*sx); y=int(sh.top/914400*sy); w=int(sh.width/914400*sx); h=int(sh.height/914400*sy); cv2.rectangle(m,(max(0,x-18),max(0,y-18)),(min(Ww-1,x+w+18),min(Hh-1,y+h+18)),255,-1)
    cv2.rectangle(m,(int(Ww*.8),int(Hh*.88)),(Ww-1,Hh-1),255,-1); out=cv2.inpaint(cv2.cvtColor(a,cv2.COLOR_RGB2BGR),m,7,cv2.INPAINT_TELEA); out=cv2.cvtColor(out,cv2.COLOR_BGR2RGB); q=TMP/'cover.png'; Image.fromarray(out).save(q); return q

prs=Presentation(); prs.slide_width=Inches(W); prs.slide_height=Inches(H); blank=prs.slide_layouts[6]
# 1 cover
s=prs.slides.add_slide(blank); s.shapes.add_picture(str(clean_cover()),0,0,width=prs.slide_width,height=prs.slide_height); text(s,'Z H I   K A O   M A N A G E R',4.36,.89,4.65,.34,13,ICE,False,PP_ALIGN.CENTER); text(s,'智烤 管家',3.22,1.47,6.2,1.22,56,WHITE,True,PP_ALIGN.CENTER,True); text(s,'让每一座烤房，都有一份可计算的“健康档案”',3.68,2.92,5.96,.42,14.2,WHITE,False,PP_ALIGN.CENTER); rich(s,[('一库清底',WHITE,True),('   |   ',CYAN,False),('双模决策',WHITE,True),('   |   ',CYAN,False),('三智落地',WHITE,True),('   |   ',CYAN,False),('四队长效',WHITE,True)],3.15,3.8,7.05,.44,13.3); text(s,'基于健康评估的烤房用管管一体化解决方案',3.62,4.45,6.25,.44,13.2,WHITE,True,PP_ALIGN.CENTER,True); text(s,'湖北烟草创客大赛  ·  决赛发布',3.4,6.59,2.95,.3,10.2,ICE,False,PP_ALIGN.CENTER); text(s,'宜昌市烟草公司  ·  智同道合创客小组',6.18,6.59,3.52,.3,10.2,ICE,False,PP_ALIGN.CENTER); text(s,'2026.10',9.85,6.59,.9,.3,10.2,ICE,False,PP_ALIGN.CENTER)
# 2 problems
s=prs.slides.add_slide(blank); bg(s); title(s,'烟区一线，最急的是这 3 件事','用房、抢修、投入——看似三件事，本质是一条管理链没有打通','01 问题')
for i,(T,G,items,ic) in enumerate([('用房难','分不到',['旺季状态不透明','口头协调慢','紧张与闲置并存'],home),('抢修慢','修不及',['报修—找人—找件环节多','人工排查慢','停运损失大'],wrench),('投入盲','投不准',['维修、新建靠经验','供需脱节','资金容易错配'],bars)]):
    x=.72+i*4.15; rect(s,x,1.7,3.68,3.9,NAVY2,'2C7FB2',True,7); ic(s,x+.7,2.45,.66); tag(s,G,x+2.35,2.06,.92); text(s,T,x+.48,2.92,2.72,.48,23,WHITE,True,PP_ALIGN.LEFT,True)
    for j,it in enumerate(items): circle(s,x+.5,3.66+j*.55,.12,YELLOW,None); text(s,it,x+.75,3.51+j*.55,2.55,.42,12.2,ICE)
rect(s,1.35,5.92,10.65,.74,'0A3557',YELLOW,True,5,1.1); rich(s,[('根因：',YELLOW,True),('资源无统筹  ｜  投入无量化  ｜  管护无闭环',WHITE,True)],1.65,6.03,10.1,.5,15)
# 3 1234
s=prs.slides.add_slide(blank); bg(s); title(s,'一套“1234”，把三难串成一个闭环','不是再加一个系统，而是重做一条从数据到服务的管理链','02 方案')
for i,(n,T,sub) in enumerate([('1','一库清底','一房一档｜底数清'),('2','双模决策','健康评估｜发展导引'),('3','三智落地','智投｜智用｜智管'),('4','四队长效','有人管｜有钱管｜有标准')]):
    x=.72+i*3.10; rect(s,x,1.9,2.68,3.3,NAVY2,BLUE,True,8); text(s,n,x+.12,1.9,.76,.7,30,YELLOW,True); text(s,T,x+.34,2.68,2.02,.55,21,WHITE,True,PP_ALIGN.CENTER,True); text(s,sub,x+.34,3.35,2.02,.6,12.2,ICE,False,PP_ALIGN.CENTER); 
    if i<3: arrow(s,x+2.72,3.55,x+3,3.55)
rect(s,1.85,5.6,9.55,.84,'0A3557',None,True,3); rich(s,[('数据',WHITE,True),('  →  ',CYAN,True),('决策',WHITE,True),('  →  ',CYAN,True),('服务',WHITE,True),('  →  ',CYAN,True),('机制',YELLOW,True)],2.15,5.78,8.95,.46,17)
# 4 data
s=prs.slides.add_slide(blank); bg(s); title(s,'一库清底：先把“家底”变成可用数据','标准化判定 + 移动端采集 + 一房一档，为后面的模型提供同一套数据','03 数据')
for i,(a,b) in enumerate([('4大模块','加热｜散热｜自控｜附属'),('18个关键部件','统一状态口径'),('3类部件状态','正常｜损坏｜缺失'),('3档烤房评价','优良｜一般｜较差')]):
    y=1.72+i; rect(s,.78,y,4.5,.76,NAVY2,BLUE,True,4); tag(s,a,.96,y+.19,1.38 if i else 1.18); text(s,b,2.52,y+.12,2.48,.5,13,WHITE,True)
rect(s,6.1,1.72,6.02,3.72,'0B2D4D',CYAN,True,2,1.2); text(s,'一房一档',7.55,1.95,3.72,.58,23,WHITE,True,PP_ALIGN.LEFT,True)
for i,(k,v) in enumerate([('基础信息','位置 / 容量 / 建设年限'),('部件状态','18部件逐项判定'),('维修记录','工单 / 配件 / 结果'),('用户评价','烘烤质量 / 运维服务')]): text(s,k,6.72,2.88+i*.6,1.1,.38,12.4,YELLOW,True); text(s,v,7.92,2.88+i*.6,3.7,.38,12.2,ICE)
tag(s,'现场拍照录入 → 实时归档 → 全程可溯',3.55,5.88,6.2)
# 5 model1
s=prs.slides.add_slide(blank); bg(s); title(s,'核心创新①：健康评估——先知道“谁最该修”','把部件状态、管护情况、现场评价统一成一个可执行的维修优先序','04 模型1'); tag(s,'核心创新',.82,1.62,1.08); text(s,'评价依据',.88,2.14,2.2,.42,16,WHITE,True,g=True)
x=.88
for lab,val,wid,col in [('部件状态','80%',4.2,YELLOW),('管护情况','10%',.55,CYAN),('现场评价','10%',.55,ICE)]: rect(s,x,2.72,wid,.62,col,None,False); text(s,val,x,2.74,wid,.3,13,NAVY,True,PP_ALIGN.CENTER); text(s,lab,x,3.03,wid,.28,9.8,NAVY,True,PP_ALIGN.CENTER); x+=wid+.04
rect(s,6.14,2.05,2.54,2.42,NAVY2,CYAN,True,2,1.2); text(s,'AHP',6.14,2.26,2.54,.6,27,WHITE,True,PP_ALIGN.CENTER,True); text(s,'层次分析法',6.14,2.85,2.54,.36,12.5,YELLOW,True,PP_ALIGN.CENTER); text(s,'建设年限｜使用时长\n养护记录｜实时工况',6.33,3.27,2.16,.74,11.3,ICE,False,PP_ALIGN.CENTER)
rect(s,9.2,1.8,3.42,3.18,'0B2D4D',BLUE,True,2); text(s,'60分',9.46,2.1,1.42,.78,34,YELLOW,True,PP_ALIGN.CENTER); text(s,'安全预警线',10.95,2.26,1.3,.44,13,WHITE,True)
for i,T in enumerate(['健康评分','风险状态','维保优先清单']): circle(s,9.55,3.1+i*.52,.12,YELLOW,None); text(s,T,9.82,2.98+i*.52,2.1,.38,12.4,ICE)
rect(s,1.46,5.54,10.35,.82,'0A3557',None,True,3); rich(s,[('过去：',MUTED,False),('盲目排查、被动抢修',WHITE,True),('     →     ',CYAN,True),('现在：',MUTED,False),('预判风险、靶向维保',YELLOW,True)],1.8,5.72,9.7,.45,15)
# 6 model2
s=prs.slides.add_slide(blank); bg(s); title(s,'核心创新②：发展导引——先算“有效烤能”，再决定建不建','把单栋健康结果折算到村级产能，再用客观算法识别真实缺口','05 模型2'); tag(s,'原创模型',.82,1.62,1.08)
for i,(T,x) in enumerate([('单栋\n健康数据',1),('村级\n有效烤能',3.33),('产业需求\n多维指标',5.66),('熵权法\n+ TOPSIS',8),('地图\n情景沙盘',10.34)]): rect(s,x,2.25,1.82,1.52,NAVY2,CYAN if i in [1,3] else BLUE,True,4,1.1); text(s,T,x+.13,2.49,1.56,.92,16,WHITE,True,PP_ALIGN.CENTER,i in [1,3]);
for x in [2.86,5.19,7.52,9.86]: arrow(s,x,3,x+.33,3)
text(s,'种植规模｜产业趋势｜集体经济｜管护能力｜政策适配',4.42,4.12,4.9,.42,11.4,MUTED,False,PP_ALIGN.CENTER); rect(s,2.05,4.78,9.25,1.18,'0B2D4D',YELLOW,True,2,1.1); text(s,'输出',2.3,5.08,.72,.4,12,YELLOW,True,PP_ALIGN.CENTER)
for i,T in enumerate(['哪里真缺','缺多少','建哪里']): text(s,T,3.28+i*2.32,4.96,1.8,.58,18,WHITE,True,PP_ALIGN.CENTER,True)
tag(s,'把“凭感觉”变成“凭数据”',4.7,6.12,3.9)
# 7 chain
s=prs.slides.add_slide(blank); bg(s); title(s,'双模不是并列：它是一条投入决策链','同一套数据，既服务维修，也服务投资','06 决策')
for i,(a,b) in enumerate([('存量体检','模型1'),('健康折损','风险分层'),('有效烤能','能力折算'),('真实缺口','模型2'),('调剂 / 维修 / 新建','行动')]):
    x=.72+i*2.53; rect(s,x,2.24,2.05,1.66,NAVY2,BLUE,True,4); text(s,a,x+.12,2.48,1.81,.46,16.5,WHITE,True,PP_ALIGN.CENTER,i in [0,2,3]); text(s,b,x+.12,3.10,1.81,.32,10.5,YELLOW if i in [0,3,4] else MUTED,True,PP_ALIGN.CENTER); 
    if i<4: arrow(s,x+2.08,3.05,x+2.4,3.05)
rect(s,1.4,4.7,10.53,1.14,'0A3557',None,True,2); rich(s,[('原则：',MUTED,False),('先盘活存量',WHITE,True),('  →  ',CYAN,True),('再精准修复',WHITE,True),('  →  ',CYAN,True),('仍有缺口才新增',YELLOW,True)],1.72,4.98,9.9,.55,16.2)
# 8 three scenes
s=prs.slides.add_slide(blank); bg(s); title(s,'三智落地：模型算完，更要真正跑起来','不同角色看到不同动作，但数据最终回到同一个闭环','07 落地')
for i,(T,R,its) in enumerate([('智投','管理部门',['年度改造清单','新建点位推演']),('智用','烟农 / 技术员',['找房调剂','一键报修']),('智管','合作社 / 管护队',['派单处置','验收评价 / 数据回写'])]):
    x=.88+i*4.12; rect(s,x,1.86,3.55,3.74,NAVY2,BLUE,True,4); tag(s,R,x+.25,2.12,1.3 if i else 1.1,YELLOW if i==0 else '2B7EAE',NAVY if i==0 else WHITE); text(s,T,x+.34,2.75,2.86,.64,26,WHITE,True,PP_ALIGN.CENTER,True)
    for j,it in enumerate(its): circle(s,x+.62,3.78+j*.7,.14,YELLOW if j==0 else CYAN,None); text(s,it,x+.92,3.62+j*.7,2.12,.44,12.4,ICE)
text(s,'烟农  —  合作社  —  管护人员  —  管理部门',1.85,5.98,9.65,.44,14.2,WHITE,True,PP_ALIGN.CENTER); tag(s,'可选：此处接15秒真实操作演示',4.62,6.47,4.1)
# 9 smart invest
s=prs.slides.add_slide(blank); bg(s); title(s,'智投：投入有据可依，布局贴合需求','双模研判把“该修什么、该建多少、建在哪里”落到年度计划','08 智投')
for i,(a,b) in enumerate([('双模研判','健康 + 缺口'),('筛选对象','低效 / 老化 / 高危'),('形成清单','升级改造 / 新建点位'),('复盘评估','投入与使用效果')]): x=.78+i*2.5; rect(s,x,2,2.16,1.3,NAVY2,BLUE,True,4); text(s,a,x+.12,2.2,1.92,.38,13.4,WHITE,True,PP_ALIGN.CENTER); text(s,b,x+.12,2.67,1.92,.32,10.2,MUTED,False,PP_ALIGN.CENTER)
kpi(s,'150栋','核心设备升级计划',2.18,4,3.25); kpi(s,'7栋','肖家坪千亩村科学新建',7.85,4,3.25); tag(s,'钱投到最该投的地方',4.56,5.8,4.2)
# 10 smart use
s=prs.slides.add_slide(blank); bg(s); title(s,'智用：找房、报修，不再靠一圈电话','把烟农最直接的“用房难、抢修慢”变成可追踪的线上流程','09 智用'); rect(s,.78,1.8,5.28,1.7,NAVY2,BLUE,True,4); tag(s,'找房',1.02,2.06,.88); text(s,'周边烤房  →  距离 / 空闲时段 / 设备评分  →  预约调剂',1.02,2.56,4.72,.5,12.4,WHITE,True); rect(s,.78,3.72,5.28,1.7,NAVY2,BLUE,True,4); tag(s,'报修',1.02,3.98,.88); text(s,'一键报修  →  匹配人员 / 配件 / 路线  →  进度反馈',1.02,4.48,4.72,.5,12.4,WHITE,True); rect(s,6.55,1.8,5.9,3.62,'0B2D4D',YELLOW,True,2,1.2); text(s,'真实案例｜磨坪站雷击停运',6.88,2.12,5.25,.5,17,WHITE,True,PP_ALIGN.LEFT,True); kpi(s,'2处','故障部件快速定位',7.1,2.92,2.1); kpi(s,'3小时','完成抢修复产',9.68,2.92,2.2); text(s,'双向评价：烘烤质量 + 运维服务',6.88,4.56,5.1,.42,12.3,ICE,False,PP_ALIGN.CENTER); tag(s,'服务不是结束于维修，而是结束于评价回写',3.22,5.88,6.88)
# 11 management
s=prs.slides.add_slide(blank); bg(s); title(s,'智管：四支队伍，让平台从“上线”变“长效”','专业队伍 + 闭环流程 + 专项基金，解决“谁来管、钱怎么管”','10 智管')
for i,T in enumerate(['技术','抢修','养护','综合']): x=.78+i*2.05; circle(s,x,2.16,1.36,'0D3659',CYAN); text(s,T,x,2.48,1.36,.42,15,WHITE,True,PP_ALIGN.CENTER,True)
text(s,'申请  →  派单  →  处置  →  验收  →  评价  →  数据回写',.92,3.94,7.55,.55,13.7,WHITE,True,PP_ALIGN.CENTER); rect(s,8.6,1.95,3.74,3.7,'0B2D4D',BLUE,True,2); kpi(s,'66人','专业管护团队',8.98,2.4,3); kpi(s,'近190万','运维基金滚动累计',8.98,3.86,3); rich(s,[('有人管',YELLOW,True),('  /  ',CYAN,True),('有钱管',WHITE,True),('  /  ',CYAN,True),('有标准',WHITE,True),('  /  ',CYAN,True),('有机制',WHITE,True)],1.1,5.72,7.18,.52,15.3)
# 12 results
s=prs.slides.add_slide(blank); bg(s); title(s,'一个烤季，结果说话','技术、服务、机制都要落到可量化的业务结果','11 成效')
for i,(n,l) in enumerate([('2442座','建档 + 健康评估'),('48次','烤房调剂'),('18次','突发抢修全部 <3h'),('95.8%','群众综合满意度')]): x=.55+i*3.18; rect(s,x,1.8,2.85,2.06,NAVY2,BLUE,True,3); text(s,n,x+.1,2.1,2.65,.7,28,YELLOW,True,PP_ALIGN.CENTER); text(s,l,x+.16,2.93,2.53,.48,11.6,WHITE,True,PP_ALIGN.CENTER)
for i,(n,l) in enumerate([('340+','待修烤房'),('150','升级纳入计划'),('1363','合作社统筹管护'),('81%','管护覆盖率')]): x=.88+i*3.02; text(s,n,x,4.42,2.42,.55,21,WHITE,True,PP_ALIGN.CENTER,True); text(s,l,x,4.95,2.42,.34,10.8,MUTED,False,PP_ALIGN.CENTER)
tag(s,'工单平均响应 0.5 小时',4.62,5.78,4.1); text(s,'看得见、调得动、修得快、投得准',2.95,6.3,7.42,.48,16.6,WHITE,True,PP_ALIGN.CENTER,True)
# 13 innovation
s=prs.slides.add_slide(blank); bg(s); title(s,'评委为什么应该记住“智烤管家”','不是一个功能工具，而是一套可复制的基础设施治理方法','12 创新')
for i,(tg,body) in enumerate([('业务创新','一库 + 三智\n把“用修管养”串成闭环'),('技术创新','健康评估 + 发展导引\n两个模型同时解决“修与建”'),('机制创新','四队 + 基金\n让系统上线后仍能长期运营'),('推广价值','标准、流程、角色可复制\n可向核心烟区与其他设施扩展')]): x=.76+(i%2)*6.1; y=1.84+(i//2)*2.22; rect(s,x,y,5.7,1.78,NAVY2,BLUE,True,3); tag(s,tg,x+.24,y+.24,1.18); text(s,body,x+1.64,y+.26,3.7,1.12,13.2,WHITE,True)
rect(s,2.22,6.03,8.9,.72,'0A3557',YELLOW,True,2,1); text(s,'一套数据标准 + 一条决策链 + 三个应用场景 + 一套长效机制',2.42,6.16,8.5,.4,14,WHITE,True,PP_ALIGN.CENTER)
# 14 closing
s=prs.slides.add_slide(blank); bg(s); text(s,'从“有多少房”  到  “有多少有效烤能”',1.1,1.16,11.15,.88,31,WHITE,True,PP_ALIGN.CENTER,True); text(s,'一库清底数  ·  双模做决策  ·  三智优服务  ·  四队保长效',1.6,2.35,10.15,.48,15.2,ICE,True,PP_ALIGN.CENTER); rect(s,2.12,3.2,9.1,1.34,'0B2D4D',BLUE,True,2); rich(s,[('让数据说话',WHITE,True),('   ·   ',CYAN,True),('让模型决策',WHITE,True),('   ·   ',CYAN,True),('让管家落地',YELLOW,True)],2.52,3.55,8.28,.58,18); text(s,'逐步向兴山、五峰等核心烟区推广，并向育苗棚等基础设施扩展',2.05,5.1,9.25,.4,12.1,MUTED,False,PP_ALIGN.CENTER); text(s,'谢谢大家',4.87,5.88,3.6,.72,27,WHITE,True,PP_ALIGN.CENTER,True); text(s,'智同道合创客小组',5.02,6.62,3.3,.32,10.4,MUTED,False,PP_ALIGN.CENTER)
prs.save(PPT)

# Word guide
slides=['封面','三大难题','1234体系','一库清底','健康评估模型','千亩村发展导引模型','双模决策链','三智落地','智投','智用','智管+四队','应用成效','创新与推广','总结']
times=[20,35,30,30,45,45,30,20,30,35,35,40,35,25]; roles=['A+B','A→B','B→A','B→A','B→A','A','B','B→A','A','A→B','A','B','A','A+B→合']
mem=['我们解决烤房“用投管”一体化问题','分不到、修不及、投不准','1234是一条闭环，不是四个功能','没有标准化数据，就没有后续智能决策','60分预警，先知道谁最该修','先算有效烤能，再判断真实缺口','两个模型共同决定修、调、建','模型最终要进入真实业务动作','150栋升级 + 7栋新建','雷击案例3小时内复产','66人 + 近190万，保证长期运营','2442 / 48 / 18 / 95.8%','业务、技术、机制三重创新，可复制','从“有多少房”到“有多少有效烤能”']
sp=[
('A：尊敬的各位领导、各位评委，大家好！今天我们带来《智烤管家——基于健康评估的烤房用管管一体化解决方案》。\nB：我们希望解决的，不是做一张新台账，而是让每一座烤房都有一份可计算、可决策、可持续管理的健康档案。','第一页沿用认可的蓝白山海光路封面，只把所有文字改成可编辑。开场不要解释1234细节。','站中间，第一句看评委；说到“健康档案”停半秒。'),
('A：烤房对烟农来说，是烘烤季最关键的生产设施。但一到一线，最怕三件事：旺季分不到房，故障修不及，维修和新建又投不准。\nB：表面是用房、抢修和投入三个问题，往下追，本质是资源无统筹、投入无量化、管护无闭环。','本页只保留三张问题卡，黄色只标“分不到、修不及、投不准”。','A按左中右三张卡各点一次；B最后一句不要看屏幕。'),
('B：所以我们没有再加一个零散功能，而是重新设计了一条闭环。\nA：用一库清底把家底摸清，用双模决策把维修和建设算准，用三智把能力送到一线，再用四支队伍和机制保障长期运行。也就是“1234”创新体系。','4个模块同构，数字用黄色，名称用蓝白渐变；这一页是全场地图，不讲算法。','顺着1→2→3→4讲完即翻页。'),
('B：什么叫一库清底？\nA：我们围绕加热、散热、自控、附属4大模块，锁定18个关键部件，统一正常、损坏、缺失等判定口径，并通过移动端现场拍照录入，把基础信息、部件状态、维修记录、用户评价沉淀为一房一档。','左边讲标准，右边讲一房一档，屏幕不列18个具体部件名称。','B提问后侧一步，把视觉中心让给A。'),
('B：有了数据，怎么知道先修谁？\nA：第一套是烤房健康评估模型。我们以部件状态80%、管护情况10%、现场评价10%为核心依据，运用AHP层次分析法综合建设年限、使用时长、养护记录和实时工况，生成每栋烤房健康评分。60分以下自动预警，并优先进入维保清单。','本页必须让评委看到80/10/10和60分两件事。AHP只说作用，不解释公式。','说“60分以下自动预警”时停顿并指右侧黄色数字。'),
('第二套是千亩村发展导引模型。我们把单栋烤房健康结果折算成村级有效烘烤产能，再结合种植规模、产业趋势、集体经济、管护能力和政策适配度，利用熵权法和TOPSIS降低主观偏差，最后通过地图情景沙盘回答三个问题：哪里真缺、缺多少、建在哪里。','这是原创性最强的一页。页面只展示一条5步链和3个输出，不放公式。','“哪里真缺、缺多少、建在哪里”三句放慢。'),
('两个模型并不是并排展示。模型一先把存量健康算清，模型二再把健康折损变成有效烤能，和真实需求对比。最终形成一条投入决策链：先盘活存量、再精准修复，仍有结构性缺口才新增。','本页用一条流程代替复杂架构图，让评委听懂双模为什么有价值。','只讲逻辑，不再重复算法名。'),
('B：模型算完以后，怎样真正落地？\nA：我们把能力拆成三类真实场景：管理部门用智投，烟农和技术员用智用，合作社和管护队伍用智管。最终打通烟农、合作社、管护人员和管理部门。','三智总览页只是转场。若现场有真实操作录屏，可在本页后插15秒，不要再放1分钟视频。','20秒内结束；有视频就直接接视频。'),
('智投解决的是“钱该投到哪里”。系统通过双模研判筛选低效、老化、高危烤房，形成年度升级和新建清单。2026年，我们依托平台数据完成150栋核心设备升级计划，并支撑肖家坪千亩村科学新建7栋烤房。','150栋和7栋是本页唯一两个大数字，其他文字全部做成4步小流程。','两个数字各停半秒。'),
('A：智用解决烟农最直接的找房和报修。烟农可以查看周边烤房的位置、空闲时段和设备评分，自主预约调剂；故障时一键报修，系统匹配人员、配件和配送路线。\nB：今年磨坪站一栋烤房遭遇雷击，平台快速定位2处故障部件，匹配配件资源，3小时内完成抢修复产。','左侧两条短流程，右侧只讲一个真实案例。','讲案例时面对评委，不要指屏幕。'),
('智管解决的是“上线以后谁来管”。我们组建技术、抢修、养护、综合四支管家队伍，把申请、派单、处置、验收、评价、数据回写串成闭环，同时建立专项管护基金。目前秭归已形成66人专业管护团队，基金滚动累计近190万元。','四队用四个圆点，不做复杂图标。66人和近190万元用黄色大数字。','最后用“有人管、有钱管、有标准、有机制”收住。'),
('一个烤季下来，项目完成2442座烤房建档和全覆盖健康评估；完成48次烤房调剂；18次突发故障抢修全部在3小时内闭环；群众综合满意度达到95.8%。同时筛查340余栋待修烤房，150栋升级纳入计划，两级合作社已统筹管护1363栋、覆盖81%。','前四个数字做第一视觉层级，后四个只作为支撑证据。','四个主KPI一项一句，千万不要连读。'),
('我们认为，智烤管家的价值不只在功能。业务上，一库加三智把用修管养串成闭环；技术上，两套模型同时解决“怎么修”和“要不要建”；机制上，四队和基金保证长期运行。更重要的是，这套标准、流程和角色分工可以复制到其他核心烟区，也能向育苗棚等基础设施延伸。','这一页专门替评委做总结：业务创新、技术创新、机制创新、推广价值。','每个创新只说一句，不展开。'),
('A：回到最初的三难，我们用一库清底、双模决策、三智落地、四队长效，把烤房管理从经验判断变成数据决策。\nB：最终，我们希望管理的不再只是“有多少房”，而是“有多少有效烤能”。\n合：让数据说话，让模型决策，让管家落地。我们的汇报完毕，谢谢大家！','结尾不再放一排复杂按钮，只保留一句大标题、一句1234和一句价值主张。','最后一句两人并肩，停一秒再鞠躬。')]
d=Document(); sec=d.sections[0]; sec.top_margin=DInches(.62); sec.bottom_margin=DInches(.62); sec.left_margin=DInches(.72); sec.right_margin=DInches(.72); d.styles['Normal'].font.name=FONT_CN; d.styles['Normal']._element.rPr.rFonts.set(dqn('w:eastAsia'),FONT_CN); d.styles['Normal'].font.size=DPt(10.5)
for nm,sz,col in [('Title',23,'0B4F7A'),('Heading 1',16,'0B4F7A'),('Heading 2',13,'17658C')]: st=d.styles[nm]; st.font.name=FONT_CN; st._element.rPr.rFonts.set(dqn('w:eastAsia'),FONT_CN); st.font.size=DPt(sz); st.font.color.rgb=DRGB.from_string(col)
p=d.add_paragraph(); p.style=d.styles['Title']; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run('智烤管家｜8分钟决赛逐页演讲稿与PPT设计说明'); r.bold=True
p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run(f'配套PPT：{PPT.name}\n生成时间：2026-10-03 17:42（北京时间）').font.size=DPt(9.5)
d.add_heading('一、整套发布怎么讲',1)
for t in ['整套PPT压缩为14页，主口播约7分35秒；若保留现场操作演示，建议仅15秒，总时长约7分50秒。','叙事只保留一条主线：三难 → 1234解法 → 一库 → 双模 → 三智 → 长效 → 成效 → 创新与推广。','屏幕承担“结构和证据”，讲述承担“解释和故事”。每页只让评委记住1个结论，不在PPT上堆讲稿。','黄色只用于最重要的词、数字和判断；其余保持蓝白。标题使用蓝白渐变，正文用普通白/冰蓝。']: d.add_paragraph(t,style='List Bullet')
d.add_heading('二、8分钟时间表',1); tb=d.add_table(rows=1,cols=5); tb.style='Table Grid'; tb.alignment=WD_TABLE_ALIGNMENT.CENTER
for i,h in enumerate(['页','页面','时间','角色','评委只需记住']): tb.rows[0].cells[i].text=h; shd=DOxmlElement('w:shd'); shd.set(dqn('w:fill'),'0B4F7A'); tb.rows[0].cells[i]._tc.get_or_add_tcPr().append(shd); [setattr(run.font.color,'rgb',DRGB(255,255,255)) for run in tb.rows[0].cells[i].paragraphs[0].runs]
for i,(nm,tm,ro,me) in enumerate(zip(slides,times,roles,mem),1):
    c=tb.add_row().cells
    for j,v in enumerate([str(i),nm,f'{tm}s',ro,me]): c[j].text=v
d.add_paragraph('主口播累计约7分35秒；加15秒演示约7分50秒。'); d.add_heading('三、逐页口播与设计理由',1)
acc=0
for i,(say,des,act) in enumerate(sp,1):
    start=acc; acc+=times[i-1]; t1=f'{start//60}:{start%60:02d}—{acc//60}:{acc%60:02d}'
    d.add_heading(f'第{i}页｜{slides[i-1]}',2); p=d.add_paragraph(); rr=p.add_run(f'时间：{t1}    角色：{roles[i-1]}'); rr.bold=True; rr.font.color.rgb=DRGB(14,91,130); p=d.add_paragraph(); p.add_run('建议口播：').bold=True; p.add_run(say); p=d.add_paragraph(); p.add_run('为什么这样设计：').bold=True; p.add_run(des); p=d.add_paragraph(); p.add_run('临场动作：').bold=True; p.add_run(act)
d.add_heading('四、素材与视频怎么补',1); tb=d.add_table(rows=1,cols=4); tb.style='Table Grid'; tb.alignment=WD_TABLE_ALIGNMENT.CENTER
for i,h in enumerate(['优先级','页面','素材','建议']): tb.rows[0].cells[i].text=h
for row in [('A','第10页 智用','15秒真实操作快剪','找房→预约→报修→工单→评价；不要做宣传片片头片尾。'),('A','第4页 一库','真实“一房一档”截图','后续有真实后台截图时，可替换右侧档案示意。'),('B','第9页 智投','肖家坪真实点位图','有真实地图时作为小证据，不再增加文字。'),('B','第11页 智管','四队合影或抢修现场','建议放答辩备份页，主发布页保持简洁。')]:
    c=tb.add_row().cells
    for i,v in enumerate(row): c[i].text=v
d.add_heading('五、超时应急版（30秒收口）',1); p=d.add_paragraph(); p.add_run('若现场被提示时间，直接说：').bold=True; p.add_run('“我们用一库清底解决底数不清，用双模决策解决维修和建设靠经验，用三智把能力送到一线，再用四队和基金保证长期运营。一个烤季完成2442座建档评估、48次调剂、18次抢修全部3小时内闭环，满意度95.8%。智烤管家最终要做的，就是让烤房管理从‘有多少房’走向‘有多少有效烤能’。谢谢大家！”')
d.save(DOC)
print(PPT); print(DOC)
