# -*- coding:utf-8 -*-
from pathlib import Path
from docx import Document
from docx.shared import Pt,Cm,RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re
HERE=Path(__file__).resolve().parent
ROOT=HERE.parents[2]
OLD=ROOT/'1019/v4/修改稿'
SOURCE=HERE/'10192340_V5_智烤管家AB决赛稿_PPT逐页对照.md'
RAW=SOURCE.read_text(encoding='utf-8')
SECTIONS=[];cur=None
for line in RAW.splitlines():
 if line.startswith('## 【P'):
  match=re.match(r'## 【(P[^】]+)】(.+)',line)
  cur=[match.group(1),match.group(2),[]]
  SECTIONS.append(cur)
 elif line.startswith(('A：','B：')):
  cur[2].append((line[0],line[2:]))
def han(x):return len(re.findall('[\u4e00-\u9fff]',x))
a=sum(han(t) for _,_,ls in SECTIONS for r,t in ls if r=='A')
b=sum(han(t) for _,_,ls in SECTIONS for r,t in ls if r=='B')
assert 1450<=a+b<=1580 and .69<=a/(a+b)<=.71,(a,b)
assert len(SECTIONS)==16
def add(d,s,style=None):
 p=d.add_paragraph(style=style);p.add_run(s);return p
def hdr(d,s):return add(d,s,'Heading 2')
def tbl(d,heads,rows):
 t=d.add_table(rows=1,cols=len(heads));t.style='Table Grid'
 for cell,v in zip(t.rows[0].cells,heads):cell.text=v
 for row in rows:
  cells=t.add_row().cells
  for cell,v in zip(cells,row):cell.text=str(v)
 return t
def title(d,text):
 p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
 r=p.add_run(text);r.bold=True;r.font.name='黑体';r.font.size=Pt(14)
 r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'黑体')
def stripped_original():
 d=Document(ROOT/'1004/智烤管家演讲讲稿9.30.docx')
 for e in list(d._element.body):
  if e.tag!=qn('w:sectPr'):d._element.body.remove(e)
 return d
def speech(stage):
 d=stripped_original()
 title(d,'智烤管家——基于健康评估的烤房用投管一体化解决方案')
 if stage:
  d.sections[0].top_margin=Cm(1.65);d.sections[0].bottom_margin=Cm(1.65)
 for page,name,ls in SECTIONS:
  if stage:
   p=d.add_paragraph();p.paragraph_format.keep_with_next=True
   p.paragraph_format.space_after=Pt(1)
   r=p.add_run(f'{page}｜{name}');r.font.name='黑体';r.font.size=Pt(9)
   r.font.color.rgb=RGBColor(85,104,133)
   if page=='P2':
    p=d.add_paragraph('（视频需裁剪至30秒，AB静默；原素材约41.5秒。）')
    for r in p.runs:r.font.size=Pt(10)
  for role,txt in ls:
   p=d.add_paragraph()
   p.paragraph_format.space_after=Pt(2 if stage else 4)
   p.paragraph_format.line_spacing=1.04 if stage else 1.12
   r=p.add_run(role+'：');r.bold=True;r.font.name='黑体';r.font.size=Pt(14)
   r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'黑体')
   r=p.add_run(txt);r.font.name='仿宋_GB2312';r.font.size=Pt(14)
   r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'仿宋_GB2312')
 return d
speech(False).save(HERE/'10192342_04_V5_正式决赛8分钟AB演讲净稿.docx')
speech(True).save(HERE/'10192342_05_V5_25页PPT对应双人舞台稿.docx')
def base(name,label):
 f=OLD/name
 assert f.exists(), f
 d=Document(str(f))
 return d
# 01 award work; every source was read in earlier versions, V5 addition points to actual named txt
d=base('10192325_01_12份获奖解说词学习与V4借鉴.docx','优秀解说词')
hdr(d,'V5专题：十二篇获奖答辩录音逐项复盘与迁移')
insights=[
('1.txt','3D智慧对标：用“数据太多无法下手”的情景引出三难，先给故事再给模型。'),
('2.txt','网络情报：角色小剧场帮助听众迅速理解异常触发，须尽快回到业务证据。'),
('3.txt','AI智能卷烟营销：从三类难题到数据驱动决策，技术名词后紧跟业务结果。'),
('4.txt','异常监管：用风险识别场景展示算法价值，而不是孤立介绍公式。'),
('5.txt','寄递数据研判：采数、析数、用数三连环，适合一库、双模、三智的逻辑。'),
('6.txt','会思考的食堂：双人自然接话增加亲近感，功能与实效须对应。'),
('7.txt','滞销品规管理：反差句让人记住业务痛点，借鉴“有房未必调得动”。'),
('15.txt','烟叶收购质量：田间实践做叙事锚点，面向非烟叶评委仍能解释清楚。'),
('19.txt','财务审核：台账繁杂与责任链形成数字化改革的现实动机。'),
('30.txt','雪茄三问：以连续提问打开业务问题，一一回应、最后回扣。'),
('9月9日我的.txt','早期本项目讲述：一库两模平台骨架已明确，V5升级为1234闭环。'),
('yus07.txt','雪茄三个没数：信息断点比抽象算法更易获得共情。')
]
for fn,explanation in insights:
 hdr(d,fn+'｜表达手法与转化')
 add(d,explanation)
 add(d,'方法迁移：V5借鉴结构与节奏，不挪用该作品的技术成果、获奖名次或业务数据。')
add(d,'V5总评：最有效的技术呈现不是术语更密，而是让评委先问“为什么”，再用模型的输入、计算、输出一步步作答。A角承担约70%技术与业绩，B角约30%提问与情感接力。')
d.save(HERE/'10192342_01_V5_十二篇获奖现场解说词学习与创新借鉴.docx')
# 02 full research based on complete V4 volume, 8 diagrams retained
d=base('10192325_02_五万字项目全文件研究与V4技术增订.docx','项目研究')
hdr(d,'V5新增专题：技术创新为什么不是两个模型简单相加')
for text in [
'模型一通过档案、部件、现场评价和维修记录生成健康分与风险部位，PPT第8页把12项算法参数用于健康计算；第6页18项为现场清查口径。两个数量不能混写。',
'综合健康分按部件状态80%、管护10%、现场评价10%形成。AHP层次分析法用于可解释的部件赋权，形成风险部位TOP3和维修清单。寿命估算在P8与P9之间有不同表述，需在正式比赛前核验现行算法。',
'第二模型的“数需评推”依次对应输入烟叶生产和候选点数据、折算有效烤能与理论需求、AHP和熵权组合赋权配合TOPSIS排序、地图沙盘推演新建数量及位置。输出是新建优先村、建议数量、候选点排序和改善情景。',
'这套方案的实质耦合：单栋存量烤房健康变化影响村级有效烤能；有效烤能影响需求缺口；缺口改变建设数量和点位的优先级。因此“先修存量、审慎新建”不是宣传口号而是模型结构带来的管理行为。',
'三智和四队是系统落地的闭环机制。烟农发起调剂和报修，合作社派单维修，现场验收和评价结果更新档案，模型再算。若四队与两项制度缺位，平台只是记录工具。',
'统计口径：2442座全口径建档；48次调剂；18次报修，另16条工单平均响应0.38小时；健康分低于60的340多座是预警研判对象；1363栋统筹管护覆盖本年度1672栋在用中的约81%；近190万元基金为总额；531份有效问卷满意度96.8%。',
'未纳入V5主稿的原稿数字包括150栋已修、肖家坪7栋已建、12478张影像、6704条评价、66人专业队伍及磨坪站雷击维修详情。并非断定这些记录错误，而是目前已核对的报送成果没有逐项给出证明，须补数据原件方可声称完成。'
]:add(d,text)
d.save(HERE/'10192342_02_V5_项目五万字级全文件研究与双模创新增订.docx')
# 03 official prior official volume >10k Han
d=base('10192325_03_官方叙述规范一万字V4校订版.docx','官方规范')
hdr(d,'V5附章：演讲激情和官方严谨的双重标准')
for text in [
'保留“匠心工坊”“资源无统筹、投入无量化、管护无闭环”等原稿关键词。用鲜明对仗提升可听性，但不把“模型建议”说成“政策批准”，不把“预警”说成“修复完成”。',
'身份定位：智烤管家是烤房用投管一体化解决方案。AHP、熵权、TOPSIS和健康分是业务决策方法，AI助手是查询辅助，不能误称“AI全自动选址”“AI替代维修决策”。',
'表达顺序：30秒视频 → 三愁三缺 → 1234体系 → 一库 → 两模 → 三智 → 四队 → 量化成效 → 可复制的机制。烟叶与信息条线评委均可沿因果链理解。',
'技术底稿务必回答：18项清查与12项算法参数的映射；P8寿命短板融合公式与P9 Weibull拟合表述的关系；16条平均响应工单与18次报修的不同统计分母。',
'正式舞台版本共1554个口述汉字；A/B角色约70/30；按220字每分钟口述、30秒视频和20秒切页估计约7分54秒，属于紧凑编排，必须实测彩排。'
]:add(d,text)
d.save(HERE/'10192342_03_V5_官方叙述一万字级研究校订版.docx')
# 06 25-page sheet
d=Document();title(d,'V5·25页PPT与双人话术对应检查')
pnames=['封面','开场视频','烟农三愁','三缺','1234总览','一房一档','双模总览','健康模型公式','健康模型三步','发展模型公式','数需评推','使用者','三智循环','调剂报修','调剂报修数字','维修清单','新建沙盘','智管闭环','AI助手','管护数据基金','四队职责制度','运行成效','创新价值','推广可行性','结束语']
mapping={}
for pg,topic,ls in SECTIONS:
 m=re.match(r'P(\d+)(?:—P(\d+))?',pg)
 for n in range(int(m.group(1)),int(m.group(2) or m.group(1))+1):mapping[n]=(pg,topic,ls)
assert sorted(mapping)==list(range(1,26))
for page,name in enumerate(pnames,1):
 hdr(d,f'第{page}页｜{name}')
 pg,topic,ls=mapping[page]
 if page==2:add(d,'AB静默播放30秒，原视频约41.5秒须先剪辑。')
 elif page==int(re.match('P(\d+)',pg).group(1)):
  for role,t in ls:add(d,f'{role}：{t}')
 else:add(d,'与前页同属一个技术或应用段落。本页展示支持画面，不额外增加独立台词。')
 if page in (8,9):add(d,'注意寿命算法图文的一致性：P8公式与P9拟合术语须由研发团队统一。')
 if page in (14,16,17,19):add(d,'此页嵌入视频或演示素材不应与本稿预算叠加完整播放。')
hdr(d,'时长及AB比例')
tbl(d,['语速（汉字/分）','口述（秒）','视频/换页（秒）','估算总时长'],[(s,round((a+b)/s*60,1),50,round((a+b)/s*60+50,1)) for s in (220,230,240)])
add(d,f'A={a}字、B={b}字，A占{a/(a+b)*100:.2f}%。只统计AB台词，不统计标题或视频字幕。')
d.save(HERE/'10192342_06_V5_25页PPT逐页对应与八分钟彩排卡.docx')
# 07 comparison and QAs
d=Document();title(d,'V5与9.30原稿：保留、替换、补强与评委备答')
rows=[
('匠心工坊、一年收成','保留，P3引出三愁','情感线不可丢'),
('资源无统筹、投入无量化、管护无闭环','原句保留，配排比','由烟农三愁向管理深因跃迁'),
('健康评估80/10/10、AHP','保留四类输入、健康分、寿命与TOP3','使技术能解释而非只报术语'),
('千亩村新建模型','补数需评推、有效烤能、AHP熵权、TOPSIS、沙盘','体现存量—增量之间两模耦合'),
('智用、智投、智管','按烟农服务到投资再到管护分层','避免功能菜单式汇报'),
('四队、基金、制度','保留并解释各队分工及管护覆盖范围','可落地与可持续'),
('原稿95.8%满意度','使用官方531份有效问卷96.8%','采用正式材料'),
('原稿0.5小时响应','限定16条平台工单平均0.38小时','区分样本/指标'),
('原稿150栋更新、肖家坪7栋新建','作为待核线索，暂不报已完成','防止计划替代验收'),
('原稿基金年底结余超100万元','用基金总额近190万元','总额不等于期末结余')
]
tbl(d,['原稿特色/数字','V5处理','理由'],rows)
hdr(d,'最可能被问的5个技术问题')
for t in [
'1. 为什么采用AHP？答：通过专家判断确定部件重要性权重，确保健康分的技术可解释性，同时以现场状态、年限和评价构成输入。',
'2. 为什么选址同时用AHP、熵权、TOPSIS？答：结合专业重要性与客观差异形成权重，再用TOPSIS比较候选点的相对优先级，最终仍需现场踏勘与审批。',
'3. 为什么要算有效烤能？答：名义栋数不等于真实可用能力；现有设备健康变化可能影响村里真实缺口。',
'4. 如何实现闭环？答：调剂报修→派单维修→验收评价→回写档案→模型复算，四队和制度负责线下执行。',
'5. P8与P9寿命估算为何不完全相同？答：现有页面同时呈现合成公式与Weibull表述，正式比赛前先核实实际实现版本，台词不编造两者关系。'
]:add(d,t)
d.save(HERE/'10192342_07_V5_9月30日原稿对照与评委技术问答.docx')
# source-manifest
lines=[f'# V5 Word QA\nA {a} B {b} all {a+b}; 30s video, 20s transitions.','']
for f in sorted(HERE.glob('10192342_*.docx')):
 z=Document(str(f));lines.append(f'- {f.name}: {len(z.paragraphs)} paragraphs, {len(z.inline_shapes)} images, {f.stat().st_size} bytes')
(HERE/'10192342_核验清单.md').write_text('\n'.join(lines),encoding='utf8')
print('\n'.join(lines))

# trigger V5 full-source Word generation and auditing
