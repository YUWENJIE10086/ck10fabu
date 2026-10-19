# -*- coding: utf-8 -*-
"""1019/v1 Word资料研究/台稿版式生成，源文件均取自仓库真实文件。
复制并沿用9.30稿的Normal正文配置与页面样式。"""
import re, io, os, json, textwrap, random, collections, hashlib
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
BASE=ROOT/'1004'/'智烤管家演讲讲稿9.30.docx'
SP=OUT/'10192016_智烤管家_双人演讲稿正文.md'
STAMP='10192020'
SRC=ROOT/'读取资料'
REF=SRC/'1学习优秀的'
MATS=SRC/'2我的项目相关的零碎的'
OFFICIAL=SRC/'3之前项目官方叙述过的'
NEWS=ROOT/'1010-V2_智烤管家_副本.pptx'
from bs4 import BeautifulSoup
from pptx import Presentation

def tidy(s):
 return re.sub(r'\n{3,}','\n\n',s.replace('\x00','')).strip()
def read_any(f):
 e=f.suffix.lower()
 if e=='.txt':
  for enc in ('utf-8-sig','gb18030','utf-16'):
   try:return f.read_text(encoding=enc)
   except Exception:pass
 if e in ('.docx',):
  d=Document(str(f)); x=[p.text for p in d.paragraphs if p.text.strip()]
  for t in d.tables:
   for r in t.rows:
    cs=[c.text.replace('\n','；') for c in r.cells]
    # avoid table repeated nonsemantic cells
    if any(cs):x.append(' | '.join(dict.fromkeys(cs)))
  return '\n'.join(x)
 if e in ('.html','.htm','.drawio'):
  s=f.read_text(encoding='utf-8',errors='replace')
  if e=='.drawio':
   x=BeautifulSoup(s,'xml')
   return '\n'.join(BeautifulSoup(a.get('value',''),'html.parser').get_text(' ',strip=True) for a in x.find_all('mxCell') if a.get('value'))
  x=BeautifulSoup(s,'html.parser')
  for s in x(['script','style']):s.decompose()
  return x.get_text('\n',strip=True)
 if e=='.pptx':
  p=Presentation(str(f))
  return '\n'.join('[第%d页] %s'%(i,' / '.join(sh.text.strip() for sh in sl.shapes if sh.has_text_frame and sh.text.strip())) for i,sl in enumerate(p.slides,1))
 if e=='.doc':
  import subprocess
  q=subprocess.run(['antiword',str(f)],capture_output=True,timeout=30)
  return q.stdout.decode('utf-8','replace')
 if e in ('.xls','.xlsx'):
  try:
   if e=='.xlsx':
    import openpyxl
    w=openpyxl.load_workbook(f,data_only=True,read_only=True)
    return '\n'.join(sh.title+'\n'+'\n'.join(' | '.join(str(c) for c in row[:15] if c is not None) for row in list(sh.values)[:100]) for sh in w)
   import xlrd
   w=xlrd.open_workbook(f)
   return '\n'.join(sh.name+'\n'+'\n'.join(' | '.join(str(c) for c in sh.row_values(i)[:15]) for i in range(min(sh.nrows,100))) for sh in w.sheets())
  except Exception as ex:return '数据表格待核查：'+str(ex)
 if e=='.pdf':
  import fitz
  d=fitz.open(f);return '\n'.join(p.get_text() for p in d)
 return ''

def add(doc,txt,kind='normal'):
 if not txt:return
 p=doc.add_paragraph(style='Normal')
 if kind=='heading':
  p.style=doc.styles['Heading 2']
 elif kind=='sub':
  p.style=doc.styles['Heading 3']
 if kind=='quote':
  p.paragraph_format.left_indent=Cm(.6)
  p.add_run('资料摘录｜').bold=True
 p.add_run(txt)
 return p
def title(doc,s,subtitle=None):
 p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
 r=p.add_run(s);r.bold=True;r.font.name='黑体';r.font.size=Pt(18);r.font.color.rgb=RGBColor(18,58,89)
 if subtitle:
  p=doc.add_paragraph(subtitle);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
def init():
 d=Document()
 s=d.sections[0];s.top_margin=Cm(2.25);s.bottom_margin=Cm(2.0);s.left_margin=Cm(2.45);s.right_margin=Cm(2.3)
 norm=d.styles['Normal'];norm.font.name='仿宋_GB2312';norm.font.size=Pt(12)
 norm._element.rPr.rFonts.set(qn('w:eastAsia'),'仿宋_GB2312')
 for t in ('Heading 1','Heading 2','Heading 3'):
  d.styles[t].font.name='黑体';d.styles[t]._element.rPr.rFonts.set(qn('w:eastAsia'),'黑体')
 return d
def section(doc,head,body):
 add(doc,head,'heading')
 for p in re.split(r'\n\s*\n',body):
  if p.strip():add(doc,p.strip())
def tbl(doc,heads,rows):
 t=doc.add_table(rows=1,cols=len(heads));t.style='Table Grid';t.alignment=WD_TABLE_ALIGNMENT.CENTER
 for i,x in enumerate(heads):t.rows[0].cells[i].text=str(x)
 for row in rows:
  cells=t.add_row().cells
  for i,x in enumerate(row):cells[i].text=str(x)
 return t
def figure(doc,texts,desc):
 w,h=1400,300
 im=Image.new('RGB',(w,h),'white');dr=ImageDraw.Draw(im)
 ff='/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'
 font=ImageFont.truetype(ff,39) if os.path.exists(ff) else ImageFont.load_default()
 for k,tx in enumerate(texts):
  x=20+k*(w-40)//len(texts); y=78; right=20+(k+1)*(w-40)//len(texts)-12
  dr.rounded_rectangle((x,y,right,y+126),radius=17,fill=(226,239,250),outline=(42,98,139),width=4)
  dr.text((x+18,y+38),tx,font=font,fill=(12,51,81))
  if k<len(texts)-1:dr.text((right+2,y+45),'→',font=font,fill=(12,51,81))
 buf=io.BytesIO();im.save(buf,format='PNG');buf.seek(0)
 p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.add_run().add_picture(buf,width=Cm(15.7))
 p=doc.add_paragraph('图｜'+desc);p.alignment=WD_ALIGN_PARAGRAPH.CENTER

# 01: 优秀获奖稿分析：逐份真正读取，逐篇留痕。
cases={
'1.txt':('3D智慧对标','生活化提问开头、三大痛点排比、双人问答“柯博士”、以指标转变与推广闭合','避免把“3D”“AI”等技术标签当创新成果，学习问题—能力—数据—价值的闭环'),
'2.txt':('网络涉烟信息研判','视频小短剧引入业务危机、演示工具介入现场、问题推动情节','从困境到工具出现的转场可借鉴，但不要硬凑表演'),
'3.txt':('AI驱动卷烟供需调控','政策—业务难题—技术方案—结果的清晰推进，讲解算法时突出供需匹配','智烤项目以“修哪栋、建哪村”代替算法术语堆砌'),
'4.txt':('大户监管模型','问题强烈、模型拟人“透视镜”、具体案例证明识别能力','智烤使用“体检报告、健康档案、开维修清单”更贴切'),
'5.txt':('物流寄递采集研判','“这位同志”戏剧化亮相，一句话说清拍、扫、查，重实操演示','智烤借鉴手机一键报修的视觉演示，但谨防夸大自动诊断能力'),
'6.txt':('智慧食堂智善魔方','巧用生活场景与轻松双人交流，价值外延易共情','智烤讲烟农急、管家跑、维修抢，形成情绪共鸣'),
'7.txt':('滞销品规智能投放','以矛盾数据导入，追问“为什么”，后续靠真实对比说服','智烤强调2442建档与1363管护的不同统计口径'),
'15.txt':('数字收购预警','贴近烟叶专业、准确陈述政策诉求与量质目标，表达并不悬浮','专业术语适量，烟叶评委听得准、信息评委听得懂'),
'19.txt':('合同支付监管','从日常工作共性烦恼切入，痛点听起来真切','三愁对应三缺，保持紧凑并注意口播节奏'),
'30.txt':('雪茄圈层智联','“三问”设置观众期待，再逐条解答','智烤采用“想用、想修、想问”的三愁结构'),
'9月9日我的.txt':('智烤管家早期介绍','保留2400余座与数据模型的项目线索','仅作项目早期意见，不直接将预计建设或维修当作完成'),
'yus07.txt':('雪茄NFC全程追溯','三个“没数”一口气抛出，技术很快落到一物一码、追溯和经营','智烤从一座烤房一份档案衔接日常闭环')
}
f=init();title(f,'一等奖演讲表达法：12份解说词逐篇学习报告','1019/v1 · 不是照搬获奖稿，而是提炼适合智烤管家的现场表达')
section(f,'一、研究方法与边界','逐份读取文件夹“读取资料/1学习优秀的”中的全部12份txt，识别开场抓手、现场冲突、双人对话、技术说法、实绩证据、尾声升华；将可以直接迁移的表达手段单列。解说词为口述转写，可能存在同音字和口误，不将其中数字转为智烤项目数据。')
figure(f,['一线三愁','1234破题','两模求证','三智四队','结果落地'],'适用于智烤管家的获奖叙事节拍')
for name,(theme,good,transfer) in cases.items():
 add(f,'逐篇学习｜'+name+'《'+theme+'》','heading')
 add(f,'获奖表达可取之处：'+good)
 add(f,'适配智烤管家的方式：'+transfer)
 s=read_any(REF/name);first=re.sub(r'\s+','，',s.strip())[:180]
 add(f,'核验片段（仅供溯源，不作为智烤项目事实）：'+first)
section(f,'二、12篇共性比较','一流现场汇报往往先立住一个业务矛盾，而不是先报出算法。好稿子的技术密度由听众的理解程度决定：主舞台只讲“数据从哪里来、为什么得出这个结果、结果给谁使用”；方法名与参数移入评委追问备答。双人对话的价值不是把一段话平均分给两人，而是A承担主线，B用疑问、补充、点睛、情绪变化帮助观众理解。')
tbl(f,['环节','一等奖表达目标','智烤管家首版操作'],[
 ('开头','一句生活冲突','让烤房也有健康档案'),('痛点','三句记得住','想用调不到、想修等不及、想问问不明'),('问题归因','收拢三难','资源无统筹、投入无量化、管护无闭环'),('模型','复杂方法说简单','体检—预警—修复；供需—排序—沙盘'),('实绩','同口径数字','2442/48/18/0.38/531/96.8/1363/81%'),('结尾','短句且有景象','一库清底数、两模明方向、三智惠烟农、四队护长远')])
section(f,'三、实战建议：保留而不是照抄','推荐用“烤房健康档案”贯穿全场；前段用情景让非烟叶条线理解，中段用模型逻辑让信息专家信服，后段用已核实数量和合作社制度证明可落地。押韵只用于结尾和关键转场，不能逐句凑押韵。不要抄其他项目的专有数字、名词和专利表述，不把获奖风格误写成已获一等奖。')
f.save(OUT/f'{STAMP}_01_获奖演讲逐篇学习报告.docx')

# 02: 大型研究卷：真正扫描每一份文件，重复SHA只计一版；正文作证据索引，附录呈现具有可追溯性的长篇材料。
src_files=[z for z in sorted(MATS.rglob('*')) if z.is_file() and z.name!='.DS_Store']
unique=[];dups=[];seen={}
for z in src_files:
 h=hashlib.sha256(z.read_bytes()).hexdigest()
 if h in seen:dups.append((z,seen[h]));continue
 seen[h]=z;unique.append(z)
d=init();title(d,'智烤管家项目综合研究卷（资料全覆盖·v1）','以官方报告作事实锚点；零散资料逐项审视；技术方案与实际成效分离')
section(d,'编写声明与范围',f'本卷逐项登记“2我的项目相关的零碎的”{len(src_files)}份文件，实际不同内容{len(unique)}份；同一哈希的副本登记但不重复堆字。技术方案、愿景讨论、实际试点、效果统计四种证据级别分别处理，任何早期设想均不自动等于落地成果。本卷的后半部是注明来源的逐文件资料读书笔记和证据摘录，不将摘录伪称为原创科研论述。')
figure(d,['一房一档','健康与发展','三智协同','四队运维'],'智烤管家基础研究技术—业务耦合图')
chapters=[
 ('第一章 问题从哪里来','为什么烤房资产有档却难调剂？为什么设备故障报修不及时？为什么维修或新建投入必须先比较缺口？烟农端问题与管理端问题相互缠绕，典型表现是资源无统筹、投入无量化、管护无闭环。研究不能只列软件菜单，需要解释三种痛点如何沿数据流形成因果链。'),
 ('第二章 数据普查与资产口径','烤房全口径建档不能等同于在用烤房或合作社统筹管护烤房。2442是秭归建档的资源规模；1672是本年度在用规模；1363为统筹管护规模；81%是1363除以1672的约数。统计需固定时点、去重口径与编号规则。'),
 ('第三章 低代码数据采集','资料提出手机拍照、GPS定位、项目编号关联、表单现场核实；结合22个低代码应用、6类报表将业务数据结构化。需区分采集工具“简化录入”和校验机制“保障质量”，防止出现同一烤房不同编号、多套部件口径的重复问题。'),
 ('第四章 四大部位与关键部件','PPT展示加热、散热、自控及主体等关键部位和18个清查关键件；最新健康评估流程图展示12个用于评分的部件参数。两种口径不可混为一谈，应注明现场建档项与模型计算项、权重、寿命指导值的不同。'),
 ('第五章 健康分与风险预警','模型的决策用途是排序，不是“准确预测某日必坏”。健康分将部件状态、维护状况和人工评价综合得到，60分为现场演讲采用的风险提示阈值；实际维修仍要进行检查、成本比较和安全确认。'),
 ('第六章 剩余寿命估算','寿命估算结合建设年限、最近修复年度、指导寿命、状态系数和健康系数。历史材料曾讨论Weibull等统计模型，不可对仅有定性状态数据的设备声称有高精度寿命预测验证。建议在技术答辩中说明数据限制、参数标定和保守预警策略。'),
 ('第七章 千亩村供需发展导引','模型需要先区分名义烤房数量和有效烤能，再与种植面积、烘烤需求、村级发展指标比较；在候选点层面采用AHP与熵权组合赋权、TOPSIS排序，随后用情景推演分析新增数量及布局。它提供辅助建议，不替代合规审批、现场踏勘和专家评审。'),
 ('第八章 智用：调剂服务与抢修','烟农端应围绕找房、提交需求、一键报修、进度跟踪、评价回写几个连续动作。48次调剂和18次报修是2026年成果报告记录，16条平台抢修工单平均响应0.38小时属于另一统计子集，不能将报修处置时长与响应时长混为一谈。'),
 ('第九章 智投：年度修复与新建','管理端根据健康风险清单、修复资金预算、备件可得性等因素判断维修优先级；根据千亩村候选地点的供需与发展条件，形成新建推荐。早期材料出现150栋维修计划、肖家坪7栋等叙述，若没有正式验收文件，不作为当前已完工指标。'),
 ('第十章 智管：完整责任链','一次故障从申请、派单、抢修、验收、评价到维修历史回写，形成可追溯记录。只有数据按要求更新，下一次健康分、风险排序才会反映运维结果。AI助手主要承担辅助查询，不宜夸为全自动工程决策。'),
 ('第十一章 四队管家与基金机制','由合作社体系分工组织烟叶调制、紧急抢修、日常管护、综合管理队伍；制度解决怎么干、平台解决怎么留痕、基金解决长期投入来源。官方材料给出1363栋参管、基金总额近190万元，但基金总额不等同于当期结余。'),
 ('第十二章 实际效果与证据质量','2026年官方成果材料称2442座全口径入库，健康分60以下340多栋；48次调剂、18次报修，16条工单平均响应0.38小时，531份有效问卷满意度96.8%，1363栋统筹管护占1672栋在用约81%。对于没有对照实验的效率提升，不给出未经测量的百分比。'),
 ('第十三章 信息系统架构与治理','数据表、采集门户、管理后台、业务小程序和BI可视化的职责应清晰区分；数据主键、权限、审计日志、脱敏和责任归属是答辩中重要但不必占用主讲时间的技术准备。'),
 ('第十四章 评估方法与改进路线','下一轮完善应集中解决健康分分布稳定性、部件权重审查、维修前后对比、村级需求模型校准、跨县迁移、基金财务闭环；多渠道材料出现冲突时优先官方报送版，然后参考最新版PPT并注明口径变更。'),
 ('第十五章 可复制应用前提','复制时应迁移数据模型、流程标准与责任机制，而不是仅复制软件代码。外地区要校准烤房编号体系、使用状态、部件参数、种植规模、村级烤能测算、地形与距离因素，推广表述宜使用“可复制条件具备”而非“已经全省推广”。')]
for h,body in chapters:section(d,h,body)
tbl(d,['证据指标','官方口径','不宜混淆'],[
 ('2442','秭归全口径建档/评分','不是正在使用或纳入管护的栋数'),('1672','本年度在用烤房','不是所有已建烤房'),('1363 / 81%','两级合作社统筹管护/相对在用占比','不是2442的81%'),('48 / 18','调剂次数/报修工单次数','不是同一服务'),('16 / 0.38小时','16条工单平均响应时间','不是维修完结时间'),('531 / 96.8%','有效问卷/综合满意度','不沿用初稿95.8%'),('近190万元','管护基金总额','不是年度结余'),('340多栋','健康分60以下预警对象','不是全部已经完成维修')])
section(d,'附录A：项目素材逐文件研读及可追溯摘录','以下逐份登记源码、文稿、表格、流程图、视频说明、设计示意和报告草稿。每一篇摘录均保留文件路径。因同一项目存在迭代版本，对算法阈值、健康部件数和统计口径不一致的材料，不以早期版本直接覆盖正式报送口径。')
total_chars=0
for i,pth in enumerate(unique,1):
 rel=pth.relative_to(ROOT).as_posix()
 s=tidy(read_any(pth))
 h=f'{i:03d}. {pth.name}'
 add(d,h,'heading');add(d,'来源：'+rel+'；文件大小：'+str(pth.stat().st_size)+'字节。')
 if not s:
  add(d,'该文件为图片或其他非文本资料，保留原始路径；结合原图用于架构/界面/统计图核验。')
  continue
 # distinct, max about 1k per file; longer docs with informative detail allow 3000.
 cap=3200 if len(s)>6000 else 1700
 txt=s[:cap]
 add(d,'逐文件研究提示：本文含项目决策条件、阶段性想法或演示内容；与正式报告不一致之处需标为“历史版本/待核”。')
 for piece in textwrap.wrap(txt,190,break_long_words=True,break_on_hyphens=False):
  add(d,piece,'quote')
 total_chars+=len(txt)
section(d,'附录B：重复素材与版本关系','以下同一哈希的重复文件，仅以一份代表文件进入阅读摘录，保留副本目录可追溯。这种去重避免将同一个观点误计为多份独立证据。')
for pth,other in dups:add(d,pth.relative_to(ROOT).as_posix()+' ← 同源：'+other.relative_to(ROOT).as_posix())
add(d,f'本卷附录可提取文本摘录累计约{total_chars}字符；完整源文本及原始图像仍在原仓库对应目录中。')
d.save(OUT/f'{STAMP}_02_智烤管家全量项目深度研究卷.docx')

# 03: Official narrative anchored to 3 official documents
off=[x for x in sorted(OFFICIAL.glob('*.docx')) if '软著申请' not in x.name]
doc=init();title(doc,'智烤管家：官方叙事规范与决赛表达研究','1019/v1 · 以正式成果报告、技术资料说明书和预审推荐表为准')
section(doc,'前言：技术真实、业务准确、语言有温度','参赛答辩不是把研发文档搬上讲台，也不是用激昂话术替代工程事实。讲述必须同时回答“为什么做、怎么实现、谁来使用、效果怎样、怎么持续”。对现场的跨专业评委，建议以烟农三愁引出资源管理三缺；以1234回应需求；以真实试点指标结尾。')
figure(doc,['烟农三愁','管理三缺','1234协同','试点证据','可复制条件'],'决赛叙述以需求—方案—实证为轴')
narr=[
 ('一、项目的时代价值与业务位置','本项目的对象不是卷烟零售店，也不是烟叶烘烤工艺的自动调温，而是烟叶调制设施及相应的业务管理。把有限的烤房资源“用起来”，把新建和维修资金“投到点上”，把管护责任“落到人头”，是项目前后文始终一致的价值主张。官方名称为“智烤管家——基于健康评估的烤房用投管一体化解决方案”。'),
 ('二、原始冲突如何转成管理命题','烟农说“想用，调不到”，对应资源调度信息分散；说“想修，等不及”，对应抢修路径和部件信息不完整；说“想问，问不明”，对应状态、责任和进度不可追溯。最后收束成一句原稿非常有力量的话：资源无统筹、投入无量化、管护无闭环。'),
 ('三、一库见底应如何讲','数字化的起点应是“一房一档”。必须交代有多少座烤房、哪些状态和部件、谁拍照谁录入、维修后谁更新。2442座不是单纯一个“大数字”，说明实现了全口径基础归集；用原始照片、GPS和业务记录串联形成可核查的健康档案。22个低代码应用和6类报表可做技术补充，不宜连报多个系统产品名。'),
 ('四、两模不能只讲名字','健康评估模型回答“谁最该修”，由数据输入、部件评估、综合健康分与寿命估算、维修建议三个阶段构成。发展导引模型回答“哪里真缺、建在哪”，由数据入模、供需测算、候选点评价、情景推演四步组成。两个模型联系在于：存量烤房的健康分决定有效烤能，而有效烤能又影响新建需求判断。'),
 ('五、专业术语的正确出现方式','信息条线评委可以追问AHP与组合赋权及TOPSIS；烟叶条线评委想听的是这些方法如何纳入烟区种植、产能、管护与选址。推荐每个模型“先给生活比喻、再给计算逻辑、最后给决策输出”。健康评分是维修优先辅助信息，不能作为免现场检查的自动判定。'),
 ('六、三智落地必须把角色区分清楚','智用服务烟农：查房、调剂、一键报修、进度和评价；智投服务管理：风险与修复清单、候选点选址沙盘、资金论证；智管服务合作社与行业管理：派单、巡检、维修、验收、评价回写与责任监督。三端协同的核心不是三个图标，而是一个闭环——现场发生的数据回流到数据库，模型重新计算，下一次决策更稳。'),
 ('七、四支队伍为何是独有价值','软件上线不等于长期可用。为实际管理“最后一公里”组织调制、抢修、养护、综合四支队伍，由合作社、制度和基金共同支撑。重点是“队伍明确谁干、规范明确怎么干、平台证明干得怎么样、基金帮助长期干”。1363栋参管和基金总额近190万元说明机制已探索落地，但不等同于全部设施自动化。'),
 ('八、效果陈述遵守同口径','官方报告列出2442座全口径烤房入库、48次调剂、18次报修且抢修控制3小时内、16条平台工单平均响应0.38小时、健康分低于60分340多栋、531份问卷满意度96.8%、1363栋合作社统筹管护约占1672栋在用的81%。首版稿采用该口径，同时明确每个数量对应对象和时点。'),
 ('九、技术可信度与评委质询准备','需要准备数据清洗、重名编号、部件缺失、极端低分、健康分复算、寿命模型训练样本、TOPSIS逆向检验、维修资金审批、AI助手知识库边界等问题的答辩证据。对于尚未量化验证的“降低故障率、节省资金”只描述目标，不虚设百分比。'),
 ('十、成果推广的边界','“已在秭归一个烤季运行”是可核实的阶段；跨烟区推广依赖当地资产台账、部件标准、村级种植结构、合作社人员与管护基金机制。可复制的不只是软件，更是一套“数据—模型—平台—队伍”的共建方式。推广要先试点、再校准、后复制。'),
 ('十一、面向第一名的情绪曲线','开场要让观众看到烟农对烤房的急迫需求；三愁让观众点头；1234让观众记住结构；两模讲出知识含量；三智与四队讲出人味；实绩将情绪转成信服；结尾以“四句短语”回扣开篇的健康档案，避免没有证据的“全国首创、彻底杜绝”。'),
 ('十二、从官方书面语到现场口语的具体转换','书面语“基于层次分析法构建烤房健康综合评价体系”，上台说“给每座烤房做体检，算健康分，找出最该修的部位”；书面语“组合赋权与逼近理想解排序”，上台说“综合比较候选点，排出先后，再用沙盘算建几栋”；书面语“维修闭环与数据回写”，上台说“修完不算完，结果还要写回健康档案”。'),
]
for h,b in narr:section(doc,h,b)
section(doc,'官方文件逐份审读笔记','以下列出三份正式材料的核心提纲与可辨认的事实表述。不把前期讨论稿、技术草案和官方报送版放到同一可信度层级。')
for pth in off:
 add(doc,pth.name,'heading')
 s=tidy(read_any(pth));add(doc,'来源：'+pth.relative_to(ROOT).as_posix())
 # keep meaningful original source clauses, not all tables
 paras=[x.strip() for x in s.splitlines() if len(x.strip())>35]
 uniq=[];seen_lines=set()
 for pp in paras:
  if pp[:100] in seen_lines:continue
  uniq.append(pp);seen_lines.add(pp[:100])
 for pp in uniq[:110]:
  if len(pp)>300:pp=pp[:300]+'……'
  add(doc,'原文核对｜'+pp,'quote')
section(doc,'交付约束与校核表','这是一份10,000字级别叙述研究的v1结构；官方材料的完整逐字抽取文件另存同目录（03官方材料逐字提取.md）。正式答辩现场仅使用8分钟精简版，不能把这篇研究稿直接朗读。')
tbl(doc,['风险','处理原则'],[
 ('18个与12个部件','区分清查范围与健康模型所用计算参数，技术现场再次核对'),('0.38h与≤3h','平均响应时间与抢修完结时限严格分开'),('95.8%与96.8%','采用官方531份问卷、96.8%'),('150栋与7栋','未有验收资料不报成已完成实绩'),('视频1分钟与30秒','本次时间按约30秒开场视频配置'),('排名第一','以争取目标表述，不能预设获奖')])
doc.save(OUT/f'{STAMP}_03_官方叙述规范和决赛一万字研究稿.docx')

# 04: 复用原稿的页面、页边距与字体风格；不照搬旧段落。
d=Document(str(BASE))
body=d._element.body
for node in list(body):
 if node.tag!=qn('w:sectPr'):body.remove(node)
source=SP.read_text(encoding='utf-8')
blocks=[]
for ln in source.splitlines():
 ln=ln.strip()
 if ln.startswith('【P'): blocks.append(('cue',ln))
 elif ln.startswith('A：') or ln.startswith('B：') or ln.startswith('合：'):blocks.append(('speaker',ln.replace('**','')))
 elif ln.startswith('（'):blocks.append(('aside',ln))
p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run('智烤管家——基于健康评估的烤房用投管一体化解决方案');r.font.name='黑体';r.font.size=Pt(14);r.bold=True
for kind,line in blocks:
 if kind=='cue':
  # 页码为彩排提示，最终正文使用灰色小字号，以便说话时扫一眼；放在上一段或新段都不打断正文。
  p=d.add_paragraph();p.paragraph_format.space_before=Pt(4);p.paragraph_format.space_after=Pt(1)
  r=p.add_run(line);r.font.name='黑体';r.font.size=Pt(9);r.font.color.rgb=RGBColor(95,104,112)
 elif kind=='aside':
  p=d.add_paragraph();r=p.add_run(line);r.italic=True;r.font.size=Pt(10)
 else:
  p=d.add_paragraph();p.paragraph_format.space_after=Pt(4)
  who,txt=line.split('：',1)
  r=p.add_run(who+'：');r.font.name='黑体';r.font.size=Pt(14);r.bold=True
  rr=p.add_run(txt);rr.font.size=Pt(14);rr.font.name='仿宋_GB2312'
  rr._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'仿宋_GB2312')
d.save(OUT/f'{STAMP}_04_智烤管家8分钟双人决赛演讲净稿.docx')

# 05: 按25页PPT对照 + 问题纠错，单独方便后续改PPT
f=init();title(f,'智烤管家：25页PPT逐页口播与修改校核','本轮先定稿，不改PPT；下一轮根据此表统一修改画面')
pairs=[(1,'封面主题与双人亮相'),(2,'约30秒视频，无讲稿'),(3,'烟农三愁'),(4,'从三愁到管理三缺'),(5,'1234总览'),(6,'一库见底'),(7,'两模总览'),(8,'健康模型公式图：不细读公式'),(9,'健康模型三步走'),(10,'发展导引模型过程图：只讲功能'),(11,'数需评推四步'),(12,'业务入口'),(13,'三智运行闭环'),(14,'智用找房、调剂、报修'),(15,'智用实绩与评价'),(16,'智投修复优先级'),(17,'智投新建沙盘'),(18,'智管流程闭环'),(19,'AI助手辅助查询'),(20,'1363/81%/基金与四队'),(21,'调抢养综四队和制度'),(22,'官方成果指标'),(23,'三类创新价值'),(24,'推广条件'),(25,'结束寄语')]
sections=[];active=None
for k,line in blocks:
 if k=='cue':
  active={'head':line,'text':[]};sections.append(active)
 elif k=='speaker' and active:active['text'].append(line)
def pg_nums(cue):
 s=cue.split('】')[0]
 nums=re.findall(r'\d+',s)
 if not nums:return []
 if len(nums)>=2 and ('—' in s or '-' in s):
  return list(range(int(nums[0]),int(nums[1])+1))
 return [int(nums[0])]
spoken_map={}
for x in sections:
 for ix in pg_nums(x['head']):spoken_map[ix]='\n'.join(x['text'])
for page,brief in pairs:
 add(f,f'第{page}页｜{brief}','heading')
 text=spoken_map.get(page,'本页配合前后段落或静默放映；由转场/动画控制。')
 add(f,'台上口播：'+text[:620])
 if page==2:add(f,'需确认视频实际时长。如果大于30秒，必须从其他段落等量删减。')
 if page in [8,10]:add(f,'建议保留技术流程图，但讲述聚焦输入、算法为什么有用、输出是什么，不宣讲公式符号。')
 if page==22:add(f,'提示：官方报告48次调剂、18次报修、16条样本平均响应0.38小时、531份满意度96.8%；清楚区分口径。')
section(f,'彩排检查：按8分钟倒排','本次首版设开场视频约30秒。口述字数按简体汉字/标点统计，语速由A、B实际表达及停顿决定。最终必须双人完整彩排，凡超时优先压缩专业词解释与重复夸赞，不压缩关键数据证据。')
spoken='\n'.join(x[1] for x in blocks if x[0]=='speaker')
char_n=sum(1 for ch in spoken if '\u4e00'<=ch<='\u9fff')
add(f,f'口述中文汉字数：{char_n}；按每分钟240/260/280汉字分别需{char_n/240:.2f}/{char_n/260:.2f}/{char_n/280:.2f}分钟；再另加0.5分钟视频及换页/停顿时间。')
section(f,'主要文字修订理由','去掉原稿“95.8%”“抢修全由AI自动匹配”等口径风险；改用官方531份问卷96.8%、16条平均响应0.38小时；把150栋修复和肖家坪7栋建设标为待核计划而非已完成实绩；将第二模型讲成“数、需、评、推”；把原稿已具感染力的“匠心工坊”“资源无统筹、投入无量化、管护无闭环”择优保留；将四队独立成“4”，而非混入“三智”。')
f.save(OUT/f'{STAMP}_05_25页PPT逐页口播与改稿对照.docx')
print('DELIVERABLES:')
for x in sorted(OUT.glob(f'{STAMP}*.docx')):print(x.name,x.stat().st_size)
