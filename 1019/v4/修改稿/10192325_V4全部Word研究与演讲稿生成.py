# -*- coding: utf-8 -*-
"""1019/V4 原稿保持版式 + 高质量正式演讲稿 / 获奖学习 / 项目研究 / 官方研究 / PPT逐页 / 原稿对照。
不依赖外部自动生成文本；研究文档使用已逐份审读的V3研究卷，
新增的V4解读注明“本轮补充”，避免将原材料摘录伪装成新创作。
"""
from pathlib import Path
from collections import defaultdict
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import re, math, shutil
ROOT=Path(__file__).resolve().parents[3]
OUT=Path(__file__).resolve().parent
SRC=OUT/"10192315_04_智烤管家V4_双人决赛稿_PPT对照正文.md"
BASE=ROOT/"1004"/"智烤管家演讲讲稿9.30.docx"
OLD=ROOT/"1019"/"v3"
STAMP="10192325"
def han(s):return len(re.findall(r'[\u4e00-\u9fff]',s))
def add(d,t,style=None):
 p=d.add_paragraph(style=style or 'Normal');p.add_run(t);return p
def heading(d,t):return add(d,t,'Heading 2')
def table(d,heads,rows):
 t=d.add_table(rows=1,cols=len(heads));t.style="Table Grid";t.alignment=WD_TABLE_ALIGNMENT.CENTER
 for c,s in zip(t.rows[0].cells,heads):c.text=s
 for r in rows:
  cells=t.add_row().cells
  for c,s in zip(cells,r):c.text=str(s)
 return t
def page_title(d,s):
 p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
 r=p.add_run(s);r.font.name="黑体";r.font.size=Pt(16);r.bold=True
def intro(d,title,summary):
 page_title(d,title);add(d,summary)
def load_md():
 data=SRC.read_text(encoding='utf-8')
 groups=[];group=None
 for line in data.splitlines():
  if line.startswith("## 【"):
   group={"title":line[3:],"spoken":[],"notes":[]}
   groups.append(group)
  elif line.startswith(("A：","B：")):
   assert group, 'Speaker before heading'
   group["spoken"].append(line)
  elif line.startswith("（") and group:
   group["notes"].append(line)
 return groups
groups=load_md()
assert len(groups)==16,(len(groups),[i["title"] for i in groups])
lines=[x for g in groups for x in g["spoken"]]
A=sum(han(s[2:]) for s in lines if s.startswith("A："))
B=sum(han(s[2:]) for s in lines if s.startswith("B："))
TOTAL=A+B
assert 1480<=TOTAL<=1620,(A,B,TOTAL)
assert .69<=A/TOTAL<=.71,(A,B,A/TOTAL)
def pagenums(title):
 z=re.search(r"【P(\d+)(?:[—-]P?(\d+))?】",title)
 assert z,title
 return list(range(int(z.group(1)),int(z.group(2) or z.group(1))+1))
allpages=[v for g in groups for v in pagenums(g["title"])]
assert sorted(allpages)==list(range(1,26)),sorted(allpages)

# ==== 原格式正式稿：原9.30版式不更改，清除旧正文再写新内容
def new_from_original():
 d=Document(str(BASE))
 for node in list(d._element.body):
  if node.tag!=qn("w:sectPr"):d._element.body.remove(node)
 return d
HIGHLIGHTS=["智烤管家","资源无统筹、投入无量化、管护无闭环","一库见底、两模决策、三智落地、四队长效","数、需、评、推","五百三十一份有效问卷","百分之九十六点八","一库清家底，两模明方向"]
PATTERN="("+("|".join(map(re.escape,sorted(HIGHLIGHTS,key=len,reverse=True))))+")"
def split_bold(p,text):
 for st in re.split(PATTERN,text):
  if not st:continue
  r=p.add_run(st);r.font.name="仿宋_GB2312";r.font.size=Pt(14)
  r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"),"仿宋_GB2312")
  if st in HIGHLIGHTS:r.bold=True
def write_role(d,l):
 who,txt=l.split("：",1)
 p=d.add_paragraph();p.paragraph_format.space_after=Pt(4);p.paragraph_format.line_spacing=1.1
 r=p.add_run(who+"：");r.bold=True;r.font.name="黑体";r.font.size=Pt(14)
 r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"),"黑体")
 split_bold(p,txt)
def doc_title(d):
 p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
 r=p.add_run("智烤管家——基于健康评估的烤房用投管一体化解决方案")
 r.bold=True;r.font.name="黑体";r.font.size=Pt(14)
 r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"),"黑体")
def export_spoken(stage):
 d=new_from_original();doc_title(d)
 for g in groups:
  if stage:
   p=d.add_paragraph();p.paragraph_format.space_before=Pt(3)
   p.paragraph_format.space_after=Pt(1);p.paragraph_format.keep_with_next=True
   r=p.add_run(g["title"]);r.font.name="黑体";r.font.size=Pt(9)
   r.font.color.rgb=RGBColor(95,105,115)
   if g["title"].startswith("【P2】"):
    note=d.add_paragraph("（开场视频建议剪至30秒；播放期间AB不另读台词）")
    for r in note.runs:r.font.size=Pt(10)
  for line in g["spoken"]:write_role(d,line)
 path=OUT/f"{STAMP}_{'04_按原稿字体格式_正式净稿' if not stage else '05_25页PPT对应_双人舞台稿'}.docx"
 d.save(str(path));return path
speech_path=export_spoken(False)
stage_path=export_spoken(True)

# ==== 01 全部12篇获奖现场学习报告在V3基础上进一步深读，语言方法+具体口播转写
def retitle(doc):
 for p in doc.paragraphs[:10]:
  for run in p.runs:
   run.text=run.text.replace("智烤管家V2表达法","智烤管家V4表达法").replace("V3官方叙述","V4官方叙述")
 return doc
d=retitle(Document(str(OLD/"10192152_01_获奖解说词逐篇学习与改稿方法.docx")))
heading(d,"V4补充｜跨作品对标结果：让业务矛盾自己带出技术")
for title,desc in [
("3D智慧对标最应学习“现场问题—一眼懂的三问—可核对的效果”","技术再新，如果开头只有系统名，评委不会跟着走。V4模仿的是节奏而非措辞：从一炕烟的紧迫感开题，转为资源、投入与管护三个业务难，再以1234依次回答。"),
("AI驱动营销作品最应学习“双层表达”","先讲业务决策哪里困难，再以数据入模—计算逻辑—输出清单证明能力；专业词在紧跟生活化翻译之后出现，不连续念缩写。模型1不是一个“体检”比喻就够，还必须具体说出四类输入、加权比例、风险输出。"),
("善用双人问答，让B角替非烟叶评委发问","B不是配音员；最好的问题是评委正要问的问题。例如“先修谁？谁喊得急就先修谁？”与“哪里真缺、建在哪儿？”答在模型之后，落到维修清单和选址排序。"),
("现场气势靠确定的事实，不靠“彻底杜绝”","真实值比形容词有力量：2442入库、16条工单平均响应0.38小时、1363栋合作社管护。结尾可以激昂，过程仍须有可核查来源和范围。"),
("V4如何增强音乐性","只在转场和高潮处设置排比：“有房未必调得动，有钱未必投得准，有人未必管得久”“一库摸家底、两模做判断、三智办实事、四队护长远”。不强迫技术句押韵，以免读错AHP和TOPSIS。"),
]:
 heading(d,title);add(d,desc)
heading(d,"优秀稿十二篇→V4台词迁移清单")
table(d,["演讲环节","过去讲法的不足","V4表达原则"],[
("出场","一上台交代系统和术语","先请出“山里的烤房管家”，用拟人带入"),
("背景","把业务困难分散罗列","三愁→三缺→1234，一条主线"),
("模型1","只说健康分，不讲由来","四类输入、80/10/10、寿命、三种输出"),
("模型2","只说精准新建，缺推理路径","数需评推＋有效烤能＋组合赋权＋沙盘"),
("实效","数字反复念、缺行动画面","48/18/16/0.38与烟农实际获得的服务关联"),
("保障","只报基金规模，机制感弱","四队各自负责什么、制度和基金如何支撑"),
("结尾","口号密集、泛泛表态","用一库两模三智四队回扣收成与烟农")
])
d.save(str(OUT/f"{STAMP}_01_12份获奖解说词学习与V4借鉴.docx"))

# ==== 02 采用V3全资料研究卷：保留源路径及图像，追加V4技术完整对照，严谨区分研究正文与来源摘录
d=Document(str(OLD/"10192152_02_智烤管家约五万字全文件深度研究.docx"))
heading(d,"第二十一章 V4进一步澄清的算法源流与现场推理链")
for title,body in [
("21.1 资产信息是模型可信的边界","第一模型的可靠性先取决于数字档案是否能证明是同一座烤房：编号、地理位置、建造年份、部件现状、维修修复年、人工评价要有统一时点。低代码工具负责降低采集门槛，不意味着每条现场数据都天然正确；还需要字段校验、重复识别与业务审核。"),
("21.2 健康分与部件寿命不是同一个概念","最新PPT第8页图显示，针对十二个计算参数逐项考虑已用年限、指导寿命、状态系数，形成部件健康得分、风险贡献排序和整栋H值。综合H=0.80P_norm+0.10M+0.10E，是当前演示图中的表达。寿命计算结合关键短板寿命、部件加权平均寿命与健康系数，结果用于关注/维修提示，而不是精确到某日故障预测。"),
("21.3 PPT内部寿命表述一致性","PPT第9页“以Weibull distribution拟合剩余寿命”的一句与第8页展示的短板加权合成公式不完全是同一个计算路径。需要研发人员回答二者是并行、替代还是不同版本；V4演讲不指定尚未在官方材料中统一的分布参数，只说“结合使用年限估算剩余寿命”。答辩时必须带可核对的模型流程和参数说明。"),
("21.4 村级需求先问“有效”，再问“新建”","名义上有多少栋不等于真正可用烤能。需要按健康状况、地点、烟叶生产需求等条件将供给折算成有效烤能Q，再与理论需求D比较。若忽略可修复存量，容易误判为必须新建；若忽略空间距离，容易名义富余实际紧张。"),
("21.5 AHP、熵权和TOPSIS的分工","AHP吸收专家对候选点指标重要性的判断，熵权依据样本区分度给出客观权重；组合赋权后，TOPSIS用于形成候选点排序。排序是候选方案的对比工具而非行政许可；模型还需要以空间地图、建设数量模拟和现场踏勘作为实施约束。"),
("21.6 两模联动带来真正的管理创新","健康评估不是只决定当年维修清单，它会改变有效烤能Q，从而反过来改变千亩村是否真正需要新建、需要多少和应当在哪里建的结论。此联系是比单独叠加AHP与TOPSIS更值得现场强调的创新点。"),
("21.7 三智的改进来自“回写”而不只来自“在线”","智用中的查询、调剂和报修形成动态业务记录；智投引用模型和情景比较形成资金配置建议；智管在派单、维修、验收和评价后更新档案。缺少回写，健康分会滞后；没有线下队伍，工单只有状态、没有处置。"),
("21.8 四队使数字系统有线下执行能力","烟叶调制队、紧急抢修队、日常管护队和综合管理队分工不同，其绩效与责任须由标准流程与基金收支共同支撑。基金总额近190万元是资金治理规模指标，不应误当可用余额，必须与实际花费、结算、服务时效分开统计。"),
("21.9 指标口径与可扩展边界","2442是全口径建档对象；1672是本年度在用分母；1363是合作社统筹管护范围，覆盖在用的约81%；18条报修与16条可计算平均响应的工单不同；531份有效问卷与综合满意度96.8%成对出现。把不同数据统一说成“全部覆盖、全部完成”不符合正式成果。"),
("21.10 原稿案例的保留方式","磨坪站雷击抢修、150栋核心设备更新、肖家坪7栋新建、12478张现场影像、6704条评价、66人团队、以及12.6天缩短至11天等案例都出现在历史口述或底稿，但尚缺与最新版报送一致的独立验收或统计核对。它们可以作为团队后续补证的清单，不擅自移植到V4主台词。")
]:
 heading(d,title);add(d,body)
heading(d,"表｜V4决赛表述可追溯与不可越界事项")
table(d,["内容","可说的边界","追问准备"],[
("一库","2442座全口径建档，四大部位18个现场清查项","场景编号一致性、采集数据校验"),
("模型1","H分和寿命估计形成风险部位排序","P8的12项算法参数与P6的18项采集映射"),
("模型2","数需评推、Q/D、组合赋权TOPSIS和沙盘","权重来源、缺口标定及候选排序解释"),
("智用","48次调剂、18次报修、16条平均响应0.38h","抢修完成时长与响应时长区别"),
("智管","队伍、派单、验收、评价回写","工单是否闭环、回写更新时间"),
("四队","1363/1672约81%、基金近190万元","制度有效性、基金账面余额")
])
d.save(str(OUT/f"{STAMP}_02_五万字项目全文件研究与V4技术增订.docx"))

# ==== 03 官方叙述研究沿用已读所有官方文档，加入独立V4修辞与指标核查
d=retitle(Document(str(OLD/"10192152_03_官方文稿一万字研究与叙述规范V3.docx")))
heading(d,"十七、V4官方表述定稿条款：激情与准确如何兼容")
for t in [
"严守对象：项目专注烟叶调制设施“用、投、管”的资产管理与业务协同，并非自动烘烤温控机器人或卷烟营销智能投放系统。获奖稿提供表达样式，而不提供本项目事实证据。",
"严守范围：2442座为全口径建档范围；健康分低于60的340多座是预警研判对象；1363是统筹管护栋数，1672为本年度在用分母；基金近190万元指总额。",
"严守动作：一键报修、调剂申请、派单、评价回写可讲功能和运行记录；对与AI有关的查询提示，不应说成自动替代审批、验收和建设决策。",
"严守算法：现场讲出第一模型输入、AHP加权、80/10/10、剩余寿命及三项输出；第二模型讲数需评推、Q、D、AHP＋熵权、TOPSIS及沙盘。P8与P9寿命算法出现不同表述，需要团队统一。",
"严守成效：16条平台工单平均响应0.38小时并不是18次报修的平均维修时间；满意度采用531份有效问卷96.8%的官方口径，不能延用9.30原稿95.8%。",
"演讲语言不虚夸：少用“彻底解决、全面消灭、全国第一”等没有对照或检索证据的词，增强力度的办法是三次事实递进：有没有档案、有没有判断、有没有队伍和资金。",
"舞台结构：前40秒让评委同情烟农，紧接用三缺建立问题张力；1234之后每介绍一种能力都立即给出业务用途；结尾将一库、两模、三智、四队重新收束为“让烟农多一分踏实”。"
]:add(d,t)
d.save(str(OUT/f"{STAMP}_03_官方叙述规范一万字V4校订版.docx"))

# ==== 06 25页对照、AB定量节奏与视频占用
d=Document();sec=d.sections[0];sec.left_margin=Cm(2.4);sec.right_margin=Cm(2.4)
page_title(d,"V4双人决赛稿：25页PPT口播及彩排对照")
add(d,f"口述汉字{TOTAL}，A角{A}（{A/TOTAL*100:.2f}%），B角{B}（{B/TOTAL*100:.2f}%）。以口述每分钟230字估算{TOTAL/230*60:.0f}秒，30秒视频和20秒换页共{TOTAL/230*60+50:.0f}秒；必须以实际放映验证。")
slides=[
"封面与项目身份","开场情景视频","烟农三愁：急、难、问不明","三缺及问题归因","1234总体亮相",
"建档、部件标准、2442入库","双模回答修谁与建哪","健康模型四类输入与算法图","健康模型三步与风险部位输出",
"千亩村发展导引算法图","数需评推，Q/D和TOPSIS","三类使用者入口","三智与回写闭环",
"智用：找房调剂报修视频","48、18、16条及0.38小时","智投资金与维修优先级","新建候选点地图沙盘",
"智管：派单、处置、验收、评价","AI助手辅助查询","1363、1672、81%、190万及四队",
"调抢养综与两套制度","运行成效、531问卷与96.8%","创新价值：三个转变","推广边界和条件","结尾口号"]
mapping={}
for g in groups:
 for i in pagenums(g["title"]):
  mapping[i]=g
for i,label in enumerate(slides,1):
 heading(d,f"第{i}页｜{label}")
 lines=mapping[i]["spoken"]
 if i==2:add(d,"【静默30秒视频，不加台词；原素材约41.5秒，须剪至30秒】")
 else:
  for l in lines:
   if i in pagenums(mapping[i]["title"]) and i!=pagenums(mapping[i]["title"])[0]:
    break
   add(d,l)
  if not lines or (i!=pagenums(mapping[i]["title"])[0]):
   add(d,"本页为同一组口述的图形/界面支撑；不增加重复台词。")
 if i in (8,9):add(d,"技术校核：P8短板/加权寿命公式和P9 Weibull文字需统一；台上不擅自声称具体拟合精度。")
 if i in (14,16,17,19):add(d,"放映建议：该页含视频或演示素材，避免与口述独立完整播放而超时。")
table(d,["内容","时长预算"],[
("A＋B口述",f"{TOTAL/230*60:.1f}秒，基于230个汉字/分钟"),
("开场视频","30秒，需剪辑"),
("25页换页及重要停顿","20秒，按连续播控"),
("预排总计",f"{TOTAL/230*60+50:.1f}秒"),
("严格约束","≤480秒；语速较慢或视频未裁剪需再减字")
])
d.save(str(OUT/f"{STAMP}_06_25页PPT逐页口播与8分钟计时表.docx"))

# ==== 07 9.30原稿 vs V4实质改稿 + 备答
d=Document();page_title(d,"V4修改对照：9.30原稿优点如何保留")
table(d,["9.30初稿要点","V4处理","理由/证据"],[
("匠心工坊、一年的收成","保留，P3用烟农三愁起势","让非烟叶评委先看到人，再看到技术"),
("资源无统筹、投入无量化、管护无闭环","基本保留于P4","原稿最有力的行业问题概括"),
("四大部位、18个零部件","保留现场清查18项，算法另有12个参数","避免不同计算口径混用"),
("部件80%＋管护10%＋评价10%","恢复到P8—P9口播","健康分技术不能被“AI体检”一句带过"),
("模型2熵权与TOPSIS","补充数需评推、有效烤能Q、理论需求D、AHP组合赋权和沙盘","让方法解释形成投入决策的逻辑链"),
("150栋维修已完成、肖家坪7栋已新建","列待核，不进主稿的完成性结果","避免把计划写成正式验收"),
("磨坪雷击案例和12478张影像","保留为可补证的答辩案例，不直接报数字","需与官方台账逐项核对"),
("综合满意度95.8%","采用531份问卷96.8%","以官方成果报告为准"),
("0.5小时响应","采用16条工单0.38小时","区分样本与响应指标"),
("四队、产权公有、专业管护、有偿服务、基金","保留；补充四队职责","工程落地与组织创新都不可少"),
("基金年底结余100万元","采用近190万“总额”","不能把预计余额与基金总额混同"),
("纵向推广全市、横向育苗棚","保留愿景，措辞改为先校准后推广","秭归一个烤季试点不等于全市完成")
])
heading(d,"答辩中的三个重点技术追问")
for a in [
"问：18项清查、12项模型参数是否冲突？答：分别属于现场清查与算法计算层级，需要提供字段映射；当前PPT没有给出完整映射，不能直接等同。",
"问：PPT第8页与第9页寿命估算究竟用什么方法？答：第8页给出状态/短板/加权合成，第9页提及Weibull拟合，现场前应由研发组确定最终实施方法与参数样本，避免双版本并存。",
"问：推广到兴山、五峰需要什么？答：应先完成资产字段、部件映射、村级理论烤能、合作社队伍及基金机制的本地校准，不能只拷贝一套后台。"
]:add(d,a)
d.save(str(OUT/f"{STAMP}_07_初稿与V4逐项修改及评委备答.docx"))

review=[
"# V4事实复核与文件校验",
f"- A角：{A}口述汉字；B角：{B}口述汉字；总数：{TOTAL}；比例：{100*A/TOTAL:.2f}%:{100*B/TOTAL:.2f}%",
f"- 230汉字/分钟＋30秒开场视频＋20秒换页：{(TOTAL/230*60+50):.1f}秒。",
"- 25页已全部建立对应位置；P8/P9算法路径需赛前核对。",
"- 注意：此时长是假设的彩排预算，不是现场实测。",
"- V4研究材料保留V3已整理的源文件证据摘录、图片和相关页码，并增加V4技术校订章节，不将重复原始摘录冒充原创研究。",
"- 9.30初稿的原始页设置、角色字体原则保留；Word渲染检查由工作流完成。",
"- 具体原稿旧数字与官方报告不一致项均在初稿对照中标注。"
]
(OUT/f"{STAMP}_README_交付与口径复核.md").write_text("\n".join(review),encoding='utf-8')
print("\n".join(review))
for file in sorted(OUT.glob(f"{STAMP}*.docx")):
 d=Document(file);print("READY",file.name, file.stat().st_size, "paragraphs",len(d.paragraphs),"images",len(d.inline_shapes))

# Run V4 Word builder after workflow setup.

# Rebuild with updated table-rich QA check.
