# -*- coding: utf-8 -*-
"""2026-10-20 fresh V6. Rebuild only after reading new source, with research evidence attributed."""
from pathlib import Path
from collections import Counter
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re, shutil, tarfile, hashlib, zipfile
ROOT=Path(__file__).resolve().parents[2]
BASE=ROOT/'.github/v6_2000_release'
V5=ROOT/'1019/v5/修改稿'
PARENT=ROOT/'1019/v6'
OUT=PARENT/'修改稿'
STAMP='10201048'
FULL=BASE/f'{STAMP}_内容完整版.md'
SAFE=BASE/f'{STAMP}_8分钟彩排版.md'
def han(s): return len(re.findall(r'[\u4e00-\u9fff]',s))
def parse(path):
    units=[]
    for line in path.read_text(encoding='utf-8').splitlines():
        if line.startswith('## P'):
            m=re.match(r'^## (P\d+(?:[—-]P\d+)?)｜(.+)',line)
            if not m:raise ValueError(line)
            units.append([m.group(1),m.group(2),[]])
        elif line.startswith(('A：','B：')):
            if not units:raise ValueError('Speaker line before heading')
            units[-1][2].append((line[0],line[2:]))
    return units
f=parse(FULL);s=parse(SAFE)
def counts(x):return (sum(han(t) for _,_,ls in x for r,t in ls if r=='A'),sum(han(t) for _,_,ls in x for r,t in ls if r=='B'))
assert counts(f)==(1384,594),counts(f)
assert counts(s)==(1098,474),counts(s)
assert '作为评委' not in FULL.read_text() and '作为评委' not in SAFE.read_text()
def page_ids(header):
    m=re.fullmatch(r'P(\d+)(?:[—-]P?(\d+))?',header)
    if not m:raise ValueError(header)
    return list(range(int(m.group(1)),int(m.group(2) or m.group(1))+1))
assert sorted(p for p,t,ls in f for p in page_ids(p))==list(range(1,26))
assert sorted(p for code,t,ls in s for p in page_ids(code))==list(range(1,26))
if PARENT.exists():shutil.rmtree(PARENT)
OUT.mkdir(parents=True)
def font(r,name,size,bold=False):
    r.bold=bold;r.font.name=name;r.font.size=Pt(size)
    r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),name)
def add(d,text,style=None):
    return d.add_paragraph(text,style=style)
def h(d,text):return add(d,text,'Heading 2')
def table(d,heads,rows):
    t=d.add_table(rows=1,cols=len(heads));t.style='Table Grid'
    for c,x in zip(t.rows[0].cells,heads):c.text=str(x)
    for row in rows:
        c=t.add_row().cells
        for x,v in zip(c,row):x.text=str(v)
    return t
def speech(units,file_name,show_pages=False):
    d=Document(ROOT/'1004/智烤管家演讲讲稿9.30.docx')
    for el in list(d._element.body):
        if el.tag!=qn('w:sectPr'):d._element.body.remove(el)
    if show_pages:
        d.sections[0].top_margin=Cm(1.65);d.sections[0].bottom_margin=Cm(1.65)
    title=d.add_paragraph();title.alignment=WD_ALIGN_PARAGRAPH.CENTER
    font(title.add_run('智烤管家——基于健康评估的烤房用投管一体化解决方案'),'黑体',14,True)
    for code,heading,lines in units:
        if show_pages:
            p=d.add_paragraph();p.paragraph_format.keep_with_next=True;p.paragraph_format.space_before=Pt(3)
            p.paragraph_format.space_after=Pt(0)
            r=p.add_run(code+'｜'+heading);font(r,'黑体',9,True);r.font.color.rgb=RGBColor(79,103,132)
            if code=='P2':add(d,'（静默播放30秒开场视频，不叠加其他视频；括号提示不朗读）')
        for role,text in lines:
            p=d.add_paragraph();p.paragraph_format.space_after=Pt(2 if show_pages else 4)
            p.paragraph_format.line_spacing=1.07 if show_pages else 1.13
            font(p.add_run(role+'：'),'黑体',14,True)
            font(p.add_run(text),'仿宋_GB2312',14)
    path=OUT/(STAMP+'_'+file_name+'.docx')
    d.save(path);return path
full=speech(f,'04_V6_约2000字专业完整版')
safe=speech(s,'05_V6_8分钟彩排净稿')
stage=speech(s,'06_V6_25页PPT对应舞台稿',True)

# Source-grounded review of all 12 excellent speech transcripts and separate QC delivery.
award=Document()
award.add_heading('获奖现场解说词及QC发布方法逐篇学习·V6重审',0)
award.add_paragraph('原始学习对象：读取资料/1学习优秀的中12份txt；另读仓库根目录QC成果发布词。下文只迁移表达方法，不挪用其他项目业务数据。')
reference={
'1.txt':('3D智慧对标','三大管理痛点→数据洞察→诊断→施策，既讲方法也讲时效，适度双人问答'),
'2.txt':('网络情报研判','开场短场景制造问题，然后回到数据、方法、业务实效'),
'3.txt':('AI驱动卷烟供需调控','9个指标/虚拟网格/算法步骤等技术过程清楚，效果以具体例子证明'),
'4.txt':('监管异常识别','业务冲突贯穿建模与风险识别，而非堆砌算法名'),
'5.txt':('物流寄递研判','采集—研判—应用形成闭环'),
'6.txt':('会思考的食堂','双人接力、生活化类比，但不削弱数据与责任'),
'7.txt':('滞销品规管理','用反差、对仗构造记忆点，转成可实施措施'),
'15.txt':('烟叶质量场景','烟区基层故事与专业方案互相验证'),
'19.txt':('财务数据治理','重复台账、管理责任倒推数字化功能'),
'30.txt':('雪茄三问','用连续问题推动论证，一问一答层层深入'),
'9月9日我的.txt':('早期本项目思路','继承业务主线，去掉不成熟完成性口号'),
'yus07.txt':('雪茄三个没数','短句识别主要矛盾，形成有节奏的过渡'),
}
src=ROOT/'读取资料/1学习优秀的'
discovered=sorted(src.glob('*.txt'))
assert len(discovered)==12,(len(discovered),[p.name for p in discovered])
for p in discovered:
    name,lesson=reference.get(p.name,(p.stem,'用真实项目材料形成因果论证'))
    txt=p.read_text(encoding='utf-8',errors='replace')
    h(award,f'{p.name}｜{name}')
    add(award,f'原稿规模：{han(txt)}个汉字，{len(txt.splitlines())}行。')
    add(award,'讲法分析：'+lesson)
    add(award,'对本项目的迁移：由烟农“三愁”归因到管理“三缺”，随后用“1库、2模、3智、4队”逐层回答；B角自然追问，不自称评委。')
    phrases=[x.strip() for x in txt.splitlines() if x.strip()]
    add(award,'可供复核的开头语料摘记：'+'；'.join(phrases[:8])[:220])
qc=ROOT/'缩短烤房设备单次维护作业耗时_QC成果发布词0907.docx'
assert qc.exists()
q=Document(qc)
qtext='\n'.join(p.text for p in q.paragraphs)
h(award,'中质协15分钟QC成果发布词专项借鉴')
add(award,'核心方法：现状调查与问题归因→目标和措施→结果验证→标准化巩固；B角色可围绕业务难度和实效自然追问。')
add(award,'本项目8分钟只借鉴证据组织与AB衔接，不把QC的维护耗时、专利或经济收益挪为智烤管家实效。')
add(award,'QC原稿可读取文字总量：'+str(han(qtext))+'汉字。')
award.save(OUT/f'{STAMP}_01_12篇优秀解说词与QC方法学习.docx')

# Deep research retains previously sourced 53k-level evidence volume; new V6 appends original-first checks.
deep=Document(V5/'10192342_02_V5_项目五万字级全文件研究与双模创新增订.docx')
h(deep,'V6原创修订专题：从9.30原稿恢复具体业务动作')
for title,text in [
('现场采集','手机拍照、GPS定位、表单录入和字段校验对应PPT第6页“现场采集—标准判定—一房一档”，不是把旧表复制到平台。'),
('18个现场清查项与12个计算参数','现场四大部位18个清查项与健康评估图中的12项权重/寿命计算参数属于不同层级，正式技术资料需要维护映射表。'),
('健康模型','输入档案、部件、评价、维修等信息；AHP构建部件权重，综合健康分80/10/10，估算寿命，输出风险烤房和部位TOP3以及维修次序。'),
('千亩村发展导引','按“数需评推”折算有效烤能Q与理论需求D，AHP+熵权组合赋权、TOPSIS排序候选点，地图沙盘模拟建设数量与位置。'),
('在线服务及双维评价','小程序查询、预约/调剂、定位报修、人员与备件匹配、维修流程跟踪，烟农对烤房设备与服务人员进行反馈。'),
('长效自治','县级1175栋、村级188栋由合作社统筹管护，四队职责、两套制度和基金提供长效服务保障。'),
('证据边界','旧原稿12478张设备影像、6704条评价、150栋改造、肖家坪7栋、66人等未在9月10日官方成果核验章节逐项确认；新V6正文不视作完成结果。'),
('PPT寿命技术核查','第8页寿命合成图与第9页Weibull表述须请研发团队核实实际实施版本，避免虚构模型精度。')
]:
    h(deep,title);add(deep,text)
# Inventory the complete actual Part-2 directory rather than asserting unseen files were read thoroughly.
folder=ROOT/'读取资料/2我的项目相关的零碎的'
items=sorted(p for p in folder.rglob('*') if p.is_file())
h(deep,f'附录：资料夹2实际文件清单（共{len(items)}份）')
table(deep,['资料相对路径','类型','体量(Byte)'],[(str(p.relative_to(folder)),p.suffix,p.stat().st_size) for p in items])
add(deep,'此清单是完整文件盘点，不将元数据盘点冒充逐字专业评审。此前研究卷含多轮正文与图片证据摘录。')
deep.save(OUT/f'{STAMP}_02_五万字项目研究底稿及V6实效增订.docx')
official=Document(V5/'10192342_03_V5_官方叙述一万字级研究校订版.docx')
h(official,'V6正式稿与原稿数据之间的可核验边界')
for t in [
'2442座全口径建档、22个低代码应用、6类报表和4大部位18项清查来自PPT/官方报告。健康模型使用12项参数，不能把现场清查范围直接当算法输入维度。',
'18次报修是年度业务记录；16条平台工单平均响应0.38小时有明确样本。响应不等于维修完成。',
'340余座为健康分低于60分的风险/预警对象，不能说已全部升级修复；150栋更新“完成”与“计划”口径未充分统一。',
'531份有效问卷综合满意度96.8%为官方成果，不能沿用原稿95.8%的旧说法。',
'县级1175栋+村级188栋=1363栋，覆盖当年在用1672栋的约81%；190万元是基金总额，不等于期末可支配余额。',
'创新表达保留AHP、熵权法、TOPSIS、有效烤能和地图沙盘，但所有排序都是辅助决策，仍由行业机制审批与现场核验。',
'原稿中的设备影像12478张、评价6704条、雷击维修案例等须以现场采集日志、工单或验收记录补证后再作为正式发布业绩。'
]:add(official,t)
official.save(OUT/f'{STAMP}_03_一万字官方叙述研究及成果数字校订.docx')

# 25-slide correspondence from full deck structure.
slides={
1:'封面',2:'30秒视频',3:'烟农三愁',4:'管理三缺',5:'1234方案',6:'1库采集与建档',
7:'双模总览',8:'健康评分及寿命公式',9:'健康模型3步',10:'千亩村模型原理',
11:'数需评推技术路线',12:'3类使用者',13:'三智协同与数据回写',14:'烟农小程序',
15:'48次调剂、18次报修与0.38小时',16:'维修清单',17:'地图沙盘',
18:'智管闭环',19:'AI助手',20:'1363栋与190万基金',21:'四队和两套制度',
22:'综合运营实效',23:'创新价值',24:'推广机制',25:'谢幕'}
map_={}
for code,heading,ls in s:
    for p in page_ids(code):map_[p]=(code,heading,ls)
d=Document();d.add_heading('V6双人稿｜最新25页PPT逐页对应与8分钟计时',0)
add(d,'安全版口述1572个汉字，A角1098、B角474，A占69.85%。以220字/分钟+30秒视频+20秒换页预计478.7秒，离8分钟仅1.3秒，必须实测；建议加快至230字/分钟或精简节奏。')
for p in range(1,26):
    h(d,f'第{p}页｜{slides[p]}')
    code,heading,ls=map_[p]
    if p==2:add(d,'静默播放剪至30秒的视频')
    elif p==page_ids(code)[0]:
        for r,text in ls:add(d,r+'：'+text)
    else:add(d,'与前页合并讲解，不增加重复口播。')
    if p in (8,9):add(d,'技术组重点核对：P8剩余寿命合成图与P9 Weibull文字表述的现行关系。')
    if p in (14,16,17,19):add(d,'如需独立播放演示视频，须从口播预算中等量减少时间。')
d.save(OUT/f'{STAMP}_06_25页PPT逐页口播与八分钟计时.docx')

# Source audit includes explicit checkmarks and acknowledged infeasible simultaneously.
d=Document();d.add_heading('V6：9.30原稿复核、官方数字边界与全部需求验收',0)
rows=[
('9.30原稿为主体','完整重写；保留匠心工坊、三缺、现场采集、双维评价、服务场景与四队'),
('1234主线','1库／2模／3智／4队一体化贯穿25页'),
('一库怎么采集','拍照、GPS、表单、编号、4部位18项、正常损坏缺失、字段校验与评价回写'),
('双模技术','AHP、12参数、80/10/10、寿命、TOP3、有效烤能Q、需求D、AHP熵权、TOPSIS、地图沙盘'),
('数据与实效','2442、340余、48、18、0.38/16、531/96.8%、1175+188=1363/1672/81%、基金近190万'),
('后续数据差异','12478影像、6704评价、150栋升级、7栋新建、66人等原稿数字在备考核材料中保留待核，不冒充官方已完成'),
('AB比例','完整版1384:594=69.97%:30.03%；安全版1098:474=69.85%:30.15%'),
('2000字和8分钟矛盾','完整版1978汉字含全部细节，正常语速会超8分钟；另制作1572汉字安全版，220字/分钟预计7分59秒仍需彩排'),
('QC稿风格','B角自然追问，专业而不突兀，无“作为评委”的身份式对白'),
('版式','沿用9.30原稿Word页设置与14磅仿宋+黑体角色字'),
('奖项学习','12篇txt逐份阅读并生成迁移总结，QC项目数据不混为本项目成果'),
('5万+1万研究','在已整理的完整研究与官方叙述底稿基础上增订并标明继承性，不宣称本轮5万字全新原创'),
('路径与覆盖','1019/v6/修改稿清空旧版后只保留本轮结果，25页全部覆盖'),
('PPT是否修改','本轮先定双人稿，不覆盖用户的最新版PPT')
]
table(d,['要求/检查点','落实说明'],rows)
h(d,'最终仍须用户或项目组核实的事项')
for item in [
'第8页健康模型寿命合成公式与第9页Weibull拟合的实际对应关系。',
'12478张设备影像、6704条评价、150栋更新和肖家坪7栋是否有最新系统日志或验收单。',
'30秒视频是否已真正剪辑；P14/P16/P17/P19嵌入的视频演示是否需要额外播放。',
'彩排语速至少要达到220汉字/分钟，且切换/停顿严格控制，安全版才能接近8分钟。'
]:add(d,item)
d.save(OUT/f'{STAMP}_07_原稿官方数据对照与逐项需求验收.docx')
assert len(list(OUT.glob('*.docx')))==8
(OUT/f'{STAMP}_08_完成说明.md').write_text(
    'V6 从9.30重新写作。\n正式内容完整版口述1978个汉字，安全彩排版1572个，AB均约70:30。'
    '\n5万/1万研究卷是来源归档及本轮增订，不是全部新写。'
    '\n旧V6目录已清理，PPT原件未改。\n',encoding='utf-8')
(OUT/f'{STAMP}_完整版正文.md').write_text(FULL.read_text(),encoding='utf-8')
(OUT/f'{STAMP}_8分钟安全版正文.md').write_text(SAFE.read_text(),encoding='utf-8')
archive=OUT/f'{STAMP}_智烤管家V6_八份Word完整包.tar.gz'
with tarfile.open(archive,'w:gz') as tf:
    for p in sorted(OUT.glob('*.docx')):tf.add(p,arcname=p.name)
    tf.add(OUT/f'{STAMP}_完整版正文.md',arcname='完整版正文.md')
    tf.add(OUT/f'{STAMP}_8分钟安全版正文.md',arcname='8分钟安全版正文.md')
print('READY',len(list(OUT.glob('*.docx'))),archive.stat().st_size,counts(f),counts(s),len(items))

# Trigger definitive V6 2000-character reset build.

# Re-trigger corrected 8-document release gate.
