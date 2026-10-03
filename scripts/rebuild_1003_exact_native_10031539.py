from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pathlib import Path
from PIL import Image
import numpy as np, cv2, io, re

STAMP="10031539"
ROOT=Path('.')
BASE=ROOT/'决赛发布材料/1002PPT决赛完善版/1003'
SRC_A=BASE/'智烤管家_142106_1比1蓝白山海光路抠图还原版_副本3.pptx'
SRC_B=BASE/'智烤管家_142106_1比1蓝白山海光路抠图还原版_副本2.pptx'
OUT=ROOT/f'智烤管家_{STAMP}_1003原图1比1字效原样+文字可编辑下层+图标分层+背景无光路版.pptx'
TMP=Path('/tmp')/f'zhikao1003_{STAMP}'; TMP.mkdir(exist_ok=True)

T={
1:{1:'Z H I  K A O  M A N A G E R',2:'智烤 管家',3:'让每一座烤房，都有一份可计算的“健康档案”',4:'一库清底  |  双模决策  |  三智落地  |  四队长效',5:'基于健康评估的烤房用管管一体化解决方案',6:'湖北烟草创客大赛  ·  决赛发布',7:'宜昌市烟草公司  ·  智同道合创客小组',8:'2026.10'},
2:{1:'01',2:'02',3:'03',4:'烟区一线，最怕这 3 件事',5:'不是“有没有烤房”，而是分配、抢修、投入仍靠经验和人工',6:'分不到',7:'旺季状态不透明｜烟农无法自主选择｜临时需求难调剂',8:'修不及',9:'故障逐级上报｜人员靠电话找｜配件靠人工调',10:'投不准',11:'哪栋该修｜哪个村真缺｜新建在哪里，缺少量化依据',12:'资源调度',13:'烤房在哪、是否空闲、谁在使用',14:'缺少一张动态资源图',15:'抢修服务',16:'故障发生以后才“找人找件”',17:'时间被消耗在协调上',18:'投入决策',19:'把“有多少栋”当“有多少能力”',20:'容易造成供需错配',21:'表面：用房、抢修、投入',22:'根因：“用—投—管”没有形成数据闭环'},
3:{1:'02',2:'三愁背后，是三条链断了',3:'问题不在“有没有烤房”，而在资源、维修、投入没有形成联动闭环',4:'01',5:'用房管理断链',6:'空闲不清 / 状态不明 / 调度靠经验',7:'维修保障断链',8:'故障发现慢 / 找人找件慢 /',9:'闭环追踪弱',10:'03',11:'投入决策断链',12:'真实需求不清 / 供需错配 /',13:'新改建难量化',14:'根因：“用—投—管”三条链没有被数据打通'},
4:{1:'一套方案，做成 4 件事',2:'从摸清家底，到科学决策，再到真实落地与长效机制',3:'02 双模决策',4:'健康评估 / 供需推演 / 投入更准',5:'03 三智落地',6:'智用 / 智投 / 智管',7:'04',8:'四队长效',9:'专班推进 / 维修服务 / 运营管护 / 基金保障',10:'01 一库清底',11:'一房一档 / 状态可视 / 家底清楚',12:'把“看不清、修不及、投不准”，变成“看得见、管得住、投得准”'},
5:{1:'一库清底：先把家底摸清楚',2:'让每一座烤房，都有一份动态更新的健康档案',3:'01',4:'现场采集',5:'位置、容量、状态、设备、使用情况',6:'02',7:'03',8:'标准判定',9:'统一口径 / 健康评分 / 风险识别',10:'一房一档',11:'动态更新 / 可查可用 / 一库清底',12:'空闲可用\n156座',13:'空闲可用\n156座',14:'烤房档案',16:'正常运行\n682座',17:'正常运行\n682座',19:'故障停用\n28座',20:'故障停用\n28座',21:'YN-5301',22:'正常运行',23:'位置',24:'XX县XX镇\n2座',25:'容量',26:'3吨',27:'设备',28:'热风炉 / 自控系统',29:'使用情况',30:'当前空闲',31:'健康评分',32:'92\n分',33:'分',34:'设备状态',35:'运行规律',36:'维护记录',37:'综合评分',38:'95',39:'90',40:'88',41:'92',42:'结果：空闲房在哪里、故障房是谁、可用能力有多少，一眼看清'},
6:{1:'模型① 烤房健康评估：从打分到估寿',2:'不是给一座烤房一个分数，而是判断还能用多久、哪里先修',3:'01',4:'状态标准化',5:'12类部件 / 状态采集 / 指标归一',6:'加热系统',7:'通风系统',8:'温湿度控制',9:'电气系统',10:'围护结构',11:'02',12:'AHP组合评分',13:'权重确定 / 综合评价 / 风险分级',14:'设备状态\n0.28',15:'设备状态\n0.28',16:'结构安全\n0.07',17:'结构安全\n0.07',18:'运行环境\n0.18',19:'运行环境\n0.18',20:'故障历史\n0.20',21:'故障历史\n0.20',22:'使用年限\n0.15',23:'使用年限\n0.15',24:'维护条件\n0.12',25:'维护条件\n0.12',26:'03',27:'Weibull估寿',28:'剩余寿命 / 失效概率 / 维修窗口',29:'生存概率\nR(t)',30:'R(t)',31:'1.0',32:'0.8',33:'0.6',34:'0.4',35:'0.2',36:'0',37:'预计剩余寿命',38:'6.8年',39:'10',40:'12',41:'6',42:'8',43:'使用年限（年）',44:'健康等级',46:'风险部位',47:'维修优先序',49:'评分只是中间量，真正输出的是寿命、风险与维修次序'},
7:{1:'模型② 千亩村发展导引：把缺口算出来',2:'不是问“想建几栋”，而是先算有效烤能，再判断真实缺口',3:'01',4:'02',5:'03',6:'算有效烤能',7:'判真实缺口',8:'可用烤房 / 健康折损 / 产能折算',9:'烟叶需求 / 可用产能 / 缺口识别',11:'做建设推演',12:'维修优先 / 调剂优先 / 新建情景',13:'调剂能力\n168栋',14:'可用烤房\n168栋',15:'烟叶需求（吨）',16:'可用产能（吨）',17:'现有烤房\n1,256栋',18:'1,256栋',19:'建设推演路径',20:'可用烤能\n682栋',21:'682栋',22:'缺口',23:'健康折损\n213栋',24:'213栋',25:'2024',26:'2025',27:'2026',28:'2027',29:'2028',30:'先调剂',31:'跨村调剂\n统筹使用',32:'统筹使用',33:'再修复',34:'优先维修\n提升健康度',35:'提升健康度',36:'仍有结构性缺口',37:'才新建',38:'按需新增\n合理布局',39:'合理布局',40:'把“有多少栋”改成“有多少有效烤能”，建设才有依据'},
8:{1:'双模决策：存量评估 + 增量推演',2:'把“经验判断”变成“量化决策”',3:'模型1',4:'健康评估',5:'模型2',6:'供需推演',7:'部件体检',8:'发现隐患\n量化评分',9:'量化评分',10:'寿命预测',11:'剩余寿命',12:'风险预警',13:'维修优先序',14:'轻重缓急\n科学排序',15:'科学排序',16:'有效烤能',17:'可用产能',18:'动态测算',19:'双模耦合输出',20:'真实缺口',21:'需求对比',22:'缺口识别',23:'建设情景模拟',24:'多情景推演',25:'投入产出评估',26:'该修哪栋',27:'该补哪里',28:'该建多少',29:'让维修有依据、投入有方向、建设有测算'},
9:{1:'01',2:'三智落地：把系统能力变成日常动作',3:'不是做一个平台，而是让智用、智投、智管真正进入业务现场',4:'02',5:'03',6:'智用',7:'空闲可见',8:'一键调度',9:'状态可视',10:'智投',11:'需求识别',12:'缺口测算',13:'投入更准',14:'让资源调得动、投入算得清、运行管得住',15:'智管',16:'运行监测',17:'维保留痕',18:'预警闭环'},
10:{1:'智用：调度与抢修，真正落到烟农手里',2:'从“打电话找房、找人、找配件”，变成“一键找房、发现可视、结果回写”',3:'01',4:'02',5:'找空闲房',6:'状态可见 / 就近可选',7:'发起调度',8:'线上申请 / 一键提交',9:'03',11:'故障报修',12:'定位问题 / 自动派单',13:'调度申请',14:'选择烟房',15:'选择人员',16:'选择配件',18:'一键提交',20:'少跑腿',21:'少等待',22:'少协调',23:'04',24:'评价回写',25:'过程留痕 / 服务闭环',26:'服务评价',27:'服务及时',28:'维修专业',29:'问题解决',30:'非常满意',31:'智用的价值，不是多一个入口，而是把调度和维修变成可追踪流程'},
11:{1:'智投：把钱投到最该投的地方',2:'存量先体检，增量再推演，选址和投入都用数据说话',3:'01',4:'存量体检',5:'健康等级 / 寿命 / 风险',6:'健康等级',7:'优秀',8:'良好',9:'一般',10:'较差',11:'32%',12:'45%',13:'18%',14:'5%',15:'寿命预测',16:'02',18:'缺口识别',19:'有效烤能 / 需求对比 / 真实缺口',20:'现状供给',21:'需求预测',22:'缺口',23:'缺口率',24:'28%',25:'03',27:'选址推演',28:'村组位置 / 服务半径 / 建设情景',29:'现有设施',30:'拟建设施',31:'服务半径',32:'服务覆盖',33:'04',34:'效果复盘',35:'投入产出 / 使用率 / 后评估',36:'投入产出比',37:'设施使用率',38:'群众满意度',39:'1:2.8',40:'88%',41:'92%',42:'风险评估',43:'先调剂、再维修、最后才新建',44:'从“凭经验争项目”，变成“拿测算做决策”'},
12:{1:'智管：让每一座烤房都有可追溯的生命周期',2:'从建档、巡检、维修到基金保障，全部沉淀为长期资产数据',3:'基础信息',4:'设备清单',5:'责任人员',7:'建档时间',8:'01 建档',9:'一房一档 / 责任到人',10:'02',11:'巡检',12:'状态更新 / 风险发现',13:'巡检记录',14:'健康评分',16:'隐患告警',17:'状态趋势',18:'基金投入',19:'使用记录',20:'优先级排序',21:'资金效益',22:'04 保障',23:'基金统筹 / 轻重缓急',25:'档案不断档，责任不断链',26:'03',27:'维修',28:'工单闭环 / 备件留痕',29:'维修工单',30:'维修历史',31:'备件更换',32:'维修成本',33:'智管不是管一次，而是让数据、责任和资金跟着烤房走完整个生命周期',34:'92'},
13:{1:'四队长效：让机制跑在平台后面',2:'把一次上线，做成长周期可持续运转',3:'04',4:'03',5:'02',6:'基金保障',7:'资金统筹 / 轻重缓急 / 长效投入',8:'01',9:'运营管护',10:'日常巡检 / 数据更新 / 使用培训',11:'维修服务',12:'工单闭环 / 备件调度 / 服务评价',13:'专班推进',14:'任务统筹 / 进度督办 / 跨部门协同',15:'机制不散、服务不断、数据不乱、保障不虚'},
14:{1:'应用成效：从经验管房，走向数据管房',2:'资源可见、维修提速、决策更准、资金更省',3:'2442座',4:'烤房入库建档',5:'1363次',6:'调度与服务闭环记录',7:'96.8%',8:'家底信息完整率',9:'≈190万',10:'投入测算参考价值',11:'48个',12:'重点烟区应用覆盖',13:'看得见、调得动、修得快、投得准，才是真正的价值'},
15:{1:'总结：一套“智烤管家”，把难题变成闭环',2:'一库清底 — 双模决策 — 三智落地 — 四队长效',3:'02',4:'双模决策',5:'03',6:'三智落地',7:'04',8:'四队长效',9:'01',10:'一库清底',11:'专班 / 服务 / 管护 / 保障',12:'智用 / 智投 / 智管',13:'一房一档\n家底清楚',14:'家底清楚',15:'都有一份可计算、',16:'评估 + 推演\n量化决策',17:'量化决策',18:'让每一座烤房，',19:'可调度、可运维、可决策的健康档案',20:'—  湖北烟草创客大赛  ·  决赛发布  —'}
}
KEEP_PICTURE={5:{15,18},6:{45,48},7:{10},10:{10,17,19},11:{17,26},12:{6,15,24}}

ROIS={
1:[(2.0,.25,9.4,4.5)],2:[(.55,1.55,1.75,4.65),(7.55,1.65,1.70,1.35),(7.55,3.10,1.70,1.35),(7.55,4.55,1.70,1.35)],
3:[(2.55,1.45,8.35,4.55)],4:[(.45,1.60,3.05,4.20),(3.45,1.55,3.00,4.25),(6.45,1.55,3.00,4.25),(9.45,1.55,3.30,4.25)],
5:[(.30,3.55,8.15,2.55),(8.10,3.20,4.80,3.05)],6:[(.25,2.15,4.25,3.75),(4.30,2.10,4.45,3.75),(8.70,2.15,4.25,3.75)],
7:[(.25,2.00,6.35,4.10),(6.10,1.95,6.75,4.20)],8:[(.30,1.85,4.95,4.05),(4.90,2.25,3.55,3.75),(8.10,1.85,4.95,4.05)],
9:[(.30,2.15,4.10,3.95),(4.40,2.15,4.10,3.95),(8.50,2.15,4.45,3.95)],10:[(.15,2.10,4.25,4.00),(3.95,1.95,5.30,4.25),(9.00,2.10,4.20,4.00)],
11:[(.15,1.95,3.25,3.95),(3.20,1.95,3.35,3.95),(6.45,1.95,3.35,3.95),(9.65,1.95,3.55,3.95)],
12:[(.20,1.65,3.25,4.55),(3.20,1.65,3.20,4.55),(6.35,1.65,3.20,4.55),(9.45,1.65,3.55,4.55)],
13:[(.45,1.80,2.85,4.20),(3.30,1.80,2.85,4.20),(6.35,1.80,2.85,4.20),(9.30,1.80,3.00,4.20)],
14:[(.45,1.85,2.60,2.15),(3.35,1.85,2.55,2.15),(6.45,1.85,2.65,2.15),(9.55,1.85,2.75,2.15),(3.00,4.25,3.15,1.65),(7.10,4.25,3.15,1.65)],
15:[(.35,1.75,3.00,3.85),(3.40,1.75,3.00,3.85),(6.45,1.75,3.00,3.85),(9.50,1.75,3.20,3.85)]}

SERIF='Noto Serif CJK SC'; SANS='Noto Sans CJK SC'; WHITE=RGBColor(242,248,255)

def alpha_from_crop(rgb,mode='art'):
    a=np.array(rgb).astype(np.int16); r,g,b=a[:,:,0],a[:,:,1],a[:,:,2]; v=np.max(a,axis=2); cyan=((g+b)//2-r)
    al=np.maximum((v-(52 if mode=='road' else 78))*(3.2 if mode=='road' else 3.0),(cyan-(12 if mode=='road' else 27))*(5.0 if mode=='road' else 4.5))
    al=np.clip(al,0,255).astype(np.uint8); al[v<58]=0
    return cv2.GaussianBlur(al,(3,3),0)

def make_clean_bg(srcprs):
    slide=srcprs.slides[0]; base=Image.open(io.BytesIO(slide.shapes[0].image.blob)).convert('RGB'); arr=np.array(base); H,W=arr.shape[:2]
    mask=np.zeros((H,W),np.uint8)
    sx=W/(srcprs.slide_width/914400); sy=H/(srcprs.slide_height/914400)
    for sh in list(slide.shapes)[1:]:
        if sh.shape_type==13:
            x=int(sh.left/914400*sx); y=int(sh.top/914400*sy); w=int(sh.width/914400*sx); h=int(sh.height/914400*sy)
            cv2.rectangle(mask,(max(0,x-18),max(0,y-12)),(min(W-1,x+w+18),min(H-1,y+h+12)),255,-1)
    y0=int(H*0.69); mask[y0:,:]=255
    bgr=cv2.cvtColor(arr,cv2.COLOR_RGB2BGR); out=cv2.inpaint(bgr,mask,9,cv2.INPAINT_TELEA); out=cv2.cvtColor(out,cv2.COLOR_BGR2RGB)
    blur=cv2.GaussianBlur(out,(0,0),sigmaX=28,sigmaY=18)
    for y in range(y0,H):
        t=(y-y0)/max(1,H-y0); a=min(0.86,0.35+0.55*t); out[y]=(out[y]*(1-a)+blur[y]*a).astype(np.uint8)
        out[y]=(out[y].astype(np.float32)*(1-0.24*t)+np.array([4,20,40])*(0.24*t)).clip(0,255).astype(np.uint8)
    return Image.fromarray(out)

def is_num(s):
    return s.strip() in {'01','02','03','04'} or bool(re.fullmatch(r'[0-9,.:≈%万座次个年]+(?:\n分)?',s.strip()))

def add_native_under(slide,pic,txt,idx,sno):
    x,y,w,h=pic.left,pic.top,pic.width,pic.height; hi=h/914400; body=hi<=0.48 and len(txt)>=8; big=hi>=0.58 or (len(txt)<=10 and hi>=0.47) or is_num(txt)
    font=SERIF if big and not body else SANS; bold=big or idx==1
    if txt.strip().startswith('Z H I'): font=SANS; bold=False
    if sno==1 and idx==2: size=58
    elif is_num(txt): size=max(16,min(46,hi*72*0.80))
    elif hi>=0.78: size=max(18,min(38,hi*72*0.52))
    elif big: size=max(14,min(31,hi*72*0.68))
    else: size=max(8,min(18,hi*72*0.55))
    tb=slide.shapes.add_textbox(x,y,w,h); tb.name=f'可编辑文字下层_S{sno:02d}_{idx:03d}_双击编辑_上方保留原字效切片'
    tf=tb.text_frame; tf.clear(); tf.word_wrap=('\n' in txt); tf.margin_left=tf.margin_right=0; tf.margin_top=tf.margin_bottom=0; tf.vertical_anchor=MSO_ANCHOR.MIDDLE; tf.auto_size=MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER; r=p.add_run(); r.text=txt; r.font.name=font; r.font.bold=bold; r.font.size=Pt(size); r.font.color.rgb=WHITE
    cNvPr=tb._element.xpath('.//p:cNvPr')[0]; cNvPr.set('hidden','1')
    sp=tb._element; pe=pic._element; parent=pe.getparent(); parent.remove(sp); parent.insert(parent.index(pe),sp)

rA=Presentation(SRC_A); rB=Presentation(SRC_B); sources=[(rA,s) for s in rA.slides]+[(rB,s) for s in rB.slides]; assert len(sources)==15
prs=Presentation(); prs.slide_width=Inches(13.3326); prs.slide_height=Inches(7.5); blank=prs.slide_layouts[6]
clean=make_clean_bg(rA).resize((1600,900),Image.Resampling.LANCZOS); bgpath=TMP/'clean_bg.jpg'; clean.save(bgpath,quality=94)
for gi,(srcprs,slide) in enumerate(sources,1):
    dst=prs.slides.add_slide(blank); bg=dst.shapes.add_picture(str(bgpath),0,0,width=prs.slide_width,height=prs.slide_height); bg.name='背景_纯山海_不含文字图标_不含底部光路'
    base=Image.open(io.BytesIO(slide.shapes[0].image.blob)).convert('RGB'); W,H=base.size; sx=W/(srcprs.slide_width/914400); sy=H/(srcprs.slide_height/914400)
    textmask=np.zeros((H,W),np.uint8); orig=[]
    for si,sh in enumerate(slide.shapes):
        if si and sh.shape_type==13:
            x=int(sh.left/914400*sx); y=int(sh.top/914400*sy); w=int(sh.width/914400*sx); h=int(sh.height/914400*sy); cv2.rectangle(textmask,(max(0,x-5),max(0,y-3)),(min(W-1,x+w+5),min(H-1,y+h+3)),255,-1); orig.append(sh)
    roadcut=int(5.45*sy)
    for ri,(x,y,w,h) in enumerate(ROIS[gi],1):
        x0=max(0,int(x*sx)); y0=max(0,int(y*sy)); x1=min(W,int((x+w)*sx)); y1=min(H,int((y+h)*sy)); crop=base.crop((x0,y0,x1,y1)); alpha=alpha_from_crop(crop); alpha[textmask[y0:y1,x0:x1]>0]=0
        if y1>roadcut: alpha[max(0,roadcut-y0):,:]=0
        if int((alpha>18).sum())<70: continue
        ap=TMP/f's{gi:02d}_visual_{ri:02d}.png'; Image.fromarray(np.dstack([np.array(crop),alpha])).save(ap); q=dst.shapes.add_picture(str(ap),Inches(x0/sx),Inches(y0/sy),width=Inches((x1-x0)/sx),height=Inches((y1-y0)/sy)); q.name=f'图标与图形组件_S{gi:02d}_{ri:02d}_独立可移动删除缩放'
    ry0=int(5.25*sy); road=base.crop((0,ry0,W,H)); ra=np.array(road); aa=alpha_from_crop(road,'road'); aa[textmask[ry0:H,:]>0]=0; hsv=cv2.cvtColor(ra,cv2.COLOR_RGB2HSV); strong=((hsv[:,:,2]>92)&(hsv[:,:,1]>40))|(hsv[:,:,2]>160); aa[~strong]=0
    if int((aa>15).sum())>100:
        rp=TMP/f's{gi:02d}_road.png'; Image.fromarray(np.dstack([ra,aa])).save(rp); q=dst.shapes.add_picture(str(rp),0,Inches(ry0/sy),width=prs.slide_width,height=Inches((H-ry0)/sy)); q.name=f'底部光路_S{gi:02d}_独立可删除_背景本身无此光路'
    for ci,sh in enumerate(orig,1):
        ip=TMP/f's{gi:02d}_existing_{ci:03d}.{sh.image.ext or "png"}'; ip.write_bytes(sh.image.blob); q=dst.shapes.add_picture(str(ip),sh.left,sh.top,width=sh.width,height=sh.height); q.rotation=sh.rotation
        if ci in KEEP_PICTURE.get(gi,set()): q.name=f'图标或纯视觉切片_S{gi:02d}_{ci:03d}_独立可移动删除缩放'
        else:
            txt=T.get(gi,{}).get(ci); q.name=f'原字效切片_S{gi:02d}_{ci:03d}_删除后可直接编辑下层文字' if txt is not None else f'原始切片_S{gi:02d}_{ci:03d}_独立可删除'
            if txt is not None: add_native_under(dst,q,txt,ci,gi)
prs.save(OUT); print(OUT)


# Also create a second deck where native text is immediately visible/editable.
DIRECT=ROOT/f'智烤管家_{STAMP}_1003文字直接可编辑版_图标分层背景无光路.pptx'
p2=Presentation(OUT)
for slide in p2.slides:
    for sh in list(slide.shapes):
        n=sh.name or ''
        if n.startswith('原字效切片_'):
            el=sh._element; el.getparent().remove(el)
        elif n.startswith('可编辑文字下层_'):
            c=sh._element.xpath('.//p:cNvPr')[0]
            c.attrib.pop('hidden',None)
            sh.name=n.replace('可编辑文字下层_','原生可编辑文字_').replace('_上方保留原字效切片','')
p2.save(DIRECT)
print(DIRECT)
