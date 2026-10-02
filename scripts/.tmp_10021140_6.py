ve a shape safely
def remove_shape(shape):
    el=shape._element; el.getparent().remove(el)

# ---------- content refinements ----------
def find_shape(slide, needle):
    for sh in slide.shapes:
        if hasattr(sh,'text') and needle in sh.text:
            return sh
    return None

def set_text(shape, text):
    if shape is None: return
    shape.text=text

# Slide 2 sharper business problem
s=prs.slides[1]
set_text(find_shape(s,'问题不是'), '不是“烤房不够”，而是可用状态看不见、故障协同跑不快、投入决策算不准')
set_text(find_shape(s,'旺季状态不透明'), '旺季资源状态不透明\n烟农找不到，技术员也难统筹')
set_text(find_shape(s,'故障突发后'), '故障突发后\n报修、诊断、人员、配件串不起来')
set_text(find_shape(s,'修哪栋、建哪里'), '维修、新建仍偏经验\n“该修谁、该建哪、该投多少”算不清')
set_text(find_shape(s,'表面是'), '表面是“用房、抢修、投入”三件事，本质是一条“用 · 投 · 管”链路没有闭环')

# Slide3 evidence -> root causes
s=prs.slides[2]
set_text(find_shape(s,'匿名问卷'), '2025匿名问卷 + 线下走访｜从三类人的“愁”，追到管理链路的根因')
set_text(find_shape(s,'找不到房'), '用房调剂难 / 故障抢修慢 / 服务体验不稳定')
set_text(find_shape(s,'状态不明'), '底数状态不清 / 修建投入凭经验 / 需求难匹配')
set_text(find_shape(s,'不敢管'), '资金不足 / 责任悬空 / 标准不一 / 队伍履职弱')
set_text(find_shape(s,'这不是三个孤立问题'), '“三愁”不是三个孤立问题，而是资源配置 → 投入决策 → 日常管护之间缺少有效衔接')

# Slide4 1234 wording
s=prs.slides[3]
set_text(find_shape(s,'一个协同平台，把数据'), '一个协同平台，把“数据底座—模型决策—业务服务—长效机制”聚能成环')

# Slide5 data layer value
s=prs.slides[4]
set_text(find_shape(s,'低代码采数｜'), '低代码采数｜统一烤房编号｜业务过程持续回写')
set_text(find_shape(s,'采集覆盖18个关键部件'), '数据采集覆盖18个关键部件｜健康模型重点计算12个关键部件')
set_text(find_shape(s,'底数不再'), '数据库不是一次性建档：每次调剂、报修、验收、评价都会反向更新“一房一档”')

# Slide6 model1 technique + business outputs
s=prs.slides[5]
set_text(find_shape(s,'回答三件事'), '三步把“人工看状态”升级为“健康—寿命—风险—维修”的动态体检')
# refine bottom output line
for sh in s.shapes:
    if hasattr(sh,'text') and '输出' in sh.text and '健康' in sh.text:
        sh.text='输出：健康等级 → 剩余寿命 → 风险部位 → 维修优先序'
# add small risk label
rb=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,emu_x(.792),emu_y(.718),emu_x(.145),emu_y(.055))
set_shape_glass(rb, fill='43141D', fill_alpha=74, line_colors=('FF6B61','FFB049'), glow='FF3A4A', glow_alpha=38, glow_pt=4, line_pt=1)
tf=rb.text_frame; tf.clear(); tf.vertical_anchor=MSO_ANCHOR.MIDDLE; p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER; r=p.add_run(); r.text='<60分：预警/危险'; r.font.name='Microsoft YaHei'; r.font.size=Pt(10.5); r.font.bold=True; r.font.color.rgb=RGBColor(255,204,104)

# Slide7 completely rebuild core content area to clear model2 logic
s=prs.slides[6]
# keep title/header/nav; remove content shapes below y=.17
for sh in list(s.shapes):
    if sh.top > emu_y(.17):
        remove_shape(sh)
# three stage cards, with cropped dashboard evidence
xpos=[.045,.352,.659]
steps=[
    ('01','算“有效烤能”','健康分 × 容量 × 使用状态','先把“有几栋”变成“真正能烤多少”'),
    ('02','判“真实缺口”','AHP主观权重 + 熵权客观校正','把规模趋势、负载压力、管护条件量化'),
    ('03','做“建设推演”','TOPSIS候选点排序 + 情景沙盘','回答建不建、建多少、建哪里、建后改善多少'),
]
# crop m2_bi into three image regions
img=str(ASSET/'m2_bi.jpg')
for i,(x,st) in enumerate(zip(xpos,steps)):
    num,title,tech,meaning=st
    card=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,emu_x(x),emu_y(.218),emu_x(.285),emu_y(.635)); set_shape_glass(card, fill='062B4A', fill_alpha=62, glow_alpha=42, glow_pt=6)
    # number chip
    chip=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,emu_x(x+.018),emu_y(.243),emu_x(.05),emu_y(.055)); set_shape_glass(chip,fill='0D4964',fill_alpha=78,line_colors=('E7FBFF','29CFFF'),glow_alpha=45,glow_pt=4)
    tf=chip.text_frame; tf.clear(); tf.vertical_anchor=MSO_ANCHOR.MIDDLE; p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER; r=p.add_run(); r.text=num; r.font.bold=True; r.font.size=Pt(10); r.font.color.rgb=RGBColor(255,208,55)
    # title
    tb=s.shapes.add_textbox(emu_x(x+.08),emu_y(.238),emu_x(.18),emu_y(.06)); tf=tb.text_frame; tf.clear(); p=tf.paragraphs[0]; r=p.add_run(); r.text=title; r.font.name='Microsoft YaHei'; r.font.bold=True; r.font.size=Pt(15); set_run_gradient(r,('FFFFFF','8DEBFF','14B5FF'))
    # picture
    pic=s.shapes.add_picture(img, emu_x(x+.02), emu_y(.335), width=emu_x(.245), height=emu_y(.205))
    # crop different areas horizontally
    if i==0: pic.crop_left=0.00; pic.crop_right=0.50; pic.crop_top=0.00; pic.crop_bottom=0.15
    if i==1: pic.crop_left=0.25; pic.crop_right=0.25; pic.crop_top=0.00; pic.crop_bottom=0.15
    if i==2: pic.crop_left=0.50; pic.crop_right=0.00; pic.crop_top=0.00; pic.crop_bottom=0.15
    add_corner_brackets(s,x+.014,.327,.257,.221)
    # tech line
    tb=s.shapes.add_textbox(emu_x(x+.02),emu_y(.566),emu_x(.245),emu_y(.072)); tf=tb.text_frame; tf.clear(); p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER; r=p.add_run(); r.text=tech; r.font.name='Microsoft YaHei'; r.font.bold=True; r.font.size=Pt(10.5); r.font.color.rgb=RGBColor(255,204,72)
    # meaning
    tb=s.shapes.add_textbox(emu_x(x+.025),emu_y(.66),emu_x(.235),emu_y(.12)); tf=tb.text_frame; tf.clear(); p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER; r=p.add_run(); r.text=meaning; r.font.name='Microsoft YaHei'; r.font.size=Pt(9.8); r.font.color.rgb=RGBColor(194,229,242)
# bottom value strip
val=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,emu_x(.12),emu_y(.875),emu_x(.76),emu_y(.07)); set_shape_glass(val,fill='083A50',fill_alpha=75,line_colors=('FFD64A','39DDFF'),glow='28D9FF',glow_alpha=35,glow_pt=4,line_pt=1.1)
tf=val.text_frame; tf.clear(); tf.vertical_anchor=MSO_ANCHOR.MIDDLE; p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER; r=p.add_run(); r.text='模型2的高明之处：先继承模型1的健康结果，再把“存量真实能力”送进新建决策'; r.font.name='Microsoft YaHei'; r.font.bold=True; r.font.size=Pt(11); r.font.color.rgb=RGBColor(255,213,71)

# Slide8 coupling more explicit
s=prs.slides[7]
set_text(find_shape(s,'这条链把'), '双模型不是两个独立打分器，而是一条连续决策链：模型1先把存量算真，模型2再把增量算准')

# Slide10 overview: audience mapping
s=prs.slides[9]
set_text(find_shape(s,'三类角色'), '三智不是三个菜单，而是把双模型直接送到三类人的工作现场')
# could refine cards descriptions if found

# Slide11 智投
s=prs.slides[10]
set_text(find_shape(s,'管理人员'), '管理人员：模型给出“修复—新建—资金”的证据链')
set_text(find_shape(s,'方案论证'), '方案论证：县级29栋｜肖家坪7栋（用于建设方案论证）')

# Slide12 智用
s=prs.slides[11]
set_text(find_shape(s,'烟农 + 技术员'), '烟农 + 技术员：从“到处问”变成“线上匹配、工单协同、结果回写”')
set_text(find_shape(s,'实战：雷击停运'), '实战：雷击停运 → 定位2处故障 → 匹配人员/配件 → 3小时内复产｜48次调剂｜18次抢修均<3h｜平台工单平均响应0.38h')

# Slide13 智管
s=prs.slides[12]
set_text(find_shape(s,'合作社：'), '合作社：线上把流程闭环，线下用“四队+基金+制度”把责任真正落下去')

# Slide14 results transform
s=prs.slides[13]
set_text(find_shape(s,'一个烘烤季'), '不只看“做了多少”，更看三类人的工作方式有没有真正改变')
set_text(find_shape(s,'找房靠问'), '找房靠问 → 在线匹配\n故障等人 → 工单协同')
set_text(find_shape(s,'凭经验排查'), '凭经验排查 → 风险排序\n凭感觉申报 → 沙盘推演')
set_text(find_shape(s,'责任悬空'), '责任悬空 → 四队履职\n一次投入 → 基金滚动