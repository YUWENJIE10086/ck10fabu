=50000)
            add_text(slide,label,x,.07,.62,.22,8,YELLOW if active else MUTED,True,PP_ALIGN.CENTER)
            x+=.68

def add_corner_sparks(slide,x,y,w,h):
    # four short cyan/yellow accent strokes
    for xx,yy,dx,dy,col in [
        (x-.02,y+.08,.24,.0,CYAN2),(x+.08,y-.02,0,.22,CYAN2),
        (x+w-.22,y+h+.01,.2,0,YELLOW),(x+w+.01,y+h-.2,0,.18,YELLOW)]:
        if dx:
            s=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(xx),Inches(yy),Inches(dx),Inches(.012))
        else:
            s=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(xx),Inches(yy),Inches(.012),Inches(dy))
        s.fill.solid();s.fill.fore_color.rgb=rgb(col);set_alpha_on_solid(s,65000);s.line.fill.background();add_glow(s,col,3,45)

def add_number_badge(slide,num,label,x,y,color=YELLOW):
    c=slide.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(.42),Inches(.42))
    c.fill.solid();c.fill.fore_color.rgb=rgb('043C59');set_alpha_on_solid(c,70000);set_gradient_line(c,CYAN2,'1C8BFF',1.2);add_glow(c,CYAN,4,45)
    add_text(slide,str(num),x,y,.42,.42,12,color,True,PP_ALIGN.CENTER)
    add_text(slide,label,x+.53,y-.01,1.6,.44,13,WHITE,True)

def add_arrow(slide,x1,y1,x2,y2,color=CYAN):
    sh=slide.shapes.add_shape(MSO_SHAPE.CHEVRON,Inches(x1),Inches(y1),Inches(x2-x1),Inches(y2-y1))
    sh.fill.solid();sh.fill.fore_color.rgb=rgb(color);set_alpha_on_solid(sh,50000); sh.line.fill.background(); add_glow(sh,color,3,45)
    return sh

def picture(slide,path,x,y,w,h):
    return slide.shapes.add_picture(str(path),Inches(x),Inches(y),Inches(w),Inches(h))

def image_frame(slide,path,x,y,w,h,label=None):
    add_panel(slide,x-.04,y-.04,w+.08,h+.08,glow=True,fill1='063553',fill2='02182B',alpha1=65000,alpha2=38000)
    pic=picture(slide,path,x,y,w,h)
    add_corner_sparks(slide,x-.03,y-.03,w+.06,h+.06)
    if label:
        lab=add_panel(slide,x+.12,y+.08,1.15,.26,glow=False,fill1='10657E',fill2='06324A',alpha1=85000,alpha2=70000)
        add_text(slide,label,x+.12,y+.08,1.15,.26,8,YELLOW,True,PP_ALIGN.CENTER)
    return pic

# Cover: preserve but enhance title gradient and glows
s=prs.slides[0]
for sh in s.shapes:
    if hasattr(sh,'text') and sh.text.strip() in ['智烤','管家']:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                set_text_gradient(r,WHITE,CYAN2); r.font.bold=True
        add_glow(sh,CYAN,9,60)
add_fade_transition(s)

# SLIDE 2 Problem
s=prs.slides[1]; remove_all_shapes(s); add_hud_header(s,'烟区一线，最怕这3件事','不是“有没有烤房”，而是资源、抢修、投入没有被数字化组织起来')
probs=[('第一“难”','分不到','旺季状态不透明\n调剂仍靠电话、口头协调','01'),('第二“难”','修不及','故障突发后\n找人、诊断、配件仍靠人工','02'),('第三“难”','投不准','修哪栋、建哪里\n缺少量化依据和供需判断','03')]
for i,(tag,big,desc,n) in enumerate(probs):
    x=.55+i*3.13
    add_panel(s,x,1.35,2.78,3.05,glow=True,fill1='094D70',fill2='02172F',alpha1=70000,alpha2=43000)
    add_corner_sparks(s,x,1.35,2.78,3.05)
    add_grad_text(s,n,x+.18,1.55,.52,.5,24,True,PP_ALIGN.CENTER,YELLOW,ORANGE)
    add_text(s,tag,x+.77,1.61,1.25,.3,10,YELLOW,True)
    add_grad_text(s,big,x+.18,2.1,2.35,.72,29,True,PP_ALIGN.CENTER,WHITE,CYAN2)
    add_text(s,desc,x+.28,3.08,2.18,.72,12,SOFT,False,PP_ALIGN.CENTER)
# root cause band
p=add_panel(s,1.05,4.68,7.9,.48,glow=True,fill1='103D54',fill2='031D35',alpha1=70000,alpha2=48000)
add_text(s,'表面是“用房 · 抢修 · 投入”三件事，本质是  用  ·  投  ·  管  没有形成闭环',1.15,4.72,7.7,.38,13,WHITE,True,PP_ALIGN.CENTER)
add_fade_transition(s)

# Slide3 Survey root causes
s=prs.slides[2]; remove_all_shapes(s); add_hud_header(s,'531份问卷，把“三难”追到根上','匿名问卷 + 线下走访｜三类关联者 · 三类忧愁')
# central orb
orb=slide_orb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(4.15),Inches(1.15),Inches(1.7),Inches(1.05)); set_gradient_fill(orb,'077B99','03294C',85000,50000);set_gradient_line(orb,CYAN2,'136DFF',1.5);add_glow(orb,CYAN,12,65)
add_grad_text(s,'531',4.28,1.25,1.44,.5,28,True,PP_ALIGN.CENTER,YELLOW,ORANGE);add_text(s,'份有效问卷',4.28,1.76,1.44,.23,9,SOFT,True,PP_ALIGN.CENTER)
roles=[('烟农愁','调剂难 · 抢修难','找不到房 / 找不到人 / 等不起'),('烟站愁','底数不清 · 投入无据','状态不明 / 修建凭经验 / 需求难匹配'),('合作社愁','资金难 · 责任难 · 标准难','不敢管 / 不愿管 / 不服管')]
for i,(a,b,c) in enumerate(roles):
    x=.48+i*3.18
    add_panel(s,x,2.5,2.85,1.55,glow=True,fill1='063E5B',fill2='02182F',alpha1=67000,alpha2=43000)
    add_text(s,a,x+.18,2.68,1.0,.3,12,YELLOW,True); add_grad_text(s,b,x+.18,3.02,2.45,.4,17,True,PP_ALIGN.LEFT,WHITE,CYAN2); add_text(s,c,x+.18,3.48,2.48,.3,9,MUTED,False)
# bottom cause chain
causes=['资源无统筹','投入无量化','管护无闭环']
for i,c in enumerate(causes):
    x=1.15+i*2.62
    add_panel(s,x,4.42,2.1,.5,glow=False,fill1='0B4057',fill2='041E35',alpha1=70000,alpha2=52000)
    add_text(s,c,x,4.46,2.1,.38,12,WHITE,True,PP_ALIGN.CENTER)
    if i<2: add_text(s,'×',x+2.17,4.46,.35,.35,16,YELLOW,True,PP_ALIGN.CENTER)
add_fade_transition(s)

# Slide4 1234
s=prs.slides[3]; remove_all_shapes(s); add_hud_header(s,'智烤管家的“1234”解法','一个协同平台，把数据、模型、场景、机制串成闭环')
items=[('1','一库清底','低代码采数\n一房一档','DATA'),('2','双模决策','存量体检\n增量推演','MODEL'),('3','三智落地','智投 · 智用\n智管','SERVICE'),('4','四队长效','调制 · 抢修\n管护 · 综合','MECHANISM')]
for i,(n,t,desc,en) in enumerate(items):
    x=.5+i*2.38
    add_panel(s,x,1.45,2.15,3.1,glow=True,fill1='084865',fill2='02172E',alpha1=67000,alpha2=43000)
    add_corner_sparks(s,x,1.45,2.15,3.1)
    add_grad_text(s,n,x+.14,1.62,.48,.48,26,True,PP_ALIGN.CENTER,YELLOW,ORANGE)
    add_grad_text(s,t,x+.22,2.2,1.72,.4,17,True,PP_ALIGN.CENTER,WHITE,CYAN2)
    add_text(s,en,x+.22,2.69,1.72,.22,8,MUTED,True,PP_ALIGN.CENTER)
    add_text(s,desc,x+.3,3.25,1.55,.62,11,SOFT,False,PP_ALIGN.CENTER)
    # tiny glowing dot
    d=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+.98),Inches(4.17),Inches(.13),Inches(.13)); d.fill.solid();d.fill.fore_color.rgb=rgb(CYAN2);d.line.fill.background();add_glow(d,CYAN,5,65)
add_text(s,'贯穿主线：先看清家底 → 再把决策算清 → 再让场景跑通 → 最后形成长效机制',1.1,4.83,7.8,.3,11,YELLOW,True,PP_ALIGN.CENTER)
add_fade_transition(s)

# Slide5 Database
s=prs.slides[4]; remove_all_shapes(s); add_hud_header(s,'一库清底：低代码把家底采上来','不是做一张静态表，而是让数据“采得快、归得准、回得来”',section=1,tabs=[('一库',True),('双模',False),('三智',False),('四队',False)])
# left flow
add_panel(s,.45,1.22,4.35,3.9,glow=True,fill1='074A63',fill2='02182F',alpha1=65000,alpha2=43000)
add_grad_text(s,'低代码采数',.72,1.42,2.1,.38,19,True,PP_ALIGN.LEFT,WHITE,CYAN2)
steps=[('01','现场采集','手机端录入 · 拍照上传\n建设/部件/使用信息'),('02','统一归档','统一烤房编号 · 一房一档\n维修/评价/管护归一'),('03','持续回写','报修/验收/评价自动回写\n数据库越用越完整')]
for i,(n,t,d) in enumerate(steps):
    y=2.05+i*.92
    add_number_badge(s,n,t,.73,y)
    add_text(s,d,2.05,y-.02,2.45,.5,10,SOFT,False)
# right evidence panel
add