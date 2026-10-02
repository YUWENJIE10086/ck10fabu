_panel(s,5.05,1.22,4.45,3.9,glow=True,fill1='073F5E',fill2='01182E',alpha1=63000,alpha2=42000)
add_text(s,'真实后台｜一房一档 · 状态可查 · 过程可追',5.27,1.43,3.95,.3,11,WHITE,True)
image_frame(s,AS/'db_web.jpg',5.28,1.9,3.08,1.78,'WEB')
image_frame(s,AS/'db_phone.jpg',8.47,1.98,.75,1.5,'MOBILE')
add_grad_text(s,'2442',5.38,3.93,1.6,.55,28,True,PP_ALIGN.LEFT,YELLOW,ORANGE);add_text(s,'座烤房全口径入库',6.85,4.04,2.0,.34,10,SOFT,True)
add_text(s,'采集覆盖18个关键部件｜模型重点抓12个关键部件',5.32,4.58,3.85,.28,9,MUTED,False)
add_fade_transition(s)

# Slide6 Model1
s=prs.slides[5]; remove_all_shapes(s); add_hud_header(s,'模型①  烤房健康评估','三步，把“看起来还能用”变成“可计算、可解释、可排序”',section=2,tabs=[('模型1',True),('模型2',False),('耦合',False)])
model1=[('1','状态标准化','12个关键部件 + 房龄/维修/使用记录','把“好不好”变成同一量纲',AS/'m1_a.jpg'),('2','AHP组合评分','专家两两比较 + 一致性校验','部件80% · 管护10% · 人评10%',AS/'m1_b.jpg'),('3','Weibull估寿','寿命曲线 + 维修复衰 + 健康校准','输出剩余年限与高风险部位',AS/'m1_c.jpg')]
for i,(n,t,tech,out,img) in enumerate(model1):
    x=.42+i*3.18
    add_panel(s,x,1.25,2.9,3.35,glow=True,fill1='073F5C',fill2='01172E',alpha1=65000,alpha2=43000)
    add_grad_text(s,n,x+.12,1.38,.42,.44,22,True,PP_ALIGN.CENTER,YELLOW,ORANGE); add_grad_text(s,t,x+.62,1.4,2.05,.38,16,True,PP_ALIGN.LEFT,WHITE,CYAN2)
    image_frame(s,img,x+.2,1.92,2.5,.8,None)
    add_text(s,tech,x+.22,2.9,2.45,.5,10,SOFT,False,PP_ALIGN.CENTER)
    add_text(s,out,x+.22,3.55,2.45,.5,10,YELLOW,True,PP_ALIGN.CENTER)
    if i<2: add_arrow(s,x+2.87,2.45,x+3.18,2.72,CYAN)
# output ribbon
add_panel(s,1.35,4.78,7.3,.45,glow=True,fill1='0B5263',fill2='08233B',alpha1=72000,alpha2=52000)
add_text(s,'输出：健康等级  ·  剩余寿命  ·  风险部位  ·  维修优先序',1.45,4.82,7.1,.34,12,WHITE,True,PP_ALIGN.CENTER)
add_text(s,'技术亮点：不是只“打一个分”，而是把评分、寿命和维修优先序连成一张可解释的“处方”',1.2,5.27,7.6,.22,9,MUTED,False,PP_ALIGN.CENTER)
add_fade_transition(s)

# Slide7 Model2
s=prs.slides[6]; remove_all_shapes(s); add_hud_header(s,'模型②  千亩村发展导引','三步，把“想建几栋”变成“真实缺口 + 最优位置 + 建后效果”',section=2,tabs=[('模型1',False),('模型2',True),('耦合',False)])
model2=[('01','算有效烤能','健康结果 + 容量 + 使用状态','从“有多少栋”到“真正能烤多少”'),('02','判真实缺口','AHP主观权重 + 熵权客观校正','叠加种烟规模、承载和发展条件'),('03','做建设推演','TOPSIS候选点排序 + 情景沙盘','先模拟，再决定建不建、建哪里')]
add_panel(s,.45,1.28,4.1,3.8,glow=True,fill1='07425D',fill2='01182E',alpha1=65000,alpha2=42000)
for i,(n,t,tech,out) in enumerate(model2):
    y=1.58+i*1.08
    add_number_badge(s,n,t,.7,y)
    add_text(s,tech,2.05,y-.03,2.25,.25,9,SOFT,False)
    add_text(s,out,2.05,y+.28,2.25,.34,9,YELLOW,True)
add_panel(s,4.85,1.28,4.65,3.8,glow=True,fill1='053E5B',fill2='01182E',alpha1=63000,alpha2=42000)
add_text(s,'情景沙盘：把未来先“演一遍”',5.08,1.48,4.18,.32,12,WHITE,True)
image_frame(s,AS/'m2_bi.jpg',5.12,1.93,3.95,2.25,'BI / SANDBOX')
add_grad_text(s,'当前  /  +2%  /  +5%  /  +10%',5.13,4.35,3.95,.34,15,True,PP_ALIGN.CENTER,YELLOW,ORANGE)
add_text(s,'实时重算烤能 · 负载 · 缺口 · 拟建数量',5.25,4.68,3.7,.26,9,SOFT,False,PP_ALIGN.CENTER)
add_fade_transition(s)

# Slide8 coupling
s=prs.slides[7]; remove_all_shapes(s); add_hud_header(s,'双模耦合：健康结果进入建设决策','真正的创新，不是两个模型并排，而是让模型①成为模型②的输入',tabs=[('模型1',False),('模型2',False),('耦合',True)])
# beam
beam=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(1.1),Inches(2.55),Inches(7.8),Inches(.06));beam.fill.solid();beam.fill.fore_color.rgb=rgb(CYAN);set_alpha_on_solid(beam,45000);beam.line.fill.background();add_glow(beam,CYAN,9,60)
nodes=[('1','存量体检','健康 · 寿命 · 风险'),('2','有效烤能','健康 × 容量 × 状态'),('3','增量推演','规模 · 负载 · 缺口'),('4','投入决策','修复 · 新建 · 选址')]
for i,(n,t,d) in enumerate(nodes):
    x=1.1+i*2.05
    halo=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+.38),Inches(1.72),Inches(.8),Inches(.8));halo.fill.solid();halo.fill.fore_color.rgb=rgb('074968');set_alpha_on_solid(halo,72000);set_gradient_line(halo,CYAN2,'1378FF',1.5);add_glow(halo,CYAN,10,65)
    add_grad_text(s,n,x+.52,1.86,.52,.42,20,True,PP_ALIGN.CENTER,YELLOW,ORANGE)
    add_grad_text(s,t,x,2.78,1.55,.38,15,True,PP_ALIGN.CENTER,WHITE,CYAN2)
    add_text(s,d,x-.05,3.2,1.65,.3,9,SOFT,False,PP_ALIGN.CENTER)
    if i<3: add_arrow(s,x+1.58,2.42,x+1.93,2.72,CYAN)
add_panel(s,1.25,4.32,7.5,.62,glow=True,fill1='104F5B',fill2='061E34',alpha1=72000,alpha2=52000,line1=YELLOW,line2=CYAN2)
add_text(s,'先调剂   →   再修复   →   仍有结构性缺口才新建',1.35,4.4,7.3,.42,15,YELLOW,True,PP_ALIGN.CENTER)
add_text(s,'决策顺序被固化进系统：先盘活存量，再讨论增量',2.1,5.08,5.8,.24,9,MUTED,False,PP_ALIGN.CENTER)
add_fade_transition(s)

# Slide9 video
s=prs.slides[8]; remove_all_shapes(s); add_hud_header(s,'从模型判断，到烟区一线','约1分钟视频｜看一张工单、一次沙盘，如何真正跑起来')
for rad,alpha in [(2.2,18000),(1.65,25000),(1.05,36000)]:
    ring=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(5-rad/2),Inches(2.75-rad/2),Inches(rad),Inches(rad));ring.fill.background(); set_gradient_line(ring,CYAN2,'126DFF',1.0); add_glow(ring,CYAN,8,35)
tri=s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,Inches(4.72),Inches(2.25),Inches(.72),Inches(.82));tri.rotation=90;tri.fill.solid();tri.fill.fore_color.rgb=rgb(CYAN2);tri.line.fill.background();add_glow(tri,CYAN,10,70)
add_grad_text(s,'VIDEO',4.1,3.47,1.8,.36,15,True,PP_ALIGN.CENTER,YELLOW,ORANGE)
for i,t in enumerate(['模型判断','工单流转','沙盘推演']): add_panel(s,2.45+i*1.77,4.17,1.5,.4,glow=False,fill1='07435A',fill2='021A30',alpha1=70000,alpha2=50000); add_text(s,t,2.45+i*1.77,4.2,1.5,.32,10,SOFT,True,PP_ALIGN.CENTER)
add_fade_transition(s)

# Slide10 3智 overview
s=prs.slides[9]; remove_all_shapes(s); add_hud_header(s,'三智落地：三类角色，一套闭环','管理人员投得准｜烟农技术员用得顺、修得快｜合作社管得久',section=3,tabs=[('智投',True),('智用',False),('智管',False)])
three=[('智投','管理人员','看健康 → 看缺口 → 看沙盘','修复 / 新建 / 资金有依据',YELLOW),('智用','烟农 + 技术员','查可用 → 预约调剂 → 一键报修','找得到房 · 找得到人 · 找得到配件',CYAN2),('智管','合作社','申请 → 派单 → 验收 → 评价回写','有人管 · 有钱管 · 有标准管',GREEN)]
for i,(t,role,flow,val,col) in enumerate(three):
    x=.5+i*3.16
    add_panel(s,x,1.38,2.82,3.35,glow=True,fill1='073E59',fill2='01182F',alpha1=65000,alpha2=43000,line1=col,line2=CYAN2)
    add_grad_text(s,t,x+.18,1.65,1.28,.48,23,True,PP_ALIGN.LEFT,col,WHITE)
    add_text(s,role,x+1.48,1.72,1.1,.3,10,SOFT,True,PP_ALIGN.RIGHT)
    # orbit badge
    c=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+1.04),Inches(2.35),Inches(.72),Inches(.72)); c.fill.solid();c.fill.fore_color.rgb=rgb('06425B');set_alpha_on_solid(c,72000);set_gradient_line(c,col,CYAN2,1.3);add_glow(c,col,8,58)
    add_text(s,str(i+1),x+1.04,2.35,.72,.72,18,col,True,PP_ALIGN.CENTER)
    add_text(s,flow,x+.3,3.25,2.22,.5,10,WHITE,True,PP_ALIGN.CENTER)
    add_text(s,val,x+.24,4.02,2.34,.4,9,col,True,PP_ALIGN.CENTER)
add_text(s,'所有动作都回写到同一份烤房数字档案，形成“决策—服务—管护—再学习”的闭环',1.15,4.98,7.7,.28,10,YELLOW,True,PP_ALIGN.CENTER)
add_fade_transition(s)

# Sli