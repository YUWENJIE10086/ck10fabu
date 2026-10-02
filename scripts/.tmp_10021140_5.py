=720, n=24):
    frames=[]
    margin=34; top=92; bottom=692; r=28
    # approximate rounded-rect perimeter points
    left=margin; right=W-margin; y0=top; y1=bottom
    pts=[]
    # top
    for x in range(left+r,right-r,8): pts.append((x,y0))
    # tr arc
    for a in range(-90,1,5):
        rad=math.radians(a); pts.append((right-r+r*math.cos(rad), y0+r+r*math.sin(rad)))
    for y in range(y0+r,y1-r,8): pts.append((right,y))
    for a in range(0,91,5):
        rad=math.radians(a); pts.append((right-r+r*math.cos(rad), y1-r+r*math.sin(rad)))
    for x in range(right-r,left+r,-8): pts.append((x,y1))
    for a in range(90,181,5):
        rad=math.radians(a); pts.append((left+r+r*math.cos(rad), y1-r+r*math.sin(rad)))
    for y in range(y1-r,y0+r,-8): pts.append((left,y))
    for a in range(180,271,5):
        rad=math.radians(a); pts.append((left+r+r*math.cos(rad), y0+r+r*math.sin(rad)))
    L=len(pts)
    for f in range(n):
        im=Image.new('RGBA',(W,H),(0,0,0,0))
        # faint static frame layer
        d=ImageDraw.Draw(im)
        d.rounded_rectangle([left,y0,right,y1], radius=r, outline=(0,190,255,70), width=2)
        # moving highlight, on separate glow layer
        glow=Image.new('RGBA',(W,H),(0,0,0,0)); gd=ImageDraw.Draw(glow)
        start=int(f*L/n)
        seg=70
        for j in range(seg):
            idx=(start+j)%L
            x,y=pts[idx]
            dist=abs(j-seg/2)/(seg/2)
            alpha=int(230*(1-dist**1.5))
            rad=int(4+7*(1-dist))
            gd.ellipse([x-rad,y-rad,x+rad,y+rad], fill=(90,235,255,max(0,alpha)))
        glow=glow.filter(ImageFilter.GaussianBlur(5))
        im=Image.alpha_composite(im,glow)
        d=ImageDraw.Draw(im)
        for j in range(seg):
            idx=(start+j)%L
            x,y=pts[idx]
            dist=abs(j-seg/2)/(seg/2)
            alpha=int(255*(1-dist))
            d.ellipse([x-1,y-1,x+1,y+1], fill=(220,255,255,max(0,alpha)))
        frames.append(im.convert('P', palette=Image.Palette.ADAPTIVE, colors=255))
    frames[0].save(path, save_all=True, append_images=frames[1:], duration=70, loop=0, transparency=0, disposal=2, optimize=True)

make_sweep_gif(GIF)

prs=Presentation(str(SRC))
SW,SH=prs.slide_width, prs.slide_height

def emu_x(fr): return int(SW*fr)
def emu_y(fr): return int(SH*fr)

# ---------- xml effects ----------
def _remove_children(parent, tags):
    for c in list(parent):
        if c.tag.split('}')[-1] in tags:
            parent.remove(c)

def set_run_gradient(run, colors=('FFFFFF','8DEBFF','1AB8FF'), angle=5400000):
    rPr=run._r.get_or_add_rPr()
    _remove_children(rPr, {'solidFill','gradFill','noFill'})
    grad=OxmlElement('a:gradFill')
    grad.set('rotWithShape','1')
    gsl=OxmlElement('a:gsLst')
    poss=[0,50000,100000]
    for pos,col in zip(poss,colors):
        gs=OxmlElement('a:gs'); gs.set('pos',str(pos))
        clr=OxmlElement('a:srgbClr'); clr.set('val',col)
        gs.append(clr); gsl.append(gs)
    grad.append(gsl)
    lin=OxmlElement('a:lin'); lin.set('ang',str(angle)); lin.set('scaled','1')
    grad.append(lin)
    rPr.append(grad)

def set_text_gradient(shape, colors=('FFFFFF','8DEBFF','1AB8FF')):
    if not hasattr(shape,'text_frame'): return
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            if run.text.strip():
                run.font.color.rgb=RGBColor(120,230,255)  # fallback
                set_run_gradient(run, colors)

def set_shape_glass(shape, fill='082B4B', fill_alpha=62, line_colors=('5EEBFF','0AA6FF'), glow='31DDF5', glow_alpha=58, glow_pt=7.5, line_pt=1.6):
    # solid translucent fill
    spPr=shape._element.spPr
    _remove_children(spPr, {'solidFill','gradFill','noFill','ln','effectLst'})
    sf=OxmlElement('a:solidFill'); sc=OxmlElement('a:srgbClr'); sc.set('val',fill)
    al=OxmlElement('a:alpha'); al.set('val',str(int(fill_alpha*1000)))
    sc.append(al); sf.append(sc); spPr.append(sf)
    # gradient border
    ln=OxmlElement('a:ln'); ln.set('w',str(int(line_pt*12700)))
    gf=OxmlElement('a:gradFill'); gsl=OxmlElement('a:gsLst')
    for pos,col,alpha in [(0,line_colors[0],90),(50000,'FFFFFF',65),(100000,line_colors[1],90)]:
        gs=OxmlElement('a:gs'); gs.set('pos',str(pos)); c=OxmlElement('a:srgbClr'); c.set('val',col)
        a=OxmlElement('a:alpha'); a.set('val',str(alpha*1000)); c.append(a); gs.append(c); gsl.append(gs)
    gf.append(gsl); lin=OxmlElement('a:lin'); lin.set('ang','0'); lin.set('scaled','1'); gf.append(lin); ln.append(gf); spPr.append(ln)
    # glow
    eff=OxmlElement('a:effectLst'); gl=OxmlElement('a:glow'); gl.set('rad',str(int(glow_pt*12700)))
    c=OxmlElement('a:srgbClr'); c.set('val',glow); a=OxmlElement('a:alpha'); a.set('val',str(glow_alpha*1000)); c.append(a); gl.append(c); eff.append(gl); spPr.append(eff)

def add_glow_line(slide, x1,y1,x2,y2, color='20D8FF', width=1.8, alpha=65):
    line=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,emu_x(x1),emu_y(y1),emu_x(x2),emu_y(y2))
    line.line.color.rgb=RGBColor.from_string(color); line.line.width=Pt(width)
    ln=line._element.spPr.find(qn('a:ln'))
    if ln is not None:
        sf=ln.find(qn('a:solidFill'))
        if sf is not None:
            c=sf.find(qn('a:srgbClr'))
            if c is not None:
                a=OxmlElement('a:alpha'); a.set('val',str(alpha*1000)); c.append(a)
    return line

def add_corner_brackets(slide,x,y,w,h,color='82F4FF'):
    L=min(w,h)*0.12
    # 8 line segments
    segs=[(x,y,x+L,y),(x,y,x,y+L),(x+w-L,y,x+w,y),(x+w,y,x+w,y+L),
          (x,y+h-L,x,y+h),(x,y+h,x+L,y+h),(x+w,y+h-L,x+w,y+h),(x+w-L,y+h,x+w,y+h)]
    for a,b,c,d in segs: add_glow_line(slide,a,b,c,d,color,1.5,85)

def add_chip(slide, text, x,y,w, active=False):
    sh=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,emu_x(x),emu_y(y),emu_x(w),emu_y(0.045))
    set_shape_glass(sh, fill='0A3257', fill_alpha=70, line_colors=('35E2FF','2A8CFF'), glow_alpha=35, glow_pt=3.5, line_pt=0.9)
    tf=sh.text_frame; tf.clear(); tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=text; r.font.name='Microsoft YaHei'; r.font.size=Pt(9.5); r.font.bold=active; r.font.color.rgb=RGBColor(255,202,40) if active else RGBColor(160,226,246)
    return sh

def add_neon_title_bar(slide):
    # top luminous line & small end ticks, inspired by the master and creator decks
    add_glow_line(slide,0.10,0.114,0.90,0.114,'19D9FF',2.0,70)
    add_glow_line(slide,0.10,0.118,0.90,0.118,'0A74D8',0.7,55)

# remo