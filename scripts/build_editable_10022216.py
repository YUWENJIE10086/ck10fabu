from pathlib import Path
import re

BASE = (Path('/mnt/data/build_editable_10022128.py') if Path('/mnt/data/build_editable_10022128.py').exists() else Path('scripts/build_editable_10022128.py'))
src = BASE.read_text(encoding='utf-8')

src = src.replace("STAMP='10022128'", "STAMP='10022216'")
src = src.replace("import math, zipfile, shutil", "import math, zipfile, shutil\nfrom PIL import Image, ImageEnhance, ImageDraw")
src = src.replace("W=13.333; H=7.5", "W=13.333; H=7.5\nBG_FULL='/tmp/zhikao_bg_full.jpg'; BG_REDUCED='/tmp/zhikao_bg_reduced.jpg'; SOURCE_PPT=(Path('/mnt/data/智烤管家_10021912_蓝白山海光路无框视觉决赛版.pptx') if Path('/mnt/data/智烤管家_10021912_蓝白山海光路无框视觉决赛版.pptx').exists() else Path('决赛发布材料/1002PPT决赛完善版/智烤管家_10021912_蓝白山海光路无框视觉决赛版.pptx'))")

new_bg = r'''def add_bg(slide,road=True,seed=1):
    bg_path = BG_FULL if road else BG_REDUCED
    pic=slide.shapes.add_picture(bg_path,0,0,width=Inches(W),height=Inches(H)); pic.name='背景图_允许不可编辑'
'''
src = re.sub(r"def add_bg\(slide,\s*road=True,\s*seed=1\):.*?\ndef icon_home", new_bg + "\ndef icon_home", src, flags=re.S)

new_title = r'''def add_title(slide,title,subtitle=None):
    add_text(slide,title,.82,.31,11.7,.72,31,WHITE,True,TITLE_FONT,PP_ALIGN.CENTER)
    line(slide,1.2,.72,2.95,.72,CYAN,1.1); line(slide,10.38,.72,12.13,.72,CYAN,1.1)
    add_text(slide,'///',.72,.56,.35,.22,9,CYAN,True,BODY_FONT,PP_ALIGN.CENTER); add_text(slide,'///',12.22,.56,.35,.22,9,CYAN,True,BODY_FONT,PP_ALIGN.CENTER)
    if subtitle: add_text(slide,subtitle,1.4,.94,10.5,.34,11,MUTED,False,BODY_FONT,PP_ALIGN.CENTER)
'''
src = re.sub(r"def add_title\(slide,title,subtitle=None\):.*?\ndef add_bg", new_title + "\ndef add_bg", src, flags=re.S)

cover = """s=ns(1,True)
    add_text(s,'Z H I   K A O   M A N A G E R',3.7,.68,6,.35,12,ICE,False,BODY_FONT,PP_ALIGN.CENTER)
    add_text(s,'智烤',2.35,1.30,4.15,1.1,54,WHITE,True,TITLE_FONT,PP_ALIGN.CENTER)
    add_text(s,'管家',6.55,1.30,4.15,1.1,54,CYAN,True,TITLE_FONT,PP_ALIGN.CENTER)
    halo=s.shapes.add_shape(MSO_SHAPE.ARC,Inches(2.25),Inches(.10),Inches(8.85),Inches(6.35)); halo.fill.background(); halo.line.color.rgb=rgb(CYAN); halo.line.width=Pt(1.0); halo.rotation=180
    line(s,1.18,2.55,1.18,5.30,CYAN,.8); line(s,12.12,2.55,12.12,5.30,CYAN,.8)
    add_text(s,'让每一座烤房，都有一份可计算的“健康档案”',2,2.58,9.3,.5,15,ICE,False,BODY_FONT,PP_ALIGN.CENTER)
    add_text(s,'一库清底   |   双模决策   |   三智落地   |   四队长效',2.2,3.42,8.9,.5,15,ICE,True,BODY_FONT,PP_ALIGN.CENTER)
    add_text(s,'基于健康评估的烤房用投管一体化解决方案',2.1,4.28,9.1,.48,15,WHITE,True,BODY_FONT,PP_ALIGN.CENTER)
    add_text(s,'湖北烟草创客大赛 · 决赛发布    宜昌市烟草公司 · 智同道合创客小组    2026.10',1.6,6.82,10.2,.32,10,MUTED,False,BODY_FONT,PP_ALIGN.CENTER)
    s=ns(2,True);"""
src = re.sub(r"s=ns\(1,True\);.*?s=ns\(2,True\);", cover, src, flags=re.S)

prep = r'''
def prepare_backgrounds():
    with zipfile.ZipFile(SOURCE_PPT) as z:
        data=z.read('ppt/media/image1.jpg')
    raw=Path('/tmp/zhikao_source_cover.jpg'); raw.write_bytes(data)
    im=Image.open(raw).convert('RGB'); w,h=im.size
    veil=Image.new('RGBA',im.size,(0,0,0,0)); d=ImageDraw.Draw(veil)
    d.rectangle((0,0,w,int(h*.60)),fill=(4,20,42,230)); d.rectangle((0,int(h*.83),w,h),fill=(4,20,42,205))
    full=Image.alpha_composite(im.convert('RGBA'),veil).convert('RGB'); full=ImageEnhance.Brightness(full).enhance(.90); full.save(BG_FULL,quality=93)
    veil2=Image.new('RGBA',im.size,(0,0,0,0)); d2=ImageDraw.Draw(veil2)
    d2.rectangle((0,0,w,int(h*.66)),fill=(4,20,42,238)); d2.rectangle((0,int(h*.55),w,h),fill=(4,20,42,185))
    red=Image.alpha_composite(im.convert('RGBA'),veil2).convert('RGB'); red=ImageEnhance.Brightness(red).enhance(.72); red.save(BG_REDUCED,quality=91)
'''
idx=src.find('build_deck(PPT_A')
src = src[:idx] + prep + '\nprepare_backgrounds()\n' + src[idx:]

src = re.sub(r"# strict check: no slide pictures/media.*?print\('CREATED',PPT_A,PPT_B,DOCX\)", r'''# validation: exactly one non-editable background picture per slide; all foreground is editable
for p in [PPT_A,PPT_B,DOCX]:
    with zipfile.ZipFile(p) as z:
        bad=z.testzip(); assert bad is None, (p,bad)
for p in [PPT_A,PPT_B]:
    with zipfile.ZipFile(p) as z:
        media=[n for n in z.namelist() if n.startswith('ppt/media/') and not n.endswith('/')]
        slides=[n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')]
        assert 1 <= len(media) <= 2, media
        for n in slides:
            assert z.read(n).count(b'<p:pic') == 1, (n,z.read(n).count(b'<p:pic'))
        print('EDITABLE_FOREGROUND_OK',p.name,'slides=15','background_media=',len(media))
print('CREATED',PPT_A,PPT_B,DOCX)''', src, flags=re.S)

patched=Path('/tmp/patched_builder_10022216.py'); patched.write_text(src,encoding='utf-8')
exec(compile(src,str(patched),'exec'),{'__name__':'__main__','__file__':str(patched)})
