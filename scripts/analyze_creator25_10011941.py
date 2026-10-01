from pathlib import Path
from PIL import Image, ImageOps, ImageDraw, ImageFont
import zipfile, csv, math, os, json

ZIP = Path('创客25图片.zip')
EXTRACT = Path('_tmp_creator25_extract')
OUT = Path('_tmp_creator25_review')
EXTRACT.mkdir(exist_ok=True)
OUT.mkdir(exist_ok=True)

with zipfile.ZipFile(ZIP, 'r') as z:
    z.extractall(EXTRACT)

exts={'.png','.jpg','.jpeg','.webp','.bmp','.tif','.tiff'}
files=sorted([p for p in EXTRACT.rglob('*') if p.suffix.lower() in exts])
rows=[]
thumbs=[]
for idx,p in enumerate(files,1):
    try:
        im=Image.open(p).convert('RGB')
    except Exception as e:
        continue
    w,h=im.size
    # create readable individual review image
    maxw,maxh=2200,1400
    scale=min(maxw/w,maxh/h,1.0)
    nw,nh=max(1,int(w*scale)),max(1,int(h*scale))
    rim=im.resize((nw,nh),Image.LANCZOS) if scale<1 else im.copy()
    canvas=Image.new('RGB',(maxw,maxh+90),(20,24,32))
    x=(maxw-nw)//2; y=65+(maxh-nh)//2
    canvas.paste(rim,(x,y))
    d=ImageDraw.Draw(canvas)
    label=f'{idx:02d}  {p.name}   {w}×{h}'
    d.rectangle([0,0,maxw,60], fill=(8,16,32))
    d.text((24,18),label,fill=(230,240,255))
    out=OUT/f'{idx:02d}_{p.stem}.jpg'
    canvas.save(out,quality=90,optimize=True)
    thumbs.append(out)
    # simple palette stats
    small=im.copy(); small.thumbnail((160,160))
    pal=small.quantize(colors=5).getpalette()[:15]
    colors=[]
    for i in range(0,len(pal),3):
        colors.append('#%02X%02X%02X'%tuple(pal[i:i+3]))
    rows.append({'index':idx,'name':p.name,'path':str(p.relative_to(EXTRACT)),'width':w,'height':h,'dominant5':' '.join(colors)})

with open(OUT/'manifest.csv','w',newline='',encoding='utf-8-sig') as f:
    wr=csv.DictWriter(f,fieldnames=['index','name','path','width','height','dominant5'])
    wr.writeheader(); wr.writerows(rows)

# 2x2 contact sheets from individual review images
per=4
for s in range(0,len(thumbs),per):
    chunk=thumbs[s:s+per]
    cellw,cellh=1100,745
    sheet=Image.new('RGB',(cellw*2,cellh*2),(10,14,24))
    for j,p in enumerate(chunk):
        im=Image.open(p).convert('RGB')
        im.thumbnail((cellw-20,cellh-20),Image.LANCZOS)
        cx=(j%2)*cellw+(cellw-im.width)//2
        cy=(j//2)*cellh+(cellh-im.height)//2
        sheet.paste(im,(cx,cy))
    sheet.save(OUT/f'sheet_{s//per+1:02d}.jpg',quality=88,optimize=True)

summary={'count':len(rows),'files':[r['name'] for r in rows]}
(OUT/'summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(summary,ensure_ascii=False))
