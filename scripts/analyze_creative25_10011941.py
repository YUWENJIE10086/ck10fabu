from pathlib import Path
from PIL import Image, ImageOps, ImageDraw, ImageFont
import csv, math, os, json

src=Path('创客25图片_解压_10011941')
out=Path('creative25_analysis_10011941')
out.mkdir(exist_ok=True)
thumb_dir=out/'thumbs'; thumb_dir.mkdir(exist_ok=True)
files=[]
for p in src.rglob('*'):
    if p.suffix.lower() in {'.png','.jpg','.jpeg','.webp','.bmp'}:
        try:
            im=Image.open(p).convert('RGB')
            files.append((p, im.size[0], im.size[1]))
            im.thumbnail((800,450))
            canvas=Image.new('RGB',(820,500),'white')
            x=(820-im.width)//2; y=15+(450-im.height)//2
            canvas.paste(im,(x,y))
            d=ImageDraw.Draw(canvas)
            label=p.relative_to(src).as_posix()
            d.rectangle((0,465,820,500),fill='white')
            d.text((10,470),label,fill='black')
            canvas.save(thumb_dir/(f'{len(files):02d}_'+p.stem[:50]+'.jpg'),quality=88)
        except Exception as e:
            print('skip',p,e)

with open(out/'manifest.csv','w',newline='',encoding='utf-8-sig') as f:
    w=csv.writer(f); w.writerow(['idx','file','width','height','ratio'])
    for i,(p,w0,h0) in enumerate(files,1):
        w.writerow([i,p.relative_to(src).as_posix(),w0,h0,round(w0/h0,4)])

# contact sheets, 6 thumbnails/sheet
thumbs=sorted(thumb_dir.glob('*.jpg'))
for k in range(0,len(thumbs),6):
    group=thumbs[k:k+6]
    sheet=Image.new('RGB',(1640,1500),(240,242,245))
    for j,p in enumerate(group):
        im=Image.open(p).convert('RGB')
        r=j//2; c=j%2
        sheet.paste(im,(c*820,r*500))
    sheet.save(out/f'contact_{k//6+1:02d}.jpg',quality=88)
print('images',len(files))
