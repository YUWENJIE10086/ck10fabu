# -*- coding: utf-8 -*-
"""原始资料可审计提取器：遍历用户指定四类资料，保留源路径并建立文字底稿。
只做忠实提取，不把历史方案当作已落地效果。"""
import os, re, io, sys, json, hashlib, subprocess, zipfile
from pathlib import Path
from html import unescape
ROOT=Path(__file__).resolve().parents[2]
OUT=Path(__file__).resolve().parent
SRC=ROOT/"读取资料"
STAMP="10192000"
for p in ["docx","pptx","bs4","openpyxl","xlrd","fitz"]:
 try: __import__(p)
 except ImportError: pass
from docx import Document
from pptx import Presentation
from bs4 import BeautifulSoup
import openpyxl, xlrd, fitz

def plain(s):
 return re.sub(r'\n{3,}','\n\n',str(s).replace('\x00','').strip())
def extract(path):
 e=path.suffix.lower()
 if e=='.txt': 
  for enc in ('utf-8-sig','gb18030','utf-16'):
   try:return path.read_text(encoding=enc)
   except Exception:continue
 if e in ('.html','.htm','.drawio'):
  b=path.read_bytes()
  txt=b.decode('utf-8','replace')
  if e=='.drawio':
   soup=BeautifulSoup(txt,'xml')
   z=[unescape(x.get('value','')) for x in soup.find_all('mxCell') if x.get('value')]
   return '\n'.join(BeautifulSoup(y,'html.parser').get_text(' ',strip=True) for y in z)
  s=BeautifulSoup(txt,'html.parser')
  for x in s(['script','style','svg','nav']): x.decompose()
  return s.get_text('\n',strip=True)
 if e=='.docx':
  d=Document(str(path)); out=[]
  for p in d.paragraphs:
   if p.text.strip():out.append(p.text)
  for table in d.tables:
   out.append('[表格]')
   for row in table.rows:out.append(' | '.join(c.text.replace('\n','；') for c in row.cells))
  return '\n'.join(out)
 if e=='.doc':
  try:
   p=subprocess.run(['antiword','-w','180',str(path)],stdout=subprocess.PIPE,stderr=subprocess.DEVNULL,timeout=35)
   return p.stdout.decode('utf-8','replace')
  except Exception as ex:return f'[DOC解析失败: {ex}]'
 if e=='.pptx':
  p=Presentation(str(path)); z=[]
  for i,s in enumerate(p.slides,1):
   z.append(f'[第{i}页]')
   for sh in s.shapes:
    if sh.has_text_frame and sh.text.strip():z.append(sh.text)
   if s.has_notes_slide:
    no=s.notes_slide.notes_text_frame.text.strip()
    if no:z.append('[备注]'+no)
  return '\n'.join(z)
 if e=='.xlsx':
  w=openpyxl.load_workbook(str(path),read_only=True,data_only=True)
  z=[]
  for sh in w:
   z.append('[工作表]'+sh.title)
   for row in sh.iter_rows(max_row=min(sh.max_row,250),values_only=True):
    v=[str(x) if x is not None else '' for x in row[:35]]
    if any(v):z.append('|'.join(v))
  return '\n'.join(z)
 if e=='.xls':
  w=xlrd.open_workbook(str(path));z=[]
  for sh in w.sheets():
   z.append('[工作表]'+sh.name)
   for i in range(min(sh.nrows,250)):
    z.append('|'.join(str(x) for x in sh.row_values(i)[:35]))
  return '\n'.join(z)
 if e=='.pdf':
  d=fitz.open(str(path));return '\n'.join(p.get_text() for p in list(d)[:80])
 return '[非可直接抽取文字的素材，已登记文件清单，需结合图像或原始软件核验]'

def dump(paths,target,max_chars=None):
 lines=[f'# {target.stem}',f'自动抽取；每段标注源文件路径；内容保留原意，除不可识别字符不进行事实校正。','','---','']
 for p in paths:
  try:
   s=plain(extract(p)); err=''
  except Exception as ex:
   s='[本文件自动解析失败，需人工核查]';err=str(ex)
  rel=p.relative_to(ROOT).as_posix()
  lines.extend([f'## 源文件：{rel}',f'原始文件字节：{p.stat().st_size}；文本字符：{len(s)}', '',s if max_chars is None else s[:max_chars], '', '---',''])
 target.write_text('\n'.join(lines),encoding='utf-8')
 return [(p.relative_to(ROOT).as_posix(),p.stat().st_size) for p in paths]

parts=[sorted((SRC/x).rglob('*')) for x in ('1学习优秀的','2我的项目相关的零碎的','3之前项目官方叙述过的','4一版初稿')]
parts=[[p for p in group if p.is_file() and p.name!='.DS_Store'] for group in parts]
a,b,c,d=parts
print("source files:",[len(x) for x in parts])
dump(a,OUT/f'{STAMP}_01获奖解说词逐份全文摘录.md')
dump(b,OUT/f'{STAMP}_02项目全部文件逐项提取.md',max_chars=25000)
dump(c,OUT/f'{STAMP}_03官方材料逐字提取.md')
chosen=[ROOT/'1004'/'智烤管家演讲讲稿9.30.docx']+d
dump([x for x in chosen if x.exists()],OUT/f'{STAMP}_04初稿全文逐段摘录.md')
styles=[]
p=ROOT/'1004'/'智烤管家演讲讲稿9.30.docx'
if p.exists():
 D=Document(str(p))
 for i,x in enumerate(D.paragraphs):
  if not x.text.strip():continue
  styles.append(dict(index=i,text=x.text,style=x.style.name,alignment=str(x.alignment),line_spacing=str(x.paragraph_format.line_spacing),space_after=str(x.paragraph_format.space_after),runs=[dict(text=r.text,font=r.font.name,size=(r.font.size.pt if r.font.size else None),bold=r.bold,color=(str(r.font.color.rgb) if r.font.color and r.font.color.rgb else None)) for r in x.runs]))
(OUT/f'{STAMP}_初稿格式参数.json').write_text(json.dumps(styles,ensure_ascii=False,indent=2),encoding='utf-8')
summary=['# 资料全量登记','',f'参考解说词：{len(a)}；项目零散材料：{len(b)}；官方文档：{len(c)}；早期初稿：{len(d)}','']
for t,group in zip(['获奖解说词','项目素材','官方表述','早期稿'],parts):
 summary.extend([f'## {t}',''])
 for p in group:summary.append(f'- {p.relative_to(ROOT).as_posix()}（{p.stat().st_size}字节）')
(OUT/f'{STAMP}_资料完整清单.md').write_text('\n'.join(summary),encoding='utf-8')
print('complete')
