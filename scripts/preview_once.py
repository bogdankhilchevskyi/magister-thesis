from pathlib import Path
import json
import fitz
from PIL import Image, ImageOps, ImageDraw
root=Path.cwd()
proof=root/'build-proof'
proof.mkdir(exist_ok=True)
doc=fitz.open(next((root/'final').glob('*.pdf')))
texts=[p.get_text() for p in doc]
selected=sorted({0,1,2,len(doc)-1}|{i for i,p in enumerate(doc) if p.get_images() or any(t in texts[i] for t in ['Таблиця 2.', 'СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ'])})
fonts=sorted({f[3] for p in doc for f in p.get_fonts()})
thumbs=Image.new('RGB',(3*310,((len(selected)+2)//3)*459),'#babab8')
for j,i in enumerate(selected):
    p=doc[i]
    pix=p.get_pixmap(matrix=fitz.Matrix(1,1),alpha=False)
    im=Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
    im.thumbnail((300,424))
    x=(j%3)*310+5
    y=(j//3)*459+3
    thumbs.paste(im,(x,y))
    ImageDraw.Draw(thumbs).text((x+12,y+425),str(i+1),fill='black')
thumbs.save(proof/'preview.jpg',quality=84)
manifest={'pages':len(doc),'fonts':fonts,'outline':doc.get_toc(), 'page_sizes':list({str(p.rect) for p in doc}), 'links':[(i+1,len(p.get_links())) for i,p in enumerate(doc) if p.get_links()][:8], 'blank':[i+1 for i,s in enumerate(texts) if len(s.strip())<5], 'headers':[p.get_text(clip=fitz.Rect(0,0,595,49)).strip() for p in doc], 'overview':{str(i+1):' '.join(texts[i].split())[:130] for i in selected}}
(proof/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False),encoding='utf-8')
