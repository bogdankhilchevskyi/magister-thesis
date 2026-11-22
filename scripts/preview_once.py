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
toc_pages={i for i in range(1,min(4,len(doc))) if doc[i].get_links()}
row_heads=['Психологічні','Педагогічні','Технологічні','Соціальні','Організаційні','Клінічна значущість і вибір запиту','Поетапне розв’язання і формувальний відгук','Безпечний діалог і наставництво','Адаптивна організація і повернення']
detailed={str(i+1):[{'kind':v['kind'],'page':v.get('page'),'name':v.get('nameddest')} for v in doc[i].get_links()] for i in toc_pages}
manifest={'toc_links':detailed,'tables':{str(i+1):' '.join(t.split())[-170:] for i,t in enumerate(texts) if any(w in t for w in ['Таблиця 2.','Джерельна опора','Джерельна','Межа застосування та джерела'])},'pages':len(doc),'fonts':fonts,'outline':doc.get_toc(), 'page_sizes':list({str(p.rect) for p in doc}), 'links':[(i+1,len(p.get_links())) for i,p in enumerate(doc) if p.get_links()][:8], 'blank':[i+1 for i,s in enumerate(texts) if len(s.strip())<5], 'headers':[p.get_text(clip=fitz.Rect(0,0,595,49)).strip() for p in doc], 'overview':{str(i+1):' '.join(texts[i].split())[:130] for i in selected}}
(proof/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False),encoding='utf-8')
