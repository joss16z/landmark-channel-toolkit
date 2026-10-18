"""Synthetic domain data and drawings; NOT a claimed getProject response/map."""
import fitz
from .model import snapshot
from .layout import parse_layout


def source_preview(layout):
    """All source Lots Available for a conspicuously synthetic layout stress test."""
    rows=[dict(project='Hurstville Rise',lotnumber=lot,propertystatus='Available',
               unit=f'TEST{lot}',bedrooms='2 Bed',bathrooms=2,internal=75,external=12,price=1000000)
          for lot in layout.positions]
    return snapshot(rows,len(rows),'Hurstville Rise',sample=True)


def demonstration():
    rows,positions = [],[]
    for level in range(9):
        for index,status in enumerate(('Available','Held','Pending','Sold')):
            lot = f'DEMO-{level}-{index}'
            rows.append(dict(project='Hurstville Rise',lotnumber=lot,propertystatus=status,
                             unit=f'DEMO-{level}0{index}',bedrooms='2 Bed',bathrooms=2,internal=75,external=12,price=1000000))
            positions.append(dict(lotnumber=lot,level=level,x=.08+index*.23,y=.55,width=.2,height=.2))
            group=0 if level==0 else 1 if level==1 else 2 if level<=4 else 3 if level<=7 else 4
            positions[-1]['box_id']=f'{group}-{index}'
    doc = fitz.open()
    for title in ('Ground Level','Level 1','Levels 2–4','Levels 5–7','Level 8'):
        page = doc.new_page(width=960,height=540)
        page.insert_text((40,55),'SYNTHETIC DEMO — NOT HURSTVILLE RISE ARCHITECTURE',fontsize=18)
        page.insert_text((40,90),title,fontsize=16)
        for i in range(4):
            page.draw_rect(fitz.Rect(80+i*220,140,260+i*220,400))
    pdf = doc.tobytes()
    doc.close()
    mapping = dict(version=1,project='Hurstville Rise',verified=True,
                   pages={'0':0,'1':1,'2':2,'3':2,'4':2,'5':3,'6':3,'7':3,'8':4},positions=positions)
    return snapshot(rows,len(rows),'Hurstville Rise',sample=True),parse_layout(mapping,5),pdf
