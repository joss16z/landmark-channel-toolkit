"""Source evidence checks for the six-page Hurstville PDF. Never use unit IDs."""
import hashlib
import math
import re
import fitz
from PIL import Image,ImageChops
from .model import FloorplateError

GROUPS = ((0,'Ground',1,4,(0,)), (1,'Level 1',5,18,(1,)),
          (2,'Levels 2–4',19,60,(2,3,4)),(3,'Levels 5–7',61,105,(5,6,7)),
          (4,'Level 8',106,115,(8,)))


def source_numbers(pdf):
    result = {}
    with fitz.open(stream=pdf,filetype='pdf') as doc:
        if len(doc) != 6 or doc.needs_pass:
            raise FloorplateError('Expected original six-page PDF')
        for group,title,start,end,levels in GROUPS:
            page=doc[group+1]
            for block in page.get_text('dict')['blocks']:
                for line in block.get('lines',[]):
                    text=''.join(s['text'] for s in line['spans']).strip()
                    if re.fullmatch(r'\d(?:\s*\d){0,2}',text):
                        lot=int(re.sub(r'\s','',text))
                        if not start <= lot <= end or str(lot) in result:
                            raise FloorplateError('Unexpected or duplicate source Lot number')
                        result[str(lot)]=(group+1,tuple(line['bbox']))
        if set(result) != {str(i) for i in range(1,116)}:
            raise FloorplateError('Source Lot coverage is not 115/115')
    return result


def verify_source(mapping,pdf):
    if mapping.get('source_sha256') != hashlib.sha256(pdf).hexdigest():
        raise FloorplateError('Mapping belongs to a different PDF')
    source=source_numbers(pdf)
    rows=mapping['positions']
    if len(rows)!=115 or {str(r['lotnumber']) for r in rows} != set(source):
        raise FloorplateError('Incomplete Lot mapping')
    for row in rows:
        bbox_values=row['source_bbox']
        values=[*bbox_values,*[row[k] for k in ('x','y','width','height','anchor_x','anchor_y')]]
        if len(bbox_values)!=4 or any(isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) for v in values):
            raise FloorplateError('Invalid source coordinate evidence')
        if type(row['source_page']) is not int or type(row['group']) is not int:
            raise FloorplateError('Invalid source page/group')
        if not (0<=row['x']<1 and 0<=row['y']<1 and row['width']>0 and row['height']>0
                and row['x']+row['width']<=1 and row['y']+row['height']<=1
                and 0<=row['anchor_x']<=1 and 0<=row['anchor_y']<=1):
            raise FloorplateError('Coordinates outside source page')
        page,bbox=source[str(row['lotnumber'])]
        if row['source_page']!=page or any(abs(a-b)>.01 for a,b in zip(bbox,row['source_bbox'])) or len(row['source_bbox'])!=4:
            raise FloorplateError('Lot evidence differs from source PDF')
    boxes={}
    with fitz.open(stream=pdf,filetype='pdf') as doc:
        for row in rows:
            group=next((g for g in GROUPS if g[2]<=int(row['lotnumber'])<=g[3]),None)
            if group is None or row['group']!=group[0] or row['level'] not in group[4]:
                raise FloorplateError('Lot outside source floor group')
            boxes.setdefault(row['box_id'],[]).append(row)
        if len(boxes)!=57:
            raise FloorplateError('Expected 57 source footprints')
        for box in boxes.values():
            first=box[0];group=GROUPS[first['group']]
            ordered=sorted(box,key=lambda r:sum(r['source_bbox'][1::2]))
            if [r['level'] for r in ordered]!=list(group[4]):
                raise FloorplateError('Shared-floor stack order differs from PDF')
            page=doc[first['source_page']]
            centers_x=[(r['source_bbox'][0]+r['source_bbox'][2])/2 for r in box]
            centers_y=[(r['source_bbox'][1]+r['source_bbox'][3])/2 for r in box]
            if max(centers_x)-min(centers_x)>=5 or max(centers_y)-min(centers_y)>=45:
                raise FloorplateError('Lots are not members of one source number stack')
            cx=sum((r['source_bbox'][0]+r['source_bbox'][2])/2 for r in box)/len(box)
            cy=sum((r['source_bbox'][1]+r['source_bbox'][3])/2 for r in box)/len(box)
            geometry=('x','y','width','height','anchor_x','anchor_y')
            if any(any(r[k]!=first[k] for k in geometry) for r in box):
                raise FloorplateError('Shared-footprint coordinates differ')
            if abs(first['anchor_x']*page.rect.width-cx)>.01 or abs(first['anchor_y']*page.rect.height-cy)>.01:
                raise FloorplateError('Anchor differs from source number stack')
        for group in range(5):
            page=doc[group+1]
            images=page.get_image_info(xrefs=True)
            if len(images)!=1:
                raise FloorplateError('Source artwork structure changed')
            entry=images[0];pix=fitz.Pixmap(doc,entry['xref'])
            if pix.n!=3:
                raise FloorplateError('Unsupported source artwork color space')
            image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
            r,g,b=image.split()
            bounds=ImageChops.darker(ImageChops.darker(r,g),b).point(lambda v:255 if v<130 else 0).getbbox()
            if bounds is None:
                raise FloorplateError('Source apartment artwork unavailable')
            rect=fitz.Rect(entry['bbox'])
            artwork=fitz.Rect(rect.x0+bounds[0]/pix.width*rect.width,rect.y0+bounds[1]/pix.height*rect.height,
                              rect.x0+bounds[2]/pix.width*rect.width,rect.y0+bounds[3]/pix.height*rect.height)
            regions=[]
            for values in boxes.values():
                row=values[0]
                if row['group']!=group:continue
                region=fitz.Rect(row['x']*page.rect.width,row['y']*page.rect.height,
                                 (row['x']+row['width'])*page.rect.width,(row['y']+row['height'])*page.rect.height)
                if region.intersects(artwork) or any(region.intersects(other) for other in regions):
                    raise FloorplateError('Information box covers apartment artwork or another box')
                regions.append(region)
    return True
