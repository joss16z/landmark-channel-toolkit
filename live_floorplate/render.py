"""One visual contract for HTML and PowerPoint; PDF is read-only background."""
import base64
from dataclasses import dataclass
from html import escape
import io
import re
import textwrap
import math
import fitz
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
from floorplate_update_report_v8 import _shape_has_coloured_solid_fill, normalize_unit
from .model import FloorplateError
from .layout import matched
from .source_mapping import GROUPS
from .grouping import information_boxes

FLOOR_GROUPS=tuple(group[4] for group in GROUPS)


def grouped_labels(snapshot,layout,levels):
    """Compatibility with the existing feature branch's five-group helper."""
    group=FLOOR_GROUPS.index(tuple(levels))
    return [(pos,label) for pos,label,apartments in information_boxes(snapshot,layout,group)]


@dataclass(frozen=True)
class Style:
    fill: str = 'C2C7AD'
    ink: str = '000000'
    font: str = 'Arial'
    size: float = 12
    slide_width: float = 16.535433
    slide_height: float = 11.692913


def walk_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == 6:
            yield from walk_shapes(shape.shapes)


def reference_style(pptx_bytes):
    """Reuse v8 label detection; copy styling only, never template data/drawings."""
    prs = Presentation(io.BytesIO(pptx_bytes))
    for slide in prs.slides:
        for shape in walk_shapes(slide.shapes):
            if (_shape_has_coloured_solid_fill(shape) and getattr(shape,'has_text_frame',False)
                    and 'unit' in shape.text.lower()):
                try:
                    color = str(shape.fill.fore_color.rgb)
                    runs = [r for p in shape.text_frame.paragraphs for r in p.runs]
                    font = runs[0].font if runs else None
                    ink = str(font.color.rgb) if font and font.color.type else '000000'
                    name = font.name or 'Arial' if font else 'Arial'
                    if not re.fullmatch(r'[A-Za-z0-9 _-]+',name):
                        raise FloorplateError('Unsupported font name')
                    return Style(color,ink,name,font.size.pt if font and font.size else 12,
                                 prs.slide_width/914400,prs.slide_height/914400)
                except (ValueError,AttributeError):
                    raise FloorplateError('Reference label uses unsupported theme styling') from None
    raise FloorplateError('No colored Unit label style found in reference PowerPoint')


def pdf_image(pdf_bytes, page_index):
    with fitz.open(stream=pdf_bytes,filetype='pdf') as doc:
        if doc.needs_pass or len(doc) > 100:
            raise FloorplateError('Unsupported PDF')
        page = doc[page_index]
        scale = min(2, 2400 / max(page.rect.width,page.rect.height))
        pix = page.get_pixmap(matrix=fitz.Matrix(scale,scale),alpha=False)
        return pix.tobytes('png'), pix.width, pix.height


def fitted_label(label,pos,style):
    # Preserve the reference PPT's one-unit-per-line hierarchy; never wrap a price row.
    for size in [style.size-i*.25 for i in range(max(1,int((style.size-7)*4)+1))]:
        if size<7: break
        chars=max(1,int((pos.width*style.slide_width*72-10)/(size*.53)))
        lines=[]
        for line in label.splitlines():
            if line.startswith('Unit ') and len(line)>chars:
                # Preserve content: long demo unit/price rows may use two lines.
                unit, separator, price = line.partition(' - ')
                if separator and len(unit)<=chars and len(price)<=chars:
                    lines.extend([unit,price])
                else:
                    lines.extend(textwrap.wrap(line,width=chars,break_long_words=False,
                                               break_on_hyphens=False) or [line])
            else:
                lines.extend(textwrap.wrap(line,width=chars,break_long_words=False,
                                           break_on_hyphens=False) or [line])
        if not lines or any(len(line)>chars for line in lines): continue
        if len(lines)*size*1.19 <= pos.height*style.slide_height*72-8:
            return '\n'.join(lines),size
    raise FloorplateError('Information box cannot fit all Available data legibly')


def html_view(snapshot,layout,pdf_bytes,group,style=Style()):
    if isinstance(group,tuple):
        group=FLOOR_GROUPS.index(group)
    png,width,height=pdf_image(pdf_bytes,layout.pages[GROUPS[group][4][0]])
    boxes=[];leaders=[]
    for pos,label,apartments in information_boxes(snapshot,layout,group):
        label,size=fitted_label(label,pos,style)
        parts=label.splitlines()
        label=''.join(f'<span class="label-row {"unit-row" if line.startswith("Unit ") else "detail-row"}">{escape(line)}</span>' for line in parts)
        x,y=pos.x+pos.width/2,pos.y+pos.height/2
        leaders.append(f'<line x1="{x*1000}" y1="{y*1000}" x2="{pos.anchor_x*1000}" y2="{pos.anchor_y*1000}" stroke="#777" stroke-width=".6"/>')
        boxes.append(f'<div class="label" style="left:{pos.x*100}%;top:{pos.y*100}%;width:{pos.width*100}%;height:{pos.height*100}%;font-size:{size*100/(style.slide_width*72)}cqw" aria-label="Available apartments"><span>{label}</span></div>')
    return f'''<!doctype html><html><head><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; script-src 'unsafe-inline'"><style>
    body{{margin:0}} .canvas{{position:relative;width:100%;aspect-ratio:{width}/{height};container-type:inline-size}}
    img,svg{{width:100%;height:100%;display:block}} svg{{position:absolute;inset:0;pointer-events:none}}
    .label{{position:absolute;box-sizing:border-box;display:flex;align-items:center;justify-content:center;text-align:center;background:#{style.fill};color:#{style.ink};font-family:"{escape(style.font)}",Arial,sans-serif;line-height:1.16;font-weight:700;border:2px solid #222;border-radius:16px;overflow:hidden;padding:4px 5px}}\n    .label>span{{display:flex;flex-direction:column;align-items:center;width:100%}} .label-row{{display:block;white-space:nowrap;width:100%;font-weight:700}}
    </style></head><body><div class="canvas"><img alt="Original floorplate artwork" src="data:image/png;base64,{base64.b64encode(png).decode()}"><svg viewBox="0 0 1000 1000" preserveAspectRatio="none">{''.join(leaders)}</svg>{''.join(boxes)}</div></body></html>'''


def powerpoint(snapshot,layout,pdf_bytes,style=Style()):
    matched(snapshot,layout)
    prs=Presentation();sw,sh=style.slide_width,style.slide_height
    prs.slide_width,prs.slide_height=Inches(sw),Inches(sh)
    for group,title,start,end,levels in GROUPS:
        slide=prs.slides.add_slide(prs.slide_layouts[6])
        png,_,_=pdf_image(pdf_bytes,layout.pages[levels[0]])
        slide.shapes.add_picture(io.BytesIO(png),0,0,Inches(sw),Inches(sh))
        for pos,label,apartments in information_boxes(snapshot,layout,group):
            from pptx.enum.shapes import MSO_CONNECTOR
            line=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                Inches((pos.x+pos.width/2)*sw),Inches((pos.y+pos.height/2)*sh),
                Inches(pos.anchor_x*sw),Inches(pos.anchor_y*sh))
            line.line.color.rgb=RGBColor.from_string('777777');line.line.width=Pt(.5)
            shape=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(pos.x*sw),Inches(pos.y*sh),
                                        Inches(pos.width*sw),Inches(pos.height*sh))
            shape.fill.solid();shape.fill.fore_color.rgb=RGBColor.from_string(style.fill)
            shape.line.color.rgb=RGBColor.from_string('222222')
            shape.line.width=Pt(1.5)
            shape.text_frame.word_wrap=False;shape.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE
            shape.text_frame.margin_left=shape.text_frame.margin_right=Inches(.04)
            shape.text_frame.margin_top=shape.text_frame.margin_bottom=Inches(.04)
            shape.text,size=fitted_label(label,pos,style)
            for paragraph in shape.text_frame.paragraphs:
                paragraph.alignment=PP_ALIGN.CENTER
                paragraph.space_before=Pt(0)
                paragraph.space_after=Pt(0)
                paragraph.font.name,paragraph.font.size=style.font,Pt(size)
                paragraph.font.bold=True
                paragraph.font.color.rgb=RGBColor.from_string(style.ink)
        footer=slide.shapes.add_textbox(Inches(3.5),Inches(sh-.23),Inches(sw-4),Inches(.18))
        footer.text=f'{"SAMPLE — " if snapshot.sample else ""}{title} | As of {snapshot.synced_at.isoformat()} | Owner review required'
        footer.text_frame.paragraphs[0].font.size=Pt(7)
    result=io.BytesIO();prs.save(result);matched(snapshot,layout)
    return result.getvalue()
