"""One visual contract for HTML and PowerPoint; PDF is read-only background."""
import base64
from dataclasses import dataclass
from html import escape
import io
import re
import fitz
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
from floorplate_update_report_v8 import _shape_has_coloured_solid_fill, normalize_unit
from .model import FloorplateError
from .layout import LEVELS, level_title, matched

FLOOR_GROUPS = ((0,), (1,), (2, 3, 4), (5, 6, 7), (8,))

def grouped_labels(snapshot, layout, levels):
    """Combine Available units at identical confirmed label positions across shared levels."""
    groups = {}
    for apartment, position in matched(snapshot, layout):
        if position.level not in levels:
            continue
        key = (position.x, position.y, position.width, position.height)
        groups.setdefault(key, []).append((position.level, apartment, position))
    return [(items[0][2], '\n'.join(a.label for _, a, _ in sorted(items, key=lambda item: (item[0], item[1].unit)))) for items in groups.values()]



@dataclass(frozen=True)
class Style:
    fill: str = 'FFF200'
    ink: str = '000000'
    font: str = 'Arial'
    size: float = 12
    slide_width: float = 13.333333
    slide_height: float = 7.5


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


def html_view(snapshot, layout, pdf_bytes, level, style=Style()):
    levels = level if isinstance(level, tuple) else (level,)
    png,width,height = pdf_image(pdf_bytes,layout.pages[levels[0]])
    boxes = []
    for pos, label_text in grouped_labels(snapshot, layout, levels):
        label = escape(label_text).replace('\\n','<br>')
        boxes.append(f'<div class="label" style="left:{pos.x*100}%;top:{pos.y*100}%;width:{pos.width*100}%;height:{pos.height*100}%" aria-label="Available apartment"><span>{label}</span></div>')
    return f'''<!doctype html><html><head><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; script-src 'unsafe-inline'"><style>
    body{{margin:0;font-family:Arial}} .canvas{{position:relative;width:100%;aspect-ratio:{width}/{height}}}
    img{{width:100%;height:100%;display:block}} .label{{position:absolute;box-sizing:border-box;display:flex;align-items:center;justify-content:center;text-align:center;background:#{style.fill};color:#{style.ink};font-family:{escape(style.font)};font-size:clamp(8px,1.2vw,16px);border:1px solid #333;overflow:auto}}
    </style></head><body><div class="canvas"><img alt="Original floorplate drawing" src="data:image/png;base64,{base64.b64encode(png).decode()}">{''.join(boxes)}</div></body></html>'''


def powerpoint(snapshot, layout, pdf_bytes, style=Style()):
    matched(snapshot,layout)
    prs = Presentation()
    sw,sh = style.slide_width,style.slide_height
    prs.slide_width,prs.slide_height = Inches(sw),Inches(sh)
    for levels in FLOOR_GROUPS:
        level = levels[0]
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        title = slide.shapes.add_textbox(Inches(.2),0,Inches(sw-.4),Inches(.5))
        group_title = level_title(level) if len(levels) == 1 else f'Levels {levels[0]}–{levels[-1]}'
        title.text = f'{"SAMPLE — " if snapshot.sample else ""}{snapshot.project} — {group_title}'
        title.text_frame.paragraphs[0].font.size = Pt(20)
        png,pw,ph = pdf_image(pdf_bytes,layout.pages[level])
        ratio = min(sw/pw,(sh-1.2)/ph)
        width,height = pw*ratio,ph*ratio
        left,top = (sw-width)/2,.6+(sh-1.2-height)/2
        slide.shapes.add_picture(io.BytesIO(png),Inches(left),Inches(top),Inches(width),Inches(height))
        for pos,label_text in grouped_labels(snapshot,layout,levels):
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(left+pos.x*width),
                Inches(top+pos.y*height),Inches(pos.width*width),Inches(pos.height*height))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(style.fill)
            shape.line.color.rgb = RGBColor(0,0,0)
            shape.text_frame.word_wrap = True
            shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            shape.text_frame.margin_left = shape.text_frame.margin_right = Inches(.03)
            shape.text = label_text
            for p in shape.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                p.font.name,p.font.size = style.font,Pt(style.size)
                p.font.color.rgb = RGBColor.from_string(style.ink)
        footer = slide.shapes.add_textbox(Inches(.2),Inches(sh-.5),Inches(sw-.4),Inches(.35))
        footer.text = f'As of {snapshot.synced_at.isoformat()} — {"Synthetic demonstration" if snapshot.sample else "Agent Portal scope; owner review required"}'
        footer.text_frame.paragraphs[0].font.size = Pt(9)
    result = io.BytesIO()
    prs.save(result)
    matched(snapshot,layout)
    return result.getvalue()
