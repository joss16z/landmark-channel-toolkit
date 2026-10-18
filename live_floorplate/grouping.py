"""A shared source footprint owns one box, containing Available apartments only."""
from collections import defaultdict
from .layout import matched
from .model import FloorplateError,format_price


def information_boxes(snapshot,layout,group):
    pairs=matched(snapshot,layout)
    grouped=defaultdict(list)
    for apartment,pos in pairs:
        if pos.group==group:
            grouped[pos.box_id].append((apartment,pos))
    boxes=[]
    for key,pairs in grouped.items():
        pairs.sort(key=lambda p:(p[1].level,p[0].lot))
        first=pairs[0][1]
        if any((p.x,p.y,p.width,p.height,p.anchor_x,p.anchor_y)!=(first.x,first.y,first.width,first.height,first.anchor_x,first.anchor_y) for a,p in pairs):
            raise FloorplateError('Shared footprint has conflicting anchors')
        apartments=[a for a,p in pairs]
        signature=lambda a:(a.configuration,a.internal,a.external)
        if len({signature(a) for a in apartments})==1:
            a=apartments[0]
            lines=[a.configuration,f'{a.internal:g}m² + {a.external:g}m²']
            lines += [f'Unit {a.unit} - {format_price(a.price)}' for a in apartments]
        else:
            lines=[]
            for a in apartments:
                lines += [a.configuration,f'{a.internal:g} + {a.external:g} m²',
                          f'Unit {a.unit} - {format_price(a.price)}']
        boxes.append((first,'\n'.join(lines),tuple(apartments)))
    return sorted(boxes,key=lambda b:b[0].box_id)
