"""Explicit, operator-reviewed lot positions; no inference from unit/lot numbers."""
from dataclasses import dataclass
import math
from .model import FloorplateError, lot_key

LEVELS = tuple(range(9))


def level_title(level):
    return 'Ground Level' if level == 0 else f'Level {level}'


@dataclass(frozen=True)
class Position:
    lot: str
    level: int
    x: float
    y: float
    width: float
    height: float
    group: int = -1
    box_id: str = ''
    anchor_x: float = 0
    anchor_y: float = 0


@dataclass(frozen=True)
class Layout:
    pages: dict
    positions: dict


def parse_layout(data, page_count, pdf_bytes=None):
    if (not isinstance(data, dict) or data.get('version') not in (1,2)
            or data.get('project') != 'Hurstville Rise' or data.get('verified') is not True):
        raise FloorplateError('Reviewed Hurstville Rise mapping required')
    if data['version']==2:
        from .source_mapping import verify_source
        if pdf_bytes is None:
            raise FloorplateError('Source PDF required to verify mapping')
        verify_source(data,pdf_bytes)
    pages = data['pages']
    if set(pages) != {str(i) for i in LEVELS}:
        raise FloorplateError('Ground through Level 8 must be mapped explicitly')
    if any(type(v) is not int or not 0 <= v < page_count for v in pages.values()):
        raise FloorplateError('Invalid PDF page mapping')
    if not (pages['2'] == pages['3'] == pages['4'] and pages['5'] == pages['6'] == pages['7']):
        raise FloorplateError('Shared floor layouts must reference the same PDF page')
    positions = {}
    for row in data['positions']:
        lot, level = lot_key(row['lotnumber']), row['level']
        if lot in positions or type(level) is not int or level not in LEVELS:
            raise FloorplateError('Duplicate or invalid lot position')
        coords = [row[k] for k in ('x','y','width','height')]
        if any(isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) for v in coords):
            raise FloorplateError('Invalid position coordinates')
        x,y,w,h = coords
        if not (0 <= x < 1 and 0 <= y < 1 and w > 0 and h > 0 and x+w <= 1 and y+h <= 1):
            raise FloorplateError('Position outside PDF bounds')
        from .source_mapping import GROUPS
        group=next(g[0] for g in GROUPS if level in g[4])
        if data['version']==2 and row['group']!=group:
            raise FloorplateError('Lot floor group mismatch')
        positions[lot] = Position(lot,level,x,y,w,h,group,row.get('box_id',lot),
                                  row.get('anchor_x',x+w/2),row.get('anchor_y',y+h/2))
    return Layout({int(k):v for k,v in pages.items()},positions)


def matched(snapshot, layout, level=None):
    if not snapshot.fresh():
        raise FloorplateError('Snapshot expired; synchronize again')
    if any(a.lot not in layout.positions for a in snapshot.apartments):
        raise FloorplateError('Available lots lack confirmed positions')
    return [(a,layout.positions[a.lot]) for a in snapshot.apartments
            if level is None or layout.positions[a.lot].level == level]
