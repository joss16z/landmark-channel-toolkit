"""Offline official-response review: structural diagnostics only, no raw values.

Run on the approved computer with an authorized response saved outside Git.
This tool makes no HTTP requests and does not certify endpoint completeness.
"""
import argparse
import json
from pathlib import Path
import re
from .model import lot_key


def inspect_response(payload):
    arrays=[];metadata=[];visited=0
    def walk(value,path,depth=0):
        nonlocal visited
        visited+=1
        if depth>12 or visited>10000:
            raise ValueError('Response exceeds inspection limits')
        if isinstance(value,dict):
            for key,item in value.items():
                if not isinstance(key,str) or not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]{0,79}',key):
                    continue
                child=f'{path}.{key}' if path else key
                if re.search(r'total|count|page|hasmore|projectid|success|errorcode|^code$|^status$',key,re.I):
                    metadata.append({'path':child,'type':type(item).__name__})
                walk(item,child,depth+1)
        elif isinstance(value,list):
            rows=[row for row in value if isinstance(row,dict)]
            if any('lotnumber' in row for row in rows):
                fields={};keys=[];invalid=0;statuses={s:0 for s in ('Available','Held','Pending','Sold','unknown')}
                def field_types(row,prefix='',depth=0):
                    if depth>8:raise ValueError('Nested field limit exceeded')
                    for key,item in row.items():
                        if isinstance(key,str) and re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]{0,79}',key):
                            name=f'{prefix}.{key}' if prefix else key
                            fields.setdefault(name,set()).add(type(item).__name__)
                            if isinstance(item,dict):field_types(item,name,depth+1)
                for row in rows:
                    field_types(row)
                    try:keys.append(lot_key(row.get('lotnumber')))
                    except ValueError:invalid+=1
                    status=row.get('propertystatus')
                    statuses[status if isinstance(status,str) and status in statuses else 'unknown']+=1
                unique=set(keys)
                arrays.append({'rows_path':path,'row_count':len(value),'object_count':len(rows),
                               'fields':{k:sorted(v) for k,v in sorted(fields.items())},
                               'valid_lot_rows':len(keys),'invalid_lot_rows':invalid,
                               'unique_lots':len(unique),'duplicate_lot_rows':len(keys)-len(unique),
                               'exact_lots_1_to_115':unique=={str(i) for i in range(1,116)}
                                   and len(value)==115 and len(keys)==115,
                               'status_counts':statuses})
            # Inspect one representative object for nested lists/metadata, not every
            # property value. Field-type coverage above still covers all rows.
            if rows:walk(rows[0],path+'[]',depth+1)
    walk(payload,'')
    return {'property_arrays':arrays,'metadata_candidates':metadata,
            'endpoint_completeness_verified':False,'contains_raw_field_values':False}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--response',type=Path,required=True)
    args=parser.parse_args()
    try:
        path=args.response.resolve()
        if path.is_relative_to(Path(__file__).resolve().parents[1]) or path.stat().st_size>10_000_000:
            raise ValueError('Use a private response outside the repository')
        result=inspect_response(json.loads(path.read_text(encoding='utf-8-sig')))
    except Exception:
        print(json.dumps({'error':'RESPONSE_INSPECTION_FAILED'}))
        raise SystemExit(1)
    print(json.dumps(result,indent=2))


if __name__=='__main__':main()
