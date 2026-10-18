"""Server-only GET adapter. No credentials, endpoint or response contract defaults.

The configuration must describe a verified Agent Portal getProject request.
No endpoint discovery, authentication automation, write requests or redirects.
"""
import json
from urllib.parse import urlsplit, urljoin, urlencode, urlunsplit
from urllib.request import Request, build_opener, HTTPRedirectHandler
from .model import FloorplateError, snapshot

FIELDS = ('project','lotnumber','propertystatus','unit','bedrooms','bathrooms','internal','external','price')


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise FloorplateError('Redirect refused')


def origin(url):
    try:
        p = urlsplit(url)
        if (p.scheme != 'https' or not p.hostname or p.username is not None or p.password is not None or p.fragment
                or '\\' in url or '%' in p.netloc or any(ord(c) <= 32 or ord(c) == 127 for c in url) or p.port == 0):
            raise ValueError()
        return p.hostname.lower(), p.port or 443
    except (ValueError, TypeError):
        raise FloorplateError('Invalid endpoint URL') from None


def path_value(value, path):
    if not isinstance(path, str):
        raise FloorplateError('Unverified response path')
    for part in path.split('.') if path else ():
        value = value[part]
    return value


def fetch_json_response(url, headers):
    """Internal transport: return status and parsed body, never transport exceptions.

    Caller must not display the body. Redirects are rejected before forwarding
    credentials; errors retain only a numeric HTTP status, when available.
    """
    from urllib.error import HTTPError
    status = None
    try:
        with build_opener(NoRedirect()).open(Request(url, headers=headers, method='GET'), timeout=20) as response:
            status = response.status
            if origin(response.url) != origin(url) or status != 200:
                return status, None
            if response.headers.get_content_type() != 'application/json':
                return status, None
            data = response.read(10_000_001)
            if len(data) > 10_000_000:
                return status, None
            return status, json.loads(data)
    except HTTPError as error:
        status = error.code
        error.close()
    except Exception:
        pass
    return status, None


def fetch_page(url, headers):
    status, payload = fetch_json_response(url, headers)
    if status != 200 or payload is None:
        raise FloorplateError('Official API request failed; inspect access and contract in the approved environment') from None
    return payload



def sync(config, fetch=fetch_page):
    """Total counts all scoped rows, including unavailable properties.

    Single-page and same-origin next-link pagination only. Other response shapes
    require an explicitly reviewed adapter; no invented API structure is assumed.
    """
    if config.get('enabled') is not True or config.get('contract_verified') is not True:
        raise FloorplateError('API contract not verified')
    if config.get('account_scope') != 'agent_portal' or config.get('method') != 'GET':
        raise FloorplateError('Only approved Agent Portal read requests supported')
    url = config['url']
    if config.get('api_profile')=='arcanite_landmark_v1':
        from .official_api import request_parameters
        url,official_query=request_parameters(config,'getProject')
        config={**config,'query':official_query}
        # The guide does not document property next-link paging. Do not follow
        # guessed links with the API key; a different contract needs verification.
        if config.get('pagination')!='single':
            raise FloorplateError('Official property pagination is not established')
    allowed = origin(url)
    # Signed query values are Secrets, never browser props/report metadata.
    query = config.get('query', {})
    if query:
        parsed = urlsplit(url)
        if parsed.query:
            raise FloorplateError('Configure query secrets separately from endpoint URL')
        if any(not isinstance(v,(str,int)) or isinstance(v,bool) or not str(v).strip()
               or 'VERIFY_' in str(v) or '{' in str(v) for v in query.values()):
            raise FloorplateError('Query configuration is incomplete')
        url = urlunsplit((parsed.scheme,parsed.netloc,parsed.path,urlencode(query),''))
    fields = config['fields']
    if set(fields) not in (set(FIELDS),set(FIELDS)-{'project'}) or fields['lotnumber'] != 'lotnumber' or fields['propertystatus'] != 'propertystatus':
        raise FloorplateError('Verified field mapping required')
    if any(not isinstance(v,str) or not v.strip() or 'VERIFY_' in v for v in fields.values()):
        raise FloorplateError('Unverified field mapping')
    headers = dict(config.get('headers', {}))
    if any(k.lower() in ('host','cookie','proxy-authorization') for k in headers):
        raise FloorplateError('Unsupported authentication header')
    if config.get('pagination') not in ('single', 'next_link'):
        raise FloorplateError('Pagination contract is unknown')
    if config['pagination'] == 'next_link' and not config.get('next_path'):
        raise FloorplateError('Missing pagination path')
    # Project identity may come from a verified single-project request context,
    # not an invented property-level project field.
    if 'project' not in fields and (config.get('project_context_verified') is not True
            or not config.get('query',{}).get('projectid')):
        raise FloorplateError('Single-project request context is unverified')
    proof=config.get('completeness','total_field')
    if proof not in ('total_field','expected_lot_set'):
        raise FloorplateError('Unknown completeness proof')
    expected=None
    if proof=='expected_lot_set':
        from .model import lot_key
        configured=config.get('expected_lots')
        if not isinstance(configured,list) or not configured:
            raise FloorplateError('An explicit reviewed complete Lot set is required')
        expected={lot_key(lot) for lot in configured}
        if len(expected)!=len(configured):
            raise FloorplateError('Ambiguous complete Lot set')
    rows, seen, total = [], set(), None
    for _ in range(200):
        if origin(url) != allowed or url in seen:
            raise FloorplateError('Unsafe or repeated pagination URL')
        seen.add(url)
        payload = fetch(url, headers)
        if config.get('success_path'):
            success=path_value(payload,config['success_path'])
            expected_success=config.get('success_value')
            if type(success) is not type(expected_success) or success!=expected_success:
                raise FloorplateError('API response does not confirm success')
        if config.get('scope_path'):
            if str(path_value(payload,config['scope_path'])) != str(config['query']['projectid']):
                raise FloorplateError('Response project identity differs from request')
        count = path_value(payload, config['total_path']) if proof=='total_field' else len(expected)
        if type(count) is not int or count < 0 or (total is not None and count != total):
            raise FloorplateError('Missing or changing total')
        total = count
        page = path_value(payload, config['rows_path'])
        if not isinstance(page, list):
            raise FloorplateError('Response schema changed')
        for item in page:
            record={k:path_value(item,path) for k,path in fields.items()}
            if 'project' not in fields:
                record['project']=config['project']
            rows.append(record)
        if len(rows) > total:
            raise FloorplateError('More rows than displayed total')
        next_url = path_value(payload, config['next_path']) if config['pagination'] == 'next_link' else None
        if next_url is None:
            if expected is not None and {lot_key(r['lotnumber']) for r in rows}!=expected:
                raise FloorplateError('Complete scoped Lot set not returned')
            return snapshot(rows, total, config['project'])
        if not isinstance(next_url, str) or not next_url or not page:
            raise FloorplateError('Incomplete pagination')
        url = urljoin(url, next_url)
    raise FloorplateError('Pagination limit exceeded')
