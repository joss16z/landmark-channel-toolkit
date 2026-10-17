"""Server-only GET adapter. No credentials, endpoint or response contract defaults.

The configuration must describe a verified Agent Portal getProject request.
No endpoint discovery, authentication automation, write requests or redirects.
"""
import json
from urllib.parse import urlsplit, urljoin, urlencode, urlunsplit
from urllib.request import Request, build_opener, HTTPRedirectHandler
from .model import FloorplateError, snapshot

FIELDS = ('project','lotnumber','propertystatus','unit','bedrooms','internal','external','price')


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


def fetch_page(url, headers):
    # Redirects rejected BEFORE sending headers to any redirect destination.
    with build_opener(NoRedirect()).open(Request(url, headers=headers, method='GET'), timeout=20) as response:
        if origin(response.url) != origin(url) or response.status != 200:
            raise FloorplateError('Endpoint access failed')
        if response.headers.get_content_type() != 'application/json':
            raise FloorplateError('Unexpected endpoint content')
        data = response.read(10_000_001)
        if len(data) > 10_000_000:
            raise FloorplateError('Endpoint response too large')
        return json.loads(data)


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
    if set(fields) != set(FIELDS) or fields['lotnumber'] != 'lotnumber' or fields['propertystatus'] != 'propertystatus':
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
    rows, seen, total = [], set(), None
    for _ in range(200):
        if origin(url) != allowed or url in seen:
            raise FloorplateError('Unsafe or repeated pagination URL')
        seen.add(url)
        payload = fetch(url, headers)
        count = path_value(payload, config['total_path'])
        if type(count) is not int or count < 0 or (total is not None and count != total):
            raise FloorplateError('Missing or changing total')
        total = count
        page = path_value(payload, config['rows_path'])
        if not isinstance(page, list):
            raise FloorplateError('Response schema changed')
        rows.extend({k:path_value(row, fields[k]) for k in FIELDS} for row in page)
        if len(rows) > total:
            raise FloorplateError('More rows than displayed total')
        next_url = path_value(payload, config['next_path']) if config['pagination'] == 'next_link' else None
        if next_url is None:
            return snapshot(rows, total, config['project'])
        if not isinstance(next_url, str) or not next_url or not page:
            raise FloorplateError('Incomplete pagination')
        url = urljoin(url, next_url)
    raise FloorplateError('Pagination limit exceeded')
