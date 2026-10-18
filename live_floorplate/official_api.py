"""Request parameters from the official Landmark Third Party Integration guide.

No HTTP, credential storage, signing guesses, or response-schema assumptions.
The documented sign parameter is the fixed API key, supplied via Secrets.
"""
from .model import FloorplateError,number

BASE='https://service.arcnet.com.au/arc-service-api/v1/project/'
OPTIONAL_LIST_PARAMS={'searchkey','category','minprice','maxprice','israngesearch',
                      'bedroomqty','bathroomqty','studymediaqty','carspaceqty',
                      'pagenum','pagesize','sortby'}

# Page 4 example order; separate from normal omission of empty optional values.
LIST_EXAMPLE_ORDER=('searchkey','category','minprice','maxprice','israngesearch',
                    'bedroomqty','bathroomqty','studymediaqty','carspaceqty',
                    'pagenum','pagesize','sortby')


def project_list_comparison_parameters(config):
    """Build the required-only and exact empty-filter example variants."""
    url, query = request_parameters(config,'getProjectList')
    if set(query) != {'agencyid','sign'}:
        raise FloorplateError('Comparison requires only agency and key configuration')
    required = {key:query[key] for key in ('agencyid','sign')}
    return url, required, {**required, **dict.fromkeys(LIST_EXAMPLE_ORDER,'')}


def request_parameters(config,endpoint='getProject'):
    if endpoint not in ('getProject','getProjectList'):
        raise FloorplateError('Unsupported official read endpoint')
    url=BASE+endpoint
    if config.get('url',url)!=url:
        raise FloorplateError('Official endpoint must be configured without a query string')
    supplied=config.get('query',{})
    required={'agencyid','sign'} | ({'projectid'} if endpoint=='getProject' else set())
    allowed=required | (OPTIONAL_LIST_PARAMS if endpoint=='getProjectList' else set())
    if not set(supplied)<=allowed:
        raise FloorplateError('Undocumented endpoint parameter')
    query={}
    for key,value in supplied.items():
        if value is None or value=='':continue
        if isinstance(value,bool) or not isinstance(value,(str,int)):
            raise FloorplateError('Invalid official request parameter')
        value=str(value).strip()
        if not value or 'VERIFY_' in value or '{' in value or any(ord(c)<32 for c in value):
            raise FloorplateError('Official request configuration is incomplete')
        query[key]=value
    if not required<=set(query):
        raise FloorplateError('Missing official request parameter')
    for key in ('agencyid','projectid'):
        if key in query and (not query[key].isascii() or not query[key].isdigit() or int(query[key])<=0):
            raise FloorplateError('Invalid official identity parameter')
    if endpoint=='getProjectList':
        if ('pagenum' in query)!=('pagesize' in query):
            raise FloorplateError('Project-list paging parameters must be paired')
        for key in ('pagenum','pagesize'):
            if key in query and (not query[key].isascii() or not query[key].isdigit() or int(query[key])<=0):
                raise FloorplateError('Invalid project-list paging parameter')
        for key in ('bedroomqty','bathroomqty','studymediaqty','carspaceqty'):
            if key in query and ('israngesearch' not in query or not query[key].isascii() or not query[key].isdigit()):
                raise FloorplateError('Quantity filter requires range mode and a single integer')
        if 'israngesearch' in query and query['israngesearch'] not in ('0','1'):
            raise FloorplateError('Invalid quantity range mode')
        if 'maxprice' in query and 'minprice' not in query:
            raise FloorplateError('Maximum price requires minimum price')
        for key in ('minprice','maxprice'):
            if key in query:number(query[key])
        if 'maxprice' in query and number(query['maxprice'])<number(query['minprice']):
            raise FloorplateError('Invalid price range')
        if 'category' in query and not set(query['category'].split(','))<= {'01','02','03','04'}:
            raise FloorplateError('Invalid documented category')
        if 'sortby' in query and query['sortby'] not in ('1','2'):
            raise FloorplateError('Invalid documented sort option')
    return url,query
