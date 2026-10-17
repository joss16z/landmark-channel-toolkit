"""Explicit server-side connection diagnostic; never creates inventory snapshots."""
from urllib.parse import urlencode
from .api import fetch_json_response
from .official_api import request_parameters
from .schema_review import inspect_response


def connection_report(config):
    # Accept the direct Secrets keys and the existing nested query template.
    # No headers, configurable host, pagination or redirects in this diagnostic.
    status = None
    try:
        query = {}
        nested = config.get('query', {})
        for key in ('agencyid', 'sign', 'projectid'):
            direct = config.get(key)
            existing = nested.get(key)
            if direct is not None and existing is not None and str(direct) != str(existing):
                raise ValueError('Conflicting configuration')
            query[key] = direct if direct is not None else existing
        url, query = request_parameters({'query': query})
        status, payload = fetch_json_response(url + '?' + urlencode(query), {})
        if status != 200 or not isinstance(payload, (dict, list)):
            return {'http_status': status, 'ok': False}
        summary = inspect_response(payload)
        # A property-shaped array is necessary for a useful diagnostic, but is
        # NOT proof of successful authentication, completeness or price validity.
        return {'http_status': status, 'ok': bool(summary['property_arrays']),
                'summary': summary}
    except Exception:
        return {'http_status': status, 'ok': False}
