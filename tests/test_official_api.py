"""Official documented parameter contract; no actual API credentials or requests."""
import unittest
from unittest.mock import Mock
from live_floorplate.official_api import BASE,request_parameters
from live_floorplate.api import sync
from live_floorplate.model import FloorplateError


def config(endpoint='getProject'):
    query={'agencyid':'123456789','sign':'SYNTHETIC_NOT_AN_API_KEY'}
    if endpoint=='getProject':query['projectid']='907440240'
    return {'url':BASE+endpoint,'query':query}


class OfficialRequestTests(unittest.TestCase):
    def test_get_project_uses_exact_documented_parameters(self):
        url,query=request_parameters(config())
        self.assertEqual(url,BASE+'getProject')
        self.assertEqual(set(query),{'agencyid','sign','projectid'})
        self.assertEqual(query['projectid'],'907440240')

    def test_missing_credentials_or_id_fail_without_exposing_values(self):
        for field in ('agencyid','sign','projectid'):
            cfg=config();cfg['query'].pop(field)
            with self.assertRaises(FloorplateError) as error:request_parameters(cfg)
            self.assertNotIn('SYNTHETIC_NOT_AN_API_KEY',str(error.exception))

    def test_get_project_does_not_accept_project_list_filters_or_unknown_endpoints(self):
        for key in ('pagenum','pagesize','category','propertystatus'):
            cfg=config();cfg['query'][key]='1'
            with self.assertRaises(FloorplateError):request_parameters(cfg)
        with self.assertRaises(FloorplateError):request_parameters(config(),'addClient')
        cfg=config();cfg['url']='https://evil.invalid/'
        with self.assertRaises(FloorplateError):request_parameters(cfg)

    def test_project_list_supports_documented_parameters_and_dependencies(self):
        cfg=config('getProjectList')
        cfg['query'].update(pagenum=1,pagesize=10,category='01,03',israngesearch=0,bedroomqty=2,
                            minprice=100000,maxprice=200000,searchkey='Synthetic',sortby=1)
        url,query=request_parameters(cfg,'getProjectList')
        self.assertEqual(url,BASE+'getProjectList')
        self.assertEqual(query['israngesearch'],'0')
        for changes in ({'pagenum':1},{'maxprice':100},{'bedroomqty':2},{'sortby':3}):
            bad=config('getProjectList');bad['query'].update(changes)
            with self.assertRaises(FloorplateError):request_parameters(bad,'getProjectList')

    def test_official_profile_never_follows_undocumented_property_paging(self):
        cfg=config();cfg.update(api_profile='arcanite_landmark_v1',enabled=True,contract_verified=True,
                                account_scope='agent_portal',method='GET',pagination='next_link')
        fetch=Mock()
        with self.assertRaises(FloorplateError):sync(cfg,fetch)
        fetch.assert_not_called()

    def test_transport_errors_redact_signed_urls_in_normal_tracebacks(self):
        import traceback
        from unittest.mock import patch
        from live_floorplate.api import fetch_page
        signed='https://fixture.invalid/?sign=SYNTHETIC_SECRET_SENTINEL'
        with patch('live_floorplate.api.build_opener') as opener:
            opener.return_value.open.side_effect=OSError(signed)
            try:fetch_page(signed,{})
            except FloorplateError as error:
                self.assertNotIn('SYNTHETIC_SECRET_SENTINEL',''.join(traceback.format_exception(error)))
            else:self.fail('Transport failure must fail closed')
