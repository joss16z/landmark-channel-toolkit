"""Synthetic diagnostics only: never contact Arcanite."""
import io
import unittest
from email.message import Message
from unittest.mock import Mock, patch
from urllib.error import HTTPError
from urllib.parse import urlsplit, parse_qs
from live_floorplate.connection_test import connection_report
from live_floorplate.api import fetch_json_response, NoRedirect
from live_floorplate.model import FloorplateError

CONFIG={'agencyid':'123456789','sign':'SYNTHETIC_SECRET','projectid':'907440240'}
PAYLOAD={'data':{'properties':[{'lotnumber':1,'unitnumber':'PRIVATE_UNIT',
    'price':'PRIVATE_PRICE','propertystatus':'Available'},
    {'lotnumber':2,'propertystatus':'PRIVATE_STATUS'}]}}


class ConnectionTests(unittest.TestCase):
    def test_documented_request_reaches_transport_as_get_query_not_body(self):
        # Exercise all three modules together, mocking only the network boundary.
        # The key is synthetic; reserved characters must survive ONE decode.
        for key in ('SYNTHETIC_SECRET', 'SYNTHETIC+a/b&c=d%25?e#f ü'):
            for config in ({**CONFIG, 'sign':key}, {'query':{**CONFIG,'sign':key}}):
                response=Mock(status=200,url='https://service.arcnet.com.au/arc-service-api/v1/project/getProject')
                response.headers=Message();response.headers['Content-Type']='application/json'
                response.read.return_value=b'{"properties":[{"lotnumber":1,"propertystatus":"Available"}]}'
                with patch('live_floorplate.api.build_opener') as opener:
                    opener.return_value.open.return_value.__enter__.return_value=response
                    report=connection_report(config)
                    self.assertTrue(report['ok'])
                    request=opener.return_value.open.call_args.args[0]
                    parsed=urlsplit(request.full_url)
                    self.assertEqual(parsed.scheme,'https')
                    self.assertEqual(parsed.netloc,'service.arcnet.com.au')
                    self.assertEqual(parsed.path,'/arc-service-api/v1/project/getProject')
                    self.assertEqual(parsed.fragment,'')
                    self.assertEqual(request.get_method(),'GET')
                    self.assertIsNone(request.data)
                    self.assertEqual(parse_qs(parsed.query),{
                        'agencyid':[CONFIG['agencyid']], 'projectid':[CONFIG['projectid']], 'sign':[key]})
                    self.assertFalse(request.has_header('Authorization'))
                    self.assertFalse(request.has_header('Cookie'))
                    self.assertEqual(opener.return_value.open.call_args.kwargs,{'timeout':20})
                    self.assertIsInstance(opener.call_args.args[0],NoRedirect)
                    response.read.assert_called_once_with(10_000_001)

    def test_http_403_is_transport_status_not_schema_failure_or_auth_classification(self):
        # The client cannot distinguish a vendor 403 from a gateway/WAF 403.
        body=io.BytesIO(b'PRIVATE_PRICE PRIVATE_UNIT SYNTHETIC_SECRET')
        error=HTTPError('https://fixture.invalid/?sign=SYNTHETIC_SECRET',403,
                        'PRIVATE_UNIT',{},body)
        with patch('live_floorplate.api.build_opener') as opener, \
                patch('live_floorplate.connection_test.inspect_response') as inspect:
            opener.return_value.open.side_effect=error
            report=connection_report(CONFIG)
            self.assertEqual(report,{'http_status':403,'ok':False})
            inspect.assert_not_called()
            opener.return_value.open.assert_called_once()
            self.assertTrue(body.closed)

    def test_structure_only_and_both_secrets_formats(self):
        for config in (CONFIG,{'query':CONFIG}):
            with patch('live_floorplate.connection_test.fetch_json_response',return_value=(200,PAYLOAD)) as fetch:
                result=connection_report(config)
                self.assertTrue(result['ok'])
                self.assertEqual(result['http_status'],200)
                array=result['summary']['property_arrays'][0]
                self.assertEqual(array['unique_lots'],2)
                self.assertEqual(array['status_counts']['Available'],1)
                self.assertEqual(array['status_counts']['unknown'],1)
                self.assertFalse(result['summary']['endpoint_completeness_verified'])
                self.assertEqual(fetch.call_args.args[1],{})
                self.assertTrue(fetch.call_args.args[0].startswith('https://service.arcnet.com.au/arc-service-api/v1/project/getProject?'))
                for value in ('SYNTHETIC_SECRET','PRIVATE_UNIT','PRIVATE_PRICE','PRIVATE_STATUS'):
                    self.assertNotIn(value,str(result))

    def test_missing_conflicting_secrets_never_request(self):
        for cfg in ({},{**CONFIG,'sign':None},{**CONFIG,'query':{**CONFIG,'sign':'OTHER'}}):
            with patch('live_floorplate.connection_test.fetch_json_response') as fetch:
                self.assertFalse(connection_report(cfg)['ok']);fetch.assert_not_called()

    def test_failure_schema_and_exceptions_are_redacted(self):
        for response in ((401,None),(403,None),(200,None),(200,{'message':'PRIVATE_PRICE'}),(200,'PRIVATE_UNIT')):
            with patch('live_floorplate.connection_test.fetch_json_response',return_value=response):
                result=connection_report(CONFIG)
                self.assertFalse(result['ok'])
                self.assertNotIn('PRIVATE_',str(result))
        with patch('live_floorplate.connection_test.fetch_json_response',side_effect=ValueError('SYNTHETIC_SECRET')):
            self.assertEqual(connection_report(CONFIG),{'http_status':None,'ok':False})

    def test_transport_http_status_and_no_error_body(self):
        for status in (401,403,302):
            error=HTTPError('https://fixture.invalid/?sign=SYNTHETIC_SECRET',status,'PRIVATE_UNIT',{},io.BytesIO(b'PRIVATE_PRICE'))
            with patch('live_floorplate.api.build_opener') as opener:
                opener.return_value.open.side_effect=error
                self.assertEqual(fetch_json_response('https://fixture.invalid/',{}),(status,None))
        with self.assertRaises(FloorplateError):
            NoRedirect().redirect_request(None,None,302,'',{},'https://other.invalid/')

    def test_transport_content_limits_and_origin(self):
        for body,content_type,url in ((b'bad','application/json','https://fixture.invalid/'),
                (b'{}','text/html','https://fixture.invalid/'),
                (b'x'*10_000_001,'application/json','https://fixture.invalid/'),
                (b'{}','application/json','https://other.invalid/')):
            response=Mock(status=200,url=url)
            response.headers=Message();response.headers['Content-Type']=content_type
            response.read.return_value=body
            with patch('live_floorplate.api.build_opener') as opener:
                opener.return_value.open.return_value.__enter__.return_value=response
                self.assertEqual(fetch_json_response('https://fixture.invalid/',{}),(200,None))

    def test_button_is_explicit_and_never_creates_snapshot(self):
        from streamlit.testing.v1 import AppTest
        from pathlib import Path
        app_path=str(Path(__file__).resolve().parents[1]/'app.py')
        with patch('live_floorplate.ui.connection_report',return_value={'http_status':403,'ok':False}) as report:
            app=AppTest.from_file(app_path,default_timeout=20)
            app.secrets['arcanite']=CONFIG
            app.run();app.sidebar.radio[0].set_value('🏢 Live Floorplate').run()
            report.assert_not_called()
            next(b for b in app.button if b.label=='Test Arcanite Connection').click().run()
            report.assert_called_once()
            self.assertEqual(len(app.exception),0)
            self.assertEqual(len(app.error),1)
            self.assertNotIn('live_floorplate_snapshot',app.session_state)
            self.assertNotIn('Configured Agent Portal API',app.radio[0].options)

    def test_successful_button_shows_only_structure(self):
        from streamlit.testing.v1 import AppTest
        from pathlib import Path
        app_path=str(Path(__file__).resolve().parents[1]/'app.py')
        with patch('live_floorplate.connection_test.fetch_json_response',return_value=(200,PAYLOAD)) as fetch:
            app=AppTest.from_file(app_path,default_timeout=20)
            app.secrets['arcanite']=CONFIG
            app.run();app.sidebar.radio[0].set_value('🏢 Live Floorplate').run()
            fetch.assert_not_called()
            next(b for b in app.button if b.label=='Test Arcanite Connection').click().run()
            self.assertEqual(len(app.exception),0)
            self.assertEqual(len(app.error),0)
            self.assertGreater(len(app.dataframe),0)
            displayed=str([frame.value.to_dict() for frame in app.dataframe])
            for sentinel in ('SYNTHETIC_SECRET','PRIVATE_UNIT','PRIVATE_PRICE','PRIVATE_STATUS'):
                self.assertNotIn(sentinel,displayed)
            self.assertNotIn('live_floorplate_snapshot',app.session_state)

    def test_missing_secrets_button_fails_without_network(self):
        from streamlit.testing.v1 import AppTest
        from pathlib import Path
        with patch('live_floorplate.connection_test.fetch_json_response') as fetch:
            app=AppTest.from_file(str(Path(__file__).resolve().parents[1]/'app.py'),default_timeout=20).run()
            app.sidebar.radio[0].set_value('🏢 Live Floorplate').run()
            next(b for b in app.button if b.label=='Test Arcanite Connection').click().run()
            self.assertEqual(len(app.exception),0)
            self.assertEqual(len(app.error),1)
            fetch.assert_not_called()
