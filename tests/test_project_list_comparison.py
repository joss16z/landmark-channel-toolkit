"""No live traffic: compare the documented query variants with synthetic Secrets."""
import unittest
from unittest.mock import patch
from urllib.parse import urlsplit, parse_qsl
from live_floorplate.connection_test import compare_project_list_requests
from live_floorplate.official_api import request_parameters

CONFIG={'agencyid':'123456789','sign':'SYNTHETIC+a/b&x=%25','projectid':'907440240'}
OPTIONALS=('searchkey','category','minprice','maxprice','israngesearch','bedroomqty',
           'bathroomqty','studymediaqty','carspaceqty','pagenum','pagesize','sortby')


class ProjectListComparisonTests(unittest.TestCase):
    def test_exact_a_b_parameters_order_empty_values_and_single_encoding(self):
        for config in (CONFIG,{'query':{**CONFIG,'searchkey':'IGNORED_PRIVATE_FILTER'}}):
            with patch('live_floorplate.connection_test.fetch_json_response',
                       side_effect=[(403,{'private':'PRIVATE_RESPONSE'}),(200,{'private':'PRIVATE_RESPONSE'})]) as fetch:
                result=compare_project_list_requests(config)
                self.assertEqual(result,{'A':403,'B':200})
                self.assertEqual(fetch.call_count,2)
                for index,call in enumerate(fetch.call_args_list):
                    url,headers=call.args
                    self.assertEqual(headers,{})
                    parsed=urlsplit(url)
                    self.assertEqual(parsed.scheme+'://'+parsed.netloc+parsed.path,
                        'https://service.arcnet.com.au/arc-service-api/v1/project/getProjectList')
                    self.assertEqual(parsed.fragment,'')
                    pairs=parse_qsl(parsed.query,keep_blank_values=True)
                    required=[('agencyid',CONFIG['agencyid']),('sign',CONFIG['sign'])]
                    self.assertEqual(pairs,required+([(name,'') for name in OPTIONALS] if index else []))
                    if index:self.assertTrue(parsed.query.endswith('sortby='))
                self.assertNotIn(CONFIG['sign'],str(result))
                self.assertNotIn('PRIVATE_RESPONSE',str(result))

    def test_normal_builder_still_omits_empty_optionals(self):
        _,query=request_parameters({'query':{'agencyid':CONFIG['agencyid'],'sign':CONFIG['sign'],
                                            **dict.fromkeys(OPTIONALS,'')}},'getProjectList')
        self.assertEqual(set(query),{'agencyid','sign'})

    def test_validation_failure_prevents_both_requests(self):
        for config in ({},{'sign':CONFIG['sign']},{**CONFIG,'query':{'sign':'CONFLICT'}}):
            with patch('live_floorplate.connection_test.fetch_json_response') as fetch:
                self.assertEqual(compare_project_list_requests(config),{'A':None,'B':None})
                fetch.assert_not_called()

    def test_transport_failure_is_redacted_and_does_not_prevent_other_variant(self):
        for responses,expected in (([OSError('PRIVATE_SIGNED_URL'),(403,None)],{'A':None,'B':403}),
                                   ([(200,None),OSError('PRIVATE_RESPONSE')],{'A':200,'B':None})):
            with patch('live_floorplate.connection_test.fetch_json_response',side_effect=responses) as fetch:
                self.assertEqual(compare_project_list_requests(CONFIG),expected)
                self.assertEqual(fetch.call_count,2)

    def test_ui_explicit_button_displays_only_two_statuses(self):
        from streamlit.testing.v1 import AppTest
        from pathlib import Path
        path=str(Path(__file__).resolve().parents[1]/'app.py')
        with patch('live_floorplate.ui.compare_project_list_requests',return_value={'A':403,'B':200}) as compare, \
                patch('live_floorplate.ui.connection_report') as report,patch('live_floorplate.ui.sync') as sync:
            app=AppTest.from_file(path,default_timeout=20)
            app.secrets['arcanite']=CONFIG
            app.run();app.sidebar.radio[0].set_value('🏢 Live Floorplate').run()
            compare.assert_not_called()
            next(button for button in app.button if button.label=='Compare Project List Requests').click().run()
            compare.assert_called_once_with(CONFIG)
            report.assert_not_called();sync.assert_not_called()
            self.assertEqual(len(app.exception),0)
            text='\n'.join(item.value for item in app.markdown)
            self.assertIn('Test A — Required parameters only — HTTP status: 403',text)
            self.assertIn('Test B — Official page 4 example: all optional parameters empty — HTTP status: 200',text)
            self.assertNotIn(CONFIG['sign'],text)
            self.assertEqual(len(app.dataframe),0)
            self.assertNotIn('live_floorplate_snapshot',app.session_state)
            self.assertNotIn('Configured Agent Portal API',app.radio[0].options)

    def test_ui_missing_secrets_does_not_request(self):
        from streamlit.testing.v1 import AppTest
        from pathlib import Path
        with patch('live_floorplate.connection_test.fetch_json_response') as fetch:
            app=AppTest.from_file(str(Path(__file__).resolve().parents[1]/'app.py'),default_timeout=20).run()
            app.sidebar.radio[0].set_value('🏢 Live Floorplate').run()
            next(button for button in app.button if button.label=='Compare Project List Requests').click().run()
            self.assertEqual(len(app.exception),0)
            self.assertTrue(any('HTTP status: unavailable' in item.value for item in app.markdown))
            fetch.assert_not_called()
