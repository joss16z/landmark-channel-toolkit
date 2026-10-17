"""Offline structural review regressions using synthetic responses only."""
import json
import unittest
from live_floorplate.schema_review import inspect_response


class SchemaReviewTests(unittest.TestCase):
    def test_structure_without_private_values_or_invented_metadata(self):
        payload={'data':{'properties':[{'lotnumber':1,'unitnumber':'PRIVATE_UNIT_SENTINEL',
                 'price':'PRIVATE_PRICE_SENTINEL','propertystatus':'Available','area':{'internal':'PRIVATE_AREA_SENTINEL'}}]},
                 'sign':'PRIVATE_AUTH_SENTINEL','request':'https://private.invalid/?sign=PRIVATE_URL_SENTINEL'}
        report=inspect_response(payload)
        self.assertEqual(report['property_arrays'][0]['rows_path'],'data.properties')
        self.assertEqual(report['property_arrays'][0]['fields']['unitnumber'],['str'])
        self.assertEqual(report['property_arrays'][0]['fields']['area.internal'],['str'])
        self.assertFalse(report['endpoint_completeness_verified'])
        self.assertEqual(report['metadata_candidates'],[])
        self.assertNotIn('PRIVATE_',json.dumps(report))

    def test_all_lots_with_no_total_or_property_project_id(self):
        report=inspect_response({'properties':[dict(lotnumber=n,propertystatus='Available') for n in range(1,116)]})
        array=report['property_arrays'][0]
        self.assertTrue(array['exact_lots_1_to_115'])
        self.assertEqual(array['unique_lots'],115)
        self.assertEqual(report['metadata_candidates'],[])
        self.assertFalse(report['endpoint_completeness_verified'])

    def test_duplicate_partial_and_unknown_status_do_not_claim_complete(self):
        report=inspect_response([dict(lotnumber=1,propertystatus='PRIVATE_STATUS'),dict(lotnumber='001',propertystatus='Sold')])
        array=report['property_arrays'][0]
        self.assertFalse(array['exact_lots_1_to_115'])
        self.assertEqual(array['duplicate_lot_rows'],1)
        self.assertEqual(array['status_counts']['unknown'],1)
        self.assertNotIn('PRIVATE_STATUS',json.dumps(report))

    def test_possible_pagination_and_counts_report_types_not_values(self):
        report=inspect_response({'total':12345,'hasmore':True,'projectid':'PRIVATE_ID','nextpage':'PRIVATE_URL'})
        self.assertEqual(len(report['metadata_candidates']),4)
        self.assertNotIn('PRIVATE_',json.dumps(report))
        self.assertNotIn('12345',json.dumps(report))
