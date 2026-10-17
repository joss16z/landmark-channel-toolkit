"""Exercise existing navigation and new sample page without any API calls."""
import unittest
from pathlib import Path
from unittest.mock import patch
from streamlit.testing.v1 import AppTest

APP = str(Path(__file__).resolve().parents[1]/'app.py')


class StreamlitTests(unittest.TestCase):
    def test_existing_pages_and_sample_floorplate_work(self):
        app=AppTest.from_file(APP,default_timeout=20).run()
        self.assertEqual(len(app.exception),0)
        self.assertIn('Floorplate Update Report',app.header[0].value)
        app.sidebar.radio[0].set_value('✅ PPT价格更新').run()
        self.assertEqual(len(app.exception),0)
        self.assertIn('PPT',app.header[0].value)
        app.sidebar.radio[0].set_value('🏢 Live Floorplate').run()
        self.assertEqual(len(app.exception),0)
        self.assertIn('Live Floorplate',app.header[0].value)
        self.assertEqual(app.radio[0].value,'Source PDF preview (synthetic inventory)')
        self.assertNotIn('Configured Agent Portal API',app.radio[0].options)
        app.radio[0].set_value('Synthetic demonstration').run()
        self.assertTrue(app.session_state['live_floorplate_snapshot'].sample)
        groups=app.selectbox[0]
        self.assertEqual(groups.options,['Ground','Level 1','Levels 2–4','Levels 5–7','Level 8'])
        for group in range(5):
            app.selectbox[0].set_value(group).run()
            self.assertEqual(len(app.exception),0)

    def test_internal_preview_never_calls_api_without_assets(self):
        with patch('live_floorplate.ui.sync') as sync:
            app=AppTest.from_file(APP,default_timeout=20).run()
            app.sidebar.radio[0].set_value('🏢 Live Floorplate').run()
            self.assertEqual(app.radio[0].value,'Source PDF preview (synthetic inventory)')
            self.assertEqual(len(app.exception),0)
            sync.assert_not_called()
            self.assertNotIn('live_floorplate_snapshot',app.session_state)

    def test_preview_server_guard_clears_snapshot_and_blocks_api(self):
        from live_floorplate.ui import synchronize_into,KEY
        from live_floorplate.sample import demonstration
        from live_floorplate.model import FloorplateError
        data,layout,_=demonstration();state={KEY:data}
        with patch.dict('os.environ',{'LANDMARK_PREVIEW_ONLY':'1'}),patch('live_floorplate.ui.sync') as sync:
            with self.assertRaises(FloorplateError):synchronize_into(state,{},layout)
            sync.assert_not_called()
            self.assertNotIn(KEY,state)

    def test_preview_drops_previous_live_mode_selection(self):
        with patch.dict('os.environ',{'LANDMARK_PREVIEW_ONLY':'1'}),patch('live_floorplate.ui.sync') as sync:
            app=AppTest.from_file(APP,default_timeout=20).run()
            app.session_state['live_floorplate_mode']='Configured Agent Portal API'
            app.sidebar.radio[0].set_value('🏢 Live Floorplate').run()
            self.assertEqual(len(app.exception),0)
            self.assertEqual(app.radio[0].value,'Source PDF preview (synthetic inventory)')
            self.assertNotIn('Configured Agent Portal API',app.radio[0].options)
            sync.assert_not_called()
