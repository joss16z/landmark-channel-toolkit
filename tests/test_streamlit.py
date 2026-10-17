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
        self.assertTrue(app.session_state['live_floorplate_snapshot'].sample)

    def test_configured_mode_does_not_call_api_without_assets(self):
        with patch('live_floorplate.ui.sync') as sync:
            app=AppTest.from_file(APP,default_timeout=20).run()
            app.sidebar.radio[0].set_value('🏢 Live Floorplate').run()
            app.radio[0].set_value('Configured Agent Portal API').run()
            self.assertEqual(len(app.exception),0)
            sync.assert_not_called()
            self.assertNotIn('live_floorplate_snapshot',app.session_state)
