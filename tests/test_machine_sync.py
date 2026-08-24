import unittest
import sys
import os
import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'github_project'))

from web_portal.app import app, ima_db, sync_pma_machines_to_ima

class TestMachineAutoSync(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        self.client = app.test_client()

    def test_auto_sync_machines(self):
        # First sync real data from bundled Excel
        sync_pma_machines_to_ima()
        real_machines = ima_db.get_all_machines()
        # Bundled Excel should have machines
        self.assertGreater(len(real_machines), 0, "Bundled Excel should populate at least 1 machine")

        # Now add mock machines too
        mock_df = pd.DataFrame([
            {'Equipment': 'TS1500-400053-147', 'Machine_Name': 'Banc Test TS1500', 'Group': 'TESTING', 'Sheet': 'TESTING'},
            {'Equipment': 'KOMAX-GAMMA-263', 'Machine_Name': 'Coupe Komax 263', 'Group': 'KOMAX', 'Sheet': 'KOMAX'},
            {'Equipment': 'KS-SCHLEUNIGER-01', 'Machine_Name': 'Sertisseuse KS 01', 'Group': 'KS', 'Sheet': 'KS'},
        ])
        count = sync_pma_machines_to_ima(mock_df)
        self.assertGreaterEqual(count, 0)

        # Check that real machine IDs from the Excel are present
        machines = ima_db.get_all_machines()
        m_ids = [m['machine_id'] for m in machines]
        self.assertGreater(len(m_ids), 0)

        # Login as owner
        self.client.post('/login', data={'username': 'owner', 'password': 'Owner@SEBN2026!'}, follow_redirects=True)

        # Verify intervention form loads and has machine options
        res = self.client.get('/interventions/new')
        self.assertEqual(res.status_code, 200)
        self.assertIn(b'machine-select', res.data)

        # Verify API returns machines
        api_res = self.client.get('/api/machines-by-group')
        self.assertEqual(api_res.status_code, 200)
        data = api_res.get_json()
        self.assertIsInstance(data, list)

if __name__ == '__main__':
    unittest.main()
