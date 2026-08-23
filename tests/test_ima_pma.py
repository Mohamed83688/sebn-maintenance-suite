"""
IMA & PMA Module Integration Tests
"""
import unittest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from web_portal.app import app, ima_db, pma_config

class TestIMAPMA(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        self.client = app.test_client()

    def test_interventions_page(self):
        self.client.post('/login', data={'username': 'owner', 'password': 'Owner@SEBN2026!'}, follow_redirects=True)
        response = self.client.get('/interventions')
        self.assertEqual(response.status_code, 200)

    def test_machines_list(self):
        self.client.post('/login', data={'username': 'owner', 'password': 'Owner@SEBN2026!'}, follow_redirects=True)
        response = self.client.get('/machines')
        self.assertEqual(response.status_code, 200)

    def test_pma_dashboard(self):
        self.client.post('/login', data={'username': 'owner', 'password': 'Owner@SEBN2026!'}, follow_redirects=True)
        response = self.client.get('/pma')
        self.assertEqual(response.status_code, 200)

if __name__ == '__main__':
    unittest.main()
