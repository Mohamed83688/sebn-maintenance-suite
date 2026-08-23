"""
Role-Based Access Control (RBAC) Tests
"""
import unittest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from web_portal.app import app

class TestRBAC(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        self.client = app.test_client()

    def test_unauthenticated_access_blocked(self):
        for endpoint in ['/dashboard', '/ebm', '/admin', '/documents', '/interventions', '/passation']:
            response = self.client.get(endpoint)
            self.assertIn(response.status_code, (302, 401, 403))

    def test_technician_cannot_access_owner_ebm(self):
        self.client.post('/login', data={'username': 'ahmed.bensalah', 'password': 'Tech@2026!'}, follow_redirects=True)
        response = self.client.get('/ebm')
        self.assertEqual(response.status_code, 403)

    def test_technician_cannot_access_admin_users(self):
        self.client.post('/login', data={'username': 'ahmed.bensalah', 'password': 'Tech@2026!'}, follow_redirects=True)
        response = self.client.get('/admin/users')
        self.assertEqual(response.status_code, 403)

    def test_owner_can_access_modules(self):
        self.client.post('/login', data={'username': 'owner', 'password': 'Owner@SEBN2026!'}, follow_redirects=True)
        for endpoint in ['/dashboard', '/ebm', '/admin', '/admin/users', '/documents', '/passation', '/pma', '/interventions']:
            response = self.client.get(endpoint)
            self.assertEqual(response.status_code, 200, f"Owner failed accessing {endpoint}")

if __name__ == '__main__':
    unittest.main()
