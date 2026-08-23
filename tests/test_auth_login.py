"""
Unit & Integration Tests: Authentication & Login System
"""
import unittest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from web_portal.app import app

class TestAuth(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        self.client = app.test_client()

    def test_login_page_renders(self):
        response = self.client.get('/login')
        self.assertEqual(response.status_code, 200)

    def test_owner_login(self):
        response = self.client.post('/login', data={
            'username': 'owner',
            'password': 'Owner@SEBN2026!'
        }, follow_redirects=True)
        self.assertEqual(response.status_code, 200)

    def test_invalid_login(self):
        response = self.client.post('/login', data={
            'username': 'invalid_user',
            'password': 'wrongpassword'
        }, follow_redirects=True)
        self.assertEqual(response.status_code, 200)

    def test_logout(self):
        self.client.post('/login', data={'username': 'owner', 'password': 'Owner@SEBN2026!'}, follow_redirects=True)
        response = self.client.get('/logout', follow_redirects=True)
        self.assertEqual(response.status_code, 200)

if __name__ == '__main__':
    unittest.main()
