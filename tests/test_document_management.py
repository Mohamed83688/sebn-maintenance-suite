"""
Document Management System Tests
"""
import unittest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from web_portal.app import app, doc_mgr

class TestDocumentManagement(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        self.client = app.test_client()

    def test_documents_page_authenticated(self):
        self.client.post('/login', data={'username': 'owner', 'password': 'Owner@SEBN2026!'}, follow_redirects=True)
        response = self.client.get('/documents')
        self.assertEqual(response.status_code, 200)

    def test_admin_documents_crud_page(self):
        self.client.post('/login', data={'username': 'owner', 'password': 'Owner@SEBN2026!'}, follow_redirects=True)
        response = self.client.get('/admin/documents')
        self.assertEqual(response.status_code, 200)

    def test_document_manager_methods(self):
        docs = doc_mgr.get_all_documents(active_only=False)
        self.assertIsInstance(docs, list)

if __name__ == '__main__':
    unittest.main()
