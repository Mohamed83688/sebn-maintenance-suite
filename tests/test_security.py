"""
Security & Path Traversal Protection Tests
"""
import unittest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from web_portal.app import app
from core.document_manager import DocumentManager
from core.config import ConfigManager

class TestSecurity(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        self.client = app.test_client()
        self.pma_config = ConfigManager()

    def test_path_traversal_blocked(self):
        doc_mgr = DocumentManager(
            db_path=os.path.join(self.pma_config.active_base, "IMA.db"),
            data_dir=self.pma_config.active_base
        )
        # Attempt traversal path check
        is_safe = doc_mgr._is_safe_path("C:\\Windows\\System32\\calc.exe")
        self.assertFalse(is_safe)

    def test_sanitized_filename(self):
        dirty_name = "../../../etc/passwd"
        safe = DocumentManager.safe_filename(dirty_name)
        self.assertNotIn("/", safe)
        self.assertNotIn("\\", safe)
        self.assertNotIn("..", safe)

if __name__ == '__main__':
    unittest.main()
