"""
Database Initialization Script
Initializes a clean SQLite schema and default configuration for first-time installation.
"""
import os
import sys
import sqlite3

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.config import ConfigManager
from core.user_manager import UserManager
from core.document_manager import DocumentManager
from core.ebm_manager import EBMManager
from core.passation_manager import PassationManager
from ima.database import IMADatabase

def init_environment():
    pma_config = ConfigManager()
    data_dir = pma_config.active_base
    db_path = os.path.join(data_dir, "IMA.db")
    
    print(f"Initializing database at: {db_path}")
    
    # Initialize IMA tables
    ima_db = IMADatabase(db_path)
    
    # Initialize Users, Documents, EBM, Passation
    user_mgr = UserManager(db_path=db_path, data_dir=data_dir)
    doc_mgr = DocumentManager(db_path=db_path, data_dir=data_dir)
    ebm_mgr = EBMManager(db_path=db_path, data_dir=data_dir)
    passation_mgr = PassationManager(db_path=db_path, data_dir=data_dir)
    
    print("Database initialization complete.")

if __name__ == "__main__":
    init_environment()
