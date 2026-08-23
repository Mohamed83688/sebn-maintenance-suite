import os
import re
import sqlite3
import shutil
import logging
from typing import Optional, List, Dict, Any, Tuple

logger = logging.getLogger('sebn-maintenance')


class DocumentManager:
    """
    Manages the SEBN-TN document library:
    - Documents table in shared IMA.db (metadata only)
    - Actual files stored under  <data_dir>/documents/<doc_id>/<filename>
    - Files are NEVER stored inside the database binary column.
    - Soft-delete only: is_active=0 keeps files for historical references.
    """

    ALLOWED_EXTENSIONS = {
        'pdf', 'xlsx', 'xls', 'xlsm',
        'docx', 'doc', 'pptx', 'ppt',
        'png', 'jpg', 'jpeg', 'gif',
        'csv', 'txt'
    }

    def __init__(self, db_path: str, data_dir: str):
        self.db_path = db_path
        self.data_dir = data_dir
        self.storage_base = os.path.join(data_dir, 'documents')
        os.makedirs(self.storage_base, exist_ok=True)
        self._init_db()

    # ── Internal ──────────────────────────────────────────────────────────────

    def _get_conn(self):
        conn = sqlite3.connect(self.db_path, timeout=15)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        """Creates the documents table if it doesn't exist, and handles schema migrations."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS documents (
                    id              INTEGER PRIMARY KEY AUTOINCREMENT,
                    display_name    TEXT    NOT NULL,
                    file_name       TEXT    NOT NULL DEFAULT '',
                    file_type       TEXT    NOT NULL DEFAULT '',
                    storage_path    TEXT    NOT NULL DEFAULT '',
                    is_active       INTEGER NOT NULL DEFAULT 1,
                    display_order   INTEGER NOT NULL DEFAULT 0,
                    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)
            # Schema migration: add columns if upgrading from older version
            cur.execute("PRAGMA table_info(documents)")
            existing = {row['name'].lower() for row in cur.fetchall()}
            migrations = {
                'display_name':  "ALTER TABLE documents ADD COLUMN display_name TEXT NOT NULL DEFAULT ''",
                'file_name':     "ALTER TABLE documents ADD COLUMN file_name TEXT NOT NULL DEFAULT ''",
                'file_type':     "ALTER TABLE documents ADD COLUMN file_type TEXT NOT NULL DEFAULT ''",
                'storage_path':  "ALTER TABLE documents ADD COLUMN storage_path TEXT NOT NULL DEFAULT ''",
                'is_active':     "ALTER TABLE documents ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1",
                'display_order': "ALTER TABLE documents ADD COLUMN display_order INTEGER NOT NULL DEFAULT 0",
                'updated_at':    "ALTER TABLE documents ADD COLUMN updated_at DATETIME",
            }
            for col, sql in migrations.items():
                if col not in existing:
                    try:
                        cur.execute(sql)
                    except Exception as e:
                        logger.warning(f"[DocMgr] Migration skip {col}: {e}")

            cur.execute("CREATE INDEX IF NOT EXISTS idx_docs_active ON documents(is_active)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_docs_order  ON documents(display_order)")
            conn.commit()

    @staticmethod
    def detect_file_type(filename: str) -> str:
        """Returns the normalized file extension (lowercase, no dot)."""
        ext = os.path.splitext(filename)[1].lstrip('.').lower()
        return ext if ext else 'bin'

    @staticmethod
    def safe_filename(filename: str) -> str:
        """Sanitizes filename for safe filesystem storage."""
        name = os.path.basename(filename)
        # Replace dangerous characters
        name = re.sub(r'[^\w\-_\. ]', '_', name)
        name = name.strip()
        return name if name else 'document'

    def _abs_path(self, storage_path: str) -> str:
        """Converts relative storage_path (documents/15/foo.xlsx) to absolute path."""
        return os.path.normpath(os.path.join(self.data_dir, storage_path))

    def _is_safe_path(self, abs_path: str) -> bool:
        """Prevents path traversal: ensures abs_path is under storage_base."""
        return os.path.normpath(abs_path).startswith(os.path.normpath(self.storage_base))

    # ── Read ──────────────────────────────────────────────────────────────────

    def get_all_documents(self, active_only: bool = True) -> List[Dict[str, Any]]:
        """Returns all documents ordered by display_order, then id."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            if active_only:
                cur.execute("""
                    SELECT * FROM documents
                    WHERE is_active = 1
                    ORDER BY display_order ASC, id ASC
                """)
            else:
                cur.execute("""
                    SELECT * FROM documents
                    ORDER BY display_order ASC, id ASC
                """)
            rows = [dict(r) for r in cur.fetchall()]

        # Enrich with live file-system metadata
        for row in rows:
            abs_p = self._abs_path(row['storage_path']) if row['storage_path'] else ''
            row['file_exists'] = os.path.isfile(abs_p) if abs_p else False
            if row['file_exists']:
                row['file_size_kb'] = round(os.path.getsize(abs_p) / 1024, 1)
            else:
                row['file_size_kb'] = 0
        return rows

    def get_document(self, doc_id: int) -> Optional[Dict[str, Any]]:
        """Fetches a single document by ID (active or not)."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM documents WHERE id = ?", (doc_id,))
            row = cur.fetchone()
        if not row:
            return None
        doc = dict(row)
        abs_p = self._abs_path(doc['storage_path']) if doc['storage_path'] else ''
        doc['file_exists'] = os.path.isfile(abs_p) if abs_p else False
        doc['abs_path'] = abs_p
        if doc['file_exists']:
            doc['file_size_kb'] = round(os.path.getsize(abs_p) / 1024, 1)
        else:
            doc['file_size_kb'] = 0
        return doc

    def get_abs_path(self, doc_id: int) -> Optional[str]:
        """Returns the absolute filesystem path for doc_id, or None if not found/not exist."""
        doc = self.get_document(doc_id)
        if not doc or not doc['storage_path']:
            return None
        abs_p = self._abs_path(doc['storage_path'])
        if not self._is_safe_path(abs_p):
            logger.warning(f"[DocMgr] Path traversal attempt for doc_id={doc_id}")
            return None
        return abs_p if doc['file_exists'] else None

    # ── Write ─────────────────────────────────────────────────────────────────

    def add_document(
        self,
        display_name: str,
        file_name: str,
        file_type: str,
        storage_path: str,
        display_order: int = 0
    ) -> Tuple[bool, str, Optional[int]]:
        """Registers a new document. The file must already be saved to storage_path."""
        display_name = display_name.strip()
        if not display_name:
            return False, "Le nom d'affichage est obligatoire.", None
        if not storage_path:
            return False, "Le chemin de stockage est manquant.", None

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO documents
                    (display_name, file_name, file_type, storage_path, is_active, display_order, created_at, updated_at)
                VALUES (?, ?, ?, ?, 1, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
            """, (display_name, file_name, file_type, storage_path, display_order))
            conn.commit()
            return True, f"Document '{display_name}' ajouté avec succès.", cur.lastrowid

    def save_uploaded_file(self, file_storage, doc_id: int) -> Tuple[bool, str, str, str]:
        """
        Saves a Flask FileStorage object to  data/documents/<doc_id>/<safe_name>.
        Returns (success, msg, relative_storage_path, file_type).
        Call this AFTER inserting the DB row so you have doc_id.
        """
        if not file_storage or file_storage.filename == '':
            return False, "Aucun fichier sélectionné.", '', ''

        raw_name = file_storage.filename
        file_type = self.detect_file_type(raw_name)

        if file_type not in self.ALLOWED_EXTENSIONS:
            return False, f"Type de fichier non autorisé: .{file_type}", '', ''

        safe_name = self.safe_filename(raw_name)
        doc_dir   = os.path.join(self.storage_base, str(doc_id))
        os.makedirs(doc_dir, exist_ok=True)

        abs_dest = os.path.join(doc_dir, safe_name)
        try:
            file_storage.save(abs_dest)
        except Exception as e:
            logger.error(f"[DocMgr] File save error: {e}")
            return False, f"Erreur lors de la sauvegarde: {e}", '', ''

        rel_path = os.path.join('documents', str(doc_id), safe_name).replace('\\', '/')
        return True, "Fichier sauvegardé.", rel_path, file_type

    def replace_file(
        self,
        doc_id: int,
        file_storage
    ) -> Tuple[bool, str]:
        """
        Replaces the file for an existing document.
        Old file is archived as <filename>.old rather than deleted.
        The display_name (button label) does NOT change.
        """
        doc = self.get_document(doc_id)
        if not doc:
            return False, "Document introuvable."

        # Archive the old file safely
        if doc['storage_path']:
            old_abs = self._abs_path(doc['storage_path'])
            if os.path.isfile(old_abs) and self._is_safe_path(old_abs):
                try:
                    shutil.move(old_abs, old_abs + '.old')
                except Exception as e:
                    logger.warning(f"[DocMgr] Could not archive old file: {e}")

        ok, msg, rel_path, file_type = self.save_uploaded_file(file_storage, doc_id)
        if not ok:
            return False, msg

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE documents
                SET file_name    = ?,
                    file_type    = ?,
                    storage_path = ?,
                    updated_at   = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (
                self.safe_filename(file_storage.filename),
                file_type,
                rel_path,
                doc_id
            ))
            conn.commit()
        return True, f"Fichier remplacé pour '{doc['display_name']}'."

    def update_display_name(self, doc_id: int, new_name: str) -> Tuple[bool, str]:
        new_name = new_name.strip()
        if not new_name:
            return False, "Le nom ne peut pas être vide."
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                "UPDATE documents SET display_name=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (new_name, doc_id)
            )
            conn.commit()
        return True, "Nom mis à jour."

    def update_order(self, doc_id: int, new_order: int) -> Tuple[bool, str]:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                "UPDATE documents SET display_order=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (int(new_order), doc_id)
            )
            conn.commit()
        return True, "Ordre mis à jour."

    def toggle_active(self, doc_id: int) -> Tuple[bool, str, int]:
        """Toggles is_active. Returns (success, msg, new_is_active)."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT is_active, display_name FROM documents WHERE id=?", (doc_id,))
            row = cur.fetchone()
            if not row:
                return False, "Document introuvable.", 0
            new_state = 0 if row['is_active'] == 1 else 1
            cur.execute(
                "UPDATE documents SET is_active=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (new_state, doc_id)
            )
            conn.commit()
            label = "activé" if new_state == 1 else "désactivé"
            return True, f"Document '{row['display_name']}' {label}.", new_state

    def soft_delete(self, doc_id: int) -> Tuple[bool, str]:
        """
        Safe delete: sets is_active=0.
        The physical file is moved to <file>.archived (not deleted) to
        protect historical references.
        """
        doc = self.get_document(doc_id)
        if not doc:
            return False, "Document introuvable."

        # Archive physical file
        if doc['storage_path']:
            abs_p = self._abs_path(doc['storage_path'])
            if os.path.isfile(abs_p) and self._is_safe_path(abs_p):
                try:
                    shutil.move(abs_p, abs_p + '.archived')
                except Exception as e:
                    logger.warning(f"[DocMgr] Could not archive file on delete: {e}")

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                "UPDATE documents SET is_active=0, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (doc_id,)
            )
            conn.commit()
        return True, f"Document '{doc['display_name']}' supprimé (archivé)."
