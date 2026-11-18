"""
preventive_doc_manager.py
=========================
Manages manual preventive maintenance table scans, photos, PDF uploads,
and their permanent validation workflow in SEBN-TN Maintenance System.

Workflow:
1. Technician or Admin fills paper preventive table.
2. Takes photo / scans document (PDF or image).
3. Uploads file linked to machine, task, week, date, and technician.
4. Preview & verification.
5. Storage in permanent storage directory (preserved across updates & restarts).
6. Admin/Owner approval or rejection.
7. High-resolution zoomable viewer with inline PDF/image rendering.
"""

import os
import re
import time
import sqlite3
import hashlib
import logging
import mimetypes
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Any

logger = logging.getLogger("sebn-maintenance")

class PreventiveDocManager:
    def __init__(self, db_path: str, storage_dir: str):
        self.db_path = db_path
        self.storage_dir = storage_dir
        os.makedirs(self.storage_dir, exist_ok=True)
        self._init_db()

    def _get_conn(self):
        try:
            from core.db_mysql import get_db_connection
            return get_db_connection(sqlite_fallback_path=self.db_path)
        except Exception:
            conn = sqlite3.connect(self.db_path, timeout=15)
            conn.row_factory = sqlite3.Row
            return conn

    def _init_db(self):
        """Initializes the preventive_documents database table."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
            CREATE TABLE IF NOT EXISTS preventive_documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                equipment TEXT NOT NULL,
                machine_name TEXT,
                sheet TEXT NOT NULL,
                raw_index TEXT,
                task_type TEXT NOT NULL DEFAULT 'Monthly',
                week TEXT,
                month TEXT,
                task_key TEXT NOT NULL,
                original_filename TEXT NOT NULL,
                stored_filename TEXT NOT NULL,
                file_path TEXT NOT NULL,
                file_type TEXT NOT NULL,
                mime_type TEXT,
                file_size INTEGER DEFAULT 0,
                technician_name TEXT NOT NULL,
                technician_matricule TEXT,
                technician_id INTEGER,
                shift TEXT DEFAULT 'A',
                completion_date TEXT,
                comments TEXT,
                validation_status TEXT DEFAULT 'PENDING',
                validated_by TEXT,
                validated_at DATETIME,
                rejection_reason TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
            """)

            cur.execute("CREATE INDEX IF NOT EXISTS idx_prev_doc_task ON preventive_documents(task_key)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_prev_doc_equip ON preventive_documents(equipment)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_prev_doc_status ON preventive_documents(validation_status)")
            conn.commit()

    @staticmethod
    def build_task_key(sheet: str, raw_index: Any, equipment: str, week: str = "", task_type: str = "") -> str:
        """Constructs a deterministic unique key for a preventive task."""
        s = str(sheet or '').strip().lower().replace(" ", "_")
        idx = str(raw_index or '').strip()
        eq = str(equipment or '').strip().lower().replace(" ", "_")
        w = str(week or '').strip().lower().replace(" ", "_")
        tt = str(task_type or '').strip().lower().replace(" ", "_")
        return f"{s}__{idx}__{eq}__{w}__{tt}"

    def save_uploaded_document(
        self,
        file_obj,
        equipment: str,
        machine_name: str,
        sheet: str,
        raw_index: Any,
        task_type: str,
        week: str,
        month: str,
        technician_name: str,
        technician_matricule: str,
        technician_id: Optional[int],
        shift: str,
        completion_date: Optional[str] = None,
        comments: str = "",
        initial_status: str = "PENDING",
        validated_by: Optional[str] = None
    ) -> Tuple[bool, str, Optional[Dict[str, Any]]]:
        """
        Saves uploaded file to disk and registers record in database.
        Accepts PDF and images (JPG, JPEG, PNG, WEBP).
        """
        if not file_obj or not file_obj.filename:
            return False, "Aucun fichier fourni.", None

        orig_filename = os.path.basename(file_obj.filename)
        ext = os.path.splitext(orig_filename)[1].lower()

        allowed_images = {'.jpg', '.jpeg', '.png', '.webp'}
        allowed_docs = {'.pdf'}

        if ext in allowed_images:
            file_type = 'image'
        elif ext in allowed_docs:
            file_type = 'pdf'
        else:
            return False, f"Format non supporté ({ext}). Formats acceptés : PDF, JPG, JPEG, PNG.", None

        # Build clean unique filename
        safe_equip = re.sub(r'[^a-zA-Z0-9_\-]', '_', str(equipment or 'equip'))
        safe_week = re.sub(r'[^a-zA-Z0-9_\-]', '_', str(week or 'week'))
        timestamp = int(time.time())
        unique_token = hashlib.md5(f"{orig_filename}_{timestamp}".encode()).hexdigest()[:8]
        stored_filename = f"prev_table_{safe_equip}_{safe_week}_{timestamp}_{unique_token}{ext}"
        stored_path = os.path.join(self.storage_dir, stored_filename)

        try:
            file_obj.save(stored_path)
            file_size = os.path.getsize(stored_path)
        except Exception as e:
            logger.error(f"Error saving uploaded document: {e}")
            return False, f"Erreur lors de l'enregistrement du fichier : {e}", None

        mime_type, _ = mimetypes.guess_type(stored_path)
        if not mime_type:
            mime_type = "application/pdf" if file_type == 'pdf' else "image/jpeg"

        task_key = self.build_task_key(sheet, raw_index, equipment, week, task_type)
        comp_date = completion_date or datetime.today().strftime('%Y-%m-%d')
        val_status = initial_status.upper().strip()
        val_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S') if val_status == 'VALIDATED' else None

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
            INSERT INTO preventive_documents (
                equipment, machine_name, sheet, raw_index, task_type, week, month,
                task_key, original_filename, stored_filename, file_path, file_type,
                mime_type, file_size, technician_name, technician_matricule, technician_id,
                shift, completion_date, comments, validation_status, validated_by, validated_at,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
            """, (
                equipment,
                machine_name or equipment,
                sheet,
                str(raw_index),
                task_type or 'Monthly',
                week or '',
                month or '',
                task_key,
                orig_filename,
                stored_filename,
                stored_path,
                file_type,
                mime_type,
                file_size,
                technician_name,
                technician_matricule or '',
                technician_id,
                shift or 'A',
                comp_date,
                comments or '',
                val_status,
                validated_by if val_status == 'VALIDATED' else None,
                val_at
            ))
            doc_id = cur.lastrowid
            conn.commit()

        doc_dict = self.get_document_by_id(doc_id)
        return True, "Fichier enregistré avec succès.", doc_dict

    def get_document_by_id(self, doc_id: int) -> Optional[Dict[str, Any]]:
        """Fetches a single preventive document by its ID."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM preventive_documents WHERE id = ?", (doc_id,))
            row = cur.fetchone()
            if not row:
                return None
            d = dict(row)
            d['formatted_size'] = self.format_file_size(d.get('file_size', 0))
            return d

    def get_documents_for_task(self, sheet: str, raw_index: Any, equipment: str, week: str = "", task_type: str = "") -> List[Dict[str, Any]]:
        """Returns all documents attached to a specific task key."""
        task_key = self.build_task_key(sheet, raw_index, equipment, week, task_type)
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
            SELECT * FROM preventive_documents 
            WHERE task_key = ? OR (equipment = ? AND sheet = ? AND raw_index = ?)
            ORDER BY created_at DESC
            """, (task_key, equipment, sheet, str(raw_index)))
            rows = cur.fetchall()
            docs = []
            for r in rows:
                d = dict(r)
                d['formatted_size'] = self.format_file_size(d.get('file_size', 0))
                docs.append(d)
            return docs

    def get_all_documents_map(self) -> Dict[str, List[Dict[str, Any]]]:
        """
        Returns a dictionary mapping task_key -> list of documents.
        Also maps fallback key `(sheet, str(raw_index))` for quick O(1) lookup.
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM preventive_documents ORDER BY created_at DESC")
            rows = cur.fetchall()
            mapping: Dict[str, List[Dict[str, Any]]] = {}
            for r in rows:
                d = dict(r)
                d['formatted_size'] = self.format_file_size(d.get('file_size', 0))
                tk = d.get('task_key')
                if tk:
                    mapping.setdefault(tk, []).append(d)
                # Also index by fallback (equipment, week) and (sheet, raw_index)
                sec_key = f"{str(d.get('sheet', '')).strip().lower()}__{str(d.get('raw_index', '')).strip()}"
                mapping.setdefault(sec_key, []).append(d)
            return mapping

    def validate_document(self, doc_id: int, validated_by: str) -> Tuple[bool, str, Optional[Dict[str, Any]]]:
        """Approves and validates a submitted manual preventive document."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM preventive_documents WHERE id = ?", (doc_id,))
            row = cur.fetchone()
            if not row:
                return False, "Document introuvable.", None

            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            cur.execute("""
            UPDATE preventive_documents 
            SET validation_status = 'VALIDATED',
                validated_by = ?,
                validated_at = ?,
                rejection_reason = NULL,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """, (validated_by, now_str, doc_id))
            conn.commit()

        updated = self.get_document_by_id(doc_id)
        return True, "Tableau validé avec succès.", updated

    def reject_document(self, doc_id: int, rejected_by: str, reason: str = "") -> Tuple[bool, str, Optional[Dict[str, Any]]]:
        """Rejects a submitted manual preventive document."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM preventive_documents WHERE id = ?", (doc_id,))
            row = cur.fetchone()
            if not row:
                return False, "Document introuvable.", None

            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            cur.execute("""
            UPDATE preventive_documents 
            SET validation_status = 'REJECTED',
                validated_by = ?,
                validated_at = ?,
                rejection_reason = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """, (rejected_by, now_str, reason.strip() or "Document non conforme ou illisible", doc_id))
            conn.commit()

        updated = self.get_document_by_id(doc_id)
        return True, "Document rejeté.", updated

    def delete_document(self, doc_id: int) -> Tuple[bool, str]:
        """Deletes document record and removes physical file from disk."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM preventive_documents WHERE id = ?", (doc_id,))
            row = cur.fetchone()
            if not row:
                return False, "Document introuvable."

            file_path = row["file_path"]
            cur.execute("DELETE FROM preventive_documents WHERE id = ?", (doc_id,))
            conn.commit()

        if file_path and os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception as e:
                logger.warning(f"Failed to delete physical file {file_path}: {e}")

        return True, "Document supprimé avec succès."

    @staticmethod
    def format_file_size(size_bytes: int) -> str:
        """Formats file size into human-readable string (KB, MB)."""
        if not size_bytes or size_bytes < 1024:
            return f"{size_bytes or 0} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.1f} MB"
