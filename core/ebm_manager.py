import os
import re
import json
import sqlite3
import logging
import datetime
from typing import Dict, Any, List, Optional, Tuple
import pandas as pd
import numpy as np

logger = logging.getLogger('sebn-maintenance')

def clean_currency(val) -> float:
    """Parses diverse currency and numeric formats into clean float values."""
    if val is None or (isinstance(val, float) and np.isnan(val)) or val == '':
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).lower().strip()
    s = re.sub(r'[\s\xa0\u202f\u2007\u200b]+', '', s)
    s = re.sub(r'[a-z€$£\u00a4]+', '', s)
    if not s:
        return 0.0
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        if re.search(r',[0-9]{3}$', s):
            s = s.replace(',', '')
        else:
            s = s.replace(',', '.')
    try:
        return float(re.sub(r'[^-0-9.]', '', s))
    except Exception:
        return 0.0

class EBMManager:
    """
    EBM (Equipment Budget Management / Suivi EBM) Engine.
    Handles Excel parsing, KPIs, validations, receptions, and action plans.
    """
    def __init__(self, db_path: str, data_dir: str):
        self.db_path = db_path
        self.data_dir = data_dir
        self.uploads_dir = os.path.join(data_dir, "EBM_uploads")
        os.makedirs(self.uploads_dir, exist_ok=True)
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
        with self._get_conn() as conn:
            cur = conn.cursor()
            # EBM Settings table
            cur.execute("""
                CREATE TABLE IF NOT EXISTS ebm_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
            """)
            # EBM Action Plans table
            cur.execute("""
                CREATE TABLE IF NOT EXISTS ebm_action_plans (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    title TEXT NOT NULL,
                    description TEXT,
                    responsible TEXT,
                    due_date TEXT,
                    priority TEXT DEFAULT 'Medium',
                    status TEXT DEFAULT 'Pending',
                    file_path TEXT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)

            cur.execute("PRAGMA table_info(ebm_action_plans)")
            ap_cols = {row["name"].lower() for row in cur.fetchall()}
            if "title" not in ap_cols:
                cur.execute("ALTER TABLE ebm_action_plans ADD COLUMN title TEXT DEFAULT ''")
                if "titre" in ap_cols:
                    cur.execute("UPDATE ebm_action_plans SET title = titre WHERE title = '' OR title IS NULL")
            if "responsible" not in ap_cols:
                cur.execute("ALTER TABLE ebm_action_plans ADD COLUMN responsible TEXT DEFAULT ''")
                if "responsable" in ap_cols:
                    cur.execute("UPDATE ebm_action_plans SET responsible = responsable WHERE responsible = '' OR responsible IS NULL")
            if "due_date" not in ap_cols:
                cur.execute("ALTER TABLE ebm_action_plans ADD COLUMN due_date TEXT DEFAULT ''")
                if "date" in ap_cols:
                    cur.execute("UPDATE ebm_action_plans SET due_date = date WHERE due_date = '' OR due_date IS NULL")
            if "priority" not in ap_cols:
                cur.execute("ALTER TABLE ebm_action_plans ADD COLUMN priority TEXT DEFAULT 'Medium'")
            if "status" not in ap_cols:
                cur.execute("ALTER TABLE ebm_action_plans ADD COLUMN status TEXT DEFAULT 'Pending'")
                if "statut" in ap_cols:
                    cur.execute("UPDATE ebm_action_plans SET status = statut WHERE status = '' OR status IS NULL")
            if "file_path" not in ap_cols:
                cur.execute("ALTER TABLE ebm_action_plans ADD COLUMN file_path TEXT DEFAULT ''")
            if "updated_at" not in ap_cols:
                cur.execute("ALTER TABLE ebm_action_plans ADD COLUMN updated_at DATETIME")
            # EBM Item Validation overrides
            cur.execute("""
                CREATE TABLE IF NOT EXISTS ebm_validations (
                    item_id TEXT PRIMARY KEY,
                    status TEXT NOT NULL,
                    validated_by TEXT,
                    comment TEXT,
                    validated_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()

    def get_setting(self, key: str, default: str = "") -> str:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT value FROM ebm_settings WHERE key = ?", (key,))
            row = cur.fetchone()
            return row["value"] if row else default

    def set_setting(self, key: str, value: str):
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("INSERT OR REPLACE INTO ebm_settings (key, value) VALUES (?, ?)", (key, str(value)))
            conn.commit()

    def get_active_excel_path(self) -> Optional[str]:
        """Returns the active EBM Excel file path."""
        p = self.get_setting("dashboard_ebm")
        if p and os.path.isfile(p):
            return p
        # Check standard file inside uploads_dir
        for f in os.listdir(self.uploads_dir):
            if f.lower().endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~$'):
                return os.path.join(self.uploads_dir, f)
        # Check fallback test file
        test_p = os.path.join(self.data_dir, "TEST_Dhashbord_EBM.xlsx")
        if os.path.isfile(test_p):
            return test_p
        return None

    def read_ebm_data(self) -> Tuple[Optional[pd.DataFrame], str, Dict[str, Any]]:
        """
        Reads active EBM Excel file, detects columns, and applies manual validation overrides.
        Returns (DataFrame, sheet_name, detected_columns_dict).
        """
        path = self.get_active_excel_path()
        if not path or not os.path.isfile(path):
            return None, "", {}

        try:
            xl = pd.ExcelFile(path)
            best_df = None
            best_sheet = ""
            max_score = -1

            header_keywords = [
                'status', 'statut', 'validation', 'valid', 'project', 'projet',
                'budget', 'ebm', 'price', 'prix', 'montant', 'ecsc', 'po',
                'description', 'equipment', 'date', 'delivered', 'reception'
            ]

            for sname in xl.sheet_names:
                df_scan = pd.read_excel(path, sheet_name=sname, header=None, nrows=25)
                sheet_score = 0
                best_row = 0
                for i, row in df_scan.iterrows():
                    non_null = row.notna().sum()
                    if non_null < 2:
                        continue
                    row_vals = [str(v).lower().strip() for v in row if pd.notnull(v)]
                    score = sum(1 for v in row_vals for k in header_keywords if k in v)
                    if any('status' in v or 'statut' in v or 'valid' in v for v in row_vals):
                        score += 10
                    if any('ebm' in v for v in row_vals):
                        score += 5
                    if score > sheet_score:
                        sheet_score = score
                        best_row = i
                if sheet_score > max_score:
                    max_score = sheet_score
                    best_df = pd.read_excel(path, sheet_name=sname, header=best_row)
                    best_sheet = sname

            if best_df is None:
                best_df = pd.read_excel(path)
                best_sheet = xl.sheet_names[0] if xl.sheet_names else "Sheet1"

            # Clean column names
            best_df.columns = [re.sub(r'\s+', ' ', str(c).strip()) for c in best_df.columns]
            best_df = best_df.dropna(how='all')

            # Column detection
            cols_map = {}
            for c in best_df.columns:
                cl = c.lower()
                if not cols_map.get('status') and any(k in cl for k in ['status', 'statut', 'statue', 'validation', 'état']):
                    cols_map['status'] = c
                elif not cols_map.get('project') and any(k in cl for k in ['project', 'projet', 'ebm', 'ecsc']):
                    cols_map['project'] = c
                elif not cols_map.get('price') and any(k in cl for k in ['price', 'prix', 'montant', 'cout', 'cost']):
                    cols_map['price'] = c
                elif not cols_map.get('po') and any(k in cl for k in ['po', 'commande', 'order', 'n° po']):
                    cols_map['po'] = c
                elif not cols_map.get('resp') and any(k in cl for k in ['resp', 'responsable', 'technicien', 'charge']):
                    cols_map['resp'] = c
                elif not cols_map.get('date') and any(k in cl for k in ['date', 'delai', 'delivery', 'livraison']):
                    cols_map['date'] = c
                elif not cols_map.get('desc') and any(k in cl for k in ['desc', 'libelle', 'designation', 'equip']):
                    cols_map['desc'] = c

            return best_df, best_sheet, cols_map
        except Exception as e:
            logger.error(f"[EBM] Read error: {e}")
            return None, "", {}

    def get_kpis(self) -> Dict[str, Any]:
        """Calculates comprehensive EBM KPIs and metrics."""
        df, sheet_name, cols = self.read_ebm_data()
        if df is None or df.empty:
            return {
                "file_loaded": False,
                "file_name": "",
                "total_items": 0,
                "total_budget": 0.0,
                "validated_count": 0,
                "validated_amount": 0.0,
                "pending_count": 0,
                "pending_amount": 0.0,
                "rejected_count": 0,
                "rejected_amount": 0.0,
                "delivered_count": 0,
                "delivered_amount": 0.0,
                "projects_breakdown": [],
                "status_breakdown": [],
                "recent_items": []
            }

    def get_validations(self) -> Dict[str, Dict[str, Any]]:
        """Retrieves manual validation overrides from the database."""
        try:
            with self._get_conn() as conn:
                cur = conn.cursor()
                cur.execute("SELECT item_id, status, validated_by, comment, validated_at FROM ebm_validations")
                return {str(r["item_id"]): dict(r) for r in cur.fetchall()}
        except Exception as e:
            logger.error(f"[EBM] Error fetching validations: {e}")
            return {}

    def set_item_validation(self, item_id: str, status: str, validated_by: str = "Owner", comment: str = "") -> bool:
        """Sets or updates manual validation override for an EBM item."""
        try:
            with self._get_conn() as conn:
                cur = conn.cursor()
                cur.execute("""
                    INSERT OR REPLACE INTO ebm_validations (item_id, status, validated_by, comment, validated_at)
                    VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
                """, (str(item_id), status, validated_by, comment))
                conn.commit()
                return True
        except Exception as e:
            logger.error(f"[EBM] Error saving validation override for item {item_id}: {e}")
            return False

    def remove_item_validation(self, item_id: str) -> bool:
        """Removes manual validation override, returning item to Excel-calculated status."""
        try:
            with self._get_conn() as conn:
                cur = conn.cursor()
                cur.execute("DELETE FROM ebm_validations WHERE item_id = ?", (str(item_id),))
                conn.commit()
                return True
        except Exception as e:
            logger.error(f"[EBM] Error removing validation override for item {item_id}: {e}")
            return False

    def get_kpis(self) -> Dict[str, Any]:
        """Calculates comprehensive EBM KPIs and metrics with accurate status categorization and validation overrides."""
        df, sheet_name, cols = self.read_ebm_data()
        if df is None or df.empty:
            return {
                "file_loaded": False,
                "file_name": "",
                "total_items": 0,
                "total_budget": 0.0,
                "validated_count": 0,
                "validated_amount": 0.0,
                "committed_count": 0,
                "committed_amount": 0.0,
                "pending_count": 0,
                "pending_amount": 0.0,
                "rejected_count": 0,
                "rejected_amount": 0.0,
                "delivered_count": 0,
                "delivered_amount": 0.0,
                "projects_breakdown": [],
                "status_breakdown": [],
                "items_list": []
            }

        price_col = cols.get('price')
        status_col = cols.get('status')
        proj_col = cols.get('project')
        po_col = cols.get('po')
        date_col = cols.get('date')
        desc_col = cols.get('desc')
        resp_col = cols.get('resp')

        # Detect columns for Comment and Delivery Date if present
        comment_col = None
        delivery_col = None
        for c in df.columns:
            cl = str(c).lower().strip()
            if 'comment' in cl:
                comment_col = c
            elif 'delevry' in cl or 'delivery' in cl or 'livraison' in cl:
                delivery_col = c

        total_items = len(df)
        total_budget = 0.0
        if price_col and price_col in df.columns:
            total_budget = df[price_col].apply(clean_currency).sum()

        validations = self.get_validations()

        validated_count = 0
        validated_amount = 0.0
        pending_count = 0
        pending_amount = 0.0
        rejected_count = 0
        rejected_amount = 0.0
        delivered_count = 0
        delivered_amount = 0.0

        status_counts = {"Validé": 0, "Livré": 0, "En cours": 0, "Refusé": 0}
        project_data = {}
        items = []

        for idx, row in df.iterrows():
            item_id = str(idx + 1)
            pr = clean_currency(row.get(price_col, 0)) if price_col else 0.0
            proj = str(row.get(proj_col, "Autre")).strip() if proj_col else "Autre"
            if not proj or proj.lower() == 'nan':
                proj = "Non Spécifié"

            st_raw = str(row.get(status_col, "")).strip() if status_col else ""
            if st_raw.lower() == 'nan':
                st_raw = ""

            cm_raw = str(row.get(comment_col, "")).strip() if comment_col else ""
            if cm_raw.lower() == 'nan':
                cm_raw = ""

            po_raw = str(row.get(po_col, "")).strip() if po_col else ""
            if po_raw.lower() == 'nan':
                po_raw = ""

            deliv_raw = str(row.get(delivery_col, "")).strip() if delivery_col else ""
            if deliv_raw.lower() == 'nan':
                deliv_raw = ""

            is_override = False
            validated_by = ""

            # Check for manual override from database
            if item_id in validations:
                cat = validations[item_id].get("status", "En cours")
                is_override = True
                validated_by = validations[item_id].get("validated_by", "Owner")
            else:
                # Combined text for keyword matching
                combo = f"{st_raw.lower()} {cm_raw.lower()} {deliv_raw.lower()}"

                # 1. Refusé / Rejeté / Canceled
                if any(k in combo for k in ['rejet', 'refus', 'reject', 'cancel', 'annul', 'nok']):
                    cat = "Refusé"
                # 2. Livré (Delivered) - must not be negative ("not yet delivred", "waiting", etc.)
                elif not any(k in combo for k in ['not yet', 'not deliv', 'pas encore', 'non livr', 'waiting', 'attente']) and \
                     (any(k in combo for k in ['delivred', 'delivered', 'livré', 'livre', 'reçu', 'received', 'delivery in complited']) or st_raw.lower() == 'done'):
                    cat = "Livré"
                # 3. Validé / Engagé (PO released, document accepted, or valid PO number exists)
                elif any(k in combo for k in ['valid', 'approuv', 'conforme', 'document accepted', 'po released']) or \
                     (po_raw and po_raw not in ['0', 'nan', 'NaN', 'None', '', '0.0']):
                    cat = "Validé"
                # 4. En attente / En cours
                else:
                    cat = "En cours"

            # Tally metrics
            if cat == "Validé":
                validated_count += 1
                validated_amount += pr
            elif cat == "Livré":
                delivered_count += 1
                delivered_amount += pr
            elif cat == "Refusé":
                rejected_count += 1
                rejected_amount += pr
            else:
                cat = "En cours"
                pending_count += 1
                pending_amount += pr

            status_counts[cat] = status_counts.get(cat, 0) + 1

            if proj not in project_data:
                project_data[proj] = {"count": 0, "amount": 0.0}
            project_data[proj]["count"] += 1
            project_data[proj]["amount"] += pr

            items.append({
                "id": idx + 1,
                "project": proj,
                "description": str(row.get(desc_col, proj) if desc_col else proj),
                "price": pr,
                "status": cat,
                "raw_status": st_raw or cm_raw or "–",
                "po": po_raw,
                "responsible": str(row.get(resp_col, '') if resp_col else ''),
                "date": str(row.get(date_col, '') if date_col else ''),
                "comment": cm_raw,
                "delivery_date": deliv_raw,
                "is_override": is_override,
                "validated_by": validated_by
            })

        projects_breakdown = [
            {"project": k, "count": v["count"], "amount": round(v["amount"], 2)}
            for k, v in sorted(project_data.items(), key=lambda x: x[1]["amount"], reverse=True)
        ]

        committed_count = validated_count + delivered_count
        committed_amount = round(validated_amount + delivered_amount, 2)

        path = self.get_active_excel_path()
        return {
            "file_loaded": True,
            "file_name": os.path.basename(path) if path else "",
            "total_items": total_items,
            "total_budget": round(total_budget, 2),
            "validated_count": validated_count,
            "validated_amount": round(validated_amount, 2),
            "committed_count": committed_count,
            "committed_amount": committed_amount,
            "pending_count": pending_count,
            "pending_amount": round(pending_amount, 2),
            "rejected_count": rejected_count,
            "rejected_amount": round(rejected_amount, 2),
            "delivered_count": delivered_count,
            "delivered_amount": round(delivered_amount, 2),
            "projects_breakdown": projects_breakdown[:10],
            "status_breakdown": [{"status": k, "count": v} for k, v in status_counts.items()],
            "items_list": items
        }

    # ── Action Plans ─────────────────────────────────────────────────────────

    def get_action_plans(self) -> List[Dict[str, Any]]:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM ebm_action_plans ORDER BY created_at DESC")
            return [dict(r) for r in cur.fetchall()]

    def add_action_plan(self, title: str, description: str, responsible: str, due_date: str, priority: str = 'Medium', file_path: str = "") -> int:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(ebm_action_plans)")
            cols = {row["name"].lower() for row in cur.fetchall()}
            if "titre" in cols and "responsable" in cols and "statut" in cols:
                cur.execute("""
                    INSERT INTO ebm_action_plans (
                        titre, description, responsable, date, statut,
                        title, responsible, due_date, priority, status, file_path, created_at, updated_at
                    )
                    VALUES (?, ?, ?, ?, 'À faire', ?, ?, ?, ?, 'Pending', ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                """, (
                    title.strip(), description.strip(), responsible.strip(), due_date.strip(),
                    title.strip(), responsible.strip(), due_date.strip(), priority, file_path
                ))
            else:
                cur.execute("""
                    INSERT INTO ebm_action_plans (title, description, responsible, due_date, priority, status, file_path, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, 'Pending', ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                """, (title.strip(), description.strip(), responsible.strip(), due_date.strip(), priority, file_path))
            conn.commit()
            return cur.lastrowid

    def toggle_action_plan(self, plan_id: int) -> Tuple[bool, str]:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT status FROM ebm_action_plans WHERE id = ?", (plan_id,))
            row = cur.fetchone()
            if not row:
                return False, "Plan introuvable."
            new_st = "Completed" if row["status"] == "Pending" else "Pending"
            cur.execute("UPDATE ebm_action_plans SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_st, plan_id))
            conn.commit()
            return True, new_st

    def delete_action_plan(self, plan_id: int) -> bool:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM ebm_action_plans WHERE id = ?", (plan_id,))
            conn.commit()
            return True
