"""
attendance_manager.py
=====================
Technician Attendance & Absence Management Engine for SEBN-TN Maintenance System.
Handles:
- Daily & Weekend Attendance (Pointage semaine & week-end samedi/dimanche)
- Leaves & Permissions (Gestion des congés, permissions, maladies)
- Missing Badge Incidents (Manque de badgage / Oubli de pointage entrée/sortie)
- Analytics & Statistics Dashboard (Taux de présence, ratios mensuels/annuels, filtres)
"""

import sqlite3
import logging
import datetime
from typing import Dict, Any, List, Optional, Tuple

logger = logging.getLogger("sebn-maintenance")

FRENCH_DAYS = {
    0: "Lundi",
    1: "Mardi",
    2: "Mercredi",
    3: "Jeudi",
    4: "Vendredi",
    5: "Samedi",
    6: "Dimanche"
}

FRENCH_MONTHS = {
    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril",
    5: "Mai", 6: "Juin", 7: "Juillet", 8: "Août",
    9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
}


class AttendanceManager:
    """
    Manages technician attendance, leaves, permissions, and missing badge logs.
    Stores data permanently in the primary SQLite database (IMA.db).
    """

    def __init__(self, db_path: str, data_dir: Optional[str] = None):
        self.db_path = db_path
        self.data_dir = data_dir
        self._init_db()

    def _get_conn(self):
        conn = sqlite3.connect(self.db_path, timeout=15)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        """Creates tables for attendance, leaves, and missing badge tracking."""
        with self._get_conn() as conn:
            cur = conn.cursor()

            # 1. Technician Attendance Table (Daily & Weekend)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS technician_attendance (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    technician_id INTEGER NOT NULL,
                    matricule TEXT NOT NULL,
                    technician_name TEXT NOT NULL,
                    shift TEXT DEFAULT 'B',
                    date TEXT NOT NULL,
                    day_of_week TEXT NOT NULL,
                    is_weekend INTEGER DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'PRESENT',
                    arrival_time TEXT DEFAULT '06:00',
                    departure_time TEXT DEFAULT '14:00',
                    hours_worked REAL DEFAULT 8.0,
                    notes TEXT,
                    created_by TEXT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(technician_id, date)
                )
            """)

            # 2. Leaves & Permissions Table
            cur.execute("""
                CREATE TABLE IF NOT EXISTS technician_leaves (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    technician_id INTEGER NOT NULL,
                    matricule TEXT NOT NULL,
                    technician_name TEXT NOT NULL,
                    leave_type TEXT NOT NULL DEFAULT 'CONGE',
                    start_date TEXT NOT NULL,
                    end_date TEXT NOT NULL,
                    duration_days REAL DEFAULT 1.0,
                    start_time TEXT,
                    end_time TEXT,
                    reason TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'APPROUVE',
                    approved_by TEXT,
                    notes TEXT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # 3. Missing Badge Tracking Table (Manque de badgage)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS technician_missing_badges (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    technician_id INTEGER NOT NULL,
                    matricule TEXT NOT NULL,
                    technician_name TEXT NOT NULL,
                    date TEXT NOT NULL,
                    badge_type TEXT NOT NULL DEFAULT 'ENTREE',
                    real_time TEXT,
                    shift TEXT,
                    reason TEXT NOT NULL,
                    is_regularized INTEGER DEFAULT 0,
                    regularized_by TEXT,
                    regularized_at DATETIME,
                    notes TEXT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)

            conn.commit()
            logger.info("[Attendance] Initialized attendance, leaves, and missing_badges tables.")

    # ── Technician Helpers ───────────────────────────────────────────────────

    def get_all_technicians(self) -> List[Dict[str, Any]]:
        """Returns list of active technicians from the users table."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, name, matricule, shift, role, technician_level, is_active
                FROM users
                WHERE role = 'TECHNICIAN' OR role IS NULL
                ORDER BY name ASC
            """)
            rows = [dict(r) for r in cur.fetchall()]
            if not rows:
                # Fallback to any user with matricule
                cur.execute("SELECT id, name, matricule, shift, role, technician_level, is_active FROM users ORDER BY name ASC")
                rows = [dict(r) for r in cur.fetchall()]
            return rows

    def get_technician_by_id(self, technician_id: int) -> Optional[Dict[str, Any]]:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, name, matricule, shift, role, technician_level FROM users WHERE id = ?", (technician_id,))
            r = cur.fetchone()
            return dict(r) if r else None

    # ── Attendance Methods ───────────────────────────────────────────────────

    @staticmethod
    def _compute_day_info(date_str: str) -> Tuple[str, int]:
        """Returns (day_of_week_fr, is_weekend). Saturday=5, Sunday=6."""
        try:
            dt = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            weekday = dt.weekday()
            day_fr = FRENCH_DAYS.get(weekday, "")
            is_wknd = 1 if weekday in (5, 6) else 0
            return day_fr, is_wknd
        except Exception:
            return "", 0

    @staticmethod
    def _compute_hours(arrival: str, departure: str) -> float:
        """Computes approximate working hours between arrival and departure."""
        try:
            t1 = datetime.datetime.strptime(arrival.strip(), "%H:%M")
            t2 = datetime.datetime.strptime(departure.strip(), "%H:%M")
            diff = (t2 - t1).total_seconds() / 3600.0
            if diff < 0:
                diff += 24.0  # Overnight shift
            return round(diff, 2)
        except Exception:
            return 8.0

    def record_attendance(
        self,
        technician_id: int,
        date: str,
        status: str = "PRESENT",
        arrival_time: str = "06:00",
        departure_time: str = "14:00",
        hours_worked: Optional[float] = None,
        notes: str = "",
        created_by: str = ""
    ) -> Dict[str, Any]:
        """
        Records or updates attendance for a technician on a specific date (Upsert).
        Auto-detects weekend and computes hours if not specified.
        """
        tech = self.get_technician_by_id(technician_id)
        if not tech:
            raise ValueError(f"Technicien avec ID {technician_id} introuvable.")

        day_fr, is_wknd = self._compute_day_info(date)
        if hours_worked is None:
            if status in ("PRESENT", "RETARD", "MISSION"):
                hours_worked = self._compute_hours(arrival_time, departure_time)
            else:
                hours_worked = 0.0

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO technician_attendance (
                    technician_id, matricule, technician_name, shift, date,
                    day_of_week, is_weekend, status, arrival_time, departure_time,
                    hours_worked, notes, created_by, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(technician_id, date) DO UPDATE SET
                    status = excluded.status,
                    arrival_time = excluded.arrival_time,
                    departure_time = excluded.departure_time,
                    hours_worked = excluded.hours_worked,
                    notes = excluded.notes,
                    updated_at = CURRENT_TIMESTAMP
            """, (
                tech['id'], tech['matricule'], tech['name'], tech.get('shift', 'B'),
                date, day_fr, is_wknd, status, arrival_time, departure_time,
                hours_worked, notes, created_by
            ))
            conn.commit()
            rec_id = cur.lastrowid

        return {"success": True, "id": rec_id, "technician": tech['name'], "date": date, "is_weekend": is_wknd}

    def get_attendance_records(
        self,
        technician_id: Optional[int] = None,
        month: Optional[int] = None,
        year: Optional[int] = None,
        status: Optional[str] = None,
        weekend_only: bool = False,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        limit: int = 300
    ) -> List[Dict[str, Any]]:
        """Queries attendance records with flexible filtering."""
        query = "SELECT * FROM technician_attendance WHERE 1=1"
        params: List[Any] = []

        if technician_id:
            query += " AND technician_id = ?"
            params.append(technician_id)

        if weekend_only:
            query += " AND is_weekend = 1"

        if status:
            query += " AND status = ?"
            params.append(status)

        if start_date:
            query += " AND date >= ?"
            params.append(start_date)

        if end_date:
            query += " AND date <= ?"
            params.append(end_date)

        if year and not (start_date or end_date):
            if month:
                prefix = f"{year:04d}-{month:02d}"
                query += " AND date LIKE ?"
                params.append(f"{prefix}%")
            else:
                prefix = f"{year:04d}"
                query += " AND date LIKE ?"
                params.append(f"{prefix}%")

        query += " ORDER BY date DESC, technician_name ASC LIMIT ?"
        params.append(limit)

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            return [dict(r) for r in cur.fetchall()]

    def get_attendance_by_id(self, record_id: int) -> Optional[Dict[str, Any]]:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM technician_attendance WHERE id = ?", (record_id,))
            r = cur.fetchone()
            return dict(r) if r else None

    def update_attendance(
        self,
        record_id: int,
        status: str,
        arrival_time: str,
        departure_time: str,
        hours_worked: Optional[float] = None,
        notes: str = ""
    ) -> bool:
        """Edits an existing attendance record."""
        if hours_worked is None:
            if status in ("PRESENT", "RETARD", "MISSION"):
                hours_worked = self._compute_hours(arrival_time, departure_time)
            else:
                hours_worked = 0.0

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE technician_attendance
                SET status = ?, arrival_time = ?, departure_time = ?,
                    hours_worked = ?, notes = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (status, arrival_time, departure_time, hours_worked, notes, record_id))
            conn.commit()
            return cur.rowcount > 0

    def delete_attendance(self, record_id: int) -> bool:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM technician_attendance WHERE id = ?", (record_id,))
            conn.commit()
            return cur.rowcount > 0

    # ── Leaves & Permissions Methods ─────────────────────────────────────────

    def record_leave(
        self,
        technician_id: int,
        leave_type: str,
        start_date: str,
        end_date: str,
        duration_days: float = 1.0,
        reason: str = "",
        start_time: Optional[str] = None,
        end_time: Optional[str] = None,
        status: str = "APPROUVE",
        approved_by: str = "",
        notes: str = ""
    ) -> Dict[str, Any]:
        """Registers a leave or authorized absence/permission."""
        tech = self.get_technician_by_id(technician_id)
        if not tech:
            raise ValueError(f"Technicien avec ID {technician_id} introuvable.")

        # If end_date is before start_date, clamp
        if not end_date:
            end_date = start_date

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO technician_leaves (
                    technician_id, matricule, technician_name, leave_type,
                    start_date, end_date, duration_days, start_time, end_time,
                    reason, status, approved_by, notes, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            """, (
                tech['id'], tech['matricule'], tech['name'], leave_type,
                start_date, end_date, duration_days, start_time, end_time,
                reason, status, approved_by, notes
            ))
            conn.commit()
            leave_id = cur.lastrowid

            # Also create corresponding attendance entries marked CONGE/PERMISSION for the dates
            try:
                dt_start = datetime.datetime.strptime(start_date, "%Y-%m-%d").date()
                dt_end = datetime.datetime.strptime(end_date, "%Y-%m-%d").date()
                curr = dt_start
                while curr <= dt_end:
                    curr_str = curr.strftime("%Y-%m-%d")
                    day_fr, is_wknd = self._compute_day_info(curr_str)
                    att_status = "CONGE" if leave_type == "CONGE" else "PERMISSION"
                    cur.execute("""
                        INSERT INTO technician_attendance (
                            technician_id, matricule, technician_name, shift, date,
                            day_of_week, is_weekend, status, arrival_time, departure_time,
                            hours_worked, notes, created_by
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, '', '', 0, ?, ?)
                        ON CONFLICT(technician_id, date) DO UPDATE SET
                            status = excluded.status,
                            notes = excluded.notes,
                            updated_at = CURRENT_TIMESTAMP
                    """, (
                        tech['id'], tech['matricule'], tech['name'], tech.get('shift', 'B'),
                        curr_str, day_fr, is_wknd, att_status,
                        f"{leave_type}: {reason}", approved_by
                    ))
                    curr += datetime.timedelta(days=1)
                conn.commit()
            except Exception as e:
                logger.warning(f"[Attendance] Could not sync leave to daily attendance: {e}")

        return {"success": True, "id": leave_id, "technician": tech['name'], "duration_days": duration_days}

    def get_leaves(
        self,
        technician_id: Optional[int] = None,
        leave_type: Optional[str] = None,
        year: Optional[int] = None,
        month: Optional[int] = None,
        limit: int = 200
    ) -> List[Dict[str, Any]]:
        """Retrieves leave and permission records."""
        query = "SELECT * FROM technician_leaves WHERE 1=1"
        params: List[Any] = []

        if technician_id:
            query += " AND technician_id = ?"
            params.append(technician_id)

        if leave_type:
            query += " AND leave_type = ?"
            params.append(leave_type)

        if year:
            if month:
                prefix = f"{year:04d}-{month:02d}"
                query += " AND (start_date LIKE ? OR end_date LIKE ?)"
                params.extend([f"{prefix}%", f"{prefix}%"])
            else:
                prefix = f"{year:04d}"
                query += " AND (start_date LIKE ? OR end_date LIKE ?)"
                params.extend([f"{prefix}%", f"{prefix}%"])

        query += " ORDER BY start_date DESC, id DESC LIMIT ?"
        params.append(limit)

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            return [dict(r) for r in cur.fetchall()]

    def get_leave_by_id(self, leave_id: int) -> Optional[Dict[str, Any]]:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM technician_leaves WHERE id = ?", (leave_id,))
            r = cur.fetchone()
            return dict(r) if r else None

    def update_leave(
        self,
        leave_id: int,
        leave_type: str,
        start_date: str,
        end_date: str,
        duration_days: float,
        reason: str,
        status: str,
        notes: str = ""
    ) -> bool:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE technician_leaves
                SET leave_type = ?, start_date = ?, end_date = ?,
                    duration_days = ?, reason = ?, status = ?, notes = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (leave_type, start_date, end_date, duration_days, reason, status, notes, leave_id))
            conn.commit()
            return cur.rowcount > 0

    def delete_leave(self, leave_id: int) -> bool:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM technician_leaves WHERE id = ?", (leave_id,))
            conn.commit()
            return cur.rowcount > 0

    # ── Missing Badge Incidents (Manque de Badgage) ───────────────────────────

    def record_missing_badge(
        self,
        technician_id: int,
        date: str,
        badge_type: str = "ENTREE",
        real_time: str = "",
        shift: Optional[str] = None,
        reason: str = "",
        is_regularized: int = 0,
        regularized_by: Optional[str] = None,
        notes: str = ""
    ) -> Dict[str, Any]:
        """Records a missing badge incident (forgot to badge entry or exit)."""
        tech = self.get_technician_by_id(technician_id)
        if not tech:
            raise ValueError(f"Technicien avec ID {technician_id} introuvable.")

        actual_shift = shift or tech.get('shift', 'B')

        with self._get_conn() as conn:
            cur = conn.cursor()
            reg_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") if is_regularized else None
            cur.execute("""
                INSERT INTO technician_missing_badges (
                    technician_id, matricule, technician_name, date,
                    badge_type, real_time, shift, reason,
                    is_regularized, regularized_by, regularized_at, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                tech['id'], tech['matricule'], tech['name'], date,
                badge_type, real_time, actual_shift, reason,
                is_regularized, regularized_by, reg_at, notes
            ))
            conn.commit()
            badge_id = cur.lastrowid

        return {"success": True, "id": badge_id, "technician": tech['name'], "badge_type": badge_type}

    def get_missing_badges(
        self,
        technician_id: Optional[int] = None,
        is_regularized: Optional[int] = None,
        year: Optional[int] = None,
        month: Optional[int] = None,
        limit: int = 200
    ) -> List[Dict[str, Any]]:
        """Retrieves missing badge incidents with filtering."""
        query = "SELECT * FROM technician_missing_badges WHERE 1=1"
        params: List[Any] = []

        if technician_id:
            query += " AND technician_id = ?"
            params.append(technician_id)

        if is_regularized is not None:
            query += " AND is_regularized = ?"
            params.append(int(is_regularized))

        if year:
            if month:
                prefix = f"{year:04d}-{month:02d}"
                query += " AND date LIKE ?"
                params.append(f"{prefix}%")
            else:
                prefix = f"{year:04d}"
                query += " AND date LIKE ?"
                params.append(f"{prefix}%")

        query += " ORDER BY date DESC, id DESC LIMIT ?"
        params.append(limit)

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            return [dict(r) for r in cur.fetchall()]

    def regularize_missing_badge(
        self,
        badge_id: int,
        regularized_by: str,
        notes: str = ""
    ) -> bool:
        """Marks a missing badge incident as regularized/corrected."""
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE technician_missing_badges
                SET is_regularized = 1, regularized_by = ?, regularized_at = ?,
                    notes = CASE WHEN ? != '' THEN notes || ' | ' || ? ELSE notes END
                WHERE id = ?
            """, (regularized_by, now_str, notes, notes, badge_id))
            conn.commit()
            return cur.rowcount > 0

    def delete_missing_badge(self, badge_id: int) -> bool:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM technician_missing_badges WHERE id = ?", (badge_id,))
            conn.commit()
            return cur.rowcount > 0

    # ── Statistics & Dashboard Analytics ──────────────────────────────────────

    def get_statistics(
        self,
        technician_id: Optional[int] = None,
        month: Optional[int] = None,
        year: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        Computes high-level and granular statistics for the attendance dashboard.
        Includes weekend tracking, leave days, permissions, missing badges,
        attendance rates, and per-technician summary cards.
        """
        # Default to current year if none provided
        if not year:
            year = datetime.date.today().year

        techs = self.get_all_technicians()

        # Date filter SQL snippet
        date_cond = ""
        params: List[Any] = []
        if month:
            prefix = f"{year:04d}-{month:02d}"
            date_cond = " AND date LIKE ?"
            params.append(f"{prefix}%")
        else:
            prefix = f"{year:04d}"
            date_cond = " AND date LIKE ?"
            params.append(f"{prefix}%")

        tech_cond = ""
        if technician_id:
            tech_cond = " AND technician_id = ?"
            params.append(technician_id)

        with self._get_conn() as conn:
            cur = conn.cursor()

            # 1. Total Attendance counts
            cur.execute(f"""
                SELECT 
                    COUNT(*) as total_records,
                    SUM(CASE WHEN status = 'PRESENT' OR status = 'RETARD' OR status = 'MISSION' THEN 1 ELSE 0 END) as total_present,
                    SUM(CASE WHEN is_weekend = 1 AND (status = 'PRESENT' OR status = 'RETARD') THEN 1 ELSE 0 END) as weekend_present,
                    SUM(CASE WHEN status = 'ABSENT' THEN 1 ELSE 0 END) as total_absent,
                    SUM(CASE WHEN status = 'CONGE' THEN 1 ELSE 0 END) as total_conge_records,
                    SUM(CASE WHEN status = 'PERMISSION' THEN 1 ELSE 0 END) as total_perm_records,
                    SUM(hours_worked) as total_hours
                FROM technician_attendance
                WHERE 1=1 {date_cond} {tech_cond}
            """, params)
            att_row = dict(cur.fetchone())

            # 2. Total Leaves (from technician_leaves table)
            leave_params: List[Any] = []
            leave_cond = ""
            if month:
                leave_cond += " AND (start_date LIKE ? OR end_date LIKE ?)"
                leave_params.extend([f"{year:04d}-{month:02d}%", f"{year:04d}-{month:02d}%"])
            else:
                leave_cond += " AND (start_date LIKE ? OR end_date LIKE ?)"
                leave_params.extend([f"{year:04d}%", f"{year:04d}%"])

            if technician_id:
                leave_cond += " AND technician_id = ?"
                leave_params.append(technician_id)

            cur.execute(f"""
                SELECT 
                    SUM(CASE WHEN leave_type = 'CONGE' THEN duration_days ELSE 0 END) as total_conge_days,
                    SUM(CASE WHEN leave_type = 'PERMISSION' THEN 1 ELSE 0 END) as total_permissions_count,
                    SUM(CASE WHEN leave_type = 'MALADIE' THEN duration_days ELSE 0 END) as total_maladie_days,
                    COUNT(*) as total_leave_records
                FROM technician_leaves
                WHERE status = 'APPROUVE' {leave_cond}
            """, leave_params)
            leave_row = dict(cur.fetchone())

            # 3. Missing badges
            badge_params: List[Any] = []
            badge_cond = ""
            if month:
                badge_cond += " AND date LIKE ?"
                badge_params.append(f"{year:04d}-{month:02d}%")
            else:
                badge_cond += " AND date LIKE ?"
                badge_params.append(f"{year:04d}%")

            if technician_id:
                badge_cond += " AND technician_id = ?"
                badge_params.append(technician_id)

            cur.execute(f"""
                SELECT 
                    COUNT(*) as total_missing_badges,
                    SUM(CASE WHEN is_regularized = 0 THEN 1 ELSE 0 END) as unregularized_badges,
                    SUM(CASE WHEN is_regularized = 1 THEN 1 ELSE 0 END) as regularized_badges,
                    SUM(CASE WHEN badge_type = 'ENTREE' THEN 1 ELSE 0 END) as missing_entry,
                    SUM(CASE WHEN badge_type = 'SORTIE' THEN 1 ELSE 0 END) as missing_exit,
                    SUM(CASE WHEN badge_type = 'LES_DEUX' THEN 1 ELSE 0 END) as missing_both
                FROM technician_missing_badges
                WHERE 1=1 {badge_cond}
            """, badge_params)
            badge_row = dict(cur.fetchone())

            # 4. Per-Technician Summary Table
            tech_summary = []
            for t in techs:
                tid = t['id']
                tparams: List[Any] = []
                t_date_cond = ""
                if month:
                    t_date_cond = " AND date LIKE ?"
                    tparams.append(f"{year:04d}-{month:02d}%")
                else:
                    t_date_cond = " AND date LIKE ?"
                    tparams.append(f"{year:04d}%")
                tparams.append(tid)

                cur.execute(f"""
                    SELECT 
                        SUM(CASE WHEN status IN ('PRESENT', 'RETARD', 'MISSION') THEN 1 ELSE 0 END) as present_days,
                        SUM(CASE WHEN is_weekend = 1 AND status IN ('PRESENT', 'RETARD') THEN 1 ELSE 0 END) as weekend_days,
                        SUM(CASE WHEN status = 'ABSENT' THEN 1 ELSE 0 END) as absent_days,
                        SUM(CASE WHEN status = 'CONGE' THEN 1 ELSE 0 END) as conge_days,
                        SUM(CASE WHEN status = 'PERMISSION' THEN 1 ELSE 0 END) as permission_days,
                        SUM(hours_worked) as hours
                    FROM technician_attendance
                    WHERE 1=1 {t_date_cond} AND technician_id = ?
                """, tparams)
                t_att = dict(cur.fetchone())

                # Missing badges for this technician
                b_tparams: List[Any] = []
                b_tdate = ""
                if month:
                    b_tdate = " AND date LIKE ?"
                    b_tparams.append(f"{year:04d}-{month:02d}%")
                else:
                    b_tdate = " AND date LIKE ?"
                    b_tparams.append(f"{year:04d}%")
                b_tparams.append(tid)

                cur.execute(f"""
                    SELECT COUNT(*) as b_total, SUM(CASE WHEN is_regularized = 0 THEN 1 ELSE 0 END) as b_unreg
                    FROM technician_missing_badges
                    WHERE 1=1 {b_tdate} AND technician_id = ?
                """, b_tparams)
                t_bad = dict(cur.fetchone())

                # Leaves for this technician
                l_tparams: List[Any] = []
                l_tdate = ""
                if month:
                    l_tdate = " AND (start_date LIKE ? OR end_date LIKE ?)"
                    l_tparams.extend([f"{year:04d}-{month:02d}%", f"{year:04d}-{month:02d}%"])
                else:
                    l_tdate = " AND (start_date LIKE ? OR end_date LIKE ?)"
                    l_tparams.extend([f"{year:04d}%", f"{year:04d}%"])
                l_tparams.append(tid)

                cur.execute(f"""
                    SELECT 
                        SUM(CASE WHEN leave_type = 'CONGE' THEN duration_days ELSE 0 END) as l_conge,
                        SUM(CASE WHEN leave_type = 'PERMISSION' THEN 1 ELSE 0 END) as l_perm
                    FROM technician_leaves
                    WHERE status = 'APPROUVE' {l_tdate} AND technician_id = ?
                """, l_tparams)
                t_l = dict(cur.fetchone())

                p_days = t_att['present_days'] or 0
                ab_days = t_att['absent_days'] or 0
                total_workable = p_days + ab_days
                rate = round((p_days / total_workable * 100), 1) if total_workable > 0 else 100.0

                tech_summary.append({
                    "id": tid,
                    "name": t['name'],
                    "matricule": t['matricule'],
                    "shift": t.get('shift', 'B'),
                    "level": t.get('technician_level', 'Level 1'),
                    "present_days": p_days,
                    "weekend_days": t_att['weekend_days'] or 0,
                    "absent_days": ab_days,
                    "conge_days": (t_l['l_conge'] or 0) or (t_att['conge_days'] or 0),
                    "permission_count": (t_l['l_perm'] or 0) or (t_att['permission_days'] or 0),
                    "total_hours": round(t_att['hours'] or 0, 1),
                    "missing_badges": t_bad['b_total'] or 0,
                    "unregularized_badges": t_bad['b_unreg'] or 0,
                    "attendance_rate": rate
                })

            # 5. Monthly Trend Breakdown (Jan-Dec for given year)
            monthly_trend = []
            for m in range(1, 13):
                m_str = f"{year:04d}-{m:02d}%"
                cur.execute("""
                    SELECT 
                        SUM(CASE WHEN status IN ('PRESENT', 'RETARD', 'MISSION') THEN 1 ELSE 0 END) as present,
                        SUM(CASE WHEN is_weekend = 1 AND status IN ('PRESENT', 'RETARD') THEN 1 ELSE 0 END) as weekend,
                        SUM(CASE WHEN status = 'ABSENT' THEN 1 ELSE 0 END) as absent
                    FROM technician_attendance
                    WHERE date LIKE ?
                """, (m_str,))
                mr = dict(cur.fetchone())

                cur.execute("SELECT COUNT(*) FROM technician_missing_badges WHERE date LIKE ?", (m_str,))
                mb_count = cur.fetchone()[0]

                cur.execute("SELECT SUM(duration_days) FROM technician_leaves WHERE status='APPROUVE' AND (start_date LIKE ? OR end_date LIKE ?)", (m_str, m_str))
                l_count = cur.fetchone()[0] or 0

                monthly_trend.append({
                    "month": m,
                    "month_name": FRENCH_MONTHS.get(m, ""),
                    "present": mr['present'] or 0,
                    "weekend": mr['weekend'] or 0,
                    "absent": mr['absent'] or 0,
                    "leaves": round(float(l_count), 1),
                    "missing_badges": mb_count
                })

        # Calculate high level attendance rate
        total_p = att_row['total_present'] or 0
        total_a = att_row['total_absent'] or 0
        total_work = total_p + total_a
        global_rate = round((total_p / total_work * 100), 1) if total_work > 0 else 100.0

        conge_days = leave_row['total_conge_days'] or att_row['total_conge_records'] or 0
        permission_count = leave_row['total_permissions_count'] or att_row['total_perm_records'] or 0

        return {
            "year": year,
            "month": month,
            "month_name": FRENCH_MONTHS.get(month, "") if month else "Toute l'année",
            "total_present_days": total_p,
            "total_weekend_days": att_row['weekend_present'] or 0,
            "total_absent_days": total_a,
            "total_conge_days": round(float(conge_days), 1),
            "total_permissions": int(permission_count),
            "total_missing_badges": badge_row['total_missing_badges'] or 0,
            "unregularized_badges": badge_row['unregularized_badges'] or 0,
            "total_hours": round(float(att_row['total_hours'] or 0), 1),
            "global_attendance_rate": global_rate,
            "technicians_summary": tech_summary,
            "monthly_trend": monthly_trend,
            "technicians_count": len(techs)
        }

    # ── Excel Export Generator ───────────────────────────────────────────────

    def export_to_excel(
        self,
        year: Optional[int] = None,
        month: Optional[int] = None,
        technician_id: Optional[int] = None
    ) -> Any:
        """
        Generates a comprehensive multi-sheet Excel workbook using openpyxl.
        Returns a BytesIO stream containing the .xlsx binary data.
        """
        import io
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        if not year:
            year = datetime.date.today().year

        stats = self.get_statistics(technician_id=technician_id, month=month, year=year)
        records = self.get_attendance_records(technician_id=technician_id, month=month, year=year, limit=2000)
        leaves = self.get_leaves(technician_id=technician_id, year=year, month=month, limit=1000)
        badges = self.get_missing_badges(technician_id=technician_id, year=year, month=month, limit=1000)

        wb = openpyxl.Workbook()

        # Styles
        hdr_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        hdr_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        title_font = Font(name="Segoe UI", size=15, bold=True, color="1E3A8A")
        sub_font = Font(name="Segoe UI", size=10, italic=True, color="475569")
        bold_font = Font(name="Segoe UI", size=10, bold=True)
        regular_font = Font(name="Segoe UI", size=10)
        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )
        card_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        wknd_fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")

        # ── SHEET 1: KPIs & Synthèse ─────────────────────────────────────────
        ws1 = wb.active
        ws1.title = "Synthèse & KPIs"
        ws1.views.sheetView[0].showGridLines = True

        ws1.cell(row=2, column=2, value="SEBN-TN Maintenance — Suivi de Présence & Absences").font = title_font
        period_str = f"{stats['month_name']} {stats['year']}" if stats['month'] else f"Année {stats['year']}"
        ws1.cell(row=3, column=2, value=f"Période: {period_str} | Date d'export: {datetime.date.today().isoformat()}").font = sub_font

        # KPI Cards row
        kpi_headers = [
            ("Jours Présents", stats['total_present_days']),
            ("Jours Week-end", stats['total_weekend_days']),
            ("Heures Travaillées", f"{stats['total_hours']}h"),
            ("Jours Congés", f"{stats['total_conge_days']}j"),
            ("Permissions", stats['total_permissions']),
            ("Manques Badgage", stats['total_missing_badges']),
            ("Taux Présence", f"{stats['global_attendance_rate']}%")
        ]
        for idx, (label, val) in enumerate(kpi_headers):
            col = 2 + idx
            c_val = ws1.cell(row=5, column=col, value=val)
            c_val.font = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")
            c_val.alignment = Alignment(horizontal='center', vertical='center')
            c_val.fill = card_fill
            c_val.border = thin_border

            c_lbl = ws1.cell(row=6, column=col, value=label)
            c_lbl.font = Font(name="Segoe UI", size=9, bold=True, color="64748B")
            c_lbl.alignment = Alignment(horizontal='center', vertical='center')
            c_lbl.fill = card_fill
            c_lbl.border = thin_border

        # Per-Technician Summary Table
        ws1.cell(row=8, column=2, value="Récapitulatif par Technicien").font = Font(name="Segoe UI", size=12, bold=True, color="1E3A8A")
        tech_cols = ["ID", "Matricule", "Nom & Prénom", "Équipe", "Niveau", "Présents", "Week-ends", "Absents", "Congés (j)", "Permissions", "Heures", "Manques Badge", "Taux Présence"]
        for c_idx, c_name in enumerate(tech_cols, start=2):
            cell = ws1.cell(row=9, column=c_idx, value=c_name)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

        row_idx = 10
        for ts in stats['technicians_summary']:
            ws1.cell(row=row_idx, column=2, value=ts['id']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=3, value=ts['matricule']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=4, value=ts['name']).alignment = Alignment(horizontal='left')
            ws1.cell(row=row_idx, column=5, value=ts['shift']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=6, value=ts['level']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=7, value=ts['present_days']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=8, value=ts['weekend_days']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=9, value=ts['absent_days']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=10, value=ts['conge_days']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=11, value=ts['permission_count']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=12, value=ts['total_hours']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=13, value=ts['missing_badges']).alignment = Alignment(horizontal='center')
            ws1.cell(row=row_idx, column=14, value=f"{ts['attendance_rate']}%").alignment = Alignment(horizontal='center')

            for c_idx in range(2, 15):
                c = ws1.cell(row=row_idx, column=c_idx)
                c.font = regular_font
                c.border = thin_border
            row_idx += 1

        # ── SHEET 2: Suivi Pointages & Présences ──────────────────────────────
        ws2 = wb.create_sheet(title="Pointages & Présences")
        ws2.views.sheetView[0].showGridLines = True
        att_cols = ["ID", "Date", "Jour", "Week-end ?", "Matricule", "Technicien", "Équipe", "Statut", "Arrivée", "Départ", "Heures", "Remarques", "Saisi Par"]
        for c_idx, c_name in enumerate(att_cols, start=1):
            cell = ws2.cell(row=1, column=c_idx, value=c_name)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        r_idx = 2
        for r in records:
            ws2.cell(row=r_idx, column=1, value=r['id']).alignment = Alignment(horizontal='center')
            ws2.cell(row=r_idx, column=2, value=r['date']).alignment = Alignment(horizontal='center')
            ws2.cell(row=r_idx, column=3, value=r['day_of_week']).alignment = Alignment(horizontal='center')
            is_wk = "OUI (Samedi/Dimanche)" if r['is_weekend'] else "Non"
            c_wk = ws2.cell(row=r_idx, column=4, value=is_wk)
            c_wk.alignment = Alignment(horizontal='center')
            if r['is_weekend']:
                c_wk.fill = wknd_fill
                c_wk.font = bold_font

            ws2.cell(row=r_idx, column=5, value=r['matricule']).alignment = Alignment(horizontal='center')
            ws2.cell(row=r_idx, column=6, value=r['technician_name']).alignment = Alignment(horizontal='left')
            ws2.cell(row=r_idx, column=7, value=r['shift']).alignment = Alignment(horizontal='center')
            ws2.cell(row=r_idx, column=8, value=r['status']).alignment = Alignment(horizontal='center')
            ws2.cell(row=r_idx, column=9, value=r['arrival_time'] or "—").alignment = Alignment(horizontal='center')
            ws2.cell(row=r_idx, column=10, value=r['departure_time'] or "—").alignment = Alignment(horizontal='center')
            ws2.cell(row=r_idx, column=11, value=r['hours_worked']).alignment = Alignment(horizontal='center')
            ws2.cell(row=r_idx, column=12, value=r['notes'] or "").alignment = Alignment(horizontal='left')
            ws2.cell(row=r_idx, column=13, value=r['created_by'] or "").alignment = Alignment(horizontal='center')

            for c_idx in range(1, 14):
                c = ws2.cell(row=r_idx, column=c_idx)
                if not (c_idx == 4 and r['is_weekend']):
                    c.font = regular_font
                c.border = thin_border
            r_idx += 1

        # ── SHEET 3: Congés & Permissions ─────────────────────────────────────
        ws3 = wb.create_sheet(title="Congés & Permissions")
        ws3.views.sheetView[0].showGridLines = True
        leave_cols = ["ID", "Matricule", "Technicien", "Type", "Date Début", "Date Fin", "Durée (Jours)", "Heures (si partiel)", "Motif / Raison", "Statut", "Approuvé Par", "Notes"]
        for c_idx, c_name in enumerate(leave_cols, start=1):
            cell = ws3.cell(row=1, column=c_idx, value=c_name)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        r_idx = 2
        for l in leaves:
            ws3.cell(row=r_idx, column=1, value=l['id']).alignment = Alignment(horizontal='center')
            ws3.cell(row=r_idx, column=2, value=l['matricule']).alignment = Alignment(horizontal='center')
            ws3.cell(row=r_idx, column=3, value=l['technician_name']).alignment = Alignment(horizontal='left')
            ws3.cell(row=r_idx, column=4, value=l['leave_type']).alignment = Alignment(horizontal='center')
            ws3.cell(row=r_idx, column=5, value=l['start_date']).alignment = Alignment(horizontal='center')
            ws3.cell(row=r_idx, column=6, value=l['end_date']).alignment = Alignment(horizontal='center')
            ws3.cell(row=r_idx, column=7, value=l['duration_days']).alignment = Alignment(horizontal='center')
            times = f"{l['start_time']} - {l['end_time']}" if l.get('start_time') else "—"
            ws3.cell(row=r_idx, column=8, value=times).alignment = Alignment(horizontal='center')
            ws3.cell(row=r_idx, column=9, value=l['reason']).alignment = Alignment(horizontal='left')
            ws3.cell(row=r_idx, column=10, value=l['status']).alignment = Alignment(horizontal='center')
            ws3.cell(row=r_idx, column=11, value=l['approved_by'] or "—").alignment = Alignment(horizontal='center')
            ws3.cell(row=r_idx, column=12, value=l['notes'] or "").alignment = Alignment(horizontal='left')

            for c_idx in range(1, 13):
                c = ws3.cell(row=r_idx, column=c_idx)
                c.font = regular_font
                c.border = thin_border
            r_idx += 1

        # ── SHEET 4: Manques de Badgage ───────────────────────────────────────
        ws4 = wb.create_sheet(title="Manques de Badgage")
        ws4.views.sheetView[0].showGridLines = True
        badge_cols = ["ID", "Date", "Matricule", "Technicien", "Équipe", "Type d'Oubli", "Heure Réelle", "Motif / Justification", "Régularisé ?", "Régularisé Par", "Date Régularisation", "Notes"]
        for c_idx, c_name in enumerate(badge_cols, start=1):
            cell = ws4.cell(row=1, column=c_idx, value=c_name)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        r_idx = 2
        for b in badges:
            ws4.cell(row=r_idx, column=1, value=b['id']).alignment = Alignment(horizontal='center')
            ws4.cell(row=r_idx, column=2, value=b['date']).alignment = Alignment(horizontal='center')
            ws4.cell(row=r_idx, column=3, value=b['matricule']).alignment = Alignment(horizontal='center')
            ws4.cell(row=r_idx, column=4, value=b['technician_name']).alignment = Alignment(horizontal='left')
            ws4.cell(row=r_idx, column=5, value=b['shift']).alignment = Alignment(horizontal='center')
            ws4.cell(row=r_idx, column=6, value=b['badge_type']).alignment = Alignment(horizontal='center')
            ws4.cell(row=r_idx, column=7, value=b['real_time'] or "—").alignment = Alignment(horizontal='center')
            ws4.cell(row=r_idx, column=8, value=b['reason']).alignment = Alignment(horizontal='left')
            is_reg = "OUI (Régularisé)" if b['is_regularized'] else "NON (En attente)"
            c_reg = ws4.cell(row=r_idx, column=9, value=is_reg)
            c_reg.alignment = Alignment(horizontal='center')
            c_reg.font = bold_font
            if not b['is_regularized']:
                c_reg.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            else:
                c_reg.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

            ws4.cell(row=r_idx, column=10, value=b['regularized_by'] or "—").alignment = Alignment(horizontal='center')
            ws4.cell(row=r_idx, column=11, value=b['regularized_at'] or "—").alignment = Alignment(horizontal='center')
            ws4.cell(row=r_idx, column=12, value=b['notes'] or "").alignment = Alignment(horizontal='left')

            for c_idx in range(1, 13):
                c = ws4.cell(row=r_idx, column=c_idx)
                c.border = thin_border
            r_idx += 1

        # Auto-adjust column widths for all sheets
        for sheet in wb.worksheets:
            for col in sheet.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        lines = str(cell.value).split("\n")
                        for l in lines:
                            if len(l) > max_len:
                                max_len = len(l)
                sheet.column_dimensions[col_letter].width = max(max_len + 3, 11)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output
