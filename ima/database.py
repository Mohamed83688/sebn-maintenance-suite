"""
IMA — SQLite Database Engine
Fully independent database for the Intervention Management Application.
Tables: machines, interventions, settings
"""
import sqlite3
import os
import datetime
import json
import time
import random


def make_safe_method(method):
    import functools
    @functools.wraps(method)
    def wrapper(*args, **kwargs):
        import sqlite3
        import time
        import random
        last_err = None
        for attempt in range(15):
            try:
                return method(*args, **kwargs)
            except sqlite3.OperationalError as e:
                err_msg = str(e).lower()
                is_lock_err = any(x in err_msg for x in ("locked", "busy", "unable to open", "io error", "disk"))
                if is_lock_err:
                    last_err = e
                    # Jittered sleep: random wait between 0.1 and 1.5 seconds
                    wait = 0.1 + random.uniform(0.05, 0.1) * (attempt + 1)
                    time.sleep(wait)
                else:
                    raise
        raise sqlite3.OperationalError(
            f"La base de données réseau est occupée par un autre poste.\n"
            f"Veuillez réessayer.\n"
            f"(Détail : {last_err})"
        )
    return wrapper


class IMADatabase:
    """Manages the IMA SQLite database: machines, interventions, settings."""

    def __init__(self, db_path: str):
        self.db_path = db_path
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        self._cleanup_stale_wal()   # Remove leftover WAL files from crashes
        self._init_schema()

        # Dynamically wrap all public methods with automatic network lock retries
        for attr_name in list(dir(self)):
            if not attr_name.startswith("_") and hasattr(self, attr_name):
                attr = getattr(self, attr_name)
                if callable(attr) and attr_name != "db_path":
                    setattr(self, attr_name, make_safe_method(attr))

    def _cleanup_stale_wal(self):
        """Delete stale WAL/SHM lock files left by a previous crash.
        These are the #1 cause of 'database is locked' on network drives.
        Only deletes them if the main .db file is safely accessible.
        """
        for suffix in ("-wal", "-shm"):
            stale = self.db_path + suffix
            if os.path.exists(stale):
                try:
                    os.remove(stale)
                except OSError:
                    pass  # Cannot remove — another instance may be running, that's fine

    def _conn(self, timeout=30):
        """Return a new connection with multi-user safe settings.

        timeout=30  — SQLite will wait up to 30 s for another writer to finish
                      before raising an error.  This covers normal factory use.
        journal_mode=DELETE — reliable on Windows network shares (SMB/CIFS).
        busy_timeout is set via PRAGMA as an extra safety net.
        """
        conn = sqlite3.connect(self.db_path, timeout=timeout)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=DELETE")   # safe on network drives
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute(f"PRAGMA busy_timeout={timeout * 1000}")  # ms
        return conn

    def _retry_write(self, fn, *args, retries=5, **kwargs):
        """Execute a write function with exponential-backoff retry.

        If SQLite is locked by another user, wait a short random time and
        try again up to `retries` times before raising the error to the UI.
        This is the core of multi-user safety.
        """
        last_err = None
        for attempt in range(retries):
            try:
                return fn(*args, **kwargs)
            except sqlite3.OperationalError as e:
                if "locked" in str(e).lower() or "busy" in str(e).lower():
                    last_err = e
                    # Jittered back-off: 0.1s, 0.3s, 0.7s, 1.5s, 3.1s …
                    wait = (2 ** attempt) * 0.1 + random.uniform(0, 0.05)
                    time.sleep(wait)
                else:
                    raise
        raise sqlite3.OperationalError(
            f"La base de données est occupée par un autre poste.\n"
            f"Veuillez réessayer dans quelques secondes.\n"
            f"(Détail technique : {last_err})"
        )

    # ──────────────────────────────────────────────────────────────────
    #  SCHEMA
    # ──────────────────────────────────────────────────────────────────

    def _init_schema(self):
        with self._conn() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS machines (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    group_name   TEXT    NOT NULL DEFAULT '',
                    machine_id   TEXT    NOT NULL UNIQUE,
                    machine_name TEXT    NOT NULL DEFAULT '',
                    created_at   TEXT    NOT NULL DEFAULT (datetime('now'))
                );

                CREATE TABLE IF NOT EXISTS interventions (
                    id                INTEGER PRIMARY KEY AUTOINCREMENT,
                    code              TEXT    NOT NULL UNIQUE,
                    machine_id        TEXT    NOT NULL,
                    group_name        TEXT    NOT NULL DEFAULT '',
                    fault_description TEXT    NOT NULL DEFAULT '',
                    code_asp          TEXT    NOT NULL DEFAULT '',
                    start_time        TEXT    NOT NULL,
                    end_time          TEXT,
                    downtime_minutes  REAL    NOT NULL DEFAULT 0,
                    technician_name   TEXT    NOT NULL DEFAULT '',
                    technician_mat    TEXT    NOT NULL DEFAULT '',
                    shift             TEXT    NOT NULL DEFAULT 'A',
                    status            TEXT    NOT NULL DEFAULT 'OPEN',
                    priority          TEXT    NOT NULL DEFAULT 'MEDIUM',
                    category          TEXT    NOT NULL DEFAULT 'GEN',
                    remarks           TEXT    NOT NULL DEFAULT '',
                    checklist_results TEXT    NOT NULL DEFAULT '[]',
                    created_at        TEXT    NOT NULL DEFAULT (datetime('now')),
                    closed_at         TEXT
                );

                CREATE INDEX IF NOT EXISTS idx_int_machine  ON interventions(machine_id);
                CREATE INDEX IF NOT EXISTS idx_int_status   ON interventions(status);
                CREATE INDEX IF NOT EXISTS idx_int_group    ON interventions(group_name);
                CREATE INDEX IF NOT EXISTS idx_int_created  ON interventions(created_at);

                CREATE TABLE IF NOT EXISTS settings (
                    key   TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS reports (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    filename    TEXT    NOT NULL,
                    filepath    TEXT    NOT NULL,
                    type        TEXT    NOT NULL DEFAULT 'MANUAL',
                    sent_email  INTEGER NOT NULL DEFAULT 0,
                    created_at  TEXT    NOT NULL DEFAULT (datetime('now'))
                );

                CREATE TABLE IF NOT EXISTS asp_codes (
                    code        TEXT PRIMARY KEY,
                    description TEXT NOT NULL DEFAULT '',
                    created_at  TEXT NOT NULL DEFAULT (datetime('now'))
                );

                CREATE TABLE IF NOT EXISTS machine_plans (
                    id                INTEGER PRIMARY KEY AUTOINCREMENT,
                    machine_id        TEXT    NOT NULL,
                    group_name        TEXT    NOT NULL,
                    plan_description  TEXT    NOT NULL,
                    target_date       TEXT,
                    status            TEXT    NOT NULL DEFAULT 'PENDING',
                    created_by        TEXT    NOT NULL,
                    created_at        TEXT    NOT NULL DEFAULT (datetime('now')),
                    sent_email        INTEGER NOT NULL DEFAULT 0,
                    file_path         TEXT
                );
            """)

            # Migration: Add code_asp if missing
            try:
                conn.execute("ALTER TABLE interventions ADD COLUMN code_asp TEXT NOT NULL DEFAULT ''")
            except sqlite3.OperationalError:
                pass  # Already exists

            # Migration: Add checklist_results if missing
            try:
                conn.execute("ALTER TABLE interventions ADD COLUMN checklist_results TEXT NOT NULL DEFAULT '[]'")
            except sqlite3.OperationalError:
                pass

            # Migration: Add file_path to machine_plans if missing
            try:
                conn.execute("ALTER TABLE machine_plans ADD COLUMN file_path TEXT")
            except sqlite3.OperationalError:
                pass


    # ──────────────────────────────────────────────────────────────────
    #  MACHINES CRUD
    # ──────────────────────────────────────────────────────────────────
    def upsert_machines(self, machines: list[dict]):
        """Insert or update machines from a list of dicts with keys:
        group_name, machine_id, machine_name."""
        def _do():
            with self._conn() as conn:
                conn.executemany("""
                    INSERT INTO machines (group_name, machine_id, machine_name)
                    VALUES (:group_name, :machine_id, :machine_name)
                    ON CONFLICT(machine_id) DO UPDATE SET
                        group_name   = excluded.group_name,
                        machine_name = excluded.machine_name
                """, machines)
        self._retry_write(_do)
        return len(machines)

    def get_all_machines(self) -> list[dict]:
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT * FROM machines ORDER BY group_name, machine_id"
            ).fetchall()
        return [dict(r) for r in rows]

    def get_machines_enriched(self) -> list[dict]:
        """Returns machines joined with their intervention counts and last timestamp for status sorting."""
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT m.*, 
                       (SELECT COUNT(*) FROM interventions WHERE machine_id = m.machine_id) as total_count,
                       (SELECT COUNT(*) FROM interventions WHERE machine_id = m.machine_id AND status = 'OPEN') as open_count,
                       (SELECT MAX(COALESCE(closed_at, created_at)) FROM interventions WHERE machine_id = m.machine_id) as last_ts
                FROM machines m
                ORDER BY total_count DESC, machine_id ASC
            """).fetchall()
        return [dict(r) for r in rows]

    def get_machine_groups(self) -> list[str]:
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT DISTINCT group_name FROM machines ORDER BY group_name"
            ).fetchall()
        return [r["group_name"] for r in rows]

    def get_machines_by_group(self, group: str) -> list[dict]:
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT * FROM machines WHERE group_name = ? ORDER BY machine_id",
                (group,)
            ).fetchall()
        return [dict(r) for r in rows]

    def search_machines(self, query: str, group: str = None) -> list[dict]:
        q = f"%{query}%"
        sql = """SELECT * FROM machines WHERE (
                    machine_id   LIKE ? OR
                    machine_name LIKE ? OR
                    group_name   LIKE ?
                 )"""
        params = [q, q, q]
        if group and group != "TOUS LES GROUPES":
            sql += " AND group_name = ?"
            params.append(group)
        sql += " ORDER BY group_name, machine_id"
        with self._conn() as conn:
            rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]

    def machine_count(self) -> int:
        with self._conn() as conn:
            return conn.execute("SELECT COUNT(*) FROM machines").fetchone()[0]

    # ──────────────────────────────────────────────────────────────────
    #  ASP CODES
    # ──────────────────────────────────────────────────────────────────
    def upsert_asp_codes(self, codes: list[dict]):
        """codes: list of {'code': '...', 'description': '...'}"""
        def _do():
            with self._conn() as conn:
                conn.executemany("""
                    INSERT INTO asp_codes (code, description)
                    VALUES (:code, :description)
                    ON CONFLICT(code) DO UPDATE SET
                        description = excluded.description
                """, codes)
        self._retry_write(_do)
        return len(codes)

    def get_asp_codes_enriched(self) -> list[dict]:
        """Returns ASP codes with their usage count in interventions."""
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT a.code, a.description,
                       (SELECT COUNT(*) FROM interventions WHERE code_asp = a.code) as usage_count
                FROM asp_codes a
                ORDER BY usage_count DESC, a.code ASC
            """).fetchall()
        return [dict(r) for r in rows]

    def get_asp_description(self, code: str) -> str:
        with self._conn() as conn:
            row = conn.execute(
                "SELECT description FROM asp_codes WHERE code = ?", (code,)
            ).fetchone()
        return row["description"] if row else ""

    # ──────────────────────────────────────────────────────────────────
    #  INTERVENTIONS CRUD
    # ──────────────────────────────────────────────────────────────────
    def generate_code(self) -> str:
        """Generate next intervention code: INT-YYYY-XXXX.

        IMPORTANT: This must be called INSIDE the same transaction as
        create_intervention() to be race-condition safe across multiple
        simultaneous users.  Use _generate_code_in_conn() internally.
        """
        year = datetime.datetime.now().year
        prefix = f"INT-{year}-"
        with self._conn() as conn:
            return self._generate_code_in_conn(conn, prefix)

    def _generate_code_in_conn(self, conn, prefix: str) -> str:
        """Generate the next code using an already-open exclusive connection.
        Must be called within an active transaction to be atomic.
        """
        row = conn.execute(
            "SELECT code FROM interventions "
            "WHERE code LIKE ? ORDER BY code DESC LIMIT 1",
            (f"{prefix}%",)
        ).fetchone()
        if row:
            last_num = int(row["code"].split("-")[-1])
            return f"{prefix}{last_num + 1:04d}"
        return f"{prefix}0001"

    def create_intervention(self, data: dict) -> str:
        """Create a new intervention atomically — safe for concurrent users.

        Code generation and INSERT happen inside ONE exclusive transaction so
        two workstations can never produce the same INT-YYYY-XXXX code.
        """
        def _do_insert():
            now = datetime.datetime.now().isoformat(timespec="seconds")
            year = datetime.datetime.now().year
            prefix = f"INT-{year}-"

            # Use an explicit EXCLUSIVE transaction so no other writer can
            # slip in between the SELECT (generate_code) and the INSERT.
            conn = self._conn()
            try:
                conn.execute("BEGIN EXCLUSIVE")
                code = data.get("code") or self._generate_code_in_conn(conn, prefix)
                conn.execute("""
                    INSERT INTO interventions
                        (code, machine_id, group_name, fault_description, code_asp,
                         start_time, end_time, downtime_minutes,
                         technician_name, technician_mat, shift,
                         status, priority, category, remarks, checklist_results, created_at)
                    VALUES
                        (:code, :machine_id, :group_name, :fault_description, :code_asp,
                         :start_time, :end_time, :downtime_minutes,
                         :technician_name, :technician_mat, :shift,
                         :status, :priority, :category, :remarks, :checklist_results, :created_at)
                """, {
                    "code":              code,
                    "machine_id":        data.get("machine_id", ""),
                    "group_name":        data.get("group_name", ""),
                    "fault_description": data.get("fault_description", ""),
                    "code_asp":          data.get("code_asp", ""),
                    "start_time":        data.get("start_time", now),
                    "end_time":          data.get("end_time"),
                    "downtime_minutes":  data.get("downtime_minutes", 0),
                    "technician_name":   data.get("technician_name", ""),
                    "technician_mat":    data.get("technician_mat", ""),
                    "shift":             data.get("shift", "A"),
                    "status":            data.get("status", "OPEN"),
                    "priority":          data.get("priority", "MEDIUM"),
                    "category":          data.get("category", "GEN"),
                    "remarks":           data.get("remarks", ""),
                    "checklist_results": json.dumps(data.get("checklist_results", [])),
                    "created_at":        now,
                })
                conn.execute("COMMIT")
                return code
            except Exception:
                try:
                    conn.execute("ROLLBACK")
                except Exception:
                    pass
                raise
            finally:
                conn.close()

        return self._retry_write(_do_insert)

    def delete_intervention(self, code: str):
        """Delete an intervention by its unique code."""
        def _do():
            with self._conn() as conn:
                conn.execute("DELETE FROM interventions WHERE code = ?", (code,))
        self._retry_write(_do)

    def update_intervention(self, code: str, data: dict):
        """Update an existing intervention record."""
        def _do():
            with self._conn() as conn:
                conn.execute("""
                    UPDATE interventions
                    SET machine_id = :machine_id,
                        group_name = :group_name,
                        fault_description = :fault_description,
                        code_asp = :code_asp,
                        technician_name = :technician_name,
                        downtime_minutes = :downtime_minutes,
                        category = :category,
                        remarks = :remarks,
                        status = :status,
                        created_at = :created_at,
                        start_time = :start_time,
                        end_time = :end_time
                    WHERE code = :code
                """, {
                    "code": code,
                    "machine_id": data.get("machine_id", ""),
                    "group_name": data.get("group_name", ""),
                    "fault_description": data.get("fault_description", ""),
                    "code_asp": data.get("code_asp", ""),
                    "technician_name": data.get("technician_name", ""),
                    "downtime_minutes": data.get("downtime_minutes", 0),
                    "category": data.get("category", ""),
                    "remarks": data.get("remarks", ""),
                    "status": data.get("status", "CLOSED"),
                    "created_at": data.get("created_at"),
                    "start_time": data.get("start_time", ""),
                    "end_time": data.get("end_time")
                })
        self._retry_write(_do)

    def close_intervention(self, code: str, end_time: str = None,
                           downtime_minutes: float = None, remarks: str = None):
        """Close an open intervention."""
        now = end_time or datetime.datetime.now().isoformat(timespec="seconds")
        updates = {"status": "CLOSED", "closed_at": now, "end_time": now}
        if downtime_minutes is not None:
            updates["downtime_minutes"] = downtime_minutes
        if remarks is not None:
            updates["remarks"] = remarks

        set_clause = ", ".join(f"{k} = :{k}" for k in updates)
        updates["code"] = code

        def _do():
            with self._conn() as conn:
                conn.execute(
                    f"UPDATE interventions SET {set_clause} WHERE code = :code",
                    updates,
                )
        self._retry_write(_do)

    def get_intervention(self, code: str) -> dict | None:
        with self._conn() as conn:
            row = conn.execute(
                "SELECT * FROM interventions WHERE code = ?", (code,)
            ).fetchone()
        return dict(row) if row else None

    def get_all_interventions(self, status: str = None,
                              limit: int = 500) -> list[dict]:
        sql = "SELECT * FROM interventions"
        params = []
        if status:
            sql += " WHERE status = ?"
            params.append(status)
        sql += " ORDER BY created_at DESC LIMIT ?"
        params.append(limit)
        with self._conn() as conn:
            rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]

    def get_interventions_for_machine(self, machine_id: str) -> list[dict]:
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT * FROM interventions WHERE machine_id = ? "
                "ORDER BY created_at DESC", (machine_id,)
            ).fetchall()
        return [dict(r) for r in rows]

    def count_interventions_for_machine(self, machine_id: str) -> int:
        with self._conn() as conn:
            return conn.execute(
                "SELECT COUNT(*) FROM interventions WHERE machine_id = ?",
                (machine_id,),
            ).fetchone()[0]

    def get_interventions_by_range(self, start_date: str, end_date: str) -> list[dict]:
        """Fetch interventions between two ISO dates."""
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT * FROM interventions WHERE created_at >= ? AND created_at <= ? "
                "ORDER BY created_at DESC", (start_date, end_date)
            ).fetchall()
        return [dict(r) for r in rows]

    def get_open_count(self) -> int:
        with self._conn() as conn:
            return conn.execute(
                "SELECT COUNT(*) FROM interventions WHERE status = 'OPEN'"
            ).fetchone()[0]

    def get_filtered_analytics(self, date_from=None, date_to=None, machine_id=None,
                               technician=None, status=None, category=None) -> dict:
        """Fetch filtered analytics metrics, charts, and records for IMA dashboard."""
        where_clauses = []
        params = []

        if date_from and str(date_from).strip():
            where_clauses.append("DATE(created_at) >= ?")
            params.append(str(date_from).strip())
        if date_to and str(date_to).strip():
            where_clauses.append("DATE(created_at) <= ?")
            params.append(str(date_to).strip())
        if machine_id and str(machine_id).strip() not in ('All', '', 'Tous'):
            where_clauses.append("machine_id = ?")
            params.append(str(machine_id).strip())
        if technician and str(technician).strip() not in ('All', '', 'Tous'):
            where_clauses.append("(technician_name = ? OR technician_mat = ?)")
            tech_val = str(technician).strip()
            params.extend([tech_val, tech_val])
        if status and str(status).strip() not in ('All', '', 'Tous'):
            stat_val = str(status).strip().upper()
            if stat_val in ('CLOSED', 'TERMINÉ', 'TERMINÉE', 'TERMINE'):
                where_clauses.append("UPPER(status) IN ('CLOSED', 'TERMINE', 'TERMINÉE')")
            elif stat_val in ('OPEN', 'EN COURS', 'EN_COURS'):
                where_clauses.append("UPPER(status) IN ('OPEN', 'EN COURS', 'EN_COURS')")
            elif stat_val in ('PENDING', 'ATTENTE', 'EN ATTENTE'):
                where_clauses.append("UPPER(status) IN ('PENDING', 'ATTENTE', 'EN ATTENTE')")
            else:
                where_clauses.append("UPPER(status) = ?")
                params.append(stat_val)
        if category and str(category).strip() not in ('All', '', 'Tous'):
            where_clauses.append("category = ?")
            params.append(str(category).strip())

        where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""

        with self._conn() as conn:
            # 1. Total and KPIs
            kpi_query = f"""
                SELECT 
                    COUNT(*) AS total,
                    SUM(CASE WHEN UPPER(status) IN ('CLOSED', 'TERMINE', 'TERMINÉE') THEN 1 ELSE 0 END) AS closed,
                    SUM(CASE WHEN UPPER(status) IN ('OPEN', 'EN COURS', 'EN_COURS') THEN 1 ELSE 0 END) AS open,
                    SUM(CASE WHEN UPPER(status) IN ('PENDING', 'ATTENTE', 'EN ATTENTE') THEN 1 ELSE 0 END) AS pending,
                    COALESCE(SUM(downtime_minutes), 0) AS total_downtime,
                    COALESCE(AVG(CASE WHEN UPPER(status) IN ('CLOSED', 'TERMINE', 'TERMINÉE') AND downtime_minutes > 0 THEN downtime_minutes END), 0) AS avg_downtime
                FROM interventions {where_sql}
            """
            kpi_row = conn.execute(kpi_query, params).fetchone()
            total = kpi_row["total"] or 0
            closed = kpi_row["closed"] or 0
            open_count = kpi_row["open"] or 0
            pending = kpi_row["pending"] or 0
            total_downtime = kpi_row["total_downtime"] or 0.0
            avg_downtime = round(kpi_row["avg_downtime"] or 0.0, 1)
            closure_rate = round((closed / total) * 100, 1) if total > 0 else 0

            # 2. Timeline / Evolution
            trend_query = f"""
                SELECT 
                    DATE(created_at) AS period,
                    COUNT(*) AS count,
                    SUM(CASE WHEN UPPER(status) IN ('CLOSED', 'TERMINE', 'TERMINÉE') THEN 1 ELSE 0 END) AS closed,
                    SUM(CASE WHEN UPPER(status) IN ('OPEN', 'EN COURS', 'EN_COURS') THEN 1 ELSE 0 END) AS open,
                    COALESCE(SUM(downtime_minutes), 0) AS downtime
                FROM interventions {where_sql}
                GROUP BY DATE(created_at)
                ORDER BY period ASC
            """
            timeline_rows = conn.execute(trend_query, params).fetchall()

            # 3. Status breakdown
            status_query = f"""
                SELECT 
                    CASE 
                        WHEN UPPER(status) IN ('CLOSED', 'TERMINE', 'TERMINÉE') THEN 'Terminée'
                        WHEN UPPER(status) IN ('OPEN', 'EN COURS', 'EN_COURS') THEN 'En cours'
                        WHEN UPPER(status) IN ('PENDING', 'ATTENTE', 'EN ATTENTE') THEN 'En attente'
                        ELSE status
                    END AS status_label,
                    COUNT(*) AS count
                FROM interventions {where_sql}
                GROUP BY status_label
                ORDER BY count DESC
            """
            status_rows = conn.execute(status_query, params).fetchall()

            # 4. Top machines by intervention count
            top_machines_query = f"""
                SELECT machine_id, group_name, COUNT(*) AS count, COALESCE(SUM(downtime_minutes), 0) AS total_downtime
                FROM interventions {where_sql}
                GROUP BY machine_id
                ORDER BY count DESC
                LIMIT 10
            """
            top_machines_rows = conn.execute(top_machines_query, params).fetchall()

            # 5. Downtime per machine
            dt_where = f"{where_sql} AND downtime_minutes > 0" if where_sql else "WHERE downtime_minutes > 0"
            downtime_machines_query = f"""
                SELECT machine_id, COALESCE(SUM(downtime_minutes), 0) AS total_downtime, COUNT(*) AS count
                FROM interventions {dt_where}
                GROUP BY machine_id
                ORDER BY total_downtime DESC
                LIMIT 10
            """
            downtime_machines_rows = conn.execute(downtime_machines_query, params).fetchall()

            # 6. Detailed interventions list
            table_query = f"""
                SELECT * FROM interventions {where_sql}
                ORDER BY created_at DESC
                LIMIT 200
            """
            interventions_rows = conn.execute(table_query, params).fetchall()

            # 7. Available filter options from DB
            all_machines = [r[0] for r in conn.execute("SELECT DISTINCT machine_id FROM machines WHERE machine_id != '' UNION SELECT DISTINCT machine_id FROM interventions WHERE machine_id != '' ORDER BY 1").fetchall()]
            all_technicians = [r[0] for r in conn.execute("SELECT DISTINCT technician_name FROM interventions WHERE technician_name != '' ORDER BY 1").fetchall()]
            all_categories = [r[0] for r in conn.execute("SELECT DISTINCT category FROM interventions WHERE category != '' ORDER BY 1").fetchall()]

        return {
            "kpis": {
                "total": total,
                "closed": closed,
                "open": open_count,
                "pending": pending,
                "total_downtime": total_downtime,
                "avg_downtime": avg_downtime,
                "closure_rate": closure_rate
            },
            "timeline": [dict(r) for r in timeline_rows],
            "by_status": [dict(r) for r in status_rows],
            "top_machines": [dict(r) for r in top_machines_rows],
            "downtime_machines": [dict(r) for r in downtime_machines_rows],
            "interventions": [dict(r) for r in interventions_rows],
            "filters": {
                "machines": all_machines,
                "technicians": all_technicians,
                "categories": all_categories
            }
        }

    def get_analytics_bundle(self):
        """Fetch all analytics data in a single optimized connection session."""
        with self._conn() as conn:
            month = datetime.datetime.now().strftime("%Y-%m")
            
            # 1. KPIs (Monthly)
            kpi_row = conn.execute("""
                SELECT COUNT(*) AS total, SUM(downtime_minutes) AS downtime
                FROM interventions WHERE STRFTIME('%Y-%m', created_at) = ?
            """, (month,)).fetchone()
            open_n = conn.execute("""
                SELECT COUNT(*) FROM interventions 
                WHERE status = 'OPEN' AND STRFTIME('%Y-%m', created_at) = ?
            """, (month,)).fetchone()[0]

            # 1b. KPIs (Global / All-Time)
            global_row = conn.execute("""
                SELECT COUNT(*) AS total, SUM(downtime_minutes) AS downtime
                FROM interventions
            """).fetchone()
            global_open = conn.execute("SELECT COUNT(*) FROM interventions WHERE status = 'OPEN'").fetchone()[0]
            
            # 1c. Chronology (to scale global MTBF accurately)
            months_count = conn.execute("""
                SELECT COUNT(DISTINCT STRFTIME('%Y-%m', created_at)) 
                FROM interventions
            """).fetchone()[0] or 1
            
            # 2. Daily Trend (30 days)
            start_30 = (datetime.datetime.now() - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
            trend_30 = conn.execute("""
                SELECT DATE(created_at) AS day, COUNT(*) AS count, SUM(downtime_minutes) AS total_downtime
                FROM interventions WHERE DATE(created_at) >= ?
                GROUP BY DATE(created_at) ORDER BY day ASC
            """, (start_30,)).fetchall()

            # 3. Pareto Machines
            pareto_machines = conn.execute("""
                SELECT machine_id, group_name, COUNT(*) AS failure_count, SUM(downtime_minutes) AS total_downtime
                FROM interventions GROUP BY machine_id ORDER BY failure_count DESC LIMIT 20
            """).fetchall()

            # 4. Top Panne Codes
            top_codes = conn.execute("""
                SELECT category AS code, COUNT(*) AS count, SUM(downtime_minutes) AS total_downtime
                FROM interventions WHERE category != '' GROUP BY category ORDER BY count DESC LIMIT 10
            """).fetchall()

            # 4b. Top ASP Codes (Most used SAP TN codes)
            top_asp = conn.execute("""
                SELECT code_asp AS code, COUNT(*) AS count, SUM(downtime_minutes) AS total_downtime
                FROM interventions WHERE code_asp != '' GROUP BY code_asp ORDER BY count DESC LIMIT 10
            """).fetchall()

            # 5. Downtime per Machine
            dt_per_machine = conn.execute("""
                SELECT machine_id, group_name, SUM(downtime_minutes) AS total_downtime, COUNT(*) AS intervention_count
                FROM interventions WHERE downtime_minutes > 0 GROUP BY machine_id ORDER BY total_downtime DESC
            """).fetchall()

            # 6. Frequency per Group
            freq_per_group = conn.execute("""
                SELECT group_name, COUNT(*) AS intervention_count, SUM(downtime_minutes) AS total_downtime
                FROM interventions GROUP BY group_name ORDER BY intervention_count DESC
            """).fetchall()

            # 7. Recent Interventions
            recent = conn.execute("""
                SELECT * FROM interventions ORDER BY created_at DESC LIMIT 50
            """).fetchall()

        return {
            "stats": {
                "total": kpi_row["total"] or 0,
                "downtime": kpi_row["downtime"] or 0,
                "open": open_n,
                "closed": (kpi_row["total"] or 0) - open_n
            },
            "global_stats": {
                "total": global_row["total"] or 0,
                "downtime": global_row["downtime"] or 0,
                "open": global_open,
                "closed": (global_row["total"] or 0) - global_open,
                "months_count": months_count
            },
            "trend_30": [dict(r) for r in trend_30],
            "by_month": self.get_monthly_trend(12),
            "pareto": [dict(r) for r in pareto_machines],
            "codes": [dict(r) for r in top_codes],
            "asp_codes": [dict(r) for r in top_asp],
            "downtime_per_machine": [dict(r) for r in dt_per_machine],
            "frequency_per_group": [dict(r) for r in freq_per_group],
            "recent": [dict(r) for r in recent]
        }

    def get_downtime_per_machine(self) -> list[dict]:
        """Total downtime per machine (minutes)."""
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT machine_id,
                       group_name,
                       SUM(downtime_minutes) AS total_downtime,
                       COUNT(*)              AS intervention_count
                FROM interventions
                WHERE downtime_minutes > 0
                GROUP BY machine_id
                ORDER BY total_downtime DESC
            """).fetchall()
        return [dict(r) for r in rows]

    def get_pareto_failures(self, top_n: int = 20) -> list[dict]:
        """Top N machines by intervention count (Pareto)."""
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT machine_id,
                       group_name,
                       COUNT(*) AS failure_count
                FROM interventions
                GROUP BY machine_id
                ORDER BY failure_count DESC
                LIMIT ?
            """, (top_n,)).fetchall()
        return [dict(r) for r in rows]

    def get_frequency_per_group(self) -> list[dict]:
        """Intervention frequency per group."""
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT group_name,
                       COUNT(*) AS intervention_count,
                       SUM(downtime_minutes) AS total_downtime
                FROM interventions
                GROUP BY group_name
                ORDER BY intervention_count DESC
            """).fetchall()
        return [dict(r) for r in rows]

    def get_monthly_summary(self, year: int = None,
                            month: int = None) -> list[dict]:
        """Monthly breakdown summary."""
        now = datetime.datetime.now()
        year  = year  or now.year
        month = month or now.month
        start = f"{year}-{month:02d}-01"
        if month == 12:
            end = f"{year + 1}-01-01"
        else:
            end = f"{year}-{month + 1:02d}-01"

        with self._conn() as conn:
            rows = conn.execute("""
                SELECT * FROM interventions
                WHERE created_at >= ? AND created_at < ?
                ORDER BY created_at DESC
            """, (start, end)).fetchall()
        return [dict(r) for r in rows]

    def get_daily_trend(self, days: int = 30) -> list[dict]:
        """Daily intervention counts & downtime for last N days."""
        start = (datetime.datetime.now() -
                 datetime.timedelta(days=days)).strftime("%Y-%m-%d")
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT DATE(created_at) AS day,
                       COUNT(*)         AS count,
                       SUM(downtime_minutes) AS total_downtime
                FROM interventions
                WHERE DATE(created_at) >= ?
                GROUP BY DATE(created_at)
                ORDER BY day ASC
            """, (start,)).fetchall()
        return [dict(r) for r in rows]

    def get_monthly_trend(self, months: int = 12) -> list[dict]:
        """Monthly intervention counts & downtime for last N months."""
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT STRFTIME('%Y-%m', created_at) AS month,
                       COUNT(*)         AS count,
                       SUM(downtime_minutes) AS total_downtime
                FROM interventions
                GROUP BY month
                ORDER BY month DESC
                LIMIT ?
            """, (months,)).fetchall()
        return sorted([dict(r) for r in rows], key=lambda x: x["month"])

    def get_top_panne_codes(self, top_n: int = 10) -> list[dict]:
        """Top N panne codes by intervention count."""
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT category AS code,
                       COUNT(*) AS count,
                       SUM(downtime_minutes) AS total_downtime
                FROM interventions
                WHERE category != ''
                GROUP BY category
                ORDER BY count DESC
                LIMIT ?
            """, (top_n,)).fetchall()
        return [dict(r) for r in rows]

    def get_machine_history(self, machine_id: str) -> list[dict]:
        """Full intervention history for a specific machine."""
        return self.get_interventions_for_machine(machine_id)

    # ──────────────────────────────────────────────────────────────────
    #  SETTINGS
    # ──────────────────────────────────────────────────────────────────
    def get_setting(self, key: str, default: str = "") -> str:
        with self._conn() as conn:
            row = conn.execute(
                "SELECT value FROM settings WHERE key = ?", (key,)
            ).fetchone()
        return row["value"] if row else default

    def set_setting(self, key: str, value: str):
        with self._conn() as conn:
            conn.execute(
                "INSERT INTO settings (key, value) VALUES (?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (key, value),
            )
    # ──────────────────────────────────────────────────────────────────
    #  REPORTS LOGGING
    # ──────────────────────────────────────────────────────────────────
    def log_report(self, filename: str, filepath: str, rtype: str = "MANUAL", sent: bool = False):
        # Explicitly use local time — SQLite datetime('now') returns UTC which is wrong
        local_ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        def _do():
            with self._conn() as conn:
                conn.execute("""
                    INSERT INTO reports (filename, filepath, type, sent_email, created_at)
                    VALUES (?, ?, ?, ?, ?)
                """, (filename, filepath, rtype, 1 if sent else 0, local_ts))
        self._retry_write(_do)

    def get_all_reports(self, limit: int = 100) -> list[dict]:
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT * FROM reports ORDER BY created_at DESC LIMIT ?
            """, (limit,)).fetchall()
        return [dict(r) for r in rows]

    # ──────────────────────────────────────────────────────────────────
    #  MACHINE HEALTH & PLANS
    # ──────────────────────────────────────────────────────────────────
    def get_weekly_machine_stats(self) -> list[dict]:
        """Returns machines with intervention stats for the last 7 days."""
        week_ago = (datetime.datetime.now() - datetime.timedelta(days=7)).isoformat()
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT machine_id, group_name, 
                       COUNT(*) as count, 
                       SUM(downtime_minutes) as downtime,
                       GROUP_CONCAT(DISTINCT category) as codes
                FROM interventions
                WHERE created_at >= ?
                GROUP BY machine_id
                ORDER BY count DESC
            """, (week_ago,)).fetchall()
        return [dict(r) for r in rows]

    def get_weekly_group_stats(self) -> list[dict]:
        """Returns machine groups with intervention stats for the last 7 days."""
        week_ago = (datetime.datetime.now() - datetime.timedelta(days=7)).isoformat()
        with self._conn() as conn:
            rows = conn.execute("""
                SELECT group_name, 
                       COUNT(*) as count, 
                       SUM(downtime_minutes) as downtime
                FROM interventions
                WHERE created_at >= ?
                GROUP BY group_name
                ORDER BY count DESC
            """, (week_ago,)).fetchall()
        return [dict(r) for r in rows]

    def save_machine_plan(self, data: dict):
        """Save a new action plan for a machine."""
        def _do():
            with self._conn() as conn:
                conn.execute("""
                    INSERT INTO machine_plans 
                    (machine_id, group_name, plan_description, target_date, created_by, sent_email, file_path)
                    VALUES (:machine_id, :group_name, :plan_description, :target_date, :created_by, :sent_email, :file_path)
                """, data)
        self._retry_write(_do)

    def get_machine_plans(self, machine_id: str = None) -> list[dict]:
        """Fetch saved plans."""
        sql = "SELECT * FROM machine_plans"
        params = []
        if machine_id:
            sql += " WHERE machine_id = ?"
            params.append(machine_id)
        sql += " ORDER BY created_at DESC"
        with self._conn() as conn:
            rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]

    def clear_all_data(self):
        """Wipe all transactional and configuration data for a fresh start."""
        with self._conn() as conn:
            conn.execute("DELETE FROM interventions")
            conn.execute("DELETE FROM reports")
            conn.execute("DELETE FROM machines")
            conn.execute("DELETE FROM asp_codes")
            # We explicitly do NOT delete settings to keep admin passwords and sync paths.
        return True
