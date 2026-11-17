"""
org_manager.py
==============
Dynamic Organizational Chart & Employee Structure Engine for SEBN-TN Maintenance System.
Handles:
- Configurable hierarchical structure: Manager -> Group Leaders -> Team Leaders -> Supervisors -> Technicians
- Shift partitioning: Shift A, Shift B, Shift C (with Shift supervisors & technicians)
- Dynamic employee assignments, shift transfers, promotions, supervisor changes
- Hiring date (Date d'embauche) & automatic real-time seniority calculation
- Complete assignment audit history
- Synchronization with unified `users` table without duplicating accounts
"""

import sqlite3
import logging
import datetime
from typing import Dict, Any, List, Optional, Tuple

logger = logging.getLogger("sebn-maintenance")

# Hierarchy Levels Definitions
POSITION_LEVELS = {
    "MANAGER": {
        "code": "MANAGER",
        "rank": 1,
        "label": "Manager / Direction",
        "badge_class": "badge-manager",
        "icon": "fa-solid fa-crown"
    },
    "GROUP_LEADER": {
        "code": "GROUP_LEADER",
        "rank": 2,
        "label": "Chef de Groupe",
        "badge_class": "badge-groupleader",
        "icon": "fa-solid fa-users-gear"
    },
    "TEAM_LEADER": {
        "code": "TEAM_LEADER",
        "rank": 3,
        "label": "Chef d'Équipe",
        "badge_class": "badge-teamleader",
        "icon": "fa-solid fa-user-tie"
    },
    "SUPERVISOR": {
        "code": "SUPERVISOR",
        "rank": 4,
        "label": "Superviseur",
        "badge_class": "badge-supervisor",
        "icon": "fa-solid fa-user-check"
    },
    "TECHNICIAN": {
        "code": "TECHNICIAN",
        "rank": 5,
        "label": "Technicien",
        "badge_class": "badge-technician",
        "icon": "fa-solid fa-wrench"
    }
}

VALID_SHIFTS = ["A", "B", "C", "NOT_ASSIGNED"]


def calculate_seniority(hire_date_str: Optional[str], as_of_date: Optional[datetime.date] = None) -> Dict[str, Any]:
    """
    Calculates employee seniority automatically from their hiring date.
    Returns structured data and human-friendly French string (e.g. '3 ans 4 mois', '8 mois', '12 jours').
    """
    if not hire_date_str or not str(hire_date_str).strip():
        return {
            "text": "Non renseignée",
            "short_text": "-",
            "years": 0,
            "months": 0,
            "days": 0,
            "total_days": 0,
            "formatted_hire_date": "-",
            "raw_hire_date": ""
        }

    raw = str(hire_date_str).strip()[:10]
    try:
        parts = [int(p) for p in raw.split("-")]
        if len(parts) == 3:
            h_date = datetime.date(parts[0], parts[1], parts[2])
        else:
            return {
                "text": "Date invalide",
                "short_text": "-",
                "years": 0,
                "months": 0,
                "days": 0,
                "total_days": 0,
                "formatted_hire_date": raw,
                "raw_hire_date": raw
            }
    except Exception:
        return {
            "text": "Date invalide",
            "short_text": "-",
            "years": 0,
            "months": 0,
            "days": 0,
            "total_days": 0,
            "formatted_hire_date": raw,
            "raw_hire_date": raw
        }

    now = as_of_date or datetime.date.today()
    if h_date > now:
        return {
            "text": "À venir (embauche future)",
            "short_text": "À venir",
            "years": 0,
            "months": 0,
            "days": 0,
            "total_days": 0,
            "formatted_hire_date": h_date.strftime("%d/%m/%Y"),
            "raw_hire_date": raw
        }

    total_days = (now - h_date).days

    # Calculate years, months, days accurately
    years = now.year - h_date.year
    months = now.month - h_date.month
    days = now.day - h_date.day

    if days < 0:
        # Borrow days from previous month
        prev_month = now.month - 1 if now.month > 1 else 12
        prev_year = now.year if now.month > 1 else now.year - 1
        # Approx days in previous month
        days_in_prev_month = 30
        try:
            import calendar
            days_in_prev_month = calendar.monthrange(prev_year, prev_month)[1]
        except Exception:
            pass
        days += days_in_prev_month
        months -= 1

    if months < 0:
        months += 12
        years -= 1

    # Format human-friendly French representation
    parts_fr = []
    if years > 0:
        parts_fr.append(f"{years} {'an' if years == 1 else 'ans'}")
    if months > 0:
        parts_fr.append(f"{months} mois")

    if not parts_fr:
        if days > 0:
            text = f"{days} {'jour' if days == 1 else 'jours'}"
        else:
            text = "Aujourd'hui"
    else:
        text = " ".join(parts_fr)

    short_text = f"{years}a {months}m" if years > 0 else f"{months}m"

    return {
        "text": text,
        "short_text": short_text,
        "years": years,
        "months": months,
        "days": days,
        "total_days": total_days,
        "formatted_hire_date": h_date.strftime("%d/%m/%Y"),
        "raw_hire_date": raw
    }


class OrgChartManager:
    """
    Manages organizational chart assignments, hierarchy nodes, shift groups, and audit history.
    """

    def __init__(self, db_path: str, data_dir: Optional[str] = None):
        self.db_path = db_path
        self.data_dir = data_dir
        self._init_db()
        self._seed_initial_structure()

    def _get_conn(self):
        conn = sqlite3.connect(self.db_path, timeout=15)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        """Ensures all tables and schema columns exist."""
        with self._get_conn() as conn:
            cur = conn.cursor()

            # Ensure hire_date column exists in users table
            cur.execute("PRAGMA table_info(users)")
            user_cols = {row["name"].lower() for row in cur.fetchall()}
            if "hire_date" not in user_cols:
                try:
                    cur.execute("ALTER TABLE users ADD COLUMN hire_date TEXT")
                    conn.commit()
                except Exception as e:
                    logger.warning(f"Could not add hire_date to users: {e}")

            # 1. Main Org Assignments Table
            cur.execute("""
                CREATE TABLE IF NOT EXISTS org_assignments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER NOT NULL UNIQUE,
                    matricule TEXT NOT NULL,
                    full_name TEXT NOT NULL,
                    position_level TEXT NOT NULL,
                    job_title TEXT NOT NULL,
                    shift TEXT DEFAULT 'NOT_ASSIGNED',
                    group_name TEXT DEFAULT '',
                    team_name TEXT DEFAULT '',
                    parent_id INTEGER DEFAULT NULL,
                    hire_date TEXT DEFAULT '',
                    notes TEXT DEFAULT '',
                    is_active INTEGER DEFAULT 1,
                    display_order INTEGER DEFAULT 0,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
                    FOREIGN KEY (parent_id) REFERENCES org_assignments(id) ON DELETE SET NULL
                )
            """)

            # 2. Org Assignment History Table
            cur.execute("""
                CREATE TABLE IF NOT EXISTS org_assignment_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    assignment_id INTEGER,
                    user_id INTEGER NOT NULL,
                    matricule TEXT NOT NULL,
                    action_type TEXT NOT NULL,
                    old_position_level TEXT,
                    new_position_level TEXT,
                    old_job_title TEXT,
                    new_job_title TEXT,
                    old_shift TEXT,
                    new_shift TEXT,
                    old_parent_id INTEGER,
                    new_parent_id INTEGER,
                    details TEXT,
                    changed_by TEXT DEFAULT 'Admin',
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # Indexes
            cur.execute("CREATE INDEX IF NOT EXISTS idx_org_user_id ON org_assignments(user_id)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_org_matricule ON org_assignments(matricule)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_org_parent_id ON org_assignments(parent_id)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_org_shift ON org_assignments(shift)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_org_level ON org_assignments(position_level)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_org_hist_user ON org_assignment_history(user_id)")

            conn.commit()

    def _seed_initial_structure(self):
        """
        Seeds an initial intelligent structure from existing users if org_assignments is empty.
        Connects Manager -> Group Leader -> Team Leader -> Supervisors (Shift A/B/C) -> Technicians.
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) as count FROM org_assignments")
            count = cur.fetchone()["count"]
            if count > 0:
                return

            # Fetch all existing users
            cur.execute("SELECT id, name, matricule, username, role, shift, photo, hire_date FROM users ORDER BY id ASC")
            users = [dict(r) for r in cur.fetchall()]
            if not users:
                return

            logger.info("Initializing Organizational Chart structure from existing users...")

            # 1. Manager (Owner or first Admin)
            owner_user = next((u for u in users if u["role"] == "OWNER"), None)
            admin_user = next((u for u in users if u["role"] == "ADMIN"), None)
            tech_users = [u for u in users if u["role"] == "TECHNICIAN"]

            manager_id = None
            group_leader_id = None

            if owner_user:
                h_date = owner_user.get("hire_date") or "2018-02-01"
                cur.execute("""
                    INSERT INTO org_assignments
                    (user_id, matricule, full_name, position_level, job_title, shift, group_name, team_name, parent_id, hire_date)
                    VALUES (?, ?, ?, 'MANAGER', 'Responsable Maintenance & Direction', 'NOT_ASSIGNED', 'Direction Technique', 'Direction', NULL, ?)
                """, (owner_user["id"], owner_user["matricule"] or "OWNER-01", owner_user["name"], h_date))
                manager_id = cur.lastrowid

                # Update hire_date on user if missing
                cur.execute("UPDATE users SET hire_date = ? WHERE id = ? AND (hire_date IS NULL OR hire_date = '')", (h_date, owner_user["id"]))

            if admin_user:
                h_date = admin_user.get("hire_date") or "2019-06-15"
                cur.execute("""
                    INSERT INTO org_assignments
                    (user_id, matricule, full_name, position_level, job_title, shift, group_name, team_name, parent_id, hire_date)
                    VALUES (?, ?, ?, 'GROUP_LEADER', 'Chef de Groupe Maintenance', 'NOT_ASSIGNED', 'Maintenance Générale', 'Coordination', ?, ?)
                """, (admin_user["id"], admin_user["matricule"] or "ADMIN-01", admin_user["name"], manager_id, h_date))
                group_leader_id = cur.lastrowid

                cur.execute("UPDATE users SET hire_date = ? WHERE id = ? AND (hire_date IS NULL OR hire_date = '')", (h_date, admin_user["id"]))

            # Distribute technicians across shifts and levels
            # If we have several technicians, designate 1 or 2 as supervisors/team leaders or keep as technicians under shifts
            # Let's assign Mohamed Issaoui (11818) or Ahmed Ben Salah as Team Leader / Supervisor if appropriate,
            # or assign them with proper supervisor hierarchy.
            seeded_parents_by_shift = {"A": None, "B": None, "C": None}

            for idx, tech in enumerate(tech_users):
                shift = (tech.get("shift") or "B").upper()
                if shift not in ("A", "B", "C"):
                    shift = "B"

                # Stagger hire dates for realistic seniority demonstration
                staggered_dates = ["2020-03-10", "2021-09-01", "2022-11-15", "2023-04-20", "2024-01-10"]
                h_date = tech.get("hire_date") or staggered_dates[idx % len(staggered_dates)]

                # For the first technician in Shift B, we can make them Supervisor of Shift B if none exists yet!
                if shift == "B" and seeded_parents_by_shift["B"] is None and idx == 0:
                    pos = "SUPERVISOR"
                    title = "Superviseur Maintenance – Shift B"
                    parent = group_leader_id or manager_id
                elif idx == 1 and seeded_parents_by_shift.get(shift) is None and len(tech_users) > 3:
                    pos = "SUPERVISOR"
                    title = f"Superviseur Maintenance – Shift {shift}"
                    parent = group_leader_id or manager_id
                else:
                    pos = "TECHNICIAN"
                    title = f"Technicien de Maintenance – Shift {shift}"
                    parent = seeded_parents_by_shift.get(shift) or group_leader_id or manager_id

                cur.execute("""
                    INSERT INTO org_assignments
                    (user_id, matricule, full_name, position_level, job_title, shift, group_name, team_name, parent_id, hire_date)
                    VALUES (?, ?, ?, ?, ?, ?, 'Maintenance Curative & Préventive', ?, ?, ?)
                """, (tech["id"], tech["matricule"], tech["name"], pos, title, shift, f"Équipe Shift {shift}", parent, h_date))
                
                new_asgn_id = cur.lastrowid
                if pos == "SUPERVISOR":
                    seeded_parents_by_shift[shift] = new_asgn_id

                cur.execute("UPDATE users SET hire_date = ? WHERE id = ? AND (hire_date IS NULL OR hire_date = '')", (h_date, tech["id"]))

                # Log initial assignment in history
                cur.execute("""
                    INSERT INTO org_assignment_history
                    (assignment_id, user_id, matricule, action_type, new_position_level, new_job_title, new_shift, new_parent_id, details, changed_by)
                    VALUES (?, ?, ?, 'AFFECTATION_INITIALE', ?, ?, ?, ?, 'Initialisation automatique de la structure', 'Système')
                """, (new_asgn_id, tech["id"], tech["matricule"], pos, title, shift, parent))

            conn.commit()

    def get_all_assignments(self, filters: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
        """
        Retrieves all organizational assignments joined with user details,
        direct supervisor info, and computed seniority.
        """
        query = """
            SELECT 
                oa.id,
                oa.user_id,
                oa.matricule,
                oa.full_name,
                oa.position_level,
                oa.job_title,
                oa.shift,
                oa.group_name,
                oa.team_name,
                oa.parent_id,
                oa.hire_date,
                oa.notes,
                oa.is_active,
                oa.display_order,
                oa.created_at,
                oa.updated_at,
                u.photo,
                u.role as user_role,
                u.is_active as user_is_active,
                u.technician_level,
                parent.full_name as parent_name,
                parent.job_title as parent_job_title,
                parent.position_level as parent_position_level,
                parent.shift as parent_shift
            FROM org_assignments oa
            INNER JOIN users u ON oa.user_id = u.id
            LEFT JOIN org_assignments parent ON oa.parent_id = parent.id
            WHERE 1=1
        """
        params = []
        filters = filters or {}

        if filters.get("shift") and filters["shift"] != "ALL":
            query += " AND oa.shift = ?"
            params.append(filters["shift"])

        if filters.get("position_level") and filters["position_level"] != "ALL":
            query += " AND oa.position_level = ?"
            params.append(filters["position_level"])

        if filters.get("group_name"):
            query += " AND oa.group_name LIKE ?"
            params.append(f"%{filters['group_name']}%")

        if filters.get("team_name"):
            query += " AND oa.team_name LIKE ?"
            params.append(f"%{filters['team_name']}%")

        if filters.get("search"):
            s = f"%{filters['search'].strip()}%"
            query += " AND (oa.full_name LIKE ? OR oa.matricule LIKE ? OR oa.job_title LIKE ?)"
            params.extend([s, s, s])

        if filters.get("is_active") is not None:
            query += " AND oa.is_active = ?"
            params.append(int(filters["is_active"]))

        query += """
            ORDER BY 
                CASE oa.position_level
                    WHEN 'MANAGER' THEN 1
                    WHEN 'GROUP_LEADER' THEN 2
                    WHEN 'TEAM_LEADER' THEN 3
                    WHEN 'SUPERVISOR' THEN 4
                    WHEN 'TECHNICIAN' THEN 5
                    ELSE 6
                END ASC,
                oa.shift ASC,
                oa.full_name ASC
        """

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            rows = [dict(r) for r in cur.fetchall()]

        results = []
        for r in rows:
            seniority_info = calculate_seniority(r.get("hire_date"))
            lvl_meta = POSITION_LEVELS.get(r["position_level"], {
                "label": r["position_level"],
                "rank": 99,
                "badge_class": "badge-neutral",
                "icon": "fa-solid fa-user"
            })
            r["level_label"] = lvl_meta["label"]
            r["level_rank"] = lvl_meta["rank"]
            r["level_badge_class"] = lvl_meta["badge_class"]
            r["level_icon"] = lvl_meta["icon"]
            r["seniority"] = seniority_info
            results.append(r)

        return results

    def get_assignment_by_id(self, assignment_id: int) -> Optional[Dict[str, Any]]:
        """Returns single assignment details with seniority and supervisor info."""
        all_items = self.get_all_assignments()
        for it in all_items:
            if it["id"] == assignment_id:
                return it
        return None

    def get_assignment_by_user_id(self, user_id: int) -> Optional[Dict[str, Any]]:
        all_items = self.get_all_assignments()
        for it in all_items:
            if it["user_id"] == user_id:
                return it
        return None

    def get_unassigned_users(self) -> List[Dict[str, Any]]:
        """
        Returns all active users who are NOT yet assigned in the organizational chart.
        Used by the Admin in the 'Assign Employee' modal to prevent duplicates.
        """
        query = """
            SELECT u.id, u.name, u.matricule, u.role, u.shift, u.hire_date, u.photo, u.technician_level
            FROM users u
            LEFT JOIN org_assignments oa ON u.id = oa.user_id AND oa.is_active = 1
            WHERE oa.id IS NULL AND u.is_active = 1
            ORDER BY u.name ASC
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(query)
            rows = [dict(r) for r in cur.fetchall()]

        for r in rows:
            r["seniority"] = calculate_seniority(r.get("hire_date"))
        return rows

    def get_org_tree(self) -> Dict[str, Any]:
        """
        Builds the complete hierarchical organization tree and shift groupings.
        Structure:
          - Managers (Level 1)
            └── Group Leaders (Level 2)
                 └── Team Leaders (Level 3)
                      ├── Shift A (Supervisors + Technicians)
                      ├── Shift B (Supervisors + Technicians)
                      └── Shift C (Supervisors + Technicians)
          - Also provides statistics: total employees, per-shift counts, average seniority.
        """
        all_employees = self.get_all_assignments({"is_active": 1})

        # Calculate statistics
        total = len(all_employees)
        shift_counts = {"A": 0, "B": 0, "C": 0, "NOT_ASSIGNED": 0}
        level_counts = {
            "MANAGER": 0,
            "GROUP_LEADER": 0,
            "TEAM_LEADER": 0,
            "SUPERVISOR": 0,
            "TECHNICIAN": 0
        }
        total_seniority_days = 0
        seniority_count = 0

        for emp in all_employees:
            shift = emp.get("shift") or "NOT_ASSIGNED"
            shift_counts[shift] = shift_counts.get(shift, 0) + 1

            lvl = emp.get("position_level") or "TECHNICIAN"
            level_counts[lvl] = level_counts.get(lvl, 0) + 1

            s_days = emp["seniority"].get("total_days", 0)
            if s_days > 0:
                total_seniority_days += s_days
                seniority_count += 1

        avg_seniority_str = "0 an"
        if seniority_count > 0:
            avg_days = total_seniority_days / seniority_count
            avg_years = int(avg_days // 365)
            avg_months = int((avg_days % 365) // 30)
            if avg_years > 0:
                avg_seniority_str = f"{avg_years} {'an' if avg_years == 1 else 'ans'} {avg_months} mois"
            else:
                avg_seniority_str = f"{avg_months} mois"

        stats = {
            "total_employees": total,
            "shift_a_count": shift_counts.get("A", 0),
            "shift_b_count": shift_counts.get("B", 0),
            "shift_c_count": shift_counts.get("C", 0),
            "unassigned_count": shift_counts.get("NOT_ASSIGNED", 0),
            "managers_count": level_counts.get("MANAGER", 0),
            "group_leaders_count": level_counts.get("GROUP_LEADER", 0),
            "team_leaders_count": level_counts.get("TEAM_LEADER", 0),
            "supervisors_count": level_counts.get("SUPERVISOR", 0),
            "technicians_count": level_counts.get("TECHNICIAN", 0),
            "avg_seniority": avg_seniority_str
        }

        # Build map by id
        emp_map = {emp["id"]: dict(emp) for emp in all_employees}
        for emp in emp_map.values():
            emp["children"] = []

        # Separate levels according to the exact 5-level SEBN-TN hierarchy:
        # Level 1: Manager
        # Level 2: Group Leader
        # Level 3: Team Leader
        # Level 4: Supervisor (Shift A, Shift B, Shift C)
        # Level 5: Technician (under Supervisor in Shift A, Shift B, Shift C)

        shifts_data = {
            "A": {"supervisors": [], "unassigned_techs": [], "total": 0},
            "B": {"supervisors": [], "unassigned_techs": [], "total": 0},
            "C": {"supervisors": [], "unassigned_techs": [], "total": 0},
            "NOT_ASSIGNED": {"supervisors": [], "unassigned_techs": [], "total": 0}
        }

        managers = []
        group_leaders = []
        team_leaders = []

        for emp in all_employees:
            lvl = emp.get("position_level")
            shift = (emp.get("shift") or "NOT_ASSIGNED").upper()
            if shift not in shifts_data:
                shift = "NOT_ASSIGNED"

            if lvl == "MANAGER":
                emp["group_leaders"] = []
                managers.append(emp)
            elif lvl == "GROUP_LEADER":
                emp["team_leaders"] = []
                group_leaders.append(emp)
            elif lvl == "TEAM_LEADER":
                emp["supervisors"] = []
                team_leaders.append(emp)
            elif lvl == "SUPERVISOR":
                emp["technicians"] = []
                shifts_data[shift]["supervisors"].append(emp)

        # Link Group Leaders to Managers
        for gl in group_leaders:
            p_id = gl.get("parent_id")
            mgr = next((m for m in managers if m["id"] == p_id), None)
            if mgr:
                mgr["group_leaders"].append(gl)

        # Link Team Leaders to Group Leaders
        for tl in team_leaders:
            p_id = tl.get("parent_id")
            gl = next((g for g in group_leaders if g["id"] == p_id), None)
            if gl:
                gl["team_leaders"].append(tl)

        # Link Supervisors to Team Leaders
        all_sups = []
        for s_data in shifts_data.values():
            all_sups.extend(s_data["supervisors"])

        for sup in all_sups:
            p_id = sup.get("parent_id")
            tl = next((t for t in team_leaders if t["id"] == p_id), None)
            if tl:
                tl["supervisors"].append(sup)

        # Link Technicians under their respective Supervisor
        for emp in all_employees:
            if emp.get("position_level") == "TECHNICIAN":
                shift = (emp.get("shift") or "NOT_ASSIGNED").upper()
                if shift not in shifts_data:
                    shift = "NOT_ASSIGNED"
                p_id = emp.get("parent_id")

                target_sup = next((s for s in shifts_data[shift]["supervisors"] if s["id"] == p_id), None)
                if not target_sup and p_id:
                    target_sup = next((s for s in all_sups if s["id"] == p_id), None)

                if target_sup:
                    target_sup["technicians"].append(emp)
                else:
                    shifts_data[shift]["unassigned_techs"].append(emp)

        for s_key, s_data in shifts_data.items():
            s_data["total"] = len(s_data["supervisors"]) + sum(len(s["technicians"]) for s in s_data["supervisors"]) + len(s_data["unassigned_techs"])

        # Build recursive hierarchy tree (linking children to parent)
        roots = []
        for emp in emp_map.values():
            p_id = emp.get("parent_id")
            if p_id and p_id in emp_map:
                emp_map[p_id]["children"].append(emp)
            else:
                roots.append(emp)

        return {
            "stats": stats,
            "all_employees": all_employees,
            "tree_roots": roots,
            "levels": {
                "managers": managers,
                "group_leaders": group_leaders,
                "team_leaders": team_leaders,
                "shifts": shifts_data
            }
        }

    def assign_employee(
        self,
        user_id: Optional[int],
        position_level: str,
        job_title: str,
        shift: str = "NOT_ASSIGNED",
        parent_id: Optional[int] = None,
        group_name: str = "",
        team_name: str = "",
        hire_date: str = "",
        notes: str = "",
        full_name: Optional[str] = None,
        matricule: Optional[str] = None,
        changed_by: str = "Admin"
    ) -> Tuple[bool, str, Optional[int]]:
        """
        Assigns an employee to the organizational chart.
        Supports selecting an existing user account OR creating/assigning an employee directly.
        Permits custom name and matricule editing independent of web authentication privileges.
        """
        position_level = position_level.upper().strip()
        if position_level not in POSITION_LEVELS:
            return False, f"Niveau de poste invalide: {position_level}", None

        shift = shift.upper().strip() if shift else "NOT_ASSIGNED"
        if shift not in VALID_SHIFTS:
            shift = "NOT_ASSIGNED"

        with self._get_conn() as conn:
            cur = conn.cursor()

            # Handle direct employee creation (without prior user account)
            if not user_id or int(user_id) <= 0:
                if not full_name or not full_name.strip():
                    return False, "Le nom et prénom de l'employé sont obligatoires.", None
                full_name = full_name.strip()
                import time, uuid
                matricule = matricule.strip() if matricule and matricule.strip() else f"EMP-{int(time.time()) % 100000}"
                unique_uname = f"emp_{matricule.lower().replace(' ', '_').replace('-', '_')}_{uuid.uuid4().hex[:4]}"
                cur.execute("""
                    INSERT INTO users (name, matricule, username, password_hash, role, shift, is_active, hire_date)
                    VALUES (?, ?, ?, 'INACTIVE_ORG_ONLY', 'TECHNICIAN', ?, 1, ?)
                """, (full_name, matricule, unique_uname, shift if shift in ('A','B','C') else 'A', hire_date))
                user_id = cur.lastrowid
            else:
                # Existing user
                cur.execute("SELECT id, name, matricule, role, shift, hire_date FROM users WHERE id = ?", (user_id,))
                u = cur.fetchone()
                if not u:
                    return False, f"Utilisateur introuvable (ID: {user_id})", None

                user_data = dict(u)
                full_name = full_name.strip() if full_name and full_name.strip() else user_data["name"]
                matricule = matricule.strip() if matricule and matricule.strip() else (user_data["matricule"] or f"USER-{user_id}")

                if not hire_date:
                    hire_date = user_data.get("hire_date") or ""

            # Check if user already has an assignment
            cur.execute("SELECT * FROM org_assignments WHERE user_id = ?", (user_id,))
            existing = cur.fetchone()

            if existing:
                asgn_id = existing["id"]
                old_lvl = existing["position_level"]
                old_shift = existing["shift"]
                old_parent = existing["parent_id"]
                old_title = existing["job_title"]

                cur.execute("""
                    UPDATE org_assignments
                    SET matricule = ?, full_name = ?, position_level = ?, job_title = ?,
                        shift = ?, group_name = ?, team_name = ?, parent_id = ?,
                        hire_date = ?, notes = ?, is_active = 1, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                """, (matricule, full_name, position_level, job_title, shift, group_name, team_name, parent_id, hire_date, notes, asgn_id))

                action_type = "PROMOTION" if POSITION_LEVELS.get(position_level, {}).get("rank", 99) < POSITION_LEVELS.get(old_lvl, {}).get("rank", 99) else "MODIFICATION"
                if old_shift != shift:
                    action_type = "CHANGEMENT_SHIFT"

                cur.execute("""
                    INSERT INTO org_assignment_history
                    (assignment_id, user_id, matricule, action_type, old_position_level, new_position_level,
                     old_job_title, new_job_title, old_shift, new_shift, old_parent_id, new_parent_id, details, changed_by)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (asgn_id, user_id, matricule, action_type, old_lvl, position_level, old_title, job_title, old_shift, shift, old_parent, parent_id, f"Mise à jour: {job_title}", changed_by))

            else:
                cur.execute("""
                    INSERT INTO org_assignments
                    (user_id, matricule, full_name, position_level, job_title, shift, group_name, team_name, parent_id, hire_date, notes, is_active)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
                """, (user_id, matricule, full_name, position_level, job_title, shift, group_name, team_name, parent_id, hire_date, notes))
                asgn_id = cur.lastrowid

                cur.execute("""
                    INSERT INTO org_assignment_history
                    (assignment_id, user_id, matricule, action_type, new_position_level, new_job_title, new_shift, new_parent_id, details, changed_by)
                    VALUES (?, ?, ?, 'NOUVELLE_AFFECTATION', ?, ?, ?, ?, ?, ?)
                """, (asgn_id, user_id, matricule, position_level, job_title, shift, parent_id, "Ajout à l'organigramme", changed_by))

            # Sync name, matricule, hire_date and shift into users table
            user_updates = ["name = ?", "matricule = ?"]
            user_params = [full_name, matricule]
            if hire_date:
                user_updates.append("hire_date = ?")
                user_params.append(hire_date)
            if shift and shift in ("A", "B", "C"):
                user_updates.append("shift = ?")
                user_params.append(shift)

            user_params.append(user_id)
            cur.execute(f"UPDATE users SET {', '.join(user_updates)} WHERE id = ?", user_params)

            conn.commit()

        # Also sync technician JSON file if exists
        self._sync_technician_file(matricule, shift, hire_date, name=full_name)

        return True, f"Employé '{full_name}' affecté avec succès au poste '{job_title}'.", asgn_id

    def update_assignment(
        self,
        assignment_id: int,
        position_level: str,
        job_title: str,
        shift: str,
        parent_id: Optional[int] = None,
        group_name: str = "",
        team_name: str = "",
        hire_date: str = "",
        notes: str = "",
        full_name: Optional[str] = None,
        matricule: Optional[str] = None,
        changed_by: str = "Admin"
    ) -> Tuple[bool, str]:
        """
        Updates an existing employee's organizational assignment, position, shift, superior, hire date,
        and allows updating their full name and matricule directly.
        """
        position_level = position_level.upper().strip()
        if position_level not in POSITION_LEVELS:
            return False, f"Niveau de poste invalide: {position_level}"

        shift = shift.upper().strip() if shift else "NOT_ASSIGNED"
        if shift not in VALID_SHIFTS:
            shift = "NOT_ASSIGNED"

        # Prevent circular parent reference
        if parent_id and parent_id == assignment_id:
            return False, "Un employé ne peut pas être son propre supérieur hiérarchique."

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM org_assignments WHERE id = ?", (assignment_id,))
            existing = cur.fetchone()
            if not existing:
                return False, "Affectation introuvable."

            user_id = existing["user_id"]
            old_mat = existing["matricule"]
            old_name = existing["full_name"]
            old_lvl = existing["position_level"]
            old_shift = existing["shift"]
            old_parent = existing["parent_id"]
            old_title = existing["job_title"]
            old_hire_date = existing["hire_date"]

            # Use new name/matricule if provided and valid, otherwise keep existing
            new_name = full_name.strip() if full_name and full_name.strip() else old_name
            new_mat = matricule.strip() if matricule and matricule.strip() else old_mat

            # Detect primary action type for audit history
            details_list = []
            action_type = "MODIFICATION"

            if old_name != new_name:
                details_list.append(f"Nom: '{old_name}' -> '{new_name}'")

            if old_mat != new_mat:
                details_list.append(f"Matricule: '{old_mat}' -> '{new_mat}'")

            if old_lvl != position_level:
                old_rank = POSITION_LEVELS.get(old_lvl, {}).get("rank", 99)
                new_rank = POSITION_LEVELS.get(position_level, {}).get("rank", 99)
                action_type = "PROMOTION" if new_rank < old_rank else "CHANGEMENT_POSTE"
                details_list.append(f"Poste: {old_lvl} -> {position_level}")

            if old_shift != shift:
                action_type = "CHANGEMENT_SHIFT"
                details_list.append(f"Shift: {old_shift} -> {shift}")

            if old_parent != parent_id:
                action_type = "CHANGEMENT_SUPERVISEUR"
                details_list.append(f"Supérieur modifié")

            if old_hire_date != hire_date:
                details_list.append(f"Date d'embauche: {hire_date}")

            cur.execute("""
                UPDATE org_assignments
                SET full_name = ?, matricule = ?, position_level = ?, job_title = ?, shift = ?, group_name = ?,
                    team_name = ?, parent_id = ?, hire_date = ?, notes = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (new_name, new_mat, position_level, job_title, shift, group_name, team_name, parent_id, hire_date, notes, assignment_id))

            # Audit History
            cur.execute("""
                INSERT INTO org_assignment_history
                (assignment_id, user_id, matricule, action_type, old_position_level, new_position_level,
                 old_job_title, new_job_title, old_shift, new_shift, old_parent_id, new_parent_id, details, changed_by)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (assignment_id, user_id, new_mat, action_type, old_lvl, position_level, old_title, job_title,
                  old_shift, shift, old_parent, parent_id, "; ".join(details_list) or "Mise à jour des informations", changed_by))

            # Sync name, matricule, hire_date and shift to users table (without altering authentication power)
            user_updates = []
            user_params = []
            if old_name != new_name:
                user_updates.append("name = ?")
                user_params.append(new_name)
            if old_mat != new_mat:
                user_updates.append("matricule = ?")
                user_params.append(new_mat)
            if hire_date:
                user_updates.append("hire_date = ?")
                user_params.append(hire_date)
            if shift and shift in ("A", "B", "C"):
                user_updates.append("shift = ?")
                user_params.append(shift)

            if user_updates and user_id:
                user_params.append(user_id)
                try:
                    cur.execute(f"UPDATE users SET {', '.join(user_updates)} WHERE id = ?", user_params)
                except Exception as ue:
                    logger.warning(f"Could not update users table for user_id={user_id}: {ue}")

            conn.commit()

        self._sync_technician_file(new_mat, shift, hire_date, name=new_name, old_matricule=old_mat)

        return True, f"Modifications enregistrées avec succès pour '{new_name}'."

    def move_employee(
        self,
        assignment_id: int,
        new_shift: Optional[str] = None,
        new_parent_id: Optional[int] = None,
        changed_by: str = "Admin"
    ) -> Tuple[bool, str]:
        """
        Quick action to move an employee from one shift to another (e.g. Shift A -> Shift B)
        or reassign their direct supervisor with one click.
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM org_assignments WHERE id = ?", (assignment_id,))
            existing = cur.fetchone()
            if not existing:
                return False, "Affectation introuvable."

            user_id = existing["user_id"]
            matricule = existing["matricule"]
            full_name = existing["full_name"]
            old_shift = existing["shift"]
            old_parent = existing["parent_id"]

            shift_to_set = old_shift
            if new_shift is not None:
                new_shift = new_shift.upper().strip()
                if new_shift in VALID_SHIFTS:
                    shift_to_set = new_shift

            parent_to_set = old_parent
            if new_parent_id is not None:
                if new_parent_id == assignment_id:
                    return False, "Un employé ne peut pas être son propre supérieur."
                parent_to_set = new_parent_id if new_parent_id > 0 else None

            action_type = "CHANGEMENT_SHIFT" if shift_to_set != old_shift else "CHANGEMENT_SUPERVISEUR"
            details = f"Transfert: Shift {old_shift} -> {shift_to_set}"

            cur.execute("""
                UPDATE org_assignments
                SET shift = ?, parent_id = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (shift_to_set, parent_to_set, assignment_id))

            cur.execute("""
                INSERT INTO org_assignment_history
                (assignment_id, user_id, matricule, action_type, old_shift, new_shift, old_parent_id, new_parent_id, details, changed_by)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (assignment_id, user_id, matricule, action_type, old_shift, shift_to_set, old_parent, parent_to_set, details, changed_by))

            # Sync shift to users table if valid
            if shift_to_set in ("A", "B", "C"):
                cur.execute("UPDATE users SET shift = ? WHERE id = ?", (shift_to_set, user_id))

            conn.commit()

        self._sync_technician_file(matricule, shift_to_set, None)

        return True, f"'{full_name}' déplacé avec succès vers Shift {shift_to_set}."

    def remove_from_org(self, assignment_id: int, changed_by: str = "Admin") -> Tuple[bool, str]:
        """
        Removes an employee from the organizational chart WITHOUT deleting their user account.
        Preserves all historical attendance, interventions, exams, and accounts.
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM org_assignments WHERE id = ?", (assignment_id,))
            existing = cur.fetchone()
            if not existing:
                return False, "Affectation introuvable."

            user_id = existing["user_id"]
            matricule = existing["matricule"]
            full_name = existing["full_name"]

            # Reassign any direct subordinates of this employee to the parent of this employee
            fallback_parent = existing["parent_id"]
            cur.execute("UPDATE org_assignments SET parent_id = ? WHERE parent_id = ?", (fallback_parent, assignment_id))

            # Log removal in history
            cur.execute("""
                INSERT INTO org_assignment_history
                (assignment_id, user_id, matricule, action_type, old_position_level, old_job_title, old_shift, details, changed_by)
                VALUES (?, ?, ?, 'RETRAIT_ORGANIGRAMME', ?, ?, ?, ?, ?)
            """, (assignment_id, user_id, matricule, existing["position_level"], existing["job_title"], existing["shift"], "Retiré de l'organigramme (compte utilisateur préservé)", changed_by))

            # Delete the assignment row from org_assignments
            cur.execute("DELETE FROM org_assignments WHERE id = ?", (assignment_id,))
            conn.commit()

        return True, f"'{full_name}' a été retiré de l'organigramme (son compte et son historique ont été préservés)."

    def update_hire_date(self, user_id: int, hire_date_str: str, changed_by: str = "Admin") -> Tuple[bool, str, Dict[str, Any]]:
        """
        Updates hiring date for an employee across both org_assignments and users tables,
        and recalculates seniority.
        """
        hire_date_str = str(hire_date_str).strip()[:10]
        seniority = calculate_seniority(hire_date_str)

        with self._get_conn() as conn:
            cur = conn.cursor()

            # Update users table
            cur.execute("UPDATE users SET hire_date = ? WHERE id = ?", (hire_date_str, user_id))

            # Update org_assignments if present
            cur.execute("UPDATE org_assignments SET hire_date = ?, updated_at = CURRENT_TIMESTAMP WHERE user_id = ?", (hire_date_str, user_id))

            # Get user info for log
            cur.execute("SELECT matricule, name FROM users WHERE id = ?", (user_id,))
            u = cur.fetchone()
            if u:
                cur.execute("""
                    INSERT INTO org_assignment_history
                    (assignment_id, user_id, matricule, action_type, details, changed_by)
                    VALUES (NULL, ?, ?, 'MODIFICATION_DATE_EMBAUCHE', ?, ?)
                """, (user_id, u["matricule"], f"Nouvelle date d'embauche: {hire_date_str} ({seniority['text']})", changed_by))

            conn.commit()

        return True, f"Date d'embauche mise à jour : {hire_date_str} (Ancienneté: {seniority['text']})", seniority

    def get_assignment_history(self, user_id: Optional[int] = None, assignment_id: Optional[int] = None, limit: int = 50) -> List[Dict[str, Any]]:
        """Returns chronological assignment audit trail."""
        query = """
            SELECT 
                h.*,
                u.name as employee_name,
                u.photo as employee_photo
            FROM org_assignment_history h
            LEFT JOIN users u ON h.user_id = u.id
            WHERE 1=1
        """
        params = []
        if user_id:
            query += " AND h.user_id = ?"
            params.append(user_id)
        elif assignment_id:
            query += " AND h.assignment_id = ?"
            params.append(assignment_id)

        query += " ORDER BY h.created_at DESC, h.id DESC LIMIT ?"
        params.append(limit)

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            return [dict(r) for r in cur.fetchall()]

    def get_groups_and_teams(self) -> Dict[str, List[str]]:
        """Returns existing distinct group and team names for autocomplete."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT DISTINCT group_name FROM org_assignments WHERE group_name != '' AND group_name IS NOT NULL")
            groups = [r["group_name"] for r in cur.fetchall()]

            cur.execute("SELECT DISTINCT team_name FROM org_assignments WHERE team_name != '' AND team_name IS NOT NULL")
            teams = [r["team_name"] for r in cur.fetchall()]

        return {"groups": groups, "teams": teams}

    def _sync_technician_file(self, matricule: str, shift: Optional[str], hire_date: Optional[str], name: Optional[str] = None, old_matricule: Optional[str] = None):
        """Helper to sync name, matricule, shift and hire_date into active_base/Technicians/{matricule}.json."""
        if not self.data_dir or not matricule:
            return
        try:
            import os
            import json
            safe_new = str(matricule).replace("/", "_").replace("\\", "_").strip()
            safe_old = str(old_matricule).replace("/", "_").replace("\\", "_").strip() if old_matricule else safe_new
            old_path = os.path.join(self.data_dir, "Technicians", f"{safe_old}.json")
            new_path = os.path.join(self.data_dir, "Technicians", f"{safe_new}.json")

            target_path = old_path if os.path.exists(old_path) else (new_path if os.path.exists(new_path) else None)
            if target_path:
                with open(target_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                updated = False
                if name and name.strip():
                    data["name"] = name.strip()
                    updated = True
                if matricule and data.get("matricule") != matricule:
                    data["matricule"] = matricule
                    updated = True
                if shift and shift in ("A", "B", "C"):
                    data["shift"] = shift
                    updated = True
                if hire_date:
                    data["hire_date"] = hire_date
                    updated = True
                if updated:
                    with open(new_path, "w", encoding="utf-8") as f:
                        json.dump(data, f, indent=4, ensure_ascii=False)
                    if safe_old != safe_new and os.path.exists(old_path) and old_path != new_path:
                        try:
                            os.remove(old_path)
                        except Exception:
                            pass
        except Exception as e:
            logger.warning(f"Could not sync technician JSON file for {matricule}: {e}")
