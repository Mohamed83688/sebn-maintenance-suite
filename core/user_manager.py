import sqlite3
import os
import time
import json
import logging
from typing import Optional, List, Dict, Any, Tuple
from werkzeug.security import generate_password_hash, check_password_hash

logger = logging.getLogger('sebn-maintenance')

class UserManager:
    """
    Unified User & Technician Authentication & Account Management for SEBN-TN Maintenance System.
    Manages all 4 roles: OWNER, ADMIN, TECHNICIAN, USER in a single centralized SQLite table.
    """
    def __init__(self, db_path: str, data_dir: str):
        self.db_path = db_path
        self.data_dir = data_dir
        self._init_db()
        self._seed_default_accounts()
        self._migrate_existing_technicians()

    def _get_conn(self):
        try:
            from core.db_mysql import get_db_connection
            return get_db_connection(sqlite_fallback_path=self.db_path)
        except Exception:
            conn = sqlite3.connect(self.db_path, timeout=15)
            conn.row_factory = sqlite3.Row
            return conn

    def _init_db(self):
        os.makedirs(os.path.dirname(os.path.abspath(self.db_path)), exist_ok=True)
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    matricule TEXT,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL DEFAULT 'TECHNICIAN',
                    shift TEXT DEFAULT 'A',
                    is_active INTEGER DEFAULT 1,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    last_login DATETIME
                )
            """)

            # Check existing columns and add missing ones if upgrading existing DB
            cur.execute("PRAGMA table_info(users)")
            existing_cols = {row["name"].lower() for row in cur.fetchall()}

            if "name" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN name TEXT DEFAULT ''")
                # If nom/prenom exist, populate name
                if "nom" in existing_cols:
                    cur.execute("UPDATE users SET name = TRIM(COALESCE(prenom, '') || ' ' || COALESCE(nom, '')) WHERE name = '' OR name IS NULL")

            if "matricule" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN matricule TEXT")

            if "is_active" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN is_active INTEGER DEFAULT 1")
                if "status" in existing_cols:
                    cur.execute("UPDATE users SET is_active = CASE WHEN status = 'ACTIVE' THEN 1 ELSE 0 END")

            if "shift" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN shift TEXT DEFAULT 'A'")

            if "last_login" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN last_login DATETIME")

            if "technician_level" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN technician_level TEXT")

            if "photo" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN photo TEXT")

            if "hire_date" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN hire_date TEXT")

            if "phone" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN phone TEXT DEFAULT ''")

            if "address" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN address TEXT DEFAULT ''")

            if "education_level" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN education_level TEXT DEFAULT ''")

            if "diploma" not in existing_cols:
                cur.execute("ALTER TABLE users ADD COLUMN diploma TEXT DEFAULT ''")

            cur.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_users_matricule ON users(matricule)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_users_role ON users(role)")
            conn.commit()

    def _seed_default_accounts(self):
        """Ensures default Owner and Admin accounts exist with secure hashes."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(users)")
            cols = {row["name"].lower() for row in cur.fetchall()}

            # --- Normalize any legacy lowercase roles to uppercase ---
            cur.execute("UPDATE users SET role = UPPER(role) WHERE role != UPPER(role)")
            cur.execute("UPDATE users SET technician_level = 'Level 1' WHERE role = 'TECHNICIAN' AND technician_level LIKE '%25%'")
            cur.execute("UPDATE users SET technician_level = 'Level 0' WHERE role = 'TECHNICIAN' AND (technician_level IS NULL OR technician_level = '')")

            # 1. Owner account
            cur.execute("SELECT id, role FROM users WHERE username = 'owner'")
            owner_row = cur.fetchone()
            if not owner_row:
                if "nom" in cols and "prenom" in cols:
                    cur.execute("""
                        INSERT INTO users (nom, prenom, name, matricule, username, password_hash, role, shift, is_active, status)
                        VALUES (?, ?, ?, ?, ?, ?, 'OWNER', 'ALL', 1, 'ACTIVE')
                    """, (
                        "Issaoui", "Mohamed",
                        "Mohamed Issaoui (Propriétaire)",
                        "OWNER-01",
                        "owner",
                        generate_password_hash("Owner@SEBN2026!")
                    ))
                else:
                    cur.execute("""
                        INSERT INTO users (name, matricule, username, password_hash, role, shift, is_active)
                        VALUES (?, ?, ?, ?, 'OWNER', 'ALL', 1)
                    """, (
                        "Mohamed Issaoui (Propriétaire)",
                        "OWNER-01",
                        "owner",
                        generate_password_hash("Owner@SEBN2026!")
                    ))
                logger.info("Created default OWNER account: 'owner'")
            elif owner_row["role"] != "OWNER":
                cur.execute("UPDATE users SET role = 'OWNER' WHERE id = ?", (owner_row["id"],))

            # 2. Admin account — only ensure it exists and has correct role;
            #    do NOT overwrite an existing password hash.
            cur.execute("SELECT id, role FROM users WHERE username = 'admin'")
            admin_row = cur.fetchone()
            if not admin_row:
                if "nom" in cols and "prenom" in cols:
                    cur.execute("""
                        INSERT INTO users (nom, prenom, name, matricule, username, password_hash, role, shift, is_active, status)
                        VALUES (?, ?, ?, ?, ?, ?, 'ADMIN', 'ALL', 1, 'ACTIVE')
                    """, (
                        "SEBN-TN", "Administrateur",
                        "Administrateur SEBN-TN",
                        "ADMIN-01",
                        "admin",
                        generate_password_hash("admin2026")
                    ))
                else:
                    cur.execute("""
                        INSERT INTO users (name, matricule, username, password_hash, role, shift, is_active)
                        VALUES (?, ?, ?, ?, 'ADMIN', 'ALL', 1)
                    """, (
                        "Administrateur SEBN-TN",
                        "ADMIN-01",
                        "admin",
                        generate_password_hash("admin2026")
                    ))
                logger.info("Created default ADMIN account: 'admin'")
            elif admin_row["role"] != "ADMIN":
                cur.execute("UPDATE users SET role = 'ADMIN' WHERE id = ?", (admin_row["id"],))

            conn.commit()



    def _migrate_existing_technicians(self):
        """Imports existing technicians from JSON files into the unified users table if not already present."""
        tech_dir = os.path.join(self.data_dir, "Technicians")
        if not os.path.isdir(tech_dir):
            return

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(users)")
            cols = {row["name"].lower() for row in cur.fetchall()}

            for fname in os.listdir(tech_dir):
                if fname.endswith(".json"):
                    try:
                        p = os.path.join(tech_dir, fname)
                        with open(p, 'r', encoding='utf-8') as f:
                            profile = json.load(f)
                        
                        mat = str(profile.get("matricule", "")).strip()
                        name = str(profile.get("name", "")).strip()
                        shift = str(profile.get("shift", "A")).strip() or "A"
                        if not mat or not name:
                            continue

                        parts = name.split()
                        prenom = parts[0]
                        nom = " ".join(parts[1:]) if len(parts) > 1 else parts[0]

                        # Generate clean username: e.g. "mohamed.issaoui" or "mat123"
                        clean_name = "".join(c if c.isalnum() else "." for c in name.lower()).strip(".")
                        username = clean_name if clean_name else f"tech.{mat.lower()}"
                        
                        cur.execute("SELECT id FROM users WHERE matricule = ? OR username = ?", (mat, username))
                        if not cur.fetchone():
                            if "nom" in cols and "prenom" in cols:
                                cur.execute("""
                                    INSERT INTO users (nom, prenom, name, matricule, username, password_hash, role, shift, is_active, status)
                                    VALUES (?, ?, ?, ?, ?, ?, 'TECHNICIAN', ?, 1, 'ACTIVE')
                                """, (
                                    nom, prenom, name, mat, username,
                                    generate_password_hash(f"Tech@{mat}!"), shift
                                ))
                            else:
                                cur.execute("""
                                    INSERT INTO users (name, matricule, username, password_hash, role, shift, is_active)
                                    VALUES (?, ?, ?, ?, 'TECHNICIAN', ?, 1)
                                """, (
                                    name, mat, username,
                                    generate_password_hash(f"Tech@{mat}!"), shift
                                ))
                            logger.info(f"Migrated technician {name} ({mat}) -> username '{username}'")
                    except Exception as e:
                        logger.error(f"Error migrating tech file {fname}: {e}")
            conn.commit()

    # ── Authenticate ─────────────────────────────────────────────────────────

    def authenticate(self, username: str, password: str) -> Optional[Dict[str, Any]]:
        """
        Authenticates a user against the unified users table using secure password hashing.
        Returns user dict if valid and active, else None.
        """
        if not username or not password:
            return None

        u_clean = username.strip()
        with self._get_conn() as conn:
            cur = conn.cursor()
            # Check by username or by matricule
            cur.execute("""
                SELECT * FROM users 
                WHERE (LOWER(username) = LOWER(?) OR LOWER(matricule) = LOWER(?)) AND is_active = 1
            """, (u_clean, u_clean))
            row = cur.fetchone()
            if not row:
                return None

            user_dict = dict(row)
            if check_password_hash(user_dict["password_hash"], password):
                # Update last login
                cur.execute("UPDATE users SET last_login = CURRENT_TIMESTAMP WHERE id = ?", (user_dict["id"],))
                conn.commit()
                return user_dict
            
        return None

    # ── User CRUD (Owner-Only) ───────────────────────────────────────────────

    def get_all_users(self) -> List[Dict[str, Any]]:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, name, matricule, username, role, shift, is_active, created_at, updated_at, last_login, technician_level, photo, hire_date, phone, address, education_level, diploma 
                FROM users 
                ORDER BY 
                    CASE role 
                        WHEN 'OWNER' THEN 1 
                        WHEN 'ADMIN' THEN 2 
                        WHEN 'TECHNICIAN' THEN 3 
                        ELSE 4 
                    END, name ASC
            """)
            return [dict(r) for r in cur.fetchall()]

    def get_user_by_id(self, user_id: int) -> Optional[Dict[str, Any]]:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM users WHERE id = ?", (user_id,))
            row = cur.fetchone()
            return dict(row) if row else None

    def get_user_by_username(self, username: str) -> Optional[Dict[str, Any]]:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM users WHERE LOWER(username) = LOWER(?)", (username.strip(),))
            row = cur.fetchone()
            return dict(row) if row else None

    @staticmethod
    def normalize_level(lvl: str) -> str:
        if not lvl:
            return 'Level 0'
        s = str(lvl).strip().upper()
        if '100' in s or 'LEVEL 4' in s or 'LEVEL4' in s or 'EXPERT' in s or s == '4':
            return 'Level 4'
        if '75' in s or 'LEVEL 3' in s or 'LEVEL3' in s or s == '3':
            return 'Level 3'
        if '50' in s or 'LEVEL 2' in s or 'LEVEL2' in s or s == '2':
            return 'Level 2'
        if '25' in s or 'LEVEL 1' in s or 'LEVEL1' in s or s == '1':
            return 'Level 1'
        if '0' in s or 'BASE' in s or 'DÉBUTANT' in s or 'DEBUTANT' in s:
            return 'Level 0'
        return 'Level 0'

    def create_user(
        self,
        name: str,
        username: str,
        password: str,
        role: str = 'TECHNICIAN',
        matricule: str = '',
        shift: str = 'A',
        is_active: bool = True,
        technician_level: str = None,
        photo: str = None,
        hire_date: str = None
    ) -> Tuple[bool, str, Optional[int]]:
        """Creates a new user with hashed password."""
        name = name.strip()
        username = username.strip()
        matricule = matricule.strip()
        role = role.strip().upper()
        if role not in ('OWNER', 'ADMIN', 'TECHNICIAN', 'USER'):
            role = 'TECHNICIAN'

        if role == 'TECHNICIAN':
            technician_level = self.normalize_level(technician_level) if technician_level else 'Level 0'
        else:
            technician_level = None

        if not name or not username or not password:
            return False, "Le nom complet, le nom d'utilisateur et le mot de passe sont obligatoires.", None

        if len(password) < 6:
            return False, "Le mot de passe doit comporter au moins 6 caractères.", None

        parts = name.split()
        prenom = parts[0]
        nom = " ".join(parts[1:]) if len(parts) > 1 else parts[0]

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(users)")
            cols = {row["name"].lower() for row in cur.fetchall()}

            # Check uniqueness
            cur.execute("SELECT id FROM users WHERE LOWER(username) = LOWER(?)", (username,))
            if cur.fetchone():
                return False, f"Le nom d'utilisateur '{username}' est déjà utilisé.", None

            if matricule:
                cur.execute("SELECT id FROM users WHERE LOWER(matricule) = LOWER(?)", (matricule,))
                if cur.fetchone():
                    return False, f"Le matricule '{matricule}' est déjà attribué à un autre compte.", None

            pwd_hash = generate_password_hash(password)
            active_int = 1 if is_active else 0
            st_text = 'ACTIVE' if is_active else 'INACTIVE'

            if "nom" in cols and "prenom" in cols:
                cur.execute("""
                    INSERT INTO users (
                        nom, prenom, name, matricule, username, password_hash, role, shift, is_active, status, technician_level, photo, hire_date, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                """, (nom, prenom, name, matricule or '', username, pwd_hash, role, shift, active_int, st_text, technician_level, photo, hire_date or ''))
            else:
                cur.execute("""
                    INSERT INTO users (
                        name, matricule, username, password_hash, role, shift, is_active, technician_level, photo, hire_date, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                """, (name, matricule or '', username, pwd_hash, role, shift, active_int, technician_level, photo, hire_date or ''))

            conn.commit()
            new_id = cur.lastrowid

            # If technician, also update/create JSON profile in Technicians/ for sync
            if role == 'TECHNICIAN' and matricule:
                tech_dir = os.path.join(self.data_dir, "Technicians")
                os.makedirs(tech_dir, exist_ok=True)
                safe_mat = str(matricule).replace("/", "_").replace("\\", "_").strip()
                p = os.path.join(tech_dir, f"{safe_mat}.json")
                profile = {
                    "name": name,
                    "matricule": matricule,
                    "shift": shift,
                    "hire_date": "",
                    "technician_level": technician_level,
                    "exams": []
                }
                if not os.path.exists(p):
                    with open(p, 'w', encoding='utf-8') as f:
                        json.dump(profile, f, indent=2, ensure_ascii=False)

            return True, f"Compte {role} '{username}' créé avec succès.", new_id

    def update_user(
        self,
        user_id: int,
        name: str,
        username: str,
        role: str,
        matricule: str = '',
        shift: str = 'A',
        is_active: bool = True,
        technician_level: str = None,
        photo: str = None,
        hire_date: str = None,
        phone: str = None,
        address: str = None,
        education_level: str = None,
        diploma: str = None
    ) -> Tuple[bool, str]:
        """Updates user profile information."""
        name = name.strip()
        username = username.strip()
        matricule = matricule.strip()
        role = role.strip().upper()
        if role not in ('OWNER', 'ADMIN', 'TECHNICIAN', 'USER'):
            role = 'TECHNICIAN'

        if role == 'TECHNICIAN':
            technician_level = self.normalize_level(technician_level)
        else:
            technician_level = None

        if not name or not username:
            return False, "Le nom et le nom d'utilisateur sont obligatoires."

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, role, phone, address, education_level, diploma, hire_date, photo FROM users WHERE id = ?", (user_id,))
            target = cur.fetchone()
            if not target:
                return False, "Utilisateur introuvable."

            target_dict = dict(target) if hasattr(target, 'keys') else {}

            # Check duplicate username
            cur.execute("SELECT id FROM users WHERE LOWER(username) = LOWER(?) AND id != ?", (username, user_id))
            if cur.fetchone():
                return False, f"Le nom d'utilisateur '{username}' est déjà pris."

            if matricule:
                cur.execute("SELECT id FROM users WHERE LOWER(matricule) = LOWER(?) AND id != ?", (matricule, user_id))
                if cur.fetchone():
                    return False, f"Le matricule '{matricule}' est déjà utilisé."

            # If photo is not provided, preserve the existing photo
            if photo is None:
                photo = target_dict.get("photo")

            final_hire_date = hire_date if hire_date is not None else target_dict.get("hire_date")
            final_phone = phone if phone is not None else target_dict.get("phone", "")
            final_address = address if address is not None else target_dict.get("address", "")
            final_edu = education_level if education_level is not None else target_dict.get("education_level", "")
            final_diploma = diploma if diploma is not None else target_dict.get("diploma", "")

            cur.execute("""
                UPDATE users 
                SET name = ?, username = ?, matricule = ?, role = ?, shift = ?, is_active = ?, 
                    technician_level = ?, photo = ?, hire_date = ?, phone = ?, address = ?, 
                    education_level = ?, diploma = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (
                name,
                username,
                matricule or None,
                role,
                shift or 'A',
                1 if is_active else 0,
                technician_level,
                photo,
                final_hire_date,
                final_phone,
                final_address,
                final_edu,
                final_diploma,
                user_id
            ))
            conn.commit()

            # Sync to JSON profile if technician
            if role == 'TECHNICIAN' and matricule:
                tech_dir = os.path.join(self.data_dir, "Technicians")
                os.makedirs(tech_dir, exist_ok=True)
                safe_mat = str(matricule).replace("/", "_").replace("\\", "_").strip()
                p = os.path.join(tech_dir, f"{safe_mat}.json")
                profile = {
                    "name": name,
                    "matricule": matricule,
                    "shift": shift,
                    "hire_date": hire_date or "",
                    "technician_level": technician_level,
                    "exams": []
                }
                # Preserve existing exams if JSON already exists
                if os.path.exists(p):
                    try:
                        with open(p, 'r', encoding='utf-8') as f:
                            old = json.load(f)
                            profile["exams"] = old.get("exams", [])
                            if hire_date is None:
                                profile["hire_date"] = old.get("hire_date", "")
                    except Exception:
                        pass
                with open(p, 'w', encoding='utf-8') as f:
                    json.dump(profile, f, indent=2, ensure_ascii=False)

            return True, f"Compte '{name}' mis à jour avec succès."

    def reset_password(self, user_id: int, new_password: str) -> Tuple[bool, str]:
        """Resets a user's password with a secure hash."""
        if not new_password or len(new_password) < 6:
            return False, "Le mot de passe doit contenir au moins 6 caractères."

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, name FROM users WHERE id = ?", (user_id,))
            row = cur.fetchone()
            if not row:
                return False, "Utilisateur introuvable."

            cur.execute("""
                UPDATE users 
                SET password_hash = ?, updated_at = CURRENT_TIMESTAMP 
                WHERE id = ?
            """, (generate_password_hash(new_password), user_id))
            conn.commit()
            return True, f"Mot de passe de '{row['name']}' réinitialisé avec succès."

    def toggle_user_status(self, user_id: int) -> Tuple[bool, str, int]:
        """Toggles user active/inactive status."""
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, name, is_active, role FROM users WHERE id = ?", (user_id,))
            row = cur.fetchone()
            if not row:
                return False, "Utilisateur introuvable.", 0

            if row["role"] == "OWNER":
                return False, "Le compte Propriétaire ne peut pas être désactivé.", row["is_active"]

            new_status = 0 if row["is_active"] == 1 else 1
            cur.execute("UPDATE users SET is_active = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_status, user_id))
            conn.commit()
            status_label = "activé" if new_status == 1 else "désactivé"
            return True, f"Le compte '{row['name']}' a été {status_label}.", new_status

    def delete_user(self, user_id: int) -> Tuple[bool, str]:
        """
        Permanently deletes a user account.
        - The OWNER account can never be deleted.
        - Cascades: also removes exam_attempts and org_assignments tied to that user.
        Returns (success, message).
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, name, role FROM users WHERE id = ?", (user_id,))
            row = cur.fetchone()
            if not row:
                return False, "Utilisateur introuvable."
            if row["role"] == "OWNER":
                return False, "Le compte Propriétaire ne peut pas être supprimé."
            name = row["name"]
            # Cascade: remove exam attempts
            try:
                cur.execute("DELETE FROM exam_attempts WHERE user_id = ?", (user_id,))
            except Exception:
                pass
            # Cascade: remove org assignments
            try:
                cur.execute("DELETE FROM org_assignments WHERE user_id = ?", (user_id,))
            except Exception:
                pass
            # Cascade: remove level history
            try:
                cur.execute("DELETE FROM technician_level_history WHERE user_id = ?", (user_id,))
            except Exception:
                pass
            # Delete the user
            cur.execute("DELETE FROM users WHERE id = ?", (user_id,))
            conn.commit()
        return True, f"Le compte de '{name}' a été supprimé définitivement."

    def get_technician_full_profile(self, user_id_or_matricule: Any) -> Optional[Dict[str, Any]]:
        """
        Loads the complete technician employee file (Fiche Technicien):
        - Personal Information: photo, name, matricule, phone, address
        - Professional Information: technician_level, hire_date, seniority (auto-calculated), shift, supervisor
        - Education & Qualifications: education_level, diploma
        - Exam/training history and progression
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            
            # Fetch user record by id or by matricule
            cur.execute("SELECT * FROM users WHERE LOWER(matricule) = LOWER(?)", (str(user_id_or_matricule).strip(),))
            row = cur.fetchone()
            if not row and (isinstance(user_id_or_matricule, int) or str(user_id_or_matricule).isdigit()):
                cur.execute("SELECT * FROM users WHERE id = ?", (int(user_id_or_matricule),))
                row = cur.fetchone()

            if not row:
                return None
                
            u = dict(row)
            user_id = u["id"]
            
            # Auto-calculate Seniority using calculate_seniority from core.org_manager
            from core.org_manager import calculate_seniority
            u["seniority"] = calculate_seniority(u.get("hire_date"))
            
            # Resolve direct supervisor from org_assignments (or by shift if unassigned)
            supervisor = {
                "name": "Non assigné",
                "title": "Superviseur Direct",
                "matricule": "—",
                "shift": u.get("shift") or "A"
            }
            try:
                cur.execute("""
                    SELECT oa.id as assignment_id, oa.job_title, oa.group_name, oa.team_name, oa.shift as org_shift,
                           parent.full_name as supervisor_name, parent.job_title as supervisor_job_title,
                           parent.matricule as supervisor_matricule, parent.shift as supervisor_shift
                    FROM org_assignments oa
                    LEFT JOIN org_assignments parent ON oa.parent_id = parent.id
                    WHERE oa.user_id = ?
                """, (user_id,))
                oa_row = cur.fetchone()
                if oa_row:
                    oa_dict = dict(oa_row)
                    u["org_info"] = oa_dict
                    if oa_dict.get("supervisor_name"):
                        supervisor = {
                            "name": oa_dict["supervisor_name"],
                            "title": oa_dict.get("supervisor_job_title") or "Supervisor",
                            "matricule": oa_dict.get("supervisor_matricule") or "—",
                            "shift": oa_dict.get("supervisor_shift") or u.get("shift") or "A"
                        }
                    else:
                        # Fallback: find supervisor of the same shift in org_assignments
                        cur.execute("""
                            SELECT full_name, job_title, matricule, shift
                            FROM org_assignments
                            WHERE position_level = 'SUPERVISOR' AND shift = ?
                            LIMIT 1
                        """, (u.get("shift") or "A",))
                        s_row = cur.fetchone()
                        if s_row:
                            s_dict = dict(s_row)
                            supervisor = {
                                "name": s_dict["full_name"],
                                "title": s_dict.get("job_title") or "Supervisor",
                                "matricule": s_dict.get("matricule") or "—",
                                "shift": s_dict.get("shift") or u.get("shift") or "A"
                            }
                else:
                    # Not in org_assignments, check shift supervisor
                    cur.execute("""
                        SELECT full_name, job_title, matricule, shift
                        FROM org_assignments
                        WHERE position_level = 'SUPERVISOR' AND shift = ?
                        LIMIT 1
                    """, (u.get("shift") or "A",))
                    s_row = cur.fetchone()
                    if s_row:
                        s_dict = dict(s_row)
                        supervisor = {
                            "name": s_dict["full_name"],
                            "title": s_dict.get("job_title") or "Supervisor",
                            "matricule": s_dict.get("matricule") or "—",
                            "shift": s_dict.get("shift") or u.get("shift") or "A"
                        }
            except Exception as e:
                logger.warning(f"Could not resolve supervisor for user {user_id}: {e}")
                
            u["supervisor"] = supervisor
            
            # Fetch exam history
            exams = []
            try:
                cur.execute("""
                    SELECT ea.*, e.title as exam_title, e.number_of_questions, e.passing_score
                    FROM exam_attempts ea
                    JOIN exams e ON ea.exam_id = e.id
                    WHERE ea.user_id = ?
                    ORDER BY ea.submitted_at DESC, ea.started_at DESC
                """, (user_id,))
                for er in cur.fetchall():
                    ed = dict(er)
                    is_pass = (ed.get("status") == "passed")
                    exams.append({
                        "name": ed.get("exam_title") or f"Examen #{ed.get('exam_id')}",
                        "type": "Validation" if is_pass else "Echec",
                        "status": "RÉUSSI" if is_pass else "ÉCHOUÉ",
                        "exam_level": u.get("technician_level") or "Level 1",
                        "date": (ed.get("submitted_at") or ed.get("started_at") or "")[:10],
                        "score": f"{ed.get('score', 0)}/{ed.get('number_of_questions', 0)}",
                        "percentage": round(ed.get("percentage") or 0, 1),
                        "notes": f"Score: {ed.get('score')}/{ed.get('number_of_questions')} ({round(ed.get('percentage') or 0, 1)}%)"
                    })
            except Exception as e:
                logger.warning(f"Could not load exam attempts for user {user_id}: {e}")
                
            # Also check tech_db JSON file if extra exams or validation events exist
            try:
                safe_mat = str(u.get("matricule") or "").replace("/", "_").replace("\\", "_").strip()
                p = os.path.join(self.data_dir, "Technicians", f"{safe_mat}.json")
                if os.path.exists(p):
                    with open(p, "r", encoding="utf-8") as f:
                        j_data = json.load(f)
                        for j_ex in j_data.get("exams", []):
                            # Add if not duplicate
                            if not any(e.get("name") == j_ex.get("name") and e.get("date") == j_ex.get("date") for e in exams):
                                exams.append(j_ex)
            except Exception:
                pass
                
            u["exams"] = exams
            
            # Calculate progression percentage (Level 0 = 0%, Level 1 = 25%, Level 2 = 50%, Level 3 = 75%, Level 4 = 100%)
            lvl_norm = self.normalize_level(u.get("technician_level"))
            lvl_map = {"Level 0": 0, "Level 1": 25, "Level 2": 50, "Level 3": 75, "Level 4": 100}
            u["progression_pct"] = lvl_map.get(lvl_norm, 0)
            u["technician_level"] = lvl_norm
            
            return u

    def update_technician_profile(
        self,
        user_id: int,
        phone: str = '',
        address: str = '',
        education_level: str = '',
        diploma: str = '',
        hire_date: Optional[str] = None,
        shift: Optional[str] = None,
        name: Optional[str] = None,
        matricule: Optional[str] = None,
        technician_level: Optional[str] = None,
        photo: Optional[str] = None
    ) -> Tuple[bool, str, Optional[Dict[str, Any]]]:
        """
        Updates technician employee profile fields in users table and syncs
        with org_assignments (if assigned) and tech_db JSON files.
        Recalculates seniority automatically based on recruitment date.
        """
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM users WHERE id = ?", (user_id,))
            target = cur.fetchone()
            if not target:
                return False, "Technicien introuvable.", None

            target = dict(target)
            new_name = (name.strip() if name else target.get("name")) or target.get("name")
            new_mat = (matricule.strip() if matricule else target.get("matricule")) or target.get("matricule")
            new_shift = (shift.strip() if shift else target.get("shift")) or target.get("shift") or "A"
            new_level = self.normalize_level(technician_level) if technician_level else target.get("technician_level") or "Level 0"
            new_photo = photo if photo is not None else target.get("photo")
            new_hire_date = hire_date.strip() if hire_date else target.get("hire_date")

            # Clean text fields
            new_phone = str(phone or "").strip()
            new_address = str(address or "").strip()
            new_edu = str(education_level or "").strip()
            new_diploma = str(diploma or "").strip()

            cur.execute("""
                UPDATE users
                SET name = ?,
                    matricule = ?,
                    shift = ?,
                    technician_level = ?,
                    photo = ?,
                    hire_date = ?,
                    phone = ?,
                    address = ?,
                    education_level = ?,
                    diploma = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (
                new_name,
                new_mat,
                new_shift,
                new_level,
                new_photo,
                new_hire_date,
                new_phone,
                new_address,
                new_edu,
                new_diploma,
                user_id
            ))

            # Also sync with org_assignments if assigned
            cur.execute("""
                UPDATE org_assignments
                SET full_name = ?,
                    matricule = ?,
                    shift = ?,
                    hire_date = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE user_id = ?
            """, (new_name, new_mat, new_shift, new_hire_date, user_id))

            conn.commit()

            # Also sync with Technicians/{matricule}.json
            try:
                safe_mat = str(new_mat or "").replace("/", "_").replace("\\", "_").strip()
                p = os.path.join(self.data_dir, "Technicians", f"{safe_mat}.json")
                t_data = {}
                if os.path.exists(p):
                    with open(p, "r", encoding="utf-8") as f:
                        t_data = json.load(f)
                t_data.update({
                    "name": new_name,
                    "matricule": new_mat,
                    "shift": new_shift,
                    "technician_level": new_level,
                    "hire_date": new_hire_date,
                    "phone": new_phone,
                    "address": new_address,
                    "education_level": new_edu,
                    "diploma": new_diploma
                })
                if "exams" not in t_data:
                    t_data["exams"] = []
                with open(p, "w", encoding="utf-8") as f:
                    json.dump(t_data, f, indent=2, ensure_ascii=False)
            except Exception as e:
                logger.warning(f"Failed to sync tech_db JSON: {e}")

            full_profile = self.get_technician_full_profile(user_id)
            return True, f"Fiche de {new_name} mise à jour avec succès.", full_profile
