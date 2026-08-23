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
                SELECT id, name, matricule, username, role, shift, is_active, created_at, updated_at, last_login 
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

    def create_user(
        self,
        name: str,
        username: str,
        password: str,
        role: str = 'TECHNICIAN',
        matricule: str = '',
        shift: str = 'A',
        is_active: bool = True
    ) -> Tuple[bool, str, Optional[int]]:
        """Creates a new user with hashed password."""
        name = name.strip()
        username = username.strip()
        matricule = matricule.strip()
        role = role.strip().upper()
        if role not in ('OWNER', 'ADMIN', 'TECHNICIAN', 'USER'):
            role = 'TECHNICIAN'

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
                        nom, prenom, name, matricule, username, password_hash, role, shift, is_active, status, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                """, (nom, prenom, name, matricule or '', username, pwd_hash, role, shift, active_int, st_text))
            else:
                cur.execute("""
                    INSERT INTO users (
                        name, matricule, username, password_hash, role, shift, is_active, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                """, (name, matricule or '', username, pwd_hash, role, shift, active_int))

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
        is_active: bool = True
    ) -> Tuple[bool, str]:
        """Updates user profile information."""
        name = name.strip()
        username = username.strip()
        matricule = matricule.strip()
        role = role.strip().upper()
        if role not in ('OWNER', 'ADMIN', 'TECHNICIAN', 'USER'):
            role = 'TECHNICIAN'

        if not name or not username:
            return False, "Le nom et le nom d'utilisateur sont obligatoires."

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, role FROM users WHERE id = ?", (user_id,))
            target = cur.fetchone()
            if not target:
                return False, "Utilisateur introuvable."

            # Check duplicate username
            cur.execute("SELECT id FROM users WHERE LOWER(username) = LOWER(?) AND id != ?", (username, user_id))
            if cur.fetchone():
                return False, f"Le nom d'utilisateur '{username}' est déjà pris."

            if matricule:
                cur.execute("SELECT id FROM users WHERE LOWER(matricule) = LOWER(?) AND id != ?", (matricule, user_id))
                if cur.fetchone():
                    return False, f"Le matricule '{matricule}' est déjà utilisé."

            cur.execute("""
                UPDATE users 
                SET name = ?, username = ?, matricule = ?, role = ?, shift = ?, is_active = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (
                name,
                username,
                matricule or None,
                role,
                shift or 'A',
                1 if is_active else 0,
                user_id
            ))
            conn.commit()
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
