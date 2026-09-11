import os
import sys
import datetime
import time
import uuid
import socket
import logging
import secrets
from functools import wraps
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_file, make_response

# Add project root to sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.config import ConfigManager
from core.owner_security import OwnerSecurityManager
from core.data_engine import DataEngine
from core.tech_db import TechnicianDatabase
from core.user_manager import UserManager
from core.exam_manager import ExamManager
from core.ebm_manager import EBMManager
from core.passation_manager import PassationManager
from core.document_manager import DocumentManager
from ima.config import IMAConfig
from ima.database import IMADatabase
from core.checklist_manager import ChecklistManager
from web_portal.utils.export_helper import export_interventions_to_excel
from web_portal.utils.email_service import send_simple_alert
from web_portal.utils.pdf_helper import export_dashboard_pdf
from web_portal.utils.vault_helper import get_vault_files
from web_portal.utils.checklist_parser import find_template, parse_tasks, save_filled_checklist, list_all_templates, GENERIC_TASKS
from web_portal.utils.manager_plan_helper import (
    load_meetings, save_meeting, update_meeting_status,
    load_actions, save_action, update_action_status, delete_action,
    get_action_kpis
)

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(name)s: %(message)s')
logger = logging.getLogger('sebn-maintenance')

# ── Flask app ─────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SEBN_SECRET_KEY", "SEBN-TN-UNIFIED-AUTH-V2-2026-R2")
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024  # 64 MB

# ── Session cookie: browser-session only (strictly no persistence) ───────────
app.config['SESSION_PERMANENT']          = False   # Never persist across browser close
app.config['PERMANENT_SESSION_LIFETIME'] = datetime.timedelta(minutes=60)
app.config['SESSION_COOKIE_HTTPONLY']    = True    # JS cannot read the cookie
app.config['SESSION_COOKIE_SAMESITE']   = 'Lax'   # CSRF protection
app.config['SESSION_COOKIE_NAME']       = 'session'
if os.environ.get('SEBN_HTTPS', '').lower() in ('1', 'true', 'yes'):
    app.config['SESSION_COOKIE_SECURE'] = True

app.config['TEMPLATES_AUTO_RELOAD']     = True
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
app.jinja_env.auto_reload = True

# ── Backends & Security ───────────────────────────────────────────────────────
pma_config = ConfigManager()
owner_security = OwnerSecurityManager(pma_config.active_base)
tech_db = TechnicianDatabase(pma_config.active_base)
ima_config = IMAConfig()

IMA_DB_PATH = os.path.join(ima_config.active_base, "IMA.db")
ima_db = IMADatabase(IMA_DB_PATH)

# Centralized User & Auth Manager (OWNER, ADMIN, TECHNICIAN, USER)
user_mgr = UserManager(db_path=IMA_DB_PATH, data_dir=pma_config.active_base)
exam_mgr = ExamManager(db_path=IMA_DB_PATH, data_dir=pma_config.active_base)
from web_portal.utils.pdf_card_generator import generate_technician_card_pdf

# EBM & Passation Managers (Owner Modules)
ebm_mgr = EBMManager(db_path=IMA_DB_PATH, data_dir=pma_config.active_base)
passation_mgr = PassationManager(db_path=IMA_DB_PATH, data_dir=pma_config.active_base)

# Document Management (all roles — admin can manage, all can view)
doc_mgr = DocumentManager(db_path=IMA_DB_PATH, data_dir=pma_config.active_base)

# Checklist & PPE Manager (SQLite + Versioning)
checklist_mgr = ChecklistManager(
    db_path=IMA_DB_PATH,
    storage_dir=pma_config.dirs.get('ppe', os.path.join(pma_config.active_base, 'PPE_Templates'))
)

# Manager Action Plan workbook path
MANAGER_PLAN_PATH = os.path.join(pma_config.active_base, "manager_plan_performance.xlsx")

# ── PMA DataEngine cache ──────────────────────────────────────────────────────
_DE_CACHE  = None
_DE_PATHS  = []
_DE_MTIMES = {}

def get_pma_engine(force=False):
    global _DE_CACHE, _DE_PATHS, _DE_MTIMES
    paths = pma_config.get_all_excel_paths()

    if not paths:
        if not _DE_CACHE:
            _DE_CACHE = DataEngine()
        return _DE_CACHE

    current_mtimes = {}
    for p in paths:
        try:
            current_mtimes[p] = os.path.getmtime(p)
        except Exception:
            current_mtimes[p] = 0.0

    needs_reload = force or (paths != _DE_PATHS) or (current_mtimes != _DE_MTIMES)

    if needs_reload:
        logger.info(f"Loading {len(paths)} PMA Excel file(s): {[os.path.basename(p) for p in paths]}")
        try:
            _DE_CACHE = DataEngine()
            _DE_CACHE.load_excel(paths if len(paths) > 1 else paths[0])
            _DE_PATHS  = paths
            _DE_MTIMES = current_mtimes
            try:
                sync_pma_machines_to_ima(_DE_CACHE.current_df)
            except Exception as e:
                logger.warning(f"Could not auto-sync machines: {e}")
        except Exception as e:
            logger.error(f"PMA engine error: {e}")
            if not _DE_CACHE:
                _DE_CACHE = DataEngine()

    return _DE_CACHE

def invalidate_pma_cache():
    global _DE_CACHE, _DE_PATHS, _DE_MTIMES
    _DE_CACHE  = None
    _DE_PATHS  = []
    _DE_MTIMES = {}

def sync_pma_machines_to_ima(df=None):
    """
    Synchronizes all unique machines from the active Preventive Calendar (PMA)
    into the Curative Maintenance (IMA) database so they immediately appear
    in intervention forms, machine lists, and analytics.
    """
    try:
        machine_map = {}

        # 1. Direct DataEngine DataFrame inspection (most accurate, 100% matches Preventive Calendar)
        try:
            target_df = df
            if target_df is None or (hasattr(target_df, 'empty') and target_df.empty):
                eng = get_pma_engine()
                target_df = eng.current_df

            if target_df is not None and not target_df.empty:
                col_equip = 'Equipment' if 'Equipment' in target_df.columns else ('equipment' if 'equipment' in target_df.columns else None)
                if col_equip:
                    for _, r in target_df.iterrows():
                        mid = str(r.get(col_equip, '')).strip()
                        if not mid or mid.lower() in ['nan', 'none', '']:
                            continue
                        mname = str(r.get('Machine_Name', mid)).strip()
                        grp = str(r.get('Sheet', 'Général')).strip()
                        machine_map[mid] = {
                            "machine_id": mid,
                            "machine_name": mname if mname and mname.lower() != 'nan' else mid,
                            "group_name": grp if grp and grp.lower() != 'nan' else "Général",
                            "location": "",
                            "description": ""
                        }
        except Exception as e:
            logger.warning(f"sync_pma: DataEngine parse step error: {e}")

        # 2. Check uploaded Excel schedule files via CalendrierReader to catch any additional machines
        try:
            from ima.excel_reader import CalendrierReader
            reader = CalendrierReader()
            paths = pma_config.get_all_excel_paths()

            for p in paths:
                try:
                    c_df = reader.read_calendrier(p)
                    if not c_df.empty:
                        for _, r in c_df.iterrows():
                            mid = str(r.get('ID Machine', '')).strip()
                            mname = str(r.get('Nom Machine', mid)).strip()
                            grp = str(r.get('Groupe', '')).strip()
                            if mid and mid.lower() not in ['nan', 'none', '']:
                                if mid not in machine_map:
                                    machine_map[mid] = {
                                        "machine_id": mid,
                                        "machine_name": mname if mname and mname.lower() != 'nan' else mid,
                                        "group_name": grp if grp and grp.lower() != 'nan' else "Général",
                                        "location": "",
                                        "description": ""
                                    }
                except Exception as ex:
                    logger.warning(f"sync_pma: CalendrierReader on {p} warning: {ex}")
        except Exception as e:
            logger.warning(f"sync_pma: CalendrierReader step error: {e}")

        if machine_map:
            c = ima_db.upsert_machines(list(machine_map.values()))
            logger.info(f"sync_pma: successfully synced {c} machines to IMA database")
            return c
        return 0
    except Exception as e:
        logger.error(f"sync_pma: FATAL ERROR: {e}")
        return 0

# (Machine sync happens on-demand per request, not at startup)


def replace_pma_machines_to_ima(df=None) -> tuple[int, int]:
    """Destructive sync: replaces the machine list to exactly match the active Excel.

    Machines in the new Excel   → kept / added / updated.
    Machines NOT in new Excel   → deleted from the machines table.
    Interventions               → NEVER touched, always preserved.

    Uses the same Excel-reading logic as sync_pma_machines_to_ima() but calls
    ima_db.replace_machines() instead of ima_db.upsert_machines().

    Safety: if the Excel yields zero machines (parse error / empty file) the
    function falls back to the additive sync and returns (0, 0) to avoid a wipe.

    Returns:
        (upserted_count, deleted_count)
    """
    try:
        machine_map = {}

        # 1. Direct DataEngine DataFrame inspection (most accurate)
        try:
            target_df = df
            if target_df is None or (hasattr(target_df, 'empty') and target_df.empty):
                eng = get_pma_engine()
                target_df = eng.current_df

            if target_df is not None and not target_df.empty:
                col_equip = ('Equipment' if 'Equipment' in target_df.columns
                             else ('equipment' if 'equipment' in target_df.columns else None))
                if col_equip:
                    for _, r in target_df.iterrows():
                        mid = str(r.get(col_equip, '')).strip()
                        if not mid or mid.lower() in ['nan', 'none', '']:
                            continue
                        mname = str(r.get('Machine_Name', mid)).strip()
                        grp   = str(r.get('Sheet', 'Général')).strip()
                        machine_map[mid] = {
                            "machine_id":   mid,
                            "machine_name": mname if mname and mname.lower() != 'nan' else mid,
                            "group_name":   grp   if grp   and grp.lower()   != 'nan' else "Général",
                            "location":     "",
                            "description":  ""
                        }
        except Exception as e:
            logger.warning(f"replace_pma: DataEngine parse step error: {e}")

        # 2. Also read via CalendrierReader to catch additional machines
        try:
            from ima.excel_reader import CalendrierReader
            reader = CalendrierReader()
            paths  = pma_config.get_all_excel_paths()
            for p in paths:
                try:
                    c_df = reader.read_calendrier(p)
                    if not c_df.empty:
                        for _, r in c_df.iterrows():
                            mid   = str(r.get('ID Machine', '')).strip()
                            mname = str(r.get('Nom Machine', mid)).strip()
                            grp   = str(r.get('Groupe', '')).strip()
                            if mid and mid.lower() not in ['nan', 'none', '']:
                                if mid not in machine_map:
                                    machine_map[mid] = {
                                        "machine_id":   mid,
                                        "machine_name": mname if mname and mname.lower() != 'nan' else mid,
                                        "group_name":   grp   if grp   and grp.lower()   != 'nan' else "Général",
                                        "location":     "",
                                        "description":  ""
                                    }
                except Exception as ex:
                    logger.warning(f"replace_pma: CalendrierReader on {p} warning: {ex}")
        except Exception as e:
            logger.warning(f"replace_pma: CalendrierReader step error: {e}")

        if not machine_map:
            # Yield gracefully — no machines parsed means we do NOT wipe the list
            logger.warning("replace_pma: zero machines found in Excel — skipping destructive sync to protect existing data")
            return (0, 0)

        upserted, deleted = ima_db.replace_machines(list(machine_map.values()))
        logger.info(f"replace_pma: {upserted} machines kept/updated, {deleted} machines removed from Parc Machines")
        return (upserted, deleted)

    except Exception as e:
        logger.error(f"replace_pma: FATAL ERROR: {e}")
        return (0, 0)



# ── Diagnostic endpoint (public, read-only) ────────────────────────────────
@app.route('/diag')
def diagnostic():
    import json as _json
    try:
        app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        active_base = pma_config.active_base
        app_data_dir = os.path.join(app_root, 'data')
        excel_paths = pma_config.get_all_excel_paths()
        machine_count = len(ima_db.get_all_machines())
        last_excel = pma_config.get_last_excel_path()

        # List files in active_base and app data dir
        def list_files(d):
            try:
                return sorted(os.listdir(d)) if os.path.isdir(d) else ['(not a directory)']
            except Exception as ex:
                return [f'ERROR: {ex}']

        info = {
            'app_root': app_root,
            'active_base': active_base,
            'app_data_dir': app_data_dir,
            'SEBN_DATA_DIR_env': os.environ.get('SEBN_DATA_DIR', '(not set)'),
            'IMA_DB_PATH': IMA_DB_PATH,
            'db_exists': os.path.isfile(IMA_DB_PATH),
            'last_excel_txt': last_excel,
            'excel_paths_found': excel_paths,
            'machine_count_in_db': machine_count,
            'active_base_files': list_files(active_base),
            'app_data_files': list_files(app_data_dir),
            'schedules_dir_files': list_files(os.path.join(active_base, 'Schedules')),
        }
        html = '<html><head><title>Diagnostic</title></head><body>'
        html += '<h2>SEBN-TN Diagnostic</h2><pre style="background:#eee;padding:20px;font-size:14px;">'
        html += _json.dumps(info, indent=2, ensure_ascii=False)
        html += '</pre></body></html>'
        return html, 200
    except Exception as ex:
        return f'<pre>DIAG ERROR: {ex}</pre>', 500

# ── Force sync endpoint (public, safe — read+write only) ─────────────────────
@app.route('/sync')
def force_sync():
    import json as _json
    try:
        log_lines = []

        # Directly look for the bundled Excel file relative to app.py
        app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        bundled_excel = os.path.join(app_root, 'data', 'current_schedule.xlsx')
        log_lines.append(f"App root: {app_root}")
        log_lines.append(f"Bundled Excel path: {bundled_excel}")
        log_lines.append(f"Bundled Excel exists: {os.path.isfile(bundled_excel)}")
        log_lines.append(f"active_base: {pma_config.active_base}")
        log_lines.append(f"IMA_DB_PATH: {IMA_DB_PATH}")
        log_lines.append(f"SEBN_DATA_DIR env: {os.environ.get('SEBN_DATA_DIR', '(not set)')}")

        machines_before = len(ima_db.get_all_machines())
        log_lines.append(f"Machines in DB before sync: {machines_before}")

        # Gather all Excel paths
        paths = pma_config.get_all_excel_paths()
        if os.path.isfile(bundled_excel) and bundled_excel not in paths:
            paths.append(bundled_excel)
        log_lines.append(f"Excel paths to try: {paths}")

        total = 0
        from ima.excel_reader import CalendrierReader
        reader = CalendrierReader()
        for p in paths:
            try:
                cdf = reader.read_calendrier(p)
                log_lines.append(f"  {os.path.basename(p)}: {len(cdf)} machines extracted")
                if not cdf.empty:
                    mlist = []
                    for _, r in cdf.iterrows():
                        mid = str(r.get('ID Machine', '')).strip()
                        mname = str(r.get('Nom Machine', mid)).strip()
                        grp = str(r.get('Groupe', '')).strip()
                        if mid and mid.lower() not in ['nan', 'none', '']:
                            mlist.append({
                                'machine_id': mid,
                                'machine_name': mname if mname and mname.lower() != 'nan' else mid,
                                'group_name': grp if grp and grp.lower() != 'nan' else 'Général',
                                'location': '', 'description': ''
                            })
                    if mlist:
                        c = ima_db.upsert_machines(mlist)
                        total += c
                        log_lines.append(f"  Synced {c} machines from {os.path.basename(p)}")
            except Exception as pe:
                log_lines.append(f"  ERROR reading {p}: {pe}")

        machines_after = len(ima_db.get_all_machines())
        log_lines.append(f"Machines in DB after sync: {machines_after}")

        color = '#d4edda' if machines_after > 0 else '#f8d7da'
        status = f"✅ SUCCESS: {machines_after} machines synced!" if machines_after > 0 else "❌ FAILED: 0 machines"
        html = f'<html><head><title>Force Sync</title></head><body style="font-family:sans-serif;padding:30px;">'
        html += f'<h2>{status}</h2>'
        html += f'<div style="background:{color};padding:15px;border-radius:8px;margin-bottom:20px;">'
        html += f'<strong>Synced {machines_after} machines total</strong>'
        html += f'</div>'
        html += f'<pre style="background:#eee;padding:20px;font-size:13px;">' + '\n'.join(log_lines) + '</pre>'
        html += f'<br><a href="/interventions/new" style="padding:10px 20px;background:#28a745;color:white;text-decoration:none;border-radius:5px;">→ Open Intervention Form</a>'
        html += '</body></html>'
        return html, 200
    except Exception as ex:
        return f'<pre>SYNC ERROR: {ex}</pre>', 500

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "localhost"

def calc_downtime(start_iso, end_iso):
    if not start_iso or not end_iso:
        return 0.0
    try:
        fmt = '%Y-%m-%dT%H:%M'
        start = datetime.datetime.strptime(start_iso[:16], fmt) if len(start_iso) >= 16 else datetime.datetime.fromisoformat(start_iso)
        end   = datetime.datetime.strptime(end_iso[:16], fmt)   if len(end_iso)   >= 16 else datetime.datetime.fromisoformat(end_iso)
        return max(0.0, (end - start).total_seconds() / 60.0)
    except Exception:
        return 0.0

# ── Session helpers ───────────────────────────────────────────────────────────
def completed_ppes():
    return set(session.get('completed_ppes', []))

def mark_ppe_done(sheet, raw_idx):
    ppes = session.get('completed_ppes', [])
    k = f"{sheet}_{raw_idx}"
    if k not in ppes:
        ppes.append(k)
        session['completed_ppes'] = ppes

def ppe_is_done(sheet, raw_idx):
    return f"{sheet}_{raw_idx}" in completed_ppes()

# ── Context processor ─────────────────────────────────────────────────────────
@app.context_processor
def inject_user():
    return {
        'user':      session.get('user'),
        'role':      session.get('role'),
        'name':      session.get('name'),
        'matricule': session.get('matricule'),
        'shift':     session.get('shift', '')
    }

# ── Before request: Inactivity Timeout & Session Invalidation Check ──────────
@app.before_request
def enforce_session_security():
    if 'user' in session:
        now = time.time()
        last_active = session.get('last_active')
        # 1. 60-minute idle expiration
        if last_active and (now - last_active > 3600):
            user = session.get('name', 'Unknown')
            logger.info(f"AUTH IDLE TIMEOUT: user={user}")
            session.clear()
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Session expired', 'redirect': url_for('login')}), 401
            flash("Session expirée par inactivité. Veuillez vous reconnecter.", "warning")
            return redirect(url_for('login'))

        # 2. Invalidate admin sessions if password was reset after session creation
        if session.get('role') in ('ADMIN', 'admin'):
            login_time = session.get('login_time', 0)
            if not owner_security.is_admin_session_valid(login_time):
                user = session.get('user', 'admin')
                logger.info(f"AUTH REVOCATION: Admin session revoked due to password reset (user={user})")
                session.clear()
                if request.path.startswith('/api/'):
                    return jsonify({'error': 'Admin session revoked', 'redirect': url_for('login')}), 401
                flash("Votre session administrateur a été invalidée suite à une réinitialisation du mot de passe.", "warning")
                return redirect(url_for('login'))

        session['last_active'] = now

# ── Security + no-cache headers for all responses ────────────────────────────
@app.after_request
def add_security_headers(response):
    response.headers['Cache-Control']          = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma']                 = 'no-cache'
    response.headers['Expires']                = '0'
    response.headers['Vary']                   = '*'
    response.headers['X-Frame-Options']        = 'SAMEORIGIN'
    response.headers['X-Content-Type-Options'] = 'nosniff'
    return response

# ── Auth decorators ───────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session or not session.get('auth_token'):
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Unauthorized', 'redirect': url_for('login')}), 401
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def owner_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session or not session.get('auth_token'):
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Unauthorized', 'redirect': url_for('login')}), 401
            return redirect(url_for('login'))
        r = str(session.get('role', '')).upper().strip()
        if r not in ('OWNER', 'PROPRIETAIRE', 'PROPRIÉTAIRE'):
            flash("Accès réservé exclusivement au Propriétaire (Owner).", "danger")
            return render_template('403.html'), 403
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session or not session.get('auth_token'):
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Unauthorized', 'redirect': url_for('login')}), 401
            return redirect(url_for('login'))
        r = str(session.get('role', '')).upper().strip()
        if r not in ('OWNER', 'ADMIN', 'ADMINISTRATEUR', 'PROPRIETAIRE', 'PROPRIÉTAIRE'):
            flash("Accès refusé : Droits Administrateur requis.", "danger")
            return render_template('403.html'), 403
        return f(*args, **kwargs)
    return decorated_function

# =============================================================
# AUTH ROUTES — ONE LOGIN FOR EVERYONE (OWNER / ADMIN / TECH / USER)
# =============================================================

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user' in session and session.get('auth_token'):
        return redirect(url_for('index'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        if not username or not password:
            flash("Veuillez renseigner le nom d'utilisateur et le mot de passe.", "danger")
            return render_template('login.html')

        user = user_mgr.authenticate(username, password)
        if user:
            session.clear()
            session.permanent     = False      # Expire on browser close
            session['user_id']    = user['id']
            session['user']       = user['username']
            session['username']   = user['username']
            session['role']       = user['role']
            session['name']       = user['name']
            session['matricule']  = user.get('matricule') or user['username']
            session['shift']      = user.get('shift') or 'A'
            session['auth_token'] = uuid.uuid4().hex
            session['login_time'] = time.time()
            session['last_active']= time.time()

            if user['role'] == 'TECHNICIAN':
                pma_config.set_last_tech(user['name'], user.get('matricule') or '')

            logger.info(f"AUTH LOGIN SUCCESS ({user['role']}): {user['username']} - {user['name']}")

            if user['role'] in ('OWNER', 'ADMIN'):
                return redirect(url_for('dashboard'))
            elif user['role'] == 'TECHNICIAN':
                return redirect(url_for('pma_dashboard'))
            else:
                return redirect(url_for('dashboard'))

        # Fallback admin check for legacy/config consistency
        creds     = pma_config.get_admin_credentials()
        ima_creds = ima_config.get_admin_credentials()

        if (username == creds.get('username') and password == creds.get('password')) or \
           (username == ima_creds.get('username') and password == ima_creds.get('password')):
            session.clear()
            session.permanent     = False      # Expire on browser close
            session['user_id']    = 1
            session['user']       = username
            session['username']   = username
            session['role']       = 'ADMIN'
            session['name']       = "Administrateur SEBN-TN"
            session['matricule']  = "ADMIN"
            session['shift']      = 'ALL'
            session['auth_token'] = uuid.uuid4().hex
            session['login_time'] = time.time()
            session['last_active']= time.time()
            logger.info(f"AUTH LOGIN SUCCESS (Fallback Admin): {username}")
            return redirect(url_for('dashboard'))

        flash("Nom d'utilisateur ou mot de passe incorrect.", "danger")

    return render_template('login.html')

@app.route('/logout')
def logout():
    user = session.get('name', 'Unknown')
    logger.info(f"AUTH LOGOUT: {user}")
    session.clear()
    response = make_response(redirect(url_for('login')))
    response.delete_cookie(
        app.session_interface.get_cookie_name(app),
        path='/',
        samesite='Lax'
    )
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Clear-Site-Data'] = '"cookies", "storage"'
    return response

# =============================================================
# OWNER-ONLY ADMIN RECOVERY ROUTES
# =============================================================

@app.route('/recovery', methods=['GET'])
def recovery():
    return render_template('recovery.html')

@app.route('/recovery/verify', methods=['POST'])
def recovery_verify():
    client_ip = request.remote_addr or '127.0.0.1'

    # Rate limiting on recovery attempts (max 5 attempts per 15 min)
    is_limited, rem_sec = owner_security.rate_limiter.is_rate_limited(
        client_ip, 'owner_recovery', max_attempts=5, window_seconds=900, lockout_seconds=900
    )
    if is_limited:
        rem_min = (rem_sec // 60) + 1
        owner_security.log_security_event("RATE_LIMIT_TRIGGERED", "OWNER_RECOVERY", client_ip, f"Recovery locked for {rem_min} min")
        flash(f"Trop de tentatives infructueuses. Accès temporairement bloqué pendant {rem_min} minute(s).", "danger")
        return redirect(url_for('recovery'))

    owner_user = request.form.get('owner_username', '').strip()
    owner_pass = request.form.get('owner_password', '')

    if not owner_user or not owner_pass:
        flash("Veuillez renseigner le nom d'utilisateur et le mot de passe Propriétaire.", "danger")
        return redirect(url_for('recovery'))

    if owner_security.verify_owner_login(owner_user, owner_pass):
        owner_security.rate_limiter.reset(client_ip, 'owner_recovery')
        owner_security.log_security_event("OWNER_LOGIN_SUCCESS", owner_user, client_ip, "Owner authenticated successfully for admin recovery")
        
        # Generate single-use, 10-minute reset token
        token = owner_security.generate_reset_token(ttl_seconds=600)
        session['owner_recovery_token'] = token
        return redirect(url_for('recovery_reset', token=token))
    else:
        owner_security.rate_limiter.record_attempt(client_ip, 'owner_recovery')
        owner_security.log_security_event("OWNER_LOGIN_FAILED", owner_user or "UNKNOWN", client_ip, "Failed owner recovery authentication attempt")
        flash("Identifiants Propriétaire incorrects.", "danger")
        return redirect(url_for('recovery'))

@app.route('/recovery/reset', methods=['GET', 'POST'])
def recovery_reset():
    client_ip = request.remote_addr or '127.0.0.1'
    token = request.args.get('token') or request.form.get('token') or session.get('owner_recovery_token', '')

    if not token or not owner_security.is_reset_token_valid(token):
        flash("Jeton d'autorisation expiré ou invalide. Veuillez recommencer l'authentification Propriétaire.", "danger")
        return redirect(url_for('recovery'))

    if request.method == 'POST':
        new_password = request.form.get('new_password', '').strip()
        confirm_password = request.form.get('confirm_password', '').strip()

        if not new_password or not confirm_password:
            flash("Veuillez remplir tous les champs.", "danger")
            return render_template('recovery_reset.html', token=token)

        if new_password != confirm_password:
            flash("Les deux mots de passe ne correspondent pas.", "danger")
            return render_template('recovery_reset.html', token=token)

        if len(new_password) < 6:
            flash("Le mot de passe Administrateur doit contenir au moins 6 caractères.", "danger")
            return render_template('recovery_reset.html', token=token)

        try:
            # Update admin credentials in PMA & IMA configs
            pma_config.save_json(os.path.join(pma_config.active_base, "admin_config.json"), {"username": "admin", "password": new_password})
            ima_config.set_admin_credentials("admin", new_password)

            # Invalidate all active admin sessions across all browsers
            owner_security.invalidate_all_admin_sessions()

            # Invalidate the one-time reset token and clear recovery session
            owner_security.consume_reset_token(token)
            session.clear()

            owner_security.log_security_event("ADMIN_PASSWORD_RESET_BY_OWNER", "OWNER", client_ip, "Admin password reset successfully by Owner")
            flash("Le mot de passe Administrateur a été réinitialisé avec succès. Toutes les anciennes sessions ont été révoquées.", "success")
            return redirect(url_for('login'))
        except Exception as e:
            logger.error(f"Error resetting admin password: {e}")
            flash("Erreur lors de la réinitialisation du mot de passe.", "danger")
            return render_template('recovery_reset.html', token=token)

    return render_template('recovery_reset.html', token=token)

@app.route('/access-denied')
def access_denied():
    return render_template('403.html'), 403

# =============================================================
# SHARED / HOME ROUTES
# =============================================================

@app.route('/')
@login_required
def index():
    role = (session.get('role') or '').upper()
    if role in ('OWNER', 'ADMIN'):
        return redirect(url_for('dashboard'))
    return redirect(url_for('pma_dashboard'))

@app.route('/dashboard')
@admin_required
def dashboard():
    try:
        ima_stats = ima_db.get_analytics_bundle()
    except Exception as e:
        logger.error(f"Dashboard analytics error: {e}")
        ima_stats = {}

    if 'stats' in ima_stats:
        ima_stats['open_interventions']   = ima_stats['stats']['open']
        ima_stats['closed_interventions'] = ima_stats['stats']['closed']
        ima_stats['total_downtime_hours'] = round(ima_stats['stats']['downtime'] / 60.0, 2)
        ima_stats['total_interventions']  = ima_stats['stats']['total']
    else:
        ima_stats['open_interventions']   = 0
        ima_stats['closed_interventions'] = 0
        ima_stats['total_downtime_hours'] = 0
        ima_stats['total_interventions']  = 0

    pma_engine = get_pma_engine()
    pma_stats  = pma_engine.get_stats()

    return render_template('dashboard.html', ima=ima_stats, pma=pma_stats)

@app.route('/analytics')
@admin_required
def analytics():
    date_from  = request.args.get('date_from', '')
    date_to    = request.args.get('date_to', '')
    machine_id = request.args.get('machine', 'All')
    technician = request.args.get('technician', 'All')
    status     = request.args.get('status', 'All')
    category   = request.args.get('category', 'All')

    data = ima_db.get_filtered_analytics(
        date_from=date_from,
        date_to=date_to,
        machine_id=machine_id,
        technician=technician,
        status=status,
        category=category
    )
    return render_template('analytics.html', data=data, filters={
        "date_from": date_from,
        "date_to": date_to,
        "machine": machine_id,
        "technician": technician,
        "status": status,
        "category": category
    })

@app.route('/api/analytics/ima')
@admin_required
def api_analytics_ima():
    date_from  = request.args.get('date_from', '')
    date_to    = request.args.get('date_to', '')
    machine_id = request.args.get('machine', 'All')
    technician = request.args.get('technician', 'All')
    status     = request.args.get('status', 'All')
    category   = request.args.get('category', 'All')

    data = ima_db.get_filtered_analytics(
        date_from=date_from,
        date_to=date_to,
        machine_id=machine_id,
        technician=technician,
        status=status,
        category=category
    )
    return jsonify(data)

@app.route('/analytics/export')
@admin_required
def analytics_export():
    import io, csv
    date_from  = request.args.get('date_from', '')
    date_to    = request.args.get('date_to', '')
    machine_id = request.args.get('machine', 'All')
    technician = request.args.get('technician', 'All')
    status     = request.args.get('status', 'All')
    category   = request.args.get('category', 'All')

    data = ima_db.get_filtered_analytics(
        date_from=date_from,
        date_to=date_to,
        machine_id=machine_id,
        technician=technician,
        status=status,
        category=category
    )
    
    output = io.StringIO()
    writer = csv.writer(output, delimiter=';')
    writer.writerow(['Code', 'Date Creation', 'Machine', 'Groupe', 'Description Panne', 'Technicien', 'Matricule', 'Debut', 'Fin', 'Duree (min)', 'Statut'])
    
    for row in data.get('interventions', []):
        writer.writerow([
            row.get('code', ''),
            row.get('created_at', ''),
            row.get('machine_id', ''),
            row.get('group_name', ''),
            row.get('fault_description', ''),
            row.get('technician_name', ''),
            row.get('technician_mat', ''),
            row.get('start_time', ''),
            row.get('end_time', ''),
            row.get('downtime_minutes', 0),
            row.get('status', '')
        ])
    
    output.seek(0)
    bytes_out = io.BytesIO(output.getvalue().encode('utf-8-sig'))
    return send_file(bytes_out, mimetype='text/csv', as_attachment=True, download_name=f'analyse_interventions_{datetime.date.today().isoformat()}.csv')

@app.route('/training')
@login_required
def training():
    # If the user is a technician, redirect to their space
    if session.get('role') == 'TECHNICIAN':
        return redirect(url_for('tech_my_exams'))
    
    # Load technicians with real progression data from the exam system
    techs = exam_mgr.get_all_technicians_with_progression()
    all_exams = exam_mgr.get_all_exams()
    
    # MVP = technician with the highest level
    level_order = exam_mgr.LEVEL_ORDER
    mvp = None
    max_idx = -1
    for t in techs:
        lvl = exam_mgr.normalize_level(t.get('technician_level'))
        idx = level_order.index(lvl) if lvl in level_order else 0
        if idx > max_idx:
            max_idx = idx
            mvp = t

    return render_template('training.html', techs=techs, mvp=mvp, all_exams=all_exams)

@app.route('/training/<matricule>')
@admin_required
def training_profile(matricule):
    profile = tech_db.get_profile(matricule)
    if not profile:
        flash("Technicien introuvable.", "danger")
        return redirect(url_for('training'))
    return render_template('training_profile.html', profile=profile)

@app.route('/training/import', methods=['POST'])
@admin_required
def training_import():
    file = request.files.get('excel_file')
    if file and file.filename:
        path = os.path.join(pma_config.dirs["archives"], "temp_import.xlsx")
        file.save(path)
        ok, msg = tech_db.import_from_excel(path)
        flash(msg, "success" if ok else "danger")
    return redirect(url_for('training'))

@app.route('/training/<matricule>/validate-level', methods=['POST'])
@admin_required
def validate_level(matricule):
    level   = request.form.get('level')
    date    = request.form.get('date', datetime.date.today().isoformat())
    profile = tech_db.get_profile(matricule)
    if profile and level:
        profile['exams'].append({
            "name":       f"Validation {level}",
            "type":       "Validation",
            "exam_level": level,
            "date":       date,
            "year":       int(date[:4]) if date else datetime.datetime.now().year
        })
        tech_db.update_tech(profile)
        flash(f"Niveau {level} validé.", "success")
    return redirect(url_for('training_profile', matricule=matricule))

@app.route('/training/<matricule>/certificate', methods=['POST'])
@admin_required
def upload_certificate(matricule):
    flash("Certificat ajouté.", "success")
    return redirect(url_for('training_profile', matricule=matricule))

@app.route('/training/<matricule>/exam', methods=['POST'])
@admin_required
def training_add_exam(matricule):
    exam_type = request.form.get('type', 'Formation')
    level     = request.form.get('exam_level', '')
    date_str  = request.form.get('date', datetime.date.today().isoformat())
    notes     = request.form.get('notes', '')

    profile = tech_db.get_profile(matricule)
    if not profile:
        flash("Technicien introuvable.", "danger")
        return redirect(url_for('training'))

    if 'exams' not in profile:
        profile['exams'] = []

    profile['exams'].append({
        "type":       exam_type,
        "exam_level": level,
        "date":       date_str,
        "notes":      notes,
        "year":       int(date_str[:4]) if date_str else datetime.datetime.now().year
    })
    tech_db.update_tech(profile)
    flash(f"Événement '{exam_type}' ajouté avec succès.", "success")
    return redirect(url_for('training_profile', matricule=matricule))


# ── Formation & Exam System Routes ──────────────────────────────────────────

@app.route('/training/my-exams')
@login_required
def tech_my_exams():
    if session.get('role') not in ('TECHNICIAN',):
        return redirect(url_for('training'))
    
    u = user_mgr.get_user_by_username(session.get('username'))
    if not u:
        flash("Compte utilisateur introuvable.", "danger")
        return redirect(url_for('logout'))
        
    all_exams = exam_mgr.get_all_exams(role='TECHNICIAN')
    user_level = exam_mgr.normalize_level(u.get('technician_level'))
    progression = exam_mgr.get_technician_progression(u['id'])
    
    with exam_mgr._get_conn() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT ea.*, e.title as exam_title, e.number_of_questions
            FROM exam_attempts ea
            JOIN exams e ON ea.exam_id = e.id
            WHERE ea.user_id = ?
            ORDER BY ea.started_at DESC
        """, (u['id'],))
        attempts = [dict(r) for r in cur.fetchall()]

    card = exam_mgr.get_technician_card_details(u['id'])
    
    return render_template('tech_exams.html', user=u, exams=all_exams, attempts=attempts, card=card, progression=progression)


@app.route('/training/exams')
@admin_required
def training_exams():
    exams = exam_mgr.get_all_exams(role=session.get('role'))
    techs = exam_mgr.get_all_technicians_with_progression()
    return render_template('training_exams.html', exams=exams, techs=techs)


@app.route('/training/exams/import', methods=['POST'])
@admin_required
def exam_import():
    file = request.files.get('exam_file')
    if file and file.filename:
        fname = file.filename.lower()
        lvl = None
        if "25%" in fname or "level 1" in fname or "level1" in fname: lvl = "Level 1"
        elif "50%" in fname or "level 2" in fname or "level2" in fname: lvl = "Level 2"
        elif "75%" in fname or "level 3" in fname or "level3" in fname: lvl = "Level 3"
        elif "100%" in fname or "level 4" in fname or "level4" in fname: lvl = "Level 4"
        
        if not lvl:
            flash("Nom du fichier incorrect. Il doit contenir Level 1, 2, 3 ou 4 (ou 25%, 50%, 75%, 100%).", "danger")
            return redirect(url_for('training_exams'))

        temp_path = os.path.join(pma_config.dirs["archives"], f"temp_{secrets.token_hex(4)}.docx")
        os.makedirs(os.path.dirname(temp_path), exist_ok=True)
        file.save(temp_path)
        
        try:
            questions = exam_mgr.parse_docx_exam(temp_path, output_images_dir=EXAM_IMAGES_DIR)
            levels_pct = {"Level 1": 25, "Level 2": 50, "Level 3": 75, "Level 4": 100,
                          "LEVEL 25%": 25, "LEVEL 50%": 50, "LEVEL 75%": 75, "LEVEL 100%": 100}
            
            exam_id = exam_mgr.create_exam_with_questions(
                title=f"Test {lvl}",
                description=f"Examen de validation pour le niveau {lvl}",
                technician_level=lvl,
                level_percentage=levels_pct.get(lvl, 25),
                source_file=file.filename,
                duration=30,
                passing_score=75,
                status='draft',
                parsed_questions=questions
            )
            flash(f"Examen importé avec succès. {len(questions)} questions détectées. Veuillez définir les réponses correctes.", "success")
            return redirect(url_for('exam_preview', exam_id=exam_id))
        except Exception as e:
            flash(f"Erreur d'importation : {e}", "danger")
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
    return redirect(url_for('training_exams'))


@app.route('/training/exams/preview/<int:exam_id>')
@admin_required
def exam_preview(exam_id: int):
    exam = exam_mgr.get_exam_by_id(exam_id)
    if not exam:
        flash("Examen introuvable.", "danger")
        return redirect(url_for('training_exams'))
    return render_template('exam_preview.html', exam=exam)


@app.route('/training/exams/save-config/<int:exam_id>', methods=['POST'])
@admin_required
def exam_save_config(exam_id: int):
    title = request.form.get('title', '').strip()
    description = request.form.get('description', '').strip()
    duration = int(request.form.get('duration', '30'))
    passing_score = int(request.form.get('passing_score', '75'))
    status = request.form.get('status', 'draft')

    exam_mgr.update_exam_details(exam_id, title, description, duration, passing_score, status)

    correct_answers_map = {}
    for key, value in request.form.items():
        if key.startswith('correct_answer_'):
            q_id = key.replace('correct_answer_', '')
            correct_answers_map[q_id] = int(value)

    exam_mgr.save_exam_answers_setup(exam_id, correct_answers_map)
    flash("Configuration de l'examen enregistrée avec succès.", "success")
    return redirect(url_for('exam_preview', exam_id=exam_id))


@app.route('/training/exams/delete/<int:exam_id>', methods=['POST'])
@admin_required
def exam_delete(exam_id: int):
    exam_mgr.delete_exam(exam_id)
    flash("Examen supprimé avec succès.", "success")
    return redirect(url_for('training_exams'))


@app.route('/training/results')
@admin_required
def training_results():
    filters = {
        'technician': request.args.get('technician', ''),
        'matricule': request.args.get('matricule', ''),
        'level': request.args.get('level', ''),
        'exam_id': request.args.get('exam_id', ''),
        'status': request.args.get('status', ''),
        'date': request.args.get('date', '')
    }
    results = exam_mgr.get_all_results(filters)
    exams = exam_mgr.get_all_exams(role='admin')
    return render_template('training_results.html', results=results, exams=exams, filters=filters)


@app.route('/training/cards')
@admin_required
def training_cards():
    with exam_mgr._get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id FROM users WHERE role = 'TECHNICIAN' ORDER BY name ASC")
        tech_ids = [r['id'] for r in cur.fetchall()]
        
    techs = []
    for t_id in tech_ids:
        details = exam_mgr.get_technician_card_details(t_id)
        if details:
            techs.append(details)
            
    return render_template('training_cards.html', techs=techs)


@app.route('/training/card/view/<int:user_id>')
@login_required
def tech_card_view(user_id: int):
    if session.get('role') == 'TECHNICIAN':
        u = user_mgr.get_user_by_username(session.get('username'))
        if not u or u['id'] != user_id:
            flash("Accès non autorisé.", "danger")
            return redirect(url_for('training'))

    card = exam_mgr.get_technician_card_details(user_id)
    if not card:
        flash("Technicien introuvable.", "danger")
        return redirect(url_for('training'))
        
    return render_template('tech_card_view.html', card=card)


@app.route('/training/card/upload-photo/<int:user_id>', methods=['POST'])
@admin_required
def tech_card_upload_photo(user_id: int):
    file = request.files.get('photo_file')
    if file and file.filename:
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in ('.png', '.jpg', '.jpeg'):
            flash("Format d'image non valide. Utilisez PNG, JPG ou JPEG.", "danger")
            return redirect(url_for('tech_card_view', user_id=user_id))
            
        u = user_mgr.get_user_by_id(user_id)
        if not u:
            flash("Utilisateur introuvable.", "danger")
            return redirect(url_for('training'))
            
        mat = u.get('matricule') or f"user_{user_id}"
        safe_mat = str(mat).replace("/", "_").replace("\\", "_").strip()
        
        photo_dir = os.path.join(pma_config.active_base, "Technicians", "photos")
        os.makedirs(photo_dir, exist_ok=True)
        photo_filename = f"photo_{safe_mat}{ext}"
        photo_path = os.path.join(photo_dir, photo_filename)
        file.save(photo_path)
        
        with exam_mgr._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE users SET photo = ? WHERE id = ?", (photo_filename, user_id))
            conn.commit()
            
        flash("Photo du profil mise à jour.", "success")
        
    return redirect(url_for('tech_card_view', user_id=user_id))


@app.route('/training/static-photo/<filename>')
@login_required
def static_photo_download(filename):
    photo_path = os.path.join(pma_config.active_base, "Technicians", "photos", filename)
    if os.path.exists(photo_path):
        return send_file(photo_path)
    return redirect(url_for('static', filename='placeholder.png'))


@app.route('/training/card/print/<int:user_id>')
@login_required
def tech_card_print(user_id: int):
    if session.get('role') == 'TECHNICIAN':
        u = user_mgr.get_user_by_username(session.get('username'))
        if not u or u['id'] != user_id:
            return "Accès non autorisé.", 403

    card = exam_mgr.get_technician_card_details(user_id)
    if not card:
        return "Technicien introuvable.", 404
        
    if card.get('photo'):
        card['photo'] = os.path.join(pma_config.active_base, "Technicians", "photos", card['photo'])

    try:
        pdf_path = generate_technician_card_pdf(card)
        return send_file(pdf_path, mimetype='application/pdf', as_attachment=False)
    except Exception as e:
        return f"Erreur lors de la génération du PDF : {e}", 500


@app.route('/training/exam/start/<int:exam_id>', methods=['GET', 'POST'])
@login_required
def exam_start(exam_id: int):
    role = session.get('role', '').upper().strip()
    tech_id = request.form.get('technician_id', type=int) or request.args.get('technician_id', type=int)

    if role in ('OWNER', 'ADMIN', 'admin'):
        if not tech_id:
            flash("Veuillez sélectionner un technicien pour passer l'examen.", "warning")
            return redirect(url_for('training'))
        u = user_mgr.get_user_by_id(tech_id)
        if not u or u.get('role') != 'TECHNICIAN':
            flash("Le compte sélectionné n'est pas un technicien valide.", "danger")
            return redirect(url_for('training'))
    else:
        u = user_mgr.get_user_by_username(session.get('username'))
        if not u or u.get('role') != 'TECHNICIAN':
            flash("Seuls les techniciens peuvent passer des examens.", "danger")
            return redirect(url_for('training'))

    exam = exam_mgr.get_exam_by_id(exam_id)
    if not exam:
        flash("Examen introuvable.", "danger")
        return redirect(url_for('training'))

    if exam.get('status') != 'published' and role not in ('OWNER', 'ADMIN', 'admin'):
        flash("Cet examen n'est pas encore disponible.", "danger")
        return redirect(url_for('training'))

    attempt = exam_mgr.start_exam_attempt(exam_id, u['id'])
    return redirect(url_for('exam_take_view', attempt_id=attempt['id']))


@app.route('/training/exam/take/<int:attempt_id>')
@login_required
def exam_take_view(attempt_id: int):
    attempt = exam_mgr.get_active_attempt(attempt_id)
    if not attempt or attempt['status'] != 'running':
        flash("Tentative d'examen introuvable ou déjà soumise.", "danger")
        return redirect(url_for('training'))

    u = user_mgr.get_user_by_id(attempt['user_id'])
    if session.get('role') == 'TECHNICIAN' and session.get('username') != u['username']:
        flash("Accès non autorisé.", "danger")
        return redirect(url_for('training'))

    exam = exam_mgr.get_exam_by_id(attempt['exam_id'])
    rem_seconds = exam_mgr.get_remaining_seconds(attempt_id, exam['duration'])
    
    if rem_seconds <= 0:
        exam_mgr.submit_exam_attempt(attempt_id)
        flash("Le temps de l'examen est écoulé. Vos réponses ont été soumises.", "warning")
        return redirect(url_for('exam_result_view', attempt_id=attempt_id))

    saved_answers = exam_mgr.get_saved_answers_for_attempt(attempt_id)
    saved_text_answers = exam_mgr.get_saved_text_answers_for_attempt(attempt_id)
    
    return render_template(
        'exam_take.html', 
        attempt=attempt, 
        exam=exam, 
        remaining_seconds=rem_seconds, 
        saved_answers=saved_answers,
        saved_text_answers=saved_text_answers,
        tech=u
    )


@app.route('/training/exam/save-answer', methods=['POST'])
@login_required
def exam_save_answer():
    data = request.get_json() or {}
    attempt_id = data.get('attempt_id')
    question_id = data.get('question_id')
    selected_answer_id = data.get('selected_answer_id')
    answer_text = data.get('answer_text')

    if attempt_id and question_id:
        attempt = exam_mgr.get_active_attempt(attempt_id)
        if attempt and attempt['status'] == 'running':
            exam_mgr.save_attempt_answer(
                attempt_id, question_id, 
                selected_answer_id=selected_answer_id,
                answer_text=answer_text
            )
            return jsonify({"status": "success"})
            
    return jsonify({"status": "error"}), 400


@app.route('/training/exam/submit/<int:attempt_id>', methods=['POST'])
@login_required
def exam_submit(attempt_id: int):
    attempt = exam_mgr.get_active_attempt(attempt_id)
    if not attempt or attempt['status'] != 'running':
        flash("Tentative déjà soumise ou inexistante.", "danger")
        return redirect(url_for('training'))

    u = user_mgr.get_user_by_id(attempt['user_id'])
    if session.get('role') == 'TECHNICIAN' and session.get('username') != u['username']:
        flash("Accès non autorisé.", "danger")
        return redirect(url_for('training'))

    # Save any form responses submitted with the form POST
    for k, v in request.form.items():
        if k.startswith('question_'):
            try:
                qid = int(k.replace('question_', ''))
                ans_id = int(v)
                exam_mgr.save_attempt_answer(attempt_id, qid, selected_answer_id=ans_id)
            except Exception:
                pass
        elif k.startswith('labeling_answers_'):
            try:
                qid = int(k.replace('labeling_answers_', ''))
                if v and str(v).strip():
                    exam_mgr.save_attempt_answer(attempt_id, qid, answer_text=str(v).strip())
            except Exception:
                pass
        elif k.startswith('text_answer_'):
            try:
                qid = int(k.replace('text_answer_', ''))
                if v and str(v).strip():
                    exam_mgr.save_attempt_answer(attempt_id, qid, answer_text=str(v).strip())
            except Exception:
                pass

    exam_mgr.submit_exam_attempt(attempt_id)
    flash("Examen soumis et corrigé avec succès.", "success")
    return redirect(url_for('exam_result_view', attempt_id=attempt_id))


@app.route('/training/exam/result/<int:attempt_id>')
@login_required
def exam_result_view(attempt_id: int):
    with exam_mgr._get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM exam_attempts WHERE id = ?", (attempt_id,))
        attempt = cur.fetchone()
        
    if not attempt:
        flash("Tentative d'examen introuvable.", "danger")
        return redirect(url_for('training'))

    attempt = dict(attempt)
    u = user_mgr.get_user_by_id(attempt['user_id'])
    if session.get('role') == 'TECHNICIAN' and session.get('username') != u['username']:
        flash("Accès non autorisé.", "danger")
        return redirect(url_for('training'))

    exam = exam_mgr.get_exam_by_id(attempt['exam_id'])
    
    return render_template('exam_result.html', attempt=attempt, exam=exam, tech=u)


# ── ADMIN EXAM MANAGEMENT ROUTES ─────────────────────────────────────────────

EXAM_IMAGES_DIR = os.path.join(os.path.dirname(__file__), 'static', 'exam_images')
ALLOWED_IMG_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}


def _allowed_img(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_IMG_EXTENSIONS


@app.route('/api/admin/exam/<int:exam_id>/questions')
@admin_required
def api_admin_exam_questions(exam_id: int):
    """Returns full exam with questions+answers+images as JSON (for modal editor)."""
    exam = exam_mgr.get_exam_by_id(exam_id)
    if not exam:
        return jsonify({'error': 'not found'}), 404
    return jsonify(exam)


@app.route('/admin/exam/<int:exam_id>/edit', methods=['POST'])
@admin_required
def admin_exam_edit(exam_id: int):
    """Update exam metadata from admin panel."""
    title        = request.form.get('title', '').strip()
    description  = request.form.get('description', '').strip()
    duration     = int(request.form.get('duration', 30))
    passing_score = int(request.form.get('passing_score', 75))
    status       = request.form.get('status', 'draft')
    if not title:
        return jsonify({'error': 'Title required'}), 400
    exam_mgr.update_exam_details(exam_id, title, description, duration, passing_score, status)
    return jsonify({'success': True})


@app.route('/admin/exam/create', methods=['POST'])
@admin_required
def admin_exam_create():
    """Create a blank draft exam."""
    title  = request.form.get('title', '').strip()
    level  = exam_mgr.normalize_level(request.form.get('technician_level', 'Level 1'))
    pct_map = {'Level 1': 25, 'Level 2': 50, 'Level 3': 75, 'Level 4': 100,
               'LEVEL 25%': 25, 'LEVEL 50%': 50, 'LEVEL 75%': 75, 'LEVEL 100%': 100}
    pct = pct_map.get(level, 25)
    if not title:
        flash("Un titre est requis pour créer un examen.", "danger")
        return redirect(url_for('admin') + '#exams')
    exam_mgr.create_blank_exam(title, level, pct)
    flash(f"Examen '{title}' créé en mode brouillon.", "success")
    return redirect(url_for('admin') + '#exams')


@app.route('/admin/exam/<int:exam_id>/delete', methods=['POST'])
@admin_required
def admin_exam_delete(exam_id: int):
    """Delete an exam from admin panel."""
    exam_mgr.delete_exam(exam_id)
    flash("Examen supprimé.", "success")
    return redirect(url_for('admin') + '#exams')


@app.route('/admin/exam/question/<int:question_id>/upload-image', methods=['POST'])
@admin_required
def admin_question_upload_image(question_id: int):
    """Upload an image for a specific exam question."""
    if 'image' not in request.files:
        return jsonify({'error': 'No file'}), 400
    f = request.files['image']
    if not f or not _allowed_img(f.filename):
        return jsonify({'error': 'Invalid file type. Allowed: png, jpg, jpeg, gif, webp'}), 400

    import uuid
    ext = f.filename.rsplit('.', 1)[1].lower()
    filename = f"q_{question_id}_{uuid.uuid4().hex[:8]}.{ext}"
    os.makedirs(EXAM_IMAGES_DIR, exist_ok=True)
    f.save(os.path.join(EXAM_IMAGES_DIR, filename))
    exam_mgr.update_question_image(question_id, filename)
    return jsonify({'success': True, 'filename': filename, 'url': f'/static/exam_images/{filename}'})


@app.route('/admin/exam/question/<int:question_id>/remove-image', methods=['POST'])
@admin_required
def admin_question_remove_image(question_id: int):
    """Remove the image from an exam question."""
    exam_mgr.remove_question_image(question_id)
    return jsonify({'success': True})


@app.route('/admin/exam/question/<int:source_question_id>/move-image-to/<int:target_question_id>', methods=['POST'])
@admin_required
def admin_question_move_image(source_question_id: int, target_question_id: int):
    """
    Move (reassign) the primary image of source_question_id to target_question_id.
    Admin safety feature: 'Associer l'image à la question'.
    Both questions must belong to the same exam.
    """
    ok, msg = exam_mgr.move_question_image(source_question_id, target_question_id)
    return jsonify({'success': ok, 'message': msg})


@app.route('/admin/exam/question/<int:question_id>/update-text', methods=['POST'])
@admin_required
def admin_question_update_text(question_id: int):
    """Update the text of a question."""
    text = (request.form.get('question_text') or request.get_json(silent=True, force=True) or {}).get('question_text', '') if request.is_json else request.form.get('question_text', '')
    if request.is_json:
        text = (request.get_json(silent=True) or {}).get('question_text', '')
    else:
        text = request.form.get('question_text', '')
    if not text:
        return jsonify({'error': 'No text'}), 400
    exam_mgr.update_question_text(question_id, text)
    return jsonify({'success': True})


@app.route('/admin/exam/question/<int:question_id>/delete', methods=['POST'])
@admin_required
def admin_question_delete(question_id: int):
    """Delete a question from an exam."""
    exam_mgr.delete_question(question_id)
    return jsonify({'success': True})


@app.route('/admin/exam/question/<int:question_id>/set-type', methods=['POST'])
@admin_required
def admin_question_set_type(question_id: int):
    """Update question_type for a question."""
    data = request.get_json(silent=True) or request.form
    new_type = data.get('question_type')
    if not new_type:
        return jsonify({'error': 'Type de question manquant'}), 400
    ok = exam_mgr.set_question_type(question_id, new_type)
    if ok:
        return jsonify({'success': True, 'question_type': new_type})
    return jsonify({'error': 'Type de question invalide'}), 400


@app.route('/admin/exam/question/<int:question_id>/set-labels', methods=['POST'])
@admin_required
def admin_question_set_labels(question_id: int):
    """Update available_labels list for an image_labeling question."""
    data = request.get_json(silent=True) or {}
    labels = data.get('labels', [])
    ok = exam_mgr.set_available_labels(question_id, labels)
    if ok:
        return jsonify({'success': True, 'labels': labels})
    return jsonify({'error': 'Erreur lors de la sauvegarde des étiquettes'}), 400


@app.route('/admin/exam/<int:exam_id>/fix-types', methods=['POST'])
@admin_required
def admin_exam_fix_types(exam_id: int):
    """Auto-detect and fix question types for a single exam."""
    result = exam_mgr.fix_question_types_for_exam(exam_id)
    return jsonify({'success': True, 'result': result})


@app.route('/admin/exam/fix-all-types', methods=['POST'])
@admin_required
def admin_exam_fix_all_types():
    """Auto-detect and fix question types for all exams."""
    result = exam_mgr.fix_all_question_types()
    return jsonify({'success': True, 'result': result})


@app.route('/admin/exam/reimport-all-official', methods=['POST'])
@admin_required
def admin_exam_reimport_all_official():
    """Re-import official exams from Downloads folder."""
    result = exam_mgr.reimport_official_exams()
    if result.get('success'):
        flash("Les examens officiels ont été réimportés avec succès avec détection automatique des types.", "success")
    else:
        flash("Erreur lors de la réimportation: " + result.get('error', result.get('message', '')), "danger")
    return redirect(url_for('admin') + '#exams')



@app.route('/admin/exam/<int:exam_id>/set-correct-answers', methods=['POST'])
@admin_required
def admin_exam_set_correct_answers(exam_id: int):
    """Save correct answer selections for all questions of an exam."""
    data = request.get_json(silent=True) or {}
    correct_map = data.get('correct_answers', {})
    exam_mgr.save_exam_answers_setup(exam_id, correct_map)
    return jsonify({'success': True})


@app.route('/admin/exam/<int:exam_id>/publish', methods=['POST'])
@admin_required
def admin_exam_publish(exam_id: int):
    """Toggle publish status for an exam."""
    new_status = exam_mgr.toggle_publish_exam(exam_id)
    if new_status == 'published':
        flash("Examen publié avec succès. Il est maintenant disponible pour les techniciens.", "success")
    elif new_status == 'draft':
        flash("Examen passé en mode brouillon.", "warning")
    else:
        flash("Examen introuvable.", "danger")
    return redirect(url_for('admin') + '#exams')


@app.route('/admin/exam/<int:exam_id>/add-question', methods=['POST'])
@admin_required
def admin_exam_add_question(exam_id: int):
    """Add a new question with 4 choices to an exam."""
    q_text = request.form.get('question_text', '').strip()
    if not q_text:
        return jsonify({'error': 'Texte de la question requis'}), 400

    choices = []
    correct_idx = request.form.get('correct_choice', type=int) or 0
    for i in range(4):
        txt = request.form.get(f'choice_{i}', '').strip()
        if txt:
            choices.append({'answer_text': txt, 'is_correct': 1 if i == correct_idx else 0})

    if not choices:
        return jsonify({'error': 'Au moins une réponse est requise'}), 400

    q_id = exam_mgr.add_question_with_choices(exam_id, q_text, choices)

    if 'image' in request.files:
        f = request.files['image']
        if f and _allowed_img(f.filename):
            ext = f.filename.rsplit('.', 1)[1].lower()
            filename = f"q_{q_id}_{uuid.uuid4().hex[:8]}.{ext}"
            os.makedirs(EXAM_IMAGES_DIR, exist_ok=True)
            f.save(os.path.join(EXAM_IMAGES_DIR, filename))
            exam_mgr.update_question_image(q_id, filename)

    return jsonify({'success': True, 'question_id': q_id})


@app.route('/admin/exam/import-docx', methods=['POST'])
@admin_required
def admin_exam_import_docx():
    """Upload and parse a .docx exam file using a two-pass boundary-aware image extractor."""
    if 'exam_file' not in request.files:
        flash("Aucun fichier sélectionné.", "danger")
        return redirect(url_for('admin') + '#exams')
    f = request.files['exam_file']
    if not f or not f.filename.endswith('.docx'):
        flash("Veuillez sélectionner un fichier Word (.docx) valide.", "danger")
        return redirect(url_for('admin') + '#exams')

    level = exam_mgr.normalize_level(request.form.get('technician_level', 'Level 1'))
    pct_map = {'Level 1': 25, 'Level 2': 50, 'Level 3': 75, 'Level 4': 100,
               'LEVEL 25%': 25, 'LEVEL 50%': 50, 'LEVEL 75%': 75, 'LEVEL 100%': 100}
    pct = pct_map.get(level, 25)
    title = request.form.get('title', '').strip() or f.filename.replace('.docx', '')

    import tempfile

    temp_dir = tempfile.mkdtemp()
    temp_path = os.path.join(temp_dir, f.filename)
    f.save(temp_path)

    try:
        # Document-structure-aware parser extracts questions, textboxes, and correctly associates images
        parsed = exam_mgr.parse_docx_exam(temp_path, output_images_dir=EXAM_IMAGES_DIR)

        ex_id = exam_mgr.create_exam_with_questions(
            title=title,
            description=f"Importé via admin pour {level}",
            technician_level=level,
            level_percentage=pct,
            source_file=f.filename,
            duration=int(request.form.get('duration', 30)),
            passing_score=int(request.form.get('passing_score', 75)),
            status='draft',
            parsed_questions=parsed
        )

        flash(f"Examen '{title}' importé avec succès ({len(parsed)} questions).", "success")
    except Exception as e:
        logger.exception(f"[EXAM IMPORT] Error: {e}")
        flash(f"Erreur lors de l'import : {e}", "danger")
    finally:
        try:
            os.remove(temp_path)
            os.rmdir(temp_dir)
        except Exception:
            pass

    return redirect(url_for('admin') + '#exams')



@app.route('/admin/tech-level-override/<int:user_id>', methods=['POST'])
@admin_required
def admin_tech_level_override(user_id: int):
    """Admin/Owner manual technician level override from the admin panel."""
    new_level   = request.form.get('new_level', '').strip()
    reason      = request.form.get('reason', '').strip()
    create_cert = (request.form.get('create_cert') == '1')
    actor       = session.get('username', 'admin')
    ok, msg     = exam_mgr.owner_override_level(
        user_id=user_id,
        new_level=new_level,
        reason=reason,
        owner_username=actor,
        create_cert=create_cert
    )
    flash(msg, "success" if ok else "danger")
    return redirect(url_for('admin') + '#exams')


@app.route('/vault')
@admin_required
def vault():
    archives = get_vault_files()
    ppes = []
    filled_dir = pma_config.dirs.get("filled", "")
    if os.path.exists(filled_dir):
        for f in os.listdir(filled_dir):
            p = os.path.join(filled_dir, f)
            if os.path.isfile(p):
                ppes.append({
                    "name": f,
                    "size": os.path.getsize(p) // 1024,
                    "date": datetime.datetime.fromtimestamp(os.path.getmtime(p)).strftime('%Y-%m-%d %H:%M')
                })

    archives.sort(key=lambda x: x['date'], reverse=True)
    ppes.sort(key=lambda x: x['date'], reverse=True)
    return render_template('vault.html', files=archives, ppes=ppes)

@app.route('/vault/download/<path:filename>')
@admin_required
def vault_download(filename):
    # Security: prevent path traversal
    base_dir     = os.path.join(pma_config.active_base, "Archives")
    safe_path    = os.path.realpath(os.path.join(base_dir, filename))
    base_real    = os.path.realpath(base_dir)
    if not safe_path.startswith(base_real):
        flash("Accès non autorisé.", "danger")
        return redirect(url_for('vault'))
    if os.path.exists(safe_path):
        return send_file(safe_path, as_attachment=True)
    flash("Fichier introuvable.", "danger")
    return redirect(url_for('vault'))

@app.route('/vault/download/ppe/<path:filename>')
@login_required
def download_ppe(filename):
    filled_dir = pma_config.dirs.get('filled', os.path.join(pma_config.active_base, 'PPE_Filled'))
    safe_path  = os.path.realpath(os.path.join(filled_dir, filename))
    base_real  = os.path.realpath(filled_dir)
    if not safe_path.startswith(base_real):
        flash("Accès non autorisé.", "danger")
        return redirect(url_for('vault'))
    if os.path.isfile(safe_path) and os.path.getsize(safe_path) > 0:
        return send_file(safe_path, as_attachment=True)
    flash("Fichier Excel introuvable ou vide.", "danger")
    return redirect(url_for('vault'))


# =============================================================
# USER & TECHNICIAN MANAGEMENT (OWNER ONLY)
# =============================================================

@app.route('/admin/users')
@owner_required
def admin_users():
    users = user_mgr.get_all_users()
    return render_template('users.html', users=users)

@app.route('/admin/users/create', methods=['POST'])
@owner_required
def admin_users_create():
    name = request.form.get('name', '').strip()
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '')
    matricule = request.form.get('matricule', '').strip()
    role = request.form.get('role', 'TECHNICIAN').strip()
    shift = request.form.get('shift', 'A').strip()
    technician_level = request.form.get('technician_level', '').strip() or None

    ok, msg, _ = user_mgr.create_user(
        name=name, username=username, password=password,
        role=role, matricule=matricule, shift=shift,
        technician_level=technician_level
    )
    flash(msg, "success" if ok else "danger")
    return redirect(url_for('admin_users'))

@app.route('/admin/users/edit/<int:user_id>', methods=['POST'])
@owner_required
def admin_users_edit(user_id: int):
    name = request.form.get('name', '').strip()
    username = request.form.get('username', '').strip()
    matricule = request.form.get('matricule', '').strip()
    role = request.form.get('role', 'TECHNICIAN').strip()
    shift = request.form.get('shift', 'A').strip()
    # NOTE: technician_level is intentionally NOT read from the edit form.
    # Level is controlled exclusively by exam results.
    # OWNER can override via /training/level-override/<user_id>.

    ok, msg = user_mgr.update_user(
        user_id=user_id, name=name, username=username,
        role=role, matricule=matricule, shift=shift
    )
    flash(msg, "success" if ok else "danger")
    return redirect(url_for('admin_users'))


@app.route('/training/progression/<int:user_id>')
@admin_required
def tech_progression(user_id: int):
    """Admin/Owner view: technician progression card + level history."""
    progression = exam_mgr.get_technician_progression(user_id)
    if not progression:
        flash("Technicien introuvable.", "danger")
        return redirect(url_for('training'))
    history = exam_mgr.get_technician_level_history(user_id)
    all_exams = exam_mgr.get_all_exams()
    return render_template('tech_progression.html', progression=progression, history=history, all_exams=all_exams)


@app.route('/training/level-override/<int:user_id>', methods=['POST'])
@owner_required
def tech_level_override(user_id: int):
    """OWNER ONLY: manual level override with reason and audit log."""
    new_level = request.form.get('new_level', '').strip()
    reason    = request.form.get('reason', '').strip()
    owner_username = session.get('username', 'owner')

    ok, msg = exam_mgr.owner_override_level(
        user_id=user_id,
        new_level=new_level,
        reason=reason,
        owner_username=owner_username
    )
    flash(msg, "success" if ok else "danger")
    return redirect(url_for('tech_progression', user_id=user_id))

@app.route('/admin/users/reset-password/<int:user_id>', methods=['POST'])
@owner_required
def admin_users_reset_password(user_id: int):
    new_password = request.form.get('new_password', '')
    ok, msg = user_mgr.reset_password(user_id, new_password)
    flash(msg, "success" if ok else "danger")
    return redirect(url_for('admin_users'))

@app.route('/admin/users/toggle/<int:user_id>', methods=['POST'])
@owner_required
def admin_users_toggle(user_id: int):
    ok, msg, _ = user_mgr.toggle_user_status(user_id)
    flash(msg, "success" if ok else "danger")
    return redirect(url_for('admin_users'))


# =============================================================
# EBM MODULE (OWNER ONLY)
# =============================================================

@app.route('/ebm')
@owner_required
def ebm_dashboard():
    kpis = ebm_mgr.get_kpis()
    return render_template('ebm_dashboard.html', kpis=kpis)

@app.route('/ebm/analytics')
@owner_required
def ebm_analytics():
    kpis = ebm_mgr.get_kpis()
    return render_template('ebm_analytics.html', kpis=kpis)

@app.route('/ebm/validation')
@owner_required
def ebm_validation():
    kpis = ebm_mgr.get_kpis()
    return render_template('ebm_validation.html', kpis=kpis)

@app.route('/ebm/reception')
@owner_required
def ebm_reception():
    kpis = ebm_mgr.get_kpis()
    return render_template('ebm_validation.html', kpis=kpis)

@app.route('/ebm/plan-action')
@owner_required
def ebm_plan_action():
    plans = ebm_mgr.get_action_plans()
    return render_template('ebm_plan_action.html', plans=plans)

@app.route('/ebm/plan-action/add', methods=['POST'])
@owner_required
def ebm_plan_action_add():
    title = request.form.get('title', '').strip()
    description = request.form.get('description', '').strip()
    responsible = request.form.get('responsible', '').strip()
    due_date = request.form.get('due_date', '').strip()
    priority = request.form.get('priority', 'Medium')
    if title:
        ebm_mgr.add_action_plan(title, description, responsible, due_date, priority)
        flash("Action EBM ajoutée avec succès.", "success")
    else:
        flash("Le titre de l'action est obligatoire.", "danger")
    return redirect(url_for('ebm_plan_action'))

@app.route('/ebm/plan-action/toggle/<int:plan_id>', methods=['POST'])
@owner_required
def ebm_plan_action_toggle(plan_id: int):
    ok, new_st = ebm_mgr.toggle_action_plan(plan_id)
    flash(f"Statut de l'action : {new_st}", "success" if ok else "danger")
    return redirect(url_for('ebm_plan_action'))

@app.route('/ebm/plan-action/delete/<int:plan_id>', methods=['POST'])
@owner_required
def ebm_plan_action_delete(plan_id: int):
    ebm_mgr.delete_action_plan(plan_id)
    flash("Action EBM supprimée.", "success")
    return redirect(url_for('ebm_plan_action'))

@app.route('/ebm/settings')
@owner_required
def ebm_settings():
    active_file = ebm_mgr.get_active_excel_path()
    return render_template('ebm_settings.html', active_file=active_file)

@app.route('/ebm/settings/upload', methods=['POST'])
@owner_required
def ebm_settings_upload():
    f = request.files.get('ebm_file')
    if f and f.filename:
        safe_name = f.filename.replace(" ", "_")
        dest_path = os.path.join(ebm_mgr.uploads_dir, safe_name)
        f.save(dest_path)
        ebm_mgr.set_setting("dashboard_ebm", dest_path)
        flash(f"Fichier EBM '{safe_name}' importé et activé avec succès.", "success")
    else:
        flash("Veuillez sélectionner un fichier Excel valide.", "danger")
    return redirect(url_for('ebm_settings'))


# =============================================================
# PASSATION DE SHIFT MODULE (OWNER ONLY)
# =============================================================

@app.route('/passation')
@login_required
def passation_dashboard():
    passations = passation_mgr.get_passations()
    return render_template('passation_dashboard.html', passations=passations)

@app.route('/passation/new')
@login_required
def passation_new():
    questions = passation_mgr.get_questions(active_only=True)
    return render_template('passation_form.html', questions=questions)

@app.route('/passation/save', methods=['POST'])
@login_required
def passation_save():
    shift = request.form.get('shift', '')
    target_shift = request.form.get('target_shift', '')
    zone_name = request.form.get('zone_name', '')
    tech_name = request.form.get('technician_name', '') or session.get('name', '')
    tech_mat = request.form.get('technician_matricule', '') or session.get('matricule', '')
    remarks = request.form.get('remarks', '')

    answers = {}
    for key, val in request.form.items():
        if key.startswith('question_'):
            try:
                qid = int(key.replace('question_', ''))
                answers[qid] = val
            except ValueError:
                pass

    pass_id = passation_mgr.create_passation(
        user_id=session.get('user_id'),
        technician_name=tech_name,
        technician_matricule=tech_mat,
        shift=shift,
        target_shift=target_shift,
        zone_name=zone_name,
        remarks=remarks,
        answers=answers
    )
    flash(f"Passation de shift #{pass_id} enregistrée avec succès.", "success")
    return redirect(url_for('passation_dashboard'))

@app.route('/passation/<int:pass_id>')
@login_required
def passation_detail(pass_id: int):
    detail = passation_mgr.get_passation_detail(pass_id)
    if not detail:
        flash("Passation introuvable.", "danger")
        return redirect(url_for('passation_dashboard'))
    return render_template('passation_detail.html', passation=detail)

@app.route('/passation/export/<int:pass_id>')
@login_required
def passation_export(pass_id: int):
    out_file = passation_mgr.export_passation_excel(pass_id)
    if out_file and os.path.isfile(out_file):
        return send_file(out_file, as_attachment=True, download_name=os.path.basename(out_file))
    flash("Erreur lors de la génération du rapport Excel de passation.", "danger")
    return redirect(url_for('passation_dashboard'))

@app.route('/passation/questions')
@admin_required
def passation_questions():
    questions = passation_mgr.get_questions(active_only=False)
    return render_template('passation_questions.html', questions=questions)

@app.route('/passation/questions/add', methods=['POST'])
@admin_required
def passation_question_add():
    text = request.form.get('text', '').strip()
    category = request.form.get('category', 'Général').strip()
    qtype = request.form.get('type', 'CHOICE').strip()
    options = request.form.get('options', '').strip()
    if text:
        passation_mgr.add_question(text, category, qtype, options)
        flash("Point de passation ajouté.", "success")
    else:
        flash("Le texte de la question est obligatoire.", "danger")
    return redirect(url_for('passation_questions'))

@app.route('/passation/questions/edit/<int:qid>', methods=['POST'])
@admin_required
def passation_question_edit(qid: int):
    text = request.form.get('text', '').strip()
    category = request.form.get('category', 'Général').strip()
    qtype = request.form.get('type', 'CHOICE').strip()
    options = request.form.get('options', '').strip()
    sort_order = int(request.form.get('sort_order', 1))
    if text:
        ok = passation_mgr.update_question(qid, text, category, qtype, options, sort_order)
        flash("Point de passation modifié avec succès.", "success" if ok else "danger")
    else:
        flash("Le texte de la question est obligatoire.", "danger")
    return redirect(url_for('passation_questions'))

@app.route('/passation/questions/delete/<int:qid>', methods=['POST'])
@admin_required
def passation_question_delete(qid: int):
    ok = passation_mgr.delete_question(qid)
    flash("Point de passation supprimé.", "success" if ok else "danger")
    return redirect(url_for('passation_questions'))

@app.route('/passation/questions/toggle/<int:qid>', methods=['POST'])
@admin_required
def passation_question_toggle(qid: int):
    ok, _ = passation_mgr.toggle_question(qid)
    flash("Statut de la question mis à jour.", "success" if ok else "danger")
    return redirect(url_for('passation_questions'))

@app.route('/admin')
@admin_required
def admin():
    excel_path     = pma_config.get_last_excel_path()
    schedule_files = pma_config.list_schedule_files()
    asp_codes      = ima_db.get_asp_codes_enriched()
    all_techs      = tech_db.get_dashboard_summary()
    shift_config   = pma_config.get_shift_passwords()
    email_config   = pma_config.load_json(os.path.join(pma_config.active_base, "email_config.json"))

    ppe_dir       = pma_config.dirs.get('ppe', os.path.join(pma_config.active_base, 'PPE_Templates'))
    ppe_templates = list_all_templates(ppe_dir)
    checklists    = checklist_mgr.get_all_definitions_admin()
    chk_history   = checklist_mgr.get_execution_history(limit=30)

    # Load all signatures/validation history from Excel schedules
    de = DataEngine()
    signatures = []
    for fp in pma_config.get_all_excel_paths():
        try:
            sigs = de.get_excel_signatures(fp)
            for s in sigs:
                s['file'] = os.path.basename(fp)
            signatures.extend(sigs)
        except Exception:
            pass
    signatures.reverse()

    # Exam management data for admin panel
    all_exams_admin = exam_mgr.get_all_exams(role='ADMIN')
    technicians_for_override = exam_mgr.get_all_technicians_for_override()

    return render_template(
        'admin.html',
        excel_path=excel_path,
        schedule_files=schedule_files,
        asp_codes=asp_codes,
        all_techs=all_techs,
        ppe_templates=ppe_templates,
        checklists=checklists,
        checklist_history=chk_history,
        shift_config=shift_config,
        email_config=email_config,
        signatures=signatures,
        all_exams_admin=all_exams_admin,
        technicians_for_override=technicians_for_override
    )

@app.route('/admin/credentials', methods=['POST'])
@admin_required
def admin_credentials():
    u = request.form.get('username', '').strip()
    p = request.form.get('password', '').strip()
    if u and p:
        pma_config.save_json(os.path.join(pma_config.active_base, "admin_config.json"), {"username": u, "password": p})
        ima_config.set_admin_credentials(u, p)
        # Invalidate any other active admin sessions
        owner_security.invalidate_all_admin_sessions()
        # Keep acting admin session valid by updating their login timestamp
        session['login_time'] = time.time()
        flash("Identifiants Administrateur mis à jour avec succès.", "success")
    else:
        flash("Nom d'utilisateur et mot de passe requis.", "danger")
    return redirect(url_for('admin'))

@app.route('/admin/shifts', methods=['POST'])
@admin_required
def admin_shifts():
    a = request.form.get('A', '')
    b = request.form.get('B', '')
    c = request.form.get('C', '')
    pma_config.save_json(os.path.join(pma_config.active_base, "shift_config.json"), {"A": a, "B": b, "C": c})
    flash("Mots de passe des équipes mis à jour.", "success")
    return redirect(url_for('admin'))

@app.route('/admin/email', methods=['POST'])
@admin_required
def admin_email():
    s = request.form.get('sender', '')
    p = request.form.get('password', '')
    r = request.form.get('recipient', '')
    pma_config.save_json(os.path.join(pma_config.active_base, "email_config.json"), {"sender": s, "password": p, "recipient": r})
    flash("Configuration email mise à jour.", "success")
    return redirect(url_for('admin'))

@app.route('/admin/asp-code/add', methods=['POST'])
@admin_required
def admin_asp_add():
    c = request.form.get('code', '').strip()
    d = request.form.get('description', '').strip()
    if c and d:
        ima_db.upsert_asp_codes([{"code": c, "description": d}])
        flash("Code ASP ajouté.", "success")
    else:
        flash("Code et description requis.", "danger")
    return redirect(url_for('admin'))

@app.route('/admin/asp-code/<code>/delete', methods=['POST'])
@admin_required
def admin_asp_delete(code):
    try:
        with ima_db._conn() as conn:
            conn.execute("DELETE FROM asp_codes WHERE code = ?", (code,))
        flash(f"Code ASP '{code}' supprimé.", "success")
    except Exception as e:
        flash(f"Erreur lors de la suppression : {e}", "danger")
    return redirect(url_for('admin'))

@app.route('/admin/ppe-template', methods=['POST'])
@admin_required
def admin_ppe_upload():
    file     = request.files.get('ppe_file')
    key_name = request.form.get('template_name', '').strip()
    if not file or not file.filename:
        flash("Aucun fichier sélectionné.", "danger")
        return redirect(url_for('admin'))
    if not key_name:
        key_name = os.path.splitext(file.filename)[0]
    # Keep original extension
    orig_ext  = os.path.splitext(file.filename)[1].lower() or '.xlsx'
    safe_name = key_name.replace('/', '_').replace('\\', '_').strip() + orig_ext
    dest      = os.path.join(pma_config.dirs.get('ppe', pma_config.active_base), safe_name)
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    file.save(dest)
    flash(f"Template PPE '{safe_name}' uploadé avec succès.", "success")
    return redirect(url_for('admin'))

@app.route('/admin/ppe-template/<path:filename>/delete', methods=['POST'])
@admin_required
def admin_ppe_delete(filename):
    ppe_dir   = pma_config.dirs.get('ppe', os.path.join(pma_config.active_base, 'PPE_Templates'))
    safe_path = os.path.realpath(os.path.join(ppe_dir, filename))
    base_real = os.path.realpath(ppe_dir)
    if not safe_path.startswith(base_real):
        flash("Opération non autorisée.", "danger")
        return redirect(url_for('admin'))
    if os.path.exists(safe_path):
        os.remove(safe_path)
        flash(f"Template '{filename}' supprimé.", "success")
    else:
        flash("Fichier introuvable.", "danger")
    return redirect(url_for('admin'))

@app.route('/admin/tech/add', methods=['POST'])
@admin_required
def admin_tech_add():
    name  = request.form.get('name', '').strip()
    mat   = request.form.get('matricule', '').strip()
    shift = request.form.get('shift', '').strip()
    if not name or not mat:
        flash("Nom et matricule requis.", "danger")
        return redirect(url_for('admin'))
    if tech_db.get_profile(mat):
        flash(f"Un technicien avec le matricule {mat} existe déjà.", "warning")
        return redirect(url_for('admin'))
    tech_db.update_tech({
        "name":      name,
        "matricule": mat,
        "shift":     shift,
        "hire_date": "",
        "exams":     []
    })
    flash(f"Technicien {name} ajouté avec succès.", "success")
    return redirect(url_for('admin'))

@app.route('/admin/tech/<matricule>/delete', methods=['POST'])
@admin_required
def admin_tech_delete(matricule):
    safe_mat = str(matricule).replace('/', '_').replace('\\', '_').strip()
    path = os.path.join(tech_db.tech_dir, f"{safe_mat}.json")
    if os.path.exists(path):
        os.remove(path)
        flash(f"Technicien {matricule} supprimé.", "success")
    else:
        flash("Technicien introuvable.", "danger")
    return redirect(url_for('admin'))

# =============================================================
# IMA — INTERVENTION ROUTES
# =============================================================

@app.route('/interventions')
@login_required
def interventions():
    status      = request.args.get('status', 'All')
    search      = request.args.get('search', '').lower().strip()
    date_from   = request.args.get('date_from', '')
    date_to     = request.args.get('date_to', '')
    page        = max(1, int(request.args.get('page', 1)))
    per_page    = 25

    ints = ima_db.get_all_interventions(status if status != 'All' else None)

    if search:
        ints = [i for i in ints if
                search in str(i.get('code', '')).lower() or
                search in str(i.get('machine_id', '')).lower() or
                search in str(i.get('machine_name', '')).lower() or
                search in str(i.get('fault_description', '')).lower() or
                search in str(i.get('technician_name', '')).lower()]

    if date_from:
        ints = [i for i in ints if str(i.get('created_at', ''))[:10] >= date_from]
    if date_to:
        ints = [i for i in ints if str(i.get('created_at', ''))[:10] <= date_to]

    total        = len(ints)
    total_pages  = max(1, (total + per_page - 1) // per_page)
    page         = min(page, total_pages)
    offset       = (page - 1) * per_page
    paged_ints   = ints[offset: offset + per_page]

    return render_template(
        'interventions.html',
        interventions=paged_ints,
        current_status=status, current_search=search,
        date_from=date_from, date_to=date_to,
        page=page, total_pages=total_pages, total=total
    )

@app.route('/interventions/new', methods=['GET', 'POST'])
@login_required
def intervention_form():
    if request.method == 'POST':
        machine_id   = request.form.get('machine_id', '').strip()
        group_name   = request.form.get('group_name', '').strip()
        machine_name = request.form.get('machine_name', machine_id).strip()

        data = {
            'machine_id':        machine_id,
            'machine_name':      machine_name or machine_id,
            'group_name':        group_name,
            'fault_description': request.form.get('fault_description', '').strip(),
            'fault_type':        request.form.get('fault_type', ''),
            'code_asp':          request.form.get('asp_code', ''),
            'technician_name':   request.form.get('technician_name', '').strip(),
            'technician_mat':    request.form.get('technician_mat', '').strip(),
            'start_time':        request.form.get('start_time') or datetime.datetime.now().strftime('%Y-%m-%dT%H:%M'),
            'shift':             request.form.get('shift', session.get('shift', '')),
            'priority':          request.form.get('priority', 'MEDIUM'),
            'category':          request.form.get('category', ''),
            'remarks':           request.form.get('remarks', '').strip(),
        }

        # Technicians can only create under their own name
        if session.get('role') != 'admin':
            data['technician_name'] = session.get('name', '')
            data['technician_mat']  = session.get('matricule', '')

        try:
            code = ima_db.create_intervention(data)
        except Exception as e:
            logger.error(f"Create intervention error: {e}")
            flash(f"Erreur lors de la création : {e}", "danger")
            groups    = ima_db.get_machine_groups()
            asp_codes = ima_db.get_asp_codes_enriched()
            machs     = ima_db.get_all_machines()
            return render_template('intervention_form.html', groups=groups, asp_codes=asp_codes,
                                   machines=machs, now=datetime.datetime.now().strftime('%Y-%m-%dT%H:%M'), form_data=data)

        end_time = request.form.get('end_time', '').strip()
        if end_time:
            dt = calc_downtime(data['start_time'], end_time)
            ima_db.close_intervention(code, end_time, dt, data.get('remarks', ''))

        flash(f"Intervention {code} créée avec succès.", "success")
        return redirect(url_for('intervention_detail', code=code))

    # --- Multi-layer machine loading — always shows machines no matter what ---
    machs = ima_db.get_all_machines()

    # Layer 1: if DB has no machines, sync from Excel then retry
    if not machs:
        try:
            sync_pma_machines_to_ima()
            machs = ima_db.get_all_machines()
        except Exception as _e:
            logger.error(f"Layer-1 sync failed: {_e}")

    # Layer 2: if still empty, read machines DIRECTLY from Excel (bypass DB entirely)
    if not machs:
        try:
            from ima.excel_reader import CalendrierReader
            _reader = CalendrierReader()
            _paths = pma_config.get_all_excel_paths()
            # Also check the app's own bundled data/ folder explicitly
            _app_data = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
            for _f in ['current_schedule.xlsx', 'calendrier.xlsx', 'planning.xlsx']:
                _fp = os.path.join(_app_data, _f)
                if os.path.isfile(_fp) and _fp not in _paths:
                    _paths.append(_fp)
            for _p in _paths:
                _cdf = _reader.read_calendrier(_p)
                if not _cdf.empty:
                    for _, _r in _cdf.iterrows():
                        _mid = str(_r.get('ID Machine', '')).strip()
                        _mname = str(_r.get('Nom Machine', _mid)).strip()
                        _grp = str(_r.get('Groupe', '')).strip()
                        if _mid and _mid.lower() not in ['nan', 'none', '']:
                            machs.append({
                                'machine_id': _mid,
                                'machine_name': _mname if _mname and _mname.lower() != 'nan' else _mid,
                                'group_name': _grp if _grp and _grp.lower() != 'nan' else 'Général',
                            })
                    if machs:
                        break
            # Try to persist these into DB for next time
            if machs:
                try:
                    ima_db.upsert_machines([{**m, 'location': '', 'description': ''} for m in machs])
                except Exception:
                    pass
        except Exception as _e:
            logger.error(f"Layer-2 direct Excel read failed: {_e}")

    # Layer 3: if still empty, get from the DataEngine (PMA engine that drives the calendar)
    if not machs:
        try:
            _eng = get_pma_engine()
            if _eng.current_df is not None and not _eng.current_df.empty:
                _df = _eng.current_df
                _col = 'Equipment' if 'Equipment' in _df.columns else None
                if _col:
                    _seen = set()
                    for _, _r in _df.iterrows():
                        _mid = str(_r.get(_col, '')).strip()
                        if _mid and _mid.lower() not in ['nan', 'none', ''] and _mid not in _seen:
                            _seen.add(_mid)
                            _grp = str(_r.get('Zone', _r.get('Sheet', 'Général'))).strip()
                            machs.append({
                                'machine_id': _mid,
                                'machine_name': _mid,
                                'group_name': _grp if _grp and _grp.lower() != 'nan' else 'Général',
                            })
        except Exception as _e:
            logger.error(f"Layer-3 DataEngine fallback failed: {_e}")

    # Deduplicate
    _seen_ids = set()
    machs = [m for m in machs if m['machine_id'] not in _seen_ids and not _seen_ids.add(m['machine_id'])]

    groups = sorted({m['group_name'] for m in machs if m.get('group_name')}) if machs else []
    asp_codes = ima_db.get_asp_codes_enriched()
    logger.info(f"intervention_form: rendering with {len(machs)} machines, {len(groups)} groups")
    return render_template('intervention_form.html', groups=groups, asp_codes=asp_codes,
                           machines=machs, now=datetime.datetime.now().strftime('%Y-%m-%dT%H:%M'), form_data={})

@app.route('/api/machines-by-group')
@login_required
def api_machines_by_group():
    grp = request.args.get('group', '').strip()
    try:
        if not ima_db.get_all_machines():
            sync_pma_machines_to_ima()
        if grp:
            machines = ima_db.get_machines_by_group(grp)
        else:
            machines = ima_db.get_all_machines()
        return jsonify(machines)
    except Exception as e:
        logger.error(f"api_machines_by_group error: {e}")
        return jsonify([]), 200

@app.route('/interventions/<code>')
@login_required
def intervention_detail(code):
    inv = ima_db.get_intervention(code)
    if not inv:
        flash("Intervention introuvable.", "danger")
        return redirect(url_for('interventions'))
    return render_template('intervention_detail.html', intervention=inv,
                           now=datetime.datetime.now().strftime('%Y-%m-%dT%H:%M'))

@app.route('/interventions/<code>/edit', methods=['GET', 'POST'])
@admin_required
def intervention_edit(code):
    inv = ima_db.get_intervention(code)
    if not inv:
        flash("Intervention introuvable.", "danger")
        return redirect(url_for('interventions'))

    if request.method == 'POST':
        updated = {
            'machine_id':        request.form.get('machine_id', inv.get('machine_id', '')).strip(),
            'group_name':        request.form.get('group_name', inv.get('group_name', '')).strip(),
            'fault_description': request.form.get('fault_description', '').strip(),
            'code_asp':          request.form.get('asp_code', inv.get('code_asp', '')),
            'technician_name':   request.form.get('technician_name', '').strip(),
            'downtime_minutes':  float(request.form.get('downtime_minutes', inv.get('downtime_minutes', 0)) or 0),
            'category':          request.form.get('category', inv.get('category', '')),
            'remarks':           request.form.get('remarks', '').strip(),
            'status':            request.form.get('status', inv.get('status', 'OPEN')),
            'start_time':        request.form.get('start_time', inv.get('start_time', '')),
            'end_time':          request.form.get('end_time', inv.get('end_time', '')) or None,
            'created_at':        inv.get('created_at', ''),
        }
        try:
            ima_db.update_intervention(code, updated)
            flash(f"Intervention {code} mise à jour.", "success")
            return redirect(url_for('intervention_detail', code=code))
        except Exception as e:
            logger.error(f"Update intervention error: {e}")
            flash(f"Erreur lors de la mise à jour : {e}", "danger")

    groups    = ima_db.get_machine_groups()
    asp_codes = ima_db.get_asp_codes_enriched()
    return render_template('intervention_edit.html', intervention=inv,
                           groups=groups, asp_codes=asp_codes,
                           now=datetime.datetime.now().strftime('%Y-%m-%dT%H:%M'))

@app.route('/interventions/<code>/close', methods=['POST'])
@login_required
def close_intervention(code):
    inv = ima_db.get_intervention(code)
    if not inv:
        flash("Intervention introuvable.", "danger")
        return redirect(url_for('interventions'))

    end_time = request.form.get('end_time', '').strip() or datetime.datetime.now().strftime('%Y-%m-%dT%H:%M')
    remarks  = request.form.get('remarks', '').strip()

    dt = calc_downtime(inv.get('start_time', ''), end_time)

    try:
        ima_db.close_intervention(code, end_time, dt, remarks)
    except Exception as e:
        logger.error(f"Close intervention error: {e}")
        flash(f"Erreur lors de la clôture : {e}", "danger")
        return redirect(url_for('intervention_detail', code=code))

    # Email alert for long interventions (> 2 hours)
    if dt > 120:
        try:
            email_cfg = pma_config.load_json(os.path.join(pma_config.active_base, "email_config.json"))
            recipients = [email_cfg.get('recipient')] if email_cfg.get('recipient') else []
            recipients += email_cfg.get('recipients', [])
            sender, pw = email_cfg.get('sender', ''), email_cfg.get('password', '')
            if sender and pw and recipients:
                subject = f"Alerte: Intervention longue ({code}) — {round(dt/60, 1)}h"
                body    = (f"L'intervention {code} clôturée après {int(dt)} min.\n"
                           f"Technicien : {inv.get('technician_name', '')}\n"
                           f"Machine    : {inv.get('machine_id', '')}\n"
                           f"Remarques  : {remarks}")
                for recip in set(recipients):
                    send_simple_alert(sender, pw, recip, subject, body)
        except Exception as e:
            logger.warning(f"Email send error: {e}")

    flash("Intervention clôturée avec succès.", "success")
    return redirect(url_for('intervention_detail', code=code))

@app.route('/interventions/<code>/delete', methods=['POST'])
@admin_required
def delete_intervention(code):
    try:
        ima_db.delete_intervention(code)
        flash(f"Intervention {code} supprimée.", "success")
    except Exception as e:
        flash(f"Erreur lors de la suppression : {e}", "danger")
    return redirect(url_for('interventions'))

@app.route('/interventions/<code>/print')
@login_required
def print_intervention(code):
    inv = ima_db.get_intervention(code)
    if not inv:
        flash("Intervention introuvable.", "danger")
        return redirect(url_for('interventions'))
    return render_template('intervention_print.html', intervention=inv)

# =============================================================
# MACHINES
# =============================================================

@app.route('/machines')
@login_required
def machines():
    group  = request.args.get('group', 'All')
    search = request.args.get('search', '').strip()

    machs = ima_db.search_machines(search, group if group != 'All' else None)
    if not machs and not search and group == 'All':
        sync_pma_machines_to_ima()
        machs = ima_db.search_machines(search, group if group != 'All' else None)

    enriched = ima_db.get_machines_enriched()
    e_dict   = {m['machine_id']: m for m in enriched}

    final_machs = []
    for m in machs:
        m.update(e_dict.get(m['machine_id'], {}))
        final_machs.append(m)

    groups = ima_db.get_machine_groups()
    return render_template('machines.html', machines=final_machs, groups=groups,
                           current_group=group, current_search=search)

@app.route('/machines/sync-pma', methods=['POST', 'GET'])
@login_required
def machines_sync_pma():
    upserted, deleted = replace_pma_machines_to_ima()
    if upserted > 0:
        msg = f"{upserted} machines synchronisées depuis le planning PMA."
        if deleted > 0:
            msg += f" {deleted} machine(s) absente(s) du nouveau fichier supprimée(s) du parc."
        flash(msg, "success")
    else:
        flash("Aucune machine trouvée dans le planning actuel. Vérifiez que le fichier Excel est bien chargé.", "warning")
    return redirect(url_for('machines'))

@app.route('/machines/import', methods=['POST'])
@admin_required
def machines_import():
    file = request.files.get('excel_file')
    if file and file.filename:
        safe_filename = file.filename.replace('/', '_').replace('\\', '_').strip()
        path = os.path.join(pma_config.active_base, safe_filename)
        file.save(path)
        pma_config.set_last_excel_path(path)
        invalidate_pma_cache()
        try:
            eng = get_pma_engine(force=True)
            upserted, deleted = replace_pma_machines_to_ima(eng.current_df)
            msg = f"{upserted} machines importées et synchronisées avec succès."
            if deleted > 0:
                msg += f" {deleted} machine(s) absente(s) du nouveau fichier supprimée(s) du parc."
            flash(msg, "success")
        except Exception as e:
            logger.error(f"Machines import error: {e}")
            flash(f"Erreur lors de l'importation : {e}", "danger")
    return redirect(url_for('machines'))

@app.route('/machines/<machine_id>')
@login_required
def machine_detail(machine_id):
    if machine_id in ('sync', 'sync-pma', 'import'):
        return redirect(url_for('machines_sync_pma'))

    history  = ima_db.get_machine_history(machine_id)
    enriched = [m for m in ima_db.get_machines_enriched() if m['machine_id'] == machine_id]

    if not enriched:
        flash("Machine introuvable.", "danger")
        return redirect(url_for('machines'))

    m            = enriched[0]
    total_dt     = sum(i.get('downtime_minutes', 0) or 0 for i in history)
    open_i       = len([i for i in history if i.get('status') == 'OPEN'])
    m['total_downtime_hours'] = round(total_dt / 60, 2)
    m['open_interventions']   = open_i

    return render_template('machine_detail.html', machine=m, history=history)

# =============================================================
# ADMIN DASHBOARD & SETTINGS ROUTES
# =============================================================

@app.route('/admin/update-config', methods=['POST'])
@admin_required
def admin_update_config():
    config_path = os.path.join(pma_config.active_base, "app_settings.json")
    cfg = pma_config.load_json(config_path, {})
    cfg['company_name'] = request.form.get("company_name", "SEBN-TN").strip()
    cfg['app_title'] = request.form.get("app_title", "Enterprise Maintenance Suite").strip()
    try:
        cfg['session_timeout'] = int(request.form.get("session_timeout", 60))
    except ValueError:
        pass
    pma_config.save_json(config_path, cfg)
    flash("Paramètres généraux enregistrés.", "success")
    return redirect(url_for('admin'))

@app.route('/admin/change-password', methods=['POST'])
@admin_required
def admin_change_password():
    curr_pw = request.form.get('current_password', '')
    new_pw  = request.form.get('new_password', '')
    conf_pw = request.form.get('confirm_password', '')
    creds   = pma_config.get_admin_credentials()

    if curr_pw != creds.get('password'):
        flash("Mot de passe actuel incorrect.", "danger")
        return redirect(url_for('admin'))
    if not new_pw or new_pw != conf_pw:
        flash("Le nouveau mot de passe et sa confirmation ne correspondent pas.", "danger")
        return redirect(url_for('admin'))

    creds['password'] = new_pw
    admin_path = os.path.join(pma_config.active_base, "admin_config.json")
    pma_config.save_json(admin_path, creds)
    flash("Mot de passe administrateur modifié avec succès.", "success")
    return redirect(url_for('admin'))

@app.route('/admin/update-email', methods=['POST'])
@admin_required
def admin_update_email():
    email_path = os.path.join(pma_config.active_base, "email_config.json")
    data = {
        "smtp_server": request.form.get("smtp_server", "smtp.gmail.com").strip(),
        "smtp_port": int(request.form.get("smtp_port", 587)),
        "sender": request.form.get("sender_email", "").strip(),
        "password": request.form.get("sender_password", "").strip(),
        "recipient": request.form.get("sender_email", "").strip()
    }
    pma_config.save_json(email_path, data)
    flash("Paramètres de messagerie mis à jour.", "success")
    return redirect(url_for('admin'))

@app.route('/admin/upload-ppe', methods=['POST'])
@admin_required
def upload_ppe_template():
    file = request.files.get('ppe_file')
    chk_name = request.form.get('checklist_name', '').strip()
    chk_code = request.form.get('checklist_code', '').strip()
    eq_pattern = request.form.get('equipment_pattern', 'ALL').strip()
    change_sum = request.form.get('change_summary', '').strip()

    if file and file.filename:
        filename = file.filename.replace('/', '_').replace('\\', '_').strip()
        if filename.lower().endswith(('.xlsx', '.xls', '.xlsm')):
            temp_path = os.path.join(pma_config.dirs['ppe'], f"tmp_{uuid.uuid4().hex[:8]}_{filename}")
            file.save(temp_path)
            try:
                ok, msg, new_vid = checklist_mgr.validate_and_import_excel(
                    file_path=temp_path,
                    original_filename=filename,
                    name=chk_name or None,
                    checklist_code=chk_code or None,
                    equipment_pattern=eq_pattern or 'ALL',
                    imported_by=session.get('name', 'admin'),
                    change_summary=change_sum
                )
                if ok:
                    flash(msg, "success")
                else:
                    flash(msg, "danger")
            finally:
                if os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except:
                        pass
        else:
            flash("Format de fichier non supporté (utilisez .xlsx ou .xls).", "danger")
    else:
        flash("Aucun fichier sélectionné.", "danger")
    return redirect(url_for('admin'))

@app.route('/admin/toggle-ppe', methods=['POST'])
@admin_required
def toggle_ppe_template():
    version_id = request.form.get('version_id', type=int)
    action = request.form.get('action', 'deactivate')
    if version_id:
        if action == 'activate':
            ok, msg = checklist_mgr.activate_version(version_id)
            flash(msg, "success" if ok else "danger")
        else:
            ok, msg = checklist_mgr.deactivate_version(version_id)
            flash(msg, "warning" if ok else "danger")
    return redirect(url_for('admin'))

@app.route('/admin/download-ppe/<int:version_id>')
@admin_required
def download_ppe_file(version_id):
    with checklist_mgr._get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT stored_filename, original_filename FROM checklist_versions WHERE id = ?", (version_id,))
        row = cur.fetchone()
        if row:
            arch_path = os.path.join(checklist_mgr.archive_dir, row['stored_filename'])
            if os.path.exists(arch_path):
                return send_file(arch_path, as_attachment=True, download_name=row['original_filename'])
            # Fallback to direct storage dir
            direct_path = os.path.join(checklist_mgr.storage_dir, row['original_filename'])
            if os.path.exists(direct_path):
                return send_file(direct_path, as_attachment=True, download_name=row['original_filename'])
    flash("Fichier source introuvable.", "danger")
    return redirect(url_for('admin'))

@app.route('/admin/checklist-items/<int:version_id>')
@admin_required
def get_checklist_items_json(version_id):
    items = checklist_mgr.get_items_for_version(version_id)
    return jsonify({"items": items})

@app.route('/admin/add-asp', methods=['POST'])
@admin_required
def add_asp_code():
    code = request.form.get('code', '').strip()
    desc = request.form.get('description', '').strip()
    if code and desc:
        if ima_db.add_asp_code(code, desc):
            flash(f"Code ASP '{code}' ajouté.", "success")
        else:
            flash("Erreur ou code ASP déjà existant.", "danger")
    else:
        flash("Le code et la description sont requis.", "danger")
    return redirect(url_for('admin'))

@app.route('/admin/delete-asp', methods=['POST'])
@admin_required
def delete_asp_code():
    code = request.form.get('code', '').strip()
    if code:
        if ima_db.delete_asp_code(code):
            flash(f"Code ASP '{code}' supprimé.", "warning")
        else:
            flash("Erreur lors de la suppression.", "danger")
    return redirect(url_for('admin'))

@app.route('/settings/reset_db', methods=['POST'])
@admin_required
def reset_db():
    try:
        ima_db.clear_all_data()
        flash("Base de données curative réinitialisée (machines conservées).", "warning")
    except Exception as e:
        flash(f"Erreur lors de la réinitialisation : {e}", "danger")
    return redirect(url_for('admin'))

@app.route('/settings/purge_pma', methods=['POST'])
@admin_required
def purge_pma():
    try:
        paths = pma_config.get_all_excel_paths()
        for p in paths:
            if os.path.exists(p):
                os.remove(p)
        pma_config.set_last_excel_path("")
        invalidate_pma_cache()
        flash("Toutes les données PMA et plannings ont été purgés avec succès.", "warning")
    except Exception as e:
        flash(f"Erreur lors de la purge : {e}", "danger")
    return redirect(url_for('admin'))

@app.route('/export/excel')
@admin_required
def export_excel():
    try:
        filepath = export_interventions_to_excel(ima_db)
        if filepath and os.path.exists(filepath):
            return send_file(filepath, as_attachment=True)
    except Exception as e:
        logger.error(f"Excel export error: {e}")
    flash("Aucune donnée à exporter ou erreur lors de l'exportation.", "warning")
    return redirect(url_for('admin'))

@app.route('/export/pdf')
@admin_required
def export_pdf():
    try:
        pma_engine = get_pma_engine()
        filepath   = export_dashboard_pdf(ima_db, pma_engine)
        if filepath and os.path.exists(filepath):
            return send_file(filepath, as_attachment=True)
    except Exception as e:
        logger.error(f"PDF export error: {e}")
    flash("Erreur lors de la génération du PDF.", "danger")
    return redirect(url_for('admin'))

# =============================================================
# PMA ROUTES
# =============================================================

@app.route('/pma')
@login_required
def pma_dashboard():
    import re as _re

    current_iso_w       = datetime.datetime.now().isocalendar()[1]
    current_iso_w_label = f"S{current_iso_w}"

    month  = request.args.get('month', '')
    sheet  = request.args.get('sheet', '')
    # Default to current week if no week filter is provided
    week   = request.args.get('week', current_iso_w_label)
    search = request.args.get('search', '').lower().strip()
    # Track if user explicitly cleared week filter
    no_week = 'week' in request.args and request.args.get('week', '') == ''
    if no_week:
        week = ''

    eng = get_pma_engine()
    df  = eng.current_df

    tasks = []
    if df is not None and not df.empty:
        col_month = 'Month' if 'Month' in df.columns else ('month' if 'month' in df.columns else None)
        col_sheet = 'Sheet' if 'Sheet' in df.columns else ('sheet' if 'sheet' in df.columns else None)
        col_week  = 'Week'  if 'Week'  in df.columns else ('week'  if 'week'  in df.columns else None)
        col_equip = 'Equipment' if 'Equipment' in df.columns else 'equipment'
        col_type  = 'Type'  if 'Type'  in df.columns else 'type'
        col_stat  = 'Status' if 'Status' in df.columns else 'status'
        col_ri    = 'Raw_Index' if 'Raw_Index' in df.columns else 'raw_index'

        filtered = df
        if month and col_month:
            filtered = filtered[filtered[col_month].astype(str) == month]
        if sheet and col_sheet:
            filtered = filtered[filtered[col_sheet].astype(str) == sheet]
        if week and col_week:
            # Normalize: digits only match so S33, 33, KW33, W33 all match week 33
            week_digits = _re.sub(r'\D', '', str(week))
            if week_digits:
                filtered = filtered[
                    filtered[col_week].astype(str).str.replace(r'\D', '', regex=True) == week_digits
                ]
            else:
                filtered = filtered[filtered[col_week].astype(str) == week]

        records = filtered.to_dict('records')

        for r in records:
            equip = str(r.get(col_equip, '')).strip()
            type_ = str(r.get(col_type, '')).strip()
            sheet_name = str(r.get(col_sheet, '') if col_sheet else '').strip()

            # Strict guard: skip any metadata / user / signature rows
            if not equip or equip.upper() in ['ADMIN', 'ROLE', 'TECHNICIAN', 'TECHNICIEN', 'SIGNATURES', 'SIGNATURE', 'USER', 'NAN'] or 'SIGNATURE' in sheet_name.upper():
                continue

            if search and search not in equip.lower() and search not in type_.lower():
                continue

            zone = str(r.get('Zone', sheet_name)).strip()
            carte = str(r.get('Carte', '')).strip()

            tasks.append({
                "equipment":    equip,
                "machine_name": str(r.get('Machine_Name', equip)).strip(),
                "zone":         zone if zone and zone.lower() != 'nan' else sheet_name,
                "group":        zone if zone and zone.lower() != 'nan' else sheet_name,
                "carte":        carte if carte and carte.lower() != 'nan' else '',
                "type":         type_,
                "week":         str(r.get(col_week, '') if col_week else ''),
                "month":        str(r.get(col_month, '') if col_month else ''),
                "status":       str(r.get(col_stat, '')),
                "sheet":        sheet_name,
                "raw_index":    str(r.get(col_ri, ''))
            })

        # Always sort: pending tasks first, completed tasks at bottom
        def sort_key(t):
            is_done = str(t.get('status', '')).upper().strip() in ('COMPLÉTÉ', 'COMPLETE', 'DONE', 'OK', 'TERMINÉ')
            w_str   = str(t.get('week', ''))
            digits  = "".join(filter(str.isdigit, w_str))
            w_num   = int(digits) if digits else 999
            return (1 if is_done else 0, w_num)

        tasks.sort(key=sort_key)

    is_filtered = bool(month or sheet or week or search)
    if tasks:
        t_total = len(tasks)
        t_done = sum(1 for t in tasks if str(t.get('status', '')).upper().strip() in ('COMPLÉTÉ', 'COMPLETE', 'DONE', 'OK', 'TERMINÉ'))
        t_pending = t_total - t_done
        t_rate = round((t_done / t_total) * 100) if t_total > 0 else 0
        stats = {
            'total': t_total,
            'done': t_done,
            'pending': t_pending,
            'rate': t_rate,
            'is_filtered': is_filtered
        }
    else:
        raw_stats = eng.get_stats()
        stats = {
            'total': raw_stats.get('total', 0) if not is_filtered else 0,
            'done': raw_stats.get('done', 0) if not is_filtered else 0,
            'pending': raw_stats.get('pending', 0) if not is_filtered else 0,
            'rate': raw_stats.get('rate', 0) if not is_filtered else 0,
            'is_filtered': is_filtered
        }

    months, sheets, weeks = [], [], []
    if eng.current_df is not None and not eng.current_df.empty:
        for m_col in ['Month', 'month']:
            if m_col in eng.current_df.columns:
                months = sorted([str(m) for m in eng.current_df[m_col].dropna().unique() if str(m).strip()])
                break
        for s_col in ['Sheet', 'sheet']:
            if s_col in eng.current_df.columns:
                sheets = sorted([
                    str(s) for s in eng.current_df[s_col].dropna().unique()
                    if str(s).strip() and not any(k in str(s).upper() for k in ['SIGNATURE', 'USER', 'ACCOUNT', 'LOGIN', 'TECH'])
                ])
                break
        for w_col in ['Week', 'week']:
            if w_col in eng.current_df.columns:
                weeks = sorted([str(w) for w in eng.current_df[w_col].dropna().unique() if str(w).strip()])
                break

    current_iso_w = datetime.datetime.now().isocalendar()[1]
    current_iso_w_label = f"S{current_iso_w}"

    return render_template('pma_dashboard.html',
                           tasks=tasks, stats=stats,
                           months=months, sheets=sheets, weeks=weeks,
                           current_month=month, current_sheet=sheet,
                           current_week=week, current_search=search,
                           current_iso_w_label=current_iso_w_label)

@app.route('/pma/upload', methods=['POST'])
@admin_required
def pma_upload():
    file = request.files.get('excel_file')
    if file and file.filename:
        safe_filename = file.filename.replace('/', '_').replace('\\', '_').strip()
        path = os.path.join(pma_config.active_base, safe_filename)
        file.save(path)
        pma_config.set_last_excel_path(path)
        invalidate_pma_cache()
        eng = get_pma_engine(force=True)
        upserted, deleted = replace_pma_machines_to_ima(eng.current_df)
        msg = f"Planning '{safe_filename}' ajouté avec succès ({upserted} machines synchronisées pour les interventions)."
        if deleted > 0:
            msg += f" {deleted} machine(s) absente(s) du nouveau fichier supprimée(s) du parc."
        flash(msg, "success")
    else:
        flash("Aucun fichier sélectionné.", "danger")
    return redirect(url_for('admin'))

@app.route('/pma/complete-task', methods=['POST'])
@login_required
def complete_task():
    data      = request.get_json() or {}
    sheet     = data.get('sheet', '')
    raw_index = data.get('raw_index', '')
    task_type = data.get('type', '')

    eng = get_pma_engine()
    try:
        ok = eng.complete_task(sheet, raw_index, task_type, session.get('name'), session.get('shift'))
        if ok:
            pma_file = pma_config.get_last_excel_path()
            if pma_file and os.path.exists(pma_file):
                eng.add_excel_signature(
                    pma_file,
                    session.get('name', 'Technicien'),
                    session.get('role', 'Technician'),
                    matricule=session.get('matricule', '')
                )
            invalidate_pma_cache()
            return jsonify({"status": "success"})
    except Exception as e:
        logger.error(f"Complete task error: {e}")
    return jsonify({"status": "error", "message": "Impossible de valider la tâche."}), 400

@app.route('/pma/reset-task', methods=['POST'])
@admin_required
def reset_task():
    data      = request.get_json() or {}
    sheet     = data.get('sheet', '')
    raw_index = data.get('raw_index', '')
    task_type = data.get('type', 'Monthly')

    eng = get_pma_engine()
    try:
        import ast
        if isinstance(raw_index, str):
            try:
                raw_index = ast.literal_eval(raw_index)
            except Exception:
                pass
        ok = eng.reset_task_status(sheet, raw_index, task_type)
        if ok:
            invalidate_pma_cache()
            return jsonify({"status": "success"})
    except Exception as e:
        logger.error(f"Reset task error: {e}")
    return jsonify({"status": "error", "message": "Impossible de réinitialiser la tâche."}), 400

@app.route('/pma/checklist')
@login_required
def checklist():
    sheet        = request.args.get('sheet', '')
    idx          = request.args.get('idx', '')
    type_        = request.args.get('type', '')
    equip        = request.args.get('equip', '')
    week         = request.args.get('week', '')
    selected_vid = request.args.get('version_id', type=int)

    active_templates = checklist_mgr.get_all_active_templates_list()

    chk_data = None
    if selected_vid:
        items = checklist_mgr.get_items_for_version(selected_vid)
        if items:
            with checklist_mgr._get_conn() as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT v.id as version_id, v.version_number, v.original_filename, d.name as checklist_name
                    FROM checklist_versions v JOIN checklist_definitions d ON v.definition_id = d.id
                    WHERE v.id = ?
                """, (selected_vid,))
                row = cur.fetchone()
                if row:
                    chk_data = dict(row)
                    chk_data["items"] = items

    template_param = request.args.get('template', '')
    if not chk_data and template_param != '__generic__':
        chk_data = checklist_mgr.get_active_checklist(equipment=equip, task_type=type_)

    if chk_data:
        tasks = chk_data.get("items", [])
        template_found = True
        active_version_id = chk_data.get("version_id")
        active_template_name = f"{chk_data.get('checklist_name')} (V{chk_data.get('version_number')})"
    else:
        tasks = GENERIC_TASKS
        template_found = False
        active_version_id = None
        active_template_name = "Checklist Générique (8 points)"

    # New checklist execution always starts 100% empty (no pre-filled answers)
    return render_template(
        'checklist.html',
        sheet=sheet, idx=idx, type=type_, equip=equip, week=week,
        tasks=tasks, template_found=template_found,
        all_templates=active_templates,
        active_version_id=active_version_id,
        active_template=active_template_name
    )

@app.route('/pma/checklist/save', methods=['POST'])
@login_required
def checklist_save():
    sheet        = request.form.get('sheet', '')
    idx          = request.form.get('idx', '')
    task_type    = request.form.get('type', '')
    equip        = request.form.get('equip', '')
    week         = request.form.get('week', '')
    version_id   = request.form.get('version_id', type=int)
    task_count   = int(request.form.get('task_count', 0))

    answers = {}
    for n in range(1, task_count + 1):
        status = request.form.get(f'status_{n}', '').strip()
        val    = request.form.get(f'val_{n}', '').strip()
        obs    = request.form.get(f'obs_{n}', '').strip()
        if status or val or obs:
            answers[str(n)] = {'status': status, 'val': val, 'obs': obs}

    # Fetch tasks from DB version or fallback
    if version_id:
        tasks = checklist_mgr.get_items_for_version(version_id)
    else:
        chk_data = checklist_mgr.get_active_checklist(equipment=equip, task_type=task_type)
        tasks = chk_data.get("items", []) if chk_data else GENERIC_TASKS
        if chk_data:
            version_id = chk_data.get("version_id")

    # Resolve master template file path
    template_path = checklist_mgr.get_template_file_path(
        version_id=version_id, equipment=equip, task_type=task_type
    )

    now_str    = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_equip = (equip or 'equip').replace('/', '_').replace(' ', '_').replace(':', '_')
    safe_week  = (week or 'week').replace('/', '_').replace(' ', '_')
    if template_path:
        orig_base = os.path.splitext(os.path.basename(template_path))[0]
        out_name = f"{orig_base}_{safe_equip}_{safe_week}_{now_str}.xlsx"
    else:
        out_name = f"checklist_{safe_equip}_{safe_week}_{now_str}.xlsx"

    filled_dir = pma_config.dirs.get('filled', os.path.join(pma_config.active_base, 'PPE_Filled'))
    os.makedirs(filled_dir, exist_ok=True)
    out_path   = os.path.join(filled_dir, out_name)

    metadata = {
        'technician':    session.get('name', 'Technicien'),
        'matricule':     session.get('matricule', ''),
        'shift':         session.get('shift', 'A'),
        'equip':         equip,
        'type':          task_type,
        'sheet':         sheet,
        'week':          week,
        'version_id':    version_id,
        'template_path': template_path or '',
    }

    # 1. PHYSICAL EXCEL SAVE (COPY MASTER TEMPLATE -> FILL CELLS -> SAVE COPY)
    excel_saved = False
    try:
        excel_saved = save_filled_checklist(tasks, answers, metadata, out_path)
    except Exception as e:
        logger.error(f"[EXCEL] Save exception: {e}", exc_info=True)
        excel_saved = False

    # 2. STRICT PHYSICAL FILE EXISTENCE VALIDATION
    if not excel_saved or not os.path.isfile(out_path) or os.path.getsize(out_path) == 0:
        logger.error(f"[EXCEL] Physical file validation failed: {out_path}")
        flash("Erreur lors de l'enregistrement du fichier Excel physique.", "danger")
        return redirect(url_for('pma_dashboard'))

    logger.info(f"[EXCEL] Physical file validated successfully: {out_path} ({os.path.getsize(out_path)} bytes)")

    # 3. RECORD EXECUTION IN DATABASE (ONLY AFTER PHYSICAL FILE VALIDATION)
    exec_id = None
    try:
        exec_id = checklist_mgr.record_execution(
            version_id=version_id,
            equipment=equip,
            task_type=task_type,
            sheet=sheet,
            week=week,
            month=datetime.datetime.now().strftime('%B'),
            technician_name=session.get('name', 'Technicien'),
            technician_matricule=session.get('matricule', ''),
            shift=session.get('shift', 'A'),
            answers=answers,
            filled_excel_path=out_path
        )
    except Exception as e:
        logger.error(f"[EXCEL] Record execution error: {e}")

    # 4. COMPLETE PMA TASK & ADD SIGNATURE
    eng = get_pma_engine()
    if eng.complete_task(sheet, idx, task_type, session.get('name'), session.get('shift')):
        pma_file = pma_config.get_last_excel_path()
        if pma_file and os.path.exists(pma_file):
            eng.add_excel_signature(
                pma_file,
                session.get('name', 'Technicien'),
                session.get('role', 'Technician'),
                matricule=session.get('matricule', '')
            )
        invalidate_pma_cache()

    flash(f"Checklist Excel enregistrée avec succès ({out_name}) — {len(answers)}/{len(tasks)} points validés.", "success")
    return redirect(url_for('pma_dashboard'))


# ─── Download completed Excel (filled copy of original master template) ───────
@app.route('/pma/checklist/download/<int:exec_id>')
@login_required
def checklist_download(exec_id: int):
    """
    Download the completed checklist Excel for a past execution.
    The file is the original master template with technician results filled in.
    If the physical file is missing, it is regenerated on-the-fly from the DB answers
    using the same master template copy-and-fill pipeline.
    """
    from flask import send_file as _send_file
    exec_record = checklist_mgr.get_execution_by_id(exec_id)
    if not exec_record:
        flash("Exécution introuvable.", "danger")
        return redirect(url_for('pma_dashboard'))

    filled_path = exec_record.get('filled_excel_path', '')
    if filled_path and os.path.isfile(filled_path):
        return _send_file(
            filled_path,
            as_attachment=True,
            download_name=os.path.basename(filled_path)
        )

    # File missing → regenerate from original master template + saved DB answers
    version_id  = exec_record.get('version_id')
    equip       = exec_record.get('equipment', '')
    task_type   = exec_record.get('task_type', '')
    week        = exec_record.get('week', '')
    answers_raw = exec_record.get('answers', {})

    if version_id:
        tasks = checklist_mgr.get_items_for_version(version_id)
    else:
        chk = checklist_mgr.get_active_checklist(equipment=equip, task_type=task_type)
        tasks = chk.get('items', []) if chk else GENERIC_TASKS
        if chk:
            version_id = chk.get('version_id')

    template_path = checklist_mgr.get_template_file_path(
        version_id=version_id, equipment=equip, task_type=task_type
    )
    if not template_path:
        flash("Template Excel introuvable. Veuillez le ré-importer dans Administration → Checklists.", "danger")
        return redirect(url_for('pma_dashboard'))

    now_str    = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_equip = equip.replace('/', '_').replace(' ', '_')
    orig_base  = os.path.splitext(os.path.basename(template_path))[0]
    out_name   = f"{orig_base}_{safe_equip}_{now_str}.xlsx"
    filled_dir = pma_config.dirs.get('filled', os.path.join(pma_config.active_base, 'PPE_Filled'))
    out_path   = os.path.join(filled_dir, out_name)

    meta = {
        'technician':    exec_record.get('technician_name', ''),
        'matricule':     exec_record.get('technician_matricule', ''),
        'shift':         exec_record.get('shift', ''),
        'equip':         equip,
        'type':          task_type,
        'sheet':         exec_record.get('sheet', ''),
        'week':          week,
        'version_id':    version_id,
        'template_path': template_path,
    }
    try:
        ok = save_filled_checklist(tasks, answers_raw, meta, out_path)
        if not ok or not os.path.isfile(out_path):
            flash("Impossible de régénérer le fichier Excel.", "danger")
            return redirect(url_for('pma_dashboard'))
    except Exception as e:
        logger.error(f"Checklist regeneration error for exec {exec_id}: {e}")
        flash("Erreur lors de la régénération du fichier Excel.", "danger")
        return redirect(url_for('pma_dashboard'))

    return _send_file(
        out_path,
        as_attachment=True,
        download_name=os.path.basename(out_path)
    )


# =============================================================
# PMA ANALYTICS
# =============================================================

@app.route('/pma/analytics')
@admin_required
def pma_analytics():
    month     = request.args.get('month', 'All')
    week      = request.args.get('week', 'All')
    sheet     = request.args.get('sheet', 'All')
    machine   = request.args.get('machine', 'All')
    status    = request.args.get('status', 'All')
    task_type = request.args.get('type', 'All')

    eng = get_pma_engine()
    data = eng.get_filtered_pma_analytics(
        month=month,
        week=week,
        sheet=sheet,
        machine=machine,
        status=status,
        task_type=task_type
    )
    pma_file = pma_config.get_last_excel_path()
    has_excel = bool(pma_file and os.path.exists(pma_file))

    return render_template('pma_analytics.html', data=data, has_excel=has_excel, filters={
        "month": month,
        "week": week,
        "sheet": sheet,
        "machine": machine,
        "status": status,
        "type": task_type
    })

@app.route('/api/analytics/pma')
@admin_required
def api_analytics_pma():
    month     = request.args.get('month', 'All')
    week      = request.args.get('week', 'All')
    sheet     = request.args.get('sheet', 'All')
    machine   = request.args.get('machine', 'All')
    status    = request.args.get('status', 'All')
    task_type = request.args.get('type', 'All')

    eng = get_pma_engine()
    data = eng.get_filtered_pma_analytics(
        month=month,
        week=week,
        sheet=sheet,
        machine=machine,
        status=status,
        task_type=task_type
    )
    return jsonify(data)

@app.route('/pma/analytics/export')
@admin_required
def pma_analytics_export():
    import io, csv
    month     = request.args.get('month', 'All')
    week      = request.args.get('week', 'All')
    sheet     = request.args.get('sheet', 'All')
    machine   = request.args.get('machine', 'All')
    status    = request.args.get('status', 'All')
    task_type = request.args.get('type', 'All')

    eng = get_pma_engine()
    data = eng.get_filtered_pma_analytics(
        month=month,
        week=week,
        sheet=sheet,
        machine=machine,
        status=status,
        task_type=task_type
    )

    output = io.StringIO()
    writer = csv.writer(output, delimiter=';')
    writer.writerow(['Equipement', 'Nom Machine', 'Zone / Ligne', 'Carte', 'Type / Frequence', 'Semaine', 'Mois', 'Statut'])

    for row in data.get('tasks', []):
        writer.writerow([
            row.get('equipment', ''),
            row.get('machine_name', ''),
            row.get('zone', ''),
            row.get('carte', ''),
            row.get('type', ''),
            row.get('week', ''),
            row.get('month', ''),
            row.get('status', '')
        ])

    output.seek(0)
    bytes_out = io.BytesIO(output.getvalue().encode('utf-8-sig'))
    return send_file(bytes_out, mimetype='text/csv', as_attachment=True, download_name=f'analyse_preventive_{datetime.date.today().isoformat()}.csv')

@app.route('/api/pma/weekly')
@admin_required
def api_pma_weekly():
    eng = get_pma_engine()
    df  = eng.current_df
    if df is None or df.empty:
        return jsonify({"tasks": [], "total": 0, "done": 0, "pending": 0})

    current_week = datetime.datetime.now().isocalendar()[1]
    week_labels  = [f"S{current_week}", f"KW {current_week}", f"KW{current_week}"]

    col_week   = 'Week'   if 'Week'   in df.columns else 'week'
    col_equip  = 'Equipment' if 'Equipment' in df.columns else 'equipment'
    col_type   = 'Type'   if 'Type'   in df.columns else 'type'
    col_sheet  = 'Sheet'  if 'Sheet'  in df.columns else 'sheet'
    col_status = 'Status' if 'Status' in df.columns else 'status'
    col_ri     = 'Raw_Index' if 'Raw_Index' in df.columns else 'raw_index'

    filtered = df[df[col_week].astype(str).isin(week_labels)]
    records  = filtered.to_dict('records')
    tasks    = []
    for r in records:
        status = str(r.get(col_status, '')).upper().strip()
        tasks.append({
            "equipment": str(r.get(col_equip, '')),
            "type":      str(r.get(col_type, '')),
            "week":      str(r.get(col_week, '')),
            "sheet":     str(r.get(col_sheet, '')),
            "status":    status,
            "raw_index": str(r.get(col_ri, ''))
        })

    done = sum(1 for t in tasks if t['status'] == 'COMPLÉTÉ')
    return jsonify({"tasks": tasks, "total": len(tasks), "done": done, "pending": len(tasks) - done})

# =============================================================
# ANOMALY REPORT (accessible to all authenticated users)
# =============================================================

@app.route('/pma/anomaly', methods=['GET', 'POST'])
@login_required
def anomaly_report():
    if request.method == 'POST':
        machine_id   = request.form.get('machine_id', '').strip()
        anomaly_type = request.form.get('anomaly_type', '').strip()
        description  = request.form.get('description', '').strip()

        if not machine_id or not description:
            flash("Veuillez remplir tous les champs obligatoires.", "danger")
            return redirect(url_for('anomaly_report'))

        try:
            data = {
                'machine_id':        machine_id,
                'machine_name':      machine_id,
                'group_name':        'ANOMALIE',
                'fault_type':        anomaly_type,
                'fault_description': f"[ANOMALIE] {anomaly_type}: {description}",
                'category':          'ANOMALY',
                'technician_name':   session.get('name', ''),
                'technician_mat':    session.get('matricule', ''),
                'start_time':        datetime.datetime.now().isoformat(),
                'end_time':          None,
                'downtime_minutes':  0,
                'status':            'OPEN',
                'code_asp':          '',
                'shift':             session.get('shift', ''),
                'remarks':           description
            }
            code = ima_db.create_intervention(data)
        except Exception as e:
            logger.error(f"Anomaly DB error: {e}")
            code = None

        # Email notification
        try:
            email_cfg  = pma_config.load_json(os.path.join(pma_config.active_base, "email_config.json"))
            recipients = [email_cfg.get('recipient')] if email_cfg.get('recipient') else []
            recipients += email_cfg.get('recipients', [])
            sender, pw = email_cfg.get('sender', ''), email_cfg.get('password', '')
            if sender and pw and recipients:
                subject = f"[ANOMALIE CRITIQUE] {machine_id} — {anomaly_type}"
                body    = (f"Anomalie critique signalée.\n\n"
                           f"Équipement : {machine_id}\nType       : {anomaly_type}\n"
                           f"Description: {description}\n"
                           f"Technicien : {session.get('name', '')} ({session.get('matricule', '')})\n"
                           f"Date       : {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")
                for recip in set(recipients):
                    send_simple_alert(sender, pw, recip, subject, body)
        except Exception as e:
            logger.warning(f"Anomaly email error: {e}")

        flash(f"Anomalie enregistrée{f' (code: {code})' if code else ''} et notification envoyée.", "success")
        return redirect(url_for('pma_dashboard'))

    all_machines = ima_db.get_all_machines()
    return render_template('anomaly_report.html', machines=all_machines)

# =============================================================
# MANAGER ACTION PLAN
# =============================================================

@app.route('/manager-plan')
@admin_required
def manager_plan():
    meetings       = load_meetings(MANAGER_PLAN_PATH)
    actions        = load_actions(MANAGER_PLAN_PATH)
    kpis           = get_action_kpis(actions)
    kanban = {
        'Backlog':  [a for a in actions if a.get('Statut') == 'Backlog'],
        'En cours': [a for a in actions if a.get('Statut') == 'En cours'],
        'Terminée': [a for a in actions if a.get('Statut') == 'Terminée'],
        'Bloquée':  [a for a in actions if a.get('Statut') == 'Bloquée'],
    }
    meeting_titles = [m.get('Titre', '') for m in meetings if m.get('Titre')]
    return render_template(
        'manager_plan.html',
        meetings=meetings, actions=actions,
        kpis=kpis, kanban=kanban,
        meeting_titles=meeting_titles
    )

@app.route('/manager-plan/meeting/add', methods=['POST'])
@admin_required
def manager_plan_add_meeting():
    meeting = {
        'Date':         request.form.get('date', datetime.date.today().isoformat()),
        'Titre':        request.form.get('titre', '').strip(),
        'Responsable':  request.form.get('responsable', '').strip(),
        'Participants': request.form.get('participants', '').strip(),
        'Statut':       request.form.get('statut', 'Planifiée'),
        'Notes':        request.form.get('notes', '').strip(),
    }
    if not meeting['Titre']:
        flash("Le titre de la réunion est requis.", "danger")
        return redirect(url_for('manager_plan'))
    if save_meeting(MANAGER_PLAN_PATH, meeting):
        flash("Réunion ajoutée avec succès.", "success")
    else:
        flash("Erreur lors de l'enregistrement de la réunion.", "danger")
    return redirect(url_for('manager_plan'))

@app.route('/manager-plan/action/add', methods=['POST'])
@admin_required
def manager_plan_add_action():
    action = {
        'Action':      request.form.get('action', '').strip(),
        'Responsable': request.form.get('responsable', '').strip(),
        'Échéance':    request.form.get('echeance', '').strip(),
        'Statut':      request.form.get('statut', 'Backlog'),
        'Notes':       request.form.get('notes', '').strip(),
        'Réunion_Ref': request.form.get('reunion_ref', '').strip(),
    }
    if not action['Action']:
        flash("La description de l'action est requise.", "danger")
        return redirect(url_for('manager_plan'))
    if save_action(MANAGER_PLAN_PATH, action):
        flash("Action ajoutée avec succès.", "success")
    else:
        flash("Erreur lors de l'enregistrement de l'action.", "danger")
    return redirect(url_for('manager_plan'))

@app.route('/manager-plan/action/<action_id>/update', methods=['POST'])
@admin_required
def manager_plan_update_action(action_id):
    new_status = request.form.get('statut', '')
    if new_status and update_action_status(MANAGER_PLAN_PATH, action_id, new_status):
        flash(f"Statut mis à jour → {new_status}.", "success")
    else:
        flash("Mise à jour échouée.", "danger")
    return redirect(url_for('manager_plan'))

@app.route('/manager-plan/action/<action_id>/delete', methods=['POST'])
@admin_required
def manager_plan_delete_action(action_id):
    if delete_action(MANAGER_PLAN_PATH, action_id):
        flash("Action supprimée.", "success")
    else:
        flash("Suppression échouée.", "danger")
    return redirect(url_for('manager_plan'))

@app.route('/api/manager-plan/kpis')
@admin_required
def api_manager_plan_kpis():
    actions = load_actions(MANAGER_PLAN_PATH)
    return jsonify(get_action_kpis(actions))

# =============================================================
# DOCUMENT MANAGEMENT
# =============================================================

@app.route('/documents')
@login_required
def documents_page():
    """Public-facing document shortcuts page — DB-driven, no hardcoded paths."""
    docs = doc_mgr.get_all_documents(active_only=True)
    return render_template('documents.html', documents=docs)


@app.route('/documents/view/<int:doc_id>')
@login_required
def document_view(doc_id):
    """
    Protected document viewer / inline file server.
    Steps:
      1. Verify session
      2. Fetch metadata from DB
      3. Verify file exists inside safe storage boundary
      4. Serve inline (for PDF/images) or render viewer page
    """
    doc = doc_mgr.get_document(doc_id)
    if not doc or not doc.get('is_active', 0):
        flash("Document introuvable ou désactivé.", "danger")
        return redirect(url_for('documents_page'))

    # ?inline=1 → raw file for iframe/img src (PDF, images)
    if request.args.get('inline') == '1':
        abs_path = doc_mgr.get_abs_path(doc_id)
        if not abs_path:
            flash("Fichier non trouvé sur le serveur.", "danger")
            return redirect(url_for('documents_page'))
        mime_map = {
            'pdf': 'application/pdf',
            'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'xls': 'application/vnd.ms-excel',
            'xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.12',
            'csv': 'text/csv; charset=utf-8',
            'txt': 'text/plain; charset=utf-8',
            'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'doc': 'application/msword',
            'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            'ppt': 'application/vnd.ms-powerpoint',
            'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'gif': 'image/gif',
        }
        mimetype = mime_map.get(doc['file_type'].lower(), 'application/octet-stream')
        return send_file(abs_path, mimetype=mimetype)

    # Full viewer page
    return render_template('document_viewer.html', doc=doc)


@app.route('/documents/download/<int:doc_id>')
@login_required
def document_download(doc_id):
    """Protected force-download route."""
    doc = doc_mgr.get_document(doc_id)
    if not doc:
        flash("Document introuvable.", "danger")
        return redirect(url_for('documents_page'))

    abs_path = doc_mgr.get_abs_path(doc_id)
    if not abs_path:
        flash("Fichier non disponible sur le serveur.", "danger")
        return redirect(url_for('documents_page'))

    logger.info(f"[DOC] Download doc_id={doc_id} user={session.get('user')} file={doc['file_name']}")
    return send_file(abs_path, as_attachment=True, download_name=doc['file_name'])


# ── Administration ──────────────────────────────────────────────────────────

@app.route('/admin/documents')
@admin_required
def admin_documents():
    """Admin document management page (all docs, including inactive)."""
    docs = doc_mgr.get_all_documents(active_only=False)
    return render_template('admin_documents.html', docs=docs)


@app.route('/admin/documents/add', methods=['POST'])
@admin_required
def admin_documents_add():
    """Upload + register a new document."""
    display_name  = request.form.get('display_name', '').strip()
    display_order = int(request.form.get('display_order', 0) or 0)
    is_active     = int(request.form.get('is_active', 1) or 1)
    file_obj      = request.files.get('file')

    if not display_name:
        flash("Le nom du bouton est obligatoire.", "danger")
        return redirect(url_for('admin_documents'))
    if not file_obj or file_obj.filename == '':
        flash("Veuillez sélectionner un fichier.", "danger")
        return redirect(url_for('admin_documents'))

    # 1. Insert DB row to get doc_id
    ok, msg, doc_id = doc_mgr.add_document(
        display_name  = display_name,
        file_name     = DocumentManager.safe_filename(file_obj.filename),
        file_type     = DocumentManager.detect_file_type(file_obj.filename),
        storage_path  = '',   # will update after file save
        display_order = display_order
    )
    if not ok:
        flash(msg, "danger")
        return redirect(url_for('admin_documents'))

    # 2. Save file to disk
    file_ok, file_msg, rel_path, file_type = doc_mgr.save_uploaded_file(file_obj, doc_id)
    if not file_ok:
        # Remove the orphan DB row
        doc_mgr.soft_delete(doc_id)
        flash(f"Erreur fichier: {file_msg}", "danger")
        return redirect(url_for('admin_documents'))

    # 3. Update DB row with real path + is_active
    from core.document_manager import DocumentManager as _DM
    import sqlite3 as _sq3
    with _sq3.connect(IMA_DB_PATH) as _conn:
        _conn.execute(
            "UPDATE documents SET storage_path=?, file_name=?, file_type=?, is_active=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
            (rel_path, DocumentManager.safe_filename(request.files['file'].filename), file_type, is_active, doc_id)
        )
        _conn.commit()

    logger.info(f"[DOC] Added doc_id={doc_id} name='{display_name}' path={rel_path} user={session.get('user')}")
    flash(f"Document '{display_name}' ajouté avec succès.", "success")
    return redirect(url_for('admin_documents'))


@app.route('/admin/documents/<int:doc_id>/replace', methods=['POST'])
@admin_required
def admin_documents_replace(doc_id):
    """Replace the physical file for a document — button label unchanged."""
    file_obj = request.files.get('file')
    if not file_obj or file_obj.filename == '':
        flash("Veuillez sélectionner un nouveau fichier.", "danger")
        return redirect(url_for('admin_documents'))

    ok, msg = doc_mgr.replace_file(doc_id, file_obj)
    flash(msg, "success" if ok else "danger")
    return redirect(url_for('admin_documents'))


@app.route('/admin/documents/<int:doc_id>/rename', methods=['POST'])
@admin_required
def admin_documents_rename(doc_id):
    """Update only the display name (button label)."""
    new_name = request.form.get('display_name', '').strip()
    ok, msg = doc_mgr.update_display_name(doc_id, new_name)
    flash(msg, "success" if ok else "danger")
    return redirect(url_for('admin_documents'))


@app.route('/admin/documents/<int:doc_id>/toggle', methods=['POST'])
@admin_required
def admin_documents_toggle(doc_id):
    """Toggle is_active (show/hide from Documents page)."""
    ok, msg, _ = doc_mgr.toggle_active(doc_id)
    flash(msg, "success" if ok else "danger")
    return redirect(url_for('admin_documents'))


@app.route('/admin/documents/<int:doc_id>/reorder', methods=['POST'])
@admin_required
def admin_documents_reorder(doc_id):
    """Update display_order for a document."""
    try:
        new_order = int(request.form.get('display_order', 0))
    except ValueError:
        new_order = 0
    doc_mgr.update_order(doc_id, new_order)
    return redirect(url_for('admin_documents'))


@app.route('/admin/documents/<int:doc_id>/delete', methods=['POST'])
@admin_required
def admin_documents_delete(doc_id):
    """
    Soft delete: sets is_active=0, archives physical file as <file>.archived.
    Does NOT permanently remove anything — safe for historical references.
    """
    ok, msg = doc_mgr.soft_delete(doc_id)
    logger.info(f"[DOC] Soft-delete doc_id={doc_id} user={session.get('user')} ok={ok}")
    flash(msg, "success" if ok else "danger")
    return redirect(url_for('admin_documents'))


# =============================================================
# ERROR HANDLERS
# =============================================================

@app.errorhandler(404)
def not_found(e):
    return render_template('403.html'), 404

@app.errorhandler(500)
def server_error(e):
    logger.error(f"500 error: {e}")
    flash("Une erreur interne est survenue. Veuillez réessayer.", "danger")
    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True, port=5000)
