import sys
import os
import sqlite3
import openpyxl

sys.path.insert(0, '.')
from web_portal.app import app, user_mgr, ebm_mgr, passation_mgr

print("=" * 70)
print("TESTING MAJOR INTEGRATION: AUTH, EBM, PASSATION, USER MANAGEMENT")
print("=" * 70)

client = app.test_client()

# ─────────────────────────────────────────────────────────────────────────────
# 1. TEST OWNER LOGIN & FULL ACCESS
# ─────────────────────────────────────────────────────────────────────────────
print("\n[TEST 1] OWNER LOGIN & FULL ACCESS")
res = client.post('/login', data={'username': 'owner', 'password': 'Owner@SEBN2026!'}, follow_redirects=True)
assert res.status_code == 200, f"Owner login failed: {res.status_code}"
print("  [PASS] Owner login SUCCESS")

# Check all standard pages
for endpoint in ['/dashboard', '/interventions', '/machines', '/pma', '/analytics', '/admin', '/vault']:
    r = client.get(endpoint)
    assert r.status_code == 200, f"Owner access failed on {endpoint}: {r.status_code}"
print("  [PASS] Owner access to all core pages (Dashboard, PMA, IMA, Analytics, Admin, Vault): 200 OK")

# Check Owner-Only Modules
for endpoint in ['/admin/users', '/ebm', '/ebm/analytics', '/ebm/validation', '/ebm/plan-action', '/ebm/settings', '/passation', '/passation/new', '/passation/questions']:
    r = client.get(endpoint)
    assert r.status_code == 200, f"Owner access failed on {endpoint}: {r.status_code}"
print("  [PASS] Owner access to Owner-Only modules (/admin/users, /ebm/*, /passation/*): 200 OK")

# ─────────────────────────────────────────────────────────────────────────────
# 2. TEST OWNER USER & TECHNICIAN MANAGEMENT
# ─────────────────────────────────────────────────────────────────────────────
print("\n[TEST 2] OWNER USER & TECHNICIAN MANAGEMENT")

# Create a new technician
res_create = client.post('/admin/users/create', data={
    'name': 'Ahmed Ben Salah',
    'username': 'ahmed.bensalah',
    'password': 'TechPassword2026!',
    'matricule': 'TN-5555',
    'role': 'TECHNICIAN',
    'shift': 'B'
}, follow_redirects=True)
assert res_create.status_code == 200

created_tech = user_mgr.get_user_by_username('ahmed.bensalah')
assert created_tech is not None, "Technician was not created in database!"
assert created_tech['role'] == 'TECHNICIAN'
assert created_tech['matricule'] == 'TN-5555'
print(f"  [PASS] Owner created technician: {created_tech['name']} (Username: {created_tech['username']})")

# Create a standard user
res_create_user = client.post('/admin/users/create', data={
    'name': 'Sami Operator',
    'username': 'sami.operator',
    'password': 'UserPassword2026!',
    'matricule': '',
    'role': 'USER',
    'shift': 'A'
}, follow_redirects=True)
assert res_create_user.status_code == 200
created_user = user_mgr.get_user_by_username('sami.operator')
assert created_user is not None
print(f"  [PASS] Owner created standard user: {created_user['name']} (Role: {created_user['role']})")

# Reset password
res_reset = client.post(f'/admin/users/reset-password/{created_tech["id"]}', data={
    'new_password': 'NewTechPassword2026!'
}, follow_redirects=True)
assert res_reset.status_code == 200
print("  [PASS] Owner successfully reset technician password")

client.get('/logout')

# ─────────────────────────────────────────────────────────────────────────────
# 3. TEST NEW TECHNICIAN LOGIN & RBAC RESTRICTIONS
# ─────────────────────────────────────────────────────────────────────────────
print("\n[TEST 3] NEW TECHNICIAN LOGIN & RBAC ENFORCEMENT")

# Test login with old password (must fail)
res_old = client.post('/login', data={'username': 'ahmed.bensalah', 'password': 'TechPassword2026!'}, follow_redirects=True)
assert "Identifiants incorrects" in res_old.data.decode('utf-8')
print("  [PASS] Old password rejected after reset: OK")

# Test login with new password (must succeed)
res_new = client.post('/login', data={'username': 'ahmed.bensalah', 'password': 'NewTechPassword2026!'}, follow_redirects=True)
assert res_new.status_code == 200
print("  [PASS] Technician login with Username + Password: SUCCESS")

# Operational pages should be accessible
for endpoint in ['/pma', '/interventions', '/machines', '/pma/anomaly']:
    r = client.get(endpoint)
    assert r.status_code == 200, f"Technician access failed on {endpoint}: {r.status_code}"
print("  [PASS] Technician access to operational pages (PMA, Interventions, Machines, Anomaly): 200 OK")

# Protected Admin & Owner pages MUST BE BLOCKED (403 Forbidden)
for endpoint in ['/admin', '/admin/users', '/ebm', '/ebm/analytics', '/passation']:
    r = client.get(endpoint)
    assert r.status_code == 403, f"SECURITY BREACH: Technician was allowed into {endpoint} (got {r.status_code}, expected 403)"
print("  [PASS] Technician strictly BLOCKED from Admin, Users, EBM, Passation (HTTP 403 Forbidden): OK")

client.get('/logout')

# ─────────────────────────────────────────────────────────────────────────────
# 4. TEST ADMIN LOGIN & PERMISSIONS
# ─────────────────────────────────────────────────────────────────────────────
print("\n[TEST 4] ADMIN LOGIN & PERMISSIONS")
res_admin = client.post('/login', data={'username': 'admin', 'password': 'admin2026'}, follow_redirects=True)
assert res_admin.status_code == 200
print("  [PASS] Admin login: SUCCESS")

# Admin can access standard administration & analytics
for endpoint in ['/dashboard', '/admin', '/analytics', '/pma/analytics', '/vault', '/training']:
    r = client.get(endpoint)
    assert r.status_code == 200, f"Admin access failed on {endpoint}: {r.status_code}"
print("  [PASS] Admin access to Admin, Analytics, Vault, Training: 200 OK")

# Admin MUST BE BLOCKED from Owner-Only modules: EBM, Passation, User Management
for endpoint in ['/admin/users', '/ebm', '/passation']:
    r = client.get(endpoint)
    assert r.status_code == 403, f"SECURITY BREACH: Admin was allowed into {endpoint} (got {r.status_code}, expected 403)"
print("  [PASS] Admin strictly BLOCKED from Owner modules (Users, EBM, Passation) (HTTP 403 Forbidden): OK")

client.get('/logout')

# ─────────────────────────────────────────────────────────────────────────────
# 5. TEST EBM & PASSATION FUNCTIONALITY AS OWNER
# ─────────────────────────────────────────────────────────────────────────────
print("\n[TEST 5] EBM & PASSATION WORKFLOWS AS OWNER")
client.post('/login', data={'username': 'owner', 'password': 'Owner@SEBN2026!'}, follow_redirects=True)

# EBM Action Plan CRUD
res_ebm_add = client.post('/ebm/plan-action/add', data={
    'title': 'Audit Budget Q3 EBM',
    'description': 'Vérification des PO avec le service achat',
    'responsible': 'Mohamed Issaoui',
    'due_date': '2026-09-30',
    'priority': 'High'
}, follow_redirects=True)
assert res_ebm_add.status_code == 200

plans = ebm_mgr.get_action_plans()
assert len(plans) > 0
plan_id = plans[0]['id']
print(f"  [PASS] EBM Action Plan created (ID #{plan_id})")

res_toggle = client.post(f'/ebm/plan-action/toggle/{plan_id}', follow_redirects=True)
assert res_toggle.status_code == 200
print("  [PASS] EBM Action Plan status toggle: OK")

# Passation Handover Workflow
pass_questions = passation_mgr.get_questions(active_only=True)
assert len(pass_questions) > 0
post_passation = {
    'shift': 'SHIFT 1 (06:00 - 14:00)',
    'target_shift': 'SHIFT 2 (14:00 - 22:00)',
    'zone_name': 'KSM MMA',
    'technician_name': 'Mohamed Issaoui',
    'technician_matricule': 'OWNER-01',
    'remarks': 'Machine KSM-02 contrôlée, rasoir pushback vérifié.',
}
for q in pass_questions:
    post_passation[f'question_{q["id"]}'] = 'Conforme' if q['type'] == 'CHOICE' else 'Vérification effectuée'

res_pass_save = client.post('/passation/save', data=post_passation, follow_redirects=True)
assert res_pass_save.status_code == 200
passations_list = passation_mgr.get_passations()
assert len(passations_list) > 0
latest_pass = passations_list[0]
print(f"  [PASS] Passation shift handover saved successfully (ID #{latest_pass['id']})")

# Passation Detail View
res_pass_detail = client.get(f'/passation/{latest_pass["id"]}')
assert res_pass_detail.status_code == 200
print(f"  [PASS] Passation detail view (/passation/{latest_pass['id']}): 200 OK")

# Passation Excel Export
res_pass_export = client.get(f'/passation/export/{latest_pass["id"]}')
assert res_pass_export.status_code == 200
assert len(res_pass_export.data) > 0
print(f"  [PASS] Passation Excel report download: 200 OK ({len(res_pass_export.data)} bytes)")

print("\n" + "=" * 70)
print("ALL INTEGRATION & AUTHENTICATION TESTS PASSED WITH 100% SUCCESS!")
print("=" * 70)
