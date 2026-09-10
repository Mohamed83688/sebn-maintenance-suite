import sys
import os
sys.path.insert(0, r'c:\Users\hama0\.gemini\antigravity\scratch\excel_monthly_viewer\project')

from web_portal.app import app

def run_tests():
    with app.test_client() as client:
        print("======================================================================")
        print("TESTING PASSATION RBAC: TECHNICIAN FILLING & ADMIN QUESTION EDITING")
        print("======================================================================")

        # 1. TECHNICIAN WORKFLOW
        # Login as technician
        r = client.post('/login', data={'username': 'ahmed.bensalah', 'password': 'Tech@2026!'}, follow_redirects=True)
        assert r.status_code == 200, f"Tech login failed: {r.status_code}"
        print("[TEST 1] Technician login: SUCCESS")

        # Tech can view passations dashboard
        r = client.get('/passation')
        assert r.status_code == 200, f"Tech /passation failed: {r.status_code}"
        print("  [PASS] Technician access to /passation: 200 OK")

        # Tech can view new passation form
        r = client.get('/passation/new')
        assert r.status_code == 200, f"Tech /passation/new failed: {r.status_code}"
        print("  [PASS] Technician access to /passation/new: 200 OK")

        # Tech can fill and submit a passation
        r = client.post('/passation/save', data={
            'shift': 'SHIFT 1 (06:00 - 14:00)',
            'target_shift': 'SHIFT 2 (14:00 - 22:00)',
            'zone_name': 'Ligne Test Tech',
            'technician_name': 'Ahmed Ben Salah',
            'technician_matricule': 'AHMED-01',
            'remarks': 'Passation remplie par le technicien sans anomalie.',
            'question_1': 'OK',
            'question_2': 'OK'
        }, follow_redirects=True)
        assert r.status_code == 200, f"Tech /passation/save failed: {r.status_code}"
        print("  [PASS] Technician submitted passation: SUCCESS")

        # Tech BLOCKED from managing questions (403 Forbidden)
        r = client.get('/passation/questions')
        assert r.status_code == 403, f"Tech was NOT blocked on /passation/questions: {r.status_code}"
        r = client.post('/passation/questions/add', data={'text': 'Test Hack'})
        assert r.status_code == 403, f"Tech was NOT blocked on /passation/questions/add: {r.status_code}"
        print("  [PASS] Technician strictly BLOCKED from /passation/questions (HTTP 403): OK")

        # Logout tech
        client.get('/logout')

        # 2. ADMIN WORKFLOW
        # Login as Admin
        r = client.post('/login', data={'username': 'admin', 'password': 'admin2026'}, follow_redirects=True)
        assert r.status_code == 200, f"Admin login failed: {r.status_code}"
        print("\n[TEST 2] Admin login: SUCCESS")

        # Admin can view questions list
        r = client.get('/passation/questions')
        assert r.status_code == 200, f"Admin /passation/questions failed: {r.status_code}"
        print("  [PASS] Admin access to /passation/questions: 200 OK")

        # Admin can add a question
        r = client.post('/passation/questions/add', data={
            'text': 'Test Point Nouveau',
            'category': 'Qualité',
            'type': 'CHOICE',
            'options': 'OK,NOK,N/A'
        }, follow_redirects=True)
        assert r.status_code == 200, f"Admin add question failed: {r.status_code}"
        print("  [PASS] Admin added new passation question: OK")

        # Get question ID from manager
        from core.passation_manager import PassationManager
        from ima.config import IMAConfig
        ima_cfg = IMAConfig()
        db_path = os.path.join(ima_cfg.active_base, "IMA.db")
        pm = PassationManager(db_path=db_path, data_dir=ima_cfg.active_base)
        qs = pm.get_questions(active_only=False)
        added_q = next((q for q in qs if q['text'] == 'Test Point Nouveau'), None)
        assert added_q is not None
        qid = added_q['id']

        # Admin can edit the question
        r = client.post(f'/passation/questions/edit/{qid}', data={
            'text': 'Test Point Modifié par Admin',
            'category': 'Qualité & Sécurité',
            'type': 'CHOICE',
            'options': 'Conforme,Non Conforme',
            'sort_order': 10
        }, follow_redirects=True)
        assert r.status_code == 200, f"Admin edit question failed: {r.status_code}"
        edited_q = pm.get_question(qid)
        assert edited_q['text'] == 'Test Point Modifié par Admin'
        print("  [PASS] Admin edited passation question: OK")

        # Admin can toggle status
        r = client.post(f'/passation/questions/toggle/{qid}', follow_redirects=True)
        assert r.status_code == 200
        toggled_q = pm.get_question(qid)
        assert toggled_q['is_active'] == 0
        print("  [PASS] Admin toggled passation question status: OK")

        # Admin can delete question
        r = client.post(f'/passation/questions/delete/{qid}', follow_redirects=True)
        assert r.status_code == 200
        assert pm.get_question(qid) is None
        print("  [PASS] Admin deleted passation question: OK")

        # Admin is still blocked from Owner modules (EBM, Users)
        r = client.get('/admin/users')
        assert r.status_code == 403
        r = client.get('/ebm')
        assert r.status_code == 403
        print("  [PASS] Admin still strictly blocked from Owner modules (EBM, Users) (HTTP 403): OK")

        # Logout
        client.get('/logout')

        print("\n======================================================================")
        print("ALL PASSATION WORKFLOW & RBAC TESTS PASSED SUCCESSFULLY!")
        print("======================================================================")

if __name__ == '__main__':
    run_tests()
