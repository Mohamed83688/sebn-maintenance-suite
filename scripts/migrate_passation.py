import os, sys, sqlite3

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)

from ima.config import IMAConfig
from ima.database import IMADatabase

ima_config = IMAConfig()
IMA_DB_PATH = os.path.join(ima_config.active_base, 'IMA.db')

ima_db = IMADatabase(IMA_DB_PATH)

SRC_CANDIDATES = [
    os.path.join('C:', os.sep, 'Users', 'hama0', '.gemini', 'antigravity-ide', 'scratch', 'sebn-tn03-passation', 'database', 'sebn_passation.db'),
    os.path.join(os.path.dirname(PROJECT_ROOT), 'sebn-tn03-passation', 'database', 'sebn_passation.db'),
]
SRC_DB = next((p for p in SRC_CANDIDATES if os.path.exists(p)), None)

def migrate():
    if not SRC_DB:
        print('[MIGRATION] Source DB not found - skipping')
        return True

    print(f'[MIGRATION] Source:      {SRC_DB}')
    print(f'[MIGRATION] Destination: {IMA_DB_PATH}')

    src = sqlite3.connect(SRC_DB)
    src.row_factory = sqlite3.Row
    dst = sqlite3.connect(IMA_DB_PATH)
    dst.row_factory = sqlite3.Row
    dst.execute('PRAGMA foreign_keys=OFF')

    src_users = src.execute('SELECT * FROM users').fetchall()
    user_id_map = {}
    migrated_users = 0
    for u in src_users:
        if str(u['role']).lower() == 'admin':
            continue
        dst.execute('INSERT OR IGNORE INTO passation_users (matricule, name, role) VALUES (?,?,?)',
                    (str(u['matricule']).upper().strip(), u['name'] or '', 'technician'))
        row = dst.execute('SELECT id FROM passation_users WHERE UPPER(matricule)=UPPER(?)',
                          (str(u['matricule']).strip(),)).fetchone()
        if row:
            user_id_map[u['id']] = row['id']
            migrated_users += 1
    dst.commit()
    print(f'[MIGRATION] Users:     {migrated_users}')

    src_qs = src.execute('SELECT * FROM questions').fetchall()
    q_id_map = {}
    migrated_qs = 0
    for q in src_qs:
        text = q['text']
        cat = q['category'] if 'category' in q.keys() and q['category'] else 'General'
        act = q['active'] if 'active' in q.keys() else 1

        ex = dst.execute('SELECT id FROM passation_questions WHERE text=?', (text,)).fetchone()
        if ex:
            q_id_map[q['id']] = ex['id']
            continue
        cur = dst.execute(
            'INSERT INTO passation_questions (text, category, type, options, active, display_order) VALUES (?,?,?,?,?,?)',
            (text, cat, 'text', '', act, q['id']))
        q_id_map[q['id']] = cur.lastrowid
        migrated_qs += 1
    dst.commit()
    print(f'[MIGRATION] Questions: {migrated_qs}')

    src_ps = src.execute('SELECT * FROM passations').fetchall()
    p_id_map = {}
    migrated_ps = 0
    for p in src_ps:
        old_uid = p['user_id']
        tech_name = tech_mat = ''
        if old_uid and old_uid in user_id_map:
            ur = dst.execute('SELECT name, matricule FROM passation_users WHERE id=?',
                             (user_id_map[old_uid],)).fetchone()
            if ur:
                tech_name, tech_mat = ur['name'], ur['matricule']
        shift = p['shift'] or 'SHIFT 1'
        su = shift.upper()
        if 'POSTE 1' in su or 'SHIFT 1' in su: shift = 'SHIFT 1 (06:00 - 14:00)'
        elif 'POSTE 2' in su or 'SHIFT 2' in su: shift = 'SHIFT 2 (14:00 - 22:00)'
        elif 'POSTE 3' in su or 'SHIFT 3' in su: shift = 'SHIFT 3 (22:00 - 06:00)'
        cur = dst.execute(
            'INSERT INTO passations (technician_matricule, technician_name, shift, remarks, timestamp) VALUES (?,?,?,?,?)',
            (tech_mat, tech_name, shift, '', p['timestamp'] or ''))
        p_id_map[p['id']] = cur.lastrowid
        migrated_ps += 1
    dst.commit()
    print(f'[MIGRATION] Passations: {migrated_ps}')

    src_rs = src.execute('SELECT * FROM responses').fetchall()
    migrated_rs = skipped_rs = 0
    for r in src_rs:
        np = p_id_map.get(r['passation_id'])
        nq = q_id_map.get(r['question_id'])
        if not np or not nq:
            skipped_rs += 1
            continue
        dst.execute('INSERT INTO passation_responses (passation_id, question_id, answer) VALUES (?,?,?)',
                    (np, nq, r['answer'] or ''))
        migrated_rs += 1
    dst.commit()
    print(f'[MIGRATION] Responses:  {migrated_rs} ok, {skipped_rs} skipped')

    try:
        src_sets = src.execute('SELECT key, value FROM settings').fetchall()
        for s in src_sets:
            dst.execute('INSERT OR REPLACE INTO passation_settings (key, value) VALUES (?,?)',
                        (s['key'], s['value'] or ''))
        dst.commit()
        print(f'[MIGRATION] Settings:   {len(src_sets)}')
    except Exception as e:
        print(f'[MIGRATION] Settings:   skipped ({e})')

    src.close()
    dst.execute('PRAGMA foreign_keys=ON')
    dst.close()
    print('[MIGRATION] SUCCESS: Historical passation data successfully migrated into IMA.db')
    return True

if __name__ == '__main__':
    migrate()
