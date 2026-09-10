import os
import sqlite3

print("=" * 70)
print("INSPECTING PASSATION DATABASE & LOGIC")
print("=" * 70)

db_paths = [
    r"C:\Users\hama0\.gemini\antigravity\scratch\sebn-tn03-passation\database\sebn_v3.db",
    r"C:\Users\hama0\.gemini\antigravity\scratch\sebn-tn03-passation\database\sebn_passation.db",
    r"C:\Users\hama0\.gemini\antigravity\scratch\sebn-tn03-passation\database\sebn_v2.db",
]

for p in db_paths:
    if os.path.exists(p):
        print(f"\n--- Checking Passation DB: {p} ---")
        conn = sqlite3.connect(p)
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [r[0] for r in cur.fetchall()]
        print("Tables:", tables)
        for t in tables:
            cur.execute(f"PRAGMA table_info({t})")
            cols = [f"{c[1]} ({c[2]})" for c in cur.fetchall()]
            cur.execute(f"SELECT COUNT(*) FROM {t}")
            cnt = cur.fetchone()[0]
            print(f"  [{t}] {cnt} rows | Cols: {cols}")
        conn.close()

print("\n" + "=" * 70)
print("INSPECTING CURRENT SEBN-TN USER / AUTHENTICATION STRUCTURE")
print("=" * 70)

curr_db = r"C:\Users\hama0\Documents\SEBN-TN IMA APP\IMA.db"
if not os.path.exists(curr_db):
    curr_db = r"C:\Users\hama0\.gemini\antigravity\scratch\excel_monthly_viewer\project\data\IMA.db"

if os.path.exists(curr_db):
    print(f"Current IMA DB: {curr_db}")
    conn = sqlite3.connect(curr_db)
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = [r[0] for r in cur.fetchall()]
    print("Tables in IMA DB:", tables)
    for t in tables:
        if 'user' in t.lower() or 'tech' in t.lower() or 'auth' in t.lower():
            cur.execute(f"PRAGMA table_info({t})")
            cols = [f"{c[1]} ({c[2]})" for c in cur.fetchall()]
            cur.execute(f"SELECT COUNT(*) FROM {t}")
            cnt = cur.fetchone()[0]
            print(f"  [{t}] {cnt} rows | Cols: {cols}")
    conn.close()
