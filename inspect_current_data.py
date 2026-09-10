import os
import sqlite3
import json

curr_db = r"C:\Users\hama0\Documents\SEBN-TN IMA APP\IMA.db"
conn = sqlite3.connect(curr_db)
cur = conn.cursor()

print("--- USERS TABLE ---")
cur.execute("SELECT * FROM users")
for r in cur.fetchall():
    print(r)

print("\n--- PASSATIONS TABLE ---")
cur.execute("SELECT * FROM passations")
for r in cur.fetchall():
    print(r)

print("\n--- EBM SETTINGS ---")
cur.execute("SELECT * FROM ebm_settings")
for r in cur.fetchall():
    print(r)

print("\n--- EBM ACTION PLANS ---")
cur.execute("SELECT * FROM ebm_action_plans")
for r in cur.fetchall():
    print(r)

conn.close()
