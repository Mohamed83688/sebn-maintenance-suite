"""
SEBN-TN Maintenance System — SQLite to MySQL Data Migration & Validation Script
Migrates all 26 tables and 1100+ records from IMA.db to MySQL/MariaDB.
Preserves all primary keys, foreign keys, timestamps, and JSON data.
Can also generate a standalone migration.sql dump file.
"""
import os
import sys
import shutil
import sqlite3
import argparse
import datetime
import logging
from typing import Dict, List, Any, Tuple

# Add project root to sys.path
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _PROJECT_ROOT)

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

from core.db_mysql import get_mysql_config, get_mysql_connection, test_mysql_connection

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
logger = logging.getLogger("migration")

DEFAULT_SQLITE_PATH = os.path.join(_PROJECT_ROOT, "data", "IMA.db")
DEFAULT_SCHEMA_PATH = os.path.join(_PROJECT_ROOT, "scripts", "mysql_schema.sql")

# Table migration order (independent tables first, then dependent tables)
MIGRATION_ORDER = [
    "asp_codes",
    "settings",
    "reports",
    "machines",
    "machine_plans",
    "interventions",
    "users",
    "level_change_log",
    "exams",
    "exam_questions",
    "exam_answers",
    "exam_attempts",
    "exam_attempt_answers",
    "exam_question_images",
    "passation_settings",
    "passations",
    "passation_questions",
    "passation_responses",
    "ebm_settings",
    "ebm_action_plans",
    "ebm_validations",
    "documents",
    "checklist_definitions",
    "checklist_versions",
    "checklist_items",
    "checklist_executions",
]


def create_sqlite_backup(sqlite_path: str) -> str:
    """Creates a timestamped backup copy of the SQLite database."""
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{sqlite_path}.backup_{ts}"
    shutil.copy2(sqlite_path, backup_path)
    logger.info(f"Created SQLite backup at: {backup_path}")
    return backup_path


def format_sql_value(val: Any) -> str:
    """Formats a Python value as a MySQL SQL literal."""
    if val is None:
        return "NULL"
    if isinstance(val, (int, float)):
        return str(val)
    if isinstance(val, (datetime.datetime, datetime.date)):
        return f"'{val.strftime('%Y-%m-%d %H:%M:%S')}'"
    # String escaping
    s = str(val)
    s = s.replace("\\", "\\\\").replace("'", "''").replace("\0", "")
    return f"'{s}'"


def export_to_sql_file(sqlite_path: str, schema_path: str, output_sql_path: str):
    """Generates a complete standalone .sql script containing schema and all data."""
    logger.info(f"Generating full SQL dump from {sqlite_path} -> {output_sql_path}")
    conn = sqlite3.connect(sqlite_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    with open(schema_path, "r", encoding="utf-8") as sf:
        schema_sql = sf.read()

    with open(output_sql_path, "w", encoding="utf-8") as out:
        out.write("-- ==============================================================\n")
        out.write(f"-- SEBN-TN Complete Migration SQL Dump\n")
        out.write(f"-- Generated: {datetime.datetime.now().isoformat()}\n")
        out.write(f"-- Source: {sqlite_path}\n")
        out.write("-- ==============================================================\n\n")
        out.write("SET FOREIGN_KEY_CHECKS = 0;\n")
        out.write("SET NAMES utf8mb4;\n\n")

        # 1. Write schema
        out.write("-- --- SCHEMA DEFINITIONS ---\n")
        out.write(schema_sql)
        out.write("\n\n-- --- DATA INSERTS ---\n\n")

        # 2. Write data for each table
        total_rows = 0
        for table in MIGRATION_ORDER:
            try:
                cur.execute(f"SELECT * FROM [{table}]")
                rows = cur.fetchall()
            except sqlite3.OperationalError:
                continue

            if not rows:
                out.write(f"-- Table `{table}`: 0 rows\n\n")
                continue

            col_names = [d[0] for d in cur.description]
            cols_escaped = ", ".join([f"`{c}`" for c in col_names])
            out.write(f"-- Table `{table}`: {len(rows)} rows\n")
            out.write(f"TRUNCATE TABLE `{table}`;\n")

            # Batch insert in chunks of 50
            chunk_size = 50
            for i in range(0, len(rows), chunk_size):
                chunk = rows[i:i + chunk_size]
                out.write(f"INSERT INTO `{table}` ({cols_escaped}) VALUES\n")
                val_lines = []
                for r in chunk:
                    vals = ", ".join([format_sql_value(r[c]) for c in col_names])
                    val_lines.append(f"  ({vals})")
                out.write(",\n".join(val_lines) + ";\n")

            total_rows += len(rows)
            out.write("\n")

        out.write("SET FOREIGN_KEY_CHECKS = 1;\n")
        out.write(f"-- Total migrated records: {total_rows}\n")

    conn.close()
    logger.info(f"Dump complete! Total {total_rows} records written to {output_sql_path}")


def run_live_migration(sqlite_path: str, schema_path: str, mysql_config: Dict[str, Any]) -> Tuple[bool, Dict[str, Any]]:
    """Connects directly to MySQL, applies schema, and copies all table data."""
    logger.info("Connecting to MySQL...")
    conn = sqlite3.connect(sqlite_path)
    conn.row_factory = sqlite3.Row
    sqlite_cur = conn.cursor()

    mysql_conn = get_mysql_connection(mysql_config)
    
    report = {
        "tables": {},
        "total_sqlite": 0,
        "total_mysql": 0,
        "key_records": {}
    }

    try:
        # 1. Apply schema
        logger.info("Applying MySQL schema...")
        with open(schema_path, "r", encoding="utf-8") as sf:
            schema_content = sf.read()

        statements = [s.strip() for s in schema_content.split(";") if s.strip()]
        with mysql_conn.cursor() as my_cur:
            my_cur.execute("SET FOREIGN_KEY_CHECKS = 0")
            for stmt in statements:
                try:
                    my_cur.execute(stmt)
                except Exception as e:
                    logger.warning(f"Schema statement warning: {e}")
            mysql_conn.commit()

        # 2. Migrate data table by table
        logger.info("Migrating table data...")
        with mysql_conn.cursor() as my_cur:
            my_cur.execute("SET FOREIGN_KEY_CHECKS = 0")
            
            for table in MIGRATION_ORDER:
                try:
                    sqlite_cur.execute(f"SELECT * FROM [{table}]")
                    rows = sqlite_cur.fetchall()
                except sqlite3.OperationalError:
                    continue

                row_count = len(rows)
                report["total_sqlite"] += row_count

                if row_count == 0:
                    report["tables"][table] = {"sqlite": 0, "mysql": 0, "status": "OK (empty)"}
                    continue

                col_names = [d[0] for d in sqlite_cur.description]
                cols_str = ", ".join([f"`{c}`" for c in col_names])
                placeholders = ", ".join(["%s"] * len(col_names))
                insert_sql = f"INSERT INTO `{table}` ({cols_str}) VALUES ({placeholders})"

                # Delete existing rows to prevent duplicate key errors on re-run
                my_cur.execute(f"DELETE FROM `{table}`")

                data = [tuple(r[c] for c in col_names) for r in rows]
                my_cur.executemany(insert_sql, data)
                mysql_conn.commit()

                # Verify count in MySQL
                my_cur.execute(f"SELECT COUNT(*) as cnt FROM `{table}`")
                res = my_cur.fetchone()
                my_cnt = res.get("cnt") or res.get(0) or 0
                report["total_mysql"] += my_cnt
                status = "MATCH" if my_cnt == row_count else f"MISMATCH ({row_count} vs {my_cnt})"
                report["tables"][table] = {"sqlite": row_count, "mysql": my_cnt, "status": status}
                logger.info(f"  [{status}] Table `{table}`: {row_count} -> {my_cnt} rows")

                # Reset AUTO_INCREMENT if 'id' column exists
                if "id" in col_names:
                    try:
                        my_cur.execute(f"SELECT COALESCE(MAX(id), 0) + 1 AS next_id FROM `{table}`")
                        next_id_row = my_cur.fetchone()
                        next_id = next_id_row.get("next_id") or next_id_row.get(0) or 1
                        my_cur.execute(f"ALTER TABLE `{table}` AUTO_INCREMENT = {next_id}")
                    except Exception:
                        pass

            my_cur.execute("SET FOREIGN_KEY_CHECKS = 1")
            mysql_conn.commit()

        # 3. Verify Key Records
        with mysql_conn.cursor() as my_cur:
            # Check machine TS1700-401113-65
            my_cur.execute("SELECT * FROM machines WHERE machine_id = 'TS1700-401113-65'")
            m_row = my_cur.fetchone()
            report["key_records"]["TS1700-401113-65"] = bool(m_row)

            # Check intervention INT-2026-0001
            my_cur.execute("SELECT * FROM interventions WHERE code = 'INT-2026-0001'")
            int_row = my_cur.fetchone()
            report["key_records"]["INT-2026-0001"] = {
                "exists": bool(int_row),
                "machine_id": int_row.get("machine_id") if int_row else None
            }

            # Check Owner and Admin
            my_cur.execute("SELECT username, role, is_active FROM users WHERE username IN ('owner', 'admin')")
            users = my_cur.fetchall()
            report["key_records"]["accounts"] = {u.get("username"): u.get("role") for u in users}

    finally:
        conn.close()
        mysql_conn.close()

    return True, report


def print_report(report: Dict[str, Any]):
    """Formats and prints the migration summary report."""
    print("\n" + "=" * 70)
    print("           SEBN-TN DATABASE MIGRATION VALIDATION REPORT")
    print("=" * 70)
    print(f"{'Table Name':<28} | {'SQLite Rows':<12} | {'MySQL Rows':<12} | Status")
    print("-" * 70)
    for tbl, info in report.get("tables", {}).items():
        st = "✅ " + info["status"] if "MATCH" in info["status"] or "OK" in info["status"] else "❌ " + info["status"]
        print(f"{tbl:<28} | {info['sqlite']:<12} | {info['mysql']:<12} | {st}")
    print("-" * 70)
    print(f"{'TOTAL':<28} | {report['total_sqlite']:<12} | {report['total_mysql']:<12} | {'✅ MATCH' if report['total_sqlite'] == report['total_mysql'] else '❌ MISMATCH'}")
    print("=" * 70)

    print("\n--- Key Records Verification ---")
    kr = report.get("key_records", {})
    ts_ok = "✅ Found" if kr.get("TS1700-401113-65") else "❌ NOT FOUND"
    print(f"  Machine TS1700-401113-65:      {ts_ok}")

    int_info = kr.get("INT-2026-0001", {})
    if isinstance(int_info, dict) and int_info.get("exists"):
        print(f"  Intervention INT-2026-0001:    ✅ Found (linked to machine: {int_info.get('machine_id')})")
    else:
        print(f"  Intervention INT-2026-0001:    ❌ NOT FOUND")

    accs = kr.get("accounts", {})
    print(f"  Default Accounts:              owner: {accs.get('owner', 'MISSING')}, admin: {accs.get('admin', 'MISSING')}")
    print("=" * 70 + "\n")


def main():
    parser = argparse.ArgumentParser(description="SEBN-TN SQLite to MySQL Migration Tool")
    parser.add_argument("--source-db", default=DEFAULT_SQLITE_PATH, help="Path to SQLite database")
    parser.add_argument("--schema-file", default=DEFAULT_SCHEMA_PATH, help="Path to MySQL schema SQL file")
    parser.add_argument("--dump-sql", help="Output path to export a standalone .sql migration file")
    parser.add_argument("--backup", action="store_true", default=True, help="Create a backup of SQLite DB first")
    parser.add_argument("--host", help="MySQL Host")
    parser.add_argument("--port", type=int, help="MySQL Port")
    parser.add_argument("--user", help="MySQL User")
    parser.add_argument("--password", help="MySQL Password")
    parser.add_argument("--database", help="MySQL Database name")
    args = parser.parse_args()

    if not os.path.exists(args.source_db):
        logger.error(f"Source database not found: {args.source_db}")
        sys.exit(1)

    if args.backup:
        create_sqlite_backup(args.source_db)

    # If --dump-sql requested, generate file
    if args.dump_sql:
        export_to_sql_file(args.source_db, args.schema_file, args.dump_sql)
        print(f"\n✅ Standalone migration SQL script generated at:\n   {os.path.abspath(args.dump_sql)}\n")
        if not (args.host or os.environ.get("DB_HOST")):
            print("No MySQL live connection credentials provided; exiting after dump generation.")
            return

    # Prepare MySQL credentials
    cfg = get_mysql_config()
    if args.host:
        cfg["host"] = args.host
    if args.port:
        cfg["port"] = args.port
    if args.user:
        cfg["user"] = args.user
    if args.password is not None:
        cfg["password"] = args.password
    if args.database:
        cfg["database"] = args.database

    print("\nConnecting to MySQL with configuration:")
    print(f"  Host:     {cfg.get('host')}")
    print(f"  Port:     {cfg.get('port')}")
    print(f"  Database: {cfg.get('database')}")
    print(f"  User:     {cfg.get('user')}\n")

    ok, msg = test_mysql_connection(cfg)
    if not ok:
        print(f"❌ Cannot connect to MySQL server: {msg}\n")
        # Also auto-generate the dump so user always has the SQL ready
        auto_dump = os.path.join(_PROJECT_ROOT, "scripts", "sebn_ima_migration.sql")
        export_to_sql_file(args.source_db, args.schema_file, auto_dump)
        print(f"💡 Generated standalone SQL file ready for import:\n   {auto_dump}\n")
        print("You can import it into MySQL/MariaDB anytime using:")
        print(f"   mysql -h {cfg.get('host')} -u {cfg.get('user')} -p {cfg.get('database')} < {auto_dump}\n")
        sys.exit(1)

    print("✅ MySQL connection successful! Starting migration...\n")
    success, report = run_live_migration(args.source_db, args.schema_file, cfg)
    print_report(report)


if __name__ == "__main__":
    main()
