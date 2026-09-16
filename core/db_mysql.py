"""
SEBN-TN Enterprise Maintenance Suite — MySQL / MariaDB Database Driver
Centralized connection manager and query compatibility layer.
Supports both MySQL and fallback SQLite for local resilience.
"""
import os
import re
import sys
import logging
from typing import Any, Dict, List, Optional, Tuple, Union
from urllib.parse import urlparse

# Attempt to load .env from project root
try:
    from dotenv import load_dotenv
    # Look for .env in current dir or up to 2 parent dirs
    _curr = os.path.dirname(os.path.abspath(__file__))
    for _ in range(3):
        _env_path = os.path.join(_curr, ".env")
        if os.path.exists(_env_path):
            load_dotenv(_env_path)
            break
        _curr = os.path.dirname(_curr)
except ImportError:
    pass

logger = logging.getLogger("sebn-db")


class DictRow:
    """
    Enhanced row class 100% compatible with sqlite3.Row:
    - Sequence unpacking: a, b = row (unpacks column values)
    - Dict key access: row['id'], row.get('name')
    - Tuple index access: row[0], row[1]
    - Mapping conversion: dict(row) -> {'id': 1, ...}
    - Column introspection: row.keys() -> ['id', 'name', ...]
    - Automatic alias: row['name'] aliases row['Field'] for SHOW COLUMNS compatibility
    """
    def __init__(self, data: Optional[Dict[str, Any]] = None, column_names: Optional[List[str]] = None):
        self._data = dict(data) if data is not None else {}
        self._keys = list(column_names) if column_names else list(self._data.keys())

    def __getitem__(self, key: Union[str, int]) -> Any:
        if isinstance(key, int):
            if 0 <= key < len(self._keys):
                return self._data[self._keys[key]]
            elif -len(self._keys) <= key < 0:
                return self._data[self._keys[key]]
            raise IndexError(f"Column index {key} out of range ({len(self._keys)} columns)")
        if key == "name" and "name" not in self._data and "Field" in self._data:
            return self._data["Field"]
        return self._data[key]

    def __setitem__(self, key: str, value: Any):
        if key not in self._data and key not in self._keys:
            self._keys.append(key)
        self._data[key] = value

    def get(self, key: Union[str, int], default: Any = None) -> Any:
        try:
            return self[key]
        except (KeyError, IndexError):
            return default

    def keys(self) -> List[str]:
        return list(self._keys)

    def values(self) -> List[Any]:
        return [self._data[k] for k in self._keys]

    def items(self) -> List[Tuple[str, Any]]:
        return [(k, self._data[k]) for k in self._keys]

    def __iter__(self):
        for k in self._keys:
            yield self._data[k]

    def __len__(self) -> int:
        return len(self._keys)

    def __contains__(self, key: Any) -> bool:
        if key in self._data or key in self._keys:
            return True
        if key == "name" and "Field" in self._data:
            return True
        return False

    def __repr__(self) -> str:
        return repr(self._data)


def convert_placeholders(sql: str) -> str:
    """
    Safely converts SQLite '?' parameter placeholders to MySQL '%s'
    while skipping '?' inside single or double-quoted string literals.
    """
    if "?" not in sql:
        return sql
    parts = []
    last_end = 0
    # Match quoted strings or '?'
    pattern = re.compile(r"('([^'\\]|\\.)*'|\"([^\"\\]|\\.)*\"|\?)")
    for match in pattern.finditer(sql):
        token = match.group(0)
        if token == "?":
            parts.append(sql[last_end:match.start()])
            parts.append("%s")
            last_end = match.end()
    parts.append(sql[last_end:])
    return "".join(parts)


def adapt_sql_dialect(sql: str) -> str:
    """
    Translates common SQLite statements into MySQL equivalents.
    """
    cleaned = sql.strip()
    # 1. PRAGMA table_info(x) -> SHOW COLUMNS FROM x
    m = re.match(r"(?i)^PRAGMA\s+table_info\s*\(\s*\[?([a-zA-Z0-9_]+)\]?\s*\)", cleaned)
    if m:
        return f"SHOW COLUMNS FROM `{m.group(1)}`"

    # 2. PRAGMA foreign_keys = ON / OFF -> ignore in MySQL
    if re.match(r"(?i)^PRAGMA\s+foreign_keys", cleaned):
        return "SELECT 1"

    # 3. PRAGMA journal_mode / busy_timeout -> ignore in MySQL
    if re.match(r"(?i)^PRAGMA\s+(journal_mode|busy_timeout)", cleaned):
        return "SELECT 1"

    # 4. sqlite_master check -> SHOW TABLES
    if "sqlite_master" in cleaned.lower() and "table" in cleaned.lower():
        cleaned = re.sub(r"(?i)SELECT\s+name\s+FROM\s+sqlite_master\s+WHERE\s+type\s*=\s*'table'", "SHOW TABLES", cleaned)

    # 5. AUTOINCREMENT -> AUTO_INCREMENT
    cleaned = re.sub(r"(?i)\bAUTOINCREMENT\b", "AUTO_INCREMENT", cleaned)

    # 6. datetime('now') -> NOW()
    cleaned = re.sub(r"(?i)datetime\s*\(\s*'now'\s*\)", "NOW()", cleaned)

    # 7. INSERT OR REPLACE INTO -> REPLACE INTO
    cleaned = re.sub(r"(?i)\bINSERT\s+OR\s+REPLACE\s+INTO\b", "REPLACE INTO", cleaned)

    # 8. ON CONFLICT(...) DO UPDATE SET -> ON DUPLICATE KEY UPDATE
    cleaned = re.sub(r"(?i)ON\s+CONFLICT\s*\([^)]*\)\s*DO\s+UPDATE\s+SET", "ON DUPLICATE KEY UPDATE", cleaned)
    cleaned = re.sub(r"(?i)\bexcluded\.([a-zA-Z0-9_]+)\b", r"VALUES(\1)", cleaned)

    # 9. Quote reserved keyword `key` in settings queries
    cleaned = re.sub(r"(?i)\bWHERE\s+key\s*=", "WHERE `key` =", cleaned)
    cleaned = re.sub(r"(?i)\(\s*key\s*,", "(`key`,", cleaned)

    # 10. Convert named parameters :name -> %(name)s (when not part of a string or ::)
    cleaned = re.sub(r"(?<!:):([a-zA-Z0-9_]+)\b", r"%(\1)s", cleaned)

    # 11. Convert placeholders (? -> %s)
    return convert_placeholders(cleaned)


def get_mysql_config() -> Dict[str, Any]:
    """Reads MySQL credentials from environment variables."""
    # Check DATABASE_URL / MYSQL_URL first
    url = os.environ.get("MYSQL_URL") or os.environ.get("DATABASE_URL") or os.environ.get("JAWSDB_URL")
    if url and (url.startswith("mysql://") or url.startswith("mysql2://")):
        parsed = urlparse(url)
        return {
            "host": parsed.hostname or "localhost",
            "port": parsed.port or 3306,
            "user": parsed.username or "root",
            "password": parsed.password or "",
            "database": (parsed.path or "/sebn_ima").lstrip("/"),
            "charset": "utf8mb4",
            "collation": "utf8mb4_unicode_ci",
            "connect_timeout": 10
        }

    return {
        "host": os.environ.get("DB_HOST", "localhost"),
        "port": int(os.environ.get("DB_PORT", 3306)),
        "user": os.environ.get("DB_USER", os.environ.get("MYSQL_USER", "root")),
        "password": os.environ.get("DB_PASSWORD", os.environ.get("MYSQL_PASSWORD", "")),
        "database": os.environ.get("DB_NAME", os.environ.get("MYSQL_DATABASE", "sebn_ima")),
        "charset": "utf8mb4",
        "collation": "utf8mb4_unicode_ci",
        "connect_timeout": 10
    }


def is_mysql_configured() -> bool:
    """Returns True if MySQL environment variables or URL are explicitly configured."""
    if os.environ.get("MYSQL_URL") or os.environ.get("DATABASE_URL"):
        return True
    return bool(os.environ.get("DB_HOST") or os.environ.get("DB_USER"))


class MySQLCursorWrapper:
    """Wraps a mysql.connector cursor to provide DictRow instances and dialect adaptation."""
    def __init__(self, raw_cursor):
        self._raw = raw_cursor

    @property
    def lastrowid(self):
        return self._raw.lastrowid

    @property
    def rowcount(self):
        return self._raw.rowcount

    @property
    def description(self):
        return self._raw.description

    def execute(self, operation: str, params: Optional[Union[Tuple, List, Dict]] = None):
        adapted_sql = adapt_sql_dialect(operation)
        if params is not None:
            # If params is a single-item tuple/list, pass it
            return self._raw.execute(adapted_sql, params)
        return self._raw.execute(adapted_sql)

    def executemany(self, operation: str, seq_of_params: List[Any]):
        adapted_sql = adapt_sql_dialect(operation)
        return self._raw.executemany(adapted_sql, seq_of_params)

    def fetchone(self) -> Optional[DictRow]:
        row = self._raw.fetchone()
        if row is None:
            return None
        if isinstance(row, dict):
            return DictRow(row)
        # Tuple cursor
        col_names = [d[0] for d in self._raw.description] if self._raw.description else []
        d = dict(zip(col_names, row))
        return DictRow(d, column_names=col_names)

    def fetchall(self) -> List[DictRow]:
        rows = self._raw.fetchall()
        if not rows:
            return []
        if isinstance(rows[0], dict):
            return [DictRow(r) for r in rows]
        col_names = [d[0] for d in self._raw.description] if self._raw.description else []
        return [DictRow(dict(zip(col_names, r)), column_names=col_names) for r in rows]

    def fetchmany(self, size: Optional[int] = None) -> List[DictRow]:
        rows = self._raw.fetchmany(size) if size else self._raw.fetchmany()
        if not rows:
            return []
        if isinstance(rows[0], dict):
            return [DictRow(r) for r in rows]
        col_names = [d[0] for d in self._raw.description] if self._raw.description else []
        return [DictRow(dict(zip(col_names, r)), column_names=col_names) for r in rows]

    def close(self):
        try:
            self._raw.close()
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


class MySQLConnectionWrapper:
    """
    Wraps mysql.connector.connection.MySQLConnection:
    - Provides with conn as conn context manager that commits on success and closes on exit
    - Provides conn.execute(sql, params) for direct query execution
    - Provides conn.cursor() yielding MySQLCursorWrapper
    """
    def __init__(self, raw_conn):
        self._raw = raw_conn

    def cursor(self, dictionary: bool = True, **kwargs):
        # We always use dictionary=True for uniform dict row access
        raw_cur = self._raw.cursor(dictionary=dictionary, **kwargs)
        return MySQLCursorWrapper(raw_cur)

    def execute(self, sql: str, params: Optional[Union[Tuple, List, Dict]] = None):
        cur = self.cursor()
        cur.execute(sql, params)
        return cur

    def executemany(self, sql: str, seq_of_params: List[Any]):
        cur = self.cursor()
        cur.executemany(sql, seq_of_params)
        return cur

    def executescript(self, script: str):
        statements = [s.strip() for s in script.split(";") if s.strip()]
        cur = self.cursor()
        for stmt in statements:
            try:
                cur.execute(stmt)
            except Exception as e:
                logger.warning(f"[DB executescript] statement warning: {e}")
        self.commit()

    @property
    def row_factory(self):
        return None

    @row_factory.setter
    def row_factory(self, val):
        pass

    def commit(self):
        self._raw.commit()

    def rollback(self):
        try:
            self._raw.rollback()
        except Exception:
            pass

    def close(self):
        try:
            self._raw.close()
        except Exception:
            pass

    def is_connected(self) -> bool:
        try:
            return self._raw.is_connected()
        except Exception:
            return False

    def ping(self, reconnect: bool = True, attempts: int = 1, delay: int = 0):
        return self._raw.ping(reconnect=reconnect, attempts=attempts, delay=delay)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None:
            self.rollback()
        else:
            try:
                self.commit()
            except Exception:
                pass
        self.close()


def get_mysql_connection(config: Optional[Dict[str, Any]] = None) -> MySQLConnectionWrapper:
    """
    Connects to MySQL/MariaDB and returns a MySQLConnectionWrapper.
    Raises mysql.connector.Error if connection fails.
    """
    import mysql.connector
    cfg = config or get_mysql_config()
    raw = mysql.connector.connect(**cfg)
    return MySQLConnectionWrapper(raw)


def test_mysql_connection(config: Optional[Dict[str, Any]] = None) -> Tuple[bool, str]:
    """
    Tests if MySQL server is reachable with the given or configured credentials.
    Returns (success: bool, message: str).
    """
    try:
        conn = get_mysql_connection(config)
        with conn.cursor() as cur:
            cur.execute("SELECT 1 AS ok")
            row = cur.fetchone()
        conn.close()
        return True, "Connexion MySQL établie avec succès."
    except Exception as e:
        return False, str(e)


def get_db_connection(sqlite_fallback_path: Optional[str] = None):
    """
    Smart connection factory:
    1. If MySQL is configured and reachable, connects to MySQL.
    2. If MySQL is not reachable or not configured, falls back to SQLite
       using sqlite_fallback_path (if provided) to ensure continuous operation.
    """
    if is_mysql_configured():
        try:
            return get_mysql_connection()
        except Exception as e:
            if not sqlite_fallback_path:
                raise
            logger.warning(f"[DB] MySQL connection failed ({e}). Falling back to SQLite: {sqlite_fallback_path}")

    # Fallback to SQLite if path provided
    if sqlite_fallback_path and os.path.exists(sqlite_fallback_path):
        import sqlite3
        conn = sqlite3.connect(sqlite_fallback_path, timeout=15)
        conn.row_factory = sqlite3.Row
        return conn

    # Attempt MySQL directly and surface the error
    return get_mysql_connection()
