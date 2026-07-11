import sqlite3
import os
import sys


def _get_db_path() -> str:
    if getattr(sys, "frozen", False):
        # Frozen exe: store DB in %APPDATA%\ReportGenerator\ (always writable)
        base_dir = os.path.join(os.environ["APPDATA"], "ReportGenerator")
    else:
        # Dev: store DB next to this file
        base_dir = os.path.dirname(os.path.abspath(__file__))

    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, "app_data.db")


def _get_connection():
    db_path = _get_db_path()
    conn = sqlite3.connect(db_path)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS config (
            key   TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS mappings (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            mapping_path TEXT NOT NULL UNIQUE
        )
    """)

    conn.commit()
    return conn


def config_get(key: str) -> str | None:
    with _get_connection() as conn:
        row = conn.execute(
            "SELECT value FROM config WHERE key = ?", (key,)
        ).fetchone()
        return row[0] if row else None


def config_set(key: str, value: str) -> None:
    with _get_connection() as conn:
        conn.execute(
            "INSERT INTO config (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, value),
        )
        conn.commit()


def mapping_add(mapping_path: str) -> int:
    with _get_connection() as conn:
        cursor = conn.execute(
            "INSERT OR IGNORE INTO mappings (mapping_path) VALUES (?)",
            (mapping_path,),
        )
        conn.commit()

        if cursor.lastrowid:
            return cursor.lastrowid

        row = conn.execute(
            "SELECT id FROM mappings WHERE mapping_path = ?",
            (mapping_path,),
        ).fetchone()

        return row[0] if row else 0


def mapping_list() -> list[tuple[int, str]]:
    with _get_connection() as conn:
        rows = conn.execute(
            "SELECT id, mapping_path FROM mappings ORDER BY id"
        ).fetchall()
        return [(row[0], row[1]) for row in rows]


def mapping_get_last() -> str | None:
    with _get_connection() as conn:
        row = conn.execute(
            "SELECT mapping_path FROM mappings ORDER BY id DESC LIMIT 1"
        ).fetchone()
        return row[0] if row else None


def mapping_delete(mapping_id: int) -> None:
    with _get_connection() as conn:
        conn.execute(
            "DELETE FROM mappings WHERE id = ?", (mapping_id,)
        )
        conn.commit()
