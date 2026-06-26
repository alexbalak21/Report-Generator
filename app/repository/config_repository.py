import sqlite3
import os


DB_PATH = os.path.join(os.path.dirname(__file__), "..", "app_data.db")


def _get_connection():
    conn = sqlite3.connect(os.path.abspath(DB_PATH))
    conn.execute("""
        CREATE TABLE IF NOT EXISTS config (
            key   TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)
    conn.commit()
    return conn


def config_get(key: str) -> str | None:
    with _get_connection() as conn:
        row = conn.execute("SELECT value FROM config WHERE key = ?", (key,)).fetchone()
        return row[0] if row else None


def config_set(key: str, value: str) -> None:
    with _get_connection() as conn:
        conn.execute(
            "INSERT INTO config (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, value),
        )
        conn.commit()
