import sqlite3
import os
import sys


# Optional override of the database file (tests set this).
DB_PATH: str | None = None


def _get_db_path() -> str:
    override = DB_PATH or os.environ.get("REPORT_GENERATOR_DB")
    if override:
        os.makedirs(os.path.dirname(os.path.abspath(override)), exist_ok=True)
        return override

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

    conn.execute("""
        CREATE TABLE IF NOT EXISTS profiles (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            slug        TEXT NOT NULL UNIQUE,
            name        TEXT NOT NULL,
            folder      TEXT NOT NULL,
            next_line   INTEGER NOT NULL DEFAULT 2,
            output_dir  TEXT,
            last_used   TEXT
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



# ── Profiles ──────────────────────────────────────────────────────────────────

_PROFILE_COLUMNS = ("slug", "name", "folder", "next_line", "output_dir", "last_used")
_PROFILE_UPDATABLE = {"name", "folder", "next_line", "output_dir", "last_used"}


def _profile_row(row) -> dict:
    return dict(zip(_PROFILE_COLUMNS, row))


def profile_list() -> list[dict]:
    with _get_connection() as conn:
        rows = conn.execute(
            f"SELECT {', '.join(_PROFILE_COLUMNS)} FROM profiles ORDER BY name COLLATE NOCASE"
        ).fetchall()
        return [_profile_row(r) for r in rows]


def profile_get(slug: str) -> dict | None:
    with _get_connection() as conn:
        row = conn.execute(
            f"SELECT {', '.join(_PROFILE_COLUMNS)} FROM profiles WHERE slug = ?", (slug,)
        ).fetchone()
        return _profile_row(row) if row else None


def profile_add(slug: str, name: str, folder: str,
                next_line: int = 2, output_dir: str | None = None) -> None:
    with _get_connection() as conn:
        conn.execute(
            "INSERT INTO profiles (slug, name, folder, next_line, output_dir) "
            "VALUES (?, ?, ?, ?, ?)",
            (slug, name, folder, next_line, output_dir),
        )
        conn.commit()


def profile_update(slug: str, **fields) -> None:
    unknown = set(fields) - _PROFILE_UPDATABLE
    if unknown:
        raise ValueError(f"Unknown profile field(s): {', '.join(sorted(unknown))}")
    if not fields:
        return
    assignments = ", ".join(f"{k} = ?" for k in fields)
    with _get_connection() as conn:
        conn.execute(
            f"UPDATE profiles SET {assignments} WHERE slug = ?",
            (*fields.values(), slug),
        )
        conn.commit()


def profile_delete(slug: str) -> None:
    with _get_connection() as conn:
        conn.execute("DELETE FROM profiles WHERE slug = ?", (slug,))
        conn.commit()
