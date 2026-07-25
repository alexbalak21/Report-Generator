import sqlite3
import json
import os
import sys
import datetime


def _get_db_path() -> str:
    # When frozen by PyInstaller, use the temporary _MEIPASS directory.
    # When running normally, use the directory of this file.
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    db_path = os.path.join(base_dir, "app_data.db")

    # Ensure the directory exists (important when running from the .exe)
    os.makedirs(os.path.dirname(db_path), exist_ok=True)

    return db_path


def _get_connection():
    db_path = _get_db_path()
    conn = sqlite3.connect(db_path)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS reports (
            id            TEXT PRIMARY KEY,
            created_at    TEXT NOT NULL,
            excel_path    TEXT NOT NULL,
            template_path TEXT NOT NULL,
            mapping_path  TEXT NOT NULL,
            row_number    INTEGER NOT NULL,
            data_json     TEXT NOT NULL
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS report_state (
            key   TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)

    conn.commit()
    return conn


# ── Report log ────────────────────────────────────────────────────────────────

def save_report(report_id: str, created_at: str, excel_path: str,
                template_path: str, mapping_path: str,
                row_number: int, data: dict) -> None:
    with _get_connection() as conn:
        conn.execute(
            """
            INSERT INTO reports
                (id, created_at, excel_path, template_path, mapping_path, row_number, data_json)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (report_id, created_at, excel_path, template_path, mapping_path,
             row_number, json.dumps(data, ensure_ascii=False)),
        )
        conn.commit()


def list_reports() -> list[dict]:
    with _get_connection() as conn:
        rows = conn.execute(
            "SELECT id, created_at, excel_path, template_path, mapping_path, row_number "
            "FROM reports ORDER BY created_at DESC"
        ).fetchall()
        return [
            {
                "id": r[0], "created_at": r[1], "excel_path": r[2],
                "template_path": r[3], "mapping_path": r[4], "row_number": r[5],
            }
            for r in rows
        ]


# ── Report state (replaces report_state.json) ─────────────────────────────────

def get_last_report_number() -> str | None:
    with _get_connection() as conn:
        row = conn.execute(
            "SELECT value FROM report_state WHERE key = 'last_report_number'"
        ).fetchone()
        return row[0] if row else None


def set_last_report_number(value: str) -> None:
    with _get_connection() as conn:
        conn.execute(
            "INSERT INTO report_state (key, value) VALUES ('last_report_number', ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (value,),
        )
        conn.commit()
