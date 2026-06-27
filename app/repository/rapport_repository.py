import sqlite3
import json
import os
import datetime

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "app_data.db")


def _get_connection():
    conn = sqlite3.connect(os.path.abspath(DB_PATH))

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
