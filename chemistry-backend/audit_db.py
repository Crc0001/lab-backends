import os
import sqlite3
from pathlib import Path
from contextlib import contextmanager

DB_PATH = Path(os.getenv("AUDIT_DB_PATH", "audit.db"))


def init_db():
    with _conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS audit_logs (
                id    INTEGER PRIMARY KEY AUTOINCREMENT,
                time  TEXT    NOT NULL,
                user  TEXT    NOT NULL,
                type  TEXT    NOT NULL,
                desc  TEXT    NOT NULL
            )
        """)


@contextmanager
def _conn():
    # check_same_thread=False + timeout 防止并发写入锁死
    conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def insert_log(time: str, user: str, log_type: str, desc: str):
    with _conn() as conn:
        conn.execute(
            "INSERT INTO audit_logs (time, user, type, desc) VALUES (?, ?, ?, ?)",
            (time, user, log_type, desc),
        )


def query_logs(user: str | None, limit: int, offset: int) -> list[dict]:
    with _conn() as conn:
        if user:
            rows = conn.execute(
                "SELECT * FROM audit_logs WHERE user=? ORDER BY id DESC LIMIT ? OFFSET ?",
                (user, limit, offset),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM audit_logs ORDER BY id DESC LIMIT ? OFFSET ?",
                (limit, offset),
            ).fetchall()
    return [dict(r) for r in rows]

