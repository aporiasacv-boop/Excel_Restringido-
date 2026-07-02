import hashlib
import os
import secrets
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

DB_PATH = Path(os.getenv("DATABASE_PATH", "data/olnatura.db"))


def _utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def hash_password(password: str, salt: str) -> str:
    payload = f"{salt}|{password}".encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def new_salt() -> str:
    return secrets.token_hex(16)


@contextmanager
def get_conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE COLLATE NOCASE,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                username TEXT NOT NULL,
                workbook TEXT,
                action TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (user_id) REFERENCES users(id)
            );
            """
        )


def seed_admin() -> None:
    with get_conn() as conn:
        row = conn.execute(
            "SELECT id FROM users WHERE username = ? COLLATE NOCASE", ("Admin",)
        ).fetchone()
        if row:
            return
        salt = new_salt()
        conn.execute(
            """
            INSERT INTO users (username, password_hash, salt, role, is_active, created_at)
            VALUES (?, ?, ?, ?, 1, ?)
            """,
            ("Admin", hash_password("Admin123!", salt), salt, "admin", _utc_now()),
        )


def find_user(username: str):
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM users WHERE username = ? COLLATE NOCASE", (username,)
        ).fetchone()


def create_user(username: str, password: str, role: str = "user") -> None:
    salt = new_salt()
    with get_conn() as conn:
        conn.execute(
            """
            INSERT INTO users (username, password_hash, salt, role, is_active, created_at)
            VALUES (?, ?, ?, ?, 1, ?)
            """,
            (username, hash_password(password, salt), salt, role, _utc_now()),
        )


def upsert_user(username: str, password: str, role: str = "user") -> str:
    """Crea o actualiza contraseña/rol. Devuelve 'created' o 'updated'."""
    salt = new_salt()
    pwd_hash = hash_password(password, salt)
    with get_conn() as conn:
        row = conn.execute(
            "SELECT id FROM users WHERE username = ? COLLATE NOCASE", (username,)
        ).fetchone()
        if row:
            conn.execute(
                """
                UPDATE users
                SET password_hash = ?, salt = ?, role = ?, is_active = 1
                WHERE id = ?
                """,
                (pwd_hash, salt, role, row["id"]),
            )
            return "updated"
        conn.execute(
            """
            INSERT INTO users (username, password_hash, salt, role, is_active, created_at)
            VALUES (?, ?, ?, ?, 1, ?)
            """,
            (username, pwd_hash, salt, role, _utc_now()),
        )
        return "created"


def verify_user(username: str, password: str):
    user = find_user(username)
    if not user:
        return None
    if not user["is_active"]:
        return None
    expected = hash_password(password, user["salt"])
    if expected != user["password_hash"]:
        return None
    return user


def write_audit(username: str, user_id: Optional[int], workbook: Optional[str], action: str) -> None:
    with get_conn() as conn:
        conn.execute(
            """
            INSERT INTO audit_log (user_id, username, workbook, action, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (user_id, username, workbook, action, _utc_now()),
        )


def list_audit(limit: int = 100):
    with get_conn() as conn:
        rows = conn.execute(
            """
            SELECT id, username, workbook, action, created_at
            FROM audit_log
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
        return [dict(row) for row in rows]


def list_users():
    with get_conn() as conn:
        rows = conn.execute(
            """
            SELECT username, role, is_active, created_at
            FROM users
            ORDER BY username COLLATE NOCASE
            """
        ).fetchall()
        return [dict(row) for row in rows]


def set_user_active(username: str, is_active: bool) -> bool:
    with get_conn() as conn:
        cur = conn.execute(
            """
            UPDATE users SET is_active = ?
            WHERE username = ? COLLATE NOCASE
            """,
            (1 if is_active else 0, username),
        )
        return cur.rowcount > 0
