import sys

import database as db


def main() -> None:
    db.init_db()
    with db.get_conn() as conn:
        rows = conn.execute(
            """
            SELECT username, role, is_active, created_at
            FROM users
            ORDER BY id
            """
        ).fetchall()

    if not rows:
        print("No hay usuarios.")
        return

    print(f"{'Usuario':<24} {'Rol':<8} {'Activo':<8} Creado")
    print("-" * 60)
    for row in rows:
        activo = "SI" if row["is_active"] else "NO"
        print(f"{row['username']:<24} {row['role']:<8} {activo:<8} {row['created_at']}")


if __name__ == "__main__":
    main()
