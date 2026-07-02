import sys
from pathlib import Path

import database as db

ROOT = Path(__file__).resolve().parent.parent
USERS_FILE = ROOT / "usuarios" / "USUARIOS.txt"


def parse_line(line: str, line_no: int):
    line = line.strip()
    if not line or line.startswith("#"):
        return None
    parts = [p.strip() for p in line.split(",")]
    if len(parts) < 2:
        print(f"Linea {line_no}: formato invalido (use usuario,contrasena,rol)")
        return None
    username, password = parts[0], parts[1]
    role = parts[2] if len(parts) > 2 and parts[2] else "user"
    if len(username) < 3:
        print(f"Linea {line_no}: usuario muy corto ({username})")
        return None
    if len(password) < 6:
        print(f"Linea {line_no}: contrasena muy corta para {username}")
        return None
    if role not in ("user", "admin"):
        print(f"Linea {line_no}: rol invalido ({role}), use user o admin")
        return None
    if username.lower() == "admin":
        print(f"Linea {line_no}: use USUARIOS.txt solo para operadores; Admin es aparte")
        return None
    return username, password, role


def main() -> None:
    if not USERS_FILE.exists():
        print(f"No existe: {USERS_FILE}")
        sys.exit(1)

    db.init_db()
    db.seed_admin()

    created = 0
    updated = 0
    skipped = 0

    for i, raw in enumerate(USERS_FILE.read_text(encoding="utf-8").splitlines(), start=1):
        parsed = parse_line(raw, i)
        if not parsed:
            if raw.strip() and not raw.strip().startswith("#"):
                skipped += 1
            continue
        username, password, role = parsed
        result = db.upsert_user(username, password, role)
        if result == "created":
            created += 1
            print(f"  + {username} ({role})")
        else:
            updated += 1
            print(f"  ~ {username} actualizado ({role})")

    print("")
    print(f"Listo: {created} nuevos, {updated} actualizados, {skipped} omitidos")
    print(f"Archivo: {USERS_FILE}")
    print("Admin: Admin / Admin123!")


if __name__ == "__main__":
    main()
