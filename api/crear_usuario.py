import sys

import database as db


def main() -> None:
    if len(sys.argv) != 3:
        print("Uso: python crear_usuario.py <usuario> <contrasena>")
        print("Ejemplo: python crear_usuario.py maria.garcia MiClave123")
        sys.exit(1)

    username = sys.argv[1].strip()
    password = sys.argv[2]

    if len(username) < 3:
        print("El usuario debe tener al menos 3 caracteres.")
        sys.exit(1)
    if len(password) < 6:
        print("La contrasena debe tener al menos 6 caracteres.")
        sys.exit(1)
    if username.lower() == "admin":
        print("Use la cuenta Admin existente o otro nombre.")
        sys.exit(1)

    db.init_db()
    if db.find_user(username):
        print(f"El usuario '{username}' ya existe.")
        sys.exit(1)

    db.create_user(username, password)
    print(f"Usuario creado: {username}")


if __name__ == "__main__":
    main()
