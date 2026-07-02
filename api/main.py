import os
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

import database as db

load_dotenv()

app = FastAPI(title="Olnatura Auth API", version="2.0.0")


class LoginRequest(BaseModel):
    username: str = Field(min_length=1, max_length=80)
    password: str = Field(min_length=1, max_length=120)
    workbook: Optional[str] = Field(default=None, max_length=200)


class AdminAuth(BaseModel):
    admin_username: str = Field(min_length=1, max_length=80)
    admin_password: str = Field(min_length=1, max_length=120)


class AdminCreateUser(AdminAuth):
    username: str = Field(min_length=3, max_length=80)
    password: str = Field(min_length=6, max_length=120)
    role: str = Field(default="user", max_length=20)


class AdminSetActive(AdminAuth):
    is_active: bool = True


def _require_admin(body: AdminAuth):
    admin = db.verify_user(body.admin_username.strip(), body.admin_password)
    if not admin or admin["role"] != "admin":
        raise HTTPException(status_code=403, detail="Solo administradores")
    return admin


@app.on_event("startup")
def startup() -> None:
    db.init_db()
    db.seed_admin()


@app.get("/health")
def health():
    return {"ok": True, "service": "olnatura-auth"}


@app.post("/login")
def login(body: LoginRequest):
    user = db.verify_user(body.username.strip(), body.password)
    if not user:
        raise HTTPException(status_code=401, detail="Credenciales invalidas o cuenta inactiva")

    db.write_audit(
        username=user["username"],
        user_id=user["id"],
        workbook=body.workbook,
        action="login",
    )
    return {
        "ok": True,
        "username": user["username"],
        "role": user["role"],
    }


@app.get("/audit")
def audit(limit: int = 50):
    return {"ok": True, "items": db.list_audit(limit=min(limit, 200))}


@app.post("/admin/users/list")
def admin_users_list(body: AdminAuth):
    _require_admin(body)
    users = db.list_users()
    lines = []
    for u in users:
        active = "1" if u["is_active"] else "0"
        lines.append(f"{u['username']}|{u['role']}|{active}|{u['created_at']}")
    return {"ok": True, "data": "\n".join(lines)}


@app.post("/admin/users/create")
def admin_users_create(body: AdminCreateUser):
    _require_admin(body)
    username = body.username.strip()
    if username.lower() == "admin":
        raise HTTPException(status_code=400, detail="Nombre reservado")
    if db.find_user(username):
        raise HTTPException(status_code=409, detail="El usuario ya existe")
    role = body.role if body.role in ("user", "admin") else "user"
    db.create_user(username, body.password, role)
    db.write_audit(username=body.admin_username.strip(), user_id=None, workbook=None, action=f"admin_create:{username}")
    return {"ok": True, "username": username, "role": role}


@app.post("/admin/users/{username}/set-active")
def admin_users_set_active(username: str, body: AdminSetActive):
    _require_admin(body)
    uname = username.strip()
    if uname.lower() == "admin":
        raise HTTPException(status_code=400, detail="No se puede desactivar Admin")
    if not db.find_user(uname):
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    db.set_user_active(uname, body.is_active)
    action = "admin_activate" if body.is_active else "admin_deactivate"
    db.write_audit(username=body.admin_username.strip(), user_id=None, workbook=None, action=f"{action}:{uname}")
    return {"ok": True, "username": uname, "is_active": body.is_active}


if __name__ == "__main__":
    import uvicorn

    host = os.getenv("API_HOST", "127.0.0.1")
    port = int(os.getenv("API_PORT", "8000"))
    uvicorn.run("main:app", host=host, port=port, reload=False)
