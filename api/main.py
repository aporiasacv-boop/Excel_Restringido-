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


if __name__ == "__main__":
    import uvicorn

    host = os.getenv("API_HOST", "127.0.0.1")
    port = int(os.getenv("API_PORT", "8000"))
    uvicorn.run("main:app", host=host, port=port, reload=False)
