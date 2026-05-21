"""
auth.py — Autenticación JWT para la Malla Validadora RIPS
Usuarios almacenados en users.json. Tokens con expiración de 8 horas (jornada laboral).
"""

from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

import bcrypt
from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from jose import JWTError, jwt

# ── Configuración ─────────────────────────────────────────────────────────────
SECRET_KEY   = os.environ.get("JWT_SECRET", "drf-malla-validadora-2275-clave-secreta-cambiar-en-produccion")
ALGORITHM    = "HS256"
TOKEN_HORAS  = 8

USERS_FILE   = Path(__file__).parent / "users.json"
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/login")

# ── Carga de usuarios ─────────────────────────────────────────────────────────

def _cargar_usuarios() -> list[dict]:
    if not USERS_FILE.exists():
        return []
    try:
        return json.loads(USERS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return []


def _guardar_usuarios(users: list[dict]) -> None:
    USERS_FILE.write_text(json.dumps(users, indent=2, ensure_ascii=False), encoding="utf-8")

# ── Contraseñas ───────────────────────────────────────────────────────────────

def verificar_password(plain: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(plain.encode(), hashed.encode())
    except Exception:
        return False


def hashear_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode(), bcrypt.gensalt()).decode()

# ── Autenticación ─────────────────────────────────────────────────────────────

def autenticar_usuario(username: str, password: str) -> Optional[dict]:
    for u in _cargar_usuarios():
        if u.get("username") == username:
            if verificar_password(password, u.get("password", "")):
                return u
    return None

# ── Tokens JWT ────────────────────────────────────────────────────────────────

def crear_token(username: str, role: str, nombre: str) -> str:
    expire = datetime.utcnow() + timedelta(hours=TOKEN_HORAS)
    payload = {
        "sub":    username,
        "role":   role,
        "nombre": nombre,
        "exp":    expire,
    }
    return jwt.encode(payload, SECRET_KEY, algorithm=ALGORITHM)


def get_usuario_actual(token: str = Depends(oauth2_scheme)) -> dict:
    exc = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Token inválido o expirado. Inicie sesión nuevamente.",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        username: str = payload.get("sub", "")
        if not username:
            raise exc
        return {
            "username": username,
            "role":     payload.get("role", "auditor"),
            "nombre":   payload.get("nombre", username),
        }
    except JWTError:
        raise exc


def solo_admin(usuario: dict = Depends(get_usuario_actual)) -> dict:
    if usuario.get("role") != "admin":
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Se requiere rol de administrador.",
        )
    return usuario

# ── Gestión de usuarios (solo admin) ─────────────────────────────────────────

def listar_usuarios() -> list[dict]:
    return [
        {"username": u["username"], "role": u["role"], "nombre": u.get("nombre", "")}
        for u in _cargar_usuarios()
    ]


def _validar_password(password: str) -> None:
    if len(password) < 8:
        raise HTTPException(status_code=400, detail="La contraseña debe tener al menos 8 caracteres.")


def crear_usuario(username: str, password: str, role: str, nombre: str) -> dict:
    _validar_password(password)
    users = _cargar_usuarios()
    if any(u["username"] == username for u in users):
        raise HTTPException(status_code=400, detail=f"El usuario '{username}' ya existe.")
    nuevo = {
        "username": username,
        "password": hashear_password(password),
        "role":     role,
        "nombre":   nombre,
    }
    users.append(nuevo)
    _guardar_usuarios(users)
    return {"username": username, "role": role, "nombre": nombre}


def eliminar_usuario(username: str) -> None:
    users = _cargar_usuarios()
    nuevos = [u for u in users if u["username"] != username]
    if len(nuevos) == len(users):
        raise HTTPException(status_code=404, detail=f"Usuario '{username}' no encontrado.")
    if not any(u["role"] == "admin" for u in nuevos):
        raise HTTPException(status_code=400, detail="No se puede eliminar el último administrador.")
    _guardar_usuarios(nuevos)


def cambiar_password(username: str, nueva: str) -> None:
    _validar_password(nueva)
    users = _cargar_usuarios()
    for u in users:
        if u["username"] == username:
            u["password"] = hashear_password(nueva)
            _guardar_usuarios(users)
            return
    raise HTTPException(status_code=404, detail=f"Usuario '{username}' no encontrado.")


def cambiar_password_verificado(username: str, actual: str, nueva: str) -> None:
    """Cambia contraseña verificando la actual. Para uso self-service."""
    _validar_password(nueva)
    users = _cargar_usuarios()
    for u in users:
        if u["username"] == username:
            if not verificar_password(actual, u.get("password", "")):
                raise HTTPException(status_code=400, detail="La contraseña actual es incorrecta.")
            u["password"] = hashear_password(nueva)
            _guardar_usuarios(users)
            return
    raise HTTPException(status_code=404, detail=f"Usuario '{username}' no encontrado.")
