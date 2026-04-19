"""Usuarios y contraseñas desde JSON (texto plano).

- Por defecto: ``auth_users.json`` junto a este archivo (no versionar: ver ``.gitignore``).
- En el servidor: definir la variable de entorno ``AUTH_USERS_PATH`` con la ruta absoluta
  al JSON que solo exista en disco del host (o volumen), sin pasar por Git.
- Si no hay archivo o está vacío: ``AUTH_ADMIN_USER`` y ``AUTH_ADMIN_PASSWORD`` (variables de entorno).

Copiá ``auth_users.example.json`` → ``auth_users.json`` en cada entorno y editá usuarios/claves.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from urllib.parse import quote

_AUTH_DIR = Path(__file__).resolve().parent
_override = (os.environ.get("AUTH_USERS_PATH") or "").strip()
AUTH_USERS_FILE = Path(_override) if _override else _AUTH_DIR / "auth_users.json"


def _normalizar_usuarios(raw: dict) -> dict[str, str]:
    out: dict[str, str] = {}
    for k, v in (raw or {}).items():
        ks = str(k).strip()
        vs = str(v).strip() if v is not None else ""
        if ks:
            out[ks] = vs
    return out


def load_users() -> dict[str, str]:
    """Devuelve mapa usuario -> contraseña (texto plano, editable en el archivo)."""
    if AUTH_USERS_FILE.is_file():
        try:
            with open(AUTH_USERS_FILE, encoding="utf-8-sig") as f:
                data = json.load(f)
            got = _normalizar_usuarios(data.get("users") or {})
            if got:
                return got
        except (OSError, json.JSONDecodeError):
            pass
    # Sin archivo o JSON vacío (p. ej. deploy sin subir claves): credenciales por entorno
    u = (os.environ.get("AUTH_ADMIN_USER") or "").strip()
    p = (os.environ.get("AUTH_ADMIN_PASSWORD") or "").strip()
    if u and p:
        return {u: p}
    return {}


def verify_credentials(username: str, password: str) -> bool:
    users = load_users()
    u = (username or "").strip()
    pwd = (password or "").strip()
    if not u:
        return False
    return users.get(u) == pwd


def whatsapp_new_user_url() -> str:
    msg = (
        "Buen día! Quiero generar mi usuario para el sistema Análisis Mis Comprobantes"
    )
    return f"https://wa.me/5493513132914?text={quote(msg)}"
