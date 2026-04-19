"""Usuarios y contraseñas desde JSON (texto plano).

- Por defecto: ``auth_users.json`` junto a este archivo (no versionar: ver ``.gitignore``).
- En el servidor: definir la variable de entorno ``AUTH_USERS_PATH`` con la ruta absoluta
  al JSON que solo exista en disco del host (o volumen), sin pasar por Git.

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


def load_users() -> dict[str, str]:
    """Devuelve mapa usuario -> contraseña (texto plano, editable en el archivo)."""
    if not AUTH_USERS_FILE.is_file():
        return {}
    try:
        with open(AUTH_USERS_FILE, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}
    return dict(data.get("users") or {})


def verify_credentials(username: str, password: str) -> bool:
    users = load_users()
    u = (username or "").strip()
    if not u or password is None:
        return False
    return users.get(u) == password


def whatsapp_new_user_url() -> str:
    msg = (
        "Buen día! Quiero generar mi usuario para el sistema Análisis Mis Comprobantes"
    )
    return f"https://wa.me/5493513132914?text={quote(msg)}"
