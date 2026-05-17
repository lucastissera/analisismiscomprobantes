"""Usuarios y contraseñas desde JSON (texto plano).

Archivo principal: **auth_users.json** (misma carpeta que este módulo). Podés editarlo,
commitearlo y subirlo al servidor; el login usa siempre ese archivo salvo override.

Opcional: variable ``AUTH_USERS_PATH`` con otra ruta absoluta al JSON.

Si el archivo no existe o ``users`` está vacío: ``AUTH_ADMIN_USER`` y ``AUTH_ADMIN_PASSWORD``
en el entorno (p. ej. solo en el servidor sin archivo).
"""

from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path
from urllib.parse import quote

_AUTH_DIR = Path(__file__).resolve().parent
_LOG = logging.getLogger(__name__)


def _auth_users_file() -> Path:
    """Ruta al JSON; se evalúa en cada lectura para respetar env y .env cargado antes del import."""
    override = (os.environ.get("AUTH_USERS_PATH") or "").strip()
    if override:
        return Path(override)
    # Portable PyInstaller: el usuario puede copiar auth_users.json junto al .exe (prioridad).
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        portable = exe_dir / "auth_users.json"
        if portable.is_file():
            return portable
        meip = getattr(sys, "_MEIPASS", None)
        if meip:
            for name in ("auth_users.json", "auth_users.example.json"):
                bundled = Path(meip) / name
                if bundled.is_file():
                    return bundled
    return _AUTH_DIR / "auth_users.json"


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
    path = _auth_users_file()
    if path.is_file():
        try:
            with open(path, encoding="utf-8-sig") as f:
                data = json.load(f)
            got = _normalizar_usuarios(data.get("users") or {})
            if got:
                return got
        except json.JSONDecodeError as exc:
            _LOG.warning("auth_users.json inválido en %s: %s", path, exc)
        except OSError as exc:
            _LOG.warning("No se pudo leer %s: %s", path, exc)
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
