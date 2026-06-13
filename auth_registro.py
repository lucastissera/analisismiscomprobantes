"""Alta de usuarios por enlace: CUIT como usuario, contraseña elegida por el cliente."""

from __future__ import annotations

import json
import logging
import os
import re
import secrets
import smtplib
import sys
import tempfile
import threading
from datetime import date, datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path
from typing import Any
from urllib.parse import quote

import bcrypt

from app_branding import APP_NAME

_LOG = logging.getLogger(__name__)
_lock = threading.Lock()

_CUIT_RE = re.compile(r"^\d{11}$")
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def dir_auth_servidor() -> Path:
    override = (os.environ.get("AUTH_REGISTRATIONS_DIR") or "").strip()
    if override:
        p = Path(override)
    elif (os.environ.get("AUTH_DATA_DIR") or "").strip():
        p = Path(os.environ["AUTH_DATA_DIR"].strip()) / "auth"
    elif getattr(sys, "frozen", False):
        from auth import _dir_datos_usuario

        p = _dir_datos_usuario()
    else:
        p = Path(tempfile.gettempdir()) / "aic_auth_data"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _path_solicitudes() -> Path:
    return dir_auth_servidor() / "solicitudes_pendientes.json"


def _path_usuarios_overlay() -> Path:
    return dir_auth_servidor() / "usuarios_registrados.json"


def _path_log_altas() -> Path:
    return dir_auth_servidor() / "altas_completadas.json"


def _token_horas() -> int:
    raw = (os.environ.get("AUTH_ALTA_TOKEN_HORAS") or "72").strip()
    try:
        return max(1, min(int(raw), 168))
    except ValueError:
        return 72


def _min_password_len() -> int:
    raw = (os.environ.get("AUTH_MIN_PASSWORD_LEN") or "8").strip()
    try:
        return max(6, min(int(raw), 128))
    except ValueError:
        return 8


def _dias_suscripcion() -> int:
    raw = (os.environ.get("AUTH_SUBSCRIPTION_DAYS") or "30").strip()
    try:
        return max(1, min(int(raw), 3660))
    except ValueError:
        return 30


def _parse_fecha_local(val: Any) -> date | None:
    from auth import _parse_fecha

    return _parse_fecha(val)


def normalizar_cuit(val: str) -> str | None:
    digits = re.sub(r"\D", "", (val or "").strip())
    if not _CUIT_RE.match(digits):
        return None
    return digits


def formatear_cuit(cuit: str) -> str:
    d = normalizar_cuit(cuit) or cuit
    if len(d) == 11:
        return f"{d[:2]}-{d[2:10]}-{d[10]}"
    return d


def _leer_json(path: Path, default: Any) -> Any:
    if not path.is_file():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        _LOG.warning("No se pudo leer %s: %s", path, exc)
        return default


def _escribir_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def hash_password(password: str) -> str:
    pwd = (password or "").encode("utf-8")
    return bcrypt.hashpw(pwd, bcrypt.gensalt(rounds=12)).decode("ascii")


def verificar_password(stored: str, password: str) -> bool:
    s = (stored or "").strip()
    pwd = (password or "").encode("utf-8")
    if s.startswith("$2"):
        try:
            return bcrypt.checkpw(pwd, s.encode("ascii"))
        except ValueError:
            return False
    return s == (password or "")


def cargar_usuarios_overlay() -> dict[str, dict[str, Any]]:
    data = _leer_json(_path_usuarios_overlay(), {"users": {}})
    users = data.get("users") if isinstance(data, dict) else {}
    return users if isinstance(users, dict) else {}


def _meta_overlay(cuit: str) -> dict[str, Any] | None:
    u = normalizar_cuit(cuit)
    if not u:
        return None
    meta = cargar_usuarios_overlay().get(u)
    return meta if isinstance(meta, dict) else None


def cuenta_pendiente_aprobacion(cuit: str) -> dict[str, Any] | None:
    meta = _meta_overlay(cuit)
    if not meta:
        return None
    if meta.get("pendiente_aprobacion") or meta.get("activo") is False:
        return meta
    return None


def verificar_acceso_overlay(cuit: str, password: str) -> str | None:
    """None = ok; 'pending_approval' | 'invalid'."""
    meta = cuenta_pendiente_aprobacion(cuit)
    if not meta:
        return None
    if verificar_password(str(meta.get("password") or ""), password):
        return "pending_approval"
    return "invalid"


def alta_publica_habilitada() -> bool:
    v = (os.environ.get("AUTH_ALTA_PUBLICA") or "1").strip().lower()
    return v in ("1", "true", "yes", "on")


def usuario_existe(cuit: str) -> bool:
    u = normalizar_cuit(cuit)
    if not u:
        return False
    overlay = cargar_usuarios_overlay()
    if u in overlay:
        return True
    from auth import _load_cuentas_sin_env_json, _usuarios_desde_env_json

    env = _usuarios_desde_env_json()
    base = env if env else _load_cuentas_sin_env_json()
    return u in base


def _cargar_solicitudes() -> dict[str, Any]:
    data = _leer_json(_path_solicitudes(), {"solicitudes": {}})
    if not isinstance(data, dict):
        return {"solicitudes": {}}
    if "solicitudes" not in data or not isinstance(data["solicitudes"], dict):
        data["solicitudes"] = {}
    return data


def crear_solicitud(
    *,
    cuit: str,
    email: str,
    nombre: str = "",
) -> tuple[str, dict[str, Any]]:
    u = normalizar_cuit(cuit)
    if not u:
        raise ValueError("cuit_invalido")
    em = (email or "").strip().lower()
    if not _EMAIL_RE.match(em):
        raise ValueError("email_invalido")
    if usuario_existe(u):
        raise ValueError("cuit_duplicado")

    token = secrets.token_urlsafe(32)
    ahora = datetime.now(timezone.utc)
    expira = ahora + timedelta(hours=_token_horas())
    registro = {
        "cuit": u,
        "email": em,
        "nombre": (nombre or "").strip(),
        "creado": ahora.isoformat(timespec="seconds"),
        "expira": expira.isoformat(timespec="seconds"),
        "usado": False,
    }

    with _lock:
        data = _cargar_solicitudes()
        # Una solicitud activa por CUIT
        for tok, sol in list(data["solicitudes"].items()):
            if not isinstance(sol, dict):
                continue
            if sol.get("cuit") == u and not sol.get("usado"):
                try:
                    exp = datetime.fromisoformat(str(sol["expira"]).replace("Z", "+00:00"))
                    if exp > ahora:
                        del data["solicitudes"][tok]
                except ValueError:
                    del data["solicitudes"][tok]
        data["solicitudes"][token] = registro
        _escribir_json(_path_solicitudes(), data)

    return token, registro


def obtener_solicitud(token: str) -> dict[str, Any] | None:
    tok = (token or "").strip()
    if not tok:
        return None
    data = _cargar_solicitudes()
    sol = data.get("solicitudes", {}).get(tok)
    if not isinstance(sol, dict) or sol.get("usado"):
        return None
    try:
        exp = datetime.fromisoformat(str(sol["expira"]).replace("Z", "+00:00"))
    except ValueError:
        return None
    if exp <= datetime.now(timezone.utc):
        return None
    return sol


def activar_cuenta(token: str, password: str) -> dict[str, Any]:
    tok = (token or "").strip()
    pwd = password or ""
    if len(pwd) < _min_password_len():
        raise ValueError("password_corta")

    with _lock:
        sol = obtener_solicitud(tok)
        if not sol:
            raise ValueError("token_invalido")
        cuit = str(sol["cuit"])
        if usuario_existe(cuit):
            raise ValueError("cuit_duplicado")

        overlay_path = _path_usuarios_overlay()
        overlay = _leer_json(overlay_path, {"version": 1, "users": {}})
        if not isinstance(overlay.get("users"), dict):
            overlay["users"] = {}
        overlay["users"][cuit] = {
            "password": hash_password(pwd),
            "email": sol.get("email"),
            "nombre": sol.get("nombre") or "",
            "valido_desde": datetime.now(timezone.utc).date().isoformat(),
            "activo": False,
            "pendiente_aprobacion": True,
            "password_definida": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        }
        overlay["updated_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        _escribir_json(overlay_path, overlay)

        data = _cargar_solicitudes()
        if tok in data.get("solicitudes", {}):
            data["solicitudes"][tok]["usado"] = True
            data["solicitudes"][tok]["activado"] = datetime.now(timezone.utc).isoformat(
                timespec="seconds"
            )
            _escribir_json(_path_solicitudes(), data)

    registro_alta = {
        "cuit": cuit,
        "email": sol.get("email"),
        "nombre": sol.get("nombre") or "",
        "activado": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "estado": "pendiente_aprobacion",
    }
    _registrar_alta_log(registro_alta)
    return registro_alta


def listar_pendientes_aprobacion() -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for cuit, meta in cargar_usuarios_overlay().items():
        if not isinstance(meta, dict):
            continue
        if not (meta.get("pendiente_aprobacion") or meta.get("activo") is False):
            continue
        out.append(
            {
                "cuit": cuit,
                "cuit_fmt": formatear_cuit(cuit),
                "email": meta.get("email") or "",
                "nombre": meta.get("nombre") or "",
                "password_definida": meta.get("password_definida") or "",
            }
        )
    out.sort(key=lambda x: x.get("password_definida") or "", reverse=True)
    return out


def aprobar_cuenta(cuit: str) -> bool:
    u = normalizar_cuit(cuit)
    if not u:
        return False
    dias = _dias_suscripcion()
    hoy = datetime.now(timezone.utc).date()
    valido_hasta = hoy + timedelta(days=dias)
    with _lock:
        path = _path_usuarios_overlay()
        overlay = _leer_json(path, {"version": 1, "users": {}})
        users = overlay.get("users")
        if not isinstance(users, dict) or u not in users:
            return False
        users[u]["activo"] = True
        users[u]["pendiente_aprobacion"] = False
        users[u]["valido_desde"] = hoy.isoformat()
        users[u]["valido_hasta"] = valido_hasta.isoformat()
        users[u]["aprobado_en"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        overlay["updated_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        _escribir_json(path, overlay)
    return True


def listar_usuarios_suscripcion() -> list[dict[str, Any]]:
    hoy = date.today()
    out: list[dict[str, Any]] = []
    for cuit, meta in cargar_usuarios_overlay().items():
        if not isinstance(meta, dict):
            continue
        if meta.get("pendiente_aprobacion") or meta.get("activo") is False:
            continue
        vh = _parse_fecha_local(meta.get("valido_hasta"))
        dias = (vh - hoy).days if vh else None
        out.append(
            {
                "cuit": cuit,
                "cuit_fmt": formatear_cuit(cuit),
                "email": meta.get("email") or "",
                "nombre": meta.get("nombre") or "",
                "valido_hasta": vh.isoformat() if vh else "",
                "valido_hasta_fmt": vh.strftime("%d/%m/%Y") if vh else "—",
                "dias_restantes": dias,
                "vencida": dias is not None and dias < 0,
            }
        )
    out.sort(key=lambda x: (x.get("dias_restantes") is None, x.get("dias_restantes") or 0))
    return out


def renovar_suscripcion(cuit: str, dias: int | None = None) -> bool:
    u = normalizar_cuit(cuit)
    if not u:
        return False
    duracion = dias if dias is not None else _dias_suscripcion()
    if duracion < 1:
        return False
    hoy = date.today()
    with _lock:
        path = _path_usuarios_overlay()
        overlay = _leer_json(path, {"version": 1, "users": {}})
        users = overlay.get("users")
        if not isinstance(users, dict) or u not in users:
            return False
        meta = users[u]
        if meta.get("pendiente_aprobacion") or meta.get("activo") is False:
            return False
        vh_actual = _parse_fecha_local(meta.get("valido_hasta"))
        base = max(hoy, vh_actual) if vh_actual else hoy
        nueva_hasta = base + timedelta(days=duracion)
        if not meta.get("valido_desde") or (vh_actual and hoy > vh_actual):
            meta["valido_desde"] = hoy.isoformat()
        meta["valido_hasta"] = nueva_hasta.isoformat()
        meta["renovado_en"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        overlay["updated_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        _escribir_json(path, overlay)
    return True


def info_suscripcion_usuario(username: str) -> dict[str, Any] | None:
    from auth import _load_cuentas, es_administrador

    u_raw = (username or "").strip()
    if not u_raw or es_administrador(u_raw):
        return None
    u = normalizar_cuit(u_raw) or u_raw
    cuenta = _load_cuentas().get(u) or _load_cuentas().get(u_raw)
    if not cuenta or not cuenta.valido_hasta:
        return None
    hoy = date.today()
    dias = (cuenta.valido_hasta - hoy).days
    return {
        "valido_hasta": cuenta.valido_hasta,
        "valido_hasta_fmt": cuenta.valido_hasta.strftime("%d/%m/%Y"),
        "dias_restantes": dias,
    }


def rechazar_cuenta(cuit: str) -> bool:
    u = normalizar_cuit(cuit)
    if not u:
        return False
    with _lock:
        path = _path_usuarios_overlay()
        overlay = _leer_json(path, {"version": 1, "users": {}})
        users = overlay.get("users")
        if not isinstance(users, dict) or u not in users:
            return False
        meta = users[u]
        if not (meta.get("pendiente_aprobacion") or meta.get("activo") is False):
            return False
        del users[u]
        overlay["updated_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        _escribir_json(path, overlay)
    return True


def _registrar_alta_log(entry: dict[str, Any]) -> None:
    path = _path_log_altas()
    data = _leer_json(path, {"altas": []})
    if not isinstance(data.get("altas"), list):
        data["altas"] = []
    data["altas"].insert(0, entry)
    data["altas"] = data["altas"][:200]
    _escribir_json(path, data)


def listar_altas_recientes(limit: int = 30) -> list[dict[str, Any]]:
    data = _leer_json(_path_log_altas(), {"altas": []})
    altas = data.get("altas") if isinstance(data, dict) else []
    if not isinstance(altas, list):
        return []
    return [a for a in altas[:limit] if isinstance(a, dict)]


def whatsapp_alta_admin_url(cuit: str, email: str, nombre: str = "") -> str:
    tel = (os.environ.get("AUTH_ADMIN_WHATSAPP") or "5493513132914").strip()
    cuit_fmt = formatear_cuit(cuit)
    nom = f" ({nombre})" if nombre else ""
    msg = (
        f"Solicitud de alta en {APP_NAME}: CUIT {cuit_fmt}{nom}, "
        f"email {email}. El usuario ya definió contraseña. "
        f"Pendiente de tu aprobación (¿pagó?)."
    )
    return f"https://wa.me/{tel}?text={quote(msg)}"


def _enviar_email(destino: str, asunto: str, cuerpo: str) -> bool:
    host = (os.environ.get("SMTP_HOST") or "").strip()
    user = (os.environ.get("SMTP_USER") or "").strip()
    password = (os.environ.get("SMTP_PASSWORD") or "").strip()
    port_raw = (os.environ.get("SMTP_PORT") or "587").strip()
    if not host or not destino:
        return False
    try:
        port = int(port_raw)
    except ValueError:
        port = 587
    remitente = (os.environ.get("SMTP_FROM") or user or f"noreply@{host}").strip()
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = remitente
    msg["To"] = destino
    msg.set_content(cuerpo)
    try:
        with smtplib.SMTP(host, port, timeout=20) as smtp:
            if port != 25:
                smtp.starttls()
            if user and password:
                smtp.login(user, password)
            smtp.send_message(msg)
        return True
    except OSError as exc:
        _LOG.warning("No se pudo enviar email a %s: %s", destino, exc)
        return False


def notificar_admin_alta(cuit: str, email: str, nombre: str = "") -> dict[str, Any]:
    cuit_fmt = formatear_cuit(cuit)
    nom_line = f"Nombre: {nombre}\n" if nombre else ""
    cuerpo = (
        f"Nueva solicitud de alta en {APP_NAME}.\n\n"
        f"CUIT (usuario): {cuit_fmt}\n"
        f"{nom_line}"
        f"Email de contacto: {email}\n\n"
        f"El usuario ya eligió contraseña por enlace.\n"
        f"La cuenta queda PENDIENTE hasta que la apruebes en el panel "
        f"«Altas de usuarios» (después de confirmar el pago).\n"
    )
    admin_mail = (os.environ.get("AUTH_ADMIN_NOTIFY_EMAIL") or "").strip()
    email_ok = False
    if admin_mail:
        email_ok = _enviar_email(
            admin_mail,
            f"[{APP_NAME}] Alta de usuario {cuit_fmt}",
            cuerpo,
        )
    return {
        "email_enviado": email_ok,
        "whatsapp_url": whatsapp_alta_admin_url(cuit, email, nombre),
    }
