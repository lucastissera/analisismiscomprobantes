"""Marcas de tiempo para logs visibles al usuario (Argentina)."""

from __future__ import annotations

from datetime import datetime
from zoneinfo import ZoneInfo

_TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")


def ahora_ar() -> datetime:
    return datetime.now(_TZ_AR)


def ahora_ar_naive() -> datetime:
    """Datetime sin tz en hora Argentina (comparar con la hora que elige el usuario)."""
    return datetime.now(_TZ_AR).replace(tzinfo=None)


def fecha_hora_ar_texto(momento: datetime | None = None) -> str:
    """Texto legible con zona explícita."""
    m = momento or ahora_ar()
    return m.strftime("%Y-%m-%d %H:%M:%S") + " (Argentina)"


def hora_log_ar() -> str:
    """``HH:MM:SS`` en hora de Argentina (independiente del TZ del servidor)."""
    return ahora_ar().strftime("%H:%M:%S")
