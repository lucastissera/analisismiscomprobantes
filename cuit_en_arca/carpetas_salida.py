"""Nombres de carpetas de salida con fecha y hora (evita colisiones el mismo día)."""

from __future__ import annotations

from datetime import date, datetime

from cuit_en_arca.hora_log import ahora_ar


def momento_carpeta_ar(hoy: date | None = None) -> datetime:
    """Datetime naive en hora Argentina para nombres de carpeta con timestamp."""
    ar = ahora_ar().replace(tzinfo=None)
    if hoy is None:
        return ar
    if isinstance(hoy, datetime):
        return hoy.replace(tzinfo=None) if hoy.tzinfo else hoy
    return datetime.combine(hoy, ar.time())


def stamp_carpeta_ejecucion(momento: datetime | None = None) -> str:
    """``yyyy-mm-dd HH-MM`` (sin ``:`` por compatibilidad en Windows)."""
    m = momento or ahora_ar()
    if m.tzinfo is not None:
        m = m.replace(tzinfo=None)
    return m.strftime("%Y-%m-%d %H-%M")
