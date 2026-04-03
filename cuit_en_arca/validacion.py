"""Validación de rango de fechas (dd/mm/yyyy)."""

from __future__ import annotations

import re
from datetime import date, timedelta

from cuit_en_arca.errores import FechaRangoInvalidaError

_PATRON_DM = re.compile(
    r"^\s*(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{4})\s*$"
)


def parsear_fecha_argentina(texto: str) -> date:
    """Parsea texto tipo d/m/aaaa o dd-mm-yyyy."""
    s = (texto or "").strip()
    m = _PATRON_DM.match(s)
    if not m:
        raise FechaRangoInvalidaError(
            "Las fechas deben tener formato dd/mm/yyyy (ej. 01/03/2025)."
        )
    d, mes, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return date(y, mes, d)
    except ValueError as exc:
        raise FechaRangoInvalidaError("Fecha inválida.") from exc


def parsear_rango_fechas_texto(texto: str | None) -> tuple[str, str] | None:
    """
    Interpreta un texto tipo ``01/01/2025 - 31/12/2025`` o ``... al ...`` / ``... hasta ...``.
    Devuelve (desde, hasta) como strings para pasar a ``parsear_fecha_argentina``.
    """
    if texto is None:
        return None
    s = str(texto).strip()
    if not s:
        return None
    patrones = (
        r"^\s*(.+?)\s+[-–—]\s+(.+?)\s*$",
        r"^\s*(.+?)\s+al\s+(.+?)\s*$",
        r"^\s*(.+?)\s+hasta\s+(.+?)\s*$",
    )
    for pat in patrones:
        m = re.match(pat, s, re.I)
        if m:
            a, b = m.group(1).strip(), m.group(2).strip()
            if a and b:
                return a, b
    return None


def validar_rango_max_un_anio(desde: date, hasta: date) -> None:
    """
    V1 del diagrama: el rango debe ser <= 1 año.
    Se interpreta como: (hasta - desde) <= 365 días (inclusive de extremos).
    """
    if hasta < desde:
        raise FechaRangoInvalidaError("La fecha hasta debe ser posterior o igual a la fecha desde.")
    if (hasta - desde) > timedelta(days=365):
        raise FechaRangoInvalidaError("Corroborar rango de fechas: el período no puede superar un año.")
