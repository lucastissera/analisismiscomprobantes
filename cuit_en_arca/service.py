"""
Orquestación de la etapa «CUIT en ARCA» (antes de sumar_imp_total).

No importa ni modifica sumar_imp_total: el flujo termina en un archivo descargable
que el usuario puede volver a subir en «Procesador de Comprobantes».
"""

from __future__ import annotations

import io
import os

from cuit_en_arca.credenciales import leer_credenciales_xlsx
from cuit_en_arca.errores import AutomatizacionNoDisponibleError, FechaRangoInvalidaError
from cuit_en_arca.validacion import parsear_fecha_argentina, validar_rango_max_un_anio


def automatizacion_cuit_arca_habilitada() -> bool:
    """Variable de entorno explícita (evita lanzar Chromium en Render sin deps)."""
    v = os.environ.get("CUIT_EN_ARCA_PLAYWRIGHT", "").strip().lower()
    return v in ("1", "true", "yes", "on")


def ejecutar_flujo_cuit_en_arca(
    archivo_credenciales: io.BytesIO,
    fecha_desde_texto: str | None = None,
    fecha_hasta_texto: str | None = None,
) -> tuple[bytes, str]:
    """
    Valida entradas y, si está habilitado, ejecuta Playwright.
    Fechas: si el Excel trae columna D (Rango Fechas) parseable, tiene prioridad;
    si no, se usan las del formulario.
    Retorna (bytes del archivo descargado, nombre sugerido).
    """
    cred = leer_credenciales_xlsx(archivo_credenciales)
    fd = (fecha_desde_texto or "").strip() or None
    fh = (fecha_hasta_texto or "").strip() or None
    if cred.rango_fecha_desde and cred.rango_fecha_hasta:
        fd, fh = cred.rango_fecha_desde, cred.rango_fecha_hasta
    if not fd or not fh:
        raise FechaRangoInvalidaError(
            "Indicá el rango en el formulario (dd/mm/yyyy) o en la columna D del Excel "
            "(Rango Fechas), por ejemplo: 01/01/2025 - 31/12/2025."
        )
    desde = parsear_fecha_argentina(fd)
    hasta = parsear_fecha_argentina(fh)
    validar_rango_max_un_anio(desde, hasta)

    if not automatizacion_cuit_arca_habilitada():
        raise AutomatizacionNoDisponibleError(
            "La descarga automática desde AFIP está deshabilitada en este servidor. "
            "Para usarla en local: definí la variable de entorno CUIT_EN_ARCA_PLAYWRIGHT=1, "
            "instalá dependencias (pip install playwright) y ejecutá: playwright install chromium"
        )

    headless = os.environ.get("CUIT_EN_ARCA_HEADLESS", "1").strip().lower() not in (
        "0",
        "false",
        "no",
    )
    from cuit_en_arca.automation_playwright import ejecutar_descarga_mis_comprobantes

    return ejecutar_descarga_mis_comprobantes(
        cred, desde, hasta, headless=headless
    )
