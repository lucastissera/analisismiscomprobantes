"""Errores de negocio alineados al diagrama de proceso CUIT en ARCA."""


class ArcaProcesoError(Exception):
    """Base para fallos controlados del flujo (mensaje listo para mostrar al usuario)."""


class FechaRangoInvalidaError(ArcaProcesoError):
    """V1 → No: rango de fechas mayor a un año."""


class CredencialesArchivoError(ArcaProcesoError):
    """Archivo .xlsx de credenciales ilegible o incompleto."""


class CuitRepresentadoNoEncontradoError(ArcaProcesoError):
    """I → No: el CUIT representado no aparece en la lista de perfiles."""


class AutomatizacionArcaError(ArcaProcesoError):
    """Fallo en navegación, selectores o descarga (detalle en args)."""


class AutomatizacionNoDisponibleError(ArcaProcesoError):
    """Playwright/Chromium no instalado o deshabilitado en este entorno."""
