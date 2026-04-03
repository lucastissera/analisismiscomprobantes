"""
CUIT en ARCA — etapa previa opcional al procesamiento con sumar_imp_total.

- ``credenciales``: lectura del .xlsx de CUIT / clave / representado.
- ``validacion``: rango de fechas <= 1 año.
- ``service``: orquestación + flag ``CUIT_EN_ARCA_PLAYWRIGHT``.
- ``automation_playwright``: navegador (selectores a mantener ante cambios AFIP).
"""

from cuit_en_arca.errores import (
    ArcaProcesoError,
    AutomatizacionArcaError,
    AutomatizacionNoDisponibleError,
    CredencialesArchivoError,
    CuitRepresentadoNoEncontradoError,
    FechaRangoInvalidaError,
)
from cuit_en_arca.service import automatizacion_cuit_arca_habilitada, ejecutar_flujo_cuit_en_arca

__all__ = [
    "ArcaProcesoError",
    "AutomatizacionArcaError",
    "AutomatizacionNoDisponibleError",
    "CredencialesArchivoError",
    "CuitRepresentadoNoEncontradoError",
    "FechaRangoInvalidaError",
    "automatizacion_cuit_arca_habilitada",
    "ejecutar_flujo_cuit_en_arca",
]
