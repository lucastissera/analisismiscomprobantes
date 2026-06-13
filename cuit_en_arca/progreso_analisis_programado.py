"""Estado de la ejecución del Análisis Programado (progreso en pantalla)."""

from __future__ import annotations

import threading
from dataclasses import dataclass, field
from typing import Any, Callable

from cuit_en_arca.hora_log import hora_log_ar

_lock = threading.Lock()
_estado: EstadoEjecucionAP | None = None

_ORDEN_SISTEMAS = ("mis_comprobantes", "dfe", "nuestra_parte")
_ETIQUETAS_SISTEMA = {
    "mis_comprobantes": "Mis Comprobantes",
    "dfe": "Domicilio Fiscal Electrónico",
    "nuestra_parte": "Nuestra Parte",
}

_MAX_LOG = 400


@dataclass
class EstadoEjecucionAP:
    estado: str = "idle"  # idle | en_progreso | ok | error
    mensaje: str = ""
    error: str | None = None
    carpeta: str | None = None
    actual: int = 0
    total: int = 0
    log: list[str] = field(default_factory=list)
    pasos: list[dict[str, str]] = field(default_factory=list)
    fallos: list[str] = field(default_factory=list)
    archivos: list[dict[str, str]] = field(default_factory=list)

    def a_dict(self) -> dict[str, Any]:
        pct = 0
        if self.total > 0:
            pct = min(100, int(round(100 * self.actual / self.total)))
        elif self.estado == "ok":
            pct = 100
        return {
            "estado": self.estado,
            "mensaje": self.mensaje,
            "error": self.error,
            "carpeta": self.carpeta,
            "actual": self.actual,
            "total": self.total,
            "porcentaje": pct,
            "log": list(self.log),
            "pasos": list(self.pasos),
            "fallos": list(self.fallos[:50]),
            "archivos": list(self.archivos),
        }


def _pasos_para_sistemas(sistemas: list[str]) -> list[tuple[str, str]]:
    pasos: list[tuple[str, str]] = [("preparacion", "Preparar carpeta y datos")]
    for clave in _ORDEN_SISTEMAS:
        if clave in sistemas:
            pasos.append((clave, _ETIQUETAS_SISTEMA[clave]))
    pasos.append(("finalizado", "Finalizar"))
    return pasos


def iniciar_ejecucion_ap(sistemas: list[str]) -> None:
    global _estado
    with _lock:
        pasos = _pasos_para_sistemas(sistemas)
        _estado = EstadoEjecucionAP(
            estado="en_progreso",
            mensaje="Iniciando análisis programado…",
            total=len(pasos),
            pasos=[
                {"clave": clave, "etiqueta": etiqueta, "estado": "pendiente"}
                for clave, etiqueta in pasos
            ],
        )


def callback_log_ap() -> Callable[[str], None]:
    def _cb(texto: str) -> None:
        with _lock:
            if _estado is None:
                return
            ts = hora_log_ar()
            _estado.log.append(f"[{ts}] {texto}")
            if len(_estado.log) > _MAX_LOG:
                _estado.log = _estado.log[-_MAX_LOG:]
            _estado.mensaje = texto

    return _cb


def marcar_paso_ap(clave: str, estado: str) -> None:
    with _lock:
        if _estado is None:
            return
        for paso in _estado.pasos:
            if paso["clave"] == clave:
                paso["estado"] = estado
                break
        ok = sum(1 for p in _estado.pasos if p["estado"] == "ok")
        _estado.actual = ok


def agregar_archivo_ap(download_id: str, ruta: str, nombre: str) -> None:
    with _lock:
        if _estado is None:
            return
        _estado.archivos.append({"id": download_id, "ruta": ruta, "nombre": nombre})


def marcar_ok_ap(*, carpeta: str, mensaje: str, fallos: list[str] | None = None) -> None:
    global _estado
    with _lock:
        if _estado is None:
            return
        for paso in _estado.pasos:
            if paso["estado"] == "pendiente":
                paso["estado"] = "ok" if paso["clave"] == "finalizado" else paso["estado"]
        if _estado.pasos:
            _estado.pasos[-1]["estado"] = "ok"
        _estado.estado = "ok"
        _estado.carpeta = carpeta
        _estado.mensaje = mensaje
        _estado.fallos = list(fallos or [])
        _estado.actual = _estado.total


def marcar_error_ap(error: str) -> None:
    global _estado
    with _lock:
        if _estado is None:
            _estado = EstadoEjecucionAP(estado="error", error=error, mensaje=error)
            return
        for paso in _estado.pasos:
            if paso["estado"] in ("pendiente", "en_curso"):
                paso["estado"] = "error"
                break
        _estado.estado = "error"
        _estado.error = error
        _estado.mensaje = error


def marcar_cancelado_ap(mensaje: str = "Descarga cancelada.") -> None:
    global _estado
    with _lock:
        if _estado is None:
            _estado = EstadoEjecucionAP(estado="cancelado", error=mensaje, mensaje=mensaje)
            return
        for paso in _estado.pasos:
            if paso["estado"] in ("pendiente", "en_curso"):
                paso["estado"] = "error"
                break
        _estado.estado = "cancelado"
        _estado.error = mensaje
        _estado.mensaje = mensaje


def resetear_ejecucion_ap() -> None:
    global _estado
    with _lock:
        _estado = None


def obtener_ejecucion_ap() -> dict[str, Any]:
    with _lock:
        if _estado is None:
            return {"estado": "idle", "mensaje": "", "porcentaje": 0, "log": [], "pasos": []}
        return _estado.a_dict()


def ejecutando_ap() -> bool:
    with _lock:
        return _estado is not None and _estado.estado == "en_progreso"
