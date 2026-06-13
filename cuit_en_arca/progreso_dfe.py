"""Estado de jobs de descarga del Domicilio Fiscal Electrónico (progreso en pantalla)."""

from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field
from typing import Any, Callable

_lock = threading.Lock()
_jobs: dict[str, dict[str, Any]] = {}

# Pasos que se muestran como checklist por cada CUIT.
PASOS_DFE: tuple[tuple[str, str], ...] = (
    ("login", "Iniciar sesión en ARCA"),
    ("ventanilla", "Abrir Domicilio Fiscal Electrónico"),
    ("listar", "Listar comunicaciones"),
    ("descargar", "Descargar / imprimir comunicaciones"),
)

_MAX_LOG = 400


@dataclass
class EstadoJobDfe:
    job_id: str
    total: int
    actual: int = 0
    mensaje: str = ""
    estado: str = "pendiente"  # pendiente | en_progreso | ok | error
    error: str | None = None
    carpeta: str | None = None
    total_archivos: int = 0
    cuits_ok: int = 0
    cuits_fallidos: int = 0
    log: list[str] = field(default_factory=list)
    resumen: list[dict[str, Any]] = field(default_factory=list)
    pasos: list[dict[str, str]] = field(default_factory=list)
    archivos: list[dict[str, str]] = field(default_factory=list)

    def a_dict(self) -> dict[str, Any]:
        pct = 0
        if self.total > 0:
            pct = min(100, int(round(100 * self.actual / self.total)))
        return {
            "job_id": self.job_id,
            "total": self.total,
            "actual": self.actual,
            "mensaje": self.mensaje,
            "estado": self.estado,
            "error": self.error,
            "carpeta": self.carpeta,
            "total_archivos": self.total_archivos,
            "cuits_ok": self.cuits_ok,
            "cuits_fallidos": self.cuits_fallidos,
            "porcentaje": pct,
            "log": list(self.log),
            "resumen": list(self.resumen),
            "pasos": list(self.pasos),
            "archivos": list(self.archivos),
        }


def crear_job_dfe(job_id: str, total: int) -> None:
    with _lock:
        _jobs[job_id] = {
            "estado": EstadoJobDfe(job_id=job_id, total=total, mensaje="Iniciando…"),
        }


def reiniciar_pasos_dfe(job_id: str) -> None:
    with _lock:
        item = _jobs.get(job_id)
        if not item:
            return
        st: EstadoJobDfe = item["estado"]
        st.pasos = [
            {"clave": clave, "etiqueta": etiqueta, "estado": "pendiente"}
            for clave, etiqueta in PASOS_DFE
        ]


def callback_paso_dfe(job_id: str) -> Callable[[str, str], None]:
    def _cb(clave: str, estado: str) -> None:
        with _lock:
            item = _jobs.get(job_id)
            if not item:
                return
            st: EstadoJobDfe = item["estado"]
            for paso in st.pasos:
                if paso["clave"] == clave:
                    paso["estado"] = estado
                    break

    return _cb


def callback_log_dfe(job_id: str) -> Callable[[str], None]:
    def _cb(texto: str) -> None:
        with _lock:
            item = _jobs.get(job_id)
            if not item:
                return
            st: EstadoJobDfe = item["estado"]
            ts = time.strftime("%H:%M:%S")
            st.log.append(f"[{ts}] {texto}")
            if len(st.log) > _MAX_LOG:
                st.log = st.log[-_MAX_LOG:]
            st.mensaje = texto

    return _cb


def progreso_cuit_dfe(job_id: str, actual: int, total: int, mensaje: str = "") -> None:
    with _lock:
        item = _jobs.get(job_id)
        if not item:
            return
        st: EstadoJobDfe = item["estado"]
        st.actual = actual
        st.total = total
        st.estado = "en_progreso"
        if mensaje:
            st.mensaje = mensaje


def agregar_resumen_cuit_dfe(
    job_id: str,
    *,
    cuit: str,
    razon_social: str | None,
    total_archivos: int,
    error: str | None,
) -> None:
    with _lock:
        item = _jobs.get(job_id)
        if not item:
            return
        st: EstadoJobDfe = item["estado"]
        st.resumen.append(
            {
                "cuit": cuit,
                "razon_social": razon_social or "",
                "total_archivos": total_archivos,
                "error": error,
            }
        )
        st.total_archivos += int(total_archivos or 0)
        if error:
            st.cuits_fallidos += 1
        else:
            st.cuits_ok += 1


def agregar_archivo_dfe(job_id: str, download_id: str, ruta: str, nombre: str) -> None:
    with _lock:
        item = _jobs.get(job_id)
        if not item:
            return
        st: EstadoJobDfe = item["estado"]
        st.archivos.append({"id": download_id, "ruta": ruta, "nombre": nombre})


def marcar_ok_dfe(job_id: str, *, carpeta: str) -> None:
    with _lock:
        item = _jobs.get(job_id)
        if not item:
            return
        st: EstadoJobDfe = item["estado"]
        st.estado = "ok"
        st.actual = st.total
        st.carpeta = carpeta
        st.mensaje = "Completado"


def marcar_error_dfe(job_id: str, error: str) -> None:
    with _lock:
        item = _jobs.get(job_id)
        if not item:
            return
        st: EstadoJobDfe = item["estado"]
        st.estado = "error"
        st.error = error


def marcar_cancelado_dfe(job_id: str, mensaje: str = "Descarga cancelada.") -> None:
    with _lock:
        item = _jobs.get(job_id)
        if not item:
            return
        st: EstadoJobDfe = item["estado"]
        st.estado = "cancelado"
        st.error = mensaje
        st.mensaje = mensaje


def obtener_job_dfe(job_id: str) -> dict[str, Any] | None:
    with _lock:
        item = _jobs.get(job_id)
        if not item:
            return None
        return item["estado"].a_dict()
