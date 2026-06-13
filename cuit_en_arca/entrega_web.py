"""Entrega incremental de archivos al navegador (servidor web / Render)."""

from __future__ import annotations

import mimetypes
import tempfile
from pathlib import Path
from typing import Callable

RegistrarArchivo = Callable[[str, bytes, str], None]

_descargas: dict | None = None


def init_descargas(store: dict) -> None:
    global _descargas
    _descargas = store


def make_registrar(agregar_estado) -> RegistrarArchivo:
    from uuid import uuid4

    def _registrar(rel: str, data: bytes, mime: str) -> None:
        did = uuid4().hex
        nombre = Path(rel).name
        if _descargas is not None:
            _descargas[did] = (data, nombre, mime)
        agregar_estado(did, rel, nombre)

    return _registrar


def carpeta_trabajo_web(job_id: str) -> Path:
    p = Path(tempfile.gettempdir()) / "aic_web_jobs" / job_id
    p.mkdir(parents=True, exist_ok=True)
    return p


def carpeta_ap_servidor() -> Path:
    p = Path(tempfile.gettempdir()) / "aic_ap_data" / "salida"
    p.mkdir(parents=True, exist_ok=True)
    return p


class EntregaWeb:
    """Escanea una carpeta en el servidor y registra archivos nuevos para descarga."""

    def __init__(self, base: Path, registrar: RegistrarArchivo) -> None:
        self.base = base
        self._registrar = registrar
        self._vistos: set[str] = set()

    def escanear(self) -> None:
        if not self.base.is_dir():
            return
        for path in sorted(self.base.rglob("*")):
            if not path.is_file():
                continue
            rel = path.relative_to(self.base).as_posix()
            if rel in self._vistos:
                continue
            try:
                data = path.read_bytes()
            except OSError:
                continue
            mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
            self._vistos.add(rel)
            self._registrar(rel, data, mime)
