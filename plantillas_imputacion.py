"""
Plantillas de imputación contable guardadas en disco (modo escritorio / .exe).

Datos bajo %LOCALAPPDATA%\\DepuracionExcelComprobantes\\plantillas_imputacion\\
o, en desarrollo, carpeta ``data_local_imputaciones`` en el proyecto si se define
``ENABLE_LOCAL_PLANTILLAS_IMPUTACION=1``.
"""

from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4


def plantillas_imputacion_disponibles() -> bool:
    """True en .exe congelado o si se fuerza con variable de entorno (pruebas locales)."""
    if getattr(sys, "frozen", False):
        return True
    v = (os.environ.get("ENABLE_LOCAL_PLANTILLAS_IMPUTACION") or "").strip().lower()
    return v in ("1", "true", "yes", "on")


def _dir_base_usuario() -> Path:
    if getattr(sys, "frozen", False):
        local = os.environ.get("LOCALAPPDATA")
        if local:
            return Path(local) / "DepuracionExcelComprobantes"
        return Path.home() / "AppData" / "Local" / "DepuracionExcelComprobantes"
    # Desarrollo: junto al proyecto o override
    override = (os.environ.get("IMPUTACIONES_DATA_DIR") or "").strip()
    if override:
        return Path(override)
    return Path(__file__).resolve().parent / "data_local_imputaciones"


def directorio_plantillas() -> Path:
    d = _dir_base_usuario() / "plantillas_imputacion"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _ruta_indice() -> Path:
    return directorio_plantillas() / "plantillas.json"


def _cargar_indice() -> dict[str, Any]:
    p = _ruta_indice()
    if not p.is_file():
        return {"version": 1, "plantillas": []}
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {"version": 1, "plantillas": []}
        pl = data.get("plantillas")
        if not isinstance(pl, list):
            data["plantillas"] = []
        return data
    except (json.JSONDecodeError, OSError):
        return {"version": 1, "plantillas": []}


def _guardar_indice(data: dict[str, Any]) -> None:
    p = _ruta_indice()
    tmp = p.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    tmp.replace(p)


def _normalizar_nombre_mostrar(nombre: str) -> str:
    s = (nombre or "").strip()
    return s[:120] if s else ""


def _nombre_archivo_seguro(stem: str, ext: str) -> str:
    ext_l = (ext or "").lower()
    if ext_l not in (".xlsx", ".csv"):
        ext_l = ".xlsx"
    safe = re.sub(r"[^\w.\-]", "_", stem)[:80] or "plantilla"
    return f"{safe}{ext_l}"


def listar_plantillas() -> list[dict[str, Any]]:
    data = _cargar_indice()
    out = [dict(x) for x in data.get("plantillas", []) if isinstance(x, dict)]
    out.sort(key=lambda r: (_normalizar_nombre_mostrar(str(r.get("nombre", ""))).lower(), str(r.get("id", ""))))
    return out


def obtener_plantilla(plantilla_id: str) -> dict[str, Any] | None:
    pid = (plantilla_id or "").strip()
    if not pid or not re.fullmatch(r"[a-f0-9]{32}", pid, re.I):
        return None
    for p in _cargar_indice().get("plantillas", []):
        if isinstance(p, dict) and str(p.get("id", "")) == pid:
            return dict(p)
    return None


def leer_bytes_plantilla(plantilla_id: str) -> tuple[bytes, str]:
    """Devuelve (contenido, nombre sugerido para leer_mapa_imputaciones)."""
    meta = obtener_plantilla(plantilla_id)
    if not meta:
        raise FileNotFoundError("plantilla")
    rel = str(meta.get("archivo") or "").strip()
    if not rel or ".." in rel or "/" in rel or "\\" in rel:
        raise FileNotFoundError("plantilla_archivo")
    path = directorio_plantillas() / rel
    if not path.is_file():
        raise FileNotFoundError("plantilla_archivo")
    data = path.read_bytes()
    nombre = str(meta.get("nombre_original") or rel)
    return data, nombre


def agregar_plantilla(nombre: str, contenido: bytes, nombre_archivo_original: str) -> dict[str, Any]:
    nombre_limpio = _normalizar_nombre_mostrar(nombre)
    if not nombre_limpio:
        raise ValueError("nombre_vacio")
    data = _cargar_indice()
    lista: list[dict[str, Any]] = list(data.get("plantillas", []))
    nl = nombre_limpio.lower()
    for p in lista:
        if str(p.get("nombre", "")).strip().lower() == nl:
            raise ValueError("nombre_duplicado")

    pid = uuid4().hex
    ext = Path(nombre_archivo_original or "").suffix.lower()
    if ext not in (".xlsx", ".csv"):
        ext = ".xlsx"
    fname = f"{pid}{ext}"
    path = directorio_plantillas() / fname
    path.write_bytes(contenido)

    rec = {
        "id": pid,
        "nombre": nombre_limpio,
        "archivo": fname,
        "nombre_original": Path(nombre_archivo_original or fname).name,
        "creado_iso": datetime.now(timezone.utc).isoformat(),
    }
    lista.append(rec)
    data["plantillas"] = lista
    _guardar_indice(data)
    return rec


def renombrar_plantilla(plantilla_id: str, nuevo_nombre: str) -> dict[str, Any]:
    pid = (plantilla_id or "").strip()
    nuevo = _normalizar_nombre_mostrar(nuevo_nombre)
    if not nuevo:
        raise ValueError("nombre_vacio")
    data = _cargar_indice()
    lista: list[dict[str, Any]] = list(data.get("plantillas", []))
    nlow = nuevo.lower()
    found = -1
    for i, p in enumerate(lista):
        if not isinstance(p, dict):
            continue
        if str(p.get("nombre", "")).strip().lower() == nlow and str(p.get("id")) != pid:
            raise ValueError("nombre_duplicado")
        if str(p.get("id", "")) == pid:
            found = i
    if found < 0:
        raise ValueError("no_existe")
    lista[found]["nombre"] = nuevo
    data["plantillas"] = lista
    _guardar_indice(data)
    return lista[found]


def reemplazar_archivo_plantilla(plantilla_id: str, contenido: bytes, nombre_archivo_original: str) -> dict[str, Any]:
    meta = obtener_plantilla(plantilla_id)
    if not meta:
        raise ValueError("no_existe")
    ext = Path(nombre_archivo_original or "").suffix.lower()
    if ext not in (".xlsx", ".csv"):
        ext = ".xlsx"
    fname_old = str(meta.get("archivo") or "")
    pid = str(meta.get("id"))
    fname_new = f"{pid}{ext}"
    root = directorio_plantillas()
    (root / fname_new).write_bytes(contenido)
    if fname_old and fname_old != fname_new:
        try:
            (root / fname_old).unlink(missing_ok=True)
        except OSError:
            pass

    data = _cargar_indice()
    lista: list[dict[str, Any]] = list(data.get("plantillas", []))
    for p in lista:
        if isinstance(p, dict) and str(p.get("id")) == pid:
            p["archivo"] = fname_new
            p["nombre_original"] = Path(nombre_archivo_original or fname_new).name
            p["actualizado_iso"] = datetime.now(timezone.utc).isoformat()
            break
    data["plantillas"] = lista
    _guardar_indice(data)
    out = obtener_plantilla(pid)
    if not out:
        raise ValueError("no_existe")
    return out


def eliminar_plantilla(plantilla_id: str) -> None:
    meta = obtener_plantilla(plantilla_id)
    if not meta:
        raise ValueError("no_existe")
    rel = str(meta.get("archivo") or "")
    root = directorio_plantillas()
    if rel and ".." not in rel and "/" not in rel and "\\" not in rel:
        try:
            (root / rel).unlink(missing_ok=True)
        except OSError:
            pass
    data = _cargar_indice()
    lista = [p for p in data.get("plantillas", []) if str(p.get("id", "")) != str(meta.get("id"))]
    data["plantillas"] = lista
    _guardar_indice(data)
