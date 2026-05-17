#!/usr/bin/env python3
"""
Vigila cambios en el código y plantillas del proyecto y ejecuta ``portable_build.py``
(PyInstaller + copia de ``auth_users.json``) con debounce, para que el portable refleje
los cambios sin pasos manuales.

También reacciona a ``auth_users.json`` en la raíz.

Ignora ``dist/``, ``build/``, ``.git``, cachés de Python, etc., para no entrar en bucle
cuando PyInstaller escribe la salida.

Requiere: pip install watchdog

Uso (desde la raíz del proyecto):
  python tools/portable_watch.py
  python tools/portable_watch.py --no-initial
  python tools/portable_watch.py --solo-claves   # solo auth_users.json (más liviano)
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
AUTH_PATH = (ROOT / "auth_users.json").resolve()
BUILD_SCRIPT = ROOT / "tools" / "portable_build.py"

_DEBOUNCE_SEC = 3.5
_IGNORE_ROOT_DIRS = frozenset(
    {
        "dist",
        "build",
        ".git",
        "__pycache__",
        ".venv",
        "venv",
        "env",
        "data_local_imputaciones",
        ".cursor",
        "terminals",
    }
)
_IGNORE_SUFFIXES = frozenset({".pyc", ".pyo", ".tmp", ".log", ".toc", ".pkg"})

_timer: threading.Timer | None = None
_timer_lock = threading.Lock()


def _run_build() -> None:
    print("\n--- Build portable (PyInstaller + claves) ---\n", flush=True)
    r = subprocess.run([sys.executable, str(BUILD_SCRIPT)], cwd=str(ROOT))
    if r.returncode != 0:
        print(f"Build terminó con código {r.returncode}\n", flush=True)
    else:
        print("--- Listo: dist\\MisComprobantesAnalisis actualizado ---\n", flush=True)


def schedule_build(reason: str) -> None:
    global _timer
    print(f"  → Programado rebuild ({reason}), debounce {_DEBOUNCE_SEC:.0f}s…", flush=True)
    with _timer_lock:
        if _timer is not None:
            _timer.cancel()
        _timer = threading.Timer(_DEBOUNCE_SEC, _timer_fire)
        _timer.daemon = True
        _timer.start()


def _timer_fire() -> None:
    global _timer
    with _timer_lock:
        _timer = None
    _run_build()


def _rel_parts(path_str: str) -> tuple[str, ...] | None:
    try:
        p = Path(path_str).resolve()
        rel = p.relative_to(ROOT)
    except (OSError, ValueError):
        return None
    return rel.parts


def _should_watch_path(path_str: str, solo_claves: bool) -> bool:
    parts = _rel_parts(path_str)
    if not parts:
        return False
    if solo_claves:
        try:
            return Path(path_str).resolve() == AUTH_PATH
        except OSError:
            return False
    if parts[0] in _IGNORE_ROOT_DIRS:
        return False
    p = Path(path_str)
    if p.is_dir():
        return False
    name = p.name
    if name.startswith("."):
        return False
    suf = p.suffix.lower()
    if suf in _IGNORE_SUFFIXES:
        return False
    return True


def main() -> int:
    try:
        from watchdog.events import FileSystemEventHandler
        from watchdog.observers import Observer
    except ImportError:
        print(
            "Falta el paquete 'watchdog'. Instalalo con:\n"
            "  python -m pip install watchdog",
            file=sys.stderr,
        )
        return 1

    ap = argparse.ArgumentParser(
        description="Vigila el repo y recompila el portable automáticamente.",
    )
    ap.add_argument(
        "--no-initial",
        action="store_true",
        help="No ejecutar PyInstaller al iniciar.",
    )
    ap.add_argument(
        "--solo-claves",
        action="store_true",
        help="Solo vigilar auth_users.json (no templates ni .py).",
    )
    args = ap.parse_args()

    if not BUILD_SCRIPT.is_file():
        print(f"ERROR: no existe {BUILD_SCRIPT}", file=sys.stderr)
        return 1

    solo = bool(args.solo_claves)

    class Handler(FileSystemEventHandler):
        def on_modified(self, event):  # type: ignore[override]
            if event.is_directory:
                return
            if _should_watch_path(event.src_path, solo):
                schedule_build(Path(event.src_path).name)

        def on_created(self, event):  # type: ignore[override]
            if event.is_directory:
                return
            if _should_watch_path(event.src_path, solo):
                schedule_build(Path(event.src_path).name)

        def on_moved(self, event):  # type: ignore[override]
            if getattr(event, "is_directory", False):
                return
            dest = getattr(event, "dest_path", None)
            if dest and _should_watch_path(dest, solo):
                schedule_build(Path(dest).name)

    if not args.no_initial:
        print("Compilación inicial…", flush=True)
        _run_build()

    handler = Handler()
    observer = Observer()
    observer.schedule(handler, str(ROOT), recursive=not solo)
    observer.start()

    if solo:
        print(
            f"Solo claves: vigilando {AUTH_PATH.name}\n"
            f"Debounce {_DEBOUNCE_SEC:.0f}s. Ctrl+C para salir.\n",
            flush=True,
        )
    else:
        print(
            f"Vigilando proyecto (excepto dist/, build/, .git, …)\n"
            f"Cualquier cambio en .py, templates, i18n, spec, etc. → rebuild tras "
            f"~{_DEBOUNCE_SEC:.0f}s sin nuevos cambios.\n"
            "Ctrl+C para salir.\n",
            flush=True,
        )
    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        print("\nDeteniendo vigilante…", flush=True)
    observer.stop()
    observer.join(timeout=5)
    with _timer_lock:
        if _timer is not None:
            _timer.cancel()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
