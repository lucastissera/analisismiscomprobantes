#!/usr/bin/env python3
"""
Vigila ``auth_users.json`` en la raíz del repo. En cada guardado, espera unos
segundos (debounce) y ejecuta ``portable_build.py`` (recompilación + copia de claves).

Requiere: pip install watchdog

Uso (desde la raíz del proyecto):
  python tools/portable_watch.py
  python tools/portable_watch.py --no-initial   # no compila al arrancar
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

_DEBOUNCE_SEC = 2.5
_timer: threading.Timer | None = None
_timer_lock = threading.Lock()


def _run_build() -> None:
    print("\n--- Ejecutando build portable (claves + PyInstaller) ---\n", flush=True)
    r = subprocess.run([sys.executable, str(BUILD_SCRIPT)], cwd=str(ROOT))
    if r.returncode != 0:
        print(f"Build terminó con código {r.returncode}", flush=True)
    else:
        print("--- Build completado ---\n", flush=True)


def schedule_build() -> None:
    global _timer
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


def _is_root_auth_file(src_path: str) -> bool:
    try:
        return Path(src_path).resolve() == AUTH_PATH
    except OSError:
        return False


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

    ap = argparse.ArgumentParser(description="Vigila auth_users.json y recompila el portable.")
    ap.add_argument(
        "--no-initial",
        action="store_true",
        help="No ejecutar PyInstaller al iniciar el vigilante.",
    )
    args = ap.parse_args()

    if not BUILD_SCRIPT.is_file():
        print(f"ERROR: no existe {BUILD_SCRIPT}", file=sys.stderr)
        return 1

    class Handler(FileSystemEventHandler):
        def on_modified(self, event):  # type: ignore[override]
            if event.is_directory:
                return
            if _is_root_auth_file(event.src_path):
                print(f"Cambio detectado: {event.src_path}", flush=True)
                schedule_build()

        def on_created(self, event):  # type: ignore[override]
            if event.is_directory:
                return
            if _is_root_auth_file(event.src_path):
                print(f"Archivo creado: {event.src_path}", flush=True)
                schedule_build()

    if not args.no_initial:
        print("Compilación inicial…", flush=True)
        _run_build()

    observer = Observer()
    handler = Handler()
    observer.schedule(handler, str(ROOT), recursive=False)
    observer.start()
    print(
        f"Vigilando {AUTH_PATH.name} en {ROOT}\n"
        f"Guardá el archivo de claves para recompilar y copiar automáticamente "
        f"(espera ~{_DEBOUNCE_SEC:.0f}s tras el último guardado).\n"
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
