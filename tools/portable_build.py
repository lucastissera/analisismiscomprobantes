#!/usr/bin/env python3
"""
Compila el portable con PyInstaller y copia ``auth_users.json`` a la carpeta del
ejecutable (junto a ``MisComprobantesAnalisis.exe``) si existe en la raíz del repo.

Uso: desde la raíz del proyecto
  python tools/portable_build.py

Lo invoca ``build_windows.bat`` y ``tools/portable_watch.py``.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DIST_DIR = ROOT / "dist" / "MisComprobantesAnalisis"
SPEC = ROOT / "MisComprobantesDesktop.spec"
AUTH_SRC = ROOT / "auth_users.json"


def main() -> int:
    if not SPEC.is_file():
        print(f"ERROR: no se encuentra {SPEC}", file=sys.stderr)
        return 1
    print("Ejecutando PyInstaller…", flush=True)
    r = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--noconfirm", str(SPEC)],
        cwd=str(ROOT),
    )
    if r.returncode != 0:
        return r.returncode
    if not DIST_DIR.is_dir():
        print(f"ERROR: no existe {DIST_DIR} tras compilar.", file=sys.stderr)
        return 1
    if AUTH_SRC.is_file():
        dest = DIST_DIR / "auth_users.json"
        shutil.copy2(AUTH_SRC, dest)
        print(f"Claves sincronizadas: {dest}", flush=True)
    else:
        print(
            "Aviso: no hay auth_users.json en la raíz del repo; "
            "el portable usará el ejemplo empaquetado o credenciales por entorno.",
            flush=True,
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
