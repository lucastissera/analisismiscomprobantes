"""Instala Chromium de Playwright si falta (servidor web / Render)."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


def _directorio_browsers_servidor() -> Path:
    explicito = os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    if explicito:
        return Path(explicito)
    raiz = Path(__file__).resolve().parents[1]
    return raiz / ".playwright-browsers"


def _chromium_instalado(base: Path) -> bool:
    if not base.is_dir():
        return False
    patrones = (
        "chromium_headless_shell-*",
        "chromium-*",
        "chrome-headless-shell-*",
    )
    for patron in patrones:
        if any(base.glob(patron)):
            return True
    return False


def asegurar_chromium_playwright(*, forzar: bool = False) -> None:
    """Descarga Chromium en PLAYWRIGHT_BROWSERS_PATH si no existe."""
    if getattr(sys, "frozen", False):
        return

    destino = _directorio_browsers_servidor()
    destino.mkdir(parents=True, exist_ok=True)
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(destino))

    if not forzar and _chromium_instalado(destino):
        return

    env = os.environ.copy()
    env["PLAYWRIGHT_BROWSERS_PATH"] = str(destino)
    cmd_install = [sys.executable, "-m", "playwright", "install", "chromium"]
    r = subprocess.run(cmd_install, env=env, check=False)
    if r.returncode != 0:
        raise RuntimeError(
            "No se pudo instalar Chromium para Playwright. "
            f"Ejecute: PLAYWRIGHT_BROWSERS_PATH={destino} playwright install chromium"
        )

    if not _chromium_instalado(destino):
        raise RuntimeError(
            f"Chromium no quedó instalado en {destino}. "
            "Revise el build del servidor (playwright install chromium)."
        )
