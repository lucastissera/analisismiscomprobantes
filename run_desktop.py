"""
Punto de entrada para la aplicación de escritorio (.exe).

El servidor Flask escucha en 127.0.0.1: el procesamiento de Excel/CSV es **100 % local**
y **no requiere Internet**.

En el .exe, por defecto se intenta abrir la interfaz en una **ventana tipo aplicación**
(Edge o Chrome con ``--app=URL``), sin barra de pestañas. Si no hay Edge/Chrome en las
rutas habituales, se usa el navegador predeterminado del sistema.

Autenticación: copiá ``auth_users.json`` junto al .exe (misma carpeta) o usá el ejemplo
incluido (ver README).
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import webbrowser


def _puerto_deseado() -> int:
    return int(os.environ.get("PORT", "8765"))


def _abrir_interfaz(url: str) -> None:
    """Abre la UI: en portable, Edge/Chrome en modo app si se puede; si no, navegador."""
    if os.environ.get("OPEN_BROWSER", "1").strip().lower() not in (
        "1",
        "true",
        "yes",
        "on",
    ):
        return

    frozen = bool(getattr(sys, "frozen", False))
    use_app = (os.environ.get("DESKTOP_APP_WINDOW") or ("1" if frozen else "0")).strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )
    if not use_app:
        try:
            webbrowser.open(url)
        except Exception:
            pass
        return

    edge_paths = (
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    )
    for edge in edge_paths:
        if os.path.isfile(edge):
            try:
                subprocess.Popen(
                    [edge, f"--app={url}", "--no-first-run"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                return
            except OSError:
                pass

    for chrome in (
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
    ):
        if os.path.isfile(chrome):
            try:
                subprocess.Popen(
                    [chrome, f"--app={url}", "--no-first-run"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                return
            except OSError:
                pass

    try:
        webbrowser.open(url)
    except Exception:
        pass


def main() -> None:
    os.environ.setdefault("ENABLE_LOCAL_PLANTILLAS_IMPUTACION", "1")
    if getattr(sys, "frozen", False):
        os.environ.setdefault("CUIT_EN_ARCA_UI", "0")
        os.environ.setdefault("CUIT_EN_ARCA_PLAYWRIGHT", "0")

    port = _puerto_deseado()
    url = f"http://127.0.0.1:{port}/"

    from app import app

    threading.Timer(1.2, lambda: _abrir_interfaz(url)).start()

    print(
        f"\n  Mis Comprobantes — análisis local\n"
        f"  Interfaz: {url}\n"
        f"  Los archivos se procesan en esta PC (sin conexión a Internet).\n"
        f"  Cerrá esta ventana de consola para detener el programa.\n",
        flush=True,
    )
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
