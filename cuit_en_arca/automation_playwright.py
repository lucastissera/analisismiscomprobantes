"""
Automatización del navegador según el diagrama CUIT en ARCA.

Los selectores de AFIP cambian con frecuencia: si falla, revisar HTML vigente
y ajustar localizadores en este módulo (sin tocar sumar_imp_total).

Requisitos: pip install playwright && playwright install chromium
Habilitar en servidor: variable de entorno CUIT_EN_ARCA_PLAYWRIGHT=1
"""

from __future__ import annotations

import re
import time
from datetime import date
from pathlib import Path

from cuit_en_arca.credenciales import CredencialesArca
from cuit_en_arca.errores import (
    AutomatizacionArcaError,
    AutomatizacionNoDisponibleError,
    CuitRepresentadoNoEncontradoError,
)

LOGIN_URL = "https://auth.afip.gob.ar/contribuyente_/login.xhtml"
ESPERA_CORTA_SEC = 5


def _playwright_disponible() -> bool:
    try:
        import playwright  # noqa: F401

        return True
    except ImportError:
        return False


def _formatear_rango_afip(d: date, h: date) -> tuple[str, str]:
    """Texto para filtros de fecha (dd/mm/yyyy)."""
    return d.strftime("%d/%m/%Y"), h.strftime("%d/%m/%Y")


def _normalizar_cuit_busqueda(s: str) -> str:
    return re.sub(r"\D", "", s)


def ejecutar_descarga_mis_comprobantes(
    cred: CredencialesArca,
    fecha_desde: date,
    fecha_hasta: date,
    *,
    headless: bool = True,
) -> tuple[bytes, str]:
    """
    Ejecuta login, Mis Comprobantes, perfil representado, Emitidos, filtro fechas y descarga.
    Retorna (contenido binario, nombre sugerido .xlsx o .csv).
    """
    if not _playwright_disponible():
        raise AutomatizacionNoDisponibleError(
            "Playwright no está instalado. En local: pip install playwright && playwright install chromium"
        )

    from playwright.sync_api import TimeoutError as PlaywrightTimeout
    from playwright.sync_api import sync_playwright

    fd, fh = _formatear_rango_afip(fecha_desde, fecha_hasta)
    cuit_repr = _normalizar_cuit_busqueda(cred.cuit_representado)

    browser = None
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=headless)
            context = browser.new_context(
                locale="es-AR",
                timezone_id="America/Argentina/Buenos_Aires",
                accept_downloads=True,
            )
            page = context.new_page()
            page.set_default_timeout(60_000)

            # C → Login
            page.goto(LOGIN_URL, wait_until="domcontentloaded")

            # D → CUIT login (varios patrones posibles en la página)
            cuit_llenado = False
            for sel in (
                'input[name*="cuit" i]',
                'input[id*="cuit" i]',
                "input#F1\\:username",
                'input[type="text"]',
            ):
                loc = page.locator(sel).first
                try:
                    if loc.count() > 0 and loc.is_visible(timeout=2000):
                        loc.fill(cred.cuit_login)
                        cuit_llenado = True
                        break
                except Exception:
                    continue
            if not cuit_llenado:
                raise AutomatizacionArcaError(
                    "No se encontró el campo de CUIT en el login de AFIP (selector desactualizado)."
                )

            # Avanzar / siguiente
            for texto_btn in ("Siguiente", "Continuar", "Ingresar", "Aceptar"):
                btn = page.get_by_role("button", name=re.compile(texto_btn, re.I))
                if btn.count():
                    btn.first.click()
                    break
            else:
                page.keyboard.press("Enter")

            page.wait_for_load_state("networkidle", timeout=60_000)

            # E → Clave fiscal
            clave_ok = False
            for sel in (
                'input[type="password"]',
                'input[name*="password" i]',
                'input[id*="password" i]',
            ):
                loc = page.locator(sel).first
                try:
                    if loc.count() > 0 and loc.is_visible(timeout=2000):
                        loc.fill(cred.clave_fiscal)
                        clave_ok = True
                        break
                except Exception:
                    continue
            if not clave_ok:
                raise AutomatizacionArcaError(
                    "No se encontró el campo de clave fiscal (selector desactualizado)."
                )

            ingresar = page.get_by_role("button", name=re.compile("ingresar|aceptar", re.I))
            if ingresar.count():
                ingresar.first.click()
            else:
                page.keyboard.press("Enter")

            page.wait_for_load_state("networkidle", timeout=90_000)

            # F → Espera 5 seg
            time.sleep(ESPERA_CORTA_SEC)

            # G → Mis Comprobantes
            link = page.get_by_role("link", name=re.compile(r"mis\s*comprobantes", re.I))
            if not link.count():
                link = page.locator("a", has_text=re.compile(r"mis\s*comprobantes", re.I))
            if not link.count():
                raise AutomatizacionArcaError(
                    "No se encontró el enlace al servicio Mis Comprobantes tras el login."
                )
            try:
                with page.expect_popup(timeout=15_000) as pop:
                    link.first.click()
                mc = pop.value
            except Exception:
                link.first.click()
                page.wait_for_load_state("networkidle", timeout=60_000)
                mc = page
            mc.wait_for_load_state("domcontentloaded", timeout=60_000)

            # H / I / J → Perfil representado
            time.sleep(1)
            filas = mc.locator("tr, li, div[role='option'], a").filter(
                has_text=re.compile(r"\d{2}[-.]?\d{8}[-.]?\d")
            )
            encontrado = False
            for i in range(min(filas.count(), 200)):
                txt = filas.nth(i).inner_text()
                if cuit_repr in re.sub(r"\D", "", txt):
                    filas.nth(i).click()
                    encontrado = True
                    break
            if not encontrado:
                # intento por texto exacto CUIT formateado
                fmt = f"{cuit_repr[:2]}-{cuit_repr[2:10]}-{cuit_repr[10]}"
                alt = mc.get_by_text(fmt, exact=False)
                if alt.count():
                    alt.first.click()
                    encontrado = True
            if not encontrado:
                raise CuitRepresentadoNoEncontradoError(
                    "Verificar datos ingresados: el CUIT representado no aparece en la lista."
                )

            time.sleep(ESPERA_CORTA_SEC)

            # L → Emitidos
            emit = mc.get_by_role("link", name=re.compile("emitidos", re.I))
            if not emit.count():
                emit = mc.locator("a", has_text=re.compile("emitidos", re.I))
            if emit.count():
                emit.first.click()
            else:
                tab = mc.get_by_role("tab", name=re.compile("emitidos", re.I))
                if tab.count():
                    tab.first.click()
                else:
                    raise AutomatizacionArcaError(
                        "No se encontró la sección Emitidos en Mis Comprobantes."
                    )

            mc.wait_for_load_state("networkidle", timeout=60_000)

            # L1 / L2 → Rango de fechas y búsqueda (selectores genéricos; afinar según HTML ARCA)
            inputs_date = mc.locator('input[type="text"], input:not([type="hidden"])')
            n_inp = min(inputs_date.count(), 24)
            filled = 0
            for i in range(n_inp):
                el = inputs_date.nth(i)
                try:
                    if not el.is_visible():
                        continue
                    ph = (el.get_attribute("placeholder") or "").lower()
                    nm = (el.get_attribute("name") or "").lower()
                    if filled == 0 and (
                        "desde" in ph or "inicio" in ph or "fecha" in nm or "emision" in nm
                    ):
                        el.fill(fd)
                        filled += 1
                    elif filled == 1 and ("hasta" in ph or "fin" in ph):
                        el.fill(fh)
                        filled += 1
                except Exception:
                    continue
            if filled < 2:
                idx = 0
                for i in range(n_inp):
                    el = inputs_date.nth(i)
                    try:
                        if not el.is_visible():
                            continue
                        if idx == 0:
                            el.fill(fd)
                            idx += 1
                        elif idx == 1:
                            el.fill(fh)
                            idx += 1
                    except Exception:
                        continue

            buscar = mc.get_by_role("button", name=re.compile("buscar|consultar|aplicar", re.I))
            if not buscar.count():
                buscar = mc.locator("button, input[type='submit']").filter(
                    has_text=re.compile("buscar|consultar", re.I)
                )
            if buscar.count():
                buscar.first.click()
            mc.wait_for_load_state("networkidle", timeout=120_000)

            # M → Excel o CSV
            with mc.expect_download(timeout=120_000) as dl_info:
                excel_btn = mc.get_by_role("button", name=re.compile("excel|xlsx", re.I))
                if not excel_btn.count():
                    excel_btn = mc.locator("a, button").filter(
                        has_text=re.compile(r"excel|\.xlsx", re.I)
                    )
                if excel_btn.count():
                    excel_btn.first.click()
                else:
                    csv_btn = mc.get_by_role("button", name=re.compile("csv", re.I))
                    if not csv_btn.count():
                        csv_btn = mc.locator("a, button").filter(has_text=re.compile("csv", re.I))
                    if csv_btn.count():
                        csv_btn.first.click()
                    else:
                        raise AutomatizacionArcaError(
                            "No se encontró botón de descarga Excel ni CSV tras la búsqueda."
                        )
            download = dl_info.value

            path = download.path()
            if path is None:
                raise AutomatizacionArcaError("La descarga no generó archivo temporal.")
            data = Path(path).read_bytes()
            sug = download.suggested_filename or "mis_comprobantes_descarga"
            return data, sug

    except CuitRepresentadoNoEncontradoError:
        raise
    except PlaywrightTimeout as exc:
        raise AutomatizacionArcaError(
            "Tiempo de espera agotado en AFIP (sitio lento o página distinta a la esperada)."
        ) from exc
    except AutomatizacionArcaError:
        raise
    except Exception as exc:
        raise AutomatizacionArcaError(f"Error en automatización: {exc}") from exc
    finally:
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass
