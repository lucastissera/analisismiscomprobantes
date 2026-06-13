"""
Automatización del navegador según el diagrama CUIT en ARCA.

Los selectores de AFIP cambian con frecuencia: si falla, revisar HTML vigente
y ajustar localizadores en este módulo (sin tocar sumar_imp_total).

Requisitos: pip install playwright && playwright install chromium
Habilitar en servidor: variable de entorno CUIT_EN_ARCA_PLAYWRIGHT=1
"""

from __future__ import annotations

import re
import sys
import time
from datetime import date
from pathlib import Path
from typing import Literal

from cuit_en_arca.credenciales import CredencialesArca
from cuit_en_arca.descarga import DescargaArcaResult
from cuit_en_arca.errores import (
    AutomatizacionArcaError,
    AutomatizacionNoDisponibleError,
    CuitRepresentadoNoEncontradoError,
    LoginArcaError,
    SinComprobantesError,
)
from cuit_en_arca.stealth import (
    STEALTH_INIT_SCRIPT,
    USER_AGENT_CHROME,
    chromium_args,
    clic_humano,
    escribir_como_humano,
    pausa_humana,
)

LOGIN_URL = "https://auth.afip.gob.ar/contribuyente_/login.xhtml"
PORTAL_ARCA_URL = "https://portalcf.cloud.afip.gob.ar/portal/app/"
# Servicio MCMP vigente (Mis Comprobantes autenticado).
URL_EMITIDOS_MCMP = "https://fes.afip.gob.ar/mcmp/jsp/comprobantesEmitidos.do"
URL_RECIBIDOS_MCMP = "https://fes.afip.gob.ar/mcmp/jsp/comprobantesRecibidos.do"
URL_MCMP_ROOT = "https://fes.afip.gob.ar/mcmp/"
# Legacy / redirección desde portal (fallback).
URL_MIS_COMPROBANTES_DIRECTA = (
    "https://serviciosweb.afip.gob.ar/genericos/comprobantes/Default.aspx",
)
ESPERA_CORTA_SEC = 3.0
TipoComprobantes = Literal["emitidos", "recibidos", "ambos"]


def _paso(on_paso, clave: str, estado: str) -> None:
    """Notifica el avance de un paso de la checklist (si hay callback)."""
    if on_paso is None:
        return
    try:
        on_paso(clave, estado)
    except Exception:
        pass


def _esperar_pagina(page, timeout: int = 42_000) -> None:
    """AFIP rara vez alcanza networkidle; domcontentloaded + pausa breve."""
    try:
        page.wait_for_load_state("domcontentloaded", timeout=timeout)
    except Exception:
        pass
    pausa_humana(0.3, 0.7)


def _dir_diagnostico() -> Path:
    """Carpeta donde dejar capturas/HTML para depurar fallos de navegación."""
    try:
        if getattr(sys, "frozen", False):
            base = Path(sys.executable).resolve().parent
        else:
            base = Path.cwd()
    except Exception:
        base = Path.cwd()
    destino = base / "diagnostico_arca"
    try:
        destino.mkdir(parents=True, exist_ok=True)
        return destino
    except Exception:
        import tempfile

        destino = Path(tempfile.gettempdir()) / "diagnostico_arca"
        destino.mkdir(parents=True, exist_ok=True)
        return destino


def _volcar_diagnostico(page, etiqueta: str) -> str:
    """Guarda captura + HTML + URL del estado actual para depurar.

    Devuelve la ruta base (sin extensión) o cadena vacía si no se pudo.
    """
    try:
        ts = time.strftime("%Y%m%d_%H%M%S")
        base = _dir_diagnostico()
        stem = f"{ts}_{re.sub(r'[^a-zA-Z0-9_-]', '_', etiqueta)}"
        ruta_base = base / stem
        try:
            page.screenshot(path=str(ruta_base) + ".png", full_page=True)
        except Exception:
            pass
        try:
            url = page.url
        except Exception:
            url = ""
        try:
            html = page.content()
        except Exception:
            html = ""
        try:
            fechas = _campos_fecha_visibles(page)
        except Exception:
            fechas = "?"
        cabecera = (
            f"<!-- URL: {url} | fechaEmision_visible: {fechas} -->\n"
        )
        try:
            (Path(str(ruta_base) + ".html")).write_text(
                cabecera + (html or ""), encoding="utf-8", errors="ignore"
            )
        except Exception:
            pass
        return str(ruta_base)
    except Exception:
        return ""

_FRASES_ERROR_LOGIN = (
    "clave incorrect",
    "cuit incorrect",
    "datos incorrect",
    "usuario o contraseña",
    "usuario o contrasena",
    "no coincide",
    "error de autent",
    "verifique los datos",
    "credenciales",
    "acceso denegado",
    "no pudimos validar",
)


def _playwright_disponible() -> bool:
    try:
        import playwright  # noqa: F401

        return True
    except ImportError:
        return False


def _formatear_rango_afip(d: date, h: date) -> tuple[str, str]:
    return d.strftime("%d/%m/%Y"), h.strftime("%d/%m/%Y")


def _normalizar_cuit_busqueda(s: str) -> str:
    return re.sub(r"\D", "", s)


_PATRON_MIS_COMPROBANTES = re.compile(r"mis\s*comprobantes", re.I)
_TERMINO_BUSQUEDA_MC = "Mis Comprobantes"


def _iter_contextos(page):
    """Página principal e iframes (Mis Comprobantes suele cargar el contenido en frames)."""
    yield page
    for frame in page.frames:
        if frame != page.main_frame:
            yield frame


def _encontrar_seccion_tipo(root, etiqueta: str):
    """Localiza Emitidos / Recibidos (enlace, botón, pestaña o celda clickeable)."""
    variantes = (
        etiqueta,
        f"Comprobantes {etiqueta.lower()}",
        f"Comprobantes {etiqueta}",
    )
    for texto in variantes:
        estrategias = (
            root.get_by_role("link", name=texto, exact=True),
            root.locator(f"a:text-is('{texto}')"),
            root.locator(
                f"input[type='submit'][value='{texto}' i], "
                f"input[type='button'][value='{texto}' i], "
                f"input[value='{texto}' i]"
            ),
            root.get_by_role("link", name=re.compile(rf"^\s*{re.escape(texto)}\s*$", re.I)),
            root.get_by_role("tab", name=re.compile(re.escape(texto), re.I)),
            root.get_by_role("button", name=re.compile(re.escape(texto), re.I)),
            root.get_by_text(re.compile(rf"^\s*{re.escape(texto)}\s*$", re.I)),
            root.locator("a, button, span, td, li, div, label").filter(
                has_text=re.compile(re.escape(texto), re.I)
            ),
            root.locator(
                f"a[href*='{etiqueta}' i], a[href*='{etiqueta.lower()}' i], "
                f"a[onclick*='{etiqueta}' i]"
            ),
            root.locator(f"a:has-text('{texto}')"),
        )
        for loc in estrategias:
            try:
                n = min(loc.count(), 12)
                for i in range(n):
                    item = loc.nth(i)
                    if item.is_visible(timeout=900):
                        txt = (item.inner_text() or "").strip().lower()
                        if etiqueta.lower() in txt or texto.lower() in txt:
                            return item
                        href = item.get_attribute("href") or ""
                        if etiqueta.lower() in href.lower():
                            return item
            except Exception:
                continue
    return None


def _locator_enlace_mis_comprobantes(root):
    candidatos = (
        root.get_by_role("link", name=_PATRON_MIS_COMPROBANTES),
        root.locator("a", has_text=_PATRON_MIS_COMPROBANTES),
        root.locator(
            "button, [role='button'], [role='option'], li, div, span",
            has_text=_PATRON_MIS_COMPROBANTES,
        ),
    )
    for loc in candidatos:
        try:
            n = min(loc.count(), 15)
            for i in range(n):
                item = loc.nth(i)
                if item.is_visible(timeout=800):
                    return item
        except Exception:
            continue
    return None


def _locator_buscador_servicios(root):
    selectores = (
        'input[placeholder*="Buscar" i]',
        'input[placeholder*="servicio" i]',
        'input[id*="buscador" i]',
        'input[name*="buscador" i]',
        'input[type="search"]',
        "#buscadorInput",
        "#inputSearch",
        ".buscador input",
        'input[aria-label*="Buscar" i]',
    )
    for sel in selectores:
        loc = root.locator(sel).first
        try:
            if loc.count() > 0 and loc.is_visible(timeout=1200):
                return loc
        except Exception:
            continue
    try:
        sb = root.get_by_role("searchbox").first
        if sb.count() > 0 and sb.is_visible(timeout=1200):
            return sb
    except Exception:
        pass
    try:
        ph = root.get_by_placeholder(re.compile(r"buscar", re.I)).first
        if ph.count() > 0 and ph.is_visible(timeout=1200):
            return ph
    except Exception:
        pass
    return None


def _click_servicio_y_obtener_pagina(page, link) -> object:
    try:
        with page.expect_popup(timeout=20_000) as pop:
            clic_humano(link)
        mc = pop.value
    except Exception:
        clic_humano(link)
        _esperar_pagina(page, timeout=42_000)
        mc = page
    _esperar_pagina(mc, timeout=42_000)
    pausa_humana(0.56, 1.12)
    return mc


def _esperar_resultado_mis_comprobantes(page, intentos: int = 10):
    for _ in range(intentos):
        for ctx in _iter_contextos(page):
            link = _locator_enlace_mis_comprobantes(ctx)
            if link is not None:
                return link
        pausa_humana(0.35, 0.7)
    return None


def _buscar_mis_comprobantes_en_portal(page):
    buscador = None
    ctx_buscador = page
    for ctx in _iter_contextos(page):
        buscador = _locator_buscador_servicios(ctx)
        if buscador is not None:
            ctx_buscador = ctx
            break
    if buscador is None:
        raise AutomatizacionArcaError(
            "No se encontró la barra de búsqueda de servicios en ARCA."
        )

    escribir_como_humano(buscador, _TERMINO_BUSQUEDA_MC)
    pausa_humana(0.5, 1.0)

    btn_clicado = False
    for ctx in (ctx_buscador, page):
        try:
            btn = ctx.locator(
                "button[type='submit'], button .fa-search, .btn-search, "
                "[class*='search'] button, button[aria-label*='buscar' i]"
            ).first
            if btn.count() > 0 and btn.is_visible(timeout=800):
                clic_humano(btn)
                btn_clicado = True
                break
        except Exception:
            pass
        try:
            btn = ctx.get_by_role("button", name=re.compile(r"buscar|search", re.I)).first
            if btn.count() > 0 and btn.is_visible(timeout=800):
                clic_humano(btn)
                btn_clicado = True
                break
        except Exception:
            pass
    if not btn_clicado:
        page.keyboard.press("Enter")

    pausa_humana(0.7, 1.3)
    _esperar_pagina(page, timeout=35_000)

    link = _esperar_resultado_mis_comprobantes(page)
    if link is None:
        raise AutomatizacionArcaError(
            "No apareció «Mis Comprobantes» en los resultados del buscador de ARCA."
        )
    return _click_servicio_y_obtener_pagina(page, link)


def _pagina_es_login_afip(page) -> bool:
    try:
        url = page.url.lower()
        if "auth.afip" in url or "login.xhtml" in url:
            return True
        pwd = page.locator('input[type="password"]')
        return bool(pwd.count() and pwd.first.is_visible(timeout=1200))
    except Exception:
        return False


def _pagina_es_constatacion(page) -> bool:
    """Página pública de constatación (no es Mis Comprobantes autenticado)."""
    try:
        url = page.url.lower()
        if "servicioscf.afip" in url and "comprobante" in url:
            return True
    except Exception:
        pass
    try:
        cuerpo = page.locator("body").inner_text(timeout=4000).lower()
    except Exception:
        return False
    return (
        "constataci" in cuerpo
        or "se ha movido" in cuerpo
        or "esta pagina se ha movido" in cuerpo
    )


def _url_por_tipo(tipo: Literal["emitidos", "recibidos"]) -> str:
    return URL_EMITIDOS_MCMP if tipo == "emitidos" else URL_RECIBIDOS_MCMP


def _pagina_es_mcmp(page) -> bool:
    try:
        return "fes.afip.gob.ar/mcmp" in page.url.lower()
    except Exception:
        return False


def _pagina_sesion_expirada_mcmp(page) -> bool:
    """Página de MCMP "Su sesión ha expirado / usuario no logueado".

    Aparece al entrar por URL directa sin haber establecido el SSO desde el
    portal. La URL es /mcmp pero NO estamos autenticados.
    """
    try:
        if (page.title() or "").strip().lower().startswith("su sesión ha expirado"):
            return True
    except Exception:
        pass
    try:
        cuerpo = page.locator("body").inner_text(timeout=2500).lower()
    except Exception:
        return False
    return (
        "no está logueado" in cuerpo
        or "no esta logueado" in cuerpo
        or "la sesión ha expirado" in cuerpo
        or "la sesion ha expirado" in cuerpo
    )


def _en_servicio_mcmp(page) -> bool:
    """Estamos dentro del servicio MCMP autenticado (no login ni constatación)."""
    if _pagina_es_login_afip(page) or _pagina_es_constatacion(page):
        return False
    if _pagina_es_mcmp(page) and _pagina_sesion_expirada_mcmp(page):
        return False
    return _pagina_es_mcmp(page)


def _cuit_activo_mcmp(page) -> str | None:
    """CUIT del contribuyente activo según '.nombre-activo' (11 dígitos) o None."""
    for ctx in _iter_contextos(page):
        try:
            txt = ctx.evaluate(
                "() => { const el = document.querySelector('.nombre-activo');"
                " return el ? el.textContent : ''; }"
            )
        except Exception:
            txt = ""
        if txt:
            digitos = re.sub(r"\D", "", txt)
            if len(digitos) >= 11:
                return digitos[-11:]
    return None


def _limpiar_razon_social(txt: str) -> str:
    """Extrae el nombre/razón social quitando el CUIT y etiquetas del texto."""
    s = re.sub(r"\s+", " ", str(txt or "")).strip()
    if not s:
        return ""
    # Quitar el prefijo "Representando a:" y cualquier corchete (p. ej. "[ ]").
    s = re.sub(r"(?i)\brepresentando\s+a\s*:?", " ", s)
    s = re.sub(r"\[[^\]]*\]", " ", s)
    # Quitar CUIT formateado (xx-xxxxxxxx-x) y corridas de 11 dígitos.
    s = re.sub(r"\b\d{2}[\s.\-]?\d{8}[\s.\-]?\d\b", " ", s)
    s = re.sub(r"\b\d{11}\b", " ", s)
    # Quitar etiquetas habituales del encabezado del portal.
    s = re.sub(
        r"(?i)\b(cuit|cuil|cdi|clave fiscal|representand[oa]|representad[oa]|nivel|usuario)\b",
        " ",
        s,
    )
    s = re.sub(r"[|/]+", " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip(" -–—:·,.[]")
    return s.strip()


def _razon_social_activa_mcmp(page) -> str:
    """Razón social del contribuyente activo según '.nombre-activo' (sin CUIT)."""
    for ctx in _iter_contextos(page):
        try:
            txt = ctx.evaluate(
                "() => { const el = document.querySelector('.nombre-activo');"
                " return el ? el.textContent : ''; }"
            )
        except Exception:
            txt = ""
        nombre = _limpiar_razon_social(txt)
        if nombre:
            return nombre
    return ""


def _selectores_campo_fecha():
    return {
        "rango": (
            "input#fechaEmision",
            "input[name='fechaEmision']",
            "input[id='fechaEmision']",
            "input[name*='fechaEmision' i]:not([name*='Desde' i]):not([name*='Hasta' i])",
        ),
        "desde": (
            "input#fechaEmisionDesde",
            "input[name='fechaEmisionDesde']",
            "input[id*='fechaEmisionDesde' i]",
            "input[name*='fechaDesde' i]",
            "input[name*='fechaEmisionDesde' i]",
        ),
        "hasta": (
            "input#fechaEmisionHasta",
            "input[name='fechaEmisionHasta']",
            "input[id*='fechaEmisionHasta' i]",
            "input[name*='fechaHasta' i]",
            "input[name*='fechaEmisionHasta' i]",
        ),
    }


def _leer_valor_campo(loc) -> str:
    try:
        return (loc.input_value() or loc.get_attribute("value") or "").strip()
    except Exception:
        return ""


def _escribir_fecha_campo(loc, valor: str) -> None:
    loc.click()
    pausa_humana(0.12, 0.28)
    try:
        loc.fill(str(valor))
    except Exception:
        try:
            loc.press("Control+A")
        except Exception:
            pass
        escribir_como_humano(loc, valor)
    pausa_humana(0.1, 0.25)
    try:
        loc.evaluate(
            "el => { el.dispatchEvent(new Event('input', { bubbles: true })); "
            "el.dispatchEvent(new Event('change', { bubbles: true })); }"
        )
    except Exception:
        pass
    try:
        loc.press("Tab")
    except Exception:
        pass
    pausa_humana(0.12, 0.3)


def _fecha_aplicada_en_campo(valor: str, fd: str, fh: str) -> bool:
    if not valor:
        return False
    norm = re.sub(r"\s+", " ", valor)
    return fd in norm and fh in norm


def _llenar_daterangepicker_mcmp(ctx, fd: str, fh: str) -> bool:
    """MCMP usa #fechaEmision con daterangepicker (no campos sueltos desde/hasta)."""
    rango = f"{fd} - {fh}"
    campo = ctx.locator("#fechaEmision, input[name='fechaEmision']").first
    if not (campo.count() and campo.is_visible(timeout=1500)):
        return False

    try:
        ok = campo.evaluate(
            """(el, args) => {
                const fd = args[0], fh = args[1], rango = args[2];
                const $ = window.jQuery || window.$;
                if ($ && typeof $ === 'function') {
                    const $el = $(el);
                    const drp = $el.data('daterangepicker');
                    if (drp && window.moment) {
                        drp.setStartDate(window.moment(fd, 'DD/MM/YYYY'));
                        drp.setEndDate(window.moment(fh, 'DD/MM/YYYY'));
                        $el.val(rango).trigger('change');
                        return ($el.val() || '').includes(fd);
                    }
                }
                el.value = rango;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                return (el.value || '').includes(fd);
            }""",
            [fd, fh, rango],
        )
        if ok and _fecha_aplicada_en_campo(_leer_valor_campo(campo), fd, fh):
            return True
    except Exception:
        pass

    clic_humano(campo)
    pausa_humana(0.35, 0.65)
    picker = ctx.locator(".daterangepicker").last
    try:
        if picker.count() and picker.is_visible(timeout=2500):
            for texto in ("Personalizado", "Rango personalizado", "Custom Range"):
                item = picker.locator(
                    f"li[data-range-key='{texto}'], li:has-text('{texto}')"
                ).first
                try:
                    if item.count() and item.is_visible(timeout=500):
                        clic_humano(item)
                        pausa_humana(0.2, 0.45)
                        break
                except Exception:
                    continue
            inputs = picker.locator("input[type='text']")
            n = inputs.count()
            if n >= 2:
                _escribir_fecha_campo(inputs.nth(0), fd)
                _escribir_fecha_campo(inputs.nth(1), fh)
            elif n == 1:
                _escribir_fecha_campo(inputs.first, rango)
            apply = picker.locator("button.applyBtn, .applyBtn").first
            if apply.count() and apply.is_visible(timeout=800):
                clic_humano(apply)
            pausa_humana(0.25, 0.5)
    except Exception:
        pass

    if not _fecha_aplicada_en_campo(_leer_valor_campo(campo), fd, fh):
        _escribir_fecha_campo(campo, rango)
    return _fecha_aplicada_en_campo(_leer_valor_campo(campo), fd, fh)


def _campos_fecha_visibles(mc) -> bool:
    sels = _selectores_campo_fecha()
    for ctx in _iter_contextos(mc):
        for grupo in sels.values():
            for sel in grupo:
                loc = ctx.locator(sel).first
                try:
                    if loc.count() and loc.is_visible(timeout=600):
                        return True
                except Exception:
                    continue
    return False


def _form_tipo_cargado(mc) -> bool:
    """¿Cargó la pantalla Emitidos/Recibidos? Tolerante: campo de fecha visible
    o presente en el DOM (el datepicker puede taparlo), o la DataTable/botón."""
    if _campos_fecha_visibles(mc):
        return True
    sels = _selectores_campo_fecha()
    for ctx in _iter_contextos(mc):
        for grupo in sels.values():
            for sel in grupo:
                try:
                    if ctx.locator(sel).first.count():
                        return True
                except Exception:
                    continue
        try:
            if ctx.locator(
                "table.dataTable, #fechaEmision, "
                "input[value='Buscar' i], button:has-text('Buscar')"
            ).first.count():
                return True
        except Exception:
            continue
    return False


def _tiene_ui_mis_comprobantes(page) -> bool:
    """Emitidos/Recibidos o filtros de fecha visibles (validación estricta)."""
    if _pagina_es_login_afip(page) or _pagina_es_constatacion(page):
        return False
    if _pagina_es_mcmp(page) and _campos_fecha_visibles(page):
        return True
    for ctx in _iter_contextos(page):
        try:
            fechas = ctx.locator(
                "input[id*='fechaEmision' i], input[name*='fechaDesde' i], "
                "input#fechaEmision, input[name='fechaEmision']"
            ).first
            if fechas.count() and fechas.is_visible(timeout=700):
                return True
        except Exception:
            pass
        if _encontrar_seccion_tipo(ctx, "Emitidos") or _encontrar_seccion_tipo(
            ctx, "Recibidos"
        ):
            return True
    return False


def _pagina_parece_mis_comprobantes(page) -> bool:
    """Indicios de que cargó el servicio Mis Comprobantes."""
    if _tiene_ui_mis_comprobantes(page):
        return True
    if _pagina_es_login_afip(page) or _pagina_es_constatacion(page):
        return False
    try:
        cuerpo = page.locator("body").inner_text(timeout=4000).lower()
        return (
            "mis comprobantes" in cuerpo
            and ("emitidos" in cuerpo or "recibidos" in cuerpo)
        )
    except Exception:
        return False


def _esperar_mis_comprobantes_listo(page, timeout_sec: float = 22) -> None:
    limite = time.time() + timeout_sec
    while time.time() < limite:
        if _pagina_parece_mis_comprobantes(page):
            return
        pausa_humana(0.45, 0.9)


def _esperar_post_login(page, timeout_sec: float = 25) -> None:
    """Espera a salir del login de AFIP antes de abrir Mis Comprobantes."""
    limite = time.time() + timeout_sec
    while time.time() < limite:
        if not _pagina_es_login_afip(page):
            return
        pausa_humana(0.4, 0.8)
    _esperar_pagina(page, timeout=15_000)


def _ir_al_portal_arca(page) -> None:
    try:
        if "portalcf.cloud.afip" in page.url.lower():
            _esperar_pagina(page, timeout=35_000)
            return
    except Exception:
        pass
    page.goto(PORTAL_ARCA_URL, wait_until="domcontentloaded")
    _esperar_pagina(page, timeout=42_000)
    pausa_humana(0.7, 1.4)


def _ir_directo_mis_comprobantes(page):
    """Entra al servicio MCMP (acepta el selector de contribuyente; rápido)."""
    for url in (URL_MCMP_ROOT, URL_EMITIDOS_MCMP, URL_MIS_COMPROBANTES_DIRECTA):
        try:
            page.goto(url, wait_until="domcontentloaded")
            _esperar_pagina(page, timeout=42_000)
            # Espera breve a que el redirect resuelva a MCMP (o al login).
            for _ in range(12):
                if _pagina_es_login_afip(page):
                    break
                if _en_servicio_mcmp(page):
                    return page
                pausa_humana(0.3, 0.6)
            if _en_servicio_mcmp(page):
                return page
        except Exception:
            continue
    return None


def _ir_a_pantalla_tipo(mc, tipo: Literal["emitidos", "recibidos"]) -> None:
    """Navega a emitidos/recibidos MCMP (URL directa o clic en el menú)."""
    url = _url_por_tipo(tipo)
    etiqueta = "Emitidos" if tipo == "emitidos" else "Recibidos"
    destino = "comprobantesemitidos" if tipo == "emitidos" else "comprobantesrecibidos"

    try:
        if destino in mc.url.lower() and _campos_fecha_visibles(mc):
            return
    except Exception:
        pass

    def _esperar_formulario(intentos: int = 16) -> bool:
        for _ in range(intentos):
            if _pagina_es_login_afip(mc) or _pagina_sesion_expirada_mcmp(mc):
                raise AutomatizacionArcaError(
                    f"La sesión expiró al abrir {etiqueta} en Mis Comprobantes."
                )
            if _form_tipo_cargado(mc):
                return True
            # Si en vez del formulario aparece el menú con enlaces de tipo,
            # hacemos clic para entrar a la pantalla correspondiente.
            if _enlaces_tipo_visibles(mc):
                try:
                    _ir_a_tipo_comprobantes(mc, tipo)
                except Exception:
                    pass
                if _form_tipo_cargado(mc):
                    return True
            pausa_humana(0.4, 0.8)
        return False

    for intento in range(3):
        try:
            mc.goto(url, wait_until="domcontentloaded")
            _esperar_pagina(mc, timeout=42_000)
            pausa_humana(0.5, 1.0)
        except Exception:
            pass

        if _esperar_formulario():
            return

        # No cargó el formulario: reintentar recargando o volviendo a la raíz
        # del servicio (por si MCMP exige pasar por el selector/menú).
        try:
            if intento == 0:
                mc.reload(wait_until="domcontentloaded")
            else:
                mc.goto(URL_MCMP_ROOT, wait_until="domcontentloaded")
            _esperar_pagina(mc, timeout=42_000)
            pausa_humana(0.5, 1.0)
        except Exception:
            pass

    _volcar_diagnostico(mc, f"sin_formulario_{tipo}")
    raise AutomatizacionArcaError(
        f"No cargó el formulario de {etiqueta} en Mis Comprobantes "
        f"({url})."
    )


def _nuevo_contexto_stealth(playwright, *, headless: bool):
    if not getattr(sys, "frozen", False):
        from cuit_en_arca.ensure_playwright import asegurar_chromium_playwright

        asegurar_chromium_playwright()
    browser = playwright.chromium.launch(
        headless=headless,
        args=chromium_args(headless),
    )
    context = browser.new_context(
        locale="es-AR",
        timezone_id="America/Argentina/Buenos_Aires",
        accept_downloads=True,
        user_agent=USER_AGENT_CHROME,
        viewport={"width": 1366, "height": 768},
        device_scale_factor=1,
        has_touch=False,
        is_mobile=False,
    )
    context.add_init_script(STEALTH_INIT_SCRIPT)
    # NOTA: no interceptamos peticiones (context.route). Interceptar todas las
    # requests ralentizaba el login de AFIP (cadena de redirects) y rompía la
    # apertura de Mis Comprobantes. El ahorro de bloquear trackers no compensa.
    return browser, context


def _detectar_fallo_login(page, cuit: str) -> None:
    pausa_humana(0.5, 1.0)
    try:
        cuerpo = page.locator("body").inner_text(timeout=5000).lower()
    except Exception:
        cuerpo = ""
    if any(f in cuerpo for f in _FRASES_ERROR_LOGIN):
        raise LoginArcaError(
            f"No se pudo ingresar a ARCA con CUIT {cuit} (clave o CUIT incorrectos)."
        )
    url = page.url.lower()
    if "login" in url or "auth.afip" in url:
        pwd = page.locator('input[type="password"]')
        try:
            if pwd.count() and pwd.first.is_visible(timeout=2000):
                raise LoginArcaError(
                    f"No se pudo ingresar a ARCA con CUIT {cuit} (clave o CUIT incorrectos)."
                )
        except LoginArcaError:
            raise
        except Exception:
            pass


def _llenar_cuit_y_avanzar(page, cuit: str) -> None:
    cuit_llenado = False
    for sel in (
        "input#F1\\:username",
        'input[name*="cuit" i]',
        'input[id*="cuit" i]',
        'input[type="text"]',
    ):
        loc = page.locator(sel).first
        try:
            if loc.count() > 0 and loc.is_visible(timeout=2000):
                escribir_como_humano(loc, cuit)
                cuit_llenado = True
                break
        except Exception:
            continue
    if not cuit_llenado:
        raise AutomatizacionArcaError(
            "No se encontró el campo de CUIT en el login de AFIP (selector desactualizado)."
        )
    btn_sig = page.locator("input#F1\\:btnSiguiente, button#F1\\:btnSiguiente").first
    if btn_sig.count() and btn_sig.is_visible(timeout=1500):
        clic_humano(btn_sig)
    else:
        for texto_btn in ("Siguiente", "Continuar", "Ingresar", "Aceptar"):
            btn = page.get_by_role("button", name=re.compile(texto_btn, re.I))
            if btn.count():
                clic_humano(btn.first)
                break
        else:
            page.keyboard.press("Enter")
    _esperar_pagina(page, timeout=42_000)
    pausa_humana(0.35, 0.84)


def _login_clave_fiscal(page, clave: str, cuit: str) -> None:
    clave_ok = False
    for sel in (
        "input#F1\\:password",
        'input[type="password"]',
        'input[name*="password" i]',
        'input[id*="password" i]',
    ):
        loc = page.locator(sel).first
        try:
            if loc.count() > 0 and loc.is_visible(timeout=2000):
                escribir_como_humano(loc, clave)
                clave_ok = True
                break
        except Exception:
            continue
    if not clave_ok:
        raise AutomatizacionArcaError(
            "No se encontró el campo de clave fiscal (selector desactualizado)."
        )
    btn_ing = page.locator("input#F1\\:btnIngresar, button#F1\\:btnIngresar").first
    if btn_ing.count():
        try:
            if btn_ing.is_visible(timeout=1500):
                clic_humano(btn_ing)
            else:
                page.keyboard.press("Enter")
        except Exception:
            page.keyboard.press("Enter")
    else:
        ingresar = page.get_by_role("button", name=re.compile("ingresar|aceptar", re.I))
        if ingresar.count():
            clic_humano(ingresar.first)
        else:
            page.keyboard.press("Enter")
    _esperar_pagina(page, timeout=63_000)
    _detectar_fallo_login(page, cuit)
    _esperar_post_login(page, timeout_sec=35)


def _abrir_mis_comprobantes(page):
    pausa_humana(ESPERA_CORTA_SEC * 0.8, ESPERA_CORTA_SEC * 1.4)
    _esperar_post_login(page)
    _esperar_pagina(page, timeout=42_000)

    # IMPORTANTE: entrar por URL directa a fes.afip.gob.ar/mcmp NO autentica
    # (muestra "Su sesión ha expirado"). El SSO se establece al abrir el servicio
    # desde el portal, así que ese es el camino principal.
    link = _esperar_resultado_mis_comprobantes(page, intentos=8)
    if link is not None:
        mc = _click_servicio_y_obtener_pagina(page, link)
        if _en_servicio_mcmp(mc) or _tiene_ui_mis_comprobantes(mc):
            return mc

    try:
        _ir_al_portal_arca(page)
        mc = _buscar_mis_comprobantes_en_portal(page)
        if _en_servicio_mcmp(mc) or _tiene_ui_mis_comprobantes(mc):
            return mc
    except AutomatizacionArcaError:
        pass

    # Reintento del enlace en el home del portal (por si el buscador falló).
    try:
        _ir_al_portal_arca(page)
        link = _esperar_resultado_mis_comprobantes(page, intentos=8)
        if link is not None:
            mc = _click_servicio_y_obtener_pagina(page, link)
            if _en_servicio_mcmp(mc) or _tiene_ui_mis_comprobantes(mc):
                return mc
    except AutomatizacionArcaError:
        pass

    # Último recurso: URL directa (solo sirve si el SSO ya quedó establecido).
    for _ in range(2):
        mc = _ir_directo_mis_comprobantes(page)
        if mc is not None and _en_servicio_mcmp(mc):
            return mc
        pausa_humana(0.8, 1.6)

    _volcar_diagnostico(page, "no_abrio_mis_comprobantes")
    raise AutomatizacionArcaError(
        "No se pudo abrir Mis Comprobantes tras el login "
        "(enlace en el portal, buscador y URL directa fallaron). "
        "Verifique que el CUIT tenga el servicio «Mis Comprobantes» habilitado."
    )


def _ya_en_pantalla_comprobantes(mc) -> bool:
    """Emitidos/Recibidos o filtros de fecha → ya no hace falta elegir contribuyente."""
    if _campos_fecha_visibles(mc):
        return True
    for ctx in _iter_contextos(mc):
        for etiqueta in ("Emitidos", "Recibidos"):
            if _encontrar_seccion_tipo(ctx, etiqueta) is not None:
                return True
    return False


def _filas_cuit_clicables(mc):
    """Elementos visibles que parecen opciones de contribuyente en un selector."""
    filas = mc.locator(
        "table tbody tr, ul li, div[role='option'], a, button, label"
    ).filter(has_text=re.compile(r"\d{2}[-.]?\d{8}[-.]?\d"))
    resultado = []
    vistos: set[str] = set()
    for i in range(min(filas.count(), 40)):
        item = filas.nth(i)
        try:
            if not item.is_visible(timeout=400):
                continue
            txt = item.inner_text()
            digitos = re.sub(r"\D", "", txt)
            if len(digitos) < 11:
                continue
            cuit = digitos[-11:]
            if cuit in vistos:
                continue
            vistos.add(cuit)
            resultado.append((item, cuit))
        except Exception:
            continue
    return resultado


def _intentar_clic_contribuyente(mc, cuit_repr_n: str) -> bool:
    fmt = f"{cuit_repr_n[:2]}-{cuit_repr_n[2:10]}-{cuit_repr_n[10]}"
    for loc in (
        mc.get_by_role("link", name=re.compile(re.escape(fmt), re.I)),
        mc.locator("a, tr, li, button").filter(has_text=re.compile(re.escape(fmt))),
        mc.locator("a, tr, li, button").filter(has_text=re.compile(re.escape(cuit_repr_n))),
    ):
        try:
            if loc.count() and loc.first.is_visible(timeout=1000):
                clic_humano(loc.first)
                pausa_humana(ESPERA_CORTA_SEC * 0.6, ESPERA_CORTA_SEC * 1.0)
                return True
        except Exception:
            continue

    for item, cuit in _filas_cuit_clicables(mc):
        if cuit == cuit_repr_n:
            clic_humano(item)
            pausa_humana(ESPERA_CORTA_SEC * 0.6, ESPERA_CORTA_SEC * 1.0)
            return True
    return False


def _elegir_perfil_representado(
    mc,
    cuit_repr: str,
    *,
    cuit_login: str | None = None,
) -> None:
    pausa_humana(0.4, 0.9)

    cuit_repr_n = _normalizar_cuit_busqueda(cuit_repr)
    cuit_login_n = _normalizar_cuit_busqueda(cuit_login) if cuit_login else cuit_repr_n

    # Si el contribuyente activo ya es el representado, no hace falta seleccionar.
    activo = _cuit_activo_mcmp(mc)
    if activo == cuit_repr_n:
        return

    # Intento directo en la pantalla actual (selector de contribuyente).
    if _intentar_clic_contribuyente(mc, cuit_repr_n):
        if _cuit_activo_mcmp(mc) in (None, cuit_repr_n):
            return

    # No estaba: vuelvo a la raíz del servicio (lista de contribuyentes) y reintento.
    try:
        mc.goto(URL_MCMP_ROOT, wait_until="domcontentloaded")
        _esperar_pagina(mc, timeout=42_000)
        pausa_humana(0.5, 1.0)
    except Exception:
        pass

    if _intentar_clic_contribuyente(mc, cuit_repr_n):
        return

    # Sin lista de perfiles o el representado coincide con el ingreso: AFIP ya usa el activo.
    activo = _cuit_activo_mcmp(mc)
    if cuit_repr_n == cuit_login_n or activo == cuit_repr_n or _ya_en_pantalla_comprobantes(mc):
        return
    if not _filas_cuit_clicables(mc):
        return

    raise CuitRepresentadoNoEncontradoError(
        "Verificar datos ingresados: el CUIT representado no aparece en la lista."
    )


def _enlaces_tipo_visibles(mc) -> bool:
    for ctx in _iter_contextos(mc):
        if _encontrar_seccion_tipo(ctx, "Emitidos") or _encontrar_seccion_tipo(
            ctx, "Recibidos"
        ):
            return True
    return False


def _restablecer_menu_tipos(mc) -> None:
    """Tras descargar Emitidos, vuelve al menú donde están Emitidos/Recibidos."""
    if _enlaces_tipo_visibles(mc):
        return

    for ctx in _iter_contextos(mc):
        for texto in (
            "Volver",
            "Nueva consulta",
            "Nueva búsqueda",
            "Menú principal",
            "Menú",
        ):
            try:
                loc = ctx.get_by_role("link", name=re.compile(texto, re.I))
                if loc.count() and loc.first.is_visible(timeout=900):
                    clic_humano(loc.first)
                    _esperar_pagina(mc, timeout=35_000)
                    pausa_humana(0.45, 0.9)
                    if _enlaces_tipo_visibles(mc):
                        return
            except Exception:
                continue

    for _ in range(3):
        try:
            mc.go_back()
            _esperar_pagina(mc, timeout=25_000)
            pausa_humana(0.45, 0.9)
            if _enlaces_tipo_visibles(mc):
                return
        except Exception:
            break

    try:
        mc.goto(URL_EMITIDOS_MCMP, wait_until="domcontentloaded")
        _esperar_pagina(mc, timeout=42_000)
        pausa_humana(1.0, 2.0)
        _esperar_mis_comprobantes_listo(mc, timeout_sec=24)
    except Exception:
        pass


def _llenar_fechas_en_contexto(ctx, fd: str, fh: str) -> int:
    """Devuelve 2 si las fechas quedaron cargadas (rango único o desde+hasta)."""
    if _llenar_daterangepicker_mcmp(ctx, fd, fh):
        return 2

    sels = _selectores_campo_fecha()
    rango_txt = f"{fd} - {fh}"

    for sel in sels["rango"]:
        loc = ctx.locator(sel).first
        try:
            if loc.count() and loc.is_visible(timeout=1200):
                _escribir_fecha_campo(loc, rango_txt)
                if _fecha_aplicada_en_campo(_leer_valor_campo(loc), fd, fh):
                    return 2
        except Exception:
            continue

    filled = 0
    for sel in sels["desde"]:
        loc = ctx.locator(sel).first
        try:
            if loc.count() and loc.is_visible(timeout=1200):
                _escribir_fecha_campo(loc, fd)
                if fd in _leer_valor_campo(loc):
                    filled += 1
                    break
        except Exception:
            continue

    for sel in sels["hasta"]:
        loc = ctx.locator(sel).first
        try:
            if loc.count() and loc.is_visible(timeout=1200):
                _escribir_fecha_campo(loc, fh)
                if fh in _leer_valor_campo(loc):
                    filled += 1
                    break
        except Exception:
            continue

    return filled


def _ir_a_tipo_comprobantes(mc, tipo: Literal["emitidos", "recibidos"]) -> None:
    etiqueta = "Emitidos" if tipo == "emitidos" else "Recibidos"
    pausa_humana(0.5, 1.0)
    _esperar_pagina(mc, timeout=42_000)

    if _pagina_es_constatacion(mc):
        raise AutomatizacionArcaError(
            f"No se encontró la sección {etiqueta} en Mis Comprobantes "
            "(se abrió la página de constatación pública en lugar del servicio autenticado)."
        )

    for intento in range(14):
        for ctx in _iter_contextos(mc):
            loc = _encontrar_seccion_tipo(ctx, etiqueta)
            if loc is not None:
                clic_humano(loc)
                _esperar_pagina(mc, timeout=42_000)
                pausa_humana(0.35, 0.7)
                return
        if intento in (3, 7, 11):
            _restablecer_menu_tipos(mc)
        pausa_humana(0.5, 1.0)

    raise AutomatizacionArcaError(
        f"No se encontró la sección {etiqueta} en Mis Comprobantes."
    )


def _llenar_campo_fecha(mc, selectores: tuple[str, ...], valor: str) -> bool:
    for sel in selectores:
        loc = mc.locator(sel).first
        try:
            if loc.count() > 0 and loc.is_visible(timeout=1500):
                _escribir_fecha_campo(loc, valor)
                return bool(_leer_valor_campo(loc))
        except Exception:
            continue
    return False


def _aplicar_filtro_fechas_y_buscar(mc, fd: str, fh: str) -> str:
    pausa_humana(0.35, 0.7)
    filled = 0
    for ctx in _iter_contextos(mc):
        filled = max(filled, _llenar_fechas_en_contexto(ctx, fd, fh))
        if filled >= 2:
            break

    if filled < 2:
        for ctx in _iter_contextos(mc):
            inputs_date = ctx.locator('input[type="text"], input:not([type="hidden"])')
            n_inp = min(inputs_date.count(), 24)
            for i in range(n_inp):
                el = inputs_date.nth(i)
                try:
                    if not el.is_visible():
                        continue
                    ph = (el.get_attribute("placeholder") or "").lower()
                    nm = (el.get_attribute("name") or "").lower()
                    el_id = (el.get_attribute("id") or "").lower()
                    if filled == 0 and (
                        "desde" in ph
                        or "inicio" in ph
                        or "desde" in nm
                        or "desde" in el_id
                        or nm == "fechaemision"
                        or el_id == "fechaemision"
                        or "emision" in nm
                    ):
                        _escribir_fecha_campo(el, fd if "emision" not in nm else f"{fd} - {fh}")
                        filled = 2 if "emision" in nm and "desde" not in nm else filled + 1
                    elif filled == 1 and (
                        "hasta" in ph or "fin" in ph or "hasta" in nm or "hasta" in el_id
                    ):
                        _escribir_fecha_campo(el, fh)
                        filled += 1
                except Exception:
                    continue
            if filled >= 2:
                break

    if filled < 2:
        raise AutomatizacionArcaError(
            f"No se pudieron cargar las fechas {fd} — {fh} en el formulario de "
            "Mis Comprobantes (campos no encontrados o no editables)."
        )

    pausa_humana(0.25, 0.55)
    buscar = None
    for ctx in _iter_contextos(mc):
        candidato = ctx.locator(
            "#buscarComprobantes, input#buscarComprobantes, "
            "button#buscarComprobantes, input[value='Buscar'], "
            "input[value='BUSCAR'], button:has-text('Buscar')"
        ).first
        try:
            if candidato.count() and candidato.is_visible(timeout=1200):
                buscar = candidato
                break
        except Exception:
            continue

    if buscar is None:
        btn = mc.get_by_role("button", name=re.compile("buscar|consultar|aplicar", re.I))
        if btn.count() and btn.first.is_visible(timeout=1000):
            buscar = btn.first
        else:
            raise AutomatizacionArcaError(
                "No se encontró el botón Buscar en Mis Comprobantes."
            )

    try:
        with mc.expect_response(
            lambda r: "ajax.do" in r.url and r.status == 200,
            timeout=90_000,
        ):
            clic_humano(buscar)
    except Exception:
        clic_humano(buscar)

    return _esperar_resultados_mcmp(mc)


_SEL_EXCEL = (
    "button.buttons-excel.buttons-html5",
    "button.buttons-excel",
    "a.buttons-excel",
    "[class*='buttons-excel']",
    "a[href*='Excel' i]",
    "button[title*='Excel' i]",
    "a[title*='Excel' i]",
    "button:has-text('Excel')",
    "a:has-text('Excel')",
)

# Respaldo: cuando hay muchos comprobantes, ARCA a veces ofrece solo CSV.
_SEL_CSV = (
    "button.buttons-csv.buttons-html5",
    "button.buttons-csv",
    "a.buttons-csv",
    "[class*='buttons-csv']",
    "a[href*='csv' i]",
    "button[title*='CSV' i]",
    "a[title*='CSV' i]",
    "button:has-text('CSV')",
    "a:has-text('CSV')",
)

_FRASES_SIN_RESULTADOS = (
    "no se encontraron",
    "sin resultados",
    "no hay datos",
    "no existen comprobantes",
    "no se hallaron",
    "0 registros",
)


def _locator_boton_por_sel(mc, selectores):
    """Devuelve el primer botón/enlace visible que matchee alguno de los selectores."""
    for ctx in _iter_contextos(mc):
        for sel in selectores:
            loc = ctx.locator(sel)
            try:
                n = min(loc.count(), 6)
            except Exception:
                n = 0
            for i in range(n):
                item = loc.nth(i)
                try:
                    if item.is_visible(timeout=400):
                        return item
                except Exception:
                    continue
    return None


def _locator_boton_excel(mc):
    """Devuelve el primer botón/enlace de exportación a Excel visible (página o frames)."""
    return _locator_boton_por_sel(mc, _SEL_EXCEL)


def _locator_boton_csv(mc):
    """Devuelve el primer botón/enlace de exportación a CSV visible (página o frames)."""
    return _locator_boton_por_sel(mc, _SEL_CSV)


def _locator_boton_descarga(mc):
    """Prefiere Excel; si ARCA solo ofrece CSV (muchos comprobantes), usa CSV.

    Devuelve (locator, formato) con formato 'xlsx' o 'csv', o (None, None).
    """
    btn = _locator_boton_excel(mc)
    if btn is not None:
        return btn, "xlsx"
    btn = _locator_boton_csv(mc)
    if btn is not None:
        return btn, "csv"
    return None, None


def _hay_boton_descarga(mc) -> bool:
    return _locator_boton_descarga(mc)[0] is not None


def _hay_resultados_en_tabla(mc) -> bool:
    for ctx in _iter_contextos(mc):
        filas = ctx.locator(
            "#tablaComprobantes tbody tr, table.dataTable tbody tr, "
            "table#tabla tbody tr, .dataTables_wrapper table tbody tr"
        )
        try:
            count = filas.count()
        except Exception:
            count = 0
        for i in range(min(count, 3)):
            try:
                txt = (filas.nth(i).inner_text(timeout=600) or "").strip()
            except Exception:
                txt = ""
            if txt and not any(f in txt.lower() for f in _FRASES_SIN_RESULTADOS):
                return True
    return False


def _sin_resultados(mc) -> bool:
    for ctx in _iter_contextos(mc):
        try:
            cuerpo = ctx.locator("body").inner_text(timeout=2000).lower()
        except Exception:
            continue
        if any(f in cuerpo for f in _FRASES_SIN_RESULTADOS):
            return True
    return False


def _procesando(mc) -> bool:
    for ctx in _iter_contextos(mc):
        try:
            proc = ctx.locator(
                ".dataTables_processing, #loading, .loading, .blockUI, "
                ".cargando, [id*='procesando' i]"
            ).first
            if proc.count() and proc.is_visible(timeout=300):
                return True
        except Exception:
            continue
    return False


def _esperar_resultados_mcmp(mc, timeout_sec: float = 95) -> str:
    """Espera fin de consulta. Devuelve 'ok', 'vacio' o 'timeout'."""
    limite = time.time() + timeout_sec
    while time.time() < limite:
        if _procesando(mc):
            pausa_humana(0.5, 1.0)
            continue
        if _hay_boton_descarga(mc) or _hay_resultados_en_tabla(mc):
            return "ok"
        if _sin_resultados(mc):
            return "vacio"
        pausa_humana(0.5, 1.0)
    _esperar_pagina(mc, timeout=15_000)
    if _hay_boton_descarga(mc) or _hay_resultados_en_tabla(mc):
        return "ok"
    if _sin_resultados(mc):
        return "vacio"
    return "timeout"


def _extraer_de_zip(data: bytes, nombre_zip: str) -> tuple[bytes, str]:
    """ARCA puede entregar la descarga como ZIP con el CSV/Excel adentro.

    Devuelve (bytes_internos, nombre_interno). Prioriza CSV, luego Excel.
    """
    import io as _io
    import zipfile

    with zipfile.ZipFile(_io.BytesIO(data)) as zf:
        nombres = [n for n in zf.namelist() if not n.endswith("/")]
        elegido = None
        for ext in (".csv", ".xlsx", ".xls"):
            for n in nombres:
                if n.lower().endswith(ext):
                    elegido = n
                    break
            if elegido:
                break
        if elegido is None:
            raise AutomatizacionArcaError(
                f"La descarga ZIP no contiene CSV ni Excel (contenido: {nombres})."
            )
        contenido = zf.read(elegido)
    return contenido, Path(elegido).name


def _es_zip(data: bytes) -> bool:
    return data[:4] == b"PK\x03\x04"


def _descargar_excel_o_csv(mc, estado: str = "ok") -> tuple[bytes, str]:
    if estado == "vacio":
        raise SinComprobantesError(
            "La consulta no devolvió comprobantes para el período indicado."
        )

    btn, formato = _locator_boton_descarga(mc)
    if btn is None:
        # Última espera por si DataTables tarda en renderizar los botones.
        for _ in range(12):
            pausa_humana(0.6, 1.1)
            btn, formato = _locator_boton_descarga(mc)
            if btn is not None:
                break
            if _sin_resultados(mc):
                raise SinComprobantesError(
                    "La consulta no devolvió comprobantes para el período indicado."
                )
    if btn is None:
        if _sin_resultados(mc):
            raise SinComprobantesError(
                "La consulta no devolvió comprobantes para el período indicado."
            )
        raise AutomatizacionArcaError(
            "No se encontró botón de descarga (Excel ni CSV) tras la búsqueda "
            "(¿terminó de cargar la consulta?)."
        )

    try:
        btn.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass

    with mc.expect_download(timeout=120_000) as dl_info:
        try:
            clic_humano(btn)
        except Exception:
            btn.click(force=True)
    download = dl_info.value
    path = download.path()
    if path is None:
        raise AutomatizacionArcaError("La descarga no generó archivo temporal.")
    data = Path(path).read_bytes()

    sug = download.suggested_filename or f"mis_comprobantes_descarga.{formato}"

    # ARCA suele entregar el CSV (muchos comprobantes) dentro de un ZIP. El
    # botón Excel da un .xlsx (también ZIP internamente), por eso solo tratamos
    # como contenedor cuando el nombre dice .zip o cuando vino por el botón CSV.
    if sug.lower().endswith(".zip") or (formato == "csv" and _es_zip(data)):
        data, sug = _extraer_de_zip(data, sug)

    es_csv = sug.lower().endswith(".csv")
    if es_csv:
        # Un CSV válido debe tener al menos encabezado + 1 fila de datos.
        texto = data.decode("utf-8", "ignore")
        if not data or texto.count("\n") < 1:
            raise AutomatizacionArcaError(
                "El CSV descargado está vacío o incompleto. "
                "Es probable que la búsqueda no haya aplicado el rango de fechas."
            )
    elif len(data) < 800:
        raise AutomatizacionArcaError(
            "El Excel descargado está vacío o incompleto. "
            "Es probable que la búsqueda no haya aplicado el rango de fechas."
        )
    return data, sug


_JS_LEER_COLUMNAS = """
() => {
  const $ = window.jQuery || window.$;
  if (!$ || !$.fn || !$.fn.dataTable) return null;
  let s = null;
  $('table').each(function () {
    try {
      if ($.fn.dataTable.isDataTable(this)) { s = $(this).DataTable().settings()[0]; return false; }
    } catch (e) {}
  });
  if (!s || !s.aoColumns) return null;
  return s.aoColumns.map(function (c) {
    return {
      title: (c.sTitle || '').replace(/\\s+/g, ' ').trim(),
      data: c.mData,
      visible: c.bVisible !== false,
    };
  });
}
"""

_JS_CONSULTAR = """
async (args) => {
  const t = args.t, cuit = args.cuit;
  const base = location.origin + '/mcmp/jsp/ajax.do?f=';
  const pad = (n) => String(n).padStart(2, '0');
  const fmt = (d) => pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear();
  const parse = (s) => { const p = s.split('/'); return new Date(+p[2], +p[1] - 1, +p[0]); };
  let ini = parse(args.fd);
  const fin = parse(args.fh);
  const out = [];
  let guard = 0;
  while (ini <= fin && guard < 12) {
    guard++;
    let cf = new Date(ini.getFullYear(), 11, 31);
    if (cf > fin) cf = fin;
    const rango = fmt(ini) + ' - ' + fmt(cf);
    const u1 = base + 'generarConsulta&t=' + t + '&fechaEmision=' +
      encodeURIComponent(rango) + '&tiposComprobantes=&cuitConsultada=' + cuit;
    let id = null;
    try {
      const r1 = await fetch(u1, { headers: { 'X-Requested-With': 'XMLHttpRequest' } });
      const j1 = await r1.json();
      id = j1 && j1.datos && j1.datos.idConsulta;
    } catch (e) { return { error: 'generarConsulta: ' + e }; }
    if (id) {
      try {
        const r2 = await fetch(base + 'listaResultados&id=' + id,
          { headers: { 'X-Requested-With': 'XMLHttpRequest' } });
        const j2 = await r2.json();
        const data = (j2 && j2.datos && j2.datos.data) || [];
        for (const row of data) out.push(row);
      } catch (e) { return { error: 'listaResultados: ' + e }; }
    }
    ini = new Date(cf.getFullYear(), cf.getMonth(), cf.getDate() + 1);
  }
  return { rows: out };
}
"""


def _leer_columnas_datatable(mc):
    """Lee la configuración de columnas (título + índice de dato) de la DataTable."""
    for ctx in _iter_contextos(mc):
        try:
            cols = ctx.evaluate(_JS_LEER_COLUMNAS)
        except Exception:
            cols = None
        if cols:
            return cols
    return None


def _columnas_datatable_ok(cols) -> bool:
    """Valida que la tabla traiga las columnas que el procesador necesita
    (incluye el desglose por alícuota). Si no, conviene usar la descarga oficial."""
    titulos = " | ".join((c.get("title") or "").lower() for c in cols)
    tiene_alicuota = "iva 21" in titulos
    requeridas = ("imp. total" in titulos and "tipo" in titulos and "fecha" in titulos)
    return tiene_alicuota and requeridas


def _consultar_json_mcmp(mc, t: str, fd: str, fh: str, cuit: str):
    """Llama generarConsulta + listaResultados vía fetch (mismo origen, con sesión)."""
    args = {"t": t, "fd": fd, "fh": fh, "cuit": re.sub(r"\D", "", cuit)}
    for ctx in _iter_contextos(mc):
        try:
            res = ctx.evaluate(_JS_CONSULTAR, args)
        except Exception:
            res = None
        if isinstance(res, dict) and "rows" in res:
            return res["rows"]
    raise AutomatizacionArcaError("La consulta JSON de Mis Comprobantes no respondió.")


def _construir_xlsx_desde_json(cols, rows) -> bytes:
    """Reconstruye el Excel oficial usando los títulos y el índice de dato de la tabla."""
    import io as _io

    from openpyxl import Workbook

    visibles = [c for c in cols if c.get("visible", True) and (c.get("title") or "").strip()]
    headers = [c["title"] for c in visibles]
    idxs = [c.get("data") for c in visibles]

    wb = Workbook()
    ws = wb.active
    ws.title = "Comprobantes"
    ws.append(headers)
    for fila in rows:
        valores = []
        for di in idxs:
            try:
                valores.append(fila[di] if isinstance(di, int) and di < len(fila) else "")
            except Exception:
                valores.append("")
        ws.append(valores)
    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _descargar_tipo_en_sesion(
    mc,
    cuit_repr: str,
    fd: str,
    fh: str,
    tipo: Literal["emitidos", "recibidos"],
    *,
    elegir_perfil: bool,
    cuit_login: str | None = None,
    on_paso=None,
) -> tuple[bytes, str]:
    if elegir_perfil:
        _paso(on_paso, "perfil", "en_curso")
        _elegir_perfil_representado(mc, cuit_repr, cuit_login=cuit_login)
        _paso(on_paso, "perfil", "ok")
    _paso(on_paso, tipo, "en_curso")
    _ir_a_pantalla_tipo(mc, tipo)

    # Vía rápida (API JSON interna de MCMP). Si la tabla no expone todas las
    # columnas requeridas o algo falla, se usa la descarga oficial (UI).
    try:
        cols = _leer_columnas_datatable(mc)
        if cols and _columnas_datatable_ok(cols):
            cuit_consulta = _cuit_activo_mcmp(mc) or _normalizar_cuit_busqueda(cuit_repr)
            t = "E" if tipo == "emitidos" else "R"
            rows = _consultar_json_mcmp(mc, t, fd, fh, cuit_consulta)
            if not rows:
                raise SinComprobantesError(
                    "La consulta no devolvió comprobantes para el período indicado."
                )
            data = _construir_xlsx_desde_json(cols, rows)
            if len(data) >= 800:
                return data, f"mis_comprobantes_{tipo}.xlsx"
    except SinComprobantesError:
        raise
    except Exception:
        pass

    # Respaldo: método UI (descarga oficial del botón Excel de DataTables).
    estado = _aplicar_filtro_fechas_y_buscar(mc, fd, fh)
    return _descargar_excel_o_csv(mc, estado)


def _flujo_post_login(
    mc,
    cuit_repr: str,
    cuit_login: str,
    fd: str,
    fh: str,
    tipo: TipoComprobantes,
    on_paso=None,
) -> DescargaArcaResult:
    if tipo == "ambos":
        avisos: list[str] = []
        razon_social = ""
        try:
            data_e, nom_e = _descargar_tipo_en_sesion(
                mc, cuit_repr, fd, fh, "emitidos",
                elegir_perfil=True, cuit_login=cuit_login, on_paso=on_paso,
            )
            emitidos = (data_e, nom_e)
            _paso(on_paso, "emitidos", "ok")
        except SinComprobantesError:
            emitidos = None
            avisos.append("Emitidos: sin comprobantes en el período")
            _paso(on_paso, "emitidos", "ok")
        if not razon_social:
            try:
                razon_social = _razon_social_activa_mcmp(mc)
            except Exception:
                razon_social = ""
        pausa_humana(0.4, 0.8)
        try:
            data_r, nom_r = _descargar_tipo_en_sesion(
                mc, cuit_repr, fd, fh, "recibidos",
                elegir_perfil=False, cuit_login=cuit_login, on_paso=on_paso,
            )
            recibidos = (data_r, nom_r)
            _paso(on_paso, "recibidos", "ok")
        except SinComprobantesError:
            recibidos = None
            avisos.append("Recibidos: sin comprobantes en el período")
            _paso(on_paso, "recibidos", "ok")
        except AutomatizacionArcaError as exc:
            recibidos = None
            avisos.append(f"falló Recibidos: {exc}")
            _paso(on_paso, "recibidos", "error")

        if not razon_social:
            try:
                razon_social = _razon_social_activa_mcmp(mc)
            except Exception:
                razon_social = ""

        if emitidos is None and recibidos is None:
            raise SinComprobantesError(
                "; ".join(avisos) or "Sin comprobantes en el período."
            )
        return DescargaArcaResult(
            emitidos=emitidos,
            recibidos=recibidos,
            aviso_parcial="; ".join(avisos) if avisos else None,
            razon_social=razon_social or None,
        )

    data, nom = _descargar_tipo_en_sesion(
        mc, cuit_repr, fd, fh, tipo,
        elegir_perfil=True, cuit_login=cuit_login, on_paso=on_paso,
    )
    _paso(on_paso, tipo, "ok")
    try:
        razon_social = _razon_social_activa_mcmp(mc)
    except Exception:
        razon_social = ""
    res = DescargaArcaResult.simple(data, nom, emitidos=(tipo == "emitidos"))
    if razon_social:
        res = DescargaArcaResult(
            emitidos=res.emitidos,
            recibidos=res.recibidos,
            aviso_parcial=res.aviso_parcial,
            razon_social=razon_social,
        )
    return res


def ejecutar_descarga_mis_comprobantes(
    cred: CredencialesArca,
    fecha_desde: date,
    fecha_hasta: date,
    *,
    headless: bool = True,
    tipo: TipoComprobantes = "emitidos",
    on_paso=None,
) -> DescargaArcaResult:
    if not _playwright_disponible():
        raise AutomatizacionNoDisponibleError(
            "Playwright no está instalado. En local: pip install playwright && playwright install chromium"
        )

    from playwright.sync_api import TimeoutError as PlaywrightTimeout
    from playwright.sync_api import sync_playwright

    fd, fh = _formatear_rango_afip(fecha_desde, fecha_hasta)
    cuit_repr = _normalizar_cuit_busqueda(cred.cuit_representado)
    cuit_login = _normalizar_cuit_busqueda(cred.cuit_login)

    browser = None
    try:
        with sync_playwright() as p:
            browser, context = _nuevo_contexto_stealth(p, headless=headless)
            page = context.new_page()
            page.set_default_timeout(60_000)

            _paso(on_paso, "login", "en_curso")
            page.goto(LOGIN_URL, wait_until="domcontentloaded")
            pausa_humana(0.56, 1.26)
            _llenar_cuit_y_avanzar(page, cred.cuit_login)
            _login_clave_fiscal(page, cred.clave_fiscal, cred.cuit_login)
            _paso(on_paso, "login", "ok")
            _paso(on_paso, "mis_comprobantes", "en_curso")
            mc = _abrir_mis_comprobantes(page)
            _paso(on_paso, "mis_comprobantes", "ok")
            return _flujo_post_login(
                mc, cuit_repr, cuit_login, fd, fh, tipo, on_paso=on_paso
            )

    except LoginArcaError:
        raise
    except CuitRepresentadoNoEncontradoError:
        raise
    except SinComprobantesError:
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
