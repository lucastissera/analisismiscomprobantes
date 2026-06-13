"""Automatización del servicio «Nuestra Parte» de ARCA (cgpf-webv2-fe.arca.gob.ar).

Flujo:
1. Login con clave fiscal (reutiliza el login de Mis Comprobantes).
2. Abre «Nuestra Parte» desde el portal (SSO).
3. Si el CUIT representado difiere del login: pantalla «Selección representado»
   (``#/seleccionRepresentado``) → tarjeta ``e-custom-card`` con el CUIT.
4. Entra a «Tu información» y selecciona el ejercicio pedido en el desplegable.
5. Recorre las 4 secciones:
   - Facturación, acreditaciones e ingresos
   - Información patrimonial
   - Declaraciones juradas presentadas
   - Inversiones y participaciones
   En cada una expande los acordeones con datos y exporta la grilla
   (XLSX preferido, si no CSV). En Información patrimonial además imprime el
   PDF de la pantalla y entra a cada «Ver detalles» (ícono de ojo) para exportar.
6. Guarda todo en «Nuestra Parte yyyy-mm-dd / <cuit> / <sección>».

En servidor web usa navegador headless; en el portable (.exe), visible por defecto
(``CUIT_EN_ARCA_HEADLESS=1`` fuerza headless también en el .exe).
"""

from __future__ import annotations

import re
import random
import time
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from cuit_en_arca.credenciales import CredencialesArca
from cuit_en_arca.errores import (
    AutomatizacionArcaError,
    AutomatizacionNoDisponibleError,
    CuitRepresentadoNoEncontradoError,
)
from cuit_en_arca.service import _headless_desde_env
from cuit_en_arca.stealth import clic_humano, escribir_como_humano, pausa_humana

NP_TERMINO_BUSQUEDA = "Nuestra Parte"
NP_HOST = "cgpf"  # cgpf-webv2-fe.arca.gob.ar
NP_BASE = "https://cgpf-webv2-fe.arca.gob.ar/#"

# Secciones (texto del encabezado de la tarjeta en «Tu información»).
SECCIONES = (
    ("facturacion", "Facturación, acreditaciones e ingresos"),
    ("patrimonial", "Información patrimonial"),
    ("ddjj", "Declaraciones juradas presentadas"),
    ("inversiones", "Inversiones y participaciones"),
)

# Textos de secciones principales (excluir al buscar subitems).
_NOMBRES_SECCION = tuple(n for _, n in SECCIONES)

# Cada CUIT puede tener distinta cantidad de ítems; tope solo anti-bucle infinito.
_NP_TOPE_ITEMS = 120
_NP_ESPERA_DOM_MS = 2000

# Ritmo entre acciones (segundos): facturación acordeón ~5 s; resto ~7–10 s.
_NP_RITMO_FACTURACION_ACORDEON = (4.8, 5.2)
_NP_RITMO_GENERAL = (7.0, 10.0)
_np_ultima_accion: float = 0.0


def _reset_ritmo_np() -> None:
    global _np_ultima_accion
    _np_ultima_accion = 0.0


def _marcar_accion_np() -> None:
    global _np_ultima_accion
    _np_ultima_accion = time.monotonic()


def _esperar_ritmo_np(modo: str = "general") -> None:
    """Espera lo necesario para respetar el intervalo mínimo desde la acción anterior."""
    global _np_ultima_accion
    if _np_ultima_accion <= 0:
        return
    lo, hi = (
        _NP_RITMO_FACTURACION_ACORDEON
        if modo == "facturacion_acordeon"
        else _NP_RITMO_GENERAL
    )
    objetivo = random.uniform(lo, hi)
    faltante = objetivo - (time.monotonic() - _np_ultima_accion)
    if faltante > 0:
        time.sleep(faltante)


def _esperar_np(np, timeout_ms: int | None = None) -> None:
    """Espera breve de DOM tras navegar (sin networkidle)."""
    limite = timeout_ms or _NP_ESPERA_DOM_MS
    try:
        np.wait_for_load_state("domcontentloaded", timeout=limite)
    except Exception:
        pass


def _cantidad(loc, tope: int | None = None) -> int:
    """Cuenta elementos del DOM respetando un tope de seguridad."""
    limite = tope or _NP_TOPE_ITEMS
    try:
        return min(loc.count(), limite)
    except Exception:
        return 0


@dataclass
class _EstadoScrollNP:
    """Evita bucles de scroll: como máximo una pasada hasta el fondo por vista."""

    fondo_alcanzado: bool = False


_np_scroll_por_pagina: dict[int, _EstadoScrollNP] = {}


def _estado_scroll(np) -> _EstadoScrollNP:
    clave = id(np)
    if clave not in _np_scroll_por_pagina:
        _np_scroll_por_pagina[clave] = _EstadoScrollNP()
    return _np_scroll_por_pagina[clave]


def _reset_scroll_np(np) -> None:
    _np_scroll_por_pagina.pop(id(np), None)


def _en_fondo_pagina(np) -> bool:
    try:
        return bool(
            np.evaluate(
                """
                () => {
                  const doc = document.scrollingElement || document.documentElement;
                  const okDoc = doc.scrollTop + window.innerHeight >= doc.scrollHeight - 8;
                  const grids = Array.from(
                    document.querySelectorAll('.e-gridcontent, .e-content, [class*="gridcontent"]')
                  );
                  if (!grids.length) return okDoc;
                  const okGrids = grids.every(
                    (g) => g.scrollHeight <= g.clientHeight + 4
                      || g.scrollTop + g.clientHeight >= g.scrollHeight - 8
                  );
                  return okDoc && okGrids;
                }
                """
            )
        )
    except Exception:
        return True


def _revelar_contenido_np(np, on_log=None) -> None:
    """Desplaza página y grillas **una vez** hasta el fondo; no repite si ya llegó."""
    est = _estado_scroll(np)
    if est.fondo_alcanzado or _en_fondo_pagina(np):
        est.fondo_alcanzado = True
        return
    for paso in range(18):
        if _en_fondo_pagina(np):
            est.fondo_alcanzado = True
            return
        try:
            movio = np.evaluate(
                """
                () => {
                  let moved = false;
                  const doc = document.scrollingElement || document.documentElement;
                  const y0 = doc.scrollTop;
                  doc.scrollBy(0, Math.max(window.innerHeight * 0.88, 420));
                  if (doc.scrollTop > y0) moved = true;
                  document.querySelectorAll(
                    '.e-gridcontent, .e-content, [class*="gridcontent"]'
                  ).forEach((g) => {
                    const t0 = g.scrollTop;
                    g.scrollBy(0, Math.max(g.clientHeight * 0.88, 200));
                    if (g.scrollTop > t0) moved = true;
                  });
                  return moved;
                }
                """
            )
        except Exception:
            est.fondo_alcanzado = True
            return
        if not movio or _en_fondo_pagina(np):
            est.fondo_alcanzado = True
            return
        pausa_humana(0.1, 0.18)
    est.fondo_alcanzado = True
    if on_log:
        _log(on_log, "  • Scroll: fondo alcanzado (no se repetirá en esta vista).")


def _click_np(np, loc, *, intentar_revelar: bool = False) -> None:
    """Clic sin bucles de scroll. Tras llegar al fondo, siempre force=True."""
    est = _estado_scroll(np)
    try:
        if loc.is_visible(timeout=400):
            loc.click(force=True, timeout=3000)
            return
    except Exception:
        pass
    if est.fondo_alcanzado:
        loc.click(force=True, timeout=3000)
        return
    if intentar_revelar:
        _revelar_contenido_np(np)
        try:
            if loc.is_visible(timeout=400):
                loc.click(force=True, timeout=3000)
                return
        except Exception:
            pass
    loc.click(force=True, timeout=3000)


def _indices_visibles_en_pantalla(np, loc) -> list[int]:
    """Índices de elementos visibles en el viewport (sin scroll)."""
    total = _cantidad(loc)
    return [i for i in range(total) if _visible(loc.nth(i))]


def _anios_en_opcion(texto: str) -> list[str]:
    """Años detectados en el texto de una opción del desplegable (fechas DD/MM/AAAA)."""
    t = (texto or "").strip()
    if not t:
        return []
    anos: list[str] = []
    for m in re.finditer(r"\d{1,2}[/\-]\d{1,2}[/\-](\d{4})", t):
        anos.append(m.group(1))
    if not anos:
        anos.extend(re.findall(r"\b(\d{4})\b", t))
    return anos


def _opcion_contiene_anio(texto: str, anio: str) -> bool:
    """True si la opción incluye el año pedido (p. ej. «al 31/12/2025» para «2025»)."""
    if not anio or not texto:
        return False
    return anio in _anios_en_opcion(texto)


def _puntuacion_opcion_ejercicio(texto: str, anio: str) -> int:
    """Mayor puntuación = mejor coincidencia cuando hay varias opciones con el mismo año."""
    if not _opcion_contiene_anio(texto, anio):
        return -1
    score = 0
    if re.search(rf"31[/\-]12[/\-]{re.escape(anio)}", texto):
        score += 100
    if re.search(rf"31[/\-]03[/\-]{re.escape(anio)}", texto):
        score += 90
    if re.search(rf"\bal\s+.*{re.escape(anio)}", texto, re.I):
        score += 80
    if re.search(rf"\b{re.escape(anio)}\b", texto):
        score += 10
    return score


def _mejor_opcion_ejercicio(textos: list[str], anio: str, ejercicio_raw: str) -> str | None:
    """Elige la opción del desplegable que contiene el año (no hace falta coincidencia exacta)."""
    limpios = [t.strip() for t in textos if (t or "").strip()]
    if not limpios:
        return None
    if anio:
        con_anio = [t for t in limpios if _opcion_contiene_anio(t, anio)]
        if not con_anio:
            return None
        return max(con_anio, key=lambda t: _puntuacion_opcion_ejercicio(t, anio))
    raw = ejercicio_raw.strip()
    for t in limpios:
        if raw.lower() in t.lower():
            return t
    return None


def _opcion_coincide_ejercicio(txt: str, objetivo: str, anio: str) -> bool:
    if not txt:
        return False
    if objetivo and (txt == objetivo or objetivo in txt or txt in objetivo):
        return True
    if anio and _opcion_contiene_anio(txt, anio):
        if not objetivo:
            return True
        return _puntuacion_opcion_ejercicio(txt, anio) >= _puntuacion_opcion_ejercicio(
            objetivo, anio
        )
    return False


def _click_sin_scroll(loc) -> None:
    """Clic sin scroll_into_view (Playwright no desplaza la página)."""
    try:
        loc.click(force=True, timeout=3000)
    except Exception:
        try:
            loc.dispatch_event("click")
        except Exception:
            pass


def _en_hub_tu_informacion(np) -> bool:
    """True si estamos en el hub «Tu información» (4 tarjetas), no dentro de una sección."""
    u = np.url or ""
    if re.search(r"tu-informacion/[^/?#]+", u, re.I):
        return False
    return bool(re.search(r"tu-informacion/?(?:\?|#|$)", u, re.I))


def _leer_ejercicio_visible(np) -> str:
    """Texto del ejercicio en el multiselect de NP (p. ej. «al 31/12/2024»)."""
    for sel in ("div.multiselect-wrapper", "div[class*='multiselect-wrapper']"):
        try:
            wrap = np.locator(sel).first
            if wrap.count():
                t = (wrap.inner_text(timeout=800) or "").strip()
                m = re.search(r"\bal\s+\d{1,2}/\d{1,2}/\d{4}", t, re.I)
                if m:
                    return m.group(0)
        except Exception:
            continue
    for sel in (
        ".e-multi-select-wrapper input",
        ".e-ddl input",
        ".e-input-group.e-control-wrapper input",
    ):
        try:
            loc = np.locator(sel).first
            if loc.count() and _visible(loc):
                v = (loc.input_value(timeout=800) or "").strip()
                if v and re.search(r"\d{4}", v):
                    return v
        except Exception:
            continue
    try:
        t = np.get_by_text(re.compile(r"\bal\s+\d{1,2}/\d{1,2}/\d{4}", re.I)).first
        if t.count():
            return (t.inner_text(timeout=1000) or "").strip()
    except Exception:
        pass
    return ""


def _popup_ejercicio_visible(np) -> bool:
    for sel in (
        "ul[id*='multiselect-options']",
        "[id*='multiselect-options']",
        ".e-popup-open",
        ".e-ddl.e-popup",
        ".e-dropdownbase.e-popup",
    ):
        try:
            pop = np.locator(sel).first
            if pop.count() and pop.is_visible(timeout=500):
                return True
        except Exception:
            continue
    return False


def _abrir_desplegable_ejercicio(np) -> bool:
    """Abre el desplegable de ejercicio (multiselect NP: caret expand_more)."""
    if _popup_ejercicio_visible(np):
        return True
    for sel in (
        "div.multiselect-wrapper span[class*='caret']",
        "div.multiselect-wrapper span.material-symbols-rounded",
        "div[class*='multiselect-wrapper']",
    ):
        try:
            loc = np.locator(sel).first
            if loc.count() and _visible(loc):
                _click_sin_scroll(loc)
                pausa_humana(0.25, 0.45)
                if _popup_ejercicio_visible(np):
                    return True
        except Exception:
            continue
    # Respaldo Syncfusion legacy
    for sel in (
        "span.e-input-group-icon.e-ddl-icon",
        ".e-multi-select-wrapper .e-input-group-icon",
        ".icon-multiselect",
    ):
        try:
            icon = np.locator(sel).first
            if icon.count() and _visible(icon):
                _click_sin_scroll(icon)
                pausa_humana(0.25, 0.45)
                if _popup_ejercicio_visible(np):
                    return True
        except Exception:
            continue
    return False


def _cerrar_popup_ejercicio(np) -> None:
    """Cierra el popup del ejercicio sin Escape (no revierte la selección)."""
    if not _popup_ejercicio_visible(np):
        return
    for sel in ("div.contenedor", "main", "h1", "div.item"):
        try:
            loc = np.locator(sel).first
            if loc.count() and _visible(loc):
                loc.click(position={"x": 12, "y": 12}, timeout=1200)
                pausa_humana(0.12, 0.25)
                break
        except Exception:
            continue
    try:
        np.locator("ul[id*='multiselect-options']").first.wait_for(state="hidden", timeout=3000)
    except Exception:
        pass


def _locator_opciones_ejercicio(np):
    return np.locator(
        "ul[id*='multiselect-options'] li, "
        "li[id*='multiselect-option'], "
        ".e-popup-open .e-list-item, .e-ddl.e-popup .e-list-item, "
        "[role='listbox'] [role='option']"
    )


def _click_opcion_ejercicio(np, o, txt: str, anio: str = "", on_log=None) -> bool:
    """Selecciona una opción del periodo (1 clic; 2.º solo si no aplicó)."""
    try:
        _click_sin_scroll(o)
    except Exception:
        try:
            o.dispatch_event("click")
        except Exception:
            return False
    pausa_humana(0.35, 0.55)
    year = anio
    if not year:
        ym = re.search(r"(\d{4})", txt or "")
        if ym:
            year = ym.group(1)
    if year and not _ejercicio_aplicado_ok(np, year, txt or ""):
        try:
            _click_sin_scroll(o)
            pausa_humana(0.28, 0.42)
        except Exception:
            pass
    _cerrar_popup_ejercicio(np)
    pausa_humana(0.3, 0.5)
    _esperar_np(np)
    if year:
        return _ejercicio_aplicado_ok(np, year, txt or "")
    return bool(_leer_ejercicio_visible(np))


def _ejercicio_aplicado_ok(np, anio: str, objetivo: str) -> bool:
    actual = _leer_ejercicio_visible(np)
    if not actual:
        return False
    if anio:
        return _opcion_contiene_anio(actual, anio)
    return objetivo.lower() in actual.lower() or actual.lower() in objetivo.lower()


def _listar_opciones_ejercicio(np) -> list[str]:
    """Lista opciones visibles del desplegable (sin scroll artificial)."""
    resultado: list[str] = []
    try:
        raw = np.evaluate(
            """
            () => Array.from(document.querySelectorAll(
              "ul[id*='multiselect-options'] li, li[id*='multiselect-option'], "
              + ".e-popup-open .e-list-item, [role='listbox'] [role='option']"
            )).map(el => (el.getAttribute('aria-label') || el.innerText || '').trim())
              .filter(Boolean)
            """
        )
        if isinstance(raw, list):
            vistos: set[str] = set()
            for x in raw:
                t = str(x).strip()
                if t and t not in vistos:
                    vistos.add(t)
                    resultado.append(t)
    except Exception:
        pass
    return resultado


def _es_texto_seccion_principal(texto: str) -> bool:
    t = re.sub(r"\s+", " ", (texto or "").strip()).lower()
    if not t:
        return True
    for nombre in _NOMBRES_SECCION:
        if nombre.lower() in t or t in nombre.lower():
            return True
    return bool(re.search(r"tu informaci|exportar|^\s*volver\s*$", t, re.I))


@dataclass
class SeccionNP:
    clave: str
    nombre: str
    archivos: list[str] = field(default_factory=list)
    nota: str = ""


@dataclass
class ResultadoNPCuit:
    cuit_login: str
    cuit_representado: str
    razon_social: str | None
    ejercicio: str
    carpeta: str
    secciones: list[SeccionNP] = field(default_factory=list)
    error: str | None = None

    @property
    def total_archivos(self) -> int:
        return sum(len(s.archivos) for s in self.secciones)


def _playwright_disponible() -> bool:
    try:
        import playwright  # noqa: F401

        return True
    except Exception:
        return False


def _log(on_log, msg: str) -> None:
    if on_log:
        try:
            on_log(msg)
        except Exception:
            pass


def _paso(on_paso, clave: str, estado: str) -> None:
    if on_paso:
        try:
            on_paso(clave, estado)
        except Exception:
            pass


# --------------------------------------------------------------------------- #
# Carpeta de destino
# --------------------------------------------------------------------------- #
def carpeta_np_base(
    hoy: date | None = None,
    base_elegida: str | Path | None = None,
    *,
    nombre_sesion: str | None = None,
) -> Path:
    from cuit_en_arca.dfe_automation import _escritorio_windows
    from cuit_en_arca.carpetas_salida import stamp_carpeta_ejecucion

    if nombre_sesion:
        nombre = nombre_sesion
    else:
        nombre = f"Nuestra Parte {stamp_carpeta_ejecucion(datetime.combine(hoy, datetime.now().time()) if hoy else None)}"
    if base_elegida:
        destino = Path(base_elegida) / nombre
        destino.mkdir(parents=True, exist_ok=True)
        return destino
    home = Path.home()
    escritorio = _escritorio_windows()
    if escritorio is None:
        candidatos = [
            home / "Desktop",
            home / "Escritorio",
            home / "OneDrive" / "Desktop",
            home / "OneDrive" / "Escritorio",
        ]
        escritorio = next((c for c in candidatos if c.is_dir()), None)
    destino = (escritorio or home) / nombre
    destino.mkdir(parents=True, exist_ok=True)
    return destino


def _nombre_seguro(nombre: str, *, fallback: str = "archivo") -> str:
    from cuit_en_arca.dfe_automation import _nombre_seguro as ns

    return ns(nombre, fallback=fallback)


# --------------------------------------------------------------------------- #
# Apertura del servicio Nuestra Parte
# --------------------------------------------------------------------------- #
def _abrir_nuestra_parte(page):
    from cuit_en_arca.automation_playwright import (
        _click_servicio_y_obtener_pagina,
        _esperar_pagina,
        _esperar_post_login,
        _iter_contextos,
        _locator_buscador_servicios,
    )

    pausa_humana(0.8, 1.4)
    _esperar_post_login(page)
    _esperar_pagina(page, timeout=42_000)

    objetivo = None
    buscador = None
    for ctx in _iter_contextos(page):
        buscador = _locator_buscador_servicios(ctx)
        if buscador is not None:
            break
    if buscador is not None:
        escribir_como_humano(buscador, NP_TERMINO_BUSQUEDA)
        pausa_humana(1.0, 1.8)
        res = page.get_by_text(re.compile(r"Nuestra Parte", re.I))
        for i in range(_cantidad(res, 12)):
            try:
                if res.nth(i).is_visible(timeout=900):
                    objetivo = _click_servicio_y_obtener_pagina(page, res.nth(i))
                    break
            except Exception:
                continue

    pausa_humana(2.2, 3.2)
    np = None
    for pg in page.context.pages:
        if NP_HOST in (pg.url or "").lower():
            np = pg
            break
    if np is None:
        np = objetivo or page.context.pages[-1]
    try:
        np.bring_to_front()
        np.set_default_timeout(40_000)
        _esperar_np(np, 12_000)
    except Exception:
        pass
    pausa_humana(1.5, 2.5)
    if NP_HOST not in (np.url or "").lower():
        raise AutomatizacionArcaError(
            "No se pudo abrir «Nuestra Parte» desde el portal. "
            "Verificá que el CUIT tenga el servicio habilitado."
        )
    return np


def _cuit_np_n(s: str) -> str:
    from cuit_en_arca.automation_playwright import _normalizar_cuit_busqueda

    return _normalizar_cuit_busqueda(s)


def _cuit_np_fmt(cuit_n: str) -> str:
    if len(cuit_n) != 11:
        return cuit_n
    return f"{cuit_n[:2]}-{cuit_n[2:10]}-{cuit_n[10]}"


def _en_seleccion_representado_np(np) -> bool:
    u = (np.url or "").lower().replace("-", "").replace("_", "")
    return "seleccionrepresentado" in u


def _seleccionar_representado_np(
    np,
    cuit_repr: str,
    *,
    cuit_login: str | None = None,
    on_log=None,
) -> None:
    """Elige el CUIT en ``#/seleccionRepresentado`` (tarjetas ``e-custom-card``)."""
    cuit_n = _cuit_np_n(cuit_repr)
    if cuit_login and _cuit_np_n(cuit_login) == cuit_n:
        return

    pausa_humana(0.6, 1.0)
    if not _en_seleccion_representado_np(np):
        try:
            np.wait_for_url(re.compile(r"seleccionRepresentado", re.I), timeout=10_000)
        except Exception:
            if not _en_seleccion_representado_np(np):
                _log(on_log, "NP: sin pantalla de selección de representado.")
                return

    fmt = _cuit_np_fmt(cuit_n)
    _log(on_log, f"NP: seleccionando representado {fmt}…")

    resultado = np.evaluate(
        """
        ([cuit, fmt]) => {
          const norm = (t) => (t || "").replace(/\\D/g, "");
          const cards = document.querySelectorAll('[id*="e-custom-card"], .e-custom-card');
          for (const card of cards) {
            const t = (card.innerText || card.textContent || "").trim();
            if (!t) continue;
            const n = norm(t);
            if (n.includes(cuit) || t.includes(fmt)) {
              card.click();
              return { ok: true, label: t.replace(/\\s+/g, " ").slice(0, 120) };
            }
          }
          return { ok: false };
        }
        """,
        [cuit_n, fmt],
    )

    if isinstance(resultado, dict) and resultado.get("ok"):
        _log(on_log, f"NP representado: {resultado.get('label') or fmt}.")
        pausa_humana(0.9, 1.5)
        try:
            np.wait_for_function(
                "() => !location.hash.toLowerCase().includes('seleccionrepresentado')",
                timeout=15_000,
            )
        except Exception:
            pass
        _esperar_np(np, 4000)
        return

    cards = np.locator('[id*="e-custom-card"], .e-custom-card')
    for i in range(min(cards.count(), 40)):
        card = cards.nth(i)
        try:
            texto = (card.inner_text(timeout=600) or "").strip()
        except Exception:
            continue
        n = re.sub(r"\D", "", texto)
        if cuit_n in n or fmt in texto:
            _click_np(np, card)
            _log(on_log, f"NP representado: {texto[:80]}.")
            pausa_humana(0.9, 1.5)
            _esperar_np(np, 4000)
            return

    raise CuitRepresentadoNoEncontradoError(
        f"El CUIT representado {fmt} no aparece en «Nuestra Parte» (selección representado)."
    )


def _leer_razon_social(np) -> str:
    try:
        js = r"""
        () => {
          const el = document.querySelector('.e-navbar, header, .navbar, body');
          const t = (el && (el.innerText||'')) || '';
          const m = t.match(/([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ .,'-]{4,})\s*\[?\d{2}-?\d{8}-?\d/);
          return m ? m[1].trim() : '';
        }
        """
        return (np.evaluate(js) or "").strip()
    except Exception:
        return ""


def _ir_tu_informacion(np) -> None:
    _reset_scroll_np(np)
    if _en_hub_tu_informacion(np):
        return
    try:
        np.goto(f"{NP_BASE}/tu-informacion", wait_until="domcontentloaded")
    except Exception:
        try:
            ti = np.get_by_text(re.compile(r"Tu informaci[oó]n", re.I)).first
            if ti.count():
                _click_np(np, ti)
        except Exception:
            pass
    pausa_humana(0.5, 0.9)
    _esperar_np(np)


def _volver_a_tu_informacion(np, on_log=None) -> None:
    """Vuelve al hub de las 4 tarjetas sin recargar ni re-seleccionar ejercicio."""
    if _en_hub_tu_informacion(np):
        return
    for _ in range(4):
        if _en_hub_tu_informacion(np):
            return
        _volver(np)
        pausa_humana(0.25, 0.4)
    if not _en_hub_tu_informacion(np):
        try:
            np.goto(f"{NP_BASE}/tu-informacion", wait_until="domcontentloaded")
            _esperar_np(np, 2000)
        except Exception as exc:
            _log(on_log, f"No se pudo volver a Tu información: {exc}")


def _seleccionar_ejercicio(np, ejercicio: str, on_log=None) -> str:
    """Elige el periodo en el multiselect NP (p. ej. «2024» → «al 31/12/2024»)."""
    ejercicio = (ejercicio or "").strip()
    seleccionado = _leer_ejercicio_visible(np)
    if not ejercicio:
        return seleccionado

    anio = ""
    m = re.search(r"(\d{4})", ejercicio)
    if m:
        anio = m.group(1)

    if anio and _ejercicio_aplicado_ok(np, anio, ejercicio):
        _log(on_log, f"Ejercicio ya activo: {seleccionado or anio}")
        return seleccionado or ejercicio

    if not _abrir_desplegable_ejercicio(np):
        _log(on_log, "No se pudo abrir el desplegable de ejercicio.")
        return seleccionado or ejercicio

    try:
        np.locator("ul[id*='multiselect-options'], .e-popup-open").first.wait_for(
            state="visible", timeout=5000
        )
    except Exception:
        if not _popup_ejercicio_visible(np):
            _abrir_desplegable_ejercicio(np)

    textos = _listar_opciones_ejercicio(np)
    objetivo = _mejor_opcion_ejercicio(textos, anio, ejercicio)
    if objetivo:
        _log(on_log, f"Opción objetivo para «{ejercicio}»: {objetivo}")

    elegida = False

    if anio:
        for sel in (
            f"li[id*='multiselect-option-{anio}']",
            f"li[aria-label*='{anio}']",
        ):
            try:
                direct = np.locator(sel).first
                if direct.count() and _visible(direct):
                    txt = (direct.get_attribute("aria-label") or direct.inner_text(timeout=500) or "").strip()
                    if _click_opcion_ejercicio(np, direct, txt or objetivo or "", anio, on_log):
                        seleccionado = _leer_ejercicio_visible(np) or txt or objetivo or ""
                        elegida = True
                        _log(on_log, f"Ejercicio seleccionado: {seleccionado}")
                        break
            except Exception:
                continue
        if elegida:
            return seleccionado or ejercicio

    opciones = _locator_opciones_ejercicio(np)
    total = _cantidad(opciones)

    for i in range(total):
        o = opciones.nth(i)
        try:
            if not _visible(o):
                continue
            txt = (o.get_attribute("aria-label") or o.inner_text(timeout=500) or "").strip()
        except Exception:
            continue
        if not _opcion_coincide_ejercicio(txt, objetivo or "", anio):
            continue
        if _click_opcion_ejercicio(np, o, txt or objetivo or "", anio, on_log):
            seleccionado = _leer_ejercicio_visible(np) or txt or objetivo or ""
            elegida = True
            _log(on_log, f"Ejercicio seleccionado: {seleccionado}")
            break

    if not elegida and anio and _abrir_desplegable_ejercicio(np):
        try:
            opt = _locator_opciones_ejercicio(np).filter(has_text=re.compile(re.escape(anio)))
            for i in range(_cantidad(opt)):
                o = opt.nth(i)
                if not _visible(o):
                    continue
                txt = (o.get_attribute("aria-label") or o.inner_text(timeout=500) or "").strip()
                if not _opcion_contiene_anio(txt, anio):
                    continue
                if _click_opcion_ejercicio(np, o, txt, anio, on_log):
                    seleccionado = _leer_ejercicio_visible(np) or txt
                    elegida = True
                    _log(on_log, f"Ejercicio seleccionado (reintento): {seleccionado}")
                    break
        except Exception:
            pass

    if not elegida:
        _log(
            on_log,
            f"No se encontró el ejercicio «{ejercicio}» "
            f"(opciones: {', '.join(textos[:6])}{'…' if len(textos) > 6 else ''}); "
            "se usa el periodo visible.",
        )
        _cerrar_popup_ejercicio(np)
    else:
        seleccionado = _leer_ejercicio_visible(np) or seleccionado

    return seleccionado or ejercicio


# --------------------------------------------------------------------------- #
# Exportación de grillas / detalles
# --------------------------------------------------------------------------- #
def _primera_opcion_visible(np, sufijo_id: str, texto_re):
    """Devuelve un locator a la opción de formato VISIBLE (menú abierto)."""
    loc = np.locator(f"[id$='{sufijo_id}']")
    try:
        n = loc.count()
    except Exception:
        n = 0
    for i in range(n):
        if _visible(loc.nth(i)):
            return loc.nth(i)
    # Alternativa por texto visible.
    alt = np.get_by_text(texto_re)
    try:
        for i in range(_cantidad(alt, 12)):
            if _visible(alt.nth(i)):
                return alt.nth(i)
    except Exception:
        pass
    return None


def _exportar_menu_actual(np, dest: Path, prefijo: str, on_log=None) -> str:
    """Con un botón EXPORTAR ya abierto, descarga XLSX (o CSV)."""
    opt = _primera_opcion_visible(np, "format-XLSX", re.compile(r"\bXLSX\b|Excel", re.I))
    fmt = "xlsx"
    if opt is None:
        opt = _primera_opcion_visible(np, "format-CSV", re.compile(r"\bCSV\b", re.I))
        fmt = "csv"
    if opt is None:
        _log(on_log, "  • No se encontró opción de formato en el menú EXPORTAR.")
        return ""
    try:
        with np.expect_download(timeout=40_000) as di:
            _click_np(np, opt)
        d = di.value
        sug = d.suggested_filename or f"{prefijo}.{fmt}"
        ext = Path(sug).suffix or f".{fmt}"
        ruta = dest / (_nombre_seguro(prefijo, fallback=prefijo) + ext)
        k = 1
        while ruta.exists():
            ruta = dest / f"{_nombre_seguro(prefijo)}_{k}{ext}"
            k += 1
        d.save_as(str(ruta))
        _log(on_log, f"  • Exportado: {ruta.name}")
        _marcar_accion_np()
        return str(ruta)
    except Exception as exc:
        _log(on_log, f"  • No se pudo exportar «{prefijo}»: {exc}")
        return ""


def _exportar_grilla_visible(
    np, dest: Path, prefijo: str, on_log=None, *, solo_viewport: bool = False
) -> list[str]:
    """Abre cada botón EXPORTAR visible y descarga su grilla."""
    guardados: list[str] = []
    botones = np.locator("button[id$='export-button']")
    try:
        n = botones.count()
    except Exception:
        n = 0
    if n == 0:
        botones = np.locator("button:has-text('EXPORTAR')")
        try:
            n = botones.count()
        except Exception:
            n = 0
    revelado = False
    for i in range(n):
        b = botones.nth(i)
        if not _visible(b):
            if solo_viewport:
                if not revelado and i == 0:
                    _revelar_contenido_np(np, on_log)
                    revelado = True
                    if _visible(b):
                        pass
                    else:
                        continue
                else:
                    continue
            elif not revelado:
                _revelar_contenido_np(np, on_log)
                revelado = True
            if not _visible(b) and _estado_scroll(np).fondo_alcanzado:
                continue
        try:
            _click_np(np, b)
            pausa_humana(0.2, 0.35)
        except Exception:
            continue
        ruta = _exportar_menu_actual(np, dest, prefijo, on_log)
        if ruta:
            guardados.append(ruta)
    return guardados


def _visible(loc) -> bool:
    try:
        return loc.is_visible(timeout=600)
    except Exception:
        return False


def _exportar_pdf_pantalla(np, dest: Path, prefijo: str, on_log=None) -> str:
    from cuit_en_arca.dfe_automation import _png_a_pdf

    try:
        png = np.screenshot(full_page=False)
        ruta = dest / (_nombre_seguro(prefijo, fallback="pantalla")[:120] + ".pdf")
        _png_a_pdf(png, ruta)
        _log(on_log, f"  • Pantalla guardada en PDF: {ruta.name}")
        return str(ruta)
    except Exception as exc:
        _log(on_log, f"  • No se pudo guardar PDF de pantalla: {exc}")
        return ""


def _procesar_acordeones(
    np,
    dest: Path,
    prefijo_sec: str,
    on_log=None,
    *,
    modo_ritmo: str = "general",
    ver_detalles: bool = False,
) -> list[str]:
    """Procesa acordeones de a uno: expandir → exportar (→ opcional ver detalles).

    ``modo_ritmo=facturacion_acordeon`` → ~5 s entre exportaciones consecutivas.
    Otros modos → ~7–10 s entre acciones.
    """
    guardados: list[str] = []
    accs = np.locator("button.accordion-button")
    try:
        n = accs.count()
    except Exception:
        n = 0
    for i in range(n):
        if i > 0:
            _esperar_ritmo_np(modo_ritmo)
        b = accs.nth(i)
        try:
            titulo = (b.inner_text(timeout=800) or "").strip().replace("\n", " ")
        except Exception:
            titulo = f"{prefijo_sec}_{i+1}"
        titulo = re.sub(r"\s+", " ", titulo)[:80] or f"{prefijo_sec}_{i+1}"
        try:
            cls = b.get_attribute("class") or ""
            if "collapsed" in cls:
                _click_sin_scroll(b)
                try:
                    np.locator("button[id$='export-button']").first.wait_for(
                        state="visible", timeout=4000
                    )
                except Exception:
                    pausa_humana(0.25, 0.4)
        except Exception:
            continue

        guardados += _exportar_grilla_visible(np, dest, titulo, on_log, solo_viewport=True)
        if ver_detalles:
            guardados += _procesar_ver_detalles(np, dest, titulo, on_log)
    return guardados


def _procesar_ver_detalles(np, dest: Path, prefijo: str, on_log=None) -> list[str]:
    """Entra a cada «Ver detalles» (ícono de ojo) y exporta lo que haya."""
    guardados: list[str] = []
    selectores = (
        "span[id*='ver detalles-option']",
        "span[id*='row-ver']",
        "span:text-is('visibility')",
        "button[title*='detalle' i]",
        "button[aria-label*='detalle' i]",
        "span.material-symbols-rounded:text-is('visibility')",
    )
    ojos = None
    for sel in selectores:
        try:
            loc = np.locator(sel)
            if loc.count():
                ojos = loc
                break
        except Exception:
            continue
    if ojos is None:
        return guardados
    _revelar_contenido_np(np, on_log)
    indices = _indices_visibles_en_pantalla(np, ojos)
    if not indices:
        indices = list(range(_cantidad(ojos)))
    for i in indices:
        try:
            o = ojos.nth(i)
            if not _visible(o) and _estado_scroll(np).fondo_alcanzado:
                continue
            txt_icono = ""
            try:
                txt_icono = (o.inner_text(timeout=300) or "").strip().lower()
            except Exception:
                pass
            if txt_icono and txt_icono not in ("visibility", "remove_red_eye", ""):
                if "visibility" not in txt_icono and "detalle" not in txt_icono:
                    continue
            _click_sin_scroll(o)
            pausa_humana(0.25, 0.4)
            _esperar_np(np, 2500)
            guardados += _exportar_grilla_visible(np, dest, f"{prefijo}_detalle_{i+1}", on_log)
            _volver(np)
        except Exception as exc:
            _log(on_log, f"  • Ver detalles {i+1}: {exc}")
            _volver(np)
    return guardados


def _procesar_subitems(np, dest: Path, prefijo: str, on_log=None) -> list[str]:
    """Entra a cada ítem/concepto de la sección (filas o tarjetas clicables)."""
    guardados: list[str] = []
    candidatos: list[tuple[str, str]] = []

    try:
        raw = np.evaluate(
            """
            () => {
              const out = [];
              const seen = new Set();
              const push = (el, label) => {
                const t = (label || el.innerText || "").trim().replace(/\\s+/g, " ");
                if (!t || t.length < 3 || seen.has(t)) return;
                if (/exportar|volver|tu informaci|facturaci|patrimonial|declaraciones|inversiones/i.test(t)) return;
                seen.add(t);
                out.push(t.slice(0, 120));
              };
              document.querySelectorAll("a[href*='tu-informacion']").forEach((a) => {
                push(a, a.innerText);
              });
              document.querySelectorAll(
                ".list-group-item, .card-body a, [class*='cursor-pointer'], .pointer"
              ).forEach((el) => push(el));
              document.querySelectorAll("tbody tr.e-row, tr[role='row']").forEach((tr) => {
                if (tr.querySelector("button[id$='export-button']")) return;
                push(tr);
              });
              return out.slice(0, 120);
            }
            """
        )
        if isinstance(raw, list):
            candidatos = [
                (t, t)
                for t in raw
                if isinstance(t, str) and t.strip() and not _es_texto_seccion_principal(t)
            ]
    except Exception:
        pass

    if not candidatos:
        loc = np.locator("a[href*='tu-informacion'], .list-group-item, .card:has(.card-body)")
        try:
            n = _cantidad(loc)
        except Exception:
            n = 0
        for i in range(n):
            try:
                el = loc.nth(i)
                if not _visible(el):
                    continue
                txt = (el.inner_text(timeout=500) or "").strip().replace("\n", " ")
                txt = re.sub(r"\s+", " ", txt)[:120]
                if len(txt) >= 3 and not _es_texto_seccion_principal(txt):
                    candidatos.append((txt, txt))
            except Exception:
                continue

    vistos: set[str] = set()
    for etiqueta, _ in candidatos:
        if etiqueta in vistos:
            continue
        vistos.add(etiqueta)
        try:
            destino = np.get_by_text(re.compile(re.escape(etiqueta[:60]), re.I)).first
            if not destino.count() or not _visible(destino):
                continue
            _click_np(np, destino)
            pausa_humana(0.5, 0.9)
            _esperar_np(np)
            pref = _nombre_seguro(etiqueta, fallback=prefijo)[:80]
            guardados += _procesar_acordeones(np, dest, pref, on_log)
            guardados += _exportar_grilla_visible(np, dest, pref, on_log)
            guardados += _procesar_ver_detalles(np, dest, pref, on_log)
            _volver(np)
            pausa_humana(0.5, 0.9)
        except Exception as exc:
            _log(on_log, f"  • Subitem «{etiqueta[:50]}»: {exc}")
            _volver(np)
    return guardados


def _procesar_filas_grilla(np, dest: Path, prefijo: str, on_log=None) -> list[str]:
    """Exporta filas de grillas Syncfusion (conceptos patrimoniales u otros)."""
    guardados: list[str] = []
    filas = np.locator(
        ".e-gridcontent tbody tr.e-row:not(.e-emptyrow), "
        "tbody tr.e-row, table tbody tr[role='row']"
    )
    try:
        n = _cantidad(filas)
    except Exception:
        n = 0
    for i in range(n):
        row = filas.nth(i)
        if not _visible(row):
            continue
        titulo = f"{prefijo}_fila_{i + 1}"
        try:
            txt = (row.inner_text(timeout=800) or "").strip().replace("\n", " ")
            txt = re.sub(r"\s+", " ", txt)[:80]
            if txt and len(txt) > 2:
                titulo = _nombre_seguro(txt, fallback=titulo)[:80]
        except Exception:
            pass
        try:
            exp = row.locator("button[id$='export-button'], button:has-text('EXPORTAR')")
            if exp.count() and _visible(exp.first):
                _click_np(np, exp.first)
                pausa_humana(0.3, 0.5)
                ruta = _exportar_menu_actual(np, dest, titulo, on_log)
                if ruta:
                    guardados.append(ruta)
                continue
        except Exception:
            pass
        try:
            _click_np(np, row, intentar_revelar=True)
            pausa_humana(0.3, 0.5)
            _esperar_np(np)
            guardados += _exportar_grilla_visible(np, dest, titulo, on_log)
            guardados += _procesar_acordeones(np, dest, titulo, on_log)
            guardados += _procesar_ver_detalles(np, dest, titulo, on_log)
            _volver(np)
        except Exception as exc:
            _log(on_log, f"  • Fila grilla «{titulo}»: {exc}")
            _volver(np)
    return guardados


def _conceptos_hub(np, segmento: str) -> list[tuple[str, str]]:
    """Filas/conceptos clicables del hub de una sección (patrimonial, facturación, etc.)."""
    try:
        raw = np.evaluate(
            """
            (seg) => {
              const out = [];
              const seen = new Set();
              const push = (text, href, el) => {
                text = (text || '').trim().replace(/\\s+/g, ' ');
                if (!text || text.length < 4 || seen.has(text)) return;
                if (/exportar|volver|tu informaci|declaraciones juradas|informaci[oó]n patrimonial/i.test(text)) return;
                const tieneChevron = el && el.querySelector(
                  '.bi-chevron-right, .material-symbols-rounded, [class*="chevron"], [class*="arrow"]'
                );
                const subRuta = href && href.includes('tu-informacion/' + seg + '/') &&
                  !href.match(new RegExp('tu-informacion/' + seg + '/?$', 'i'));
                if (!subRuta && !tieneChevron && !(el && el.tagName === 'A' && href)) return;
                seen.add(text);
                out.push({ text: text.slice(0, 120), href: href || '' });
              };
              document.querySelectorAll('a[href]').forEach((a) => {
                const h = (a.getAttribute('href') || '').toLowerCase();
                if (h.includes('tu-informacion/' + seg + '/')) {
                  push(a.innerText, h, a);
                }
              });
              document.querySelectorAll(
                '.card, .list-group-item, .list-group-item-action, [class*="cursor-pointer"], tr.e-row'
              ).forEach((el) => {
                const html = el.innerHTML || '';
                const t = el.innerText || '';
                if (/chevron|arrow_forward|keyboard_arrow|visibility/i.test(html)) {
                  push(t, '', el);
                }
              });
              return out;
            }
            """,
            segmento,
        )
        if isinstance(raw, list):
            return [
                (str(x.get("text", "")).strip(), str(x.get("href", "")).strip())
                for x in raw
                if isinstance(x, dict) and str(x.get("text", "")).strip()
            ]
    except Exception:
        pass
    return []


def _ir_concepto_hub(np, etiqueta: str, href: str) -> None:
    if href:
        destino = href
        if href.startswith("#"):
            destino = f"{NP_BASE}{href.lstrip('#')}"
        elif href.startswith("/"):
            destino = f"https://cgpf-webv2-fe.arca.gob.ar{href}"
        elif not href.startswith("http"):
            destino = f"{NP_BASE}/{href.lstrip('/')}"
        np.goto(destino, wait_until="domcontentloaded")
    else:
        loc = np.get_by_text(re.compile(re.escape(etiqueta[:55]), re.I)).first
        if loc.count() and _visible(loc):
            _click_np(np, loc)
        else:
            raise AutomatizacionArcaError(f"No se encontró el concepto «{etiqueta[:40]}».")
    pausa_humana(0.4, 0.7)
    _esperar_np(np)


def _procesar_detalle_concepto(np, dest: Path, prefijo: str, on_log=None) -> list[str]:
    """Dentro de un concepto: acordeones, export, ojo."""
    guardados: list[str] = []
    guardados += _procesar_acordeones(np, dest, prefijo, on_log)
    guardados += _exportar_grilla_visible(np, dest, prefijo, on_log, solo_viewport=True)
    guardados += _procesar_ver_detalles(np, dest, prefijo, on_log)
    if not guardados:
        guardados += _procesar_filas_grilla(np, dest, prefijo, on_log)
    return guardados


def _procesar_hub_conceptos(
    np, dest: Path, prefijo: str, segmento: str, on_log=None
) -> list[str]:
    """Recorre conceptos del hub (filas con chevron, como en Facturación / Patrimonial)."""
    guardados: list[str] = []
    conceptos = _conceptos_hub(np, segmento)
    if not conceptos:
        return guardados
    _log(on_log, f"  • {len(conceptos)} concepto(s) detectados.")
    for etiqueta, href in conceptos:
        if _es_texto_seccion_principal(etiqueta):
            continue
        if guardados:
            _esperar_ritmo_np("general")
        try:
            _ir_concepto_hub(np, etiqueta, href)
            sub = _nombre_seguro(etiqueta, fallback=prefijo)[:80]
            guardados += _procesar_detalle_concepto(np, dest, sub, on_log)
            _volver(np)
            pausa_humana(0.4, 0.7)
        except Exception as exc:
            _log(on_log, f"  • Concepto «{etiqueta[:45]}»: {exc}")
            _volver(np)
    return guardados


def _procesar_ddjj(np, dest: Path, prefijo: str, on_log=None) -> list[str]:
    """Declaraciones juradas: exportar directo sin recorrer subitems (evita scroll largo)."""
    guardados = _exportar_grilla_visible(np, dest, prefijo, on_log, solo_viewport=True)
    if guardados:
        _log(on_log, f"  • DDJJ: {len(guardados)} archivo(s) exportado(s).")
        return guardados
    guardados += _procesar_acordeones(np, dest, prefijo, on_log)
    if not guardados:
        guardados += _procesar_ver_detalles(np, dest, prefijo, on_log)
    return guardados


def _procesar_patrimonial(np, dest: Path, prefijo: str, on_log=None) -> list[str]:
    """Información patrimonial: expandir grilla → ojo por fila → exportar (flujo grabado)."""
    guardados: list[str] = []
    try:
        expand = np.locator("span:text-is('keyboard_arrow_down')").first
        if expand.count() and _visible(expand):
            _click_sin_scroll(expand)
            pausa_humana(0.5, 0.8)
    except Exception:
        pass

    ojos = np.locator("span[id*='ver detalles-option'], span[id*='row-ver']")
    _revelar_contenido_np(np, on_log)
    indices = _indices_visibles_en_pantalla(np, ojos)
    if not indices:
        indices = list(range(_cantidad(ojos)))
    for idx, i in enumerate(indices):
        if idx > 0:
            _esperar_ritmo_np("general")
        o = ojos.nth(i)
        if not _visible(o) and _estado_scroll(np).fondo_alcanzado:
            continue
        try:
            _click_sin_scroll(o)
            pausa_humana(0.25, 0.4)
            _esperar_np(np, 2500)
            sub = f"{prefijo}_fila_{i + 1}"
            guardados += _exportar_grilla_visible(np, dest, sub, on_log, solo_viewport=True)
            _volver(np)
        except Exception as exc:
            _log(on_log, f"  • Patrimonial fila {i + 1}: {exc}")
            _volver(np)

    guardados += _exportar_grilla_visible(np, dest, prefijo, on_log, solo_viewport=True)
    return guardados


def _procesar_contenido_seccion(
    np, dest: Path, prefijo: str, on_log=None, *, clave: str = ""
) -> list[str]:
    """Recorre subitems, acordeones, grillas y «Ver detalles» de una sección."""
    if clave == "patrimonial":
        return _procesar_patrimonial(np, dest, prefijo, on_log)
    if clave == "ddjj":
        return _procesar_ddjj(np, dest, prefijo, on_log)
    if clave in ("facturacion", "inversiones"):
        ritmo = "facturacion_acordeon" if clave == "facturacion" else "general"
        guardados = _procesar_acordeones(
            np, dest, prefijo, on_log, modo_ritmo=ritmo, ver_detalles=False
        )
        if not guardados:
            guardados += _exportar_grilla_visible(np, dest, prefijo, on_log, solo_viewport=True)
        return guardados
    guardados: list[str] = []
    try:
        n_acc = np.locator("button.accordion-button").count()
    except Exception:
        n_acc = 0
    if n_acc:
        guardados += _procesar_acordeones(np, dest, prefijo, on_log, ver_detalles=True)
    if not guardados:
        guardados += _exportar_grilla_visible(np, dest, prefijo, on_log)
        guardados += _procesar_ver_detalles(np, dest, prefijo, on_log)
    return guardados


def _volver(np) -> None:
    try:
        v = np.get_by_role("button", name=re.compile(r"^\s*volver\s*$", re.I))
        if v.count() and _visible(v.first):
            _click_np(np, v.first)
            try:
                v.first.wait_for(state="hidden", timeout=3000)
            except Exception:
                pass
            _reset_scroll_np(np)
            return
    except Exception:
        pass
    try:
        np.go_back()
        _esperar_np(np, 2000)
        _reset_scroll_np(np)
    except Exception:
        pass


def _abrir_seccion_desde_tu_info(np, nombre_card: str, on_log=None) -> bool:
    """Hace clic en la tarjeta de la sección desde «Tu información»."""
    _volver_a_tu_informacion(np, on_log)
    pat = re.compile(re.escape(nombre_card[:50]), re.I)
    try:
        cards = np.locator(
            "div.contenedor div.item div[class*='e-card'], div[class*='e-card']"
        ).filter(has_text=pat)
        for i in range(_cantidad(cards, 8)):
            card = cards.nth(i)
            if _visible(card):
                _click_sin_scroll(card)
                _esperar_np(np, 2500)
                return not _en_hub_tu_informacion(np)
    except Exception:
        pass
    try:
        loc = np.get_by_text(pat)
        for i in range(_cantidad(loc, 8)):
            if _visible(loc.nth(i)):
                _click_sin_scroll(loc.nth(i))
                _esperar_np(np, 2500)
                return not _en_hub_tu_informacion(np)
    except Exception:
        pass
    return False


def _procesar_seccion(
    np, clave: str, nombre: str, dest_base: Path, ejercicio: str = "", on_log=None
) -> SeccionNP:
    sec = SeccionNP(clave=clave, nombre=nombre)
    dest = dest_base / _nombre_seguro(nombre, fallback=clave)
    dest.mkdir(parents=True, exist_ok=True)

    _reset_scroll_np(np)
    _esperar_ritmo_np("general")
    _log(on_log, f"Sección: {nombre}")
    if not _abrir_seccion_desde_tu_info(np, nombre, on_log):
        sec.nota = "No se pudo abrir o no tiene datos."
        _log(on_log, "  • No se pudo abrir la sección (o sin datos).")
        try:
            next(dest.iterdir())
        except StopIteration:
            try:
                dest.rmdir()
            except OSError:
                pass
        return sec

    _reset_scroll_np(np)
    sec.archivos += _procesar_contenido_seccion(np, dest, clave, on_log, clave=clave)

    if clave == "patrimonial":
        pdf = _exportar_pdf_pantalla(np, dest, f"{clave}_resumen", on_log)
        if pdf:
            sec.archivos.append(pdf)

    if not sec.archivos:
        sec.nota = "Sin archivos exportables (sección sin datos)."
        try:
            next(dest.iterdir())
        except StopIteration:
            try:
                dest.rmdir()
            except OSError:
                pass
    return sec


# --------------------------------------------------------------------------- #
# Orquestación
# --------------------------------------------------------------------------- #
def ejecutar_descarga_nuestra_parte(
    cred: CredencialesArca,
    ejercicio: str,
    *,
    carpeta_destino: Path,
    headless: bool | None = None,
    on_log: Callable[[str], None] | None = None,
    on_paso: Callable[[str, str], None] | None = None,
) -> ResultadoNPCuit:
    """Descarga las 4 secciones de «Nuestra Parte» para un CUIT."""
    headless = _headless_desde_env() if headless is None else headless
    if not _playwright_disponible():
        raise AutomatizacionNoDisponibleError(
            "Playwright no está instalado. En local: pip install playwright && playwright install chromium"
        )

    from playwright.sync_api import sync_playwright

    from cuit_en_arca.automation_playwright import (
        LOGIN_URL,
        _llenar_cuit_y_avanzar,
        _login_clave_fiscal,
        _nuevo_contexto_stealth,
    )

    carpeta_destino.mkdir(parents=True, exist_ok=True)
    resultado = ResultadoNPCuit(
        cuit_login=cred.cuit_login,
        cuit_representado=cred.cuit_representado,
        razon_social=None,
        ejercicio=ejercicio,
        carpeta=str(carpeta_destino),
    )

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

            _paso(on_paso, "servicio", "en_curso")
            np = _abrir_nuestra_parte(page)
            if _cuit_np_n(cred.cuit_representado) != _cuit_np_n(cred.cuit_login):
                _seleccionar_representado_np(
                    np,
                    cred.cuit_representado,
                    cuit_login=cred.cuit_login,
                    on_log=on_log,
                )
            resultado.razon_social = _leer_razon_social(np) or None
            _paso(on_paso, "servicio", "ok")

            _paso(on_paso, "tu_informacion", "en_curso")
            _ir_tu_informacion(np)
            sel = _seleccionar_ejercicio(np, ejercicio, on_log)
            if sel:
                resultado.ejercicio = sel
            _log(on_log, f"Ejercicio: {resultado.ejercicio or '(actual)'}")
            _paso(on_paso, "tu_informacion", "ok")

            _paso(on_paso, "descargar", "en_curso")
            _reset_ritmo_np()
            for clave, nombre in SECCIONES:
                try:
                    sec = _procesar_seccion(
                        np, clave, nombre, carpeta_destino, ejercicio, on_log
                    )
                except Exception as exc:
                    sec = SeccionNP(clave=clave, nombre=nombre, nota=f"Error: {exc}")
                    _log(on_log, f"  • Error en sección {nombre}: {exc}")
                resultado.secciones.append(sec)
            _paso(on_paso, "descargar", "ok")

            _log(on_log, f"Listo. Archivos: {resultado.total_archivos} en {carpeta_destino}")
            return resultado
    finally:
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass


def ejecutar_nuestra_parte_lote(
    filas,
    *,
    headless: bool | None = None,
    on_log: Callable[[str], None] | None = None,
    on_paso: Callable[[str, str], None] | None = None,
    on_reiniciar_pasos: Callable[[], None] | None = None,
    on_progreso: Callable[[int, int, str], None] | None = None,
    on_cuit_fin: Callable[[str, str | None, int, str | None], None] | None = None,
    carpeta_base: str | Path | None = None,
    job_id: str | None = None,
    modo_ap: bool = False,
    nombre_carpeta_sesion: str | None = None,
) -> Path:
    """Procesa varias filas (CUIT) de «Nuestra Parte». Tolerante a errores."""
    headless = _headless_desde_env() if headless is None else headless
    base = carpeta_np_base(
        base_elegida=carpeta_base,
        nombre_sesion=nombre_carpeta_sesion,
    )
    total = len(filas)
    _log(on_log, f"Carpeta de destino: {base}")
    resumen_lote: list[dict] = []

    from cuit_en_arca.cancelacion import verificar_cancelacion

    for idx, fila in enumerate(filas, start=1):
        if job_id:
            verificar_cancelacion(job_id)
        elif modo_ap:
            verificar_cancelacion(ap=True)
        cuit_log = getattr(fila, "cuit_login", "")
        cuit_repr = getattr(fila, "cuit_representado", "") or cuit_log
        ejercicio = getattr(fila, "ejercicio", "") or ""
        if on_progreso:
            on_progreso(idx - 1, total, f"CUIT {cuit_repr} ({idx}/{total})")
        if on_reiniciar_pasos:
            on_reiniciar_pasos()

        cred = CredencialesArca(
            cuit_login=cuit_log,
            clave_fiscal=getattr(fila, "clave_fiscal", ""),
            cuit_representado=cuit_repr,
        )
        dest = base / _nombre_seguro(cuit_repr, fallback=cuit_log or f"cuit_{idx}")
        dest.mkdir(parents=True, exist_ok=True)

        try:
            res = ejecutar_descarga_nuestra_parte(
                cred,
                ejercicio,
                carpeta_destino=dest,
                headless=headless,
                on_log=on_log,
                on_paso=on_paso,
            )
            if on_cuit_fin:
                on_cuit_fin(cuit_repr, res.razon_social, res.total_archivos, None)
            resumen_lote.append(
                {"cuit": cuit_repr, "razon_social": res.razon_social or "", "error": None}
            )
        except Exception as exc:
            _log(on_log, f"CUIT {cuit_repr}: ERROR {exc}")
            if on_paso:
                try:
                    on_paso("descargar", "error")
                except Exception:
                    pass
            if on_cuit_fin:
                on_cuit_fin(cuit_repr, None, 0, str(exc))
            resumen_lote.append(
                {"cuit": cuit_repr, "razon_social": "", "error": str(exc)}
            )

        if on_progreso:
            on_progreso(idx, total, f"CUIT {cuit_repr} completado ({idx}/{total})")

    from cuit_en_arca.fallos_arca import escribir_fallos_txt

    escribir_fallos_txt(base, resumen_cuits=resumen_lote)
    return base
