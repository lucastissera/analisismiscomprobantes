"""Automatización del Domicilio Fiscal Electrónico (Ventanilla Electrónica de ARCA).

Flujo:
1. Login con clave fiscal (reutiliza el login de Mis Comprobantes).
2. Abre el servicio «Domicilio Fiscal Electrónico» **desde el portal** (el SSO no
   funciona entrando por URL directa a ``ve.cloud.afip.gob.ar``).
3. Cierra el popup inicial ("Recordar más tarde").
4. Si el CUIT representado difiere del login: pestaña «Comunicaciones de mis
   representados» y selección del CUIT en el desplegable.
5. Aplica el rango de fechas y lista las comunicaciones recibidas.
6. Por cada comunicación:
   - Si tiene archivo adjunto (``a#adjunto-nombre``: .txt / .pdf), lo **descarga**.
   - Si es solo informativa en pantalla, **imprime la pantalla a PDF**.
7. Guarda todo en una carpeta del escritorio: ``DFE yyyy-mm-dd``.

Pensado para ejecutarse en modo **visible** (headless=False) en la PC del usuario.
"""

from __future__ import annotations

import io
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from cuit_en_arca.credenciales import CredencialesArca
from cuit_en_arca.errores import (
    AutomatizacionArcaError,
    AutomatizacionNoDisponibleError,
    CuitRepresentadoNoEncontradoError,
    LoginArcaError,
)
from cuit_en_arca.stealth import clic_humano, escribir_como_humano, pausa_humana

_DFE_ESPERA_TABLA_MS = 10_000
VE_TERMINO_BUSQUEDA = "Domicilio Fiscal Electrónico"


@dataclass
class ComunicacionDfe:
    numero: str
    asunto: str
    fecha: str
    tipo: str  # 'adjunto' | 'pantalla_pdf'
    archivos: list[str] = field(default_factory=list)


@dataclass
class ResultadoDfeCuit:
    cuit_login: str
    cuit_representado: str
    razon_social: str | None
    carpeta: str
    comunicaciones: list[ComunicacionDfe] = field(default_factory=list)
    error: str | None = None

    @property
    def total_archivos(self) -> int:
        return sum(len(c.archivos) for c in self.comunicaciones)


def _playwright_disponible() -> bool:
    try:
        import playwright  # noqa: F401

        return True
    except Exception:
        return False


def _log(on_log: Callable[[str], None] | None, msg: str) -> None:
    if on_log:
        try:
            on_log(msg)
        except Exception:
            pass


def _escritorio_windows() -> Path | None:
    """Ruta del Escritorio según el registro (soporta escritorios reubicados)."""
    if sys.platform != "win32":
        return None
    import os

    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
        ) as k:
            val, _ = winreg.QueryValueEx(k, "Desktop")
            val = os.path.expandvars(val)
            p = Path(val)
            if p.is_dir():
                return p
    except Exception:
        return None
    return None


def carpeta_dfe_escritorio(
    hoy: date | None = None,
    base_elegida: str | Path | None = None,
    *,
    nombre_sesion: str | None = None,
) -> Path:
    """Carpeta ``DFE yyyy-mm-dd HH-MM``."""
    from cuit_en_arca.carpetas_salida import stamp_carpeta_ejecucion

    if nombre_sesion:
        nombre = nombre_sesion
    else:
        momento = (
            datetime.now()
            if hoy is None
            else datetime.combine(hoy, datetime.now().time())
        )
        nombre = f"DFE {stamp_carpeta_ejecucion(momento)}"
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
    base = escritorio or home
    destino = base / nombre
    destino.mkdir(parents=True, exist_ok=True)
    return destino


_INVALIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _nombre_seguro(nombre: str, *, fallback: str = "archivo") -> str:
    nombre = (nombre or "").strip()
    nombre = _INVALIDOS.sub("_", nombre)
    nombre = re.sub(r"\s+", " ", nombre).strip(" .")
    return nombre or fallback


def _fecha_a_iso(recibido: str) -> str:
    m = re.search(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", recibido or "")
    if m:
        d, mth, y = m.groups()
        return f"{y}-{int(mth):02d}-{int(d):02d}"
    return ""


def _parsear_fecha_recibido(recibido: str) -> date | None:
    m = re.search(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", recibido or "")
    if not m:
        return None
    d, mth, y = m.groups()
    try:
        return date(int(y), int(mth), int(d))
    except ValueError:
        return None


def _png_a_pdf(png_bytes: bytes, ruta: Path) -> None:
    from PIL import Image

    img = Image.open(io.BytesIO(png_bytes))
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.save(str(ruta), "PDF", resolution=100.0)


# --------------------------------------------------------------------------- #
# Apertura del servicio DFE desde el portal
# --------------------------------------------------------------------------- #
def _abrir_dfe(page):
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

    ve = None
    buscador = None
    for ctx in _iter_contextos(page):
        buscador = _locator_buscador_servicios(ctx)
        if buscador is not None:
            break

    if buscador is not None:
        escribir_como_humano(buscador, VE_TERMINO_BUSQUEDA)
        pausa_humana(1.0, 1.8)
        res = page.get_by_text(re.compile(r"Domicilio Fiscal Electr", re.I))
        for i in range(min(res.count(), 8)):
            try:
                if res.nth(i).is_visible(timeout=900):
                    ve = _click_servicio_y_obtener_pagina(page, res.nth(i))
                    break
            except Exception:
                continue

    if ve is None:
        loc = page.locator("a.accesoPrincipal[title*='Domicilio' i]")
        if loc.count():
            try:
                ve = _click_servicio_y_obtener_pagina(page, loc.first)
            except Exception:
                ve = None

    # Elegir la pestaña real de la ventanilla (ve.cloud).
    for pg in page.context.pages:
        if "ve.cloud" in (pg.url or ""):
            ve = pg
            break
    if ve is None:
        ve = page

    ve.set_default_timeout(40_000)
    try:
        ve.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        pass
    pausa_humana(2.0, 3.2)

    if "ve.cloud" not in (ve.url or ""):
        raise AutomatizacionArcaError(
            "No se pudo abrir el Domicilio Fiscal Electrónico desde el portal. "
            "Verifique que el CUIT tenga el servicio habilitado."
        )
    return ve


_DFE_PANEL_REPRESENTADOS = "#representados-comunicaciones-tab"
_DFE_BTN_APLICAR_REPRESENTADOS = "button-aplicar-fil-superior"
_DFE_BTN_VOLVER_REPRESENTADOS = "comunicacion-btn-volver"
_DFE_ESPERA_FILTROS_REPRESENTADO_MS = 5_000


def _cuit_dfe_n(s: str) -> str:
    from cuit_en_arca.automation_playwright import _normalizar_cuit_busqueda

    return _normalizar_cuit_busqueda(s)


def _cuit_dfe_fmt(cuit_n: str) -> str:
    if len(cuit_n) != 11:
        return cuit_n
    return f"{cuit_n[:2]}-{cuit_n[2:10]}-{cuit_n[10]}"


def _activar_pestana_representados_dfe(ve, on_log=None) -> None:
    tab = ve.locator("#representados-comunicaciones-tab___BV_tab_button__")
    if not tab.count():
        tab = ve.get_by_role("tab", name=re.compile(r"comunicaciones de mis representados", re.I))
    if not tab.count():
        tab = ve.get_by_text(re.compile(r"comunicaciones de mis representados", re.I))
    if not tab.count():
        raise AutomatizacionArcaError(
            "No se encontró la pestaña «Comunicaciones de mis representados» en DFE."
        )
    clic_humano(tab.first)
    pausa_humana(0.45, 0.85)
    ve.locator(_DFE_PANEL_REPRESENTADOS).wait_for(state="visible", timeout=12_000)
    _log(on_log, "Pestaña «Comunicaciones de mis representados» activa.")


def _seleccionar_cuit_representado_dfe(ve, cuit_repr: str, on_log=None) -> None:
    cuit_n = _cuit_dfe_n(cuit_repr)
    fmt = _cuit_dfe_fmt(cuit_n)
    panel = ve.locator(_DFE_PANEL_REPRESENTADOS)
    panel.wait_for(state="visible", timeout=12_000)

    resultado = ve.evaluate(
        """
        ([cuit, fmt]) => {
          const norm = (t) => (t || "").replace(/\\D/g, "");
          const panel = document.querySelector("#representados-comunicaciones-tab");
          if (!panel) return { ok: false, reason: "sin_panel" };
          const coincide = (texto) => {
            const t = (texto || "").trim();
            if (!t) return false;
            const n = norm(t);
            return n === cuit || t.includes(fmt) || t.includes(cuit);
          };
          for (const sel of panel.querySelectorAll("select")) {
            for (const opt of sel.options) {
              const text = (opt.textContent || opt.label || "").trim();
              const val = (opt.value || "").trim();
              if (coincide(text) || coincide(val)) {
                sel.value = opt.value;
                sel.dispatchEvent(new Event("input", { bubbles: true }));
                sel.dispatchEvent(new Event("change", { bubbles: true }));
                return { ok: true, method: "select", label: text.slice(0, 120) };
              }
            }
          }
          return { ok: false, reason: "sin_select" };
        }
        """,
        [cuit_n, fmt],
    )

    if isinstance(resultado, dict) and resultado.get("ok"):
        _log(
            on_log,
            f"CUIT representado seleccionado ({resultado.get('label') or cuit_n}).",
        )
        pausa_humana(0.15, 0.3)
        return

    # Desplegable custom (Bootstrap / Vue): abrir y elegir opción visible.
    cands = panel.locator("select, [role='combobox'], .custom-select")
    for i in range(cands.count()):
        candidato = cands.nth(i)
        try:
            if not candidato.is_visible(timeout=800):
                continue
            clic_humano(candidato)
            pausa_humana(0.25, 0.45)
        except Exception:
            continue

    for patron in (fmt, cuit_n):
        for loc in (
            panel.get_by_role("option", name=re.compile(re.escape(patron), re.I)),
            panel.locator("option, li, a").filter(has_text=re.compile(re.escape(patron), re.I)),
            ve.get_by_role("option", name=re.compile(re.escape(patron), re.I)),
        ):
            try:
                if loc.count() and loc.first.is_visible(timeout=1200):
                    clic_humano(loc.first)
                    _log(on_log, f"CUIT representado seleccionado: {cuit_n}.")
                    pausa_humana(0.15, 0.3)
                    return
            except Exception:
                continue

    raise CuitRepresentadoNoEncontradoError(
        f"El CUIT representado {fmt} no aparece en el desplegable de DFE."
    )


def _panel_representados_dfe(ve):
    return ve.locator(_DFE_PANEL_REPRESENTADOS)


def _campo_fecha_representados(ve, cual: str):
    """Campo de fecha dentro del panel representados (evita la pestaña titular)."""
    panel = _panel_representados_dfe(ve)
    cid = "daterange-fechas-desde" if cual == "desde" else "daterange-fechas-hasta"
    return panel.locator(f"#{cid}").first


def _escribir_fecha_dfe(campo, valor: date) -> None:
    """Vue datepicker: clic + tipeo humano (``fill`` solo no dispara el filtro)."""
    texto = valor.strftime("%d/%m/%Y")
    escribir_como_humano(campo, texto)
    try:
        campo.evaluate(
            """(el, v) => {
              el.value = v;
              el.dispatchEvent(new Event("input", { bubbles: true }));
              el.dispatchEvent(new Event("change", { bubbles: true }));
              el.blur();
            }""",
            texto,
        )
    except Exception:
        pass
    pausa_humana(0.12, 0.25)


def _click_aplicar_representados(ve, on_log=None, *, motivo: str = "") -> None:
    root = _panel_representados_dfe(ve)
    if not _click_boton_aplicar_dfe(
        ve, root, _DFE_BTN_APLICAR_REPRESENTADOS, solo_panel=True
    ):
        raise AutomatizacionArcaError(
            "No se pudo hacer clic en «Aplicar» (comunicaciones de representados)."
        )
    if motivo:
        _log(on_log, f"Aplicar ({motivo}).")
    pausa_humana(0.35, 0.6)
    _esperar_tabla_representados(ve)


def _configurar_dfe_representados(ve, cuit_repr: str, on_log=None) -> None:
    _activar_pestana_representados_dfe(ve, on_log)
    _seleccionar_cuit_representado_dfe(ve, cuit_repr, on_log)
    _esperar_filtros_representados_dfe(ve, on_log)
    # Tras elegir el CUIT hay que Aplicar para cargar la grilla (grabación 121636).
    _click_aplicar_representados(ve, on_log, motivo="CUIT representado")


def _esperar_filtros_representados_dfe(ve, on_log=None) -> None:
    """Espera hasta 5 s a que estén visibles fechas y botón Aplicar."""
    panel = ve.locator(_DFE_PANEL_REPRESENTADOS)
    try:
        panel.locator("#daterange-fechas-desde").wait_for(
            state="visible", timeout=_DFE_ESPERA_FILTROS_REPRESENTADO_MS
        )
        panel.locator(f"#{_DFE_BTN_APLICAR_REPRESENTADOS}").wait_for(
            state="visible", timeout=_DFE_ESPERA_FILTROS_REPRESENTADO_MS
        )
        _log(on_log, "Filtros de fechas listos.")
    except Exception as exc:
        _log(on_log, f"Espera de filtros DFE representados: {exc}")


def _cerrar_popup_dfe(ve, on_log=None) -> None:
    try:
        btn = ve.get_by_role("button", name=re.compile(r"recordar m[aá]s tarde", re.I))
        if btn.count() and btn.first.is_visible(timeout=4000):
            clic_humano(btn.first)
            _log(on_log, "Popup inicial cerrado (Recordar más tarde).")
            pausa_humana(0.25, 0.45)
            return
    except Exception:
        pass
    # Alternativa: "Entendido"
    try:
        btn = ve.get_by_role("button", name=re.compile(r"^entendido$", re.I))
        if btn.count() and btn.first.is_visible(timeout=2000):
            clic_humano(btn.first)
            pausa_humana(0.2, 0.4)
    except Exception:
        pass


def _es_select_paginacion(s) -> bool:
    """Evita confundir el desplegable de CUIT representado con «registros por página»."""
    try:
        opts = s.locator("option")
        n = opts.count()
        if n < 2 or n > 10:
            return False
        vals: list[int] = []
        for j in range(n):
            v = (opts.nth(j).get_attribute("value") or "").strip()
            if not v.isdigit():
                return False
            vals.append(int(v))
        return bool(vals) and max(vals) <= 500
    except Exception:
        return False


def _maximizar_pagina_representados(ve, on_log=None) -> None:
    """Muestra hasta 100 comunicaciones por página (#per-page-select en representados)."""
    panel = _panel_representados_dfe(ve)
    sel = panel.locator("#per-page-select").first
    try:
        if not sel.count() or not sel.is_visible(timeout=4000):
            return
        for val in ("100", "50"):
            try:
                sel.select_option(val, timeout=8_000)
                _log(on_log, f"Paginación representados: {val} por página.")
                pausa_humana(0.3, 0.5)
                _esperar_tabla_representados(ve, timeout_ms=8_000)
                return
            except Exception:
                continue
    except Exception as exc:
        _log(on_log, f"No se pudo ajustar paginación representados: {exc}")


def _maximizar_registros(ve) -> None:
    """Intenta mostrar más registros por página (select de paginación, no CUIT)."""
    try:
        sels = ve.locator("select")
        for i in range(min(sels.count(), 12)):
            s = sels.nth(i)
            if not _es_select_paginacion(s):
                continue
            opts = s.locator("option")
            vals: list[str] = []
            for j in range(opts.count()):
                v = (opts.nth(j).get_attribute("value") or "").strip()
                if v.isdigit():
                    vals.append(v)
            if vals:
                mejor = max(vals, key=lambda x: int(x))
                if int(mejor) > 10:
                    s.select_option(mejor, timeout=8_000)
                    pausa_humana(0.25, 0.45)
                return
    except Exception:
        pass


def _click_boton_aplicar_dfe(
    ve, root, id_boton: str | None = None, *, solo_panel: bool = False
) -> bool:
    botones: list = []
    if id_boton:
        botones.append(root.locator(f"#{id_boton}").first)
        if not solo_panel:
            botones.append(ve.locator(f"#{id_boton}").first)
    botones.append(root.get_by_role("button", name=re.compile(r"^\s*aplicar\s*$", re.I)).first)
    if not solo_panel:
        botones.append(
            ve.get_by_role("button", name=re.compile(r"^\s*aplicar\s*$", re.I)).first
        )
    for btn in botones:
        try:
            if not btn.count():
                continue
            btn.wait_for(state="visible", timeout=_DFE_ESPERA_FILTROS_REPRESENTADO_MS)
            btn.scroll_into_view_if_needed(timeout=3000)
            try:
                clic_humano(btn)
            except Exception:
                btn.click(timeout=8000)
            return True
        except Exception:
            continue
    return False


def _aplicar_rango_representados(
    ve,
    desde: date,
    hasta: date,
    on_log=None,
) -> None:
    """En representados: fecha desde → Aplicar → fecha hasta → Aplicar (grabación ARCA)."""
    di = _campo_fecha_representados(ve, "desde")
    dh = _campo_fecha_representados(ve, "hasta")
    try:
        di.wait_for(state="visible", timeout=_DFE_ESPERA_FILTROS_REPRESENTADO_MS)
        dh.wait_for(state="visible", timeout=_DFE_ESPERA_FILTROS_REPRESENTADO_MS)
    except Exception as exc:
        raise AutomatizacionArcaError(
            "No se encontraron los campos de fecha en comunicaciones de representados."
        ) from exc
    _escribir_fecha_dfe(di, desde)
    _click_aplicar_representados(ve, on_log, motivo="fecha desde")
    _escribir_fecha_dfe(dh, hasta)
    _click_aplicar_representados(ve, on_log, motivo="fecha hasta")
    _log(on_log, f"Rango aplicado (representados): {desde:%d/%m/%Y} – {hasta:%d/%m/%Y}.")
    pausa_humana(0.2, 0.35)


def _esperar_tabla_representados(ve, *, timeout_ms: int | None = None) -> None:
    limite = timeout_ms if timeout_ms is not None else _DFE_ESPERA_TABLA_MS
    try:
        ve.locator(f"{_DFE_PANEL_REPRESENTADOS} a[id^='sistema[']").first.wait_for(
            state="attached", timeout=limite
        )
    except Exception:
        pass
    pausa_humana(0.15, 0.3)


def _leer_filas_representados(ve) -> list[dict]:
    """Lista de comunicaciones en pestaña representados (enlaces ``sistema[id]``)."""
    try:
        raw = ve.evaluate(
            """
            () => {
              const panel = document.querySelector("#representados-comunicaciones-tab");
              if (!panel) return [];
              const out = [];
              panel.querySelectorAll('a[id^="sistema["]').forEach((a) => {
                const m = (a.id || "").match(/sistema\\[(\\d+)\\]/);
                if (!m) return;
                let recibido = "";
                const tr = a.closest("tr");
                if (tr) {
                  const dm = (tr.innerText || "").match(/\\d{1,2}\\/\\d{1,2}\\/\\d{4}/);
                  if (dm) recibido = dm[0];
                }
                out.push({
                  id: m[1],
                  sistema_id: a.id,
                  asunto: (a.innerText || a.textContent || "").trim(),
                  recibido,
                });
              });
              return out;
            }
            """
        )
        if isinstance(raw, list):
            return [dict(x) for x in raw if x.get("id")]
    except Exception:
        pass
    out: list[dict] = []
    panel = ve.locator(_DFE_PANEL_REPRESENTADOS)
    links = panel.locator("a[id^='sistema[']")
    n = min(links.count(), 20)
    for i in range(n):
        a = links.nth(i)
        try:
            sid = a.get_attribute("id") or ""
            m = re.search(r"sistema\[(\d+)\]", sid)
            if not m:
                continue
            asunto = a.inner_text(timeout=150).strip()
            out.append({"id": m.group(1), "sistema_id": sid, "asunto": asunto, "recibido": ""})
        except Exception:
            continue
    return out


def _abrir_detalle_representado(ve, com: dict) -> bool:
    sid = com.get("sistema_id") or f"sistema[{com['id']}]"
    panel = _panel_representados_dfe(ve)
    link = panel.locator(f'a[id="{sid}"]').first
    if not link.count():
        asunto = (com.get("asunto") or "")[:40]
        if asunto:
            link = panel.locator("a[id^='sistema[']").filter(
                has_text=re.compile(re.escape(asunto[:25]), re.I)
            ).first
    if not link.count():
        return False
    try:
        link.scroll_into_view_if_needed(timeout=5000)
        link.wait_for(state="visible", timeout=5000)
        clic_humano(link)
    except Exception:
        try:
            link.click(timeout=8000)
        except Exception:
            return False
    pausa_humana(0.45, 0.85)
    try:
        ve.wait_for_load_state("domcontentloaded", timeout=8_000)
    except Exception:
        pass
    return True


def _aplicar_rango(
    ve,
    desde: date | None,
    hasta: date | None,
    on_log=None,
    *,
    panel: str | None = None,
    id_boton_aplicar: str | None = None,
) -> None:
    if not desde or not hasta:
        return
    try:
        root = ve.locator(panel) if panel else ve
        di = root.locator("#daterange-fechas-desde").first
        dh = root.locator("#daterange-fechas-hasta").first
        if not di.count() or not dh.count():
            di = ve.locator("#daterange-fechas-desde").first
            dh = ve.locator("#daterange-fechas-hasta").first
        if di.count() and dh.count():
            di.fill(desde.strftime("%d/%m/%Y"))
            pausa_humana(0.15, 0.3)
            dh.fill(hasta.strftime("%d/%m/%Y"))
            pausa_humana(0.15, 0.3)
            if not _click_boton_aplicar_dfe(ve, root, id_boton_aplicar):
                raise AutomatizacionArcaError(
                    "No se pudo hacer clic en el botón «Aplicar» del filtro de fechas."
                )
            _log(on_log, f"Rango aplicado: {desde:%d/%m/%Y} – {hasta:%d/%m/%Y}.")
            _esperar_tabla_comunicaciones(ve, timeout_ms=_DFE_ESPERA_TABLA_MS)
            pausa_humana(0.2, 0.35)
        else:
            raise AutomatizacionArcaError(
                "No se encontraron los campos de fecha (#daterange-fechas-desde/hasta)."
            )
    except AutomatizacionArcaError:
        raise
    except Exception as exc:
        _log(on_log, f"No se pudo aplicar el rango de fechas ({exc}).")
        raise AutomatizacionArcaError(
            f"No se pudo aplicar el rango de fechas: {exc}"
        ) from exc


def _esperar_tabla_comunicaciones(ve, *, timeout_ms: int | None = None) -> None:
    """Espera a que la grilla tenga filas (tope configurable, default 10 s)."""
    limite = timeout_ms if timeout_ms is not None else _DFE_ESPERA_TABLA_MS
    try:
        ve.locator("tr[role='row'][aria-rowindex] [id^='fechaNotificacion[']").first.wait_for(
            state="attached", timeout=limite
        )
    except Exception:
        pass
    pausa_humana(0.15, 0.3)


def _preparar_listado_dfe(
    ve,
    fecha_desde: date | None,
    fecha_hasta: date | None,
    on_log=None,
    *,
    representados: bool = False,
) -> list[dict]:
    """Tras cerrar el popup: filtrar, maximizar paginación y leer filas."""
    if representados:
        if fecha_desde and fecha_hasta:
            _aplicar_rango_representados(ve, fecha_desde, fecha_hasta, on_log)
        else:
            _esperar_tabla_representados(ve)
        _maximizar_pagina_representados(ve, on_log)
        filas = _leer_filas_representados(ve)
        _log(on_log, f"Comunicaciones representados en grilla: {len(filas)}.")
        return filas
    if fecha_desde and fecha_hasta:
        _aplicar_rango(ve, fecha_desde, fecha_hasta, on_log)
    else:
        _esperar_tabla_comunicaciones(ve, timeout_ms=_DFE_ESPERA_TABLA_MS)
    _maximizar_registros(ve)
    return _leer_filas(ve)


def _leer_filas(ve) -> list[dict]:
    """Lee filas de comunicaciones en bloque (rápido; evita ~50 s en grillas grandes)."""
    try:
        raw = ve.evaluate(
            """
            () => {
              const out = [];
              document.querySelectorAll("tr[role='row'][aria-rowindex]").forEach((tr) => {
                const idEl = tr.querySelector("[id^='fechaNotificacion[']");
                if (!idEl) return;
                const m = (idEl.id || "").match(/\\[(\\d+)\\]/);
                if (!m) return;
                const asunto = (tr.querySelector("td[aria-colindex='2']")?.innerText || "").trim();
                let recibido = "";
                for (const col of ["7", "6", "5"]) {
                  const txt = (tr.querySelector(`td[aria-colindex='${col}']`)?.innerText || "").trim();
                  if (/\\d{1,2}[\\/\\-]\\d{1,2}[\\/\\-]\\d{4}/.test(txt)) {
                    recibido = txt;
                    break;
                  }
                }
                out.push({ id: m[1], asunto, recibido });
              });
              return out;
            }
            """
        )
        if isinstance(raw, list):
            return [dict(x) for x in raw if x.get("id")]
    except Exception:
        pass
    out: list[dict] = []
    filas = ve.locator("tr[role='row'][aria-rowindex]")
    n = min(filas.count(), 40)
    for i in range(n):
        f = filas.nth(i)
        cid = ""
        try:
            idel = f.locator("[id^='fechaNotificacion[']").first
            if idel.count():
                idattr = idel.get_attribute("id") or ""
                m = re.search(r"\[(\d+)\]", idattr)
                cid = m.group(1) if m else ""
        except Exception:
            cid = ""
        if not cid:
            continue
        asunto = ""
        try:
            asunto = f.locator("td[aria-colindex='2']").first.inner_text(timeout=300).strip()
        except Exception:
            asunto = ""
        recibido = ""
        for col in ("7", "6", "5"):
            try:
                txt = f.locator(f"td[aria-colindex='{col}']").first.inner_text(timeout=200).strip()
                if re.search(r"\d{1,2}[/\-]\d{1,2}[/\-]\d{4}", txt):
                    recibido = txt
                    break
            except Exception:
                continue
        out.append({"id": cid, "asunto": asunto, "recibido": recibido})
    return out


def _abrir_detalle(ve, cid: str) -> bool:
    fila = ve.locator(f"tr:has([id='fechaNotificacion[{cid}]'])").first
    if not fila.count():
        fila = ve.locator(f"tr:has([id='adjunto[{cid}]'])").first
    if not fila.count():
        return False
    celda = fila.locator("td[aria-colindex='2']").first
    try:
        clic_humano(celda)
    except Exception:
        try:
            celda.click(timeout=4000)
        except Exception:
            return False
    pausa_humana(0.5, 0.9)
    try:
        ve.wait_for_load_state("domcontentloaded", timeout=8_000)
    except Exception:
        pass
    return True


def _volver_a_lista(ve, *, representados: bool = False) -> None:
    if representados:
        try:
            volver = ve.locator(f"#{_DFE_BTN_VOLVER_REPRESENTADOS}")
            if volver.count() and volver.first.is_visible(timeout=3000):
                clic_humano(volver.first)
                pausa_humana(0.45, 0.85)
                _esperar_tabla_representados(ve, timeout_ms=8000)
                return
        except Exception:
            pass
    try:
        volver = ve.get_by_role("button", name=re.compile(r"^\s*volver\s*$", re.I))
        if volver.count():
            clic_humano(volver.first)
            pausa_humana(0.5, 0.9)
            return
    except Exception:
        pass
    try:
        ve.go_back()
        pausa_humana(0.5, 0.9)
    except Exception:
        pass


def _descargar_o_imprimir(ve, com: dict, dest: Path, on_log=None) -> ComunicacionDfe:
    cid = com["id"]
    asunto = com.get("asunto", "")
    recibido = com.get("recibido", "")
    res = ComunicacionDfe(numero=cid, asunto=asunto, fecha=recibido, tipo="pantalla_pdf")

    anchors = ve.locator("a#adjunto-nombre")
    na = anchors.count()
    if na > 0:
        res.tipo = "adjunto"
        for i in range(na):
            a = anchors.nth(i)
            nombre = (a.get_attribute("title") or "").strip()
            try:
                with ve.expect_download(timeout=30_000) as di:
                    a.click()
                d = di.value
                fn = _nombre_seguro(nombre or d.suggested_filename, fallback=f"adjunto_{cid}_{i}")
                ruta = dest / fn
                if ruta.exists():
                    ruta = dest / f"{ruta.stem}_{i}{ruta.suffix}"
                d.save_as(str(ruta))
                res.archivos.append(str(ruta))
                _log(on_log, f"  • Adjunto descargado: {ruta.name}")
            except Exception as exc:
                _log(on_log, f"  • No se pudo descargar adjunto de {cid}: {exc}")
    else:
        # Comunicación informativa → imprimir pantalla a PDF.
        try:
            png = ve.screenshot(full_page=True)
            iso = _fecha_a_iso(recibido)
            base = f"{iso + '_' if iso else ''}{_nombre_seguro(asunto, fallback='comunicacion')}_{cid}"
            ruta = dest / (base[:120] + ".pdf")
            _png_a_pdf(png, ruta)
            res.archivos.append(str(ruta))
            _log(on_log, f"  • Pantalla impresa a PDF: {ruta.name}")
        except Exception as exc:
            _log(on_log, f"  • No se pudo imprimir a PDF la comunicación {cid}: {exc}")
    return res


def _seleccionar_objetivo(
    filas: list[dict],
    desde: date | None,
    hasta: date | None,
    *,
    representados: bool = False,
) -> list[dict]:
    if representados and filas:
        if desde and hasta:
            en_rango = [
                f
                for f in filas
                if (_parsear_fecha_recibido(f.get("recibido", "")) is None)
                or (desde <= _parsear_fecha_recibido(f.get("recibido", "")) <= hasta)
            ]
            return en_rango or filas
        return filas
    if desde and hasta:
        # Si no se pudo parsear la fecha, igual se incluye (el filtro del sitio ya aplicó).
        en_rango = [
            f
            for f in filas
            if (_parsear_fecha_recibido(f.get("recibido", "")) is None)
            or (desde <= _parsear_fecha_recibido(f.get("recibido", "")) <= hasta)
        ]
        if en_rango:
            return en_rango
    # Sin rango o sin resultados en rango → últimas 3 (las que haya).
    return filas[:3]


def ejecutar_descarga_dfe(
    cred: CredencialesArca,
    fecha_desde: date | None,
    fecha_hasta: date | None,
    *,
    carpeta_destino: Path,
    headless: bool = False,
    on_log: Callable[[str], None] | None = None,
    on_paso: Callable[[str, str], None] | None = None,
) -> ResultadoDfeCuit:
    """Descarga las comunicaciones del DFE de un CUIT a ``carpeta_destino``."""
    if not _playwright_disponible():
        raise AutomatizacionNoDisponibleError(
            "Playwright no está instalado. En local: pip install playwright && playwright install chromium"
        )

    from playwright.sync_api import TimeoutError as PlaywrightTimeout
    from playwright.sync_api import sync_playwright

    from cuit_en_arca.automation_playwright import (
        LOGIN_URL,
        _llenar_cuit_y_avanzar,
        _login_clave_fiscal,
        _nuevo_contexto_stealth,
        _razon_social_activa_mcmp,
    )

    carpeta_destino.mkdir(parents=True, exist_ok=True)
    resultado = ResultadoDfeCuit(
        cuit_login=cred.cuit_login,
        cuit_representado=cred.cuit_representado,
        razon_social=None,
        carpeta=str(carpeta_destino),
    )

    def paso(clave: str, estado: str) -> None:
        if on_paso:
            try:
                on_paso(clave, estado)
            except Exception:
                pass

    browser = None
    try:
        with sync_playwright() as p:
            browser, context = _nuevo_contexto_stealth(p, headless=headless)
            page = context.new_page()
            page.set_default_timeout(60_000)

            paso("login", "en_curso")
            _log(on_log, f"Iniciando sesión en ARCA (CUIT {cred.cuit_login})…")
            page.goto(LOGIN_URL, wait_until="domcontentloaded")
            pausa_humana(0.6, 1.2)
            _llenar_cuit_y_avanzar(page, cred.cuit_login)
            _login_clave_fiscal(page, cred.clave_fiscal, cred.cuit_login)
            paso("login", "ok")

            paso("ventanilla", "en_curso")
            _log(on_log, "Abriendo Domicilio Fiscal Electrónico…")
            ve = _abrir_dfe(page)
            _cerrar_popup_dfe(ve, on_log)
            try:
                ve.bring_to_front()
            except Exception:
                pass
            usar_representados = _cuit_dfe_n(cred.cuit_representado) != _cuit_dfe_n(
                cred.cuit_login
            )
            if usar_representados:
                _log(
                    on_log,
                    f"DFE: consultando comunicaciones del representado "
                    f"{_cuit_dfe_fmt(_cuit_dfe_n(cred.cuit_representado))}…",
                )
                _configurar_dfe_representados(ve, cred.cuit_representado, on_log)
            try:
                resultado.razon_social = _razon_social_activa_mcmp(ve) or None
            except Exception:
                resultado.razon_social = None
            paso("ventanilla", "ok")

            paso("listar", "en_curso")
            filas = _preparar_listado_dfe(
                ve,
                fecha_desde,
                fecha_hasta,
                on_log,
                representados=usar_representados,
            )
            _log(on_log, f"Comunicaciones detectadas: {len(filas)}.")
            objetivo = _seleccionar_objetivo(
                filas, fecha_desde, fecha_hasta, representados=usar_representados
            )
            _log(on_log, f"A procesar: {len(objetivo)} comunicación(es).")
            paso("listar", "ok")

            paso("descargar", "en_curso")
            for idx, com in enumerate(objetivo, start=1):
                _log(on_log, f"[{idx}/{len(objetivo)}] {com.get('asunto','')[:60]} "
                             f"(N° {com['id']}, {com.get('recibido','')})")
                if usar_representados:
                    ok_det = _abrir_detalle_representado(ve, com)
                else:
                    ok_det = _abrir_detalle(ve, com["id"])
                if not ok_det:
                    _log(on_log, f"  • No se pudo abrir la comunicación {com['id']}.")
                    continue
                res_com = _descargar_o_imprimir(ve, com, carpeta_destino, on_log)
                resultado.comunicaciones.append(res_com)
                _volver_a_lista(ve, representados=usar_representados)
            paso("descargar", "ok")
            _log(on_log, f"Listo. Archivos guardados: {resultado.total_archivos} en {carpeta_destino}")

            return resultado

    except LoginArcaError:
        raise
    except CuitRepresentadoNoEncontradoError:
        raise
    except PlaywrightTimeout as exc:
        raise AutomatizacionArcaError(
            "Tiempo de espera agotado en ARCA (sitio lento o pantalla distinta a la esperada)."
        ) from exc
    except AutomatizacionArcaError:
        raise
    except Exception as exc:
        raise AutomatizacionArcaError(f"Error en automatización DFE: {exc}") from exc
    finally:
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass


def _fecha_de(texto: str) -> date | None:
    from cuit_en_arca.validacion import parsear_fecha_argentina

    try:
        return parsear_fecha_argentina(texto)
    except Exception:
        return None


_PLANTILLA_DFE_NOMBRE = "Formato DFE.xlsx"
_PLANTILLA_DFE_CARPETA = "Formato DFE"


def ruta_plantilla_dfe_excel() -> Path:
    """Ubica el modelo Excel de importación DFE (desarrollo y PyInstaller)."""
    candidatos: list[Path] = []
    if getattr(sys, "frozen", False):
        bundle = Path(getattr(sys, "_MEIPASS", ""))
        candidatos.extend(
            [
                bundle / _PLANTILLA_DFE_CARPETA / _PLANTILLA_DFE_NOMBRE,
                bundle / _PLANTILLA_DFE_NOMBRE,
            ]
        )
    raiz = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent
    dev_raiz = Path(__file__).resolve().parent.parent
    candidatos.extend(
        [
            raiz / _PLANTILLA_DFE_CARPETA / _PLANTILLA_DFE_NOMBRE,
            raiz / _PLANTILLA_DFE_NOMBRE,
            dev_raiz / _PLANTILLA_DFE_CARPETA / _PLANTILLA_DFE_NOMBRE,
            dev_raiz / _PLANTILLA_DFE_NOMBRE,
        ]
    )
    for p in candidatos:
        if p.is_file():
            return p
    return candidatos[0]


def ejecutar_dfe_lote(
    filas,
    *,
    headless: bool = False,
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
    """Procesa varias filas (CUIT) del DFE y guarda todo en ``DFE yyyy-mm-dd``.

    Si ``carpeta_base`` se indica, la carpeta ``DFE yyyy-mm-dd`` se crea allí;
    si no, en el escritorio. Tolerante: un CUIT con error no frena el resto.
    """
    base = carpeta_dfe_escritorio(
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
        if on_progreso:
            on_progreso(idx - 1, total, f"CUIT {cuit_repr} ({idx}/{total})")
        if on_reiniciar_pasos:
            on_reiniciar_pasos()

        cred = CredencialesArca(
            cuit_login=cuit_log,
            clave_fiscal=getattr(fila, "clave_fiscal", ""),
            cuit_representado=cuit_repr,
        )
        fd = _fecha_de(getattr(fila, "fecha_desde", "") or "")
        fh = _fecha_de(getattr(fila, "fecha_hasta", "") or "")

        # Subcarpeta por CUIT (evita choques de nombres entre contribuyentes).
        dest = base / _nombre_seguro(cuit_repr, fallback=cuit_log or f"cuit_{idx}")
        dest.mkdir(parents=True, exist_ok=True)

        try:
            res = ejecutar_descarga_dfe(
                cred,
                fd,
                fh,
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
                # marcar el paso actual como error visualmente
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

