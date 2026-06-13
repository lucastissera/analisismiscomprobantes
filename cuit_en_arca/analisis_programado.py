"""Programación y ejecución automática de análisis ARCA (Mis Comprobantes, DFE, Nuestra Parte)."""

from __future__ import annotations

import json
import logging
import os
import sys
import threading
import time
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

from cuit_en_arca.hora_log import ahora_ar_naive, fecha_hora_ar_texto

_LOG = logging.getLogger(__name__)

_lock = threading.Lock()
_ejecutando = False
_scheduler_iniciado = False

SISTEMAS_VALIDOS = frozenset({"mis_comprobantes", "dfe", "nuestra_parte"})

DIAS_SEMANA = (
    (0, "Lunes"),
    (1, "Martes"),
    (2, "Miércoles"),
    (3, "Jueves"),
    (4, "Viernes"),
    (5, "Sábado"),
    (6, "Domingo"),
)


def _directorio_datos() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # Web (Render/gunicorn): misma raíz que carpeta_ap_servidor (temp/aic_ap_data).
    if os.environ.get("CUIT_EN_ARCA_UI", "").strip() == "1":
        from cuit_en_arca.entrega_web import carpeta_ap_servidor

        return carpeta_ap_servidor().parent
    return Path(__file__).resolve().parent.parent


def _ruta_config() -> Path:
    return _directorio_datos() / "analisis_programado.json"


def ruta_plantilla_excel() -> Path:
    """Ubica el modelo Excel (desarrollo y PyInstaller / _MEIPASS)."""
    candidatos: list[Path] = []
    if getattr(sys, "frozen", False):
        bundle = Path(getattr(sys, "_MEIPASS", ""))
        candidatos.extend(
            [
                bundle / "Formato Analisis Programado.xlsx",
                bundle / "static" / "Formato Analisis Programado.xlsx",
            ]
        )
    raiz = _directorio_datos()
    dev_raiz = Path(__file__).resolve().parent.parent
    candidatos.extend(
        [
            raiz / "Formato Analisis Programado.xlsx",
            raiz / "static" / "Formato Analisis Programado.xlsx",
            dev_raiz / "Formato Analisis Programado.xlsx",
            dev_raiz / "static" / "Formato Analisis Programado.xlsx",
        ]
    )
    for p in candidatos:
        if p.is_file():
            return p
    return candidatos[0]


def _config_completa(cfg: ConfigAnalisisProgramado) -> bool:
    return bool(cfg.sistemas and cfg.carpeta_destino and cfg.filas)


@dataclass
class ConfigAnalisisProgramado:
    activo: bool = True
    dia_semana: int = 0  # 0=lunes … 6=domingo (datetime.weekday)
    hora: int = 9
    minuto: int = 0
    sistemas: list[str] = field(default_factory=list)
    carpeta_destino: str = ""
    filas: list[dict[str, Any]] = field(default_factory=list)
    ultima_ejecucion: str | None = None
    ultimo_resultado: dict[str, Any] | None = None

    def a_dict(self) -> dict[str, Any]:
        d = asdict(self)
        # No exponer claves fiscales en respuestas de estado (solo al guardar local).
        return d

    def a_dict_publico(self) -> dict[str, Any]:
        d = self.a_dict()
        filas_pub = []
        for f in d.get("filas") or []:
            filas_pub.append(
                {
                    "fila_excel": f.get("fila_excel"),
                    "cuit_login": f.get("cuit_login"),
                    "cuit_representado": f.get("cuit_representado"),
                    "tiene_clave": bool(f.get("clave_fiscal")),
                    "fechas_mis_comprobantes": f.get("fechas_mis_comprobantes"),
                    "fecha_dfe_desde": f.get("fecha_dfe_desde"),
                    "fecha_dfe_hasta": f.get("fecha_dfe_hasta"),
                    "ejercicio_nuestra_parte": f.get("ejercicio_nuestra_parte"),
                }
            )
        d["filas"] = filas_pub
        d["total_filas"] = len(filas_pub)
        d["programacion_lista"] = _config_completa(self)
        d["scheduler"] = scheduler_estado(self)
        return d


def cargar_config() -> ConfigAnalisisProgramado:
    ruta = _ruta_config()
    if not ruta.is_file():
        return ConfigAnalisisProgramado()
    try:
        data = json.loads(ruta.read_text(encoding="utf-8"))
        sistemas = [s for s in (data.get("sistemas") or []) if s in SISTEMAS_VALIDOS]
        return ConfigAnalisisProgramado(
            activo=bool(data.get("activo", True)),
            dia_semana=int(data.get("dia_semana", 0)),
            hora=int(data.get("hora", 9)),
            minuto=int(data.get("minuto", 0)),
            sistemas=sistemas,
            carpeta_destino=str(data.get("carpeta_destino") or ""),
            filas=list(data.get("filas") or []),
            ultima_ejecucion=data.get("ultima_ejecucion"),
            ultimo_resultado=data.get("ultimo_resultado"),
        )
    except Exception:
        return ConfigAnalisisProgramado()


def guardar_config(cfg: ConfigAnalisisProgramado) -> None:
    ruta = _ruta_config()
    ruta.write_text(
        json.dumps(cfg.a_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def limpiar_cache_programacion(cfg: ConfigAnalisisProgramado | None = None) -> ConfigAnalisisProgramado:
    """Desactiva la programación tras cancelación/error (sin borrar datos cargados)."""
    cfg = cfg or cargar_config()
    cfg.activo = False
    guardar_config(cfg)
    return cfg


def pausar_programacion_tras_ejecucion(
    cfg: ConfigAnalisisProgramado,
    *,
    ultimo_resultado: dict[str, Any] | None = None,
) -> ConfigAnalisisProgramado:
    """Marca la ejecución como hecha y pausa hasta que el usuario guarde de nuevo."""
    cfg.activo = False
    cfg.ultima_ejecucion = ahora_ar_naive().isoformat(timespec="seconds")
    if ultimo_resultado is not None:
        cfg.ultimo_resultado = ultimo_resultado
    guardar_config(cfg)
    return cfg


def registrar_ejecucion_manual(
    cfg: ConfigAnalisisProgramado,
    *,
    ultimo_resultado: dict[str, Any] | None = None,
) -> ConfigAnalisisProgramado:
    """Registra resultado sin desactivar la programación guardada."""
    cfg.ultima_ejecucion = ahora_ar_naive().isoformat(timespec="seconds")
    if ultimo_resultado is not None:
        cfg.ultimo_resultado = ultimo_resultado
    guardar_config(cfg)
    return cfg


def limpiar_configuracion_completa() -> ConfigAnalisisProgramado:
    """Borra toda la configuración guardada (cancela futuros análisis programados)."""
    cfg = ConfigAnalisisProgramado()
    guardar_config(cfg)
    return cfg


def _filas_desde_config(cfg: ConfigAnalisisProgramado):
    from cuit_en_arca.planilla_analisis_programado import FilaAnalisisProgramado

    out: list[FilaAnalisisProgramado] = []
    for i, raw in enumerate(cfg.filas, start=1):
        out.append(
            FilaAnalisisProgramado(
                fila_excel=int(raw.get("fila_excel") or i),
                cuit_login=str(raw.get("cuit_login") or ""),
                clave_fiscal=str(raw.get("clave_fiscal") or ""),
                cuit_representado=str(raw.get("cuit_representado") or raw.get("cuit_login") or ""),
                fechas_mis_comprobantes=str(raw.get("fechas_mis_comprobantes") or ""),
                fecha_dfe_desde=str(raw.get("fecha_dfe_desde") or ""),
                fecha_dfe_hasta=str(raw.get("fecha_dfe_hasta") or ""),
                ejercicio_nuestra_parte=str(raw.get("ejercicio_nuestra_parte") or ""),
            )
        )
    return out


def _carpeta_ejecucion(base: str | Path) -> Path:
    from cuit_en_arca.carpetas_salida import stamp_carpeta_ejecucion

    dest = Path(base) / f"Análisis Programado {stamp_carpeta_ejecucion()}"
    dest.mkdir(parents=True, exist_ok=True)
    return dest


def _motivo_espera_ejecucion(cfg: ConfigAnalisisProgramado, ahora: datetime) -> str:
    if not cfg.activo:
        return "La programación está pausada; guardá de nuevo para activarla."
    if not _config_completa(cfg):
        return "Faltan sistemas, datos o carpeta de destino."
    if ahora.weekday() != cfg.dia_semana:
        hoy = DIAS_SEMANA[ahora.weekday()][1]
        prog = DIAS_SEMANA[cfg.dia_semana][1]
        return f"Hoy es {hoy}; programaste {prog}."
    programado = ahora.replace(hour=cfg.hora, minute=cfg.minuto, second=0, microsecond=0)
    if ahora < programado:
        return (
            f"Esperando las {cfg.hora:02d}:{cfg.minuto:02d} "
            f"(ahora {ahora.strftime('%H:%M')} Argentina)."
        )
    if cfg.ultima_ejecucion:
        try:
            ult = datetime.fromisoformat(cfg.ultima_ejecucion)
            if ult.date() == ahora.date():
                return "Ya se ejecutó hoy."
        except Exception:
            pass
    return ""


def debe_ejecutar_ahora(cfg: ConfigAnalisisProgramado, ahora: datetime | None = None) -> bool:
    if not cfg.activo or not _config_completa(cfg):
        return False
    ahora = ahora or ahora_ar_naive()
    if ahora.weekday() != cfg.dia_semana:
        return False
    programado = ahora.replace(hour=cfg.hora, minute=cfg.minuto, second=0, microsecond=0)
    if ahora < programado:
        return False
    if cfg.ultima_ejecucion:
        try:
            ult = datetime.fromisoformat(cfg.ultima_ejecucion)
            if ult.date() == ahora.date():
                return False
        except Exception:
            pass
    return True


def ejecutar_analisis_programado(
    cfg: ConfigAnalisisProgramado | None = None,
    *,
    on_log: Callable[[str], None] | None = None,
    manual: bool = False,
    _reservado: bool = False,
) -> dict[str, Any]:
    """Ejecuta los sistemas seleccionados. Devuelve resumen del resultado."""
    global _ejecutando
    from cuit_en_arca.progreso_analisis_programado import (
        callback_log_ap,
        iniciar_ejecucion_ap,
        marcar_cancelado_ap,
        marcar_error_ap,
        marcar_ok_ap,
        marcar_paso_ap,
    )
    from cuit_en_arca.cancelacion import reset_cancelacion_ap, verificar_cancelacion
    from cuit_en_arca.errores import CancelacionUsuarioError

    if not _reservado:
        with _lock:
            if _ejecutando:
                return {"ok": False, "mensaje": "Ya hay una ejecución programada en curso."}
            _ejecutando = True

    cfg = cfg or cargar_config()
    activo_previo = cfg.activo
    resultado: dict[str, Any] = {
        "ok": True,
        "mensaje": "",
        "sistemas": {},
        "carpeta": "",
        "fallos": [],
    }

    log_prog = callback_log_ap()
    entrega_ref: list = [None]

    def log(msg: str) -> None:
        log_prog(msg)
        ent = entrega_ref[0]
        if ent is not None:
            from cuit_en_arca.entrega_web import log_indica_archivo_nuevo

            if log_indica_archivo_nuevo(msg):
                ent.escanear()
        if on_log:
            try:
                on_log(msg)
            except Exception:
                pass

    iniciar_ejecucion_ap(list(cfg.sistemas))
    reset_cancelacion_ap()

    try:
        if not cfg.carpeta_destino:
            raise ValueError("No hay carpeta de destino configurada.")
        if not cfg.sistemas:
            raise ValueError("No hay sistemas seleccionados.")
        if not cfg.filas:
            raise ValueError("No hay filas de datos cargadas.")

        marcar_paso_ap("preparacion", "en_curso")
        filas_ap = _filas_desde_config(cfg)
        base = _carpeta_ejecucion(cfg.carpeta_destino)
        resultado["carpeta"] = str(base)
        entrega = None
        if not getattr(sys, "frozen", False):
            from cuit_en_arca.entrega_web import EntregaWeb, make_registrar
            from cuit_en_arca.progreso_analisis_programado import agregar_archivo_ap

            entrega = EntregaWeb(base, make_registrar(agregar_archivo_ap))
            entrega_ref[0] = entrega
            entrega.escanear()
        log(f"Inicio análisis programado → {base}")
        if manual:
            log("Ejecución manual (Ejecutar ahora).")
        marcar_paso_ap("preparacion", "ok")

        from cuit_en_arca.planilla_analisis_programado import (
            filas_dfe,
            filas_mis_comprobantes,
            filas_nuestra_parte,
        )
        from cuit_en_arca.service import _headless_desde_env

        headless = _headless_desde_env()

        if "mis_comprobantes" in cfg.sistemas:
            verificar_cancelacion(ap=True)
            marcar_paso_ap("mis_comprobantes", "en_curso")
            mc, err_mc = filas_mis_comprobantes(filas_ap)
            resultado["sistemas"]["mis_comprobantes"] = {"filas": len(mc), "errores_planilla": err_mc}
            if mc:
                from cuit_en_arca.lote import ejecutar_lote_arca

                log(f"Mis Comprobantes: {len(mc)} fila(s)…")
                try:
                    res = ejecutar_lote_arca(
                        mc,
                        errores_planilla=err_mc,
                        carpeta_destino=base / "Mis Comprobantes",
                        headless=headless,
                        modo_ap=True,
                        on_log=log,
                    )
                    resultado["sistemas"]["mis_comprobantes"]["descargas_ok"] = res.descargas_ok
                    resultado["sistemas"]["mis_comprobantes"]["fallos"] = list(res.ingresos_fallidos)
                    resultado["fallos"].extend(res.ingresos_fallidos)
                    if res.advertencias:
                        resultado["fallos"].extend(res.advertencias)
                    marcar_paso_ap("mis_comprobantes", "ok")
                except CancelacionUsuarioError as exc:
                    marcar_cancelado_ap(str(exc))
                    if not manual:
                        limpiar_cache_programacion(cfg)
                    return resultado
                except Exception as exc:
                    resultado["sistemas"]["mis_comprobantes"]["error"] = str(exc)
                    resultado["fallos"].append(f"Mis Comprobantes: {exc}")
                    marcar_paso_ap("mis_comprobantes", "error")
            elif err_mc:
                resultado["fallos"].extend(err_mc)
                marcar_paso_ap("mis_comprobantes", "error")
            else:
                log("Mis Comprobantes: sin filas con fechas en la planilla.")
                marcar_paso_ap("mis_comprobantes", "ok")
            if entrega:
                entrega.escanear()

        if "dfe" in cfg.sistemas:
            verificar_cancelacion(ap=True)
            marcar_paso_ap("dfe", "en_curso")
            dfe, err_dfe = filas_dfe(filas_ap)
            resultado["sistemas"]["dfe"] = {"filas": len(dfe), "errores_planilla": err_dfe}
            if dfe:
                from cuit_en_arca.dfe_automation import ejecutar_dfe_lote

                log(f"DFE: {len(dfe)} fila(s)…")
                try:
                    carpeta_dfe = ejecutar_dfe_lote(
                        dfe,
                        headless=headless,
                        on_log=log,
                        carpeta_base=base / "DFE",
                        modo_ap=True,
                    )
                    resultado["sistemas"]["dfe"]["carpeta"] = str(carpeta_dfe)
                    marcar_paso_ap("dfe", "ok")
                except CancelacionUsuarioError as exc:
                    marcar_cancelado_ap(str(exc))
                    if not manual:
                        limpiar_cache_programacion(cfg)
                    return resultado
                except Exception as exc:
                    resultado["sistemas"]["dfe"]["error"] = str(exc)
                    resultado["fallos"].append(f"DFE: {exc}")
                    marcar_paso_ap("dfe", "error")
            elif err_dfe:
                resultado["fallos"].extend(err_dfe)
                marcar_paso_ap("dfe", "error")
            else:
                log("DFE: sin filas con fechas en la planilla.")
                marcar_paso_ap("dfe", "ok")
            if entrega:
                entrega.escanear()

        if "nuestra_parte" in cfg.sistemas:
            verificar_cancelacion(ap=True)
            marcar_paso_ap("nuestra_parte", "en_curso")
            np, err_np = filas_nuestra_parte(filas_ap)
            resultado["sistemas"]["nuestra_parte"] = {"filas": len(np), "errores_planilla": err_np}
            if np:
                from cuit_en_arca.nuestra_parte_automation import ejecutar_nuestra_parte_lote

                log(f"Nuestra Parte: {len(np)} fila(s)…")
                try:
                    carpeta_np = ejecutar_nuestra_parte_lote(
                        np,
                        headless=headless,
                        on_log=log,
                        carpeta_base=base / "Nuestra Parte",
                        modo_ap=True,
                    )
                    resultado["sistemas"]["nuestra_parte"]["carpeta"] = str(carpeta_np)
                    marcar_paso_ap("nuestra_parte", "ok")
                except CancelacionUsuarioError as exc:
                    marcar_cancelado_ap(str(exc))
                    if not manual:
                        limpiar_cache_programacion(cfg)
                    return resultado
                except Exception as exc:
                    resultado["sistemas"]["nuestra_parte"]["error"] = str(exc)
                    resultado["fallos"].append(f"Nuestra Parte: {exc}")
                    marcar_paso_ap("nuestra_parte", "error")
            elif err_np:
                resultado["fallos"].extend(err_np)
                marcar_paso_ap("nuestra_parte", "error")
            else:
                log("Nuestra Parte: sin filas con ejercicio en la planilla.")
                marcar_paso_ap("nuestra_parte", "ok")
            if entrega:
                entrega.escanear()

        from cuit_en_arca.fallos_arca import escribir_fallos_txt

        marcar_paso_ap("finalizado", "en_curso")
        if resultado["fallos"]:
            escribir_fallos_txt(base, otros=resultado["fallos"])
            if entrega:
                entrega.escanear()
            resultado["ok"] = any(
                s.get("descargas_ok", 0) > 0 or s.get("carpeta")
                for s in resultado["sistemas"].values()
            )
            resultado["mensaje"] = (
                f"Completado con {len(resultado['fallos'])} aviso(s)/fallo(s). "
                f"Revisá ingresos_fallidos.txt en la carpeta."
            )
        else:
            resultado["mensaje"] = "Análisis programado completado sin errores."

        log(resultado["mensaje"])
        marcar_ok_ap(
            carpeta=resultado["carpeta"],
            mensaje=resultado["mensaje"],
            fallos=resultado["fallos"],
        )
        resumen = {
            "ok": resultado["ok"],
            "mensaje": resultado["mensaje"],
            "carpeta": resultado["carpeta"],
        }
        if manual:
            cfg.activo = activo_previo
            registrar_ejecucion_manual(cfg, ultimo_resultado=resumen)
        else:
            pausar_programacion_tras_ejecucion(cfg, ultimo_resultado=resumen)
        return resultado

    except CancelacionUsuarioError as exc:
        marcar_cancelado_ap(str(exc))
        if not manual:
            limpiar_cache_programacion(cfg)
        return resultado
    except Exception as exc:
        resultado["ok"] = False
        resultado["mensaje"] = str(exc)
        log(f"ERROR: {exc}")
        marcar_error_ap(str(exc))
        if not manual:
            limpiar_cache_programacion(cfg)
        return resultado
    finally:
        with _lock:
            _ejecutando = False


_INTERVALO_SCHEDULER_CERCA = 15.0
_INTERVALO_SCHEDULER_LEJOS = 60.0


def _segundos_espera_scheduler(cfg: ConfigAnalisisProgramado) -> float:
    """Cuánto esperar hasta el próximo chequeo (≈15 s tras la hora programada)."""
    if not cfg.activo or not _config_completa(cfg):
        return _INTERVALO_SCHEDULER_LEJOS
    ahora = ahora_ar_naive()
    if ahora.weekday() != cfg.dia_semana:
        return _INTERVALO_SCHEDULER_LEJOS
    if cfg.ultima_ejecucion:
        try:
            if datetime.fromisoformat(cfg.ultima_ejecucion).date() == ahora.date():
                return _INTERVALO_SCHEDULER_LEJOS
        except Exception:
            pass
    programado = ahora.replace(hour=cfg.hora, minute=cfg.minuto, second=0, microsecond=0)
    if ahora >= programado:
        return _INTERVALO_SCHEDULER_CERCA
    delta = (programado - ahora).total_seconds()
    if delta <= _INTERVALO_SCHEDULER_CERCA:
        return max(1.0, delta)
    if delta <= 120:
        return _INTERVALO_SCHEDULER_CERCA
    return _INTERVALO_SCHEDULER_LEJOS


def lanzar_ejecucion_ap(
    cfg: ConfigAnalisisProgramado,
    *,
    manual: bool = False,
) -> tuple[bool, str]:
    """Ejecuta en un hilo de fondo. Devuelve (ok, mensaje_error)."""
    global _ejecutando

    with _lock:
        if _ejecutando:
            return False, "Ya hay una ejecución en curso."
        _ejecutando = True

    def _worker() -> None:
        try:
            ejecutar_analisis_programado(cfg, manual=manual, _reservado=True)
        except Exception:
            _LOG.exception("Error en ejecución de análisis programado")

    threading.Thread(
        target=_worker,
        daemon=True,
        name="ap-ejecucion-manual" if manual else "ap-ejecucion",
    ).start()
    return True, ""


def _loop_scheduler() -> None:
    while True:
        cfg: ConfigAnalisisProgramado | None = None
        try:
            cfg = cargar_config()
            if debe_ejecutar_ahora(cfg):
                _LOG.info(
                    "Disparando análisis programado (%s %02d:%02d)",
                    DIAS_SEMANA[cfg.dia_semana][1],
                    cfg.hora,
                    cfg.minuto,
                )
                lanzar_ejecucion_ap(cfg, manual=False)
        except Exception:
            _LOG.exception("Error en el scheduler de análisis programado")
        try:
            espera = _segundos_espera_scheduler(cfg or cargar_config())
        except Exception:
            espera = _INTERVALO_SCHEDULER_CERCA
        time.sleep(espera)


def scheduler_estado(cfg: ConfigAnalisisProgramado | None = None) -> dict[str, Any]:
    """Diagnóstico del scheduler (útil en web)."""
    cfg = cfg or cargar_config()
    ahora = ahora_ar_naive()
    dia_nombre = DIAS_SEMANA[cfg.dia_semana][1] if 0 <= cfg.dia_semana <= 6 else "?"
    debe = debe_ejecutar_ahora(cfg, ahora)
    motivo = "" if debe else _motivo_espera_ejecucion(cfg, ahora)
    return {
        "hilo_iniciado": _scheduler_iniciado,
        "config_activa": cfg.activo,
        "config_completa": _config_completa(cfg),
        "debe_ejecutar_ahora": debe,
        "motivo_espera": motivo,
        "zona_horaria": "America/Argentina/Buenos_Aires",
        "ahora_servidor": fecha_hora_ar_texto(),
        "dia_hoy": DIAS_SEMANA[ahora.weekday()][1],
        "programado": f"{dia_nombre} {cfg.hora:02d}:{cfg.minuto:02d}",
        "ultima_ejecucion": cfg.ultima_ejecucion,
    }


def iniciar_scheduler() -> None:
    global _scheduler_iniciado
    if _scheduler_iniciado:
        return
    _scheduler_iniciado = True
    t = threading.Thread(target=_loop_scheduler, daemon=True, name="analisis-programado")
    t.start()
    _LOG.info("Scheduler de análisis programado iniciado")
