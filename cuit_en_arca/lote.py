"""Descarga masiva Mis Comprobantes (emitidos + recibidos) desde planilla Excel."""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from cuit_en_arca.credenciales import CredencialesArca
from cuit_en_arca.empaquetado import (
    NOMBRE_ERRORES,
    SUBCARPETA_PROCESADOS,
    empaquetar_descargas,
)
from cuit_en_arca.errores import AutomatizacionNoDisponibleError, LoginArcaError
from cuit_en_arca.planilla_lote import (
    FilaPlanillaArca,
    leer_planilla_lote_con_errores,
)
from cuit_en_arca.service import _headless_desde_env, _requiere_playwright
from cuit_en_arca.stealth import pausa_entre_filas_lote
from cuit_en_arca.validacion import parsear_fecha_argentina

OnProgresoLote = Callable[[int, int, str, bool], None]
OnPasoLote = Callable[[str, str], None]
OnLogLote = Callable[[str], None]


def _log_lote(on_log: OnLogLote | None, msg: str) -> None:
    if on_log is None:
        return
    try:
        on_log(msg)
    except Exception:
        pass


@dataclass
class ResultadoLoteArca:
    contenido: bytes
    nombre_archivo: str
    mimetype: str
    total_filas: int
    descargas_ok: int
    ingresos_fallidos: list[str] = field(default_factory=list)
    advertencias: list[str] = field(default_factory=list)
    carpeta: str | None = None


def _carpeta_mis_comprobantes(
    base: str | Path,
    hoy: date | None = None,
    *,
    nombre_sesion: str | None = None,
) -> Path:
    """Carpeta ``Mis Comprobantes yyyy-mm-dd HH-MM`` dentro de ``base``."""
    from cuit_en_arca.carpetas_salida import momento_carpeta_ar, stamp_carpeta_ejecucion

    if nombre_sesion:
        destino = Path(base) / nombre_sesion
    else:
        destino = Path(base) / f"Mis Comprobantes {stamp_carpeta_ejecucion(momento_carpeta_ar(hoy))}"
    destino.mkdir(parents=True, exist_ok=True)
    return destino


def _nombre_seguro(cuit: str, tipo: str, nombre_sug: str) -> str:
    base = nombre_sug if nombre_sug else f"mis_comprobantes_{tipo}"
    if not base.lower().endswith((".xlsx", ".csv")):
        ext = ".csv" if ".csv" in base.lower() else ".xlsx"
        base = f"{base}{ext}"
    stem = base.rsplit(".", 1)[0]
    ext = base.rsplit(".", 1)[-1]
    return f"{cuit}_{tipo}_{stem}.{ext}"


def _mensaje_sin_descargas(
    advertencias: list[str],
    ingresos_fallidos: list[str],
) -> str:
    base = "No se obtuvo ninguna descarga."
    if advertencias:
        muestra = "; ".join(advertencias[:3])
        extra = f" (+{len(advertencias) - 3} más)" if len(advertencias) > 3 else ""
        return f"{base} Detalle: {muestra}{extra}"
    if ingresos_fallidos:
        muestra = "; ".join(ingresos_fallidos[:3])
        extra = f" (+{len(ingresos_fallidos) - 3} más)" if len(ingresos_fallidos) > 3 else ""
        return f"{base} Ingresos fallidos: {muestra}{extra}"
    return f"{base} Revisá la planilla (CUIT, clave y rango de fechas en columna D)."


def ejecutar_lote_arca(
    filas: list[FilaPlanillaArca],
    *,
    errores_planilla: list[str] | None = None,
    on_progreso: OnProgresoLote | None = None,
    on_paso: OnPasoLote | None = None,
    on_log: OnLogLote | None = None,
    on_reiniciar_pasos: Callable[[], None] | None = None,
    mapa_imputaciones: dict[str, tuple[str, str]] | None = None,
    carpeta_destino: str | Path | None = None,
    headless: bool | None = None,
    job_id: str | None = None,
    modo_ap: bool = False,
    nombre_carpeta_sesion: str | None = None,
) -> ResultadoLoteArca:
    _requiere_playwright()

    from cuit_en_arca.automation_playwright import ejecutar_descarga_mis_comprobantes
    from cuit_en_arca.resumen_cuit import (
        NOMBRE_SALIDA as NOMBRE_RESUMEN,
        ResumenCuitAcumulador,
        construir_resumen_cuit_xlsx,
    )

    from sumar_imp_total import procesar_comprobantes_a_excel_y_resumen

    total = len(filas)
    archivos: dict[str, bytes] = {}
    procesados: dict[str, bytes] = {}
    resumen_cuit = ResumenCuitAcumulador()

    # Modo carpeta: se escriben los archivos a disco a medida que se generan,
    # en «Mis Comprobantes yyyy-mm-dd» dentro de la carpeta elegida (sin .rar).
    carpeta = (
        _carpeta_mis_comprobantes(carpeta_destino, nombre_sesion=nombre_carpeta_sesion)
        if carpeta_destino
        else None
    )
    dir_proc = None
    if carpeta is not None:
        dir_proc = carpeta / SUBCARPETA_PROCESADOS
        dir_proc.mkdir(exist_ok=True)
        _log_lote(on_log, f"Carpeta de destino: {carpeta}")
    # Filas inválidas (CUIT/clave/fechas): se reportan, no frenan el lote.
    ingresos_fallidos: list[str] = list(errores_planilla or [])
    advertencias: list[str] = []
    descargas_ok = 0

    headless = _headless_desde_env() if headless is None else headless

    from cuit_en_arca.cancelacion import verificar_cancelacion

    _log_lote(on_log, f"Lote: {total} fila(s) a procesar.")
    if errores_planilla:
        _log_lote(on_log, f"Planilla: {len(errores_planilla)} fila(s) con error de formato.")

    for i, fila in enumerate(filas):
        if job_id:
            verificar_cancelacion(job_id)
        elif modo_ap:
            verificar_cancelacion(ap=True)
        if i > 0:
            pausa_entre_filas_lote()

        if on_reiniciar_pasos:
            on_reiniciar_pasos()

        if on_progreso:
            on_progreso(
                i + 1,
                total,
                f"CUIT {fila.cuit_representado} (fila {fila.fila_excel})…",
                False,
            )
        _log_lote(
            on_log,
            f"— Fila {i + 1}/{total} (Excel {fila.fila_excel}): "
            f"CUIT {fila.cuit_representado} · {fila.fecha_desde} – {fila.fecha_hasta} —",
        )

        cred = CredencialesArca(
            cuit_login=fila.cuit_login,
            clave_fiscal=fila.clave_fiscal,
            cuit_representado=fila.cuit_representado,
        )
        desde = parsear_fecha_argentina(fila.fecha_desde)
        hasta = parsear_fecha_argentina(fila.fecha_hasta)

        try:
            resultado = ejecutar_descarga_mis_comprobantes(
                cred,
                desde,
                hasta,
                headless=headless,
                tipo="ambos",
                on_paso=on_paso,
                on_log=on_log,
            )
        except LoginArcaError as exc:
            ingresos_fallidos.append(
                f"CUIT ingreso {fila.cuit_login} (fila {fila.fila_excel}): {exc}"
            )
            _log_lote(on_log, f"Ingreso fallido (fila {fila.fila_excel}): {exc}")
            if on_progreso:
                on_progreso(i + 1, total, f"Fila {fila.fila_excel}: ingreso fallido", True)
            continue
        except Exception as exc:
            advertencias.append(
                f"CUIT {fila.cuit_representado} (fila {fila.fila_excel}): {exc}"
            )
            _log_lote(on_log, f"Error en fila {fila.fila_excel}: {exc}")
            if on_progreso:
                on_progreso(i + 1, total, f"Fila {fila.fila_excel}: error", True)
            continue

        cuit = fila.cuit_representado
        razon_social_arca = (resultado.razon_social or "").strip()
        nuevos: list[tuple[str, bytes, bool]] = []  # (nombre, datos, emitidos)
        if resultado.emitidos:
            data_e, nom_e = resultado.emitidos
            nombre_e = _nombre_seguro(cuit, "emitidos", nom_e)
            archivos[nombre_e] = data_e
            nuevos.append((nombre_e, data_e, True))
            if carpeta is not None:
                (carpeta / nombre_e).write_bytes(data_e)
            _log_lote(on_log, f"  • Emitidos guardado: {nombre_e}")
        if resultado.recibidos:
            data_r, nom_r = resultado.recibidos
            nombre_r = _nombre_seguro(cuit, "recibidos", nom_r)
            archivos[nombre_r] = data_r
            nuevos.append((nombre_r, data_r, False))
            if carpeta is not None:
                (carpeta / nombre_r).write_bytes(data_r)
            _log_lote(on_log, f"  • Recibidos guardado: {nombre_r}")
        if resultado.aviso_parcial:
            advertencias.append(
                f"CUIT {cuit} (fila {fila.fila_excel}): {resultado.aviso_parcial}"
            )
        if resultado.emitidos or resultado.recibidos:
            descargas_ok += 1

        # Procesamiento automático de los archivos recién descargados.
        if nuevos:
            if on_paso:
                on_paso("procesamiento", "en_curso")
            _log_lote(on_log, "Procesando archivos descargados…")
            fallo_proc = False
            for nombre, datos, es_emit in nuevos:
                try:
                    # Las imputaciones contables, por ahora, solo a recibidos.
                    excel_proc, resumen = procesar_comprobantes_a_excel_y_resumen(
                        datos,
                        nombre,
                        emitidos=es_emit,
                        mapa_imputaciones=None if es_emit else mapa_imputaciones,
                    )
                    # El procesado siempre es un Excel, aunque el original sea .csv.
                    nombre_proc = f"{nombre.rsplit('.', 1)[0]}.xlsx"
                    procesados[nombre_proc] = excel_proc
                    if dir_proc is not None:
                        (dir_proc / nombre_proc).write_bytes(excel_proc)
                    _log_lote(on_log, f"  • Procesado: {nombre_proc}")
                    resumen_cuit.agregar(
                        cuit,
                        emitidos=es_emit,
                        # Prioridad: razón social leída de ARCA al seleccionar el
                        # CUIT; si no, la del propio archivo (emisor/receptor).
                        razon_social=razon_social_arca
                        or resumen.get("razon_social", ""),
                        por_mes=resumen.get("por_mes", {}),
                    )
                except Exception as exc:  # no abortar el lote por un archivo
                    fallo_proc = True
                    advertencias.append(
                        f"CUIT {cuit} (fila {fila.fila_excel}): "
                        f"no se pudo procesar «{nombre}»: {exc}"
                    )
            if on_paso:
                on_paso("procesamiento", "error" if fallo_proc else "ok")
            if fallo_proc:
                _log_lote(on_log, "Procesamiento completado con advertencias.")
            else:
                _log_lote(on_log, "Procesamiento completado.")

        if on_progreso:
            on_progreso(
                i + 1,
                total,
                f"Fila {fila.fila_excel} completada",
                True,
            )
        _log_lote(on_log, f"Fila {fila.fila_excel} completada.")

    lineas_errores = []
    if ingresos_fallidos:
        lineas_errores.append("=== Ingresos a ARCA no exitosos (CUIT / clave) ===")
        lineas_errores.extend(ingresos_fallidos)
        lineas_errores.append("")
    if advertencias:
        lineas_errores.append("=== Otros avisos ===")
        lineas_errores.extend(advertencias)
        lineas_errores.append("")
    if not lineas_errores:
        lineas_errores.append("Sin errores de ingreso registrados.")
    texto_errores = "\n".join(lineas_errores)

    if not archivos:
        # Si se procesaron filas pero ninguna descarga, igual se entrega el log.
        if not filas:
            raise AutomatizacionNoDisponibleError(
                _mensaje_sin_descargas(advertencias, ingresos_fallidos)
            )
        if not ingresos_fallidos and not advertencias:
            raise AutomatizacionNoDisponibleError(
                _mensaje_sin_descargas(advertencias, ingresos_fallidos)
            )

    extra_raiz: dict[str, bytes] = {}
    try:
        resumen_xlsx = construir_resumen_cuit_xlsx(resumen_cuit)
        if resumen_xlsx:
            extra_raiz[NOMBRE_RESUMEN] = resumen_xlsx
    except Exception as exc:  # el resumen no debe frenar la entrega
        advertencias.append(f"No se pudo generar «{NOMBRE_RESUMEN}»: {exc}")

    # Modo carpeta: ya se escribieron los archivos a medida que se generaron;
    # solo falta el resumen y el detalle de errores. No se genera .rar.
    if carpeta is not None:
        if extra_raiz.get(NOMBRE_RESUMEN):
            (carpeta / NOMBRE_RESUMEN).write_bytes(extra_raiz[NOMBRE_RESUMEN])
        (carpeta / NOMBRE_ERRORES).write_text(texto_errores, encoding="utf-8")
        _log_lote(on_log, f"Listo. Archivos en {carpeta}")
        return ResultadoLoteArca(
            contenido=b"",
            nombre_archivo=carpeta.name,
            mimetype="",
            total_filas=len(filas),
            descargas_ok=descargas_ok,
            ingresos_fallidos=ingresos_fallidos,
            advertencias=advertencias,
            carpeta=str(carpeta),
        )

    contenido, nombre, mime = empaquetar_descargas(
        archivos, texto_errores, procesados=procesados, extra=extra_raiz
    )

    return ResultadoLoteArca(
        contenido=contenido,
        nombre_archivo=nombre,
        mimetype=mime,
        total_filas=len(filas),
        descargas_ok=descargas_ok,
        ingresos_fallidos=ingresos_fallidos,
        advertencias=advertencias,
    )


def ejecutar_lote_planilla_arca(
    buf: io.BytesIO,
    *,
    on_progreso: OnProgresoLote | None = None,
    on_paso: OnPasoLote | None = None,
    on_log: OnLogLote | None = None,
    on_reiniciar_pasos: Callable[[], None] | None = None,
    mapa_imputaciones: dict[str, tuple[str, str]] | None = None,
    carpeta_destino: str | Path | None = None,
) -> ResultadoLoteArca:
    filas, errores = leer_planilla_lote_con_errores(buf)
    return ejecutar_lote_arca(
        filas,
        errores_planilla=errores,
        on_progreso=on_progreso,
        on_paso=on_paso,
        on_log=on_log,
        on_reiniciar_pasos=on_reiniciar_pasos,
        mapa_imputaciones=mapa_imputaciones,
        carpeta_destino=carpeta_destino,
    )
