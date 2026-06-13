import io
import os
import sys
import threading
import time
from datetime import timedelta
from pathlib import Path
from uuid import uuid4

_APP_ROOT = Path(__file__).resolve().parent
try:
    from dotenv import load_dotenv

    load_dotenv(_APP_ROOT / ".env")
    if getattr(sys, "frozen", False):
        load_dotenv(Path(sys.executable).resolve().parent / ".env")
except ImportError:
    pass

# Habilitar descarga ARCA por defecto (local, portable y servidor web).
# Desactivar solo con CUIT_EN_ARCA_PLAYWRIGHT=0
os.environ.setdefault("CUIT_EN_ARCA_PLAYWRIGHT", "1")
os.environ.setdefault("CUIT_EN_ARCA_UI", "1")

if not getattr(sys, "frozen", False):
    try:
        from cuit_en_arca.ensure_playwright import asegurar_chromium_playwright

        asegurar_chromium_playwright()
    except Exception:
        pass

if getattr(sys, "frozen", False):
    from cuit_en_arca.playwright_env import aplicar_entorno_playwright_portable

    aplicar_entorno_playwright_portable()

from flask import (
    Flask,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
    Response,
)

if getattr(sys, "frozen", False):
    _bundle = Path(getattr(sys, "_MEIPASS", _APP_ROOT))
    _tpl = _bundle / "templates"
    _static = _bundle / "static"
    app = Flask(
        __name__,
        root_path=str(_bundle),
        template_folder=str(_tpl),
        static_folder=str(_static),
    )
else:
    app = Flask(__name__)

from auth import (
    es_administrador,
    export_users_payload,
    iniciar_sincronizacion_usuarios,
    verificar_acceso,
    verificar_token_remoto,
    whatsapp_new_user_url,
)

try:
    iniciar_sincronizacion_usuarios()
except Exception:
    pass
from cuit_en_arca import ArcaProcesoError, CancelacionUsuarioError, ejecutar_lote_arca
from cuit_en_arca.planilla_lote import (
    leer_planilla_lote_con_errores,
    parsear_entrada_manual,
    parsear_entradas_manuales,
)
from cuit_en_arca.progreso_lote import (
    agregar_archivo_lote,
    callback_paso,
    callback_progreso,
    crear_job,
    marcar_error,
    marcar_cancelado,
    marcar_ok,
    obtener_job,
    reiniciar_pasos,
)
from cuit_en_arca.progreso_dfe import (
    agregar_archivo_dfe,
    agregar_resumen_cuit_dfe,
    callback_log_dfe,
    callback_paso_dfe,
    crear_job_dfe,
    marcar_error_dfe,
    marcar_cancelado_dfe,
    marcar_ok_dfe,
    obtener_job_dfe,
    progreso_cuit_dfe,
    reiniciar_pasos_dfe,
)
from cuit_en_arca.planilla_nuestra_parte import (
    leer_planilla_np_con_errores,
    parsear_entradas_manuales_np,
)
from cuit_en_arca.progreso_nuestra_parte import (
    agregar_archivo_np,
    agregar_resumen_cuit_np,
    callback_log_np,
    callback_paso_np,
    crear_job_np,
    marcar_error_np,
    marcar_cancelado_np,
    marcar_ok_np,
    obtener_job_np,
    progreso_cuit_np,
    reiniciar_pasos_np,
)
from i18n import (
    LANG_LABELS,
    MESES,
    SUPPORTED_LANGS,
    normalize_lang,
    tr,
    tr_js_bundle,
)
from plantillas_imputacion import (
    agregar_plantilla,
    eliminar_plantilla,
    leer_bytes_plantilla,
    listar_plantillas,
    plantillas_imputacion_disponibles,
    renombrar_plantilla,
    reemplazar_archivo_plantilla,
)

from sumar_imp_total import (
    COLUMNAS_A_AJUSTAR,
    COLUMNAS_DETALLE_SIN_RESUMEN,
    COLUMNAS_TOTAL_RESUMEN,
    enriquecer_contrapartes_con_imputacion,
    escribir_excel_informe_completo,
    escribir_excel_informe_dual,
    leer_mapa_imputaciones_desde_archivo,
    periodos_orden_crono,
    procesar_archivo,
    resumen_totales_por_imputacion,
    total_resumen_pantalla,
    totales_resumen_por_periodo,
)

app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-cambiar-en-produccion")
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(minutes=30)
app.config["SESSION_REFRESH_EACH_REQUEST"] = True
# download_id -> (bytes, nombre_archivo, mimetype)
DESCARGAS: dict[str, tuple[bytes, str, str]] = {}

from cuit_en_arca.entrega_web import init_descargas  # noqa: E402

init_descargas(DESCARGAS)

# Inactividad: sin peticiones al servidor durante este tiempo → cerrar sesión.
# Cada petición (refresco, nueva pestaña con la misma app, navegación) renueva el plazo.
_SESSION_IDLE_SEC = 30 * 60


def _es_app_escritorio() -> bool:
    return getattr(sys, "frozen", False)


def _nombre_carpeta_web_sesion(prefijo: str, raw: str | None = None) -> str | None:
    """Nombre de subcarpeta acordado con el navegador (web), p. ej. «Mis Comprobantes 2026-06-12 23-05»."""
    if _es_app_escritorio():
        return None
    if raw is None:
        raw = (request.form.get("web_carpeta_sesion") or "").strip()
    if not raw or not raw.startswith(f"{prefijo} "):
        return None
    if any(c in raw for c in "/\\") or ".." in raw:
        return None
    if len(raw) > 120:
        return None
    return raw


def _fabricar_entrega(
    job_id: str,
    carpeta_form: str | None,
    agregar_estado,
):
    from cuit_en_arca.entrega_web import EntregaWeb, carpeta_trabajo_web, make_registrar

    if _es_app_escritorio():
        p = (carpeta_form or "").strip()
        if not p:
            return None, None
        return Path(p), None
    base = carpeta_trabajo_web(job_id)
    return base, EntregaWeb(base, make_registrar(agregar_estado))


def _wrap_progreso_con_entrega(on_prog, entrega):
    if entrega is None or on_prog is None:
        return on_prog

    def _cb(actual, total, mensaje, fila_terminada=False):
        on_prog(actual, total, mensaje, fila_terminada)
        if fila_terminada:
            entrega.escanear()

    return _cb


def _safe_internal_path(target: str | None) -> str:
    if not target or not isinstance(target, str):
        return url_for("index")
    t = target.strip()
    if t.startswith("/") and not t.startswith("//"):
        return t
    return url_for("index")


@app.before_request
def _session_idle_and_login():
    if request.endpoint == "static" or (
        request.path and request.path.startswith("/static")
    ):
        return None

    now = time.time()
    username = session.get("user")
    if username:
        last = session.get("last_activity")
        if last is None:
            session["last_activity"] = now
            session.modified = True
        elif (now - float(last)) > _SESSION_IDLE_SEC:
            session.pop("user", None)
            session.pop("last_activity", None)
        else:
            session["last_activity"] = now
            session.modified = True

    if request.endpoint in ("login", "set_lang", "desktop_quit", "logout", "api_auth_users", None):
        return None
    if session.get("user"):
        return None
    return redirect(url_for("login", next=request.path))


def _entero_miles_punto(n: int) -> str:
    s = str(abs(int(n)))
    if len(s) <= 3:
        return s if n >= 0 else "-" + s
    partes = []
    while s:
        partes.append(s[-3:])
        s = s[:-3]
    out = ".".join(reversed(partes))
    return out if n >= 0 else "-" + out


@app.template_filter("fmt_ar")
def fmt_num_ar_argentina(value: object) -> str:
    """Miles con punto, decimales con coma (visualización en pantalla)."""
    try:
        x = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return str(value)
    neg = x < 0
    x = abs(x)
    centavos = int(round(x * 100 + 1e-9))
    ent = centavos // 100
    dec = centavos % 100
    body = f"{_entero_miles_punto(ent)},{dec:02d}"
    return f"-{body}" if neg else body


def _mostrar_ui_cuit_arca() -> bool:
    v = os.environ.get("CUIT_EN_ARCA_UI", "").strip().lower()
    return v in ("1", "true", "yes", "on")


@app.context_processor
def _inject_ui_flags():
    return {
        "mostrar_cuit_arca_ui": _mostrar_ui_cuit_arca(),
        "ejecutable_escritorio_frozen": getattr(sys, "frozen", False),
        "modo_escritorio": getattr(sys, "frozen", False),
    }


@app.context_processor
def _inject_i18n():
    lg = normalize_lang(session.get("lang"))

    def t(key: str, **kwargs):
        return tr(lg, key, **kwargs)

    return {
        "t": t,
        "current_lang": lg,
        "current_user": session.get("user"),
        "es_administrador": bool(session.get("es_admin")),
        "nombres_meses": MESES[lg],
        "langs": SUPPORTED_LANGS,
        "lang_labels": LANG_LABELS,
        "i18n_js": tr_js_bundle(lg),
    }


def _mapa_imputaciones_desde_peticion(
    lg: str,
) -> tuple[dict[str, tuple[str, str]] | None, str | None, bytes | None, str | None]:
    """
    Devuelve (mapa_cuit_imputacion | None, mensaje_error | None, bytes_archivo_si_subido, nombre_orig_archivo).
    """
    f_imp = request.files.get("excel_imputaciones")
    has_file = bool(
        f_imp and getattr(f_imp, "filename", None) and str(f_imp.filename).strip()
    )
    plantilla_id = (request.form.get("plantilla_imputacion_id") or "").strip()

    if has_file and plantilla_id and plantillas_imputacion_disponibles():
        return None, tr(lg, "err_imputacion_archivo_y_plantilla"), None, None

    if plantillas_imputacion_disponibles() and plantilla_id and not has_file:
        try:
            raw, nombre = leer_bytes_plantilla(plantilla_id)
            buf = io.BytesIO(raw)
            mapa = leer_mapa_imputaciones_desde_archivo(
                buf, nombre_archivo=nombre, ui_lang=lg
            )
            return mapa, None, None, None
        except FileNotFoundError:
            return None, tr(lg, "err_plantilla_imputacion_no_encontrada"), None, None
        except ValueError as exc:
            return None, str(exc), None, None

    if has_file:
        nombre_imp = Path(f_imp.filename).name
        nl = nombre_imp.lower()
        if not (nl.endswith(".xlsx") or nl.endswith(".csv")):
            return None, tr(lg, "err_only_xlsx_csv"), None, None
        datos = f_imp.read()
        buf_imp = io.BytesIO(datos)
        try:
            mapa = leer_mapa_imputaciones_desde_archivo(
                buf_imp, nombre_archivo=nombre_imp, ui_lang=lg
            )
            return mapa, None, datos, nombre_imp
        except ValueError as exc:
            return None, str(exc), None, None

    return None, None, None, None


@app.context_processor
def _inject_plantillas_imputacion():
    if plantillas_imputacion_disponibles():
        try:
            lista = listar_plantillas()
        except OSError:
            lista = []
    else:
        lista = []
    return {
        "plantillas_imputacion_ui": plantillas_imputacion_disponibles(),
        "plantillas_imputacion_lista": lista,
    }


@app.get("/set-lang/<code>")
def set_lang(code: str):
    session["lang"] = normalize_lang(code)
    nxt = request.args.get("next") or "/"
    if isinstance(nxt, str) and nxt.startswith("/") and not nxt.startswith("//"):
        return redirect(nxt)
    return redirect(url_for("index"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("user"):
        return redirect(_safe_internal_path(request.args.get("next")))
    if request.method == "POST":
        next_val = (request.form.get("next") or "").strip()
        user = (request.form.get("usuario") or "").strip()
        pwd = request.form.get("password") or ""
        motivo = verificar_acceso(user, pwd)
        if motivo is None:
            session["user"] = user
            session["es_admin"] = es_administrador(user)
            session["last_activity"] = time.time()
            session.permanent = True
            session.modified = True
            return redirect(_safe_internal_path(next_val or request.args.get("next")))
        lg = normalize_lang(session.get("lang"))
        return render_template(
            "login.html",
            login_error=motivo == "invalid",
            login_error_expired=motivo in ("expired", "not_yet"),
            login_error_msg=(
                tr(lg, "login_error_expired")
                if motivo in ("expired", "not_yet")
                else tr(lg, "login_error_bad")
            ),
            next=next_val,
            whatsapp_url=whatsapp_new_user_url(),
        )
    next_val = (request.args.get("next") or "").strip()
    return render_template(
        "login.html",
        next=next_val,
        whatsapp_url=whatsapp_new_user_url(),
    )


def _limpiar_sesion_flask() -> None:
    """Cierra sesión Flask conservando solo el idioma elegido."""
    lang = session.get("lang")
    session.clear()
    if lang:
        session["lang"] = lang
    session.modified = True


def _aplicar_borrado_cookie_sesion(resp: Response) -> Response:
    resp.delete_cookie(
        app.config.get("SESSION_COOKIE_NAME", "session"),
        path=app.config.get("SESSION_COOKIE_PATH") or "/",
    )
    return resp


def _iniciar_cierre_proceso_desktop() -> None:
    """Cierra navegador (modo app), borra cookies locales y termina el .exe."""

    def _salir() -> None:
        time.sleep(0.6)
        try:
            from cuit_en_arca.browser_desktop import (
                cerrar_navegador_desktop,
                limpiar_cookies_localhost,
            )

            cerrar_navegador_desktop()
            limpiar_cookies_localhost()
        except Exception:
            pass
        os._exit(0)

    threading.Thread(target=_salir, daemon=True).start()


def _respuesta_cierre_desktop() -> Response:
    """Cierra el proceso; la ventana del navegador se termina por PID/perfil."""
    _limpiar_sesion_flask()
    resp = Response("", mimetype="text/html; charset=utf-8")
    _aplicar_borrado_cookie_sesion(resp)
    _iniciar_cierre_proceso_desktop()
    return resp


@app.get("/logout")
def logout():
    _limpiar_sesion_flask()
    if getattr(sys, "frozen", False):
        return _respuesta_cierre_desktop()
    resp = redirect(url_for("login"))
    return _aplicar_borrado_cookie_sesion(resp)


@app.route("/desktop-quit", methods=["GET", "POST"])
def desktop_quit():
    """Solo .exe local: cierra el proceso (sin consola no hay otra forma obvia de salir)."""
    if not getattr(sys, "frozen", False):
        abort(404)
    ra = (request.remote_addr or "").replace("::ffff:", "")
    if ra not in ("127.0.0.1", "::1"):
        abort(403)
    _limpiar_sesion_flask()
    return _respuesta_cierre_desktop()


MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@app.get("/")
def index():
    return render_template("inicio.html")


@app.get("/procesador")
def procesador():
    return render_template("index.html")


@app.get("/descargar/<download_id>")
def descargar(download_id: str):
    item = DESCARGAS.get(download_id)
    if not item:
        lg = normalize_lang(session.get("lang"))
        return render_template("index.html", error=tr(lg, "err_download_gone"))

    contenido, nombre_salida, mime = item
    return send_file(
        io.BytesIO(contenido),
        as_attachment=True,
        download_name=nombre_salida,
        mimetype=mime,
    )


@app.post("/procesar")
def procesar():
    lg = normalize_lang(session.get("lang"))
    f_rec = request.files.get("excel_recibidos")
    f_emit = request.files.get("excel_emitidos")
    has_r = bool(f_rec and (f_rec.filename or "").strip())
    has_e = bool(f_emit and (f_emit.filename or "").strip())
    if not has_r and not has_e:
        return render_template("index.html", error=tr(lg, "err_select_file"))

    def _ext_ok(n: str) -> bool:
        nl = n.lower()
        return nl.endswith(".xlsx") or nl.endswith(".csv")

    mapa_imputaciones, err_imp, datos_imp_bytes, imp_nombre_orig = (
        _mapa_imputaciones_desde_peticion(lg)
    )
    if err_imp:
        return render_template("index.html", error=err_imp)

    nombre_guardar = (request.form.get("nombre_nueva_plantilla_imputacion") or "").strip()
    if (
        plantillas_imputacion_disponibles()
        and nombre_guardar
        and datos_imp_bytes is not None
        and imp_nombre_orig
    ):
        try:
            agregar_plantilla(nombre_guardar, datos_imp_bytes, imp_nombre_orig)
            flash(tr(lg, "flash_plantilla_guardada_ok", nombre=nombre_guardar), "success")
        except ValueError as exc:
            if str(exc) == "nombre_duplicado":
                flash(tr(lg, "flash_plantilla_nombre_duplicado"), "warning")
            elif str(exc) == "nombre_vacio":
                pass

    con_cols_imp = mapa_imputaciones is not None

    if has_r and has_e:
        nombre_r = Path(f_rec.filename).name
        nombre_e = Path(f_emit.filename).name
        if not _ext_ok(nombre_r) or not _ext_ok(nombre_e):
            return render_template("index.html", error=tr(lg, "err_only_xlsx_csv"))
        try:
            buf_r = io.BytesIO(f_rec.read())
            buf_e = io.BytesIO(f_emit.read())
            (
                df_r,
                tot_r,
                tpp_r,
                nce_r,
                tabla_r,
            ) = procesar_archivo(
                buf_r,
                0,
                nombre_archivo=nombre_r,
                ui_lang=lg,
                emitidos=False,
            )
            (
                df_e,
                tot_e,
                tpp_e,
                nce_e,
                tabla_e,
            ) = procesar_archivo(
                buf_e,
                0,
                nombre_archivo=nombre_e,
                ui_lang=lg,
                emitidos=True,
            )
        except ValueError as exc:
            return render_template("index.html", error=str(exc))
        except Exception as exc:
            return render_template(
                "index.html", error=tr(lg, "err_processing", exc=exc)
            )

        tabla_r = enriquecer_contrapartes_con_imputacion(tabla_r, mapa_imputaciones)
        tabla_e = enriquecer_contrapartes_con_imputacion(tabla_e, mapa_imputaciones)
        res_imp_r = (
            resumen_totales_por_imputacion(tabla_r) if con_cols_imp else None
        )
        res_imp_e = (
            resumen_totales_por_imputacion(tabla_e) if con_cols_imp else None
        )

        per_r = periodos_orden_crono(
            tpp_r,
            nce_r.get("neto_nc_por_periodo", {}),
            nce_r.get("iva_nc_por_periodo", {}),
        )
        per_e = periodos_orden_crono(
            tpp_e,
            nce_e.get("neto_nc_por_periodo", {}),
            nce_e.get("iva_nc_por_periodo", {}),
        )
        tres_r = {c: tot_r[c] for c in COLUMNAS_TOTAL_RESUMEN}
        tdet_r = {c: tot_r[c] for c in COLUMNAS_DETALLE_SIN_RESUMEN}
        tres_e = {c: tot_e[c] for c in COLUMNAS_TOTAL_RESUMEN}
        tdet_e = {c: tot_e[c] for c in COLUMNAS_DETALLE_SIN_RESUMEN}

        salida = io.BytesIO()
        escribir_excel_informe_dual(
            salida,
            df_recibidos=df_r,
            totales_por_periodo_rec=tpp_r,
            periodos_orden_rec=per_r,
            notas_credito_extras_rec=nce_r,
            totales_resumen_rec=tres_r,
            totales_detalle_rec=tdet_r,
            suma_total_rec=round(total_resumen_pantalla(tot_r), 2),
            tabla_contrapartes_rec=tabla_r,
            df_emitidos=df_e,
            totales_por_periodo_emit=tpp_e,
            periodos_orden_emit=per_e,
            notas_credito_extras_emit=nce_e,
            totales_resumen_emit=tres_e,
            totales_detalle_emit=tdet_e,
            suma_total_emit=round(total_resumen_pantalla(tot_e), 2),
            tabla_contrapartes_emit=tabla_e,
            columnas_orden=COLUMNAS_A_AJUSTAR,
            resumen_imputacion_rec=res_imp_r,
            resumen_imputacion_emit=res_imp_e,
            con_columnas_imputacion_en_contrapartes=con_cols_imp,
            mapa_imputaciones=mapa_imputaciones,
        )
        contenido = salida.getvalue()
        nombre_salida = f"{Path(nombre_r).stem}_{Path(nombre_e).stem}_ajustado.xlsx"
        download_id = uuid4().hex
        DESCARGAS[download_id] = (contenido, nombre_salida, MIME_XLSX)

        return render_template(
            "index.html",
            mostrar_resultado=True,
            procesamiento_dual=True,
            totales_resumen_recibidos=tres_r,
            totales_detalle_recibidos=tdet_r,
            suma_total_recibidos=round(total_resumen_pantalla(tot_r), 2),
            totales_resumen_emitidos=tres_e,
            totales_detalle_emitidos=tdet_e,
            suma_total_emitidos=round(total_resumen_pantalla(tot_e), 2),
            columnas_orden=COLUMNAS_A_AJUSTAR,
            totales_por_periodo_recibidos=tpp_r,
            periodos_orden_recibidos=per_r,
            resumen_total_periodo_recibidos=totales_resumen_por_periodo(tpp_r),
            total_neto_nc_recibidos=nce_r["total_neto_nc"],
            total_iva_nc_recibidos=nce_r["total_iva_nc"],
            neto_nc_por_periodo_recibidos=nce_r["neto_nc_por_periodo"],
            iva_nc_por_periodo_recibidos=nce_r["iva_nc_por_periodo"],
            totales_por_periodo_emitidos=tpp_e,
            periodos_orden_emitidos=per_e,
            resumen_total_periodo_emitidos=totales_resumen_por_periodo(tpp_e),
            total_neto_nc_emitidos=nce_e["total_neto_nc"],
            total_iva_nc_emitidos=nce_e["total_iva_nc"],
            neto_nc_por_periodo_emitidos=nce_e["neto_nc_por_periodo"],
            iva_nc_por_periodo_emitidos=nce_e["iva_nc_por_periodo"],
            tabla_contrapartes_recibidos=tabla_r,
            tabla_contrapartes_emitidos=tabla_e,
            download_id=download_id,
            nombre_salida=nombre_salida,
            imputacion_activa=con_cols_imp,
            resumen_imputacion_recibidos=res_imp_r,
            resumen_imputacion_emitidos=res_imp_e,
        )

    emitidos = bool(has_e)
    archivo = f_emit if emitidos else f_rec
    nombre = Path(archivo.filename).name
    if not _ext_ok(nombre):
        return render_template("index.html", error=tr(lg, "err_only_xlsx_csv"))

    try:
        datos = archivo.read()
        buffer = io.BytesIO(datos)
        (
            df_ajustado,
            totales,
            totales_por_periodo,
            notas_credito_extras,
            tabla_contrapartes,
        ) = procesar_archivo(
            buffer,
            0,
            nombre_archivo=nombre,
            ui_lang=lg,
            emitidos=emitidos,
        )
    except ValueError as exc:
        return render_template("index.html", error=str(exc))
    except Exception as exc:  # fallback para errores no esperados
        return render_template(
            "index.html", error=tr(lg, "err_processing", exc=exc)
        )

    tabla_contrapartes = enriquecer_contrapartes_con_imputacion(
        tabla_contrapartes,
        mapa_imputaciones,
    )
    res_imp = (
        resumen_totales_por_imputacion(tabla_contrapartes) if con_cols_imp else None
    )

    salida = io.BytesIO()
    periodos_orden = periodos_orden_crono(
        totales_por_periodo,
        notas_credito_extras.get("neto_nc_por_periodo", {}),
        notas_credito_extras.get("iva_nc_por_periodo", {}),
    )
    totales_resumen = {c: totales[c] for c in COLUMNAS_TOTAL_RESUMEN}
    totales_detalle = {c: totales[c] for c in COLUMNAS_DETALLE_SIN_RESUMEN}
    escribir_excel_informe_completo(
        df_ajustado,
        salida,
        emitidos=emitidos,
        totales=totales,
        totales_por_periodo=totales_por_periodo,
        periodos_orden=periodos_orden,
        notas_credito_extras=notas_credito_extras,
        totales_resumen=totales_resumen,
        totales_detalle=totales_detalle,
        suma_total=round(total_resumen_pantalla(totales), 2),
        columnas_orden=COLUMNAS_A_AJUSTAR,
        tabla_contrapartes=tabla_contrapartes,
        resumen_imputacion=res_imp,
        con_columnas_imputacion_en_contrapartes=con_cols_imp,
        mapa_imputaciones=mapa_imputaciones,
    )
    contenido = salida.getvalue()

    nombre_salida = f"{Path(nombre).stem}_ajustado.xlsx"
    download_id = uuid4().hex
    DESCARGAS[download_id] = (contenido, nombre_salida, MIME_XLSX)

    resumen_total_periodo = totales_resumen_por_periodo(totales_por_periodo)

    return render_template(
        "index.html",
        mostrar_resultado=True,
        procesamiento_dual=False,
        emitidos=emitidos,
        totales_resumen=totales_resumen,
        totales_detalle=totales_detalle,
        columnas_orden=COLUMNAS_A_AJUSTAR,
        suma_total=round(total_resumen_pantalla(totales), 2),
        totales_por_periodo=totales_por_periodo,
        periodos_orden=periodos_orden,
        resumen_total_periodo=resumen_total_periodo,
        total_neto_nc=notas_credito_extras["total_neto_nc"],
        total_iva_nc=notas_credito_extras["total_iva_nc"],
        neto_nc_por_periodo=notas_credito_extras["neto_nc_por_periodo"],
        iva_nc_por_periodo=notas_credito_extras["iva_nc_por_periodo"],
        tabla_contrapartes=tabla_contrapartes,
        download_id=download_id,
        nombre_salida=nombre_salida,
        imputacion_activa=con_cols_imp,
        resumen_imputacion=res_imp,
    )


@app.route("/plantillas-imputaciones", methods=["GET", "POST"])
def plantillas_imputaciones():
    if not plantillas_imputacion_disponibles():
        abort(404)
    lg = normalize_lang(session.get("lang"))
    if request.method == "POST":
        accion = (request.form.get("accion") or "").strip()
        pid = (request.form.get("plantilla_id") or "").strip()
        try:
            if accion == "renombrar":
                nuevo = (request.form.get("nuevo_nombre") or "").strip()
                renombrar_plantilla(pid, nuevo)
                flash(tr(lg, "flash_plantilla_renombrada"), "success")
            elif accion == "reemplazar":
                f_rep = request.files.get("nuevo_archivo")
                fn = (
                    (getattr(f_rep, "filename", None) or "").strip()
                    if f_rep
                    else ""
                )
                if not f_rep or not fn:
                    flash(tr(lg, "err_plantilla_archivo_falta"), "warning")
                else:
                    nl = fn.lower()
                    if not (nl.endswith(".xlsx") or nl.endswith(".csv")):
                        flash(tr(lg, "err_only_xlsx_csv"), "warning")
                    else:
                        reemplazar_archivo_plantilla(
                            pid, f_rep.read(), Path(fn).name
                        )
                        flash(tr(lg, "flash_plantilla_archivo_ok"), "success")
            elif accion == "eliminar":
                eliminar_plantilla(pid)
                flash(tr(lg, "flash_plantilla_eliminada"), "success")
            else:
                flash(tr(lg, "err_plantilla_accion"), "warning")
        except ValueError as exc:
            code = str(exc)
            if code == "nombre_duplicado":
                flash(tr(lg, "flash_plantilla_nombre_duplicado"), "warning")
            elif code == "nombre_vacio":
                flash(tr(lg, "err_plantilla_nombre_vacio"), "warning")
            elif code == "no_existe":
                flash(tr(lg, "err_plantilla_no_existe"), "warning")
            else:
                flash(tr(lg, "err_plantilla_generico"), "warning")
        return redirect(url_for("plantillas_imputaciones"))
    try:
        plantillas = listar_plantillas()
    except OSError:
        plantillas = []
    return render_template(
        "plantillas_imputaciones.html",
        plantillas=plantillas,
    )


def _filas_arca_desde_peticion(
    lg: str,
) -> tuple[list, list[str], str | None]:
    """Devuelve (filas, errores_parciales, mensaje_error | None)."""
    planilla = request.files.get("planilla_arca")
    has_file = bool(
        planilla and getattr(planilla, "filename", None) and str(planilla.filename).strip()
    )

    if has_file:
        if not Path(planilla.filename).name.lower().endswith(".xlsx"):
            return [], [], tr(lg, "err_arca_xlsx")
        try:
            filas, errores = leer_planilla_lote_con_errores(
                io.BytesIO(planilla.read())
            )
        except ArcaProcesoError as exc:
            return [], [], str(exc)
        if not filas:
            msg = "; ".join(errores) or tr(lg, "err_arca_xlsx")
            return [], errores, msg
        return filas, errores, None

    cuits_login = request.form.getlist("arca_cuit_login")
    claves = request.form.getlist("arca_clave_fiscal")
    cuits_repr = request.form.getlist("arca_cuit_representado")
    rangos = request.form.getlist("arca_rango_fechas")

    hay_algo = any(
        (v or "").strip()
        for lista in (cuits_login, claves, cuits_repr, rangos)
        for v in lista
    )
    if not hay_algo:
        return [], [], tr(lg, "err_arca_sin_datos")

    filas, errores = parsear_entradas_manuales(
        cuits_login, claves, cuits_repr, rangos
    )
    if not filas:
        msg = "; ".join(errores) or tr(lg, "err_arca_manual_incompleto")
        return [], errores, msg
    return filas, errores, None


@app.get("/arca-descarga-lote/plantilla")
def arca_plantilla():
    from cuit_en_arca.plantillas_importacion import ruta_plantilla_arca_excel

    ruta = ruta_plantilla_arca_excel()
    if not ruta.is_file():
        abort(404)
    return send_file(
        ruta,
        as_attachment=True,
        download_name="Formato Analisis Comprobantes.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/arca-descarga-lote")
def arca_descarga_lote():
    lg = normalize_lang(session.get("lang"))
    if not _mostrar_ui_cuit_arca():
        if request.headers.get("X-Requested-With") == "fetch":
            return jsonify({"error": tr(lg, "err_arca_disabled")}), 403
        return (
            render_template(
                "index.html",
                error=tr(lg, "err_arca_disabled"),
            ),
            403,
        )

    filas, _errores_planilla, err_msg = _filas_arca_desde_peticion(lg)
    if err_msg:
        if request.headers.get("X-Requested-With") == "fetch":
            return jsonify({"error": err_msg}), 400
        return render_template("index.html", error=err_msg)

    # Imputación contable opcional: si se adjuntó un Excel en la solapa de
    # imputación o se eligió una plantilla guardada, se aplica al lote
    # (solo a comprobantes recibidos). Si no, el lote se procesa sin imputar.
    mapa_imputaciones, err_imp, _datos_imp, _imp_nom = (
        _mapa_imputaciones_desde_peticion(lg)
    )
    if err_imp:
        if request.headers.get("X-Requested-With") == "fetch":
            return jsonify({"error": err_imp}), 400
        return render_template("index.html", error=err_imp)

    carpeta_form = (request.form.get("carpeta_destino") or "").strip() or None

    job_id = uuid4().hex
    base, entrega = _fabricar_entrega(
        job_id,
        carpeta_form,
        lambda did, rel, nom: agregar_archivo_lote(job_id, did, rel, nom),
    )
    if base is None:
        msg = tr(lg, "carpeta_cancelada")
        if request.headers.get("X-Requested-With") == "fetch":
            return jsonify({"error": msg}), 400
        return render_template("index.html", error=msg)
    carpeta_destino = str(base)
    nombre_sesion_mc = _nombre_carpeta_web_sesion("Mis Comprobantes")

    def _err_inesperado(exc: Exception) -> str:
        return tr(lg, "err_arca_unexpected", exc=exc)

    from cuit_en_arca.cancelacion import reset_cancelacion

    reset_cancelacion(job_id)
    crear_job(job_id, len(filas))
    on_prog = _wrap_progreso_con_entrega(callback_progreso(job_id), entrega)
    on_paso = callback_paso(job_id)

    def _on_reiniciar() -> None:
        reiniciar_pasos(job_id)

    def _worker() -> None:
        try:
            resultado = ejecutar_lote_arca(
                filas,
                errores_planilla=_errores_planilla,
                on_progreso=on_prog,
                on_paso=on_paso,
                on_reiniciar_pasos=_on_reiniciar,
                mapa_imputaciones=mapa_imputaciones,
                carpeta_destino=carpeta_destino,
                job_id=job_id,
                nombre_carpeta_sesion=nombre_sesion_mc,
            )
            if entrega:
                entrega.escanear()
            if resultado.carpeta:
                # Modo carpeta: los archivos ya están en disco, sin descarga.
                fallos = list(resultado.ingresos_fallidos) + list(resultado.advertencias)
                marcar_ok(
                    job_id,
                    nombre_archivo=resultado.nombre_archivo,
                    carpeta=resultado.carpeta,
                    descargas_ok=resultado.descargas_ok,
                    ingresos_fallidos=len(resultado.ingresos_fallidos),
                    fallos_detalle=fallos,
                )
                return
            did = uuid4().hex
            DESCARGAS[did] = (
                resultado.contenido,
                resultado.nombre_archivo,
                resultado.mimetype,
            )
            marcar_ok(
                job_id,
                download_id=did,
                nombre_archivo=resultado.nombre_archivo,
                descargas_ok=resultado.descargas_ok,
                ingresos_fallidos=len(resultado.ingresos_fallidos),
                fallos_detalle=list(resultado.ingresos_fallidos) + list(resultado.advertencias),
            )
        except CancelacionUsuarioError as exc:
            marcar_cancelado(job_id, str(exc))
        except ArcaProcesoError as exc:
            marcar_error(job_id, str(exc))
        except Exception as exc:
            marcar_error(job_id, _err_inesperado(exc))

    threading.Thread(target=_worker, daemon=True).start()

    if request.headers.get("X-Requested-With") == "fetch":
        return jsonify({"job_id": job_id, "total": len(filas)})

    return render_template(
        "index.html",
        arca_job_id=job_id,
        arca_job_total=len(filas),
    )


@app.get("/arca-lote-estado/<job_id>")
def arca_lote_estado(job_id: str):
    estado = obtener_job(job_id)
    if estado is None:
        return jsonify({"error": "job_not_found"}), 404
    return jsonify(estado)


# --------------------------------------------------------------------------- #
# Domicilio Fiscal Electrónico (Ventanilla Electrónica)
# --------------------------------------------------------------------------- #
def _filas_dfe_desde_peticion(lg: str):
    """Devuelve (filas, errores_planilla, mensaje_error)."""
    f = request.files.get("dfe_excel")
    has_file = bool(f and getattr(f, "filename", None) and str(f.filename).strip())
    if has_file:
        nombre = Path(f.filename).name
        if not nombre.lower().endswith(".xlsx"):
            return [], [], tr(lg, "err_only_xlsx_csv")
        try:
            filas, errores = leer_planilla_lote_con_errores(io.BytesIO(f.read()))
        except ArcaProcesoError as exc:
            return [], [], str(exc)
        if not filas:
            return [], errores, "; ".join(errores) or tr(lg, "dfe_err_sin_datos")
        return filas, errores, None

    cuits = request.form.getlist("dfe_cuit_login")
    claves = request.form.getlist("dfe_clave_fiscal")
    reprs = request.form.getlist("dfe_cuit_representado")
    desdes = request.form.getlist("dfe_fecha_desde")
    hastas = request.form.getlist("dfe_fecha_hasta")

    hay_algo = any(
        (v or "").strip()
        for lista in (cuits, claves, reprs, desdes, hastas)
        for v in lista
    )
    if not hay_algo:
        return [], [], tr(lg, "dfe_err_sin_datos")

    def _at(lista, i):
        return (lista[i] if i < len(lista) else "").strip()

    n = max(len(cuits), len(claves), len(reprs), len(desdes), len(hastas))
    rangos = []
    for i in range(n):
        d = _at(desdes, i)
        h = _at(hastas, i)
        rangos.append(f"{d} - {h}" if (d or h) else "")

    filas, errores = parsear_entradas_manuales(cuits, claves, reprs, rangos)
    if not filas:
        return [], errores, "; ".join(errores) or tr(lg, "dfe_err_manual_incompleto")
    return filas, errores, None


@app.get("/domicilio-fiscal")
def domicilio_fiscal():
    return render_template("dfe.html")


@app.get("/domicilio-fiscal/plantilla")
def dfe_plantilla():
    from cuit_en_arca.dfe_automation import ruta_plantilla_dfe_excel

    ruta = ruta_plantilla_dfe_excel()
    if not ruta.is_file():
        abort(404)
    return send_file(
        ruta,
        as_attachment=True,
        download_name="Formato DFE.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/dfe-descargar")
def dfe_descargar():
    lg = normalize_lang(session.get("lang"))
    es_fetch = request.headers.get("X-Requested-With") == "fetch"

    if not _mostrar_ui_cuit_arca():
        if es_fetch:
            return jsonify({"error": tr(lg, "err_arca_disabled")}), 403
        return render_template("dfe.html", error=tr(lg, "err_arca_disabled")), 403

    filas, _errores, err_msg = _filas_dfe_desde_peticion(lg)
    if err_msg:
        if es_fetch:
            return jsonify({"error": err_msg}), 400
        return render_template("dfe.html", error=err_msg)

    from cuit_en_arca.dfe_automation import ejecutar_dfe_lote
    from cuit_en_arca.service import _headless_desde_env

    headless = _headless_desde_env()

    carpeta_form = (request.form.get("carpeta_destino") or "").strip() or None

    job_id = uuid4().hex
    base, entrega = _fabricar_entrega(
        job_id,
        carpeta_form,
        lambda did, rel, nom: agregar_archivo_dfe(job_id, did, rel, nom),
    )
    if base is None:
        msg = tr(lg, "carpeta_cancelada")
        if es_fetch:
            return jsonify({"error": msg}), 400
        return render_template("dfe.html", error=msg)
    carpeta_destino = str(base)
    nombre_sesion_dfe = _nombre_carpeta_web_sesion("DFE")

    def _err_inesperado(exc: Exception) -> str:
        return tr(lg, "err_arca_unexpected", exc=exc)

    from cuit_en_arca.cancelacion import reset_cancelacion

    reset_cancelacion(job_id)
    crear_job_dfe(job_id, len(filas))
    reiniciar_pasos_dfe(job_id)
    on_log = callback_log_dfe(job_id)
    on_paso = callback_paso_dfe(job_id)

    def _reinit() -> None:
        reiniciar_pasos_dfe(job_id)

    def _prog(actual: int, total: int, msg: str) -> None:
        progreso_cuit_dfe(job_id, actual, total, msg)

    def _cuit_fin(cuit, razon_social, total_archivos, error) -> None:
        agregar_resumen_cuit_dfe(
            job_id,
            cuit=cuit,
            razon_social=razon_social,
            total_archivos=total_archivos,
            error=error,
        )
        if entrega:
            entrega.escanear()

    def _worker() -> None:
        try:
            carpeta = ejecutar_dfe_lote(
                filas,
                headless=headless,
                on_log=on_log,
                on_paso=on_paso,
                on_reiniciar_pasos=_reinit,
                on_progreso=_prog,
                on_cuit_fin=_cuit_fin,
                carpeta_base=carpeta_destino,
                job_id=job_id,
                nombre_carpeta_sesion=nombre_sesion_dfe,
            )
            if entrega:
                entrega.escanear()
            marcar_ok_dfe(job_id, carpeta=str(carpeta))
        except CancelacionUsuarioError as exc:
            marcar_cancelado_dfe(job_id, str(exc))
        except ArcaProcesoError as exc:
            marcar_error_dfe(job_id, str(exc))
        except Exception as exc:
            marcar_error_dfe(job_id, _err_inesperado(exc))

    threading.Thread(target=_worker, daemon=True).start()

    if es_fetch:
        return jsonify({"job_id": job_id, "total": len(filas)})
    return render_template("dfe.html", dfe_job_id=job_id)


@app.get("/dfe-estado/<job_id>")
def dfe_estado(job_id: str):
    estado = obtener_job_dfe(job_id)
    if estado is None:
        return jsonify({"error": "job_not_found"}), 404
    return jsonify(estado)


# --------------------------------------------------------------------------- #
# Nuestra Parte
# --------------------------------------------------------------------------- #
def _filas_np_desde_peticion(lg: str):
    """Devuelve (filas, errores_planilla, mensaje_error) para Nuestra Parte."""
    f = request.files.get("np_excel")
    has_file = bool(f and getattr(f, "filename", None) and str(f.filename).strip())
    if has_file:
        nombre = Path(f.filename).name
        if not nombre.lower().endswith(".xlsx"):
            return [], [], tr(lg, "err_only_xlsx_csv")
        try:
            filas, errores = leer_planilla_np_con_errores(io.BytesIO(f.read()))
        except ArcaProcesoError as exc:
            return [], [], str(exc)
        if not filas:
            return [], errores, "; ".join(errores) or tr(lg, "np_err_sin_datos")
        return filas, errores, None

    cuits = request.form.getlist("np_cuit_login")
    claves = request.form.getlist("np_clave_fiscal")
    reprs = request.form.getlist("np_cuit_representado")
    ejercicios = request.form.getlist("np_ejercicio")

    hay_algo = any(
        (v or "").strip()
        for lista in (cuits, claves, reprs, ejercicios)
        for v in lista
    )
    if not hay_algo:
        return [], [], tr(lg, "np_err_sin_datos")

    filas, errores = parsear_entradas_manuales_np(cuits, claves, reprs, ejercicios)
    if not filas:
        return [], errores, "; ".join(errores) or tr(lg, "np_err_manual_incompleto")
    return filas, errores, None


@app.get("/nuestra-parte/plantilla")
def np_plantilla():
    from cuit_en_arca.plantillas_importacion import ruta_plantilla_np_excel

    ruta = ruta_plantilla_np_excel()
    if not ruta.is_file():
        abort(404)
    return send_file(
        ruta,
        as_attachment=True,
        download_name="Formato Nuestra Parte.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/nuestra-parte")
def nuestra_parte():
    return render_template("nuestra_parte.html")


@app.post("/np-descargar")
def np_descargar():
    lg = normalize_lang(session.get("lang"))
    es_fetch = request.headers.get("X-Requested-With") == "fetch"

    if not _mostrar_ui_cuit_arca():
        if es_fetch:
            return jsonify({"error": tr(lg, "err_arca_disabled")}), 403
        return render_template("nuestra_parte.html", error=tr(lg, "err_arca_disabled")), 403

    filas, _errores, err_msg = _filas_np_desde_peticion(lg)
    if err_msg:
        if es_fetch:
            return jsonify({"error": err_msg}), 400
        return render_template("nuestra_parte.html", error=err_msg)

    from cuit_en_arca.nuestra_parte_automation import ejecutar_nuestra_parte_lote
    from cuit_en_arca.service import _headless_desde_env

    headless = _headless_desde_env()
    carpeta_form = (request.form.get("carpeta_destino") or "").strip() or None

    job_id = uuid4().hex
    base, entrega = _fabricar_entrega(
        job_id,
        carpeta_form,
        lambda did, rel, nom: agregar_archivo_np(job_id, did, rel, nom),
    )
    if base is None:
        msg = tr(lg, "carpeta_cancelada")
        if es_fetch:
            return jsonify({"error": msg}), 400
        return render_template("nuestra_parte.html", error=msg)
    carpeta_destino = str(base)
    nombre_sesion_np = _nombre_carpeta_web_sesion("Nuestra Parte")

    def _err_inesperado(exc: Exception) -> str:
        return tr(lg, "err_arca_unexpected", exc=exc)

    from cuit_en_arca.cancelacion import reset_cancelacion

    reset_cancelacion(job_id)
    crear_job_np(job_id, len(filas))
    reiniciar_pasos_np(job_id)
    on_log = callback_log_np(job_id)
    on_paso = callback_paso_np(job_id)

    def _reinit() -> None:
        reiniciar_pasos_np(job_id)

    def _prog(actual: int, total: int, msg: str) -> None:
        progreso_cuit_np(job_id, actual, total, msg)

    def _cuit_fin(cuit, razon_social, total_archivos, error) -> None:
        agregar_resumen_cuit_np(
            job_id,
            cuit=cuit,
            razon_social=razon_social,
            total_archivos=total_archivos,
            error=error,
        )
        if entrega:
            entrega.escanear()

    def _worker() -> None:
        try:
            carpeta = ejecutar_nuestra_parte_lote(
                filas,
                headless=headless,
                on_log=on_log,
                on_paso=on_paso,
                on_reiniciar_pasos=_reinit,
                on_progreso=_prog,
                on_cuit_fin=_cuit_fin,
                carpeta_base=carpeta_destino,
                job_id=job_id,
                nombre_carpeta_sesion=nombre_sesion_np,
            )
            if entrega:
                entrega.escanear()
            marcar_ok_np(job_id, carpeta=str(carpeta))
        except CancelacionUsuarioError as exc:
            marcar_cancelado_np(job_id, str(exc))
        except ArcaProcesoError as exc:
            marcar_error_np(job_id, str(exc))
        except Exception as exc:
            marcar_error_np(job_id, _err_inesperado(exc))

    threading.Thread(target=_worker, daemon=True).start()

    if es_fetch:
        return jsonify({"job_id": job_id, "total": len(filas)})
    return render_template("nuestra_parte.html", np_job_id=job_id)


@app.get("/np-estado/<job_id>")
def np_estado(job_id: str):
    estado = obtener_job_np(job_id)
    if estado is None:
        return jsonify({"error": "job_not_found"}), 404
    return jsonify(estado)


# --------------------------------------------------------------------------- #
# Análisis Programado
# --------------------------------------------------------------------------- #
def _filas_ap_desde_peticion(lg: str):
    from cuit_en_arca.planilla_analisis_programado import leer_planilla_analisis_programado

    f = request.files.get("ap_excel")
    has_file = bool(f and getattr(f, "filename", None) and str(f.filename).strip())
    if has_file:
        nombre = Path(f.filename).name
        if not nombre.lower().endswith(".xlsx"):
            return [], [], tr(lg, "err_only_xlsx_csv")
        try:
            filas, errores = leer_planilla_analisis_programado(io.BytesIO(f.read()))
        except ArcaProcesoError as exc:
            return [], [], str(exc)
        if not filas:
            return [], errores, "; ".join(errores) or tr(lg, "ap_err_sin_datos")
        return filas, errores, None

    return _filas_ap_desde_manual(lg)


def _filas_ap_desde_manual(lg: str):
    from cuit_en_arca.planilla_analisis_programado import parsear_entradas_manuales_ap

    cuits = request.form.getlist("ap_cuit")
    claves = request.form.getlist("ap_clave")
    reprs = request.form.getlist("ap_repr")
    fechas_mc = request.form.getlist("ap_fechas_mc")
    dfe_desde = request.form.getlist("ap_dfe_desde")
    dfe_hasta = request.form.getlist("ap_dfe_hasta")
    ejercicios = request.form.getlist("ap_ejercicio")

    hay = any(
        (v or "").strip()
        for lst in (cuits, claves, reprs, fechas_mc, dfe_desde, dfe_hasta, ejercicios)
        for v in lst
    )
    if not hay:
        return [], [], tr(lg, "ap_err_sin_datos")

    filas, errores = parsear_entradas_manuales_ap(
        cuits, claves, reprs, fechas_mc, dfe_desde, dfe_hasta, ejercicios
    )
    if not filas:
        return [], errores, "; ".join(errores) or tr(lg, "ap_err_sin_datos")
    return filas, errores, None


def _filas_ap_a_dict(filas) -> list[dict]:
    from dataclasses import asdict

    return [asdict(f) for f in filas]


@app.get("/analisis-programado")
def analisis_programado():
    from cuit_en_arca.analisis_programado import cargar_config

    cfg = cargar_config()
    return render_template(
        "analisis_programado.html",
        config=cfg.a_dict_publico(),
    )


@app.get("/analisis-programado/plantilla")
def analisis_programado_plantilla():
    from cuit_en_arca.analisis_programado import ruta_plantilla_excel

    ruta = ruta_plantilla_excel()
    if not ruta.is_file():
        abort(404)
    return send_file(
        ruta,
        as_attachment=True,
        download_name="Formato Analisis Programado.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/analisis-programado/estado")
def analisis_programado_estado():
    from cuit_en_arca.analisis_programado import cargar_config

    return jsonify(cargar_config().a_dict_publico())


@app.get("/analisis-programado/ejecucion")
def analisis_programado_ejecucion():
    from cuit_en_arca.progreso_analisis_programado import obtener_ejecucion_ap

    return jsonify(obtener_ejecucion_ap())


@app.post("/analisis-programado/guardar")
def analisis_programado_guardar():
    from cuit_en_arca.analisis_programado import (
        ConfigAnalisisProgramado,
        cargar_config,
        guardar_config,
    )

    lg = normalize_lang(session.get("lang"))
    es_fetch = request.headers.get("X-Requested-With") == "fetch"

    sistemas = [
        s
        for s in request.form.getlist("ap_sistemas")
        if s in ("mis_comprobantes", "dfe", "nuestra_parte")
    ]
    if not sistemas:
        msg = tr(lg, "ap_err_sin_sistema")
        if es_fetch:
            return jsonify({"error": msg}), 400
        return render_template(
            "analisis_programado.html",
            error=msg,
            config=cargar_config().a_dict_publico(),
        )

    carpeta = (request.form.get("carpeta_destino") or "").strip()
    if _es_app_escritorio() and not carpeta:
        msg = tr(lg, "ap_err_sin_carpeta")
        if es_fetch:
            return jsonify({"error": msg}), 400
        return render_template(
            "analisis_programado.html",
            error=msg,
            config=cargar_config().a_dict_publico(),
        )
    if not _es_app_escritorio():
        from cuit_en_arca.entrega_web import carpeta_ap_servidor

        carpeta = str(carpeta_ap_servidor())

    filas, _errores, err_msg = _filas_ap_desde_peticion(lg)
    if err_msg:
        if es_fetch:
            return jsonify({"error": err_msg}), 400
        return render_template(
            "analisis_programado.html",
            error=err_msg,
            config=cargar_config().a_dict_publico(),
        )

    try:
        dia = int(request.form.get("ap_dia_semana", "0"))
        hora = int(request.form.get("ap_hora", "9"))
        minuto = int(request.form.get("ap_minuto", "0"))
    except ValueError:
        dia, hora, minuto = 0, 9, 0

    cfg = ConfigAnalisisProgramado(
        activo=True,
        dia_semana=max(0, min(6, dia)),
        hora=max(0, min(23, hora)),
        minuto=max(0, min(59, minuto)),
        sistemas=sistemas,
        carpeta_destino=carpeta,
        filas=_filas_ap_a_dict(filas),
        ultima_ejecucion=None,
        ultimo_resultado=None,
    )
    try:
        guardar_config(cfg)
    except OSError as exc:
        msg = tr(lg, "ap_err_guardar") + f" ({exc})"
        if es_fetch:
            return jsonify({"error": msg}), 500
        return render_template(
            "analisis_programado.html",
            error=msg,
            config=cargar_config().a_dict_publico(),
        )

    if es_fetch:
        return jsonify({"ok": True, "mensaje": tr(lg, "ap_ok_guardado"), "config": cfg.a_dict_publico()})
    return render_template(
        "analisis_programado.html",
        ok=tr(lg, "ap_ok_guardado"),
        config=cfg.a_dict_publico(),
    )


@app.post("/api/cancelar-descarga")
def cancelar_descarga():
    from cuit_en_arca.browser_desktop import cerrar_navegador_desktop
    from cuit_en_arca.cancelacion import solicitar_cancelacion, solicitar_cancelacion_ap

    payload = request.get_json(silent=True) or {}
    tipo = (payload.get("tipo") or request.form.get("tipo") or "").strip()
    job_id = (payload.get("job_id") or request.form.get("job_id") or "").strip()

    if tipo == "ap":
        solicitar_cancelacion_ap()
    elif job_id:
        solicitar_cancelacion(job_id)
    else:
        lg = normalize_lang(session.get("lang"))
        return jsonify({"error": tr(lg, "err_arca_unexpected", exc="sin job")}), 400

    try:
        cerrar_navegador_desktop()
    except Exception:
        pass
    return jsonify({"ok": True})


@app.get("/api/auth-users")
def api_auth_users():
    """Listado de usuarios para sync de portables (Bearer AUTH_USERS_REMOTE_TOKEN)."""
    if not verificar_token_remoto(request.headers.get("Authorization")):
        return jsonify({"error": "unauthorized"}), 401
    return jsonify(export_users_payload())


@app.post("/analisis-programado/limpiar")
def analisis_programado_limpiar():
    from cuit_en_arca.analisis_programado import limpiar_configuracion_completa
    from cuit_en_arca.progreso_analisis_programado import resetear_ejecucion_ap

    lg = normalize_lang(session.get("lang"))
    cfg = limpiar_configuracion_completa()
    resetear_ejecucion_ap()
    return jsonify({
        "ok": True,
        "mensaje": tr(lg, "ap_ok_limpiado"),
        "config": cfg.a_dict_publico(),
    })


@app.get("/elegir-carpeta")
def elegir_carpeta():
    """Abre un diálogo nativo del sistema para elegir la carpeta de descarga.

    Solo tiene sentido en el escritorio (servidor = PC del usuario), por eso se
    restringe a peticiones locales.
    """
    ra = (request.remote_addr or "").replace("::ffff:", "")
    if ra not in ("127.0.0.1", "::1"):
        return jsonify({"error": "solo_local"}), 403

    from urllib.parse import unquote

    titulo = unquote(request.args.get("titulo") or "Elegir carpeta de descarga").strip()
    from cuit_en_arca.elegir_carpeta import elegir_carpeta_dialogo

    ruta = elegir_carpeta_dialogo(titulo)
    if not ruta:
        return jsonify({"cancelado": True})
    return jsonify({"carpeta": ruta})


if __name__ == "__main__":
    import threading
    import webbrowser

    try:
        from cuit_en_arca.analisis_programado import iniciar_scheduler

        iniciar_scheduler()
    except Exception:
        pass

    puerto = int(os.environ.get("PORT", 5000))
    url = f"http://127.0.0.1:{puerto}/"
    print(f"\n  Servidor: {url}\n  (Abrí esa dirección en el navegador si no se abre sola.)\n", flush=True)
    if os.environ.get("OPEN_BROWSER", "1").strip().lower() in ("1", "true", "yes", "on"):
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="0.0.0.0", port=puerto, debug=False)
