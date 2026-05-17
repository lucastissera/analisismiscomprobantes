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
except ImportError:
    pass

# En local (sin RENDER) habilitar Playwright/AFIP si no definiste la variable.
# En Render: definí CUIT_EN_ARCA_PLAYWRIGHT=1 en Environment si querés la descarga automática.
if os.environ.get("RENDER", "").strip().lower() not in ("true", "1", "yes"):
    os.environ.setdefault("CUIT_EN_ARCA_PLAYWRIGHT", "1")

from flask import (
    Flask,
    abort,
    flash,
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
    app = Flask(__name__, root_path=str(_bundle), template_folder=str(_tpl))
else:
    app = Flask(__name__)

from auth import verify_credentials, whatsapp_new_user_url
from cuit_en_arca import ArcaProcesoError, ejecutar_flujo_cuit_en_arca
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

# Inactividad: sin peticiones al servidor durante este tiempo → cerrar sesión.
# Cada petición (refresco, nueva pestaña con la misma app, navegación) renueva el plazo.
_SESSION_IDLE_SEC = 30 * 60


def _safe_internal_path(target: str | None) -> str:
    """Solo rutas relativas del mismo sitio (evita redirecciones abiertas)."""
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

    if request.endpoint in ("login", "set_lang", "desktop_quit", None):
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
        if verify_credentials(user, pwd):
            session["user"] = user
            session["last_activity"] = time.time()
            session.permanent = True
            session.modified = True
            return redirect(_safe_internal_path(next_val or request.args.get("next")))
        return render_template(
            "login.html",
            login_error=True,
            next=next_val,
            whatsapp_url=whatsapp_new_user_url(),
        )
    next_val = (request.args.get("next") or "").strip()
    return render_template(
        "login.html",
        next=next_val,
        whatsapp_url=whatsapp_new_user_url(),
    )


@app.get("/logout")
def logout():
    session.pop("user", None)
    session.pop("last_activity", None)
    return redirect(url_for("login"))


@app.route("/desktop-quit", methods=["GET", "POST"])
def desktop_quit():
    """Solo .exe local: cierra el proceso (sin consola no hay otra forma obvia de salir)."""
    if not getattr(sys, "frozen", False):
        abort(404)
    ra = (request.remote_addr or "").replace("::ffff:", "")
    if ra not in ("127.0.0.1", "::1"):
        abort(403)
    lg = normalize_lang(session.get("lang"))

    def _salir() -> None:
        time.sleep(0.2)
        os._exit(0)

    threading.Thread(target=_salir, daemon=True).start()
    return Response(
        tr(lg, "desktop_quit_cerrando") + "\n",
        mimetype="text/plain; charset=utf-8",
    )


MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _mimetype_por_nombre(nombre: str) -> str:
    nl = nombre.lower()
    if nl.endswith(".csv"):
        return "text/csv; charset=utf-8"
    return MIME_XLSX


@app.get("/")
def index():
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


@app.post("/cuit-en-arca")
def cuit_en_arca():
    lg = normalize_lang(session.get("lang"))
    if not _mostrar_ui_cuit_arca():
        return (
            render_template(
                "index.html",
                error=tr(lg, "err_arca_disabled"),
            ),
            403,
        )

    cred_file = request.files.get("credenciales")
    fecha_desde = (request.form.get("fecha_desde") or "").strip()
    fecha_hasta = (request.form.get("fecha_hasta") or "").strip()

    if not cred_file or cred_file.filename == "":
        return render_template(
            "index.html",
            error=tr(lg, "err_arca_cred_missing"),
        )
    if not Path(cred_file.filename).name.lower().endswith(".xlsx"):
        return render_template(
            "index.html",
            error=tr(lg, "err_arca_xlsx"),
        )
    try:
        buf = io.BytesIO(cred_file.read())
        data, nombre_sug = ejecutar_flujo_cuit_en_arca(
            buf,
            fecha_desde or None,
            fecha_hasta or None,
        )
    except ArcaProcesoError as exc:
        return render_template("index.html", error=str(exc))
    except Exception as exc:
        return render_template(
            "index.html",
            error=tr(lg, "err_arca_unexpected", exc=exc),
        )

    nombre_out = Path(nombre_sug).name if nombre_sug else "mis_comprobantes_descarga.xlsx"
    if not nombre_out.lower().endswith((".xlsx", ".csv")):
        nombre_out = f"{nombre_out}.xlsx"

    did = uuid4().hex
    DESCARGAS[did] = (data, nombre_out, _mimetype_por_nombre(nombre_out))

    return render_template(
        "index.html",
        cuit_arca_ok=True,
        cuit_arca_download_id=did,
        cuit_arca_nombre=nombre_out,
    )


if __name__ == "__main__":
    import threading
    import webbrowser

    puerto = int(os.environ.get("PORT", 5000))
    url = f"http://127.0.0.1:{puerto}/"
    print(f"\n  Servidor: {url}\n  (Abrí esa dirección en el navegador si no se abre sola.)\n", flush=True)
    if os.environ.get("OPEN_BROWSER", "1").strip().lower() in ("1", "true", "yes", "on"):
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="0.0.0.0", port=puerto, debug=False)
