/** Carpeta de destino: diálogo nativo (.exe) o File System Access API (web). */
(function (global) {
  var PREFIJOS_SISTEMA = {
    mis_comprobantes: "Mis Comprobantes",
    dfe: "DFE",
    nuestra_parte: "Nuestra Parte",
    analisis_programado: "Análisis Programado",
  };

  var _dirHandle = null;
  var _dirNombre = "";
  var _parentNombre = "";
  var _subcarpetaSesion = null;
  var _sistemaActual = null;

  function esModoEscritorio() {
    return global.MC_MODO_ESCRITORIO === true;
  }

  function soporteCarpetaWeb() {
    return typeof global.showDirectoryPicker === "function";
  }

  function requiereCarpeta(ruta) {
    if (esModoEscritorio()) {
      return !String(ruta || "").trim();
    }
    return !_dirHandle;
  }

  function obtenerHandleWeb() {
    return _dirHandle;
  }

  function obtenerNombreWeb() {
    return _dirNombre;
  }

  function obtenerSubcarpetaSesion() {
    return _subcarpetaSesion;
  }

  function stampCarpetaEjecucion(fecha) {
    var d = fecha || new Date();
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, "0");
    var day = String(d.getDate()).padStart(2, "0");
    var h = String(d.getHours()).padStart(2, "0");
    var min = String(d.getMinutes()).padStart(2, "0");
    return y + "-" + m + "-" + day + " " + h + "-" + min;
  }

  function nombreSubcarpetaSistema(sistema) {
    var pref = PREFIJOS_SISTEMA[sistema];
    if (!pref) return null;
    return pref + " " + stampCarpetaEjecucion();
  }

  function esErrorCarpetaRestringida(err) {
    if (!err) return false;
    var nombre = String(err.name || "");
    var msg = String(err.message || "").toLowerCase();
    if (nombre === "SecurityError" || nombre === "NotAllowedError") {
      if (
        msg.indexOf("system") >= 0 ||
        msg.indexOf("sistema") >= 0 ||
        msg.indexOf("restricted") >= 0 ||
        msg.indexOf("restring") >= 0
      ) {
        return true;
      }
    }
    return (
      msg.indexOf("system file") >= 0 ||
      msg.indexOf("archivos del sistema") >= 0 ||
      msg.indexOf("contains system") >= 0 ||
      msg.indexOf("contiene archivos") >= 0
    );
  }

  function mensajeError(err, textos) {
    textos = textos || {};
    var cancelada =
      textos.cancelada ||
      textos.carpeta_cancelada ||
      "No se eligió ninguna carpeta.";
    var noSoportado =
      textos.noSoportado ||
      textos.err_carpeta_web_api ||
      textos.errCarpetaWebApi ||
      cancelada;
    var restringida =
      textos.restringida ||
      textos.err_carpeta_restringida ||
      textos.errCarpetaRestringida ||
      cancelada;

    if (!err) return cancelada;
    if (err.message === "picker_web_no_soportado") return noSoportado;
    if (err.message === "picker_carpeta_restringida" || esErrorCarpetaRestringida(err)) {
      return restringida;
    }
    return cancelada;
  }

  function actualizarCampoSesion(sesionInputId) {
    if (!sesionInputId) return;
    var el = document.getElementById(sesionInputId);
    if (el) el.value = _subcarpetaSesion || "";
  }

  function prepararCarpetaWeb(parentHandle, sistema) {
    _sistemaActual = sistema || null;
    _parentNombre = parentHandle.name || "Carpeta";

    if (sistema === "analisis_programado") {
      _dirHandle = parentHandle;
      _dirNombre = _parentNombre;
      _subcarpetaSesion = null;
      return Promise.resolve(_dirNombre);
    }

    var subNombre = nombreSubcarpetaSistema(sistema || "mis_comprobantes");
    if (!subNombre) {
      _dirHandle = parentHandle;
      _dirNombre = _parentNombre;
      _subcarpetaSesion = null;
      return Promise.resolve(_dirNombre);
    }

    return parentHandle.getDirectoryHandle(subNombre, { create: true }).then(function (subHandle) {
      _dirHandle = subHandle;
      _subcarpetaSesion = subNombre;
      _dirNombre = _parentNombre + " / " + subNombre;
      return _dirNombre;
    });
  }

  function elegirCarpetaNativa(titulo) {
    var q = encodeURIComponent(titulo || "Elegir carpeta de descarga");
    return fetch("/elegir-carpeta?titulo=" + q, { credentials: "same-origin" })
      .then(function (r) {
        if (!r.ok) throw new Error("picker_http");
        return r.json();
      })
      .then(function (data) {
        if (data && data.error) throw new Error(data.error);
        if (!data || !data.carpeta) return null;
        return data.carpeta;
      });
  }

  function elegirCarpetaWeb(titulo, sistema) {
    if (!soporteCarpetaWeb()) {
      return Promise.reject(new Error("picker_web_no_soportado"));
    }
    return global
      .showDirectoryPicker({ mode: "readwrite" })
      .then(function (handle) {
        return prepararCarpetaWeb(handle, sistema);
      })
      .catch(function (err) {
        if (err && err.name === "AbortError") return null;
        if (esErrorCarpetaRestringida(err)) {
          throw new Error("picker_carpeta_restringida");
        }
        throw err;
      });
  }

  function elegirCarpeta(titulo, sistema) {
    if (esModoEscritorio()) {
      return elegirCarpetaNativa(titulo);
    }
    return elegirCarpetaWeb(titulo, sistema);
  }

  function configurarUiWeb(opts) {
    if (esModoEscritorio()) return;
    if (!soporteCarpetaWeb() && opts.onError) {
      opts.onError(new Error("picker_web_no_soportado"));
    }
  }

  function enlazar(opts) {
    var input = document.getElementById(opts.inputId);
    var label = opts.labelId ? document.getElementById(opts.labelId) : null;
    var btn = opts.btnId ? document.getElementById(opts.btnId) : null;
    var titulo = opts.titulo || "Elegir carpeta de descarga";
    var textos = opts.textos || null;
    var sistema = opts.sistema || "mis_comprobantes";
    var sesionInputId = opts.sesionInputId || null;

    configurarUiWeb(opts);

    function aplicar(ruta) {
      if (input) input.value = ruta;
      if (label) label.textContent = ruta;
      actualizarCampoSesion(sesionInputId);
    }

    function picker() {
      if (btn) btn.disabled = true;
      return elegirCarpeta(titulo, sistema)
        .then(function (ruta) {
          if (btn) btn.disabled = false;
          if (!ruta) {
            if (opts.onCancel) opts.onCancel();
            return null;
          }
          aplicar(ruta);
          if (opts.onOk) opts.onOk(ruta);
          return ruta;
        })
        .catch(function (err) {
          if (btn) btn.disabled = false;
          if (opts.onError) opts.onError(err, mensajeError(err, textos));
          throw err;
        });
    }

    if (btn) {
      btn.addEventListener("click", function () {
        picker().catch(function () {});
      });
    }

    return {
      elegir: picker,
      obtener: function () {
        if (esModoEscritorio()) {
          return ((input && input.value) || "").trim();
        }
        return _dirNombre || "";
      },
      aplicar: aplicar,
    };
  }

  function resolverCarpeta(opts) {
    var api = opts.api;
    if (esModoEscritorio()) {
      var actual = api.obtener();
      if (actual) return Promise.resolve(actual);
      return api.elegir();
    }
    if (_dirHandle) return Promise.resolve(_dirNombre);
    return api.elegir();
  }

  function limpiarCarpetaWeb() {
    _dirHandle = null;
    _dirNombre = "";
    _parentNombre = "";
    _subcarpetaSesion = null;
    _sistemaActual = null;
    document.querySelectorAll('input[name="web_carpeta_sesion"]').forEach(function (el) {
      el.value = "";
    });
  }

  global.McElegirCarpeta = {
    elegir: elegirCarpeta,
    enlazar: enlazar,
    resolver: resolverCarpeta,
    requiereCarpeta: requiereCarpeta,
    esModoEscritorio: esModoEscritorio,
    soporteCarpetaWeb: soporteCarpetaWeb,
    obtenerHandleWeb: obtenerHandleWeb,
    obtenerSubcarpetaSesion: obtenerSubcarpetaSesion,
    limpiarCarpetaWeb: limpiarCarpetaWeb,
    mensajeError: mensajeError,
    esErrorCarpetaRestringida: esErrorCarpetaRestringida,
    nombreSubcarpetaSistema: nombreSubcarpetaSistema,
  };
})(window);
