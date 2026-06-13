/** Carpeta de destino: diálogo nativo (.exe) o File System Access API (web). */
(function (global) {
  var _dirHandle = null;
  var _dirNombre = "";

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

  function confirmarPermisoEscritura(handle) {
    if (!handle || typeof handle.requestPermission !== "function") {
      return Promise.resolve("granted");
    }
    return handle.requestPermission({ mode: "readwrite" }).then(function (perm) {
      if (perm === "granted") return perm;
      throw new Error("picker_carpeta_restringida");
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

  function elegirCarpetaWeb(titulo) {
    if (!soporteCarpetaWeb()) {
      return Promise.reject(new Error("picker_web_no_soportado"));
    }
    return global
      .showDirectoryPicker({ mode: "readwrite" })
      .then(function (handle) {
        return confirmarPermisoEscritura(handle).then(function () {
          _dirHandle = handle;
          _dirNombre = handle.name || "Carpeta";
          return _dirNombre;
        });
      })
      .catch(function (err) {
        if (err && err.name === "AbortError") return null;
        if (esErrorCarpetaRestringida(err)) {
          throw new Error("picker_carpeta_restringida");
        }
        throw err;
      });
  }

  function elegirCarpeta(titulo) {
    if (esModoEscritorio()) {
      return elegirCarpetaNativa(titulo);
    }
    return elegirCarpetaWeb(titulo);
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

    configurarUiWeb(opts);

    function aplicar(ruta) {
      if (input) input.value = ruta;
      if (label) label.textContent = ruta;
    }

    function picker() {
      if (btn) btn.disabled = true;
      return elegirCarpeta(titulo)
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
  }

  global.McElegirCarpeta = {
    elegir: elegirCarpeta,
    enlazar: enlazar,
    resolver: resolverCarpeta,
    requiereCarpeta: requiereCarpeta,
    esModoEscritorio: esModoEscritorio,
    soporteCarpetaWeb: soporteCarpetaWeb,
    obtenerHandleWeb: obtenerHandleWeb,
    limpiarCarpetaWeb: limpiarCarpetaWeb,
    mensajeError: mensajeError,
    esErrorCarpetaRestringida: esErrorCarpetaRestringida,
  };
})(window);
