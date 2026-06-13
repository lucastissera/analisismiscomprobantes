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
    return global.showDirectoryPicker({ mode: "readwrite", id: "aic-carpeta" })
      .then(function (handle) {
        _dirHandle = handle;
        _dirNombre = handle.name || "Carpeta";
        return _dirNombre;
      })
      .catch(function (err) {
        if (err && err.name === "AbortError") return null;
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
          if (opts.onError) opts.onError(err);
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
  };
})(window);
