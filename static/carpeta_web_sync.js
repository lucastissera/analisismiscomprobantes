/** Guarda en la carpeta elegida (web) los archivos que el servidor va generando. */
(function (global) {
  var _descargados = Object.create(null);

  function clave(jobKey, fileId) {
    return String(jobKey || "job") + ":" + String(fileId || "");
  }

  async function escribirEnRuta(rootHandle, rutaRelativa, blob) {
    var partes = String(rutaRelativa || "")
      .replace(/\\/g, "/")
      .split("/")
      .filter(Boolean);
    if (!partes.length) return;
    var dir = rootHandle;
    var i;
    for (i = 0; i < partes.length - 1; i++) {
      dir = await dir.getDirectoryHandle(partes[i], { create: true });
    }
    var fh = await dir.getFileHandle(partes[partes.length - 1], { create: true });
    var writable = await fh.createWritable();
    await writable.write(blob);
    await writable.close();
  }

  async function syncArchivos(archivos, jobKey) {
    if (global.McElegirCarpeta && global.McElegirCarpeta.esModoEscritorio()) {
      return;
    }
    var handle = global.McElegirCarpeta && global.McElegirCarpeta.obtenerHandleWeb
      ? global.McElegirCarpeta.obtenerHandleWeb()
      : null;
    if (!handle || !archivos || !archivos.length) return;

    for (var i = 0; i < archivos.length; i++) {
      var a = archivos[i];
      if (!a || !a.id) continue;
      var k = clave(jobKey, a.id);
      if (_descargados[k]) continue;
      var resp = await fetch("/descargar/" + encodeURIComponent(a.id), {
        credentials: "same-origin",
      });
      if (!resp.ok) continue;
      var blob = await resp.blob();
      await escribirEnRuta(handle, a.ruta || a.nombre, blob);
      _descargados[k] = true;
    }
  }

  function syncDesdeEstado(st, jobKey) {
    if (!st || !st.archivos || !st.archivos.length) return Promise.resolve();
    return syncArchivos(st.archivos, jobKey || st.job_id || "ap").catch(function () {});
  }

  function reset(jobKey) {
    if (!jobKey) {
      _descargados = Object.create(null);
      return;
    }
    var pref = String(jobKey) + ":";
    Object.keys(_descargados).forEach(function (k) {
      if (k.indexOf(pref) === 0) delete _descargados[k];
    });
  }

  global.McCarpetaWebSync = {
    syncDesdeEstado: syncDesdeEstado,
    syncArchivos: syncArchivos,
    reset: reset,
  };
})(window);
