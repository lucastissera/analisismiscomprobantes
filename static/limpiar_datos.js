window.McLimpiarDatos = function (opts) {
  if (!opts) return;
  var confirmMsg = opts.confirm || "¿Borrar los datos cargados?";
  if (!confirm(confirmMsg)) return;

  var form = opts.formId ? document.getElementById(opts.formId) : null;

  if (typeof opts.onLimpiar === "function") {
    opts.onLimpiar();
  }

  if (form) {
    form.reset();
  }

  if (opts.fileIds && opts.fileIds.length) {
    opts.fileIds.forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.value = "";
    });
  }

  if (opts.fileNames && opts.fileNames.length) {
    opts.fileNames.forEach(function (name) {
      var el = form ? form.querySelector('[name="' + name + '"]') : null;
      if (!el && form) {
        el = form.querySelector('input[name="' + name + '"]');
      }
      if (el) el.value = "";
    });
  }

  if (opts.manualContainerId) {
    var cont = document.getElementById(opts.manualContainerId);
    if (cont) {
      var filas = cont.querySelectorAll(opts.manualRowSelector || ".fila-manual, .arca-fila-manual");
      for (var i = filas.length - 1; i > 0; i--) {
        filas[i].remove();
      }
      if (filas[0]) {
        filas[0].querySelectorAll("input").forEach(function (inp) {
          inp.value = "";
          if (inp.dataset) {
            inp.dataset.touched = "";
          }
        });
        if (window.resetClaveFields) window.resetClaveFields(filas[0]);
      }
    }
  }

  if (opts.carpetaInputId) {
    var carp = document.getElementById(opts.carpetaInputId);
    if (carp) carp.value = "";
  }
  if (opts.carpetaLabelId) {
    var lab = document.getElementById(opts.carpetaLabelId);
    if (lab) lab.textContent = "—";
  }

  if (window.McElegirCarpeta && window.McElegirCarpeta.esModoEscritorio && !window.McElegirCarpeta.esModoEscritorio()) {
    if (window.McElegirCarpeta.limpiarCarpetaWeb) window.McElegirCarpeta.limpiarCarpetaWeb();
  }
  if (window.McCarpetaWebSync && window.McCarpetaWebSync.reset) {
    window.McCarpetaWebSync.reset();
  }

  if (typeof opts.onListo === "function") {
    opts.onListo();
  }
};
