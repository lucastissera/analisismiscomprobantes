import io
import os
from pathlib import Path
from uuid import uuid4

from flask import Flask, render_template, request, send_file

from sumar_imp_total import (
    COLUMNAS_A_AJUSTAR,
    COLUMNAS_DETALLE_SIN_RESUMEN,
    COLUMNAS_TOTAL_RESUMEN,
    NOMBRES_MESES,
    procesar_archivo,
    total_resumen_pantalla,
    totales_resumen_por_mes,
)


app = Flask(__name__)
DESCARGAS: dict[str, tuple[bytes, str]] = {}


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/descargar/<download_id>")
def descargar(download_id: str):
    item = DESCARGAS.get(download_id)
    if not item:
        return render_template("index.html", error="El archivo a descargar ya no está disponible.")

    contenido, nombre_salida = item
    return send_file(
        io.BytesIO(contenido),
        as_attachment=True,
        download_name=nombre_salida,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/procesar")
def procesar():
    archivo = request.files.get("excel")
    if not archivo or archivo.filename == "":
        return render_template("index.html", error="Debes seleccionar un archivo .xlsx o .csv")

    nombre = Path(archivo.filename).name
    if not (nombre.lower().endswith(".xlsx") or nombre.lower().endswith(".csv")):
        return render_template("index.html", error="Solo se permiten archivos .xlsx o .csv")

    try:
        datos = archivo.read()
        buffer = io.BytesIO(datos)
        df_ajustado, totales, totales_por_mes = procesar_archivo(
            buffer, 0, nombre_archivo=nombre
        )
    except ValueError as exc:
        return render_template("index.html", error=str(exc))
    except Exception as exc:  # fallback para errores no esperados
        return render_template(
            "index.html", error=f"Ocurrió un error al procesar el archivo: {exc}"
        )

    salida = io.BytesIO()
    df_ajustado.to_excel(salida, index=False)
    contenido = salida.getvalue()

    nombre_salida = f"{Path(nombre).stem}_ajustado.xlsx"
    download_id = uuid4().hex
    DESCARGAS[download_id] = (contenido, nombre_salida)

    resumen_total_mes = totales_resumen_por_mes(totales_por_mes)
    totales_resumen = {c: totales[c] for c in COLUMNAS_TOTAL_RESUMEN}
    totales_detalle = {c: totales[c] for c in COLUMNAS_DETALLE_SIN_RESUMEN}
    meses_idx = list(range(1, 13))

    return render_template(
        "index.html",
        mostrar_resultado=True,
        totales_resumen=totales_resumen,
        totales_detalle=totales_detalle,
        columnas_orden=COLUMNAS_A_AJUSTAR,
        suma_total=round(total_resumen_pantalla(totales), 2),
        totales_por_mes=totales_por_mes,
        nombres_meses=NOMBRES_MESES,
        meses_idx=meses_idx,
        resumen_total_mes=resumen_total_mes,
        download_id=download_id,
        nombre_salida=nombre_salida,
    )


if __name__ == "__main__":
    puerto = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=puerto, debug=False)
