"""
Lee un archivo Excel (.xlsx) o CSV, ajusta signos según Tipo y suma columnas indicadas.
En .xlsx se prueba encabezado en fila 1 o fila 2 y se usa el que tenga todas las columnas requeridas.
En .csv los encabezados suelen estar en fila 1.
Las filas con Tipo = nota de crédito (por código numérico) se consideran en negativo.

Uso:
  python sumar_imp_total.py <ruta_al_archivo.xlsx> [hoja] [archivo_salida.xlsx]
"""

import io
import re
import sys
import csv
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Columnas a evaluar: se les aplica signo por Tipo y multiplicación por Tipo Cambio
COLUMNAS_A_AJUSTAR = [
    "Neto Grav. IVA 0%",
    "IVA 2,5%",
    "Neto Grav. IVA 2,5%",
    "IVA 5%",
    "Neto Grav. IVA 5%",
    "IVA 10,5%",
    "Neto Grav. IVA 10,5%",
    "IVA 21%",
    "Neto Grav. IVA 21%",
    "IVA 27%",
    "Neto Grav. IVA 27%",
    "Neto Gravado Total",
    "Neto No Gravado",
    "Op. Exentas",
    "Otros Tributos",
    "Total IVA",
    "Imp. Total",
]

# Totales que se suman en la línea "Total" del resumen en pantalla (no IVA desglosado ni Imp. Total)
COLUMNAS_TOTAL_RESUMEN = [
    "Neto Grav. IVA 0%",
    "Neto Gravado Total",
    "Neto No Gravado",
    "Op. Exentas",
    "Otros Tributos",
    "Total IVA",
]

# Columnas de detalle (IVA por alícuota e Imp. Total): no entran en la fila "Total (resumen)"
COLUMNAS_DETALLE_SIN_RESUMEN = [
    c for c in COLUMNAS_A_AJUSTAR if c not in COLUMNAS_TOTAL_RESUMEN
]

# Códigos numéricos de la columna "Tipo" que se consideran nota de crédito (suma en negativo)
# Se matchea por número (sin ceros a la izquierda): "003", "3" y "03" son el mismo código
CODIGOS_NOTA_CREDITO = {
    3, 8, 13, 21, 38, 43, 44, 48, 53,
    110, 112, 113, 114, 203, 206, 208, 211, 213,
}

# Tipos B/C (y afines): Imp. Total → Neto Grav. IVA 0% (no Neto Gravado Total). Columna Tipo.
CODIGOS_IMP_TOTAL_EN_NETO_IVA_0 = frozenset(
    (
        6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 18, 19, 20, 21, 25, 26, 28, 29,
        40, 41, 42, 43, 44, 46, 47, 61, 64, 82, 83, 90, 91, 109, 110, 111,
        113, 114, 116, 117, 206, 207, 208, 211, 212, 213,
    )
)

# Códigos AFIP típicos de comprobantes con letra B (excl. redirección Imp.→Neto IVA 0% si emitidos).
CODIGOS_LETRA_B_AFIP = frozenset(
    {
        6,
        7,
        8,
        9,
        10,
    }
)

NOMBRES_MESES = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]

def limpiar_nombre_columna_bruto(nombre: str) -> str:
    """Corrige nombres de columna típicos (mojibake UTF-8 como latin-1, etc.)."""
    s = str(nombre).strip()
    s = s.replace("EmisiÃ³n", "Emisión")
    s = s.replace("Ã³", "ó").replace("Ã­", "í").replace("Ã¡", "á").replace("Ã©", "é")
    s = s.replace("Ãº", "ú").replace("Ã±", "ñ")
    return s


# Alias de columnas para soportar variaciones entre .xlsx y .csv (ARCA)
ALIAS_COLUMNAS = {
    "Fecha Emisión": [
        "Fecha Emisión",
        "Fecha de Emisión",
        "Fecha de emisión",
        "Fecha",
    ],
    "Tipo": ["Tipo", "Tipo de Comprobante"],
    "Tipo Cambio": ["Tipo Cambio"],
    "Neto Grav. IVA 0%": ["Neto Grav. IVA 0%", "Imp. Neto Gravado IVA 0%"],
    "IVA 2,5%": ["IVA 2,5%"],
    "Neto Grav. IVA 2,5%": ["Neto Grav. IVA 2,5%", "Imp. Neto Gravado IVA 2,5%"],
    "IVA 5%": ["IVA 5%"],
    "Neto Grav. IVA 5%": ["Neto Grav. IVA 5%", "Imp. Neto Gravado IVA 5%"],
    "IVA 10,5%": ["IVA 10,5%"],
    "Neto Grav. IVA 10,5%": ["Neto Grav. IVA 10,5%", "Imp. Neto Gravado IVA 10,5%"],
    "IVA 21%": ["IVA 21%"],
    "Neto Grav. IVA 21%": ["Neto Grav. IVA 21%", "Imp. Neto Gravado IVA 21%"],
    "IVA 27%": ["IVA 27%"],
    "Neto Grav. IVA 27%": ["Neto Grav. IVA 27%", "Imp. Neto Gravado IVA 27%"],
    "Neto Gravado Total": ["Neto Gravado Total", "Imp. Neto Gravado Total"],
    "Neto No Gravado": ["Neto No Gravado", "Imp. Neto No Gravado"],
    "Op. Exentas": ["Op. Exentas", "Imp. Op. Exentas"],
    "Otros Tributos": ["Otros Tributos"],
    "Total IVA": ["Total IVA"],
    "Imp. Total": ["Imp. Total"],
}


def parsear_numero_importe(val) -> float:
    """
    Convierte valores típicos de exportación ARCA/CSV argentino a float.
    Soporta coma decimal (10156,44), notación científica con coma (7,50154E+13)
    y separador de miles con punto (1.234,56).
    También acepta escalares numéricos de numpy/pandas ya leídos por read_csv.
    """
    if pd.isna(val):
        return float("nan")
    if isinstance(val, bool):
        return float("nan")
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val)

    if isinstance(val, np.generic):
        return float(val)

    try:
        return float(val)
    except (TypeError, ValueError):
        pass

    s = str(val).strip()
    if not s or s in ("-", "–", "$"):
        return float("nan")

    # Notación científica con coma en la mantisa: 7,50154E+13
    if re.search(r"[Ee]", s):
        s = re.sub(
            r"^([-+]?\d+),(\d+[Ee][+-]?\d+)$",
            r"\1.\2",
            s,
            flags=re.IGNORECASE,
        )
        try:
            return float(s)
        except ValueError:
            return float("nan")

    s_clean = s.replace(" ", "")
    n_comma = s_clean.count(",")
    n_dot = s_clean.count(".")

    if n_comma == 1 and n_dot == 0:
        s_clean = s_clean.replace(",", ".")
    elif n_comma > 0 and n_dot > 0:
        if s_clean.rfind(",") > s_clean.rfind("."):
            s_clean = s_clean.replace(".", "").replace(",", ".")
        else:
            s_clean = s_clean.replace(",", "")
    elif n_comma > 1 and n_dot == 0:
        s_clean = s_clean.replace(",", ".", 1).replace(",", "", n_comma - 1)

    try:
        return float(s_clean)
    except ValueError:
        return float("nan")


def serie_a_float_importe(serie: pd.Series) -> pd.Series:
    """Convierte serie a float; prioriza dtype numérico y luego formato AR con coma."""
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")
    parsed = serie.map(parsear_numero_importe)
    # Si todo falló (p. ej. strings que pandas no interpretó), reintento vectorizado
    if parsed.notna().any() or serie.empty:
        return parsed
    str_serie = serie.astype(str).str.strip()
    str_serie = str_serie.replace("", pd.NA).replace("nan", pd.NA)
    coma_decimal = str_serie.str.contains(",", na=False) & ~str_serie.str.contains(
        r"[Ee]", na=False, regex=True
    )
    tmp = str_serie.copy()
    tmp.loc[coma_decimal] = tmp.loc[coma_decimal].str.replace(".", "", regex=False)
    tmp = tmp.str.replace(",", ".", regex=False)
    fallback = pd.to_numeric(tmp, errors="coerce")
    return fallback


def total_resumen_pantalla(totales: dict[str, float]) -> float:
    """Suma solo las columnas que deben aparecer en el total del resumen UI."""
    return float(sum(totales[c] for c in COLUMNAS_TOTAL_RESUMEN if c in totales))


def totales_resumen_por_mes(
    totales_por_mes: dict[int, dict[str, float]],
) -> dict[int, float]:
    """Total (resumen) por mes, misma regla que COLUMNAS_TOTAL_RESUMEN."""
    return {
        m: float(sum(totales_por_mes[m][c] for c in COLUMNAS_TOTAL_RESUMEN if c in totales_por_mes[m]))
        for m in range(1, 13)
    }


def limpiar_argumento_ruta(valor: str) -> str:
    """Normaliza saltos de línea/tabulaciones accidentales en argumentos de ruta."""
    return valor.replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()


def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Renombra columnas conocidas a un nombre canónico interno."""
    df = df.copy()
    df.columns = [limpiar_nombre_columna_bruto(c) for c in df.columns]
    renombres = {}
    columnas_actuales = list(df.columns)
    for canonica, aliases in ALIAS_COLUMNAS.items():
        for alias in aliases:
            if alias in columnas_actuales:
                renombres[alias] = canonica
                break
    if renombres:
        df = df.rename(columns=renombres)
    # Evita df["Imp. Total"] como DataFrame si hubo nombres duplicados tras renombrar
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()].copy()
    return df


def _columnas_requeridas_lectura() -> set[str]:
    return set(COLUMNAS_A_AJUSTAR + ["Tipo", "Tipo Cambio", "Fecha Emisión"])


def serie_codigo_tipo_comprobante(serie_tipo: pd.Series) -> pd.Series:
    """
    Código numérico AFIP en la columna Tipo (para isin con CODIGOS_*).
    - Excel a veces guarda el código como número (6, 11).
    - En texto: '6 - Factura B', '006 - ...', o guiones Unicode (– —) mal interpretados
      por split(' - '): se normalizan y, si hace falta, se toman dígitos iniciales.
    """
    s = serie_tipo
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")

    t = s.astype(str).str.strip()
    t = t.replace({"nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})
    t = t.str.replace("\u2013", "-").str.replace("\u2014", "-")
    part0 = t.str.split(r"\s*-\s*", n=1, regex=True).str[0].str.strip()
    codigo = pd.to_numeric(part0, errors="coerce")
    vac = codigo.isna() & t.notna()
    if vac.any():
        dig = t.loc[vac].str.extract(r"^(\d+)", expand=False)
        codigo.loc[vac] = pd.to_numeric(dig, errors="coerce")
    return codigo


def serie_es_comprobante_letra_b(
    codigo_num: pd.Series, serie_tipo: pd.Series
) -> pd.Series:
    """
    Indica filas de comprobantes con letra B (AFIP).
    Usado en emitidos: esas filas no redirigen Imp. Total a Neto Grav. IVA 0%;
    el neto y el IVA siguen las columnas del archivo (con signo y tipo de cambio).
    """
    por_codigo = codigo_num.isin(CODIGOS_LETRA_B_AFIP)
    t = serie_tipo.astype(str)
    up = t.str.upper()
    por_texto = (
        up.str.contains(r"FACTURA\s+B", regex=True, na=False)
        | up.str.contains(r"NOTA\s+DE\s+D[EÉ]BITO\s+B", regex=True, na=False)
        | up.str.contains(r"NOTA\s+DE\s+CR[EÉ]DITO\s+B", regex=True, na=False)
        | up.str.contains(r"RECIBO\s+B", regex=True, na=False)
        | up.str.contains(r"TIQUE.*FACTURA\s+B", regex=True, na=False)
    )
    return por_codigo | por_texto


def _serie_fecha_emision_a_datetime(serie: pd.Series) -> pd.Series:
    """
    Convierte la columna Fecha Emisión a datetime64 (naïve o con TZ AR si aplica).
    Misma lógica para .xlsx (General/serial) y .csv (texto d/m/y, ISO, etc.).
    """
    serie = serie.reset_index(drop=True)
    n = len(serie)
    non_null = serie.notna()
    if n == 0 or not non_null.any():
        return pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")

    # --- 1) Serial Excel ---
    sn = pd.to_numeric(serie, errors="coerce")
    if non_null.any() and bool((sn.notna() == non_null).all()):
        vmin, vmax = float(sn[non_null].min()), float(sn[non_null].max())
        if vmin > 20000 and vmax < 80000:
            return pd.to_datetime(sn, unit="D", origin="1899-12-30", errors="coerce")

    # --- 2) datetime64 ---
    if pd.api.types.is_datetime64_any_dtype(serie.dtype):
        fechas = pd.to_datetime(serie, errors="coerce")
        if pd.api.types.is_datetime64tz_dtype(fechas):
            try:
                fechas = fechas.dt.tz_convert("America/Argentina/Buenos_Aires")
            except Exception:
                pass
        return fechas

    # --- 3) Texto ---
    s = serie.astype(str).str.strip()
    s = s.replace({"nan": pd.NA, "None": pd.NA, "<NA>": pd.NA, "NaT": pd.NA})
    s = s.str.replace(r"\s+\d{1,2}:\d{2}.*$", "", regex=True)
    # ISO primero (evita warning dayfirst con yyyy-mm-dd típico de CSV exportado)
    fechas = pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
    pend = fechas.isna() & s.notna()
    if pend.any():
        fechas.loc[pend] = pd.to_datetime(s.loc[pend], dayfirst=True, errors="coerce")
    pend = fechas.isna() & s.notna()
    if pend.any():
        fechas.loc[pend] = pd.to_datetime(
            s.loc[pend], format="%d/%m/%Y", errors="coerce"
        )
    pend = fechas.isna() & s.notna()
    if pend.any():
        fechas.loc[pend] = pd.to_datetime(
            s.loc[pend], format="%d-%m-%Y", errors="coerce"
        )
    pend = fechas.isna() & s.notna()
    if pend.any():
        try:
            fechas.loc[pend] = pd.to_datetime(
                s.loc[pend], dayfirst=True, errors="coerce", format="mixed"
            )
        except (TypeError, ValueError):
            pass

    return fechas


def _mes_fila_fecha_emision(
    df: pd.DataFrame, nombre_archivo: str | None
) -> pd.Series:
    """
    Mes calendario (1-12) por fila según Fecha Emisión (misma lógica .xlsx y .csv).
    """
    _ = nombre_archivo
    fechas = _serie_fecha_emision_a_datetime(df["Fecha Emisión"])
    mes = fechas.dt.month.astype(float)
    return mes.where(fechas.notna(), np.nan).reset_index(drop=True)


def _formatear_fecha_emision_salida_excel(df: pd.DataFrame) -> None:
    """
    Deja Fecha Emisión como texto dd/mm/yyyy en el DataFrame que se exporta a .xlsx,
    igual que suele verse al procesar .xlsx (evita que CSV salga como yyyy-mm-dd).
    No altera totales: debe llamarse después de _totales_anuales_y_por_mes.
    """
    if "Fecha Emisión" not in df.columns:
        return
    orig = df["Fecha Emisión"]
    fechas = _serie_fecha_emision_a_datetime(orig)
    texto = fechas.dt.strftime("%d/%m/%Y")
    df["Fecha Emisión"] = np.where(
        fechas.notna(),
        texto,
        np.where(orig.notna(), orig.astype(str), ""),
    )


def _totales_anuales_y_por_mes(
    df_ajustado: pd.DataFrame,
    columnas: list[str],
    nombre_archivo: str | None,
) -> tuple[dict[str, float], dict[int, dict[str, float]]]:
    """
    Suma por columna (todas las filas) y acumulado por mes 1-12 según Fecha Emisión
    de cada fila del propio df ajustado (mismo criterio que el archivo de salida).
    """
    block = df_ajustado[columnas].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    block_arr = block.to_numpy(dtype=np.float64, copy=False)
    resultado = {c: float(block_arr[:, j].sum()) for j, c in enumerate(columnas)}

    totales_por_mes: dict[int, dict[str, float]] = {
        m: {c: 0.0 for c in columnas} for m in range(1, 13)
    }
    mes_fila = _mes_fila_fecha_emision(df_ajustado, nombre_archivo)
    mes_np = mes_fila.to_numpy(dtype=float, copy=False)
    n = block_arr.shape[0]
    if len(mes_np) != n:
        aligned = np.full(n, np.nan, dtype=np.float64)
        k = min(n, len(mes_np))
        if k > 0:
            aligned[:k] = mes_np[:k]
        mes_np = aligned
    for pos in range(n):
        m = mes_np[pos]
        if not np.isfinite(m):
            continue
        mi = int(m)
        if mi < 1 or mi > 12:
            continue
        for j, c in enumerate(columnas):
            totales_por_mes[mi][c] += float(block_arr[pos, j])

    return resultado, totales_por_mes


def _mensaje_procesamiento(ui_lang: str, *, es: str, en: str) -> str:
    """Errores de validación: español solo si la UI es es; en cualquier otro idioma, inglés."""
    return es if ui_lang == "es" else en


def _mejor_dataframe_excel(entrada, hoja: str | int, ui_lang: str = "en") -> pd.DataFrame:
    """
    Prueba header=0 y header=1 en .xlsx y elige el que tenga todas las columnas requeridas.
    Si ninguno las tiene todas, elige el que menos falte (misma lógica que CSV vs xlsx).
    """
    req = _columnas_requeridas_lectura()
    opciones: list[tuple[int, int, pd.DataFrame]] = []

    for header_row in (0, 1):
        if hasattr(entrada, "seek"):
            entrada.seek(0)
        try:
            raw = pd.read_excel(entrada, sheet_name=hoja, header=header_row)
        except Exception:
            continue
        raw.columns = raw.columns.astype(str).str.strip()
        cand = normalizar_columnas(raw)
        faltan = len(req - set(cand.columns))
        opciones.append((faltan, header_row, cand))

    if not opciones:
        raise ValueError(
            _mensaje_procesamiento(
                ui_lang,
                es="No se pudo leer el archivo Excel.",
                en="Could not read the Excel file.",
            )
        )

    opciones.sort(key=lambda t: (t[0], t[1]))
    mejor_faltan, _hdr, mejor_df = opciones[0]
    if mejor_faltan > 0:
        nombres = ", ".join(mejor_df.columns.astype(str))
        faltantes = list(req - set(mejor_df.columns))
        raise ValueError(
            _mensaje_procesamiento(
                ui_lang,
                es=(
                    f"No se encontraron las columnas: {faltantes}. "
                    f"Columnas en el archivo: {nombres}"
                ),
                en=(
                    f"Missing columns: {faltantes}. "
                    f"Columns in the file: {nombres}"
                ),
            )
        )
    return mejor_df


def leer_tabla(
    entrada,
    hoja: str | int = 0,
    nombre_archivo: str | None = None,
    ui_lang: str = "en",
) -> pd.DataFrame:
    """
    Lee un .xlsx o .csv con formato:
    - .xlsx: fila 1 encabezado general, fila 2 encabezados de columnas, fila 3+ datos
    - .csv: fila 1 encabezados de columnas, fila 2+ datos
    """
    nombre = (nombre_archivo or str(entrada)).lower()
    if nombre.endswith(".csv"):
        # CSV: encabezados en fila 1 (excepto archivos con primera línea "sep=;")
        # Se prueban varias combinaciones y se elige la que contiene columnas requeridas.
        columnas_requeridas = _columnas_requeridas_lectura()
        muestra = ""
        delimitadores = [";", ",", "\t", "|"]
        skiprows_opciones = [0]

        try:
            if hasattr(entrada, "seek"):
                entrada.seek(0)
            if hasattr(entrada, "read"):
                muestra = entrada.read(8192)
                if isinstance(muestra, bytes):
                    muestra = muestra.decode("utf-8", errors="ignore")
        finally:
            if hasattr(entrada, "seek"):
                entrada.seek(0)

        lineas = [ln.strip() for ln in muestra.splitlines() if ln.strip()]
        primera = lineas[0] if lineas else ""
        if primera.lower().startswith("sep="):
            sep_decl = primera.split("=", 1)[1].strip()
            if sep_decl:
                delim_decl = sep_decl[0]
                delimitadores = [delim_decl] + [d for d in delimitadores if d != delim_decl]
            skiprows_opciones = [1, 0]
        else:
            try:
                dialecto = csv.Sniffer().sniff(muestra or "", delimiters=";,|\t,")
                delim_sniff = dialecto.delimiter
                delimitadores = [delim_sniff] + [d for d in delimitadores if d != delim_sniff]
            except Exception:
                pass

        primer_df = None
        df = None
        for skiprows in skiprows_opciones:
            for delimitador in delimitadores:
                for on_bad_lines in (None, "skip"):
                    if hasattr(entrada, "seek"):
                        entrada.seek(0)
                    kwargs = {
                        "header": 0,
                        "skiprows": skiprows,
                        "sep": delimitador,
                        "engine": "python",
                        "skipinitialspace": True,
                        "encoding": "utf-8-sig",
                        # ARCA / Argentina: día primero en fechas ambiguas al inferir dtypes
                        "dayfirst": True,
                    }
                    if on_bad_lines is not None:
                        kwargs["on_bad_lines"] = on_bad_lines
                    try:
                        candidato = pd.read_csv(entrada, **kwargs)
                    except pd.errors.ParserError:
                        continue
                    except Exception:
                        continue

                    candidato.columns = candidato.columns.astype(str).str.strip()
                    candidato = normalizar_columnas(candidato)
                    if primer_df is None:
                        primer_df = candidato
                    if columnas_requeridas.issubset(set(candidato.columns)):
                        df = candidato
                        break
                if df is not None:
                    break
            if df is not None:
                break

        if df is None:
            # Fallback: primera lectura exitosa aunque no tenga todas las columnas.
            if primer_df is not None:
                df = primer_df
            else:
                raise ValueError(
                    _mensaje_procesamiento(
                        ui_lang,
                        es="No se pudo leer el CSV con un formato válido.",
                        en="Could not read the CSV in a valid format.",
                    )
                )
    else:
        df = _mejor_dataframe_excel(entrada, hoja, ui_lang=ui_lang)

    df.columns = df.columns.astype(str).str.strip()
    return normalizar_columnas(df)


def procesar_archivo(
    ruta_excel: str,
    hoja: str | int = 0,
    nombre_archivo: str | None = None,
    ui_lang: str = "en",
    emitidos: bool = False,
) -> tuple[pd.DataFrame, dict[str, float], dict[int, dict[str, float]]]:
    """
    Lee un archivo Excel y devuelve la sumatoria de las columnas indicadas.
    Fila 1 = encabezado general, fila 2 = encabezados de columnas, datos desde fila 3.
    Las filas con Tipo = nota de crédito se suman en valor negativo.
    Comprobantes CODIGOS_IMP_TOTAL_EN_NETO_IVA_0 (B/C y afines), en .xlsx y .csv: el importe
    suele estar solo en Imp. Total; se refleja en Neto Grav. IVA 0% y en Imp. Total (misma base
    antes de signo NC y tipo de cambio), Neto Gravado Total en 0 y el resto de columnas ajustadas en 0.

    Si emitidos=True: los comprobantes con letra B no aplican esa redirección (neto e IVA en sus
    columnas); los tipo C siguen la misma lógica que en recibidos. Siempre se aplica tipo de cambio.

    Args:
        ruta_excel: Ruta al archivo .xlsx
        hoja: Nombre o índice de la hoja (0 por defecto)

    Returns:
        Tuple con:
        - DataFrame ajustado (columnas numéricas con signo aplicado según Tipo)
        - Diccionario con el nombre de cada columna y su suma.
        - Diccionario mes (1-12) -> totales por columna (mismas claves que totales).
    """
    # header=1: la fila 2 del archivo (índice 1) tiene los nombres de columnas; datos desde fila 3
    df = leer_tabla(
        ruta_excel, hoja=hoja, nombre_archivo=nombre_archivo, ui_lang=ui_lang
    )

    # Comprobar que existan todas las columnas necesarias
    columnas_requeridas = COLUMNAS_A_AJUSTAR + ["Tipo", "Tipo Cambio", "Fecha Emisión"]
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        nombres = ", ".join(df.columns.astype(str))
        raise ValueError(
            _mensaje_procesamiento(
                ui_lang,
                es=(
                    f"No se encontraron las columnas: {faltantes}. "
                    f"Columnas en el archivo: {nombres}"
                ),
                en=(
                    f"Missing columns: {faltantes}. "
                    f"Columns in the file: {nombres}"
                ),
            )
        )

    df = df.reset_index(drop=True)

    # Signo: -1 si es nota de crédito, +1 si no (código numérico AFIP)
    codigo_num = serie_codigo_tipo_comprobante(df["Tipo"])
    es_nota_credito = codigo_num.isin(CODIGOS_NOTA_CREDITO)
    signo = (1 - 2 * es_nota_credito.astype(int)).reset_index(drop=True)

    base_imp_neto_iva_0 = codigo_num.isin(CODIGOS_IMP_TOTAL_EN_NETO_IVA_0)
    if emitidos:
        es_letra_b = serie_es_comprobante_letra_b(codigo_num, df["Tipo"])
        es_imp_en_neto_iva_0 = (base_imp_neto_iva_0 & ~es_letra_b).reset_index(drop=True)
    else:
        es_imp_en_neto_iva_0 = base_imp_neto_iva_0.reset_index(drop=True)

    # Factor de conversión por fila: vacíos/no numéricos se toman como 0 solo para cálculo
    tipo_cambio = serie_a_float_importe(df["Tipo Cambio"]).fillna(0).reset_index(
        drop=True
    )

    # Ajustar signos y tipo de cambio en el DataFrame de salida, luego acumular totales
    df_ajustado = df.copy()
    resultado: dict[str, float] = {}
    imp_total_num = serie_a_float_importe(df["Imp. Total"]).fillna(0).reset_index(
        drop=True
    )
    neto_grav_num = serie_a_float_importe(df["Neto Gravado Total"]).fillna(0).reset_index(
        drop=True
    )
    neto_iva0_num = serie_a_float_importe(df["Neto Grav. IVA 0%"]).fillna(0).reset_index(
        drop=True
    )

    for col in COLUMNAS_A_AJUSTAR:
        if col == "Neto Grav. IVA 0%":
            valores = neto_iva0_num.where(~es_imp_en_neto_iva_0, imp_total_num)
        elif col == "Neto Gravado Total":
            valores = neto_grav_num.where(~es_imp_en_neto_iva_0, 0.0)
        elif col == "Imp. Total":
            # Misma base que Neto Grav. IVA 0% en B/C: coherencia tras signo y tipo de cambio
            valores = imp_total_num
        else:
            base = serie_a_float_importe(df[col]).fillna(0).reset_index(drop=True)
            valores = base.where(~es_imp_en_neto_iva_0, 0.0)
        valores_ajustados = (valores * signo * tipo_cambio).astype(float)
        df_ajustado[col] = valores_ajustados.values

    resultado, totales_por_mes = _totales_anuales_y_por_mes(
        df_ajustado, COLUMNAS_A_AJUSTAR, nombre_archivo
    )
    _formatear_fecha_emision_salida_excel(df_ajustado)

    return df_ajustado, resultado, totales_por_mes


# Contabilidad (Excel): alineación, negativos entre paréntesis, cero como guión; sin símbolo de moneda.
# Equivalente a “Contabilidad” sin divisa en la cinta de Excel.
_FORMATO_CONTABILIDAD_SIN_MONEDA = (
    r'_ * #,##0.00_ ;_ * (#,##0.00)_ ;_ * "-"??_ ;_ @_ '
)

_EPS_CERO = 1e-12


def _es_cero_numerico(val) -> bool:
    if val is None or val == "":
        return False
    if type(val) is bool:
        return False
    if isinstance(val, (int, float)):
        return abs(float(val)) < _EPS_CERO
    return False


def _longitud_texto_celda_excel(val) -> int:
    """Longitud aproximada del texto mostrado (encabezados; importes como en contabilidad sin moneda)."""
    if val is None or val == "":
        return 0
    if type(val) is bool:
        return len(str(val))
    if isinstance(val, (int, float)):
        if _es_cero_numerico(val):
            return 3
        fv = float(val)
        if fv < 0:
            return len(f"({abs(fv):,.2f})")
        return len(f"{fv:,.2f}")
    return len(str(val))


def escribir_excel_ajustado_con_formato(
    df: pd.DataFrame, destino: io.BytesIO | Path | str
) -> None:
    """
    Escribe el DataFrame en .xlsx con:
    - Fila 1 (encabezados) en negrita.
    - Celdas numéricas de COLUMNAS_A_AJUSTAR: formato contabilidad sin moneda (alineación, negativos
      entre paréntesis, cero como guión).
    - Ancho de cada una de esas columnas según el contenido más largo (encabezado o valores).
    - Fila de encabezado fija al desplazarse (freeze panes en fila 1).
    - Autofiltro activo sobre la tabla (fila de títulos con filtros en Excel).
    """
    temp = io.BytesIO()
    df.to_excel(temp, index=False, engine="openpyxl")
    temp.seek(0)
    wb = load_workbook(temp)
    ws = wb.active
    negrita = Font(bold=True)

    for cell in ws[1]:
        cell.font = negrita

    encabezados = [c.value for c in ws[1]]
    for nombre_col in COLUMNAS_A_AJUSTAR:
        try:
            idx = encabezados.index(nombre_col) + 1
        except ValueError:
            continue
        for fila in range(2, ws.max_row + 1):
            celda = ws.cell(row=fila, column=idx)
            if celda.value is not None and celda.value != "":
                celda.number_format = _FORMATO_CONTABILIDAD_SIN_MONEDA

        max_long = 0
        for fila in range(1, ws.max_row + 1):
            max_long = max(
                max_long,
                _longitud_texto_celda_excel(ws.cell(row=fila, column=idx).value),
            )
        if max_long > 0:
            ws.column_dimensions[get_column_letter(idx)].width = max_long + 1.5

    ws.freeze_panes = "A2"

    if ws.max_row >= 1 and ws.max_column >= 1:
        ultima_col = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A1:{ultima_col}{ws.max_row}"

    if isinstance(destino, io.BytesIO):
        destino.seek(0)
        destino.truncate(0)
        wb.save(destino)
        destino.seek(0)
    else:
        wb.save(destino)


def construir_ruta_salida(ruta_excel: str, salida_arg: str | None) -> Path:
    """Devuelve la ruta de salida para el excel ajustado."""
    if salida_arg:
        return Path(salida_arg)

    origen = Path(ruta_excel)
    return origen.with_name(f"{origen.stem}_ajustado.xlsx")


def main():
    if len(sys.argv) < 2:
        print("Uso: python sumar_imp_total.py <archivo.xlsx|archivo.csv> [hoja] [salida.xlsx]")
        print("  hoja: nombre o número de hoja (opcional, por defecto la primera)")
        print("  salida.xlsx: ruta del excel de salida (opcional)")
        print("Columnas sumadas desde la fila 3:", ", ".join(COLUMNAS_A_AJUSTAR))
        sys.exit(1)

    ruta = limpiar_argumento_ruta(sys.argv[1])
    hoja = sys.argv[2] if len(sys.argv) > 2 else 0
    salida_arg = limpiar_argumento_ruta(sys.argv[3]) if len(sys.argv) > 3 else None
    if isinstance(hoja, str) and hoja.isdigit():
        hoja = int(hoja)

    try:
        df_ajustado, totales, _ = procesar_archivo(ruta, hoja, nombre_archivo=ruta)
        ruta_salida = construir_ruta_salida(ruta, salida_arg)
        escribir_excel_ajustado_con_formato(df_ajustado, ruta_salida)

        print("Sumas (desde fila 3):")
        for col, total in totales.items():
            print(f"  {col}: {total:,.2f}")
        print("  ---")
        print(f"  Suma total (resumen): {total_resumen_pantalla(totales):,.2f}")
        print(f"Excel ajustado generado en: {ruta_salida}")
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo '{ruta}'", file=sys.stderr)
        sys.exit(2)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(3)


if __name__ == "__main__":
    main()
