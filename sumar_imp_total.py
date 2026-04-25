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
from collections import defaultdict
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
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

# Notas de crédito: neto gravado e IVA por alícuota (2,5%–27%; resumen y reparto mensual)
COLUMNAS_NETO_NC_ALICUOTA = [
    "Neto Grav. IVA 2,5%",
    "Neto Grav. IVA 5%",
    "Neto Grav. IVA 10,5%",
    "Neto Grav. IVA 21%",
    "Neto Grav. IVA 27%",
]
COLUMNAS_IVA_NC_ALICUOTA = [
    "IVA 2,5%",
    "IVA 5%",
    "IVA 10,5%",
    "IVA 21%",
    "IVA 27%",
]

# Suma "Neto" en informe por proveedor/cliente (mismas columnas; valores ya ajustados)
COLUMNAS_NETO_INFORME_CONTRAPARTE = [
    "Neto Gravado Total",
    "Neto Grav. IVA 0%",
    "Neto No Gravado",
    "Op. Exentas",
]

# Códigos numéricos de la columna "Tipo" que se consideran nota de crédito (suma en negativo)
# Se matchea por número (sin ceros a la izquierda): "003", "3" y "03" son el mismo código
CODIGOS_NOTA_CREDITO = {
    3, 8, 13, 21, 38, 43, 44, 48, 53,
    110, 112, 113, 114, 203, 206, 208, 211, 213,
}

# Tipos B/C (y afines): Imp. Total → Neto Grav. IVA 0% (no Neto Gravado Total). Columna Tipo.
# Comprobantes recibidos: se aplica a todo este conjunto (export ARCA suele concentrar el total en Imp. Total).
CODIGOS_IMP_TOTAL_EN_NETO_IVA_0 = frozenset(
    (
        6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 18, 19, 20, 21, 25, 26, 28, 29,
        40, 41, 42, 43, 44, 46, 47, 61, 64, 82, 83, 90, 91, 109, 110, 111,
        113, 114, 116, 117, 206, 207, 208, 211, 212, 213,
    )
)

# Solo clase C (AFIP Libro IVA / RG 4290 y afines; FCE MiPyME C = 211–213).
# Comprobantes emitidos: únicamente estos replican Imp. Total en Neto Grav. IVA 0% y anulan el resto
# de bases imponibles en esa fila; el resto de tipos usan las columnas del archivo (con signo NC).
CODIGOS_IMP_TOTAL_EN_NETO_IVA_0_SOLO_C = frozenset(
    (
        11,
        12,
        13,  # Factura / ND / NC C
        15,
        16,  # Recibo C, nota venta al contado C
        109,
        111,
        114,
        117,  # Tique C, tique factura C, tique NC/ND C
        211,
        212,
        213,  # FCE electrónica MiPyME C
    )
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
    "Denominación Emisor": [
        "Denominación Emisor",
        "Denominacion Emisor",
        "Razón Social Emisor",
    ],
    "Denominación Receptor": [
        "Denominación Receptor",
        "Denominacion Receptor",
        "Razón Social Receptor",
    ],
    "Nro. Doc. Emisor": [
        "Nro. Doc. Emisor",
        "Nro Doc. Emisor",
        "Nro Doc Emisor",
    ],
    "Nro. Doc. Receptor": [
        "Nro. Doc. Receptor",
        "Nro Doc. Receptor",
        "Nro Doc Receptor",
    ],
    "Número Hasta": ["Número Hasta", "Numero Hasta", "Nro. Hasta"],
    "Cód. Autorización": ["Cód. Autorización", "Cod. Autorización", "Cod Autorizacion"],
    "Tipo Doc. Emisor": ["Tipo Doc. Emisor", "Tipo Documento Emisor"],
    "Tipo Doc. Receptor": ["Tipo Doc. Receptor", "Tipo Documento Receptor"],
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


def totales_resumen_por_periodo(
    totales_por_periodo: dict[str, dict[str, float]],
) -> dict[str, float]:
    """Total (resumen) por mes-año, misma regla que COLUMNAS_TOTAL_RESUMEN."""
    return {
        p: float(
            sum(
                totales_por_periodo[p][c]
                for c in COLUMNAS_TOTAL_RESUMEN
                if c in totales_por_periodo[p]
            )
        )
        for p in totales_por_periodo
    }


def periodos_orden_crono(*dicts) -> list[str]:
    """Unión de claves YYYY-MM presentes en diccionarios, orden cronológico."""
    keys: set[str] = set()
    for d in dicts:
        if d:
            keys |= set(d.keys())
    return sorted(keys)


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


def _totales_anuales_y_por_periodo(
    df_ajustado: pd.DataFrame,
    columnas: list[str],
    nombre_archivo: str | None,
) -> tuple[dict[str, float], dict[str, dict[str, float]]]:
    """
    Suma por columna (todas las filas) y acumulado por (año, mes) según Fecha Emisión.
    Claves de período: ``\"YYYY-MM\"`` (solo períodos con al menos un comprobante con fecha).
    """
    _ = nombre_archivo
    block = df_ajustado[columnas].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    block_arr = block.to_numpy(dtype=np.float64, copy=False)
    resultado = {c: float(block_arr[:, j].sum()) for j, c in enumerate(columnas)}

    totales_por_periodo: dict[str, dict[str, float]] = {}

    fechas = _serie_fecha_emision_a_datetime(df_ajustado["Fecha Emisión"])
    años = fechas.dt.year
    meses = fechas.dt.month
    n = block_arr.shape[0]
    for pos in range(n):
        if pd.isna(fechas.iloc[pos]):
            continue
        y = int(años.iloc[pos])
        m = int(meses.iloc[pos])
        if m < 1 or m > 12:
            continue
        key = f"{y:04d}-{m:02d}"
        if key not in totales_por_periodo:
            totales_por_periodo[key] = {c: 0.0 for c in columnas}
        for j, c in enumerate(columnas):
            totales_por_periodo[key][c] += float(block_arr[pos, j])

    return resultado, totales_por_periodo


def totales_notas_credito_neto_e_iva_por_periodo(
    df_ajustado: pd.DataFrame,
    nombre_archivo: str | None,
) -> tuple[float, float, dict[str, float], dict[str, float]]:
    """
    Suma solo filas clasificadas como nota de crédito, sobre:
    - neto gravado 2,5% … 27% (sin 0%),
    - IVA 2,5% … 27%.

    Usa los valores ya en ``df_ajustado`` (misma regla B/C, signo NC y tipo de cambio
    que el resto). Período por Fecha de emisión, clave ``\"YYYY-MM\"``.
    Debe llamarse antes de formatear Fecha Emisión a texto si se desea; Tipo
    se sigue leyendo para el mask NC.
    """
    _ = nombre_archivo
    codigo_num = serie_codigo_tipo_comprobante(df_ajustado["Tipo"])
    es_nc = codigo_num.isin(CODIGOS_NOTA_CREDITO).to_numpy(dtype=bool)

    block_neto = df_ajustado[COLUMNAS_NETO_NC_ALICUOTA].apply(
        pd.to_numeric, errors="coerce"
    ).fillna(0.0)
    block_iva = df_ajustado[COLUMNAS_IVA_NC_ALICUOTA].apply(
        pd.to_numeric, errors="coerce"
    ).fillna(0.0)
    neto_row = block_neto.to_numpy(dtype=np.float64, copy=False).sum(axis=1)
    iva_row = block_iva.to_numpy(dtype=np.float64, copy=False).sum(axis=1)

    total_neto = float(neto_row[es_nc].sum()) if es_nc.any() else 0.0
    total_iva = float(iva_row[es_nc].sum()) if es_nc.any() else 0.0

    neto_por: dict[str, float] = defaultdict(float)
    iva_por: dict[str, float] = defaultdict(float)

    fechas = _serie_fecha_emision_a_datetime(df_ajustado["Fecha Emisión"])
    años = fechas.dt.year
    meses = fechas.dt.month
    n = int(block_neto.shape[0])
    if len(es_nc) > n:
        es_nc = es_nc[:n]
    elif len(es_nc) < n:
        es_nc = np.pad(es_nc, (0, n - len(es_nc)), constant_values=False)

    for pos in range(n):
        if not es_nc[pos]:
            continue
        if pd.isna(fechas.iloc[pos]):
            continue
        y = int(años.iloc[pos])
        m = int(meses.iloc[pos])
        if m < 1 or m > 12:
            continue
        key = f"{y:04d}-{m:02d}"
        neto_por[key] += float(neto_row[pos])
        iva_por[key] += float(iva_row[pos])

    return total_neto, total_iva, dict(neto_por), dict(iva_por)


def _normalizar_clave_cuit_doc(val) -> str:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    s = re.sub(r"\D", "", str(val).strip())
    return s


def acumulado_por_contraparte(
    df_ajustado: pd.DataFrame, emitidos: bool
) -> list[dict[str, Any]]:
    """
    Un acumulado por proveedor (recibidos) o cliente (emitidos): Nombre, CUIT,
    Neto (cuatro conceptos de neto), IVA (Total IVA) y total (Imp. Total).
    Valores de ``df_ajustado`` (signo y tipo de cambio ya aplicados).
    """
    col_nom = "Denominación Receptor" if emitidos else "Denominación Emisor"
    col_doc = "Nro. Doc. Receptor" if emitidos else "Nro. Doc. Emisor"
    for c in [col_nom, col_doc, "Total IVA", "Imp. Total"] + COLUMNAS_NETO_INFORME_CONTRAPARTE:
        if c not in df_ajustado.columns:
            return []

    d = df_ajustado[
        [col_nom, col_doc, "Total IVA", "Imp. Total"] + COLUMNAS_NETO_INFORME_CONTRAPARTE
    ].copy()
    d["_k"] = d[col_doc].map(_normalizar_clave_cuit_doc)
    neto_sum = (
        d[COLUMNAS_NETO_INFORME_CONTRAPARTE]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0)
        .sum(axis=1)
    )
    d["_neto"] = neto_sum
    d["_iva"] = pd.to_numeric(d["Total IVA"], errors="coerce").fillna(0)
    d["_tot"] = pd.to_numeric(d["Imp. Total"], errors="coerce").fillna(0)

    out: list[dict[str, Any]] = []
    for cuit_key, grp in d.groupby("_k", sort=False):
        noms = grp[col_nom].astype(str).str.strip()
        noms = noms.replace("", pd.NA).replace("nan", pd.NA)
        if noms.notna().any():
            mod = noms.mode()
            nombre = str(mod.iloc[0]) if len(mod) else str(noms.dropna().iloc[0])
        else:
            nombre = ""
        cuit_show = str(grp[col_doc].astype(str).str.strip().iloc[0]) if cuit_key else ""
        out.append(
            {
                "nombre": nombre,
                "cuit": cuit_show,
                "neto": float(grp["_neto"].sum()),
                "iva": float(grp["_iva"].sum()),
                "total": float(grp["_tot"].sum()),
            }
        )
    out.sort(
        key=lambda r: ((r["nombre"] or "zzz").lower(), r["cuit"] or "")
    )
    return out


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
    ruta_excel: str | io.BytesIO,
    hoja: str | int = 0,
    nombre_archivo: str | None = None,
    ui_lang: str = "en",
    *,
    emitidos: bool = False,
) -> tuple[
    pd.DataFrame,
    dict[str, float],
    dict[str, dict[str, float]],
    dict,
    list[dict[str, Any]],
]:
    """
    Lee un archivo Excel y devuelve la sumatoria de las columnas indicadas.
    Fila 1 = encabezado general, fila 2 = encabezados de columnas, datos desde fila 3.
    Las filas con Tipo = nota de crédito se suman en valor negativo.
    Comprobantes recibidos (emitidos=False): tipos en CODIGOS_IMP_TOTAL_EN_NETO_IVA_0 (B/C y afines):
    el importe suele estar solo en Imp. Total; se refleja en Neto Grav. IVA 0% y en Imp. Total,
    Neto Gravado Total en 0 y el resto de columnas numéricas de esa fila en 0 (antes de signo NC y TC).

    Comprobantes emitidos (emitidos=True): solo los de clase C (CODIGOS_IMP_TOTAL_EN_NETO_IVA_0_SOLO_C)
    siguen esa regla; el resto suma con los valores de cada columna del archivo (Factura A/B, FCE A/B, etc.).

    Args:
        ruta_excel: Ruta al archivo .xlsx o buffer legible por pandas (p. ej. BytesIO)
        hoja: Nombre o índice de la hoja (0 por defecto)
        emitidos: True si el archivo es export de comprobantes emitidos.

    Returns:
        Tuple con:
        - DataFrame ajustado (columnas numéricas con signo aplicado según Tipo)
        - Diccionario con el nombre de cada columna y su suma.
        - Diccionario período (``\"YYYY-MM\"``) -> totales por columna.
        - Diccionario con «Neto/IVA notas de crédito» (anual y por período) alícuotas 2,5%–27%.
        - Tabla de acumulados por proveedor o cliente (Nombre, CUIT, Neto, IVA, Total).
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

    codigos_neto_iva_0 = (
        CODIGOS_IMP_TOTAL_EN_NETO_IVA_0_SOLO_C
        if emitidos
        else CODIGOS_IMP_TOTAL_EN_NETO_IVA_0
    )
    es_imp_en_neto_iva_0 = codigo_num.isin(codigos_neto_iva_0).reset_index(drop=True)

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

    resultado, totales_por_periodo = _totales_anuales_y_por_periodo(
        df_ajustado, COLUMNAS_A_AJUSTAR, nombre_archivo
    )
    t_neto_nc, t_iva_nc, neto_nc_p, iva_nc_p = totales_notas_credito_neto_e_iva_por_periodo(
        df_ajustado, nombre_archivo
    )
    notas_credito_extras = {
        "total_neto_nc": t_neto_nc,
        "total_iva_nc": t_iva_nc,
        "neto_nc_por_periodo": neto_nc_p,
        "iva_nc_por_periodo": iva_nc_p,
    }
    tabla_contrapartes = acumulado_por_contraparte(df_ajustado, emitidos)
    _formatear_fecha_emision_salida_excel(df_ajustado)

    return (
        df_ajustado,
        resultado,
        totales_por_periodo,
        notas_credito_extras,
        tabla_contrapartes,
    )


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


# Columnas a ocultar en hoja de comprobantes (exporte)
_COLUMNA_OCULTA_COMUN = [
    "Número Hasta",
    "Cód. Autorización",
    "Tipo Doc. Emisor",
    "Tipo Doc. Receptor",
]
_COLUMNA_OCULTA_SI_RECIBIDOS = ["Nro. Doc. Receptor"]
_ANCHO_DENOM = 40
_SHEET_COMPR = "Comprobantes"


def _nombres_columnas_ocultar(emitidos: bool) -> set[str]:
    s = set(_COLUMNA_OCULTA_COMUN)
    if not emitidos:
        s |= set(_COLUMNA_OCULTA_SI_RECIBIDOS)
    return s


def _aplicar_hoja_comprobantes_excel(
    _wb,
    ws,
    encabezados: list,
    emitidos: bool,
) -> None:
    negrita = Font(bold=True)
    for cell in ws[1]:
        cell.font = negrita

    ocultar = _nombres_columnas_ocultar(emitidos)
    denom_anchura = {
        "Denominación Emisor",
        "Denominación Receptor",
    }

    for col_i, nombre in enumerate(encabezados, start=1):
        if not nombre:
            continue
        letra = get_column_letter(col_i)
        dim = ws.column_dimensions[letra]
        if nombre in ocultar:
            dim.hidden = True
            dim.width = 0.5
            continue
        if nombre in denom_anchura:
            dim.width = float(_ANCHO_DENOM)
            for fila in range(2, ws.max_row + 1):
                c = ws.cell(row=fila, column=col_i)
                if c.value and isinstance(c.value, str) and len(c.value) > _ANCHO_DENOM * 1.2:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
            continue

        max_long = 0
        for fila in range(1, ws.max_row + 1):
            v = ws.cell(row=fila, column=col_i).value
            max_long = max(max_long, _longitud_texto_celda_excel(v), len(str(nombre) if nombre else ""))
        if max_long > 0:
            dim.width = min(max_long + 2, 60)

    for col_i, nombre in enumerate(encabezados, start=1):
        if nombre in COLUMNAS_A_AJUSTAR:
            for fila in range(2, ws.max_row + 1):
                cel = ws.cell(row=fila, column=col_i)
                if cel.value is not None and cel.value != "":
                    cel.number_format = _FORMATO_CONTABILIDAD_SIN_MONEDA

    ws.title = _SHEET_COMPR
    ws.freeze_panes = "A2"
    if ws.max_row >= 1 and ws.max_column >= 1:
        ult = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A1:{ult}{ws.max_row}"


def _celda_num(ws, r: int, c: int, v: object) -> None:
    cl = ws.cell(row=r, column=c, value=v)
    if v is not None and isinstance(v, (int, float)) and v != "":
        cl.number_format = _FORMATO_CONTABILIDAD_SIN_MONEDA


def _fila_monto(
    wsr,
    fila: int,
    etiq: str,
    valor: object,
    bold_l: bool = False,
) -> int:
    c1 = wsr.cell(row=fila, column=1, value=etiq)
    if bold_l:
        c1.font = Font(bold=True)
    cl = wsr.cell(row=fila, column=2, value=valor)
    if isinstance(valor, (int, float)):
        cl.number_format = _FORMATO_CONTABILIDAD_SIN_MONEDA
    return fila + 1


def escribir_excel_informe_completo(
    df_ajustado: pd.DataFrame,
    destino: io.BytesIO | Path | str,
    *,
    emitidos: bool,
    totales: dict[str, float],
    totales_por_periodo: dict[str, dict[str, float]],
    periodos_orden: list[str],
    notas_credito_extras: dict,
    totales_resumen: dict,
    totales_detalle: dict,
    suma_total: float,
    columnas_orden: list[str],
    tabla_contrapartes: list[dict],
) -> None:
    temp = io.BytesIO()
    df_ajustado.to_excel(temp, index=False, engine="openpyxl", sheet_name=_SHEET_COMPR)
    temp.seek(0)
    wb = load_workbook(temp)
    ws0 = wb.active
    encab = [c.value for c in ws0[1]]
    _aplicar_hoja_comprobantes_excel(wb, ws0, encab, emitidos)

    negrita = Font(bold=True)
    wsr = wb.create_sheet("Resumen", 1)
    fila = 1
    t1 = wsr.cell(row=1, column=1, value="Resumen (base del total)")
    t1.font = negrita
    t1.alignment = Alignment(horizontal="center", vertical="center")
    wsr.merge_cells("A1:B1")
    fila = 2
    wsr.cell(row=fila, column=1, value="Concepto").font = negrita
    wsr.cell(row=fila, column=2, value="Importe").font = negrita
    fila_tabla0 = 2
    fila = 3
    for k, v in totales_resumen.items():
        wsr.cell(row=fila, column=1, value=k)
        cl = wsr.cell(row=fila, column=2, value=v)
        cl.number_format = _FORMATO_CONTABILIDAD_SIN_MONEDA
        fila += 1
    ctot1 = wsr.cell(row=fila, column=1, value="Total (resumen)")
    ctot1.font = negrita
    ctot2 = wsr.cell(row=fila, column=2, value=suma_total)
    ctot2.font = negrita
    ctot2.number_format = _FORMATO_CONTABILIDAD_SIN_MONEDA
    fila += 1
    fila = _fila_monto(wsr, fila, "Neto Notas de Crédito", notas_credito_extras.get("total_neto_nc", 0))
    fila = _fila_monto(wsr, fila, "IVA Notas de Crédito", notas_credito_extras.get("total_iva_nc", 0))
    u_res = fila - 1
    if totales_detalle:
        fila += 1
        wsr.cell(row=fila, column=1, value="Detalle por alícuota")
        wsr.cell(row=fila, column=1).font = negrita
        wsr.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=2)
        fila += 1
        wsr.cell(row=fila, column=1, value="Concepto").font = negrita
        wsr.cell(row=fila, column=2, value="Importe").font = negrita
        fila += 1
        for k, v in totales_detalle.items():
            wsr.cell(row=fila, column=1, value=k)
            cl = wsr.cell(row=fila, column=2, value=v)
            cl.number_format = _FORMATO_CONTABILIDAD_SIN_MONEDA
            fila += 1
        u_res = fila - 1
    wsr.auto_filter.ref = f"A{fila_tabla0}:B{u_res}"

    wsd = wb.create_sheet("Distribución mensual", 2)
    wsd.append(["Concepto"] + [""] * len(periodos_orden))
    for i, per in enumerate(periodos_orden, start=2):
        y, m = map(int, per.split("-"))
        mes_lbl = NOMBRES_MESES[m - 1].capitalize()
        chead = wsd.cell(row=1, column=i, value=f"{mes_lbl}\n{y}")
        chead.font = negrita
        chead.alignment = Alignment(
            wrap_text=True, horizontal="center", vertical="center"
        )
    c0 = wsd.cell(row=1, column=1, value="Concepto")
    c0.font = negrita
    c0.alignment = Alignment(horizontal="left", vertical="center")
    wsd.row_dimensions[1].height = 32
    fila = 2
    for col in columnas_orden:
        wsd.cell(row=fila, column=1, value=col)
        for j, per in enumerate(periodos_orden, start=2):
            vcel = totales_por_periodo.get(per, {}).get(col, 0.0)
            _celda_num(wsd, fila, j, vcel)
        fila += 1
    wsd.cell(row=fila, column=1, value="Total (resumen)").font = negrita
    for j, per in enumerate(periodos_orden, start=2):
        tr = 0.0
        for cname in COLUMNAS_TOTAL_RESUMEN:
            tr += float(totales_por_periodo.get(per, {}).get(cname, 0.0))
        cc = wsd.cell(row=fila, column=j, value=tr)
        cc.font = negrita
        cc.number_format = _FORMATO_CONTABILIDAD_SIN_MONEDA
    fila += 1
    wsd.cell(row=fila, column=1, value="Neto Notas de Crédito")
    ncp = notas_credito_extras.get("neto_nc_por_periodo", {})
    for j, per in enumerate(periodos_orden, start=2):
        _celda_num(wsd, fila, j, ncp.get(per, 0.0))
    fila += 1
    wsd.cell(row=fila, column=1, value="IVA Notas de Crédito")
    icp = notas_credito_extras.get("iva_nc_por_periodo", {})
    for j, per in enumerate(periodos_orden, start=2):
        _celda_num(wsd, fila, j, icp.get(per, 0.0))
    if wsd.max_column >= 1 and wsd.max_row >= 1:
        wsd.auto_filter.ref = f"A1:{get_column_letter(wsd.max_column)}{wsd.max_row}"
    for c in range(1, wsd.max_column + 1):
        le = get_column_letter(c)
        if c == 1:
            wsd.column_dimensions[le].width = 24
        else:
            wsd.column_dimensions[le].width = 12

    nom_total = "Total clientes" if emitidos else "Total proveedores"
    wsp = wb.create_sheet(nom_total, 3)
    wsp.append(["Nombre", "CUIT", "Neto", "IVA", "Total"])
    for cell in wsp[1]:
        cell.font = negrita
    for r, trow in enumerate(tabla_contrapartes, start=2):
        wsp.cell(row=r, column=1, value=trow.get("nombre", ""))
        wsp.cell(row=r, column=2, value=trow.get("cuit", ""))
        _celda_num(wsp, r, 3, trow.get("neto", 0))
        _celda_num(wsp, r, 4, trow.get("iva", 0))
        _celda_num(wsp, r, 5, trow.get("total", 0))
    end_r = max(1, len(tabla_contrapartes) + 1)
    wsp.auto_filter.ref = f"A1:E{end_r}"
    wsp.column_dimensions["A"].width = 32
    wsp.column_dimensions["B"].width = 16
    for col_l in "CDE":
        wsp.column_dimensions[col_l].width = 16

    if isinstance(destino, io.BytesIO):
        destino.seek(0)
        destino.truncate(0)
        wb.save(destino)
        destino.seek(0)
    else:
        wb.save(destino)


def escribir_excel_ajustado_con_formato(
    df: pd.DataFrame, destino: io.BytesIO | Path | str
) -> None:
    """
    Escribe solo la hoja de comprobantes con formato (compat. con llamadas
    que no construyen hojas de resumen / distribución).
    """
    temp = io.BytesIO()
    df.to_excel(temp, index=False, engine="openpyxl", sheet_name=_SHEET_COMPR)
    temp.seek(0)
    wb = load_workbook(temp)
    ws = wb.active
    encab = [c.value for c in ws[1]]
    _aplicar_hoja_comprobantes_excel(wb, ws, encab, emitidos=False)
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
        (
            df_ajustado,
            totales,
            tpp,
            nce,
            tabla_c,
        ) = procesar_archivo(ruta, hoja, nombre_archivo=ruta)
        ruta_salida = construir_ruta_salida(ruta, salida_arg)
        per_keys = periodos_orden_crono(
            tpp, nce.get("neto_nc_por_periodo", {}), nce.get("iva_nc_por_periodo", {})
        )
        totales_res = {c: totales[c] for c in COLUMNAS_TOTAL_RESUMEN}
        totales_det = {c: totales[c] for c in COLUMNAS_DETALLE_SIN_RESUMEN}
        escribir_excel_informe_completo(
            df_ajustado,
            ruta_salida,
            emitidos=False,
            totales=totales,
            totales_por_periodo=tpp,
            periodos_orden=per_keys,
            notas_credito_extras=nce,
            totales_resumen=totales_res,
            totales_detalle=totales_det,
            suma_total=total_resumen_pantalla(totales),
            columnas_orden=COLUMNAS_A_AJUSTAR,
            tabla_contrapartes=tabla_c,
        )

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
