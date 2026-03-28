"""
Lee un archivo Excel (.xlsx) o CSV, ajusta signos según Tipo y suma columnas indicadas.
En .xlsx se prueba encabezado en fila 1 o fila 2 y se usa el que tenga todas las columnas requeridas.
En .csv los encabezados suelen estar en fila 1.
Las filas con Tipo = nota de crédito (por código numérico) se consideran en negativo.

Uso:
  python sumar_imp_total.py <ruta_al_archivo.xlsx> [hoja] [archivo_salida.xlsx]
"""

import re
import sys
import csv
from pathlib import Path
import numpy as np
import pandas as pd

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

# Comprobantes régimen B / C (código numérico en columna Tipo): Imp. Total pasa a Neto Gravado Total
CODIGOS_GRUPO_B = {6, 7, 8, 9, 206, 208}
CODIGOS_GRUPO_C = {11, 12, 13, 211, 213}

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


def _mejor_dataframe_excel(entrada, hoja: str | int) -> pd.DataFrame:
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
        raise ValueError("No se pudo leer el archivo Excel.")

    opciones.sort(key=lambda t: (t[0], t[1]))
    mejor_faltan, _hdr, mejor_df = opciones[0]
    if mejor_faltan > 0:
        nombres = ", ".join(mejor_df.columns.astype(str))
        faltantes = list(req - set(mejor_df.columns))
        raise ValueError(
            f"No se encontraron las columnas: {faltantes}. "
            f"Columnas en el archivo: {nombres}"
        )
    return mejor_df


def leer_tabla(entrada, hoja: str | int = 0, nombre_archivo: str | None = None) -> pd.DataFrame:
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
                raise ValueError("No se pudo leer el CSV con un formato válido.")
    else:
        df = _mejor_dataframe_excel(entrada, hoja)

    df.columns = df.columns.astype(str).str.strip()
    return normalizar_columnas(df)


def procesar_archivo(
    ruta_excel: str, hoja: str | int = 0, nombre_archivo: str | None = None
) -> tuple[pd.DataFrame, dict[str, float], dict[int, dict[str, float]]]:
    """
    Lee un archivo Excel y devuelve la sumatoria de las columnas indicadas.
    Fila 1 = encabezado general, fila 2 = encabezados de columnas, datos desde fila 3.
    Las filas con Tipo = nota de crédito se suman en valor negativo.
    Comprobantes B/C: el importe de Imp. Total se usa como Neto Gravado Total (ajustado luego).

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
    df = leer_tabla(ruta_excel, hoja=hoja, nombre_archivo=nombre_archivo)

    # Comprobar que existan todas las columnas necesarias
    columnas_requeridas = COLUMNAS_A_AJUSTAR + ["Tipo", "Tipo Cambio", "Fecha Emisión"]
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        nombres = ", ".join(df.columns.astype(str))
        raise ValueError(
            f"No se encontraron las columnas: {faltantes}. "
            f"Columnas en el archivo: {nombres}"
        )

    df = df.reset_index(drop=True)

    # Signo: -1 si es nota de crédito, +1 si no (código como número, sin ceros a la izquierda)
    tipo_str = df["Tipo"].astype(str).str.strip()
    codigo_str = tipo_str.str.split(" - ", n=1).str[0].str.strip()  # ej. "003 - NOTA..." -> "003"
    codigo_num = pd.to_numeric(codigo_str, errors="coerce")  # "003" y "3" -> 3
    es_nota_credito = codigo_num.isin(CODIGOS_NOTA_CREDITO)
    signo = (1 - 2 * es_nota_credito.astype(int)).reset_index(drop=True)

    es_grupo_b_o_c = codigo_num.isin(CODIGOS_GRUPO_B | CODIGOS_GRUPO_C).reset_index(
        drop=True
    )

    # Factor de conversión por fila: vacíos/no numéricos se toman como 0 solo para cálculo
    tipo_cambio = serie_a_float_importe(df["Tipo Cambio"]).fillna(0).reset_index(
        drop=True
    )

    try:
        fechas = pd.to_datetime(
            df["Fecha Emisión"], dayfirst=True, errors="coerce", format="mixed"
        )
    except (TypeError, ValueError):
        fechas = pd.to_datetime(df["Fecha Emisión"], dayfirst=True, errors="coerce")
    mes_fila = fechas.dt.month.reset_index(drop=True)

    # Ajustar signos y tipo de cambio en el DataFrame de salida, luego acumular totales
    df_ajustado = df.copy()
    resultado: dict[str, float] = {}
    imp_total_num = serie_a_float_importe(df["Imp. Total"]).fillna(0).reset_index(
        drop=True
    )
    neto_grav_num = serie_a_float_importe(df["Neto Gravado Total"]).fillna(0).reset_index(
        drop=True
    )

    for col in COLUMNAS_A_AJUSTAR:
        if col == "Neto Gravado Total":
            valores = neto_grav_num.where(~es_grupo_b_o_c, imp_total_num)
        else:
            valores = serie_a_float_importe(df[col]).fillna(0).reset_index(drop=True)
        valores_ajustados = (valores * signo * tipo_cambio).astype(float)
        df_ajustado[col] = valores_ajustados.values

    # Totales y por mes desde el mismo DataFrame que se exporta (evita desvíos con el frontend)
    for col in COLUMNAS_A_AJUSTAR:
        resultado[col] = float(pd.to_numeric(df_ajustado[col], errors="coerce").fillna(0).sum())

    totales_por_mes: dict[int, dict[str, float]] = {
        m: {c: 0.0 for c in COLUMNAS_A_AJUSTAR} for m in range(1, 13)
    }
    for pos in range(len(df)):
        m = mes_fila.iloc[pos]
        if pd.isna(m):
            continue
        mi = int(m)
        if mi < 1 or mi > 12:
            continue
        for col in COLUMNAS_A_AJUSTAR:
            celda = pd.to_numeric(df_ajustado[col].iloc[pos], errors="coerce")
            totales_por_mes[mi][col] += 0.0 if pd.isna(celda) else float(celda)

    return df_ajustado, resultado, totales_por_mes


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
        df_ajustado.to_excel(ruta_salida, index=False)

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
