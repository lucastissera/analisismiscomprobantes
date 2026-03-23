"""
Lee un archivo Excel (.xlsx), ajusta signos según Tipo y suma columnas indicadas.
Estructura del archivo: fila 1 = encabezado general, fila 2 = encabezados de columnas,
datos desde la fila 3.
Las filas con Tipo = nota de crédito (por código numérico) se consideran en negativo.

Uso:
  python sumar_imp_total.py <ruta_al_archivo.xlsx> [hoja] [archivo_salida.xlsx]
"""

import sys
from pathlib import Path
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

# Códigos numéricos de la columna "Tipo" que se consideran nota de crédito (suma en negativo)
# Se matchea por número (sin ceros a la izquierda): "003", "3" y "03" son el mismo código
CODIGOS_NOTA_CREDITO = {
    3, 8, 13, 21, 38, 43, 44, 48, 53,
    110, 112, 113, 114, 203, 206, 208, 211, 213,
}


def limpiar_argumento_ruta(valor: str) -> str:
    """Normaliza saltos de línea/tabulaciones accidentales en argumentos de ruta."""
    return valor.replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()


def leer_tabla(entrada, hoja: str | int = 0, nombre_archivo: str | None = None) -> pd.DataFrame:
    """
    Lee un .xlsx o .csv con formato:
    - fila 1: encabezado general
    - fila 2: encabezados de columnas
    - fila 3+: datos
    """
    nombre = (nombre_archivo or str(entrada)).lower()
    if nombre.endswith(".csv"):
        df = pd.read_csv(entrada, header=1, sep=None, engine="python")
    else:
        df = pd.read_excel(entrada, sheet_name=hoja, header=1)

    df.columns = df.columns.astype(str).str.strip()
    return df


def procesar_archivo(
    ruta_excel: str, hoja: str | int = 0, nombre_archivo: str | None = None
) -> tuple[pd.DataFrame, dict[str, float]]:
    """
    Lee un archivo Excel y devuelve la sumatoria de las columnas indicadas.
    Fila 1 = encabezado general, fila 2 = encabezados de columnas, datos desde fila 3.
    Las filas con Tipo = nota de crédito se suman en valor negativo.

    Args:
        ruta_excel: Ruta al archivo .xlsx
        hoja: Nombre o índice de la hoja (0 por defecto)

    Returns:
        Tuple con:
        - DataFrame ajustado (columnas numéricas con signo aplicado según Tipo)
        - Diccionario con el nombre de cada columna y su suma.
    """
    # header=1: la fila 2 del archivo (índice 1) tiene los nombres de columnas; datos desde fila 3
    df = leer_tabla(ruta_excel, hoja=hoja, nombre_archivo=nombre_archivo)

    # Comprobar que existan todas las columnas necesarias
    columnas_requeridas = COLUMNAS_A_AJUSTAR + ["Tipo", "Tipo Cambio"]
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        nombres = ", ".join(df.columns.astype(str))
        raise ValueError(
            f"No se encontraron las columnas: {faltantes}. "
            f"Columnas en el archivo: {nombres}"
        )

    # Signo: -1 si es nota de crédito, +1 si no (código como número, sin ceros a la izquierda)
    tipo_str = df["Tipo"].astype(str).str.strip()
    codigo_str = tipo_str.str.split(" - ", n=1).str[0].str.strip()  # ej. "003 - NOTA..." -> "003"
    codigo_num = pd.to_numeric(codigo_str, errors="coerce")  # "003" y "3" -> 3
    es_nota_credito = codigo_num.isin(CODIGOS_NOTA_CREDITO)
    signo = 1 - 2 * es_nota_credito.astype(int)  # True -> -1, False -> 1

    # Factor de conversión por fila
    tipo_cambio = pd.to_numeric(df["Tipo Cambio"], errors="coerce")

    # Ajustar signos y tipo de cambio en el DataFrame de salida, luego acumular totales
    df_ajustado = df.copy()
    resultado = {}
    for col in COLUMNAS_A_AJUSTAR:
        valores = pd.to_numeric(df[col], errors="coerce")
        valores_ajustados = valores * signo * tipo_cambio
        df_ajustado[col] = valores_ajustados
        resultado[col] = float(valores_ajustados.sum())

    return df_ajustado, resultado


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
        df_ajustado, totales = procesar_archivo(ruta, hoja, nombre_archivo=ruta)
        ruta_salida = construir_ruta_salida(ruta, salida_arg)
        df_ajustado.to_excel(ruta_salida, index=False)

        print("Sumas (desde fila 3):")
        for col, total in totales.items():
            print(f"  {col}: {total:,.2f}")
        print("  ---")
        print(f"  Suma total: {sum(totales.values()):,.2f}")
        print(f"Excel ajustado generado en: {ruta_salida}")
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo '{ruta}'", file=sys.stderr)
        sys.exit(2)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(3)


if __name__ == "__main__":
    main()
