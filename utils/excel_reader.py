"""
Lector de Excel: parsea el Registro de Compras y Ventas en el formato de Pedro.
Compatible con el formato del sistema contable que exporta (encabezado en fila 10).
"""

import pandas as pd
from datetime import datetime, date
import re
from typing import Optional

# Columnas del Registro de Compras (en orden del Excel de Pedro)
COLS_COMPRAS = [
    "LLAVE", "NUM CORRELATIVO", "FECHA DE EMISIÓN", "FECHA DE VENCIMIENTO",
    "TIPO DOC", "SERIE", "AÑO DUA", "NRO COMPROBANTE",
    "TIPO DOC PROVEEDOR", "NRO DOC PROVEEDOR", "RAZÓN SOCIAL", "DESCRIPCIÓN",
    "ADQ GRA DES OPE GRA - BASE", "ADQ GRA DES OPE GRA - IGV",
    "ADQ GRA DES OPE - BASE", "ADQ GRA DES OPE - IGV",
    "ADQ GRA DES OPE NO GRA - BASE", "ADQ GRA DES OPE NO GRA - IGV",
    "VALOR NO GRABADO", "ISC", "ICBPER", "OTROS TRIBUTOS",
    "IMPORTE TOTAL MN", "IMPORTE TOTAL ME",
    "NRO DE COMPROB NO DOMICILIADO", "NRO CONSTANCIA DETRACCIÓN",
    "FECHA DETRACCIÓN", "TIPO DE CAMBIO",
    "FECHA DOC REFERENCIA", "TIPO DOC REFERENCIA",
    "SERIE DOC REFERENCIA", "NRO DOC REFERENCIA",
    "DESCRIPCIÓN TIPO DOC", "CUENTA CONTABLE", "NOMBRE CUENTA CONTABLE",
    "CENTRO DE COSTO", "NOMBRE CENTRO DE COSTO",
    "CUENTA DESTINO DEBE", "NOMBRE CUENTA DESTINO DEBE",
    "CUENTA DESTINO HABER", "NOMBRE CUENTA DESTINO HABER",
    "SUB CENTRO DE COSTO", "NOMBRE SUB CENTRO DE COSTO",
    "PROYECTO", "MONEDA", "ORIGENPK",
]

# Columnas del Registro de Ventas
COLS_VENTAS = [
    "LLAVE", "NUM CORRELATIVO", "FECHA DE EMISIÓN", "FECHA DE VENCIMIENTO",
    "TIPO DOC", "SERIE", "NRO COMPROBANTE",
    "TIPO DOC CLIENTE", "NRO DOC CLIENTE", "RAZÓN SOCIAL CLIENTE", "DESCRIPCIÓN",
    "BASE GRAVADA", "IGV", "EXONERADO", "INAFECTO", "ISC", "ICBPER",
    "OTROS TRIBUTOS", "IMPORTE TOTAL MN", "IMPORTE TOTAL ME",
    "TIPO DE CAMBIO", "FECHA DOC REFERENCIA", "TIPO DOC REFERENCIA",
    "SERIE DOC REFERENCIA", "NRO DOC REFERENCIA",
    "MONEDA", "CUENTA CONTABLE", "NOMBRE CUENTA CONTABLE",
]


def _excel_date_to_date(val) -> Optional[date]:
    """Convierte número de serie Excel o string de fecha a date."""
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        try:
            # Excel serial date (Windows epoch)
            return (datetime(1899, 12, 30) + pd.Timedelta(days=int(val))).date()
        except Exception:
            return None
    if isinstance(val, (datetime, date)):
        return val.date() if isinstance(val, datetime) else val
    if isinstance(val, str):
        for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"]:
            try:
                return datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _safe_float(val) -> float:
    """Convierte a float de forma segura."""
    if val is None or val == "" or (isinstance(val, float) and val != val):  # NaN
        return 0.0
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def _safe_str(val) -> str:
    """Convierte a str limpio."""
    if val is None:
        return ""
    s = str(val).strip()
    # Remover ".0" de números enteros mal formateados
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]
    return s


def detectar_fila_header(df_raw: pd.DataFrame, keywords: list = None) -> int:
    """
    Busca la fila que contiene los encabezados del Registro.
    Busca palabras clave como 'NUM CORRELATIVO', 'TIPO DOC', 'FECHA DE EMISIÓN'.
    Retorna el índice de la fila (0-based).
    """
    if keywords is None:
        keywords = ["NUM CORRELATIVO", "TIPO DOC", "RAZÓN SOCIAL", "TIPO DE CAMBIO"]

    for i, row in df_raw.iterrows():
        row_vals = [str(v).upper().strip() for v in row.values]
        matches = sum(1 for kw in keywords if any(kw in v for v in row_vals))
        if matches >= 3:
            return i
    return 0


def leer_registro_compras(filepath: str, nombre_hoja: str = None) -> pd.DataFrame:
    """
    Lee el Registro de Compras desde un archivo Excel.
    Detecta automáticamente la hoja correcta y la fila de encabezados.

    Retorna DataFrame limpio con las columnas estándar.
    """
    # Leer todas las hojas disponibles
    xl = pd.ExcelFile(filepath)
    hojas = xl.sheet_names

    # Buscar hoja que parezca Registro de Compras
    if nombre_hoja:
        hoja = nombre_hoja
    else:
        candidatos = [h for h in hojas if "COMPRA" in h.upper() or "II-1" in h.upper()]
        hoja = candidatos[0] if candidatos else hojas[0]

    # Leer raw para detectar encabezados
    df_raw = pd.read_excel(filepath, sheet_name=hoja, header=None, dtype=str)
    fila_header = detectar_fila_header(df_raw)

    # Leer con el header correcto
    df = pd.read_excel(filepath, sheet_name=hoja, header=fila_header, dtype=str)
    df = df.dropna(how="all")

    # Limpiar columnas duplicadas
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]

    # Filtrar filas que parecen datos (tienen NUM CORRELATIVO numérico)
    if "NUM CORRELATIVO" in df.columns:
        df = df[df["NUM CORRELATIVO"].apply(lambda x: str(x).strip().isdigit())]

    # Convertir columnas numéricas
    cols_numericas = [
        "ADQ GRA DES OPE GRA - BASE", "ADQ GRA DES OPE GRA - IGV",
        "ADQ GRA DES OPE - BASE", "ADQ GRA DES OPE - IGV",
        "ADQ GRA DES OPE NO GRA - BASE", "ADQ GRA DES OPE NO GRA - IGV",
        "VALOR NO GRABADO", "ISC", "ICBPER", "OTROS TRIBUTOS",
        "IMPORTE TOTAL MN", "IMPORTE TOTAL ME", "TIPO DE CAMBIO",
    ]
    for col in cols_numericas:
        if col in df.columns:
            df[col] = df[col].apply(_safe_float)

    # Convertir fechas
    cols_fecha = ["FECHA DE EMISIÓN", "FECHA DE VENCIMIENTO", "FECHA DETRACCIÓN", "FECHA DOC REFERENCIA"]
    for col in cols_fecha:
        if col in df.columns:
            df[col] = df[col].apply(_excel_date_to_date)

    # Limpiar strings
    cols_str = ["TIPO DOC", "SERIE", "NRO COMPROBANTE", "NRO DOC PROVEEDOR",
                "RAZÓN SOCIAL", "DESCRIPCIÓN", "MONEDA", "CUENTA CONTABLE",
                "NRO CONSTANCIA DETRACCIÓN"]
    for col in cols_str:
        if col in df.columns:
            df[col] = df[col].apply(_safe_str)

    # Rellenar TIPO DOC con "01" si está vacío (factura por defecto)
    if "TIPO DOC" in df.columns:
        df["TIPO DOC"] = df["TIPO DOC"].apply(lambda x: x.zfill(2) if x and x.isdigit() else x or "01")

    # Rellenar MONEDA
    if "MONEDA" in df.columns:
        df["MONEDA"] = df["MONEDA"].fillna("PEN").replace("", "PEN")

    df = df.reset_index(drop=True)
    return df, hoja


def leer_registro_ventas(filepath: str, nombre_hoja: str = None) -> pd.DataFrame:
    """
    Lee el Registro de Ventas desde un archivo Excel.
    """
    xl = pd.ExcelFile(filepath)
    hojas = xl.sheet_names

    if nombre_hoja:
        hoja = nombre_hoja
    else:
        candidatos = [h for h in hojas if "VENTA" in h.upper() or "II-2" in h.upper()]
        hoja = candidatos[0] if candidatos else hojas[0]

    df_raw = pd.read_excel(filepath, sheet_name=hoja, header=None, dtype=str)
    fila_header = detectar_fila_header(
        df_raw,
        keywords=["NUM CORRELATIVO", "TIPO DOC", "BASE GRAVADA", "MONEDA"],
    )

    df = pd.read_excel(filepath, sheet_name=hoja, header=fila_header, dtype=str)
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]

    if "NUM CORRELATIVO" in df.columns:
        df = df[df["NUM CORRELATIVO"].apply(lambda x: str(x).strip().isdigit())]

    # Numéricas ventas
    cols_numericas = ["BASE GRAVADA", "IGV", "EXONERADO", "INAFECTO",
                      "ISC", "ICBPER", "OTROS TRIBUTOS",
                      "IMPORTE TOTAL MN", "IMPORTE TOTAL ME", "TIPO DE CAMBIO"]
    for col in cols_numericas:
        if col in df.columns:
            df[col] = df[col].apply(_safe_float)

    cols_fecha = ["FECHA DE EMISIÓN", "FECHA DE VENCIMIENTO"]
    for col in cols_fecha:
        if col in df.columns:
            df[col] = df[col].apply(_excel_date_to_date)

    df = df.reset_index(drop=True)
    return df, hoja


def obtener_hojas_disponibles(filepath: str) -> list:
    """Retorna lista de hojas en el Excel."""
    xl = pd.ExcelFile(filepath)
    return xl.sheet_names
