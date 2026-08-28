"""
Lógica pura (sin Streamlit) del dashboard de inventario de Amparo.

Separada del archivo de la app para poder testearla con datos reales sin
tener que levantar un servidor de Streamlit. `amparo_dashboard.py` importa
todo esto y le agrega la interfaz.

Formato de archivo esperado (el real de Toni, Base_articulos.xlsx): un
único libro de Excel con dos solapas —

- "Base Articulos": el catálogo. Columnas Rubro, Codigo, Descripción, UxB,
  StkInicial, Ingresos, Egresos, StkActual, Familia, Subfamilia, Marca,
  Proveedor. StkInicial/Ingresos/Egresos ya vienen automatizadas con
  fórmulas dentro del propio Excel (ver add_automation.py) — acá no se
  recalculan, solo se leen sus valores. StkActual es la que vale como stock.
- "Ingresos y Egresos semanal": la carga semanal de movimientos (Codigo,
  EAN, Descripcion, Ingresos, Egresos) que alimenta las fórmulas de arriba.

Pedido explícito de Toni: en el dashboard no tienen que aparecer StkInicial/
Ingresos/Egresos (columnas E, F, G de "Base Articulos") — solo StkActual
(columna H) como stock, más todas las demás columnas, en el orden en que
están en el Excel.
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd

# ---------------------------------------------------------------------------
# Paleta validada (dataviz skill) — usada solo para codificar datos en los
# gráficos. Los colores de marca de Amparo (header) viven en el archivo de la
# app, separados a propósito.
# ---------------------------------------------------------------------------
PALETTE = {
    "blue": "#2a78d6",
    "orange": "#eb6834",
    "good": "#0ca30c",
    "warning": "#fab219",
    "critical": "#d03b3b",
    "ink": "#0b0b0b",
    "ink_secondary": "#52514e",
    "muted": "#898781",
    "grid": "#e1e0d9",
    "surface": "#fcfcfb",
}

UMBRAL_BAJO_DEFAULT = 10.0

BASE_SHEET_NAME = "Base Articulos"
MOV_SHEET_NAME = "Ingresos y Egresos semanal"

# Columnas que se muestran en el dashboard, EN ESTE ORDEN (pedido explícito
# de Toni: todas las columnas de "Base Articulos" salvo StkInicial/Ingresos/
# Egresos — E, F, G — que quedan afuera; StkActual es la que vale como
# Stock). "estado" es un indicador calculado que se agrega al final, no una
# columna del Excel.
DISPLAY_COLUMNS = [
    "rubro", "codigo", "descripcion", "uxb", "stock_actual",
    "familia", "subfamilia", "marca", "proveedor",
]
DISPLAY_LABELS = {
    "rubro": "Rubro",
    "codigo": "Código",
    "descripcion": "Descripción",
    "uxb": "UxB",
    "stock_actual": "Stock",
    "familia": "Familia",
    "subfamilia": "Subfamilia",
    "marca": "Marca",
    "proveedor": "Proveedor",
    "estado": "Estado",
}

BASE_COLUMN_ALIASES: dict[str, list[str]] = {
    "rubro": ["rubro"],
    "codigo": ["codigo", "cod articulo", "sku"],
    "descripcion": ["descripcion"],
    "uxb": ["uxb", "u x b", "unidades x bulto"],
    "stock_inicial": ["stkinicial", "stock inicial"],
    "ingresos": ["ingresos"],
    "egresos": ["egresos"],
    "stock_actual": ["stkactual", "stock actual", "stkact", "stock act"],
    "familia": ["familia"],
    "subfamilia": ["subfamilia", "sub familia"],
    "marca": ["marca"],
    "proveedor": ["proveedor"],
}
REQUIRED_BASE_MIN = {"rubro", "codigo", "descripcion"}
# Columnas de "Base Articulos" que el dashboard tiene explícitamente
# prohibido mostrar (columnas E, F, G del Excel real).
HIDDEN_BASE_COLUMNS = {"stock_inicial", "ingresos", "egresos"}

MOV_COLUMN_ALIASES: dict[str, list[str]] = {
    "codigo": ["codigo"],
    "ean": ["ean"],
    "descripcion": ["descripcion"],
    "ingresos": ["ingresos"],
    "egresos": ["egresos"],
}
REQUIRED_MOV_MIN = {"codigo"}

STATUS_LABEL = {
    "ok": "✓ OK",
    "low": "▲ Bajo",
    "zero": "✕ Sin stock",
    "negative": "✕ Negativo",
}
STATUS_COLOR = {
    "ok": PALETTE["good"],
    "low": PALETTE["warning"],
    "zero": PALETTE["critical"],
    "negative": PALETTE["critical"],
}


def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def _norm_header(s: str) -> str:
    s = _strip_accents(str(s)).lower().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


def clean(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


def parse_number(v) -> float | None:
    """Parsea números en formato AR ('1.234,56') o simple ('1234.56')."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = clean(v).replace(" ", "")
    if s == "":
        return None
    s = re.sub(r"\.(?=\d{3}(?:\D|$))", "", s)
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def map_columns(df: pd.DataFrame, aliases: dict[str, list[str]]) -> dict[str, str]:
    normalized = {_norm_header(c): c for c in df.columns}
    mapping: dict[str, str] = {}
    for logical, alias_list in aliases.items():
        for alias in alias_list:
            if alias in normalized:
                mapping[logical] = normalized[alias]
                break
    return mapping


def _find_sheet(sheet_names: list[str], target: str) -> str | None:
    """Encuentra la solapa por nombre, tolerando espacios extra/mayúsculas
    (el resto del archivo respeta el nombre exacto que usa Toni)."""
    norm_target = _norm_header(target)
    for name in sheet_names:
        if _norm_header(name) == norm_target:
            return name
    return None


def parse_base_sheet(raw: pd.DataFrame) -> pd.DataFrame:
    raw = raw.copy()
    raw.columns = [clean(c) for c in raw.columns]
    mapping = map_columns(raw, BASE_COLUMN_ALIASES)

    missing = REQUIRED_BASE_MIN - mapping.keys()
    if missing:
        raise ValueError(
            "No encontré estas columnas en la solapa '" + BASE_SHEET_NAME + "': "
            + ", ".join(sorted(missing))
        )

    out = pd.DataFrame()
    out["rubro"] = raw[mapping["rubro"]].map(clean).replace("", "Sin rubro")
    out["codigo"] = raw[mapping["codigo"]].map(clean)
    out["descripcion"] = raw[mapping["descripcion"]].map(clean)
    out["uxb"] = raw[mapping["uxb"]].map(clean) if "uxb" in mapping else ""
    for opt_col in ("familia", "subfamilia", "marca", "proveedor"):
        out[opt_col] = (
            raw[mapping[opt_col]].map(clean).replace("", "Sin dato") if opt_col in mapping else "Sin dato"
        )

    stk_inicial = raw[mapping["stock_inicial"]].map(parse_number) if "stock_inicial" in mapping else pd.Series([None] * len(raw))
    ingresos = raw[mapping["ingresos"]].map(parse_number) if "ingresos" in mapping else pd.Series([None] * len(raw))
    egresos = raw[mapping["egresos"]].map(parse_number) if "egresos" in mapping else pd.Series([None] * len(raw))

    # StkActual (columna H) es la que vale. Si vino vacía -- por ejemplo,
    # el archivo se editó sin recalcular fórmulas -- se reconstruye con la
    # misma fórmula que ya trae el Excel (E + F - G), fila por fila.
    fallback = stk_inicial.fillna(0) + ingresos.fillna(0) - egresos.fillna(0)
    if "stock_actual" in mapping:
        stock_actual = raw[mapping["stock_actual"]].map(parse_number)
        stock_actual = stock_actual.where(stock_actual.notna(), fallback)
    else:
        stock_actual = fallback
    out["stock_actual"] = stock_actual

    total_antes = len(out)
    out = out[out["codigo"] != ""].reset_index(drop=True)
    # Metadato (no una columna) para que la UI pueda avisar si el reporte
    # traía filas sin Código -> típicamente separadores/subtotales del
    # sistema de gestión, no productos reales.
    out.attrs["filas_descartadas"] = total_antes - len(out)
    # Metadato de diagnóstico: si StkInicial (E) está vacío/0 para casi todo
    # el catálogo, StkActual (H) es solo el neto de Ingresos-Egresos desde
    # que se empezó a cargar movimientos, no el stock físico real -> la UI
    # debe avisarlo (no se puede notar mirando solo H, que es lo único que
    # se muestra).
    if len(stk_inicial) > 0:
        frac_en_cero = float((stk_inicial.fillna(0) == 0).mean())
    else:
        frac_en_cero = 0.0
    out.attrs["stock_inicial_sin_cargar"] = frac_en_cero >= 0.8
    return out


def parse_mov_sheet(raw: pd.DataFrame) -> pd.DataFrame:
    raw = raw.copy()
    raw.columns = [clean(c) for c in raw.columns]
    mapping = map_columns(raw, MOV_COLUMN_ALIASES)

    missing = REQUIRED_MOV_MIN - mapping.keys()
    if missing:
        raise ValueError(
            "No encontré estas columnas en la solapa '" + MOV_SHEET_NAME + "': "
            + ", ".join(sorted(missing))
        )

    out = pd.DataFrame()
    out["codigo"] = raw[mapping["codigo"]].map(clean)
    out["descripcion"] = raw[mapping["descripcion"]].map(clean) if "descripcion" in mapping else ""
    out["ingresos"] = raw[mapping["ingresos"]].map(parse_number) if "ingresos" in mapping else 0.0
    out["egresos"] = raw[mapping["egresos"]].map(parse_number) if "egresos" in mapping else 0.0

    # Igual que en Base Articulos: las filas sin Código son separadores o
    # subtotales del reporte ("TOTAL FAMILIA:", "TOTAL MOVIMIENTOS:", etc.),
    # no movimientos reales -> se descartan.
    out = out[out["codigo"] != ""].reset_index(drop=True)
    return out


def parse_workbook_bytes(name: str, data: bytes) -> tuple[pd.DataFrame, pd.DataFrame | None]:
    """Lee el libro de Toni y devuelve (base_articulos, movimientos).

    `movimientos` puede ser None si el archivo no trae esa solapa (se
    permite igual: el dashboard es útil solo con el catálogo, el aviso de
    códigos huérfanos simplemente no se muestra en ese caso).
    """
    import io

    xls = pd.ExcelFile(io.BytesIO(data))
    base_name = _find_sheet(xls.sheet_names, BASE_SHEET_NAME)
    if base_name is None:
        raise ValueError(
            f"No encontré la solapa '{BASE_SHEET_NAME}' en {name}. "
            f"Solapas encontradas: {', '.join(xls.sheet_names)}"
        )
    raw_base = pd.read_excel(xls, sheet_name=base_name, dtype=str)
    base_df = parse_base_sheet(raw_base)
    base_df.attrs["archivo_origen"] = name

    mov_name = _find_sheet(xls.sheet_names, MOV_SHEET_NAME)
    mov_df = None
    if mov_name is not None:
        raw_mov = pd.read_excel(xls, sheet_name=mov_name, dtype=str)
        mov_df = parse_mov_sheet(raw_mov)

    return base_df, mov_df


def find_orphan_codes(base_df: pd.DataFrame, mov_df: pd.DataFrame | None) -> pd.DataFrame:
    """Códigos que registran movimientos en la solapa semanal pero no
    existen en el catálogo (Base Articulos) -> ese Ingreso/Egreso no se le
    sumó/restó a ningún stock. Ver 'Códigos a revisar' en el Excel."""
    if mov_df is None or mov_df.empty:
        return pd.DataFrame(columns=["codigo", "descripcion", "ingresos", "egresos"])
    codigos_base = set(base_df["codigo"])
    orphans = mov_df[~mov_df["codigo"].isin(codigos_base)].copy()
    return orphans.reset_index(drop=True)


def stock_status(row, umbral: float) -> str:
    v = row["stock_actual"]
    if v is None or pd.isna(v):
        return "zero"
    if v < 0:
        return "negative"
    if v == 0:
        return "zero"
    if v <= umbral:
        return "low"
    return "ok"


def compute_kpis(df: pd.DataFrame, umbral: float = UMBRAL_BAJO_DEFAULT) -> dict:
    df = df.copy()
    df["estado"] = df.apply(lambda r: stock_status(r, umbral), axis=1)
    return {
        "productos": len(df),
        "stock_acumulado": df["stock_actual"].fillna(0).sum(),
        "stock_bajo": int((df["estado"] == "low").sum()),
        "sin_stock_negativo": int(df["estado"].isin(["zero", "negative"]).sum()),
        "rubros": df["rubro"].nunique(),
        "familias": df.loc[df["familia"] != "Sin dato", "familia"].nunique(),
        "proveedores": df.loc[df["proveedor"] != "Sin dato", "proveedor"].nunique(),
    }
