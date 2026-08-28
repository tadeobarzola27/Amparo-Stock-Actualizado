"""
Smoke test de amparo_logic.py contra el archivo real de Toni,
Base_articulos.xlsx (3621 productos en 'Base Articulos', 668 filas de carga
en 'Ingresos y Egresos semanal').

Correr con:  python3 test_amparo_logic.py
"""

from pathlib import Path

import pandas as pd

import amparo_logic as L

HERE = Path(__file__).parent
REAL = HERE / "Base_articulos.xlsx"

failures = []


def check(label, condition, detail=""):
    status = "OK " if condition else "FAIL"
    print(f"[{status}] {label} {detail}")
    if not condition:
        failures.append(label)


# --- parse_number ---------------------------------------------------------
check("parse_number: entero simple", L.parse_number("123") == 123.0)
check("parse_number: miles con punto", L.parse_number("1.234") == 1234.0)
check("parse_number: miles+decimales AR", L.parse_number("1.234,56") == 1234.56)
check("parse_number: numero python", L.parse_number(546455) == 546455.0)
check("parse_number: vacio -> None", L.parse_number("") is None)
check("parse_number: None -> None", L.parse_number(None) is None)
check("parse_number: negativo", L.parse_number("-12") == -12.0)

# --- map_columns ------------------------------------------------------------
df_headers = pd.DataFrame(columns=[
    "Rubro", "Codigo", "Descripción", "UxB", "StkInicial", "Ingresos", "Egresos",
    "StkActual", "Familia", "Subfamilia", "Marca", "Proveedor",
])
mapping = L.map_columns(df_headers, L.BASE_COLUMN_ALIASES)
check(
    "map_columns (Base Articulos): reconoce las 12 columnas reales",
    set(mapping.keys()) == set(L.BASE_COLUMN_ALIASES.keys()),
    detail=str(mapping),
)

df_mov_headers = pd.DataFrame(columns=[" Codigo", "         EAN", "Descripcion", "Ingresos", " Egresos"])
mapping_mov = L.map_columns(df_mov_headers, L.MOV_COLUMN_ALIASES)
check(
    "map_columns (movimientos): tolera el padding real del export",
    set(mapping_mov.keys()) == {"codigo", "ean", "descripcion", "ingresos", "egresos"},
    detail=str(mapping_mov),
)

# --- parse_workbook_bytes contra el archivo real ---------------------------
assert REAL.exists(), f"No se encontró {REAL}"
base_df, mov_df = L.parse_workbook_bytes(REAL.name, REAL.read_bytes())

# valores de referencia calculados directamente con pandas sobre el mismo
# archivo (ya recalculado con recalc.py), para cruzar contra la lógica
raw_base = pd.read_excel(REAL, sheet_name=L.BASE_SHEET_NAME, dtype=str)
raw_base.columns = [str(c).strip() for c in raw_base.columns]
h = pd.to_numeric(raw_base["StkActual"], errors="coerce")

check("Base Articulos: 3621 productos, ninguno descartado", len(base_df) == 3621, detail=f"{len(base_df)} filas")
check("Base Articulos: no se descartó ninguna fila", base_df.attrs.get("filas_descartadas") == 0)
check(
    "Base Articulos: columnas de display presentes en orden",
    list(base_df.columns[: len(L.DISPLAY_COLUMNS)]) == L.DISPLAY_COLUMNS
    or set(L.DISPLAY_COLUMNS) <= set(base_df.columns),
    detail=str(list(base_df.columns)),
)
for hidden in L.HIDDEN_BASE_COLUMNS:
    check(f"Base Articulos: '{hidden}' NO se expone (pedido explícito de Toni)", hidden not in base_df.columns)

check(
    "Base Articulos: suma de Stock (StkActual/H) coincide con el Excel",
    abs(base_df["stock_actual"].sum() - h.sum()) < 0.01,
    detail=f"logic={base_df['stock_actual'].sum()} excel={h.sum()}",
)
check("Base Articulos: cuenta de negativos", int((base_df["stock_actual"] < 0).sum()) == 361)
check("Base Articulos: cuenta de ceros", int((base_df["stock_actual"] == 0).sum()) == 3159)
check(
    "Base Articulos: cuenta de stock bajo (umbral default)",
    int(((base_df["stock_actual"] > 0) & (base_df["stock_actual"] <= L.UMBRAL_BAJO_DEFAULT)).sum()) == 33,
)
check("Base Articulos: Rubros distintos", base_df["rubro"].nunique() == 11)
check("Base Articulos: Familias distintas", base_df["familia"].nunique() == 81)
check("Base Articulos: Proveedores distintos", base_df["proveedor"].nunique() == 240)
check("Base Articulos: Marcas distintas", base_df["marca"].nunique() == 698)

check(
    "Diagnóstico: detecta que StkInicial (E) no está cargado (100% en 0 en este archivo)",
    base_df.attrs.get("stock_inicial_sin_cargar") is True,
)

# --- Ingresos y Egresos semanal --------------------------------------------
check("Movimientos: se leyó la solapa", mov_df is not None)
check("Movimientos: 535 filas válidas (se descartan separadores/subtotales)", len(mov_df) == 535, detail=f"{len(mov_df)} filas")
check("Movimientos: ingresos total coincide con Base Articulos (F)", abs(mov_df["ingresos"].sum() - 26925.48) < 0.01)
check("Movimientos: egresos total coincide con lo cargado", abs(mov_df["egresos"].sum() - 38433.78) < 0.5,
      detail=f"{mov_df['egresos'].sum()}")

# --- find_orphan_codes -------------------------------------------------------
orphans = L.find_orphan_codes(base_df, mov_df)
check("Huérfanos: exactamente 5 códigos (hallazgo confirmado en el archivo real)", len(orphans) == 5,
      detail=str(sorted(orphans["codigo"].tolist())))
check(
    "Huérfanos: son los 5 códigos esperados",
    set(orphans["codigo"]) == {"3072", "393", "6690862", "6691866", "604"},
    detail=str(sorted(orphans["codigo"].tolist())),
)
check("Huérfanos: ingresos sin capturar == 0", abs(orphans["ingresos"].sum() - 0) < 0.01)
check("Huérfanos: egresos sin capturar == 4589 (la brecha detectada)", abs(orphans["egresos"].sum() - 4589) < 0.5,
      detail=str(orphans["egresos"].sum()))
check(
    "find_orphan_codes: con mov_df=None no rompe, devuelve vacío",
    L.find_orphan_codes(base_df, None).empty,
)

# --- stock_status / compute_kpis ------------------------------------------
base_df = base_df.copy()
base_df["estado"] = base_df.apply(lambda r: L.stock_status(r, L.UMBRAL_BAJO_DEFAULT), axis=1)
check("stock_status: negativos", int((base_df["estado"] == "negative").sum()) == 361)
check("stock_status: sin stock (cero)", int((base_df["estado"] == "zero").sum()) == 3159)
check("stock_status: bajo", int((base_df["estado"] == "low").sum()) == 33)

kpis = L.compute_kpis(base_df.drop(columns="estado"))
check("compute_kpis: productos == filas", kpis["productos"] == 3621)
check("compute_kpis: stock_acumulado coincide", abs(kpis["stock_acumulado"] - h.sum()) < 0.01)
check("compute_kpis: rubros", kpis["rubros"] == 11)
check("compute_kpis: familias", kpis["familias"] == 81)
check("compute_kpis: proveedores", kpis["proveedores"] == 240)
check(
    "compute_kpis: no expone ingresos_total/egresos_total (E/F/G fuera del dashboard)",
    "ingresos_total" not in kpis and "egresos_total" not in kpis,
)

# --- columnas faltantes -> error legible -----------------------------------
try:
    L.parse_base_sheet(pd.DataFrame({"Descripcion": ["Foo"]}))
    check("Base Articulos, columnas faltantes: debía lanzar ValueError", False)
except ValueError as e:
    check("Base Articulos, columnas faltantes: ValueError con mensaje claro", "rubro" in str(e) or "codigo" in str(e), detail=str(e))

try:
    L.parse_mov_sheet(pd.DataFrame({"Descripcion": ["Foo"]}))
    check("Movimientos, columnas faltantes: debía lanzar ValueError", False)
except ValueError as e:
    check("Movimientos, columnas faltantes: ValueError con mensaje claro", "codigo" in str(e), detail=str(e))

import io as _io
buf = _io.BytesIO()
with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
    pd.DataFrame({"Rubro": ["A"], "Codigo": ["1"], "Descripcion": ["Prod"]}).to_excel(
        writer, sheet_name="Otra hoja", index=False
    )
try:
    L.parse_workbook_bytes("sin_solapa.xlsx", buf.getvalue())
    check("Archivo sin solapa 'Base Articulos': debía lanzar ValueError", False)
except ValueError as e:
    check("Archivo sin solapa 'Base Articulos': ValueError con mensaje claro", "Base Articulos" in str(e), detail=str(e))

print()
if failures:
    print(f"❌ {len(failures)} check(s) fallaron: {failures}")
    raise SystemExit(1)
else:
    print(f"✅ Todos los checks pasaron contra el archivo real (3621 productos, 535 movimientos válidos).")
