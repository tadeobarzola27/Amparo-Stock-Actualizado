"""
Amparo · Inventario semanal — dashboard en Streamlit
======================================================

Se alimenta directamente de Base_articulos.xlsx: el mismo archivo que ya
tenés, con la solapa "Base Articulos" (catálogo + stock) y "Ingresos y
Egresos semanal" (carga semanal de movimientos). El cruce entre las dos
solapas — descontar/sumar cada movimiento del stock del código que
corresponde — ya está automatizado adentro del propio Excel con fórmulas;
esta app solo lee el resultado y lo muestra.

Cómo correrlo
--------------
    pip install streamlit pandas plotly openpyxl xlsxwriter
    streamlit run amparo_dashboard.py

El archivo que subís nunca sale de tu computadora: Streamlit lo procesa
localmente en el proceso que corre en tu máquina.

La lógica de parseo y cálculo vive en amparo_logic.py (sin dependencia de
Streamlit) para poder testearla por separado — ver test_amparo_logic.py.
"""

from __future__ import annotations

import hashlib
import io
import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

from amparo_logic import (
    DISPLAY_LABELS,
    PALETTE,
    STATUS_COLOR,
    STATUS_LABEL,
    UMBRAL_BAJO_DEFAULT,
    compute_kpis,
    find_orphan_codes,
    parse_workbook_bytes,
    stock_status,
)

# Colores de marca "Amparo" (solo para el header/chrome — nunca para
# codificar datos, eso lo hace PALETTE).
BRAND = {
    "navy": "#142d59",
    "blue": "#315fcb",
    "teal": "#118c7e",
}

DEFAULT_DATA_FILE_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1KqELEklzzrZrnL2cZ5VH0NDQdas_lkRe3SdPaoZYG0U/edit?usp=sharing"
)

HISTORY_FILE = Path(__file__).with_name("historial_amparo.csv")
HISTORY_COLUMNS = ["fecha", "productos", "stock_acumulado", "stock_bajo", "sin_stock_neg", "rubros"]

GROUP_DIMENSIONS = {
    "Rubro": "rubro",
    "Familia": "familia",
    "Subfamilia": "subfamilia",
    "Marca": "marca",
    "Proveedor": "proveedor",
}


@st.cache_data(show_spinner=False)
def load_workbook_bytes(name: str, data: bytes) -> tuple[pd.DataFrame, pd.DataFrame | None]:
    return parse_workbook_bytes(name, data)


def file_signature(name: str, data: bytes) -> str:
    """Huella del archivo central para no repetir snapshots en un rerun."""
    h = hashlib.sha256()
    h.update(name.encode("utf-8", "ignore"))
    h.update(data)
    return h.hexdigest()


def get_setting(name: str) -> str:
    """Lee una configuración de Streamlit Secrets o del entorno local."""
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = ""
    return str(value or os.getenv(name, "")).strip()


def normalize_data_url(url: str) -> str:
    """Convierte enlaces de Google Sheets o Drive en descargas Excel."""
    if not url:
        return url

    sheet_match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", url)
    if sheet_match:
        sheet_id = sheet_match.group(1)
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    file_match = re.search(r"/file/d/([A-Za-z0-9_-]+)", url)
    id_match = re.search(r"[?&]id=([A-Za-z0-9_-]+)", url)
    file_id = file_match.group(1) if file_match else (id_match.group(1) if id_match else "")
    if file_id and ("drive.google.com" in url or "docs.google.com" in url):
        return f"https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t"
    return url


@st.cache_data(ttl=300, show_spinner=False)
def download_central_workbook(url: str) -> tuple[str, bytes, str]:
    """Descarga el Excel central. La caché se renueva cada cinco minutos."""
    response = requests.get(normalize_data_url(url), timeout=90)
    response.raise_for_status()
    data = response.content
    content_type = response.headers.get("content-type", "").lower()

    if not data.startswith(b"PK") or "text/html" in content_type:
        raise ValueError(
            "El enlace no devolvió un archivo Excel. Verificá que el archivo de Drive "
            "esté compartido como 'Cualquier persona con el enlace: Lector'."
        )

    updated = response.headers.get("last-modified") or datetime.now().strftime("%d/%m/%Y %H:%M")
    return "Base_articulos.xlsx", data, updated


def append_history(kpis: dict) -> None:
    row = {
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "productos": kpis["productos"],
        "stock_acumulado": kpis["stock_acumulado"],
        "stock_bajo": kpis["stock_bajo"],
        "sin_stock_neg": kpis["sin_stock_negativo"],
        "rubros": kpis["rubros"],
    }
    hist = pd.DataFrame([row])[HISTORY_COLUMNS]
    if HISTORY_FILE.exists():
        hist.to_csv(HISTORY_FILE, mode="a", header=False, index=False)
    else:
        hist.to_csv(HISTORY_FILE, mode="w", header=True, index=False)


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Amparo · Inventario semanal",
    page_icon="📦",
    layout="wide",
)

st.markdown(
    f"""
    <style>
    .hero {{
        background: linear-gradient(120deg, {BRAND['navy']}, {BRAND['blue']} 62%, {BRAND['teal']});
        color: #fff; padding: 26px 30px; border-radius: 18px; margin-bottom: 18px;
    }}
    .hero .eyebrow {{ font-size: 11px; font-weight: 800; letter-spacing: .12em;
        text-transform: uppercase; opacity: .75; }}
    .hero h1 {{ margin: 6px 0 4px; font-size: 32px; }}
    .hero p {{ margin: 0; opacity: .85; font-size: 13px; }}
    span.status-badge {{
        display: inline-block; padding: 3px 9px; border-radius: 7px;
        font-weight: 700; font-size: 12.5px; color: #fff;
    }}
    </style>
    <div class="hero">
        <div class="eyebrow">Amparo · Gestión de mercadería</div>
        <h1>Inventario semanal</h1>
        <p>Consulta rápida para pedidos de reposición de las 8 tiendas — información centralizada por la oficina.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

data_file_url = get_setting("DATA_FILE_URL") or DEFAULT_DATA_FILE_URL

refresh_col, source_col = st.columns([1, 4], vertical_alignment="center")
with refresh_col:
    if st.button("🔄 Actualizar datos", use_container_width=True):
        download_central_workbook.clear()
        load_workbook_bytes.clear()
        st.rerun()
with source_col:
    st.caption(
        "Los datos se cargan automáticamente desde el archivo central administrado por la oficina. "
        "Las tiendas no necesitan subir archivos."
    )

if not data_file_url:
    st.error(
        "La fuente central todavía no está configurada. La oficina debe agregar el enlace del "
        "Excel en la variable privada DATA_FILE_URL de Streamlit."
    )
    st.stop()

try:
    with st.spinner("Actualizando inventario desde la fuente central..."):
        workbook_name, workbook_data, source_updated = download_central_workbook(data_file_url)
        base_df, mov_df = load_workbook_bytes(workbook_name, workbook_data)
except (requests.RequestException, ValueError) as e:
    st.error(f"No se pudo actualizar el inventario central. {e}")
    st.stop()

st.success(f"Inventario central disponible · última lectura: {source_updated}", icon="✅")

if base_df.empty:
    st.warning("No se pudo leer ningún producto de 'Base Articulos' en el archivo central.")
    st.stop()

filas_descartadas = base_df.attrs.get("filas_descartadas", 0)
if filas_descartadas:
    st.caption(
        f"ℹ️ Se ignoraron {filas_descartadas} fila(s) sin Código en 'Base Articulos' (separadores o "
        "subtotales del reporte, no productos)."
    )

if base_df.attrs.get("stock_inicial_sin_cargar"):
    st.info(
        "ℹ️ La columna **Stock Inicial** (E) está en 0 para casi todo el catálogo. Mientras no la "
        "cargues, el **Stock** que ves acá es solo el neto de Ingresos menos Egresos desde que "
        "empezaste a cargar movimientos — no el stock físico real. Cargá el stock inicial real "
        "(de tu conteo o de tu sistema de gestión) en la solapa 'Base Articulos' para que el "
        "número quede exacto.",
        icon="ℹ️",
    )

# ---------------------------------------------------------------------------
# Aviso de códigos huérfanos (movimientos sin catálogo)
# ---------------------------------------------------------------------------

orphans = find_orphan_codes(base_df, mov_df)
if not orphans.empty:
    st.warning(
        f"⚠ {len(orphans)} código(s) tuvieron Ingresos o Egresos cargados esta semana pero no están "
        "en tu catálogo (Base Articulos) — ese movimiento no se sumó ni restó a ningún stock.",
        icon="⚠️",
    )
    with st.expander("Ver códigos a revisar"):
        st.dataframe(
            orphans.rename(columns={
                "codigo": "Código", "descripcion": "Descripción",
                "ingresos": "Ingresos", "egresos": "Egresos",
            }),
            hide_index=True,
            width="stretch",
        )
        st.caption(
            "Este mismo detalle está en la solapa 'Códigos a revisar' del Excel. Agregalos como fila "
            "nueva en 'Base Articulos' (con su Rubro, Descripción, etc.) y se van a sumar solos la "
            "próxima vez que se abra el archivo."
        )

# ---------------------------------------------------------------------------
# Filtros
# ---------------------------------------------------------------------------

st.subheader("Filtros")
f1, f2, f3, f4 = st.columns(4)
with f1:
    rubros_sel = st.multiselect("Rubro", sorted(base_df["rubro"].unique(), key=str.lower), default=[])
with f2:
    familias_sel = st.multiselect("Familia", sorted(base_df["familia"].unique(), key=str.lower), default=[])
with f3:
    marcas_sel = st.multiselect("Marca", sorted(base_df["marca"].unique(), key=str.lower), default=[])
with f4:
    proveedores_sel = st.multiselect("Proveedor", sorted(base_df["proveedor"].unique(), key=str.lower), default=[])

g1, g2, g3, g4 = st.columns(4)
with g1:
    texto = st.text_input("Código o descripción", placeholder="Ej. 1637 o queso...")
with g2:
    stock_min = st.number_input("Stock mínimo", value=None, step=1.0, format="%.0f")
with g3:
    stock_max = st.number_input("Stock máximo", value=None, step=1.0, format="%.0f")
with g4:
    umbral_bajo = st.number_input("Umbral stock bajo", value=UMBRAL_BAJO_DEFAULT, step=1.0, format="%.0f")

base_df = base_df.copy()
base_df["estado"] = base_df.apply(lambda r: stock_status(r, umbral_bajo), axis=1)

filtered = base_df
if rubros_sel:
    filtered = filtered[filtered["rubro"].isin(rubros_sel)]
if familias_sel:
    filtered = filtered[filtered["familia"].isin(familias_sel)]
if marcas_sel:
    filtered = filtered[filtered["marca"].isin(marcas_sel)]
if proveedores_sel:
    filtered = filtered[filtered["proveedor"].isin(proveedores_sel)]
if texto:
    t = texto.strip().lower()
    filtered = filtered[
        filtered["codigo"].str.lower().str.contains(t, na=False)
        | filtered["descripcion"].str.lower().str.contains(t, na=False)
    ]
if stock_min is not None:
    filtered = filtered[filtered["stock_actual"].fillna(0) >= stock_min]
if stock_max is not None:
    filtered = filtered[filtered["stock_actual"].fillna(0) <= stock_max]

sig = file_signature(workbook_name, workbook_data)
if st.session_state.get("last_history_sig") != sig:
    append_history(compute_kpis(base_df, umbral_bajo))
    st.session_state["last_history_sig"] = sig

# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------

k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
k1.metric("Productos visibles", f"{len(filtered):,}".replace(",", "."))
k2.metric("Stock acumulado", f"{filtered['stock_actual'].fillna(0).sum():,.0f}".replace(",", "."))
k3.metric("Stock bajo", int((filtered["estado"] == "low").sum()), help=f"Umbral ≤ {umbral_bajo:.0f}")
k4.metric("Sin stock / negativo", int(filtered["estado"].isin(["zero", "negative"]).sum()))
k5.metric("Rubros visibles", filtered["rubro"].nunique())
k6.metric("Familias visibles", filtered.loc[filtered["familia"] != "Sin dato", "familia"].nunique())
k7.metric("Proveedores visibles", filtered.loc[filtered["proveedor"] != "Sin dato", "proveedor"].nunique())

tab_resumen, tab_detalle, tab_tendencia = st.tabs(["📊 Resumen", "📋 Detalle", "📈 Tendencia"])

# ---------------------------------------------------------------------------
# Tab: Resumen (gráficos)
# ---------------------------------------------------------------------------

with tab_resumen:
    colA, colB = st.columns(2)

    with colA:
        dim_label = st.selectbox("Agrupar por", list(GROUP_DIMENSIONS.keys()), index=0)
        dim = GROUP_DIMENSIONS[dim_label]
        por_dim = (
            filtered.groupby(dim, as_index=False)["stock_actual"]
            .sum(numeric_only=True)
            .sort_values("stock_actual", ascending=False)
        )
        truncado = len(por_dim) > 20
        por_dim = por_dim.head(20).sort_values("stock_actual", ascending=True)
        fig = go.Figure(
            go.Bar(
                x=por_dim["stock_actual"],
                y=por_dim[dim],
                orientation="h",
                marker_color=PALETTE["blue"],
                hovertemplate="%{y}<br>Stock: %{x:,.0f}<extra></extra>",
            )
        )
        fig.update_layout(
            title=f"Stock actual por {dim_label.lower()}" + (" (top 20)" if truncado else ""),
            plot_bgcolor=PALETTE["surface"],
            paper_bgcolor=PALETTE["surface"],
            font_color=PALETTE["ink"],
            xaxis=dict(gridcolor=PALETTE["grid"], title=None),
            yaxis=dict(title=None),
            margin=dict(l=10, r=10, t=40, b=10),
            height=max(320, 28 * len(por_dim)),
        )
        st.plotly_chart(fig, width="stretch")

    with colB:
        urgentes = filtered[filtered["estado"].isin(["low", "zero", "negative"])]
        top_urgentes = urgentes.nsmallest(15, "stock_actual").sort_values("stock_actual", ascending=False)
        if not top_urgentes.empty:
            fig2 = go.Figure(
                go.Bar(
                    x=top_urgentes["stock_actual"],
                    y=top_urgentes["descripcion"].str.slice(0, 32),
                    orientation="h",
                    marker_color=[STATUS_COLOR[e] for e in top_urgentes["estado"]],
                    hovertemplate="%{y}<br>Stock: %{x:,.0f}<extra></extra>",
                )
            )
            fig2.update_layout(
                title="Productos que necesitan atención (menor stock)",
                plot_bgcolor=PALETTE["surface"],
                paper_bgcolor=PALETTE["surface"],
                font_color=PALETTE["ink"],
                xaxis=dict(gridcolor=PALETTE["grid"], title=None),
                yaxis=dict(title=None),
                margin=dict(l=10, r=10, t=40, b=10),
                height=460,
            )
            st.plotly_chart(fig2, width="stretch")
        else:
            st.success("No hay productos en alerta con los filtros actuales.")

    alerta = filtered[filtered["estado"].isin(["low", "zero", "negative"])].sort_values(
        "stock_actual", ascending=True
    )
    st.markdown("**Productos en alerta (bajo, sin stock o negativo)**")
    if alerta.empty:
        st.success("No hay productos en alerta con los filtros actuales.")
    else:
        st.dataframe(
            alerta[["rubro", "codigo", "descripcion", "stock_actual", "estado"]]
            .assign(estado=lambda d: d["estado"].map(STATUS_LABEL))
            .rename(columns={
                "rubro": "Rubro", "codigo": "Código", "descripcion": "Descripción",
                "stock_actual": "Stock", "estado": "Estado",
            }),
            width="stretch",
            hide_index=True,
            height=280,
        )

# ---------------------------------------------------------------------------
# Tab: Detalle (tabla completa + export) — exactamente las columnas pedidas,
# en el orden del Excel, más un indicador de Estado al final.
# ---------------------------------------------------------------------------

with tab_detalle:
    st.caption(f"{len(filtered):,} productos · ordenable haciendo clic en cada columna".replace(",", "."))

    tabla = filtered[[
        "rubro", "codigo", "descripcion", "uxb", "stock_actual",
        "familia", "subfamilia", "marca", "proveedor", "estado",
    ]].rename(columns=DISPLAY_LABELS)
    tabla["Estado"] = tabla["Estado"].map(STATUS_LABEL)

    def _row_style(row):
        color = STATUS_COLOR.get(
            [k for k, v in STATUS_LABEL.items() if v == row["Estado"]][0], PALETTE["muted"]
        )
        return [f"background-color: {color}22" if col == "Estado" else "" for col in row.index]

    def _fmt_ar(v) -> str:
        # Separador de miles con punto, como en el resto de la app (convención AR).
        return f"{v:,.0f}".replace(",", ".")

    st.dataframe(
        tabla.style.apply(_row_style, axis=1).format({"Stock": _fmt_ar}, na_rep="—"),
        width="stretch",
        hide_index=True,
        height=560,
    )

    exp1, exp2 = st.columns(2)
    with exp1:
        csv_bytes = tabla.to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button(
            "⬇️ Descargar CSV filtrado",
            data=csv_bytes,
            file_name="inventario_filtrado.csv",
            mime="text/csv",
        )
    with exp2:
        xlsx_buf = io.BytesIO()
        with pd.ExcelWriter(xlsx_buf, engine="xlsxwriter") as writer:
            tabla.to_excel(writer, index=False, sheet_name="Inventario filtrado")
        st.download_button(
            "⬇️ Descargar Excel filtrado",
            data=xlsx_buf.getvalue(),
            file_name="inventario_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---------------------------------------------------------------------------
# Tab: Tendencia (historial entre cargas)
# ---------------------------------------------------------------------------

with tab_tendencia:
    if HISTORY_FILE.exists():
        hist = pd.read_csv(HISTORY_FILE)
        if len(hist) >= 2:
            fig3 = go.Figure(
                go.Scatter(
                    x=hist["fecha"], y=hist["stock_acumulado"],
                    mode="lines+markers", line_color=PALETTE["blue"],
                    hovertemplate="%{x}<br>Stock: %{y:,.0f}<extra></extra>",
                )
            )
            fig3.update_layout(
                title="Stock acumulado a lo largo del tiempo",
                plot_bgcolor=PALETTE["surface"], paper_bgcolor=PALETTE["surface"],
                font_color=PALETTE["ink"], xaxis=dict(gridcolor=PALETTE["grid"]),
                yaxis=dict(gridcolor=PALETTE["grid"]), margin=dict(l=10, r=10, t=40, b=10),
            )
            st.plotly_chart(fig3, width="stretch")
            st.dataframe(hist.tail(20), width="stretch", hide_index=True)
        else:
            st.info(
                "Todavía hay un solo snapshot guardado. Subí el archivo de la próxima semana "
                "a la fuente central para empezar a ver la tendencia."
            )
    else:
        st.info("Todavía no hay historial guardado. Se generará al leer la fuente central.")

st.caption(
    "Fuente: archivo central administrado por la oficina · Actualización automática cada 5 minutos. "
    "Columnas: Rubro · Código · Descripción · UxB · Stock (StkActual) · Familia · "
    "Subfamilia · Marca · Proveedor."
)
