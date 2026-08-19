"""
Ventas de marcas Febeca — Dashboard Streamlit
==============================================

Reescritura en Python/Streamlit del dashboard interactivo original (HTML +
Excel). Reproduce la misma lógica de negocio que se validó a lo largo de la
construcción del dashboard:

- "Venta" y "Contribución" llegan en máscara de miles de USD -> se
  multiplican x1000.
- "Clientes Activados" se SUMA (nunca se promedia) en todos los niveles.
- "Rotación" se PROMEDIA (nunca se suma) en todos los niveles.
- El mes de una semana ISO ("2026-31") es el mes que contiene su jueves
  (regla ISO 8601), usado tanto para agrupar por mes como para el
  comparativo de Presupuesto.
- "Ritmo" = % de presupuesto alcanzado en el mes en curso ÷ % de semanas
  del mes ya transcurridas (regla de negocio "alcance de las semanas").
- La hoja "Acciones" trae bloques de celdas combinadas (una marca por
  bloque); se reconstruye rellenando la marca hacia abajo. Las marcas sin
  ninguna acción quedan con el campo vacío (nunca texto de relleno).
- El archivo recurrente que se sube para refrescar datos siempre se llama
  "Ventas actualizadas.xlsx" y trae 3 hojas: Sheet1 (ventas), Presupuesto,
  Acciones. Cada carga reemplaza por completo a la anterior.

Ejecutar localmente:  streamlit run app.py
"""

import io
import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook

# ----------------------------------------------------------------------
# Configuración de página y paleta visual (misma paleta usada en el
# dashboard HTML/Excel original, para mantener continuidad de marca).
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="Ventas de marcas Febeca",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BLUE = "#2A78D6"
ORANGE = "#EB6834"
AQUA = "#1BAF7A"
VIOLET = "#4A3AA7"
RED = "#E34948"
MUTED = "#898781"
GRID = "#E1E0D9"

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

DATA_DIR = Path(__file__).parent / "data"


# ----------------------------------------------------------------------
# Reglas de negocio compartidas
# ----------------------------------------------------------------------
def semana_a_mes(semana: str):
    """Mes ISO 8601 de una semana 'YYYY-WW': el mes que contiene su jueves."""
    y, w = semana.split("-")
    d = datetime.date.fromisocalendar(int(y), int(w), 4)
    return f"{MESES_ES[d.month - 1]} {d.year}", d.year, d.month


def weeks_in_month(year: int, month: int):
    """Todas las semanas ISO cuyo jueves cae dentro de (year, month)."""
    out = []
    for w in range(1, 54):
        try:
            d = datetime.date.fromisocalendar(year, w, 4)
        except ValueError:
            continue
        if d.year == year and d.month == month:
            out.append(f"{year}-{w:02d}")
    return out


def estado_ritmo(ritmo):
    if ritmo is None or pd.isna(ritmo):
        return "Sin presupuesto asignado"
    if ritmo >= 1:
        return "Adelantado / en línea"
    if ritmo >= 0.85:
        return "Levemente atrasado"
    return "Atrasado"


# ----------------------------------------------------------------------
# Carga de datos: por defecto desde /data (CSVs ya extraídos), o desde un
# archivo "Ventas actualizadas.xlsx" recién subido (reemplaza todo).
# ----------------------------------------------------------------------
@st.cache_data
def load_default_data():
    raw = pd.read_csv(DATA_DIR / "raw_clean.csv")
    presu = pd.read_csv(DATA_DIR / "presupuesto_raw.csv", parse_dates=["Mes_Fecha"])
    acciones = pd.read_csv(DATA_DIR / "acciones_raw.csv")
    return raw, presu, acciones


def _singularize(s):
    """Normaliza un plural simple ('acciones'->'accion', 'marcas'->'marca',
    'semanas'->'semana') para comparar encabezados de forma tolerante."""
    if s.endswith("es"):
        return s[:-2]
    if s.endswith("s"):
        return s[:-1]
    return s


def _find_header_row(ws, must_have, max_scan=15):
    """Busca la fila de encabezado que contenga todas las etiquetas de `must_have`
    (case-insensitive, tolerando plurales) en alguna columna. Devuelve (fila, {etiqueta: columna})."""
    must_have_lower = [m.lower() for m in must_have]
    for r in range(1, max_scan + 1):
        found = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            vs = str(v).strip().lower()
            for label in must_have_lower:
                if vs == label or _singularize(vs) == _singularize(label):
                    found[label] = c
        if all(m in found for m in must_have_lower):
            return r, found
    return None, {}


def parse_ventas_sheet(ws):
    header_row, cols = _find_header_row(ws, ["marca", "semana"])
    if header_row is None:
        raise ValueError('No se encontró una fila de encabezado con "Marca" y "Semana".')
    # columnas conocidas, toleran variantes de nombre
    col_map = {}
    for r in range(header_row, header_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            vs = str(v).strip().lower()
            if vs in ("marca",):
                col_map["Marca"] = c
            elif vs in ("semana",):
                col_map["Semana"] = c
            elif vs in ("grupo compra", "grupo de compra", "grupo"):
                col_map["Grupo compra"] = c
            elif vs in ("clientes activados",):
                col_map["Clientes activados"] = c
            elif vs.startswith("contribuci"):
                col_map["Contribución"] = c
            elif vs.startswith("rotaci"):
                col_map["Rotación"] = c
            elif vs in ("venta neta", "venta real", "venta"):
                col_map["Venta Neta"] = c

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        marca = ws.cell(row=r, column=col_map.get("Marca", 0)).value if "Marca" in col_map else None
        semana = ws.cell(row=r, column=col_map.get("Semana", 0)).value if "Semana" in col_map else None
        if marca is None or semana is None or str(semana).strip().lower() == "semana":
            continue
        row = {
            "Marca": str(marca).strip(),
            "Semana": str(semana).strip(),
            "Grupo compra": ws.cell(row=r, column=col_map["Grupo compra"]).value if "Grupo compra" in col_map else "N/A",
            "Clientes activados": ws.cell(row=r, column=col_map["Clientes activados"]).value if "Clientes activados" in col_map else None,
            "Contribución": ws.cell(row=r, column=col_map["Contribución"]).value if "Contribución" in col_map else None,
            "Rotación": ws.cell(row=r, column=col_map["Rotación"]).value if "Rotación" in col_map else None,
            "Venta Neta": ws.cell(row=r, column=col_map["Venta Neta"]).value if "Venta Neta" in col_map else None,
        }
        rows.append(row)
    if not rows:
        raise ValueError('La hoja de ventas no tiene filas válidas.')
    return pd.DataFrame(rows)


def parse_presupuesto_sheet(ws):
    header_row, cols = _find_header_row(ws, ["marca"])
    if header_row is None:
        return None
    marca_col = cols["marca"]
    # columnas de mes: cualquier celda con valor datetime en la fila de encabezado
    month_cols = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, datetime.datetime):
            month_cols.append((c, v.date()))
    if not month_cols:
        return None
    origen_col = grupo_col = margen_col = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        vs = str(v).strip().lower()
        if vs == "origen":
            origen_col = c
        elif vs in ("grupo compra", "grupo de compra"):
            grupo_col = c
        elif vs.startswith("margen"):
            margen_col = c

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        marca = ws.cell(row=r, column=marca_col).value
        if not marca:
            continue
        for c, mes_fecha in month_cols:
            val = ws.cell(row=r, column=c).value
            rows.append({
                "Marca": str(marca).strip(),
                "Origen": ws.cell(row=r, column=origen_col).value if origen_col else None,
                "Grupo Compra": ws.cell(row=r, column=grupo_col).value if grupo_col else None,
                "Margen_Objetivo": ws.cell(row=r, column=margen_col).value if margen_col else None,
                "Mes_Fecha": pd.Timestamp(mes_fecha),
                "Presupuesto_Miles": val,
            })
    return pd.DataFrame(rows) if rows else None


def parse_acciones_sheet(ws):
    header_row, cols = _find_header_row(ws, ["marca", "accion"])
    if header_row is None:
        return None
    marca_col = cols["marca"]
    accion_col = cols["accion"]

    # celdas combinadas de la columna Marca: se propaga el valor de la
    # celda superior del bloque a todas las filas que abarca.
    merge_fill = {}
    for rng in ws.merged_cells.ranges:
        if rng.min_col == marca_col and rng.max_col == marca_col:
            top_val = ws.cell(row=rng.min_row, column=marca_col).value
            for r in range(rng.min_row, rng.max_row + 1):
                merge_fill[r] = top_val

    rows = []
    last_marca = None
    for r in range(header_row + 1, ws.max_row + 1):
        marca_val = ws.cell(row=r, column=marca_col).value
        if marca_val is None:
            marca_val = merge_fill.get(r)
        if marca_val is not None and str(marca_val).strip() != "":
            last_marca = str(marca_val).strip()
        if last_marca is None:
            continue
        accion_val = ws.cell(row=r, column=accion_col).value
        if accion_val is None or str(accion_val).strip() == "":
            continue
        rows.append({"Marca": last_marca, "Accion": str(accion_val).strip()})
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["Marca", "Accion"])


def parse_uploaded_workbook(file_bytes):
    """Lee 'Ventas actualizadas.xlsx' (Sheet1 + Presupuesto + Acciones opcionales).
    Devuelve (raw_df, presupuesto_df_or_None, acciones_df_or_None, avisos)."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    avisos = []

    ventas_sheet_name = wb.sheetnames[0]
    raw = parse_ventas_sheet(wb[ventas_sheet_name])

    presu = None
    presu_name = next((n for n in wb.sheetnames if n.strip().lower() == "presupuesto"), None)
    if presu_name:
        presu = parse_presupuesto_sheet(wb[presu_name])
        if presu is None:
            avisos.append('No se pudo leer la hoja "Presupuesto" (no se encontraron columnas de mes).')
    else:
        avisos.append('El archivo no trae una hoja "Presupuesto"; se conserva el presupuesto anterior.')

    acciones = None
    acciones_name = next((n for n in wb.sheetnames if n.strip().lower() == "acciones"), None)
    if acciones_name:
        acciones = parse_acciones_sheet(wb[acciones_name])
        if acciones is None:
            avisos.append('No se pudo leer la hoja "Acciones" (revisa que tenga columnas "Marca" y "Accion").')
    else:
        avisos.append('El archivo no trae una hoja "Acciones"; se conservan las acciones anteriores.')

    return raw, presu, acciones, avisos


def build_full_acciones(acciones_hits: pd.DataFrame, all_brands):
    """Completa acciones_hits con Accion=None para marcas sin ninguna acción."""
    con_accion = set(acciones_hits["Marca"].unique()) if len(acciones_hits) else set()
    faltantes = sorted(set(all_brands) - con_accion)
    extra = pd.DataFrame({"Marca": faltantes, "Accion": [None] * len(faltantes)})
    full = pd.concat([acciones_hits, extra], ignore_index=True)
    return full.sort_values("Marca", kind="stable").reset_index(drop=True)


# ----------------------------------------------------------------------
# Derivación de métricas sobre el DataFrame de ventas
# ----------------------------------------------------------------------
def derive_raw(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    df["Venta_USD"] = pd.to_numeric(df["Venta Neta"], errors="coerce") * 1000
    df["Contribucion_USD"] = pd.to_numeric(df["Contribución"], errors="coerce") * 1000
    df["Clientes"] = pd.to_numeric(df["Clientes activados"], errors="coerce")
    df["Rotacion"] = pd.to_numeric(df["Rotación"], errors="coerce")
    mes_info = df["Semana"].apply(semana_a_mes)
    df["Mes"] = mes_info.apply(lambda t: t[0])
    df["Mes_Year"] = mes_info.apply(lambda t: t[1])
    df["Mes_Month"] = mes_info.apply(lambda t: t[2])
    df.rename(columns={"Grupo compra": "Grupo"}, inplace=True)
    return df


def aggregate_by_brand(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["Marca", "Grupo"], as_index=False).agg(
        venta=("Venta_USD", "sum"),
        contribucion=("Contribucion_USD", "sum"),
        clientes=("Clientes", "sum"),
        rotacion=("Rotacion", "mean"),
        rot_count=("Rotacion", "count"),
    )
    total_venta = g["venta"].sum()
    g["participacion"] = np.where(total_venta != 0, g["venta"] / total_venta, 0)
    g["margen"] = np.where(g["venta"] != 0, g["contribucion"] / g["venta"], 0)
    return g


def aggregate_by_week(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("Semana", as_index=False).agg(
        venta=("Venta_USD", "sum"),
        contribucion=("Contribucion_USD", "sum"),
        clientes=("Clientes", "sum"),
        rotacion=("Rotacion", "mean"),
    )
    return g.sort_values("Semana")


def aggregate_by_grupo(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("Grupo", as_index=False).agg(
        venta=("Venta_USD", "sum"),
        contribucion=("Contribucion_USD", "sum"),
        clientes=("Clientes", "sum"),
        rotacion=("Rotacion", "mean"),
        marcas=("Marca", "nunique"),
    )
    return g


def compute_budget_comparison(df_all: pd.DataFrame, presu: pd.DataFrame, grupos_sel, marcas_sel):
    """Presupuesto vs. Real del mes en curso completo (ignora filtro de Semana,
    pero sí respeta Grupo y Marca), igual que en el HTML/Excel originales."""
    if df_all.empty or presu is None or presu.empty:
        return None

    weeks_sorted = sorted(df_all["Semana"].unique())
    mes_label, year, month = semana_a_mes(weeks_sorted[-1])
    weeks_mes = weeks_in_month(year, month)
    reported = [w for w in weeks_mes if w in set(df_all["Semana"])]
    pct_avance = len(reported) / len(weeks_mes) if weeks_mes else 0

    venta_mes = (
        df_all[df_all["Mes"] == mes_label]
        .groupby("Marca", as_index=False)["Venta_USD"].sum()
        .rename(columns={"Venta_USD": "venta_real"})
    )

    presu_mes = presu[presu["Mes_Fecha"].dt.to_period("M") == pd.Timestamp(year=year, month=month, day=1).to_period("M")]
    presu_mes = presu_mes[presu_mes["Grupo Compra"].isin(grupos_sel) & presu_mes["Marca"].isin(marcas_sel)]
    if presu_mes.empty:
        return None
    presu_mes = presu_mes[["Marca", "Grupo Compra", "Presupuesto_Miles"]].copy()
    presu_mes["presupuesto"] = presu_mes["Presupuesto_Miles"] * 1000

    rows = presu_mes.merge(venta_mes, on="Marca", how="left")
    rows["venta_real"] = rows["venta_real"].fillna(0)
    rows["cumplimiento"] = np.where(rows["presupuesto"] != 0, rows["venta_real"] / rows["presupuesto"], np.nan)
    rows["ritmo"] = rows["cumplimiento"] / pct_avance if pct_avance > 0 else np.nan
    rows["estado"] = rows["ritmo"].apply(estado_ritmo)
    rows = rows.rename(columns={"Grupo Compra": "Grupo"})[
        ["Marca", "Grupo", "presupuesto", "venta_real", "cumplimiento", "ritmo", "estado"]
    ]

    total_presu = rows["presupuesto"].sum()
    total_real = rows["venta_real"].sum()
    total_cumpl = total_real / total_presu if total_presu else 0
    total_ritmo = total_cumpl / pct_avance if pct_avance > 0 else 0

    return {
        "mes_label": mes_label,
        "weeks_mes": weeks_mes,
        "reported": reported,
        "pct_avance": pct_avance,
        "rows": rows.sort_values("presupuesto", ascending=False).reset_index(drop=True),
        "total_presu": total_presu,
        "total_real": total_real,
        "total_cumpl": total_cumpl,
        "total_ritmo": total_ritmo,
    }


def order_acciones_by_budget(acciones_full: pd.DataFrame, presu: pd.DataFrame, df_all: pd.DataFrame):
    """Ordena la tabla de Acciones de mayor a menor presupuesto del mes en
    curso; las marcas sin presupuesto quedan al final, en orden alfabético."""
    budget_map = {}
    if presu is not None and not presu.empty and not df_all.empty:
        weeks_sorted = sorted(df_all["Semana"].unique())
        _, year, month = semana_a_mes(weeks_sorted[-1])
        presu_mes = presu[presu["Mes_Fecha"].dt.to_period("M") == pd.Timestamp(year=year, month=month, day=1).to_period("M")]
        budget_map = dict(zip(presu_mes["Marca"], presu_mes["Presupuesto_Miles"] * 1000))

    def sort_key(marca):
        if marca in budget_map:
            return (0, -budget_map[marca], marca)
        return (1, 0, marca)

    marcas_unicas = sorted(acciones_full["Marca"].unique(), key=sort_key)
    rank = {m: i for i, m in enumerate(marcas_unicas)}
    out = acciones_full.copy()
    out["_rank"] = out["Marca"].map(rank)
    return out.sort_values("_rank", kind="stable").drop(columns="_rank").reset_index(drop=True)


# ----------------------------------------------------------------------
# Gráficos (Plotly), misma paleta y criterio visual que el HTML original
# ----------------------------------------------------------------------
def fmt_usd(v):
    return f"${v:,.0f}"


def fmt_num(v):
    return f"{v:,.0f}"


def fmt_num1(v):
    return f"{v:,.1f}"


def fmt_pct(v):
    return f"{v:.1%}"


def hbar_chart(labels, values, color, fmt=fmt_usd, title=None):
    values = list(values)
    fig = go.Figure(
        go.Bar(
            x=values,
            y=list(labels),
            orientation="h",
            marker_color=color,
            text=[fmt(v) for v in values],
            textposition="outside",
            cliponaxis=False,
        )
    )
    fig.update_layout(
        title=title,
        template="plotly_white",
        margin=dict(l=10, r=30, t=40 if title else 10, b=10),
        yaxis=dict(autorange="reversed", showgrid=False),
        xaxis=dict(showgrid=True, gridcolor=GRID, zeroline=False, tickformat="~s"),
        height=max(220, 34 * len(values)),
        showlegend=False,
    )
    return fig


def line_chart(labels, values, color, fmt=fmt_usd, title=None):
    values = list(values)
    fig = go.Figure(
        go.Scatter(
            x=list(labels), y=values, mode="lines+markers+text",
            line=dict(color=color, width=3),
            marker=dict(color=color, size=8),
            text=[fmt(v) for v in values],
            textposition="top center",
        )
    )
    fig.update_layout(
        title=title,
        template="plotly_white",
        margin=dict(l=10, r=10, t=40 if title else 10, b=10),
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=True, gridcolor=GRID, zeroline=False, tickformat="~s"),
        height=320,
        showlegend=False,
    )
    return fig


def vbar_chart(labels, values, color, fmt=fmt_num, title=None):
    values = list(values)
    fig = go.Figure(
        go.Bar(
            x=list(labels), y=values, marker_color=color,
            text=[fmt(v) for v in values],
            textposition="outside", cliponaxis=False,
        )
    )
    fig.update_layout(
        title=title,
        template="plotly_white",
        margin=dict(l=10, r=10, t=40 if title else 10, b=10),
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=True, gridcolor=GRID, zeroline=False, tickformat="~s"),
        height=320,
        showlegend=False,
    )
    return fig


# ----------------------------------------------------------------------
# Estado de sesión: datos activos (por defecto o reemplazados por upload)
# ----------------------------------------------------------------------
if "raw" not in st.session_state:
    raw0, presu0, acciones0 = load_default_data()
    st.session_state.raw = raw0
    st.session_state.presu = presu0
    st.session_state.acciones_hits = acciones0
    st.session_state.data_label = "data 100.xlsx (datos de ejemplo incluidos)"

with st.sidebar:
    st.header("📤 Actualizar datos")
    st.caption(
        'Sube el archivo recurrente **"Ventas actualizadas.xlsx"** '
        '(hojas Sheet1, Presupuesto y Acciones). Reemplaza por completo '
        'los datos actuales.'
    )
    uploaded = st.file_uploader("Archivo .xlsx", type=["xlsx"], label_visibility="collapsed")
    if uploaded is not None:
        try:
            raw_new, presu_new, acciones_new, avisos = parse_uploaded_workbook(uploaded.read())
            st.session_state.raw = raw_new
            if presu_new is not None:
                st.session_state.presu = presu_new
            if acciones_new is not None:
                st.session_state.acciones_hits = acciones_new
            st.session_state.data_label = uploaded.name
            st.success(f"✓ Datos actualizados desde “{uploaded.name}” — {len(raw_new)} filas.")
            for a in avisos:
                st.info(a)
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")

    st.divider()
    st.caption(
        "El presupuesto y las acciones se refrescan solo si el archivo "
        "subido trae esas hojas; de lo contrario se conservan los últimos "
        "valores cargados."
    )

raw = derive_raw(st.session_state.raw)
presu = st.session_state.presu
acciones_hits = st.session_state.acciones_hits

ALL_WEEKS = sorted(raw["Semana"].unique())
ALL_GRUPOS = sorted(raw["Grupo"].unique())
ALL_BRANDS = sorted(raw["Marca"].unique())
BRAND_GRUPO = dict(zip(raw["Marca"], raw["Grupo"]))

acciones_full = build_full_acciones(acciones_hits, ALL_BRANDS)

# ----------------------------------------------------------------------
# Encabezado + filtros
# ----------------------------------------------------------------------
st.title("Ventas de marcas Febeca")
st.caption(
    f"Venta, Activación de clientes, Rotación y Contribución por marca · "
    f"Semanas {ALL_WEEKS[0]} a {ALL_WEEKS[-1]} · Fuente: {st.session_state.data_label}"
)

f1, f2, f3, f4 = st.columns([1.3, 1, 1.3, 0.8])
with f1:
    weeks_sel = st.multiselect("Semana", ALL_WEEKS, default=ALL_WEEKS)
with f2:
    grupos_sel = st.multiselect("Grupo de compra", ALL_GRUPOS, default=ALL_GRUPOS)
with f3:
    brands_sel = st.multiselect("Marca", ALL_BRANDS, default=ALL_BRANDS)
with f4:
    top_n = st.selectbox("Top N", [5, 10, 15, len(ALL_BRANDS)], index=1)

if not weeks_sel or not grupos_sel or not brands_sel:
    st.warning("Selecciona al menos una Semana, un Grupo de Compra y una Marca.")
    st.stop()

df_filtered = raw[raw["Semana"].isin(weeks_sel) & raw["Grupo"].isin(grupos_sel) & raw["Marca"].isin(brands_sel)]
# Presupuesto vs. Real ignora el filtro de Semana (mira siempre el mes completo)
df_for_budget = raw[raw["Grupo"].isin(grupos_sel) & raw["Marca"].isin(brands_sel)]

brand_agg = aggregate_by_brand(df_filtered)
week_agg = aggregate_by_week(df_filtered)
grupo_agg = aggregate_by_grupo(df_filtered)

# ----------------------------------------------------------------------
# Selector de vista: Presupuesto / KPI
# ----------------------------------------------------------------------
view = st.segmented_control(
    "Vista", ["💰 Presupuesto", "📈 KPI"], default="💰 Presupuesto", label_visibility="collapsed"
)
if view is None:
    view = "💰 Presupuesto"

if view == "💰 Presupuesto":
    st.subheader("Presupuesto vs. Real")
    budget = compute_budget_comparison(df_for_budget, presu, grupos_sel, brands_sel)
    if budget is None:
        st.info(
            "No hay marcas con presupuesto asignado para los filtros de Grupo/Marca "
            "seleccionados, o no se cargó una hoja de Presupuesto todavía."
        )
    else:
        st.caption(
            f"Mes en curso: **{budget['mes_label']}** — compara siempre el mes completo "
            "(no responde al filtro de Semana), pero sí a Grupo de Compra y Marca. "
            f"Semanas reportadas: {len(budget['reported'])} de {len(budget['weeks_mes'])} "
            f"({budget['pct_avance']:.1%} de avance del mes)."
        )
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric(f"Presupuesto {budget['mes_label']} (USD)", f"${budget['total_presu']:,.0f}")
        c2.metric(f"Venta Real {budget['mes_label']} (USD)", f"${budget['total_real']:,.0f}")
        c3.metric("% Cumplimiento presupuesto", f"{budget['total_cumpl']:.1%}")
        c4.metric("Avance del mes (semanas)", f"{budget['pct_avance']:.1%}")
        c5.metric("Ritmo vs. calendario", f"{budget['total_ritmo']:.1%}")

        rows_sorted = budget["rows"].sort_values("cumplimiento", ascending=True)
        st.plotly_chart(
            hbar_chart(
                rows_sorted["Marca"], rows_sorted["cumplimiento"], AQUA,
                fmt=fmt_pct, title="% Cumplimiento de presupuesto por marca (mes en curso)",
            ),
            width='stretch',
        )

        with st.expander("Ver tabla de presupuesto por marca"):
            tabla = budget["rows"].copy()
            tabla["presupuesto"] = tabla["presupuesto"].map(lambda v: f"${v:,.0f}")
            tabla["venta_real"] = tabla["venta_real"].map(lambda v: f"${v:,.0f}")
            tabla["cumplimiento"] = tabla["cumplimiento"].map(lambda v: f"{v:.1%}" if pd.notna(v) else "-")
            tabla["ritmo"] = tabla["ritmo"].map(lambda v: f"{v:.1%}" if pd.notna(v) else "-")
            tabla.columns = ["Marca", "Grupo", "Presupuesto (USD)", "Venta Real (USD)", "% Cumplimiento", "Ritmo vs. calendario", "Estado"]
            st.dataframe(tabla, width='stretch', hide_index=True)

else:
    st.subheader("Indicadores y KPI")
    total_venta = brand_agg["venta"].sum()
    total_contrib = brand_agg["contribucion"].sum()
    margen = total_contrib / total_venta if total_venta else 0
    total_clientes = brand_agg["clientes"].sum()
    marcas_activas = (brand_agg["venta"] > 0).sum()
    top3 = brand_agg.nlargest(3, "venta")["venta"].sum()
    conc = top3 / total_venta if total_venta else 0

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Venta total (USD)", f"${total_venta:,.0f}")
    k2.metric("Contribución total (USD)", f"${total_contrib:,.0f}")
    k3.metric("Margen de contribución", f"{margen:.1%}")
    k4.metric("Clientes activados", f"{total_clientes:,.0f}")
    k5.metric("Marcas activas", f"{marcas_activas} / {len(brand_agg)}")
    k6.metric("Concentración Top 3 marcas", f"{conc:.1%}")

    top_venta = brand_agg.nlargest(top_n, "venta")
    top_contrib = brand_agg.nlargest(top_n, "contribucion")
    top_clientes = brand_agg.nlargest(top_n, "clientes")
    top_rot = brand_agg[brand_agg["rot_count"] > 0].nlargest(top_n, "rotacion")

    cc1, cc2 = st.columns(2)
    with cc1:
        st.plotly_chart(hbar_chart(top_venta["Marca"], top_venta["venta"], BLUE, title=f"Top {len(top_venta)} marcas por venta (USD)"), width='stretch')
        st.plotly_chart(hbar_chart(top_clientes["Marca"], top_clientes["clientes"], AQUA, fmt=fmt_num, title=f"Top {len(top_clientes)} marcas por clientes activados"), width='stretch')
    with cc2:
        st.plotly_chart(hbar_chart(top_contrib["Marca"], top_contrib["contribucion"], ORANGE, title=f"Top {len(top_contrib)} marcas por contribución (USD)"), width='stretch')
        st.plotly_chart(hbar_chart(top_rot["Marca"], top_rot["rotacion"], RED, fmt=fmt_num1, title=f"Top {len(top_rot)} marcas por índice de rotación"), width='stretch')

    cc3, cc4 = st.columns(2)
    with cc3:
        st.plotly_chart(line_chart(week_agg["Semana"], week_agg["venta"], BLUE, title="Evolución semanal de venta (USD)"), width='stretch')
    with cc4:
        st.plotly_chart(vbar_chart(week_agg["Semana"], week_agg["clientes"], VIOLET, title="Clientes activados por semana (total)"), width='stretch')

    st.plotly_chart(vbar_chart(grupo_agg["Grupo"], grupo_agg["venta"], BLUE, title="Venta por grupo de compra (USD)"), width='stretch')

# ----------------------------------------------------------------------
# Tabla de datos por marca + matriz de rotación por mes (siempre visibles)
# ----------------------------------------------------------------------
with st.expander("Ver tabla de datos por marca (ordenable)"):
    tabla = brand_agg.copy()
    tabla["venta"] = tabla["venta"].map(lambda v: f"${v:,.0f}")
    tabla["participacion"] = tabla["participacion"].map(lambda v: f"{v:.1%}")
    tabla["contribucion"] = tabla["contribucion"].map(lambda v: f"${v:,.0f}")
    tabla["margen"] = tabla["margen"].map(lambda v: f"{v:.1%}")
    tabla["clientes"] = tabla["clientes"].map(lambda v: f"{v:,.0f}")
    tabla["rotacion"] = brand_agg.apply(lambda r: f"{r['rotacion']:,.1f}" if r["rot_count"] > 0 else "-", axis=1)
    tabla = tabla[["Marca", "Grupo", "venta", "participacion", "contribucion", "margen", "clientes", "rotacion"]]
    tabla.columns = ["Marca", "Grupo", "Venta (USD)", "% Participación", "Contribución (USD)", "Margen %", "Clientes Activados", "Rotación"]
    st.dataframe(tabla, width='stretch', hide_index=True)

with st.expander("Ver rotación promedio por mes y marca (matriz)"):
    st.caption(
        'Promedio simple de "Rotación" por marca dentro de cada mes '
        "(nunca se suma)."
    )
    pivot = df_filtered.pivot_table(index="Marca", columns="Mes", values="Rotacion", aggfunc="mean")
    # incluye todos los meses presentes en los datos filtrados aunque ningún
    # valor de Rotación caiga en ese mes (se muestra la columna en blanco,
    # igual que el "-" del Excel), en vez de omitirla silenciosamente.
    meses_presentes = df_filtered["Mes"].unique().tolist()
    meses_orden = sorted(meses_presentes, key=lambda m: (int(m.split()[-1]), MESES_ES.index(m.split()[0])))
    pivot = pivot.reindex(columns=meses_orden).astype(float).round(1)
    # se formatea a texto explícitamente ("-" para celdas sin datos) porque
    # una columna 100% NaN se puede renderizar como "None" en st.dataframe.
    pivot_display = pivot.map(lambda v: "-" if pd.isna(v) else f"{v:,.1f}")
    st.dataframe(pivot_display, width='stretch')

# ----------------------------------------------------------------------
# Acciones por marca — "ventana aparte" (modal nativo de Streamlit)
# ----------------------------------------------------------------------
@st.dialog("📋 Acciones por Marca", width="large")
def acciones_dialog():
    st.caption(
        "Ordenada de mayor a menor presupuesto del mes en curso (las marcas sin "
        "presupuesto quedan al final, alfabético). Respeta los filtros de Grupo "
        "de Compra y Marca de arriba, no el de Semana."
    )
    filtro_marca = st.text_input("Buscar marca…", key="acc_f_marca")
    ordered = order_acciones_by_budget(acciones_full, presu, df_for_budget)
    view_df = ordered[ordered["Marca"].isin(brands_sel) & ordered["Marca"].map(BRAND_GRUPO).isin(grupos_sel)].copy()
    if filtro_marca:
        view_df = view_df[view_df["Marca"].str.contains(filtro_marca, case=False, na=False)]
    view_df["Grupo"] = view_df["Marca"].map(BRAND_GRUPO)
    view_df["Accion"] = view_df["Accion"].fillna("")
    view_df = view_df[["Marca", "Grupo", "Accion"]].rename(columns={"Accion": "Acción"})
    st.dataframe(view_df, width='stretch', hide_index=True, height=420)


if st.button("📋 Acciones por marca"):
    acciones_dialog()

# ----------------------------------------------------------------------
# Notas al pie
# ----------------------------------------------------------------------
st.divider()
st.caption(
    '**Notas:** "Venta" corresponde a la columna "Venta Neta" del archivo original '
    "(máscara miles de USD, convertida a USD); \"Contribución\" también se convirtió "
    'de miles de USD a USD. "Clientes Activados" se SUMA por marca/grupo/semana; '
    '"Rotación" siempre se PROMEDIA, nunca se suma — se reporta mayormente en semanas '
    "sin venta y debe interpretarse con cautela. \"Grupo de Compra\" es un atributo "
    'propio de cada marca. La sección "Presupuesto vs. Real" compara siempre el mes '
    "en curso completo y solo incluye marcas con presupuesto asignado; \"Ritmo\" "
    "compara el % de presupuesto alcanzado contra el % de semanas ya transcurridas "
    "del mes. Las marcas sin ninguna acción registrada se muestran con la celda de "
    "Acción vacía."
)
