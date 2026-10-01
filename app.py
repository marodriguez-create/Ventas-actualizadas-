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
- "Inventario" (columna opcional en Sheet1) llega en máscara de miles de USD
  igual que Venta y Contribución -> se multiplica x1000. Es una foto del
  nivel de stock: se muestra el último valor reportado por marca, nunca se
  suma entre semanas (igual criterio que Rotación).
- "Presupuesto de contribución" (columna opcional en Sheet1) es un total por
  marca, no una serie semanal: se compara contra la Contribución real del mes
  en curso con el mismo criterio de Presupuesto vs. Real (cumplimiento,
  ritmo, semáforo) que ya se usa para Venta.
- "Región" (columna opcional en Sheet1) puede traer varias filas por
  Marca+Semana (una por Región). Venta, Contribución y Clientes Activados se
  SUMAN normalmente entre esas filas (el total por marca no cambia); pero
  "Rotación" y "Presupuesto de contribución" llegan replicados/repartidos
  entre Regiones de una misma Marca+Semana, así que primero se colapsan a
  nivel de semana (promedio o suma según corresponda) ANTES de agregarlos
  por marca o por mes, para no sesgar el resultado según cuántas Regiones
  se hayan reportado esa semana. Hay una vista aparte "KPI por Región" que
  permite filtrar los mismos KPI por una o varias Regiones; si el archivo no
  trae la columna "Región", todo se trata como una sola región ("N/A").
- "Sheet2" (hoja opcional): venta COMPLETA por Marca y Mes calendario, tal
  como la calcula el sistema de origen (columnas Marca + una columna de mes
  tipo "8-2026" + "Venta Neta"). Se usa para corregir el monto de "Venta" de
  cualquier semana "bisagra" compartida entre dos meses: en vez de contar esa
  semana completa en el mes que se elija, se usa directamente el total de
  Sheet2 de ese Marca+Mes — así, la misma semana bisagra aporta un monto
  distinto al filtrar Agosto que al filtrar Septiembre, cada uno con su
  propio total oficial. Si Sheet2 no trae un Marca+Mes (archivo sin esa hoja,
  o mes sin dato todavía), NO se cuenta ninguna semana "bisagra" de ese mes
  (se suman solo las semanas cargadas que tocan EXCLUSIVAMENTE ese mes) hasta
  que Sheet2 confirme cuánto le corresponde a cada mes — así un mes sin datos
  cargados todavía no muestra Venta Real solo porque su semana bisagra con el
  mes anterior ya está cargada. Esta corrección solo aplica a "Venta" (no a
  Contribución, Clientes ni Rotación). En la vista "KPI por Región" se aplica
  igual, con una salvedad: el total de Sheet2 (que no trae desglose regional)
  solo se usa cuando están seleccionadas TODAS las Regiones; si se filtra a
  un subconjunto de Regiones, igual se excluye la semana bisagra no
  confirmada, aunque sin la precisión exacta de Sheet2 para ese subconjunto.

Ejecutar localmente:  streamlit run app.py
"""

import io
import re
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
    """Todas las semanas ISO cuyo jueves cae dentro de (year, month). Se usa
    solo para construir la lista de meses disponibles; para decidir qué
    semanas pertenecen a un mes seleccionado en el filtro se usa
    weeks_for_meses, que considera semanas "bisagra" en los dos meses que
    tocan."""
    out = []
    for w in range(1, 54):
        try:
            d = datetime.date.fromisocalendar(year, w, 4)
        except ValueError:
            continue
        if d.year == year and d.month == month:
            out.append(f"{year}-{w:02d}")
    return out


def mes_label_de(year: int, month: int) -> str:
    return f"{MESES_ES[month - 1]} {year}"


def semana_bounds(semana: str):
    """(lunes, domingo) de una semana ISO 'YYYY-WW'."""
    y, w = semana.split("-")
    lunes = datetime.date.fromisocalendar(int(y), int(w), 1)
    domingo = datetime.date.fromisocalendar(int(y), int(w), 7)
    return lunes, domingo


def meses_tocados_por_semana(semana: str):
    """Meses (year, month) que toca el rango lunes-domingo de una semana ISO
    — normalmente uno solo; dos si la semana es "bisagra" entre dos meses
    (p. ej. "2026-40", lunes 28-sep a domingo 4-oct, toca Septiembre Y
    Octubre)."""
    lunes, domingo = semana_bounds(semana)
    return {(lunes.year, lunes.month), (domingo.year, domingo.month)}


def meses_disponibles(all_weeks):
    """Lista ordenada de (year, month) para todos los meses que toca, aunque
    sea parcialmente, alguna semana cargada. Alimenta las opciones del filtro
    de Mes."""
    meses = set()
    for w in all_weeks:
        meses |= meses_tocados_por_semana(w)
    return sorted(meses)


def weeks_for_meses(all_weeks, meses_sel):
    """Todas las semanas (de las cargadas) cuyo rango lunes-domingo toca
    alguno de los meses (year, month) en meses_sel. Una semana "bisagra"
    entre dos meses aparece si seleccionas cualquiera de los dos — pero sin
    duplicarse en el resultado (es un conjunto de semanas), así que
    seleccionar varios meses junto con su semana bisagra compartida no cuenta
    esa semana dos veces."""
    meses_set = set(meses_sel)
    return sorted(w for w in all_weeks if meses_tocados_por_semana(w) & meses_set)


def mes_en_curso_default(all_weeks):
    """Mes (year, month) preseleccionado por defecto en el filtro de Mes: el
    mes calendario real de hoy, salvo que los datos cargados sean más viejos
    que hoy (archivo histórico o de prueba), en cuyo caso se usa el mes
    (regla ISO del jueves) de la última semana cargada, para no dejar el
    dashboard vacío por defecto."""
    hoy = datetime.date.today()
    year, month = hoy.year, hoy.month
    if all_weeks:
        _, ult_year, ult_month = semana_a_mes(sorted(all_weeks)[-1])
        if (ult_year, ult_month) < (year, month):
            year, month = ult_year, ult_month
    return year, month


def venta_real_marca_mes(df_all: pd.DataFrame, venta_mensual_lookup: dict, all_weeks, marca: str, year: int, month: int, usar_sheet2: bool = True):
    """Venta real de una Marca para un mes calendario completo (year, month).

    Si la hoja "Sheet2" (venta completa por Marca y Mes, tal como la calcula
    el sistema de origen) trae el dato de esa Marca+Mes Y `usar_sheet2` es
    True, se usa DIRECTAMENTE — ya es el total correcto del mes completo,
    incluyendo automáticamente la porción que le corresponde de cualquier
    semana "bisagra" compartida con el mes vecino (equivale a restarle a ese
    total las semanas que no chocan, como pidió María Antonieta, pero sin
    tener que aislar cada semana bisagra una por una: la resta y el total
    dan el mismo resultado). Por eso una misma semana bisagra puede aportar
    un monto distinto según se esté mirando Agosto o Septiembre — cada mes
    usa su propio total de Sheet2.

    `usar_sheet2=False` se usa cuando `df_all` ya viene filtrado a un
    subconjunto (p. ej. una o varias Región(es) específicas) para el que
    Sheet2 no tiene desglose: el total de Sheet2 es de TODAS las regiones
    juntas, así que aplicarlo a un subconjunto de regiones daría un número
    incorrecto. En ese caso (o cuando Sheet2 simplemente no trae esa
    Marca+Mes) se suman únicamente las semanas cargadas que tocan
    EXCLUSIVAMENTE ese mes (sin compartirlo con el mes vecino); cualquier
    semana "bisagra" se deja afuera hasta que se pueda confirmar cuánto le
    corresponde a cada mes, en vez de contarla completa como antes — así,
    por ejemplo, Octubre no muestra como "Venta Real" la semana 2026-40 (que
    es mayormente de Septiembre) solo porque esa semana ya está cargada."""
    key = (marca, year, month)
    if usar_sheet2 and venta_mensual_lookup and key in venta_mensual_lookup:
        return venta_mensual_lookup[key]
    weeks_mes_exclusivas = [
        w for w in weeks_for_meses(all_weeks, [(year, month)])
        if meses_tocados_por_semana(w) == {(year, month)}
    ]
    sub = df_all[(df_all["Marca"] == marca) & (df_all["Semana"].isin(weeks_mes_exclusivas))]
    return sub["Venta_USD"].sum()


def venta_real_por_marca(df_all: pd.DataFrame, venta_mensual_lookup: dict, all_weeks, marcas, meses_sel, usar_sheet2: bool = True) -> dict:
    """{Marca: venta_real} sumando `venta_real_marca_mes` sobre todos los
    meses seleccionados (si se eligieron varios meses, se suman sus
    totales — cada mes de Sheet2 ya es una porción exclusiva, así que no hay
    riesgo de contar dos veces una semana bisagra compartida entre dos meses
    que estén AMBOS seleccionados). Ver `venta_real_marca_mes` para el uso
    de `usar_sheet2`."""
    return {
        marca: sum(
            venta_real_marca_mes(df_all, venta_mensual_lookup, all_weeks, marca, y, m, usar_sheet2=usar_sheet2)
            for (y, m) in meses_sel
        )
        for marca in marcas
    }


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
    venta_mensual_path = DATA_DIR / "venta_mensual_raw.csv"
    venta_mensual = pd.read_csv(venta_mensual_path) if venta_mensual_path.exists() else None
    return raw, presu, acciones, venta_mensual


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
            elif vs.startswith("presupuesto de contribuci"):
                col_map["Presupuesto Contribucion"] = c
            elif vs.startswith("contribuci"):
                col_map["Contribución"] = c
            elif vs.startswith("rotaci"):
                col_map["Rotación"] = c
            elif vs in ("venta neta", "venta real", "venta"):
                col_map["Venta Neta"] = c
            elif vs in ("inventario",):
                col_map["Inventario"] = c
            elif vs in ("región", "region"):
                col_map["Región"] = c

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
            "Región": str(ws.cell(row=r, column=col_map["Región"]).value).strip() if "Región" in col_map and ws.cell(row=r, column=col_map["Región"]).value is not None else "N/A",
            "Clientes activados": ws.cell(row=r, column=col_map["Clientes activados"]).value if "Clientes activados" in col_map else None,
            "Contribución": ws.cell(row=r, column=col_map["Contribución"]).value if "Contribución" in col_map else None,
            "Rotación": ws.cell(row=r, column=col_map["Rotación"]).value if "Rotación" in col_map else None,
            "Venta Neta": ws.cell(row=r, column=col_map["Venta Neta"]).value if "Venta Neta" in col_map else None,
            "Inventario": ws.cell(row=r, column=col_map["Inventario"]).value if "Inventario" in col_map else None,
            "Presupuesto Contribucion": ws.cell(row=r, column=col_map["Presupuesto Contribucion"]).value if "Presupuesto Contribucion" in col_map else None,
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


def parse_venta_mensual_sheet(ws):
    """Hoja opcional "Sheet2": venta COMPLETA por Marca y Mes calendario (no
    por semana ISO), tal como la calcula el sistema de origen. Se usa para
    corregir el monto de "Venta" de cualquier semana "bisagra" compartida
    entre dos meses — ver `venta_real_marca_mes`. La columna de mes suele
    traer un encabezado técnico (p. ej. "'Dim_Calendario'[Nro Mes]"), así que
    en vez de buscarla por nombre se detecta la columna cuyos valores tienen
    forma "M-AAAA" o "MM-AAAA" (p. ej. "8-2026" = Agosto 2026)."""
    header_row, cols = _find_header_row(ws, ["marca"])
    if header_row is None:
        return None
    marca_col = cols["marca"]

    venta_col = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        vs = str(v).strip().lower()
        if vs in ("venta neta", "venta real", "venta", "venta completa"):
            venta_col = c
    if venta_col is None:
        return None

    mes_pat = re.compile(r"^\s*(\d{1,2})\s*-\s*(\d{4})\s*$")
    mes_col = None
    for c in range(1, ws.max_column + 1):
        if c in (marca_col, venta_col):
            continue
        sample = ws.cell(row=header_row + 1, column=c).value
        if sample is not None and mes_pat.match(str(sample)):
            mes_col = c
            break
    if mes_col is None:
        return None

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        marca = ws.cell(row=r, column=marca_col).value
        mes_val = ws.cell(row=r, column=mes_col).value
        if marca is None or mes_val is None:
            continue
        m = mes_pat.match(str(mes_val))
        if not m:
            continue
        rows.append({
            "Marca": str(marca).strip(),
            "Mes_Month": int(m.group(1)),
            "Mes_Year": int(m.group(2)),
            "VentaCompleta_Miles": ws.cell(row=r, column=venta_col).value,
        })
    return pd.DataFrame(rows) if rows else None


def parse_uploaded_workbook(file_bytes):
    """Lee 'Ventas actualizadas.xlsx' (Sheet1 + Presupuesto + Acciones +
    Sheet2 opcionales). Devuelve (raw_df, presupuesto_df_or_None,
    acciones_df_or_None, venta_mensual_df_or_None, avisos)."""
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

    venta_mensual = None
    venta_mensual_name = next((n for n in wb.sheetnames if n.strip().lower() == "sheet2"), None)
    if venta_mensual_name:
        venta_mensual = parse_venta_mensual_sheet(wb[venta_mensual_name])
        if venta_mensual is None:
            avisos.append(
                'No se pudo leer la hoja "Sheet2" (se esperan columnas "Marca", una columna de mes '
                'tipo "8-2026" y "Venta Neta"); se conserva la venta mensual anterior, si había.'
            )
    # Sheet2 es opcional: si no viene, simplemente se usa la regla anterior
    # para la(s) semana(s) "bisagra" (no hace falta avisar cada vez).

    return raw, presu, acciones, venta_mensual, avisos


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
    # Inventario y Presupuesto Contribucion vienen en máscara de miles de USD,
    # igual que Venta y Contribución -> se multiplican x1000.
    # Ambas columnas son opcionales — si el archivo/CSV no las trae, quedan en NaN.
    df["Inventario_Val"] = (
        pd.to_numeric(df["Inventario"], errors="coerce") * 1000
        if "Inventario" in df.columns else np.nan
    )
    df["PresupContrib_Miles"] = (
        pd.to_numeric(df["Presupuesto Contribucion"], errors="coerce")
        if "Presupuesto Contribucion" in df.columns else np.nan
    )
    # "Región" es opcional: si el archivo/CSV no la trae, se trata todo como
    # una sola región ("N/A") para que el resto de la lógica no tenga que
    # distinguir casos.
    if "Región" not in df.columns:
        df["Región"] = "N/A"
    else:
        df["Región"] = df["Región"].fillna("N/A").astype(str).str.strip().replace("", "N/A")
    mes_info = df["Semana"].apply(semana_a_mes)
    df["Mes"] = mes_info.apply(lambda t: t[0])
    df["Mes_Year"] = mes_info.apply(lambda t: t[1])
    df["Mes_Month"] = mes_info.apply(lambda t: t[2])
    df.rename(columns={"Grupo compra": "Grupo"}, inplace=True)
    return df


def collapse_region_rotacion(df: pd.DataFrame) -> pd.DataFrame:
    """Una fila por Marca+Grupo+Semana+Mes (colapsando Región) con el
    promedio de Rotación de esa semana. Cuando el archivo trae varias filas
    por Marca+Semana (una por Región), "Rotación" llega replicada con el
    mismo valor en cada una; promediar aquí ANTES de promediar entre semanas
    evita que una semana con más Regiones reportadas (o una Marca sin
    desglose regional) pese distinto que otra en el promedio final. El
    criterio de negocio sigue siendo el mismo: Rotación nunca se suma,
    siempre se promedia — ahora en dos pasos (semana, luego marca/grupo/mes)
    en vez de uno solo."""
    cols = ["Marca", "Grupo", "Semana", "Mes", "Mes_Year", "Mes_Month"]
    return df.groupby(cols, as_index=False)["Rotacion"].mean()


def aggregate_by_brand(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["Marca", "Grupo"], as_index=False).agg(
        venta=("Venta_USD", "sum"),
        contribucion=("Contribucion_USD", "sum"),
        clientes=("Clientes", "sum"),
    )
    rot = collapse_region_rotacion(df).groupby(["Marca", "Grupo"], as_index=False).agg(
        rotacion=("Rotacion", "mean"),
        rot_count=("Rotacion", "count"),
    )
    g = g.merge(rot, on=["Marca", "Grupo"], how="left")
    total_venta = g["venta"].sum()
    g["participacion"] = np.where(total_venta != 0, g["venta"] / total_venta, 0)
    g["margen"] = np.where(g["venta"] != 0, g["contribucion"] / g["venta"], 0)
    return g


def aggregate_by_week(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("Semana", as_index=False).agg(
        venta=("Venta_USD", "sum"),
        contribucion=("Contribucion_USD", "sum"),
        clientes=("Clientes", "sum"),
    )
    rot = collapse_region_rotacion(df).groupby("Semana", as_index=False)["Rotacion"].mean().rename(columns={"Rotacion": "rotacion"})
    g = g.merge(rot, on="Semana", how="left")
    return g.sort_values("Semana")


def aggregate_by_week_region(df: pd.DataFrame) -> pd.DataFrame:
    """Venta semanal desglosada por Región, para comparar su evolución
    semana a semana (solo en la vista 'KPI por Región'). Suma cruda de
    Venta_USD por Semana y Región, igual criterio que `aggregate_by_week`:
    no aplica el resguardo de Sheet2 (ese resguardo es para totales
    mensuales, no para la serie semanal cruda)."""
    return df.groupby(["Semana", "Región"], as_index=False).agg(venta=("Venta_USD", "sum")).sort_values("Semana")


def aggregate_by_grupo(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("Grupo", as_index=False).agg(
        venta=("Venta_USD", "sum"),
        contribucion=("Contribucion_USD", "sum"),
        clientes=("Clientes", "sum"),
        marcas=("Marca", "nunique"),
    )
    rot = collapse_region_rotacion(df).groupby("Grupo", as_index=False)["Rotacion"].mean().rename(columns={"Rotacion": "rotacion"})
    g = g.merge(rot, on="Grupo", how="left")
    return g


def kpi_por_region_tabla(
    df_region: pd.DataFrame, df_region_budget: pd.DataFrame, region_sel,
    all_weeks=None, meses_sel=None, venta_mensual_lookup=None, semana_filtro_es_mes_completo: bool = False,
) -> pd.DataFrame:
    """Una fila por Región (de las seleccionadas) con los mismos indicadores
    que las tarjetas de KPI: Venta, Contribución, Margen de contribución,
    Clientes activados, Marcas activas y Concentración Top 3. NO incluye
    Inventario: en el archivo, Inventario llega como un dato único por
    Marca+Semana que se repite igual en las filas de las 4 Regiones (no es
    un valor reportado por Región), así que desglosarlo por Región mostraría
    el mismo número repetido en cada fila, dando la impresión equivocada de
    que cada Región tiene su propio inventario independiente.

    La Venta de cada Región usa el mismo resguardo de la semana "bisagra"
    que el total general: si el filtro de Semana cubre el mes completo y
    Sheet2 todavía no confirma ese Marca+Mes, la semana bisagra se excluye.
    A diferencia del total general, aquí NUNCA se sustituye por el total de
    Sheet2 (ese total es de todas las Regiones juntas, no de una sola), así
    que ninguna fila de Región puede mostrar, por sí sola, el total del mes
    completo."""
    filas = []
    for region in region_sel:
        df_r = df_region[df_region["Región"] == region]
        brand_agg_r = aggregate_by_brand(df_r)
        # Mismo resguardo de semana "bisagra" que el total general, pero
        # aplicado por Marca ANTES de calcular Top 3 / Marcas activas, para
        # que esos indicadores sean consistentes con el total de Venta de la
        # fila (si no, el Top 3 podía sumar más que el total y dar una
        # "Concentración" por encima de 100%).
        if semana_filtro_es_mes_completo and meses_sel and not brand_agg_r.empty:
            df_rb = df_region_budget[df_region_budget["Región"] == region]
            venta_dict_r = venta_real_por_marca(
                df_rb, venta_mensual_lookup or {}, all_weeks or [], brand_agg_r["Marca"], meses_sel, usar_sheet2=False,
            )
            brand_agg_r = brand_agg_r.copy()
            brand_agg_r["venta"] = brand_agg_r["Marca"].map(venta_dict_r).fillna(brand_agg_r["venta"])
        venta = brand_agg_r["venta"].sum()
        contrib = brand_agg_r["contribucion"].sum()
        margen = contrib / venta if venta else 0
        clientes = brand_agg_r["clientes"].sum()
        marcas_activas = int((brand_agg_r["venta"] > 0).sum())
        marcas_total = len(brand_agg_r)
        top3 = brand_agg_r.nlargest(3, "venta")["venta"].sum() if not brand_agg_r.empty else 0
        concentracion = top3 / venta if venta else 0
        filas.append({
            "Región": region, "venta": venta, "contribucion": contrib, "margen": margen,
            "clientes": clientes, "marcas_activas": marcas_activas, "marcas_total": marcas_total,
            "concentracion": concentracion,
        })
    return pd.DataFrame(filas).sort_values("venta", ascending=False).reset_index(drop=True)


def compute_inventario_actual(df: pd.DataFrame) -> pd.DataFrame:
    """Último nivel de Inventario reportado por marca, en USD (foto del stock
    más reciente; nunca se suma entre semanas, igual que Rotación nunca se
    suma)."""
    if "Inventario_Val" not in df.columns or df["Inventario_Val"].notna().sum() == 0:
        return pd.DataFrame({"Marca": pd.Series(dtype="object"), "inventario": pd.Series(dtype="float64")})
    sub = df[df["Inventario_Val"].notna()].sort_values("Semana")
    out = sub.groupby("Marca", as_index=False)["Inventario_Val"].last()
    return out.rename(columns={"Inventario_Val": "inventario"})


def compute_contribucion_budget_comparison(df_all: pd.DataFrame, grupos_sel, marcas_sel, weeks_mes, mes_label):
    """Presupuesto de Contribución vs. Real del/los mes(es) seleccionado(s)
    en el filtro de Mes. A diferencia del Presupuesto de Venta (que viene en
    su propia hoja), este presupuesto llega como una columna más de Sheet1
    con un único valor total por marca (no una serie semanal) — se toma el
    último valor no vacío reportado para esa marca, nunca se suma entre
    semanas. `weeks_mes` ya viene calculado (vía weeks_for_meses) incluyendo
    cualquier semana "bisagra" que toque el/los mes(es) elegido(s)."""
    if df_all.empty or "PresupContrib_Miles" not in df_all.columns or not weeks_mes:
        return None
    if df_all["PresupContrib_Miles"].notna().sum() == 0:
        return None

    reported = [w for w in weeks_mes if w in set(df_all["Semana"])]
    pct_avance = len(reported) / len(weeks_mes) if weeks_mes else 0

    # Se filtra por la lista de semanas del mes elegido (weeks_mes), no por
    # la etiqueta "Mes" de cada fila: así una semana "bisagra" que toque el
    # mes seleccionado se incluye aunque su jueves (regla ISO) caiga en el
    # mes vecino.
    df_mes = df_all[df_all["Semana"].isin(weeks_mes)]

    contrib_real = (
        df_mes.groupby("Marca", as_index=False)["Contribucion_USD"].sum()
        .rename(columns={"Contribucion_USD": "contribucion_real"})
    )

    # El presupuesto de contribución es un total por marca, no una cifra
    # semanal: se busca en TODAS las semanas cargadas (no solo el mes en
    # curso), porque el archivo puede etiquetarlo en cualquier semana del
    # rango exportado (p. ej. la primera), incluso si esa semana cae en un
    # mes distinto al mes en curso según la regla ISO del jueves. Si el
    # archivo trae varias filas por Marca+Semana (una por Región), el total
    # de esa semana viene repartido entre ellas, así que primero se SUMAN
    # las Regiones de cada Marca+Semana antes de tomar el último valor
    # (chronológicamente) por marca.
    presu_rows = df_all[df_all["PresupContrib_Miles"].notna()]
    if presu_rows.empty:
        return None
    presu_semana = (
        presu_rows.groupby(["Marca", "Semana"], as_index=False)["PresupContrib_Miles"].sum()
        .sort_values("Semana")
    )
    presu_contrib = presu_semana.groupby("Marca", as_index=False)["PresupContrib_Miles"].last()
    presu_contrib["presupuesto_contrib"] = presu_contrib["PresupContrib_Miles"] * 1000

    marca_grupo = df_all.drop_duplicates("Marca").set_index("Marca")["Grupo"]
    presu_contrib["Grupo"] = presu_contrib["Marca"].map(marca_grupo)
    presu_contrib = presu_contrib[
        presu_contrib["Marca"].isin(marcas_sel) & presu_contrib["Grupo"].isin(grupos_sel)
    ]
    if presu_contrib.empty:
        return None

    rows = presu_contrib.merge(contrib_real, on="Marca", how="left")
    rows["contribucion_real"] = rows["contribucion_real"].fillna(0)
    rows["cumplimiento"] = np.where(
        rows["presupuesto_contrib"] != 0, rows["contribucion_real"] / rows["presupuesto_contrib"], np.nan
    )
    rows["ritmo"] = rows["cumplimiento"] / pct_avance if pct_avance > 0 else np.nan
    rows["estado"] = rows["ritmo"].apply(estado_ritmo)
    rows = rows[["Marca", "Grupo", "presupuesto_contrib", "contribucion_real", "cumplimiento", "ritmo", "estado"]]

    total_presu = rows["presupuesto_contrib"].sum()
    total_real = rows["contribucion_real"].sum()
    total_cumpl = total_real / total_presu if total_presu else 0
    total_ritmo = total_cumpl / pct_avance if pct_avance > 0 else 0

    return {
        "mes_label": mes_label,
        "weeks_mes": weeks_mes,
        "reported": reported,
        "pct_avance": pct_avance,
        "rows": rows.sort_values("presupuesto_contrib", ascending=False).reset_index(drop=True),
        "total_presu": total_presu,
        "total_real": total_real,
        "total_cumpl": total_cumpl,
        "total_ritmo": total_ritmo,
    }


def compute_budget_comparison(df_all: pd.DataFrame, presu: pd.DataFrame, grupos_sel, marcas_sel, meses_sel, weeks_mes, mes_label, venta_mensual_lookup=None, all_weeks=None):
    """Presupuesto vs. Real del/los mes(es) seleccionado(s) en el filtro de
    Mes (ignora el filtro de Semana, pero sí respeta Grupo y Marca), igual
    que en el HTML/Excel originales. Si se seleccionan varios meses, el
    presupuesto se suma por marca entre esos meses."""
    if df_all.empty or presu is None or presu.empty or not weeks_mes:
        return None

    # La Venta Real de cada Marca se calcula mes a mes con
    # `venta_real_por_marca`: usa el total oficial de la hoja "Sheet2"
    # cuando está disponible (ya corrige correctamente cualquier semana
    # "bisagra" compartida con el mes vecino), y si no, cae de respaldo en
    # sumar "Venta_USD" únicamente de las semanas cargadas que tocan
    # EXCLUSIVAMENTE ese mes (la semana bisagra se deja afuera hasta que
    # Sheet2 confirme el mes).
    marcas_presentes = sorted(df_all["Marca"].unique())
    venta_dict = venta_real_por_marca(df_all, venta_mensual_lookup or {}, all_weeks or weeks_mes, marcas_presentes, meses_sel)
    venta_mes = pd.DataFrame({"Marca": list(venta_dict.keys()), "venta_real": list(venta_dict.values())})

    # "Semanas reportadas" / "% de avance": igual criterio que la Venta Real
    # de arriba, para que no muestren mensajes contradictorios (p. ej. "100%
    # de avance" con "Venta Real $0"). Si Sheet2 ya confirma el total de
    # alguna Marca para este/estos mes(es), el mes se considera resuelto
    # (cualquier semana bisagra ya quedó repartida correctamente dentro de
    # ese total) y cuenta el avance normal por semanas cargadas. Si Sheet2
    # todavía no trae nada de este mes, una semana "bisagra" no cuenta como
    # reportada hasta que Sheet2 la confirme (mismo resguardo que
    # `venta_real_marca_mes`).
    sheet2_confirma_mes = bool(venta_mensual_lookup) and any(
        (marca, y, m) in venta_mensual_lookup for marca in marcas_presentes for (y, m) in meses_sel
    )
    if sheet2_confirma_mes:
        reported = [w for w in weeks_mes if w in set(df_all["Semana"])]
    else:
        semanas_exclusivas = {w for w in weeks_mes if meses_tocados_por_semana(w) <= set(meses_sel)}
        reported = [w for w in weeks_mes if w in semanas_exclusivas and w in set(df_all["Semana"])]
    pct_avance = len(reported) / len(weeks_mes) if weeks_mes else 0

    periodos_sel = {pd.Timestamp(year=y, month=m, day=1).to_period("M") for (y, m) in meses_sel}
    presu_mes = presu[presu["Mes_Fecha"].dt.to_period("M").isin(periodos_sel)]
    presu_mes = presu_mes[presu_mes["Grupo Compra"].isin(grupos_sel) & presu_mes["Marca"].isin(marcas_sel)]
    if presu_mes.empty:
        return None
    # Si se seleccionó más de un mes, se suma el presupuesto de cada marca
    # entre esos meses (nunca se promedia ni se toma solo el último).
    presu_mes = presu_mes.groupby(["Marca", "Grupo Compra"], as_index=False)["Presupuesto_Miles"].sum()
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


def order_acciones_by_budget(acciones_full: pd.DataFrame, presu: pd.DataFrame, df_all: pd.DataFrame, meses_sel):
    """Ordena la tabla de Acciones de mayor a menor presupuesto del/los
    mes(es) seleccionado(s) en el filtro de Mes; las marcas sin presupuesto
    quedan al final, en orden alfabético. Si hay varios meses seleccionados,
    se suma el presupuesto de cada marca entre esos meses."""
    budget_map = {}
    if presu is not None and not presu.empty and not df_all.empty and meses_sel:
        periodos_sel = {pd.Timestamp(year=y, month=m, day=1).to_period("M") for (y, m) in meses_sel}
        presu_mes = presu[presu["Mes_Fecha"].dt.to_period("M").isin(periodos_sel)]
        budget_map = (presu_mes.groupby("Marca")["Presupuesto_Miles"].sum() * 1000).to_dict()

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


def multi_line_chart(df: pd.DataFrame, x_col: str, y_col: str, series_col: str, fmt=fmt_usd, title=None):
    """Gráfico de líneas con una serie (línea) por cada valor distinto de
    `series_col` (p. ej. una línea por Región), todas compartiendo el mismo
    eje de Semana — para comparar su evolución semana a semana. Reutiliza la
    misma paleta de colores que el resto del dashboard, repitiéndola si hay
    más series que colores."""
    palette = [BLUE, ORANGE, AQUA, VIOLET, RED, MUTED]
    fig = go.Figure()
    for i, (serie, sub) in enumerate(df.groupby(series_col)):
        sub = sub.sort_values(x_col)
        color = palette[i % len(palette)]
        fig.add_trace(go.Scatter(
            x=sub[x_col], y=sub[y_col], mode="lines+markers",
            name=str(serie), line=dict(color=color, width=2.5), marker=dict(color=color, size=6),
            hovertemplate="%{x}<br>%{y:,.0f}<extra>" + str(serie) + "</extra>",
        ))
    fig.update_layout(
        title=title,
        template="plotly_white",
        margin=dict(l=10, r=10, t=40 if title else 10, b=50),
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=True, gridcolor=GRID, zeroline=False, tickformat="~s"),
        height=400,
        showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5),
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


def render_kpi_view(brand_agg: pd.DataFrame, week_agg: pd.DataFrame, grupo_agg: pd.DataFrame, top_n: int, region_breakdown: pd.DataFrame = None, week_region_agg: pd.DataFrame = None):
    """Tarjetas KPI + gráficos Top-N + evolución semanal + venta por grupo.
    Se reutiliza tanto en la vista '📈 KPI' (todas las Regiones) como en
    '🗺️ KPI por Región' (Región(es) seleccionada(s)); solo cambian las
    tablas ya agregadas que se le pasan. `region_breakdown` (opcional, solo
    en la vista por Región) agrega una tabla con Venta/Contribución/etc. por
    cada Región seleccionada, debajo de las tarjetas de totales.
    `week_region_agg` (opcional, solo en la vista por Región) agrega un
    gráfico de evolución semanal con una línea por Región, para comparar su
    avance semana a semana (no solo el total agregado del mes)."""
    total_venta = brand_agg["venta"].sum()
    total_contrib = brand_agg["contribucion"].sum()
    margen = total_contrib / total_venta if total_venta else 0
    total_clientes = brand_agg["clientes"].sum()
    marcas_activas = (brand_agg["venta"] > 0).sum()
    top3 = brand_agg.nlargest(3, "venta")["venta"].sum()
    conc = top3 / total_venta if total_venta else 0

    tiene_inventario = "inventario" in brand_agg.columns and brand_agg["inventario"].notna().any()
    total_inventario = brand_agg["inventario"].sum(skipna=True) if tiene_inventario else 0

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Venta total (USD)", f"${total_venta:,.0f}")
    k2.metric("Contribución total (USD)", f"${total_contrib:,.0f}")
    k3.metric("Margen de contribución", f"{margen:.1%}")
    k4.metric("Clientes activados", f"{total_clientes:,.0f}")
    k5.metric("Marcas activas", f"{marcas_activas} / {len(brand_agg)}")
    k6.metric("Concentración Top 3 marcas", f"{conc:.1%}")

    if tiene_inventario:
        k7, _, _ = st.columns(3)
        k7.metric("Inventario actual (USD)", f"${total_inventario:,.0f}")

    if region_breakdown is not None and not region_breakdown.empty:
        st.markdown("**Indicadores por Región**")
        tabla_r = region_breakdown.copy()
        tabla_r["marcas_activas_total"] = tabla_r.apply(
            lambda r: f"{int(r['marcas_activas'])} / {int(r['marcas_total'])}", axis=1
        )
        tabla_r["venta"] = tabla_r["venta"].map(lambda v: f"${v:,.0f}")
        tabla_r["contribucion"] = tabla_r["contribucion"].map(lambda v: f"${v:,.0f}")
        tabla_r["margen"] = tabla_r["margen"].map(lambda v: f"{v:.1%}")
        tabla_r["clientes"] = tabla_r["clientes"].map(lambda v: f"{v:,.0f}")
        tabla_r["concentracion"] = tabla_r["concentracion"].map(lambda v: f"{v:.1%}")
        cols = ["Región", "venta", "contribucion", "margen", "clientes", "marcas_activas_total", "concentracion"]
        nombres = ["Región", "Venta (USD)", "Contribución (USD)", "Margen de contribución", "Clientes activados", "Marcas activas", "Concentración Top 3"]
        tabla_r = tabla_r[cols]
        tabla_r.columns = nombres
        st.dataframe(tabla_r, width='stretch', hide_index=True)
        st.caption(
            "La Venta de cada Región usa el mismo resguardo de la semana \"bisagra\" que el "
            "total de arriba, pero nunca se sustituye por el total de Sheet2 (ese total es de "
            "todas las Regiones juntas, no de una sola), así que ninguna fila por sí sola refleja "
            "el total exacto del mes — para eso está el total combinado de arriba. No se incluye "
            "Inventario porque en el archivo llega como un dato único por Marca+Semana (no por "
            "Región), igual que Rotación."
        )

    if brand_agg.empty:
        st.info("No hay datos para los filtros seleccionados.")
        return

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

    if tiene_inventario:
        top_inv = brand_agg[brand_agg["inventario"].notna()].nlargest(top_n, "inventario")
        st.plotly_chart(
            hbar_chart(top_inv["Marca"], top_inv["inventario"], MUTED, fmt=fmt_usd, title=f"Top {len(top_inv)} marcas por inventario (USD, nivel más reciente)"),
            width='stretch',
        )

    cc3, cc4 = st.columns(2)
    with cc3:
        st.plotly_chart(line_chart(week_agg["Semana"], week_agg["venta"], BLUE, title="Evolución semanal de venta (USD)"), width='stretch')
    with cc4:
        st.plotly_chart(vbar_chart(week_agg["Semana"], week_agg["clientes"], VIOLET, title="Clientes activados por semana (total)"), width='stretch')

    if week_region_agg is not None and not week_region_agg.empty:
        st.plotly_chart(
            multi_line_chart(week_region_agg, "Semana", "venta", "Región", title="Evolución semanal de venta por Región (USD)"),
            width='stretch',
        )

    st.plotly_chart(vbar_chart(grupo_agg["Grupo"], grupo_agg["venta"], BLUE, title="Venta por grupo de compra (USD)"), width='stretch')


# ----------------------------------------------------------------------
# Estado de sesión: datos activos (por defecto o reemplazados por upload)
# ----------------------------------------------------------------------
if "raw" not in st.session_state:
    raw0, presu0, acciones0, venta_mensual0 = load_default_data()
    st.session_state.raw = raw0
    st.session_state.presu = presu0
    st.session_state.acciones_hits = acciones0
    st.session_state.venta_mensual = venta_mensual0
    st.session_state.data_label = "Ventas actualizadas.xlsx (datos incluidos por defecto)"

with st.sidebar:
    st.header("📤 Actualizar datos")
    st.caption("Carga ventas actualizadas")
    uploaded = st.file_uploader("Archivo .xlsx", type=["xlsx"], label_visibility="collapsed")
    if uploaded is not None:
        try:
            raw_new, presu_new, acciones_new, venta_mensual_new, avisos = parse_uploaded_workbook(uploaded.read())
            st.session_state.raw = raw_new
            if presu_new is not None:
                st.session_state.presu = presu_new
            if acciones_new is not None:
                st.session_state.acciones_hits = acciones_new
            if venta_mensual_new is not None:
                st.session_state.venta_mensual = venta_mensual_new
            st.session_state.data_label = uploaded.name
            st.success(f"✓ Datos actualizados desde “{uploaded.name}” — {len(raw_new)} filas.")
            for a in avisos:
                st.info(a)
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")

    st.divider()
    st.caption(
        "El presupuesto, las acciones y la venta mensual (hoja Sheet2) se "
        "refrescan solo si el archivo subido trae esas hojas; de lo "
        "contrario se conservan los últimos valores cargados."
    )

raw = derive_raw(st.session_state.raw)
presu = st.session_state.presu
acciones_hits = st.session_state.acciones_hits
venta_mensual = st.session_state.get("venta_mensual")

# Diccionario {(Marca, year, month): venta_completa_USD} a partir de Sheet2,
# si se cargó. Se asume la misma máscara que "Venta Neta" en Sheet1 (miles de
# USD) -> se multiplica x1000. Si no hay Sheet2 todavía, queda vacío y toda
# "Venta" usa la regla de respaldo (semana bisagra completa en el mes
# elegido), igual que antes de este cambio.
VENTA_MENSUAL_LOOKUP = {}
if venta_mensual is not None and not venta_mensual.empty:
    for _, _row in venta_mensual.iterrows():
        VENTA_MENSUAL_LOOKUP[(_row["Marca"], int(_row["Mes_Year"]), int(_row["Mes_Month"]))] = (
            pd.to_numeric(_row["VentaCompleta_Miles"], errors="coerce") * 1000
        )

ALL_WEEKS = sorted(raw["Semana"].unique())
ALL_GRUPOS = sorted(raw["Grupo"].unique())
ALL_BRANDS = sorted(raw["Marca"].unique())
ALL_REGIONES = sorted(raw["Región"].unique())
BRAND_GRUPO = dict(zip(raw["Marca"], raw["Grupo"]))

acciones_full = build_full_acciones(acciones_hits, ALL_BRANDS)

# Meses disponibles según las semanas cargadas (una semana "bisagra" entre
# dos meses hace que ambos aparezcan como opción), y el mes que viene
# preseleccionado por defecto (el mes real de hoy, o el de la última semana
# cargada si los datos son más viejos que hoy).
MESES_DISPONIBLES = meses_disponibles(ALL_WEEKS)
MES_LABELS = {t: mes_label_de(*t) for t in MESES_DISPONIBLES}
MES_LABEL_TO_TUPLE = {v: k for k, v in MES_LABELS.items()}
default_mes_tuple = mes_en_curso_default(ALL_WEEKS)
if default_mes_tuple not in MES_LABELS and MESES_DISPONIBLES:
    default_mes_tuple = MESES_DISPONIBLES[-1]

# ----------------------------------------------------------------------
# Encabezado + filtros
# ----------------------------------------------------------------------
header_izq, header_der = st.columns([5, 1])
with header_izq:
    st.title("Ventas de marcas Febeca")
with header_der:
    st.markdown(
        "<div style='text-align:right; padding-top:22px; color:#888; "
        "font-style:italic; font-size:0.95rem; font-weight:bold;'>By MARS</div>",
        unsafe_allow_html=True,
    )
st.caption(
    f"Venta, Activación de clientes, Rotación y Contribución por marca · "
    f"Semanas {ALL_WEEKS[0]} a {ALL_WEEKS[-1]} · Fuente: {st.session_state.data_label}"
)

f0, f1, f2, f3, f4 = st.columns([1.3, 1.3, 1, 1.3, 0.8])
with f0:
    meses_labels_sel = st.multiselect(
        "Mes", [MES_LABELS[t] for t in MESES_DISPONIBLES],
        default=[MES_LABELS[default_mes_tuple]] if default_mes_tuple in MES_LABELS else [],
        help='Si una semana cae en dos meses (p. ej. empieza en septiembre y '
             'termina en octubre), aparece en el mes que elijas aquí.',
    )
with f1:
    weeks_sel = st.multiselect("Semana", ALL_WEEKS, default=ALL_WEEKS)
with f2:
    grupos_sel = st.multiselect("Grupo de compra", ALL_GRUPOS, default=ALL_GRUPOS)
with f3:
    brands_sel = st.multiselect("Marca", ALL_BRANDS, default=ALL_BRANDS)
with f4:
    top_n = st.selectbox("Top N", [5, 10, 15, len(ALL_BRANDS)], index=1)

if not meses_labels_sel or not weeks_sel or not grupos_sel or not brands_sel:
    st.warning("Selecciona al menos un Mes, una Semana, un Grupo de Compra y una Marca.")
    st.stop()

meses_sel = sorted(MES_LABEL_TO_TUPLE[l] for l in meses_labels_sel)
# Semanas que tocan el/los mes(es) elegido(s) (incluye semanas "bisagra").
weeks_mes_filtro = weeks_for_meses(ALL_WEEKS, meses_sel)
mes_label_sel = " + ".join(MES_LABELS[t] for t in meses_sel)

# KPI, tablas y gráficos: respetan Mes, Semana, Grupo y Marca (Semana permite
# acotar más fino dentro del/los mes(es) elegido(s)).
df_filtered = raw[
    raw["Semana"].isin(weeks_sel)
    & raw["Semana"].isin(weeks_mes_filtro)
    & raw["Grupo"].isin(grupos_sel)
    & raw["Marca"].isin(brands_sel)
]
# Presupuesto vs. Real ignora el filtro de Semana (mira siempre el/los
# mes(es) elegido(s) completos), pero sí respeta Mes, Grupo y Marca.
df_for_budget = raw[raw["Grupo"].isin(grupos_sel) & raw["Marca"].isin(brands_sel)]

brand_agg = aggregate_by_brand(df_filtered)
week_agg = aggregate_by_week(df_filtered)
grupo_agg = aggregate_by_grupo(df_filtered)
# Inventario es una foto del nivel actual: se calcula sobre Grupo+Marca
# filtrados pero ignorando el filtro de Semana (igual criterio que
# Presupuesto vs. Real), y se agrega a la tabla por marca.
brand_agg = brand_agg.merge(compute_inventario_actual(df_for_budget), on="Marca", how="left")

# Corrección de "Venta" con la hoja "Sheet2" (venta completa por Marca y
# Mes): solo se aplica cuando el filtro de Semana no excluye manualmente
# ninguna semana del/los mes(es) elegido(s) -- si Semana acota a un
# subconjunto más fino que el mes completo, Sheet2 no permite aislar esa
# porción y se deja la suma cruda de Sheet1. La vista "KPI por Región"
# (brand_agg_region) aplica la misma corrección más abajo, con la salvedad
# de que solo usa el total de Sheet2 cuando están seleccionadas TODAS las
# regiones (Sheet2 no trae desglose por Región); si se filtra a un
# subconjunto de regiones, igual excluye la semana "bisagra" no confirmada
# para no sobre-contarla, aunque sin la precisión exacta de Sheet2.
semana_filtro_es_mes_completo = set(weeks_mes_filtro) <= set(weeks_sel)
if semana_filtro_es_mes_completo and VENTA_MENSUAL_LOOKUP:
    _venta_dict_global = venta_real_por_marca(df_for_budget, VENTA_MENSUAL_LOOKUP, ALL_WEEKS, brand_agg["Marca"], meses_sel)
    brand_agg["venta"] = brand_agg["Marca"].map(_venta_dict_global).fillna(brand_agg["venta"])
    _total_venta_corr = brand_agg["venta"].sum()
    brand_agg["participacion"] = np.where(_total_venta_corr != 0, brand_agg["venta"] / _total_venta_corr, 0)
    brand_agg["margen"] = np.where(brand_agg["venta"] != 0, brand_agg["contribucion"] / brand_agg["venta"], 0)

# ----------------------------------------------------------------------
# Selector de vista: Presupuesto / KPI
# ----------------------------------------------------------------------
view = st.segmented_control(
    "Vista", ["💰 Presupuesto", "📈 KPI", "🗺️ KPI por Región"], default="💰 Presupuesto", label_visibility="collapsed"
)
if view is None:
    view = "💰 Presupuesto"

if view == "💰 Presupuesto":
    st.subheader("Presupuesto de Venta vs. Real")
    budget = compute_budget_comparison(
        df_for_budget, presu, grupos_sel, brands_sel, meses_sel, weeks_mes_filtro, mes_label_sel,
        venta_mensual_lookup=VENTA_MENSUAL_LOOKUP, all_weeks=ALL_WEEKS,
    )
    if budget is None:
        st.info(
            "No hay marcas con presupuesto asignado para los filtros de Mes/Grupo/Marca "
            "seleccionados, o no se cargó una hoja de Presupuesto todavía."
        )
    else:
        st.caption(
            f"Mes(es) seleccionado(s): **{budget['mes_label']}** — compara siempre el/los "
            "mes(es) completo(s) (no responde al filtro de Semana), pero sí al filtro de "
            "Mes, Grupo de Compra y Marca. "
            f"Semanas reportadas: {len(budget['reported'])} de {len(budget['weeks_mes'])} "
            f"({budget['pct_avance']:.1%} de avance)."
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
                fmt=fmt_pct, title=f"% Cumplimiento de presupuesto por marca ({mes_label_sel})",
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

    st.divider()
    st.subheader("Presupuesto de Contribución vs. Real")
    budget_c = compute_contribucion_budget_comparison(df_for_budget, grupos_sel, brands_sel, weeks_mes_filtro, mes_label_sel)
    if budget_c is None:
        st.info(
            "No hay marcas con presupuesto de contribución asignado para los filtros "
            'seleccionados, o el archivo cargado no trae la columna "Presupuesto de '
            'contribución" en Sheet1.'
        )
    else:
        st.caption(
            f"Mes(es) seleccionado(s): **{budget_c['mes_label']}** — mismo criterio que el "
            "presupuesto de venta: compara siempre el/los mes(es) completo(s) (no responde "
            "al filtro de Semana), pero sí al filtro de Mes, Grupo de Compra y Marca. El "
            "presupuesto de contribución es un total por marca (no una serie semanal), así que nunca "
            "se suma entre semanas."
        )
        d1, d2, d3, d4, d5 = st.columns(5)
        d1.metric(f"Presup. Contribución {budget_c['mes_label']} (USD)", f"${budget_c['total_presu']:,.0f}")
        d2.metric(f"Contribución Real {budget_c['mes_label']} (USD)", f"${budget_c['total_real']:,.0f}")
        d3.metric("% Cumplimiento presupuesto", f"{budget_c['total_cumpl']:.1%}")
        d4.metric("Avance del mes (semanas)", f"{budget_c['pct_avance']:.1%}")
        d5.metric("Ritmo vs. calendario", f"{budget_c['total_ritmo']:.1%}")

        rows_sorted_c = budget_c["rows"].sort_values("cumplimiento", ascending=True)
        st.plotly_chart(
            hbar_chart(
                rows_sorted_c["Marca"], rows_sorted_c["cumplimiento"], VIOLET,
                fmt=fmt_pct, title=f"% Cumplimiento de presupuesto de contribución por marca ({mes_label_sel})",
            ),
            width='stretch',
        )

        with st.expander("Ver tabla de presupuesto de contribución por marca"):
            tabla_c = budget_c["rows"].copy()
            tabla_c["presupuesto_contrib"] = tabla_c["presupuesto_contrib"].map(lambda v: f"${v:,.0f}")
            tabla_c["contribucion_real"] = tabla_c["contribucion_real"].map(lambda v: f"${v:,.0f}")
            tabla_c["cumplimiento"] = tabla_c["cumplimiento"].map(lambda v: f"{v:.1%}" if pd.notna(v) else "-")
            tabla_c["ritmo"] = tabla_c["ritmo"].map(lambda v: f"{v:.1%}" if pd.notna(v) else "-")
            tabla_c.columns = ["Marca", "Grupo", "Presupuesto Contribución (USD)", "Contribución Real (USD)", "% Cumplimiento", "Ritmo vs. calendario", "Estado"]
            st.dataframe(tabla_c, width='stretch', hide_index=True)

elif view == "📈 KPI":
    st.subheader("Indicadores y KPI")
    render_kpi_view(brand_agg, week_agg, grupo_agg, top_n)

else:
    st.subheader("Indicadores y KPI por Región")
    region_col, _ = st.columns([1.3, 3.7])
    with region_col:
        region_sel = st.multiselect(
            "Región", ALL_REGIONES, default=ALL_REGIONES,
            help="Puedes ver una Región sola o combinar varias; los KPI se suman entre las Regiones seleccionadas.",
        )
    if not region_sel:
        st.warning("Selecciona al menos una Región.")
    else:
        df_region = df_filtered[df_filtered["Región"].isin(region_sel)]
        df_region_budget = df_for_budget[df_for_budget["Región"].isin(region_sel)]
        brand_agg_region = aggregate_by_brand(df_region)
        brand_agg_region = brand_agg_region.merge(
            compute_inventario_actual(df_region_budget), on="Marca", how="left"
        )
        week_agg_region = aggregate_by_week(df_region)
        grupo_agg_region = aggregate_by_grupo(df_region)

        # Misma corrección de "Venta" que la vista KPI principal. El total de
        # Sheet2 es de TODAS las regiones juntas, así que solo se usa
        # directamente cuando region_sel cubre todas las regiones (equivale
        # al universo completo); si es un subconjunto, se excluye igual la
        # semana "bisagra" no confirmada (sin poder aislar el monto exacto
        # de Sheet2 para ese subconjunto de regiones).
        region_es_todas = len(region_sel) == len(ALL_REGIONES)
        if semana_filtro_es_mes_completo and (not region_es_todas or VENTA_MENSUAL_LOOKUP):
            _venta_dict_region = venta_real_por_marca(
                df_region_budget, VENTA_MENSUAL_LOOKUP, ALL_WEEKS, brand_agg_region["Marca"], meses_sel,
                usar_sheet2=region_es_todas,
            )
            brand_agg_region["venta"] = brand_agg_region["Marca"].map(_venta_dict_region).fillna(brand_agg_region["venta"])
            _total_venta_corr_region = brand_agg_region["venta"].sum()
            brand_agg_region["participacion"] = np.where(
                _total_venta_corr_region != 0, brand_agg_region["venta"] / _total_venta_corr_region, 0
            )
            brand_agg_region["margen"] = np.where(
                brand_agg_region["venta"] != 0, brand_agg_region["contribucion"] / brand_agg_region["venta"], 0
            )

        region_label = "Todas las regiones" if len(region_sel) == len(ALL_REGIONES) else " + ".join(region_sel)
        st.caption(
            f"Región(es) seleccionada(s): **{region_label}** — respeta también los filtros de "
            "Mes, Semana, Grupo de Compra y Marca de arriba."
        )
        region_breakdown = kpi_por_region_tabla(
            df_region, df_region_budget, region_sel,
            all_weeks=ALL_WEEKS, meses_sel=meses_sel, venta_mensual_lookup=VENTA_MENSUAL_LOOKUP,
            semana_filtro_es_mes_completo=semana_filtro_es_mes_completo,
        )
        week_region_agg = aggregate_by_week_region(df_region)
        render_kpi_view(
            brand_agg_region, week_agg_region, grupo_agg_region, top_n,
            region_breakdown=region_breakdown, week_region_agg=week_region_agg,
        )

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
    tabla["inventario"] = (
        brand_agg["inventario"].map(lambda v: f"${v:,.0f}" if pd.notna(v) else "-")
        if "inventario" in brand_agg.columns else "-"
    )
    tabla = tabla[["Marca", "Grupo", "venta", "participacion", "contribucion", "margen", "clientes", "rotacion", "inventario"]]
    tabla.columns = ["Marca", "Grupo", "Venta (USD)", "% Participación", "Contribución (USD)", "Margen %", "Clientes Activados", "Rotación", "Inventario (USD)"]
    st.dataframe(tabla, width='stretch', hide_index=True)

with st.expander("Ver rotación promedio por mes y marca (matriz)"):
    st.caption(
        'Promedio simple de "Rotación" por marca dentro de cada mes '
        "(nunca se suma)."
    )
    rot_semanal = collapse_region_rotacion(df_filtered)
    pivot = rot_semanal.pivot_table(index="Marca", columns="Mes", values="Rotacion", aggfunc="mean")
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
        f"Ordenada de mayor a menor presupuesto de {mes_label_sel} (las marcas sin "
        "presupuesto quedan al final, alfabético). Respeta los filtros de Mes, Grupo "
        "de Compra y Marca de arriba, no el de Semana."
    )
    filtro_marca = st.text_input("Buscar marca…", key="acc_f_marca")
    ordered = order_acciones_by_budget(acciones_full, presu, df_for_budget, meses_sel)
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
with st.expander("Notas", expanded=False):
    st.caption(
        '**Notas:** "Venta" corresponde a la columna "Venta Neta" del archivo original '
        "(máscara miles de USD, convertida a USD); \"Contribución\" también se convirtió "
        'de miles de USD a USD. "Clientes Activados" se SUMA por marca/grupo/semana; '
        '"Rotación" siempre se PROMEDIA, nunca se suma — se reporta mayormente en semanas '
        "sin venta y debe interpretarse con cautela. \"Grupo de Compra\" es un atributo "
        'propio de cada marca. La sección "Presupuesto vs. Real" compara siempre el '
        "mes (o meses) completos elegidos en el filtro de Mes y solo incluye marcas "
        'con presupuesto asignado; "Ritmo" compara el % de presupuesto alcanzado '
        "contra el % de semanas ya transcurridas del mes. Una semana que cae en dos "
        "meses (bisagra): si el archivo trae la hoja \"Sheet2\" (venta completa por "
        "Marca y Mes), el monto de \"Venta\" de esa semana se toma del total oficial "
        "de Sheet2 para el mes que selecciones — por eso puede ser distinto al filtrar "
        "Agosto que al filtrar Septiembre; si no hay Sheet2 para ese mes, se cuenta la "
        "semana completa en el mes que selecciones (nunca en los dos a la vez). Las "
        "marcas sin ninguna acción registrada se muestran con la celda de Acción "
        'vacía. "Inventario" también viene en miles de USD (convertido a USD) y es el '
        "último nivel reportado por marca (nunca se suma entre semanas). "
        '"Presupuesto de Contribución" es un total por marca '
        "(no semanal) y se compara contra la Contribución real del mes elegido con "
        "el mismo criterio que el Presupuesto de Venta."
    )
