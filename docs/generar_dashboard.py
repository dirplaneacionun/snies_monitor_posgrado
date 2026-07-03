#!/usr/bin/env python3
"""
generar_dashboard.py
Genera el dashboard HTML para SNIES Posgrado.

Estructura y lógica idénticas al dashboard de pregrado (mismo CSS, mismo
JS, mismos filtros e interacciones) — sólo cambian los textos ("Posgrado"
en vez de "Pregrado"), las rutas de datos y el patrón de snapshots.

Produce:
  docs/index.html
  docs/nuevos.html
  docs/inactivos.html
  docs/modificados.html
  docs/modificados_creditos.html
  docs/modificados_costos.html

Ejecutar desde la raíz del repositorio:
  python docs/generar_dashboard.py
"""
import json
import glob
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent.parent
DOCS = Path(__file__).parent
NOVEDADES = ROOT / "data" / "novedades"
PROGRAMAS = ROOT / "Programas"

# ── Columnas ───────────────────────────────────────────────────────────────────

COLS_DETAIL = [
    "FECHA_OBTENCION", "CÓDIGO_SNIES_DEL_PROGRAMA", "NOMBRE_DEL_PROGRAMA",
    "NOMBRE_INSTITUCIÓN", "SECTOR", "MODALIDAD", "DEPARTAMENTO_OFERTA_PROGRAMA",
    "MUNICIPIO_OFERTA_PROGRAMA", "NÚMERO_CRÉDITOS", "COSTO_MATRÍCULA_ESTUD_NUEVOS",
    "PERIODICIDAD", "FECHA_DE_REGISTRO_EN_SNIES", "DIVISIÓN UNINORTE",
    "CINE_F_2013_AC_CAMPO_ESPECÍFIC", "NÚMERO_PERIODOS_DE_DURACIÓN",
    "NIVEL_DE_FORMACIÓN",
]

COLS_MOD_DETAIL = [
    "FECHA_OBTENCION", "CÓDIGO_SNIES_DEL_PROGRAMA", "NOMBRE_DEL_PROGRAMA",
    "NOMBRE_INSTITUCIÓN", "SECTOR", "MODALIDAD", "DEPARTAMENTO_OFERTA_PROGRAMA",
    "MUNICIPIO_OFERTA_PROGRAMA", "NÚMERO_CRÉDITOS", "COSTO_MATRÍCULA_ESTUD_NUEVOS",
    "PERIODICIDAD", "FECHA_DE_REGISTRO_EN_SNIES", "DIVISIÓN UNINORTE",
    "CINE_F_2013_AC_CAMPO_ESPECÍFIC", "NÚMERO_PERIODOS_DE_DURACIÓN",
    "NIVEL_DE_FORMACIÓN", "QUE_CAMBIO", "NÚMERO_CRÉDITOS_ANTERIOR",
]

SNAPSHOT_COLS = [
    "CÓDIGO_SNIES_DEL_PROGRAMA", "NOMBRE_DEL_PROGRAMA", "NOMBRE_INSTITUCIÓN",
    "SECTOR", "MODALIDAD", "DEPARTAMENTO_OFERTA_PROGRAMA",
    "NÚMERO_PERIODOS_DE_DURACIÓN", "PERIODICIDAD",
]

COLS_IDX_PREVIEW = [
    "FECHA_OBTENCION", "CÓDIGO_SNIES_DEL_PROGRAMA", "NOMBRE_DEL_PROGRAMA",
    "NOMBRE_INSTITUCIÓN", "SECTOR", "MODALIDAD", "DEPARTAMENTO_OFERTA_PROGRAMA",
    "DIVISIÓN UNINORTE", "NÚMERO_PERIODOS_DE_DURACIÓN",
    "CINE_F_2013_AC_CAMPO_ESPECÍFIC",
]
COLS_IDX_PREVIEW_MOD = [
    "FECHA_OBTENCION", "CÓDIGO_SNIES_DEL_PROGRAMA", "NOMBRE_DEL_PROGRAMA",
    "NOMBRE_INSTITUCIÓN", "SECTOR", "MODALIDAD", "DEPARTAMENTO_OFERTA_PROGRAMA",
    "QUE_CAMBIO",
]

# Columnas del universo enriquecido usado por las páginas de créditos/costos
COLS_UNIVERSO = [
    "FECHA_OBTENCION", "CÓDIGO_SNIES_DEL_PROGRAMA", "NOMBRE_DEL_PROGRAMA",
    "NOMBRE_INSTITUCIÓN", "SECTOR", "DEPARTAMENTO_OFERTA_PROGRAMA",
    "DIVISIÓN UNINORTE", "CINE_F_2013_AC_CAMPO_ESPECÍFIC", "QUE_CAMBIO",
]

# ── Helpers ────────────────────────────────────────────────────────────────────

def _to_records(df, cols):
    available = [c for c in cols if c in df.columns]
    df2 = df[available].copy()
    records = []
    for _, row in df2.iterrows():
        rec = {}
        for c in available:
            v = row[c]
            if pd.isna(v):
                rec[c] = ""
            elif isinstance(v, float) and v == int(v):
                rec[c] = str(int(v))
            else:
                rec[c] = str(v)
        records.append(rec)
    return records


def _count_last_run(df):
    if df.empty or "FECHA_OBTENCION" not in df.columns:
        return 0
    ultima = df["FECHA_OBTENCION"].iloc[-1] if not df.empty else None
    fechas_validas = df["FECHA_OBTENCION"].dropna()
    if fechas_validas.empty:
        return 0
    ultima = sorted(fechas_validas.unique(), key=_parse_fecha)[-1]
    return int((df["FECHA_OBTENCION"] == ultima).sum())


def _parse_fecha(s):
    try:
        d, m, y = str(s).split("/")
        return datetime(int(y), int(m), int(d))
    except (ValueError, AttributeError):
        return datetime(1970, 1, 1)


def _num(v):
    try:
        s = str(v).replace(",", "").strip()
        if s in ("", "nan", "None"):
            return None
        return float(s)
    except (ValueError, TypeError):
        return None


def _cambia(row, col_new, col_old):
    if col_new not in row.index or col_old not in row.index:
        return None
    vn, va = row[col_new], row[col_old]
    if pd.isna(vn) or pd.isna(va):
        return None
    return str(vn).strip() != str(va).strip()


def leer_historico():
    """Cuenta el total de programas activos por snapshot en Programas/, y
    de paso el desglose por modalidad y el detalle sector×departamento×
    modalidad de cada fila (para el gráfico de evolución por modalidad y el
    cross-tab histórico que alimenta el filtro global de index.html)."""
    snaps = glob.glob(str(PROGRAMAS / "Programas postgrado *.xlsx"))
    rows = []
    for s in snaps:
        m = re.search(r"(\d{2}-\d{2}-\d{2,4})", s)
        if not m:
            continue
        raw = m.group(1)
        dd, mm, yy = raw.split("-")
        if len(yy) == 2:
            yy = "20" + yy
        try:
            dt = datetime(int(yy), int(mm), int(dd))
        except ValueError:
            continue
        try:
            df = pd.read_excel(
                s, sheet_name="Programas",
                usecols=["NOMBRE_DEL_PROGRAMA", "MODALIDAD", "SECTOR", "DEPARTAMENTO_OFERTA_PROGRAMA"])
            df = df.iloc[:-2]
        except Exception:
            continue
        modalidad_counts = df["MODALIDAD"].fillna("Sin definir").value_counts().to_dict()
        cruzado = list(zip(
            df["SECTOR"].fillna("Sin definir"),
            df["DEPARTAMENTO_OFERTA_PROGRAMA"].fillna("Sin definir"),
            df["MODALIDAD"].fillna("Sin definir"),
        ))
        rows.append((dt, len(df), s, modalidad_counts, cruzado))
    rows.sort(key=lambda x: x[0])
    return [
        {"fecha": r[0].strftime("%Y-%m-%d"), "total": r[1], "_path": r[2],
         "_modalidad": r[3], "_cruzado": r[4]}
        for r in rows
    ]


def _modalidad_bucketing(historico, top_n=6):
    """Calcula el top N de modalidades por total histórico y una función de
    bucketeo que colapsa el resto en 'Otras'. La usan tanto el gráfico de
    evolución como el filtro global, para que un mismo valor de modalidad
    signifique lo mismo en toda la página."""
    totales = {}
    for h in historico:
        for modalidad, count in h["_modalidad"].items():
            totales[modalidad] = totales.get(modalidad, 0) + count
    top_modalidades = sorted(totales, key=totales.get, reverse=True)[:top_n]
    top_set = set(top_modalidades)
    orden = list(top_modalidades)
    if "Otras" not in orden and any(m not in top_set for m in totales):
        orden.append("Otras")

    def bucket(m):
        return m if m in top_set else "Otras"

    return orden, bucket


def historico_por_modalidad(historico, top_n=6):
    """Series de evolución de programas activos por modalidad (top N + Otras)
    a través de los snapshots históricos, para las mini-tarjetas del index."""
    fechas = [h["fecha"] for h in historico]
    orden, bucket = _modalidad_bucketing(historico, top_n)
    series_vals = {name: [] for name in orden}
    for h in historico:
        counts = {}
        for modalidad, count in h["_modalidad"].items():
            b = bucket(modalidad)
            counts[b] = counts.get(b, 0) + count
        for name in orden:
            series_vals[name].append(counts.get(name, 0))
    return {"fechas": fechas, "series": [{"name": n, "values": series_vals[n]} for n in orden]}


def historico_cruzado(historico, top_n=6):
    """Cross-tab histórico fecha×sector×departamento×modalidad (bucketeada
    top N + 'Otras') a través de todos los snapshots. Sirve para que 'Total
    acumulado' y 'Evolución de modalidad' respondan al filtro global sin
    tener que traer el detalle fila-a-fila de cada snapshot al navegador.
    Se serializa como tuplas cortas [fecha, sector, depto, modalidad, n]
    porque el volumen (decenas de miles de combinaciones) hace que las
    claves de un dict repitan mucho peso en el JSON."""
    _, bucket = _modalidad_bucketing(historico, top_n)
    filas = []
    for h in historico:
        combo_counts = {}
        for sector, depto, modalidad in h["_cruzado"]:
            key = (sector, depto, bucket(modalidad))
            combo_counts[key] = combo_counts.get(key, 0) + 1
        for (sector, depto, modb), n in combo_counts.items():
            filas.append([h["fecha"], sector, depto, modb, n])
    return filas


def leer_novedades(nombre):
    path = NOVEDADES / nombre
    if not path.exists():
        return pd.DataFrame()
    return pd.read_excel(path)


def leer_snapshot_actual(historico):
    if not historico:
        return pd.DataFrame()
    path = historico[-1]["_path"]
    df = pd.read_excel(path, sheet_name="Programas")
    df = df.iloc[:-2].copy()
    cols_ok = [c for c in SNAPSHOT_COLS if c in df.columns]
    return df[cols_ok].copy()


def enriquecer_modificados(df_m):
    """Agrega columnas derivadas (_delta, _cambia_*) usadas por las páginas
    de análisis de créditos y costos."""
    df = df_m.copy()
    df["_cred_antes"] = df.get("NÚMERO_CRÉDITOS_ANTERIOR", pd.Series(dtype=object)).apply(_num)
    df["_cred_despues"] = df.get("NÚMERO_CRÉDITOS", pd.Series(dtype=object)).apply(_num)
    df["_costo_antes"] = df.get("COSTO_MATRÍCULA_ESTUD_NUEVOS_ANTERIOR", pd.Series(dtype=object)).apply(_num)
    df["_costo_despues"] = df.get("COSTO_MATRÍCULA_ESTUD_NUEVOS", pd.Series(dtype=object)).apply(_num)

    df["_delta"] = df.apply(
        lambda r: (r["_cred_despues"] - r["_cred_antes"])
        if (pd.notna(r["_cred_antes"]) and pd.notna(r["_cred_despues"])) else None, axis=1)

    df["_cambia_credito"] = df.apply(
        lambda r: bool(pd.notna(r["_cred_antes"]) and pd.notna(r["_cred_despues"])
                       and r["_cred_antes"] != r["_cred_despues"]), axis=1)
    df["_cambia_costo"] = df.apply(
        lambda r: bool(pd.notna(r["_costo_antes"]) and pd.notna(r["_costo_despues"])
                       and r["_costo_antes"] != r["_costo_despues"]), axis=1)
    df["_cambia_periodo"] = df.apply(
        lambda r: _cambia(r, "NÚMERO_PERIODOS_DE_DURACIÓN", "NÚMERO_PERIODOS_DE_DURACIÓN_ANTERIOR"), axis=1)
    df["_cambia_modalidad"] = df.apply(lambda r: _cambia(r, "MODALIDAD", "MODALIDAD_ANTERIOR"), axis=1)
    df["_cambia_municipio"] = df.apply(
        lambda r: _cambia(r, "MUNICIPIO_OFERTA_PROGRAMA", "MUNICIPIO_OFERTA_PROGRAMA_ANTERIOR"), axis=1)

    return df


def _universo_records(df_enr, extra_cols):
    cols = [c for c in COLS_UNIVERSO if c in df_enr.columns] + extra_cols
    out = []
    for _, row in df_enr.iterrows():
        rec = {}
        for c in cols:
            v = row[c]
            if pd.isna(v):
                rec[c] = None
            elif isinstance(v, float) and v == int(v):
                rec[c] = int(v)
            else:
                rec[c] = v if not isinstance(v, float) else v
                if isinstance(rec[c], (pd.Timestamp,)):
                    rec[c] = str(rec[c])
        out.append(rec)
    return out


PLOTLY_CDN = '<script src="https://cdn.plot.ly/plotly-2.35.2.min.js" charset="utf-8"></script>'


# ── index.html ─────────────────────────────────────────────────────────────────

INDEX_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SNIES Monitor · Uninorte</title>
__PLOTLY_CDN__
<style>
:root {
  --bg:#f1f5f9; --surface:#fff; --hdr1:#15284b; --hdr2:#2d5b9e;
  --text:#0f172a; --muted:#64748b; --border:#e2e8f0;
  --blue:#2d5b9e; --green:#fcc10e; --red:#ae1e22; --amber:#bd900b;
  --radius:0.75rem;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);font-size:14px}
header{background:linear-gradient(135deg,var(--hdr1),var(--hdr2));color:#fff;
  padding:1.25rem 2rem;display:flex;justify-content:space-between;align-items:center}
header h1{font-size:1.35rem;font-weight:700;letter-spacing:-.01em}
header .sub{font-size:.8rem;opacity:.75;margin-top:.2rem}
.badge-update{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);
  padding:.5rem 1rem;border-radius:2rem;font-size:.75rem;text-align:right;white-space:nowrap}
.badge-update strong{display:block;font-size:.9rem}
main{max-width:1380px;margin:0 auto;padding:1.5rem 2rem}
section{margin-bottom:1.5rem}
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:1rem}
.kpi{background:var(--surface);border-radius:var(--radius);padding:1.25rem 1.5rem;
  box-shadow:0 1px 3px rgba(0,0,0,.07);border-left:4px solid var(--blue)}
.kpi.g{border-left-color:var(--green)}.kpi.r{border-left-color:var(--red)}.kpi.a{border-left-color:var(--amber)}
.kpi.link{cursor:pointer;transition:transform .15s,box-shadow .15s}
.kpi.link:hover{transform:translateY(-3px);box-shadow:0 6px 16px rgba(0,0,0,.13)}
.kpi-label{font-size:.7rem;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:.4rem}
.kpi-val{font-size:2rem;font-weight:700;line-height:1}
.kpi-sub{font-size:.72rem;color:var(--muted);margin-top:.35rem}
.kpi-hint{font-size:.65rem;color:var(--blue);margin-top:.3rem;opacity:.8}
.card{background:var(--surface);border-radius:var(--radius);padding:1.25rem;
  box-shadow:0 1px 3px rgba(0,0,0,.07)}
.card-title{font-size:.7rem;font-weight:600;text-transform:uppercase;
  letter-spacing:.06em;color:var(--muted);margin-bottom:.9rem}
.chart-2col{display:grid;grid-template-columns:1fr 2fr;gap:1rem}
.mini-sub{font-size:.72rem;color:var(--muted);margin:-.4rem 0 1rem}
.mini-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:1rem}
.mini-card{background:var(--surface);border-radius:var(--radius);padding:1rem 1.1rem;
  box-shadow:0 1px 3px rgba(0,0,0,.07);border:1px solid var(--border)}
.mini-head{display:flex;justify-content:space-between;align-items:flex-start;gap:.5rem;margin-bottom:.15rem}
.mini-title{font-size:.72rem;font-weight:600;color:var(--muted);line-height:1.3}
.mini-badge{font-size:.72rem;font-weight:700;white-space:nowrap}
.mini-badge.up{color:var(--text)}
.mini-badge.down{color:var(--red)}
.mini-val{font-size:1.5rem;font-weight:700;margin-bottom:.35rem}
.mini-chart{height:72px}
.tab-nav{display:flex;border-bottom:1px solid var(--border);padding:0 1.5rem;background:var(--surface);
  border-radius:var(--radius) var(--radius) 0 0}
.tab-btn{padding:.85rem 1.25rem;border:none;background:none;cursor:pointer;font-size:.8rem;
  font-weight:500;color:var(--muted);border-bottom:2px solid transparent;transition:.15s}
.tab-btn.on{color:var(--blue);border-bottom-color:var(--blue)}
.tab-btn .n{font-size:.68rem;background:var(--bg);padding:.1rem .4rem;border-radius:1rem;margin-left:.35rem}
.tab-btn.on .n{background:var(--blue);color:#fff}
.tab-pane{display:none;padding:1.25rem 1.5rem;background:var(--surface);
  border-radius:0 0 var(--radius) var(--radius)}
.tab-pane.on{display:block}
.f-input{flex:1;min-width:170px;max-width:300px;padding:.5rem .8rem;border:1px solid var(--border);
  border-radius:.4rem;font-size:.8rem;outline:none}
.f-input:focus{border-color:var(--blue)}
.f-sel{padding:.45rem .6rem;border:1px solid var(--border);border-radius:.4rem;
  font-size:.77rem;background:var(--surface);outline:none;cursor:pointer;max-width:175px}
.f-sel:focus{border-color:var(--blue)}
.f-btn{padding:.45rem .85rem;border:1px solid var(--border);border-radius:.4rem;
  font-size:.77rem;background:var(--surface);cursor:pointer;color:var(--muted);white-space:nowrap}
.f-btn:hover{background:var(--bg)}
.tbl-wrap{max-height:420px;overflow-y:auto;border:1px solid var(--border);border-radius:.5rem}
table{width:100%;border-collapse:collapse;font-size:.78rem}
th{background:var(--bg);padding:.65rem .9rem;text-align:left;font-size:.68rem;
  text-transform:uppercase;letter-spacing:.05em;color:var(--muted);cursor:pointer;
  user-select:none;position:sticky;top:0;z-index:1;white-space:nowrap}
th:hover{background:#e2e8f0}
td{padding:.65rem .9rem;border-bottom:1px solid var(--border);vertical-align:top;
  max-width:260px;word-break:break-word}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8fafc}
.empty{text-align:center;color:var(--muted);padding:2.5rem}
.global-filter{position:sticky;top:0;z-index:40;background:var(--surface);
  border-bottom:1px solid var(--border);padding:.75rem 2rem;display:flex;gap:.5rem;
  align-items:center;flex-wrap:wrap;box-shadow:0 2px 6px rgba(0,0,0,.06)}
.gf-count{font-size:.75rem;color:var(--muted);margin-left:auto;white-space:nowrap}
@media(max-width:900px){
  .kpi-grid{grid-template-columns:repeat(2,1fr)}
  .chart-2col{grid-template-columns:1fr}
  main{padding:1rem}
  header{flex-direction:column;gap:.75rem;text-align:center}
  .global-filter{padding:.65rem 1rem}
  .gf-count{margin-left:0;width:100%}
}
</style>
</head>
<body>
<header>
  <div>
    <h1>📊 SNIES Monitor · Uninorte</h1>
    <div class="sub">Programas de posgrado en Colombia</div>
  </div>
  <div class="badge-update">
    <span style="opacity:.7;font-size:.68rem">Última actualización</span>
    <strong id="fecha-update">–</strong>
  </div>
</header>
<div class="global-filter">
  <input class="f-input" id="gf-q" style="max-width:320px" placeholder="Buscar por nombre, institución, código SNIES…" oninput="applyGlobalFilter()">
  <select id="gf-sector" class="f-sel" onchange="applyGlobalFilter()"><option value="">Todos los sectores</option></select>
  <select id="gf-depto" class="f-sel" onchange="applyGlobalFilter()"><option value="">Todos los departamentos</option></select>
  <select id="gf-modalidad" class="f-sel" onchange="applyGlobalFilter()"><option value="">Todas las modalidades</option></select>
  <button class="f-btn" onclick="resetGlobalFilter()">✕ Limpiar filtros</button>
  <span class="gf-count" id="gf-count"></span>
</div>
<main>
  <section class="kpi-grid">
    <div class="kpi">
      <div class="kpi-label">Programas activos</div>
      <div class="kpi-val" id="k-total">–</div>
      <div class="kpi-sub">de posgrado hoy</div>
    </div>
    <div class="kpi g link" onclick="location.href='nuevos.html'" title="Ver detalle">
      <div class="kpi-label">Nuevos (último run)</div>
      <div class="kpi-val" id="k-nue">–</div>
      <div class="kpi-sub" id="k-nue-sub">acumulado: –</div>
      <div class="kpi-hint">Ver detalle →</div>
    </div>
    <div class="kpi r link" onclick="location.href='inactivos.html'" title="Ver detalle">
      <div class="kpi-label">Inactivos (último run)</div>
      <div class="kpi-val" id="k-ina">–</div>
      <div class="kpi-sub" id="k-ina-sub">acumulado: –</div>
      <div class="kpi-hint">Ver detalle →</div>
    </div>
    <div class="kpi a link" onclick="location.href='modificados.html'" title="Ver detalle">
      <div class="kpi-label">Modificados (último run)</div>
      <div class="kpi-val" id="k-mod">–</div>
      <div class="kpi-sub" id="k-mod-sub">acumulado: –</div>
      <div class="kpi-hint">Ver detalle →</div>
    </div>
  </section>

  <section class="chart-2col" style="grid-template-columns:1fr 1fr">
    <div class="card">
      <div class="card-title">Total acumulado de programas activos</div>
      <div id="ch-historico" style="height:260px"></div>
    </div>
    <div class="card">
      <div class="card-title">Aperturas vs. cierres netos por periodo</div>
      <div id="ch-flujo" style="height:260px"></div>
      <div id="flujo-note" style="font-size:.7rem;color:#64748b;margin-top:.6rem;line-height:1.4"></div>
    </div>
  </section>

  <section class="card">
    <div class="card-title">Evolución de la modalidad de los programas activos</div>
    <div class="mini-sub">Cada panel tiene su propia escala — ordenados de mayor a menor crecimiento relativo desde el primer registro</div>
    <div class="mini-grid" id="ch-modalidad-mini"></div>
  </section>

  <section class="chart-2col">
    <div class="card">
      <div class="card-title">Por sector</div>
      <div id="ch-sector" style="height:260px"></div>
    </div>
    <div class="card">
      <div class="card-title">Top 15 departamentos de oferta</div>
      <div id="ch-depto" style="height:260px"></div>
    </div>
  </section>

  <section class="card">
    <div class="card-title">Distribución de programas activos por duración (periodos requeridos) — clic para filtrar</div>
    <div id="ch-periodos" style="height:260px"></div>
    <div id="periodos-selector" style="display:flex;flex-wrap:wrap;gap:.35rem;margin-top:.6rem"></div>
  </section>

  <div id="periodos-chip" style="display:none;align-items:center;gap:.6rem;
      padding:.55rem 1.5rem;background:#eaf1f8;border:1px solid #c7d7ea;
      border-radius:.5rem;margin-bottom:.75rem">
    <span style="font-size:.78rem;color:#15284b">Filtrado por duración:
      <strong id="periodos-chip-val"></strong></span>
    <button onclick="clearPeriodosFilter()"
      style="background:none;border:none;cursor:pointer;color:#2d5b9e;font-size:.82rem;padding:0">
      ✕ Limpiar filtro</button>
  </div>

  <section id="snap-section" style="display:none">
    <div class="card">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:.75rem;flex-wrap:wrap;gap:.5rem">
        <div class="card-title" style="margin-bottom:0">Programas activos — <span id="snap-title"></span></div>
      </div>
      <div class="tbl-wrap" id="snap-tbl"></div>
    </div>
  </section>

  <section>
    <div class="tab-nav">
      <button class="tab-btn on" onclick="tab('nue',this)">
        Nuevos <span class="n" id="bn-nue">0</span>
      </button>
      <button class="tab-btn" onclick="tab('ina',this)">
        Inactivos <span class="n" id="bn-ina">0</span>
      </button>
      <button class="tab-btn" onclick="tab('mod',this)">
        Modificados <span class="n" id="bn-mod">0</span>
      </button>
    </div>
    <div id="tp-nue" class="tab-pane on">
      <div class="tbl-wrap" id="tw-nue"></div>
    </div>
    <div id="tp-ina" class="tab-pane">
      <div class="tbl-wrap" id="tw-ina"></div>
    </div>
    <div id="tp-mod" class="tab-pane">
      <div class="tbl-wrap" id="tw-mod"></div>
    </div>
  </section>
</main>

<script>
const D = __DATA__;
const fmt = n => (n ?? 0).toLocaleString('es-CO');
const _norm = s => String(s==null?'':s).normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').toLowerCase();
function _rowMatches(r, tokens) {
  if (!tokens.length) return true;
  const hay = _norm(Object.values(r).join(' '));
  return tokens.every(t => hay.includes(t));
}
function gv(id) { const el = document.getElementById(id); return el ? el.value : ''; }
function getSem(s) {
  if (!s || !s.trim()) return null;
  let y, m;
  const dmy = s.match(/^(\\d{2})\\/(\\d{2})\\/(\\d{4})/);
  if (dmy) { y = +dmy[3]; m = +dmy[2]; }
  if (!y || y < 2014 || y > 2035) return null;
  return y + '-' + (m <= 6 ? '1' : '2');
}
function uniq(arr) { return [...new Set(arr.filter(v => v && String(v).trim() !== ''))].sort(); }
function addOpts(id, vals) {
  const el = document.getElementById(id); if (!el) return;
  vals.forEach(v => { const o = document.createElement('option'); o.value = o.textContent = v; el.appendChild(o); });
}
document.getElementById('fecha-update').textContent = D.ultima_actualizacion;
document.getElementById('k-total').textContent = fmt(D.kpis.total_activos);
document.getElementById('k-nue').textContent   = fmt(D.kpis.nuevos_ultimo);
document.getElementById('k-ina').textContent   = fmt(D.kpis.inactivos_ultimo);
document.getElementById('k-mod').textContent   = fmt(D.kpis.mods_ultimo);
document.getElementById('k-nue-sub').textContent = 'acumulado: ' + fmt(D.kpis.nuevos_total);
document.getElementById('k-ina-sub').textContent = 'acumulado: ' + fmt(D.kpis.inactivos_total);
document.getElementById('k-mod-sub').textContent = 'acumulado: ' + fmt(D.kpis.mods_total);


function _emptyChart(el, msg) {
  if (!el) return;
  el.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#64748b;font-size:.82rem">'+(msg||'Sin datos')+'</div>';
}

const PERIODOS_TOPE = 12;
function _bucketPeriodo(v) {
  const n = Math.round(parseFloat(v));
  if (isNaN(n)) return null;
  return n >= PERIODOS_TOPE ? PERIODOS_TOPE + '+' : String(n);
}
function _periodoSortKey(l) { return l === PERIODOS_TOPE + '+' ? Infinity : parseInt(l, 10); }

function renderSector(data) {
  const counts = {};
  data.forEach(r => { const k = r['SECTOR'] || 'Sin definir'; counts[k] = (counts[k] || 0) + 1; });
  const labels = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);
  if (!labels.length) { _emptyChart(document.getElementById('ch-sector')); return; }
  Plotly.newPlot('ch-sector', [{
    labels, values: labels.map(l => counts[l]),
    type:'pie', hole:0.45,
    marker:{colors:['#2d5b9e','#fcc10e','#bd900b','#ae1e22','#6e91b9','#214174']},
    textinfo:'label+percent',
    hovertemplate:'%{label}<br><b>%{value:,}</b><extra></extra>'
  }], {
    margin:{t:10,r:10,b:10,l:10}, showlegend:false,
    plot_bgcolor:'white', paper_bgcolor:'white'
  }, {responsive:true, displayModeBar:false});
}

function renderDepto(data) {
  const counts = {};
  data.forEach(r => { const k = r['DEPARTAMENTO_OFERTA_PROGRAMA']; if (!k) return; counts[k] = (counts[k] || 0) + 1; });
  const top = Object.keys(counts).sort((a, b) => counts[b] - counts[a]).slice(0, 15).reverse();
  if (!top.length) { _emptyChart(document.getElementById('ch-depto')); return; }
  Plotly.newPlot('ch-depto', [{
    y: top, x: top.map(l => counts[l]),
    type:'bar', orientation:'h',
    marker:{color:'#2d5b9e', opacity:0.82},
    hovertemplate:'%{y}<br><b>%{x:,}</b><extra></extra>'
  }], {
    margin:{t:10,r:20,b:40,l:170},
    xaxis:{showgrid:true, gridcolor:'#e2e8f0', tickfont:{size:11}},
    yaxis:{tickfont:{size:11}},
    plot_bgcolor:'white', paper_bgcolor:'white', bargap:0.3
  }, {responsive:true, displayModeBar:false});
}

function renderPeriodos(data) {
  const combos = {};
  const periodicidades = new Set();
  data.forEach(r => {
    const label = _bucketPeriodo(r['NÚMERO_PERIODOS_DE_DURACIÓN']);
    if (label == null) return;
    const per = r['PERIODICIDAD'] || 'Sin definir';
    periodicidades.add(per);
    (combos[label] = combos[label] || {})[per] = (combos[label][per] || 0) + 1;
  });
  const labels = Object.keys(combos).sort((a, b) => _periodoSortKey(a) - _periodoSortKey(b));
  const sel = document.getElementById('periodos-selector');
  if (!labels.length) {
    _emptyChart(document.getElementById('ch-periodos'));
    if (sel) sel.innerHTML = '';
    return;
  }
  const series = [...periodicidades].map(per => ({
    name: per, values: labels.map(l => (combos[l] && combos[l][per]) || 0)
  })).sort((a, b) => b.values.reduce((x, y) => x + y, 0) - a.values.reduce((x, y) => x + y, 0));

  const PCOLORS = ['#2d5b9e','#fcc10e','#bd900b','#ae1e22','#6e91b9',
                   '#214174','#d56f18','#948e56','#15284b','#7a1518'];
  const traces = series.map((s, i) => ({
    x: labels, y: s.values, name: s.name, type: 'bar',
    marker: {color: PCOLORS[i % PCOLORS.length], opacity: 0.85},
    hovertemplate: s.name + '<br>%{x} periodos — <b>%{y:,}</b> programas<extra></extra>'
  }));
  Plotly.newPlot('ch-periodos', traces, {
    barmode: 'stack',
    margin: {t:10, r:20, b:90, l:70},
    xaxis: {title: 'Periodos', type: 'category', tickmode: 'array',
            tickvals: labels, tickfont: {size:11}},
    yaxis: {title: 'N. Programas', showgrid: true,
            gridcolor: '#e2e8f0', tickfont: {size:11}},
    plot_bgcolor: 'white', paper_bgcolor: 'white', bargap: 0.25,
    legend: {orientation: 'h', y: -0.28, font: {size: 11}, entrywidth: 120, entrywidthmode: 'pixels'},
    hovermode: 'x unified'
  }, {responsive: true, displayModeBar: false}).then(gd => {
    gd.on('plotly_click', function(ev) { setPeriodosFilter(ev.points[0].x); });
  });

  const totals = {};
  labels.forEach(l => { totals[l] = combos[l] ? Object.values(combos[l]).reduce((a, b) => a + b, 0) : 0; });
  if (sel) {
    sel.innerHTML = labels.map(l =>
      '<button id="pchip-'+l+'" onclick="setPeriodosFilter(\\''+l+'\\')" '+
      'style="padding:.22rem .65rem;background:#f1f5f9;border:1px solid #e2e8f0;'+
      'border-radius:2rem;font-size:.72rem;cursor:pointer;transition:all .15s;white-space:nowrap">'+
      l+' sem&nbsp;<span style="opacity:.6">('+totals[l].toLocaleString('es-CO')+')</span></button>'
    ).join('');
    _updateChipStyles();
  }
}

function _historicoAgregado() {
  const filas = D.historico_cruzado || [];
  const fechas = (D.historico || []).map(h => h.fecha);
  const totalPorFecha = {}; fechas.forEach(f => { totalPorFecha[f] = 0; });
  const porModalidad = {};
  filas.forEach(fila => {
    const fecha = fila[0], sector = fila[1], depto = fila[2], modalidad = fila[3], n = fila[4];
    if (FILTRO.sector && sector !== FILTRO.sector) return;
    if (FILTRO.depto && depto !== FILTRO.depto) return;
    if (FILTRO.modalidad && modalidad !== FILTRO.modalidad) return;
    totalPorFecha[fecha] = (totalPorFecha[fecha] || 0) + n;
    (porModalidad[modalidad] = porModalidad[modalidad] || {})[fecha] = (porModalidad[modalidad][fecha] || 0) + n;
  });
  return {
    fechas,
    totales: fechas.map(f => totalPorFecha[f] || 0),
    series: MODALIDADES_TOP.filter(name => porModalidad[name]).map(name => ({
      name, values: fechas.map(f => (porModalidad[name] && porModalidad[name][f]) || 0)
    }))
  };
}

function renderHistorico(agg) {
  const el = document.getElementById('ch-historico');
  if (!agg.fechas.length) { _emptyChart(el); return; }
  Plotly.newPlot('ch-historico', [{
    x: agg.fechas, y: agg.totales,
    type: 'scatter', mode: 'lines+markers',
    line: {color: '#2d5b9e', width: 2.5}, marker: {color: '#2d5b9e', size: 6},
    hovertemplate: '%{x}<br><b>%{y:,}</b> programas activos<extra></extra>'
  }], {
    margin: {t:10, r:20, b:40, l:55},
    xaxis: {showgrid: false, tickfont: {size: 10}},
    yaxis: {showgrid: true, gridcolor: '#e2e8f0', tickfont: {size: 11}},
    plot_bgcolor: 'white', paper_bgcolor: 'white', hovermode: 'x unified'
  }, {responsive: true, displayModeBar: false});
}

function renderModalidadMini(agg) {
  const cont = document.getElementById('ch-modalidad-mini');
  if (!agg.series.length) { _emptyChart(cont); return; }

  const withGrowth = agg.series.map(s => {
    const vals = s.values;
    const first = vals.length ? vals[0] : null;
    const last = vals.length ? vals[vals.length - 1] : null;
    let pct = null, isNew = false;
    if (first === 0) { if (last > 0) isNew = true; else pct = 0; }
    else if (first != null) { pct = ((last - first) / first) * 100; }
    const sortKey = isNew ? Infinity : (pct == null ? -Infinity : pct);
    return {name: s.name, values: vals, last, pct, isNew, sortKey};
  });
  withGrowth.sort((a, b) => b.sortKey - a.sortKey);

  cont.innerHTML = withGrowth.map((s, i) => {
    let badge;
    if (s.isNew) badge = '<span class="mini-badge up">▲ nuevo</span>';
    else {
      const cls = s.pct >= 0 ? 'up' : 'down';
      const arrow = s.pct >= 0 ? '▲' : '▼';
      badge = '<span class="mini-badge ' + cls + '">' + arrow + ' ' + (s.pct >= 0 ? '+' : '') + Math.round(s.pct) + '%</span>';
    }
    return '<div class="mini-card">' +
      '<div class="mini-head"><div class="mini-title">' + s.name + '</div>' + badge + '</div>' +
      '<div class="mini-val">' + fmt(s.last) + '</div>' +
      '<div class="mini-chart" id="mini-chart-' + i + '"></div>' +
    '</div>';
  }).join('');

  withGrowth.forEach((s, i) => {
    Plotly.newPlot('mini-chart-' + i, [{
      x: agg.fechas, y: s.values, type: 'scatter', mode: 'lines',
      line: {color: '#2d5b9e', width: 2},
      hovertemplate: '%{x}<br><b>%{y:,}</b> programas<extra></extra>'
    }], {
      margin: {t:4, r:4, b:20, l:32},
      xaxis: {showgrid: false, tickfont: {size: 9}, type: 'date', tickformat: '%Y', dtick: 'M12'},
      yaxis: {showgrid: false, tickfont: {size: 9}, zeroline: false, rangemode: 'tozero'},
      plot_bgcolor: 'white', paper_bgcolor: 'white', hovermode: 'x'
    }, {responsive: true, displayModeBar: false});
  });
}

function renderFlujo() {
  const nue = rows.nue.filter(passesGlobalFilter);
  const ina = rows.ina.filter(passesGlobalFilter);
  const porSem = {};
  const runsPorSem = {};
  const trackRun = (r, key) => {
    const s = getSem(r['FECHA_OBTENCION']); if (!s) return null;
    (porSem[s] = porSem[s] || {nuevos: 0, inactivos: 0})[key]++;
    (runsPorSem[s] = runsPorSem[s] || new Set()).add(r['FECHA_OBTENCION']);
    return s;
  };
  nue.forEach(r => trackRun(r, 'nuevos'));
  ina.forEach(r => trackRun(r, 'inactivos'));
  const sems = Object.keys(porSem).sort();
  const noteEl = document.getElementById('flujo-note');
  if (!sems.length) { _emptyChart(document.getElementById('ch-flujo')); if (noteEl) noteEl.textContent = ''; return; }
  const aperturas = sems.map(s => porSem[s].nuevos);
  const cierres = sems.map(s => -porSem[s].inactivos);
  const netos = sems.map((s, i) => aperturas[i] + cierres[i]);
  const runs = sems.map(s => runsPorSem[s].size);

  const UMBRAL_RUNS = 5;
  const incompletos = sems.map((s, i) => runs[i] < UMBRAL_RUNS);
  const opacityAp = incompletos.map(f => f ? .35 : .85);
  const opacityCi = incompletos.map(f => f ? .35 : .85);

  Plotly.newPlot('ch-flujo', [
    {x: sems, y: aperturas, name: 'Aperturas (nuevos)', type: 'bar',
     marker: {color: '#1a9e6b', opacity: opacityAp},
     customdata: runs,
     hovertemplate: '%{x}<br>+%{y} aperturas<br>(%{customdata} corridas del monitor)<extra></extra>'},
    {x: sems, y: cierres, name: 'Cierres (inactivos)', type: 'bar',
     marker: {color: '#ae1e22', opacity: opacityCi},
     customdata: runs,
     hovertemplate: '%{x}<br>%{y} cierres<br>(%{customdata} corridas del monitor)<extra></extra>'},
    {x: sems, y: netos, name: 'Neto', type: 'scatter', mode: 'lines+markers',
     line: {color: '#15284b', width: 2}, marker: {color: '#15284b', size: 6},
     hovertemplate: '%{x}<br>Neto: <b>%{y}</b><extra></extra>'},
  ], {
    barmode: 'relative',
    margin: {t:10, r:20, b:40, l:45},
    xaxis: {showgrid: false, tickfont: {size: 10}, type: 'category'},
    yaxis: {showgrid: true, gridcolor: '#e2e8f0', zeroline: true, zerolinecolor: '#94a3b8', tickfont: {size: 11}},
    plot_bgcolor: 'white', paper_bgcolor: 'white', hovermode: 'x unified',
    legend: {orientation: 'h', y: -0.22, font: {size: 10}},
    annotations: sems.map((s, i) => incompletos[i] ? {
      x: s, y: Math.max(aperturas[i], 0), yshift: 16, xref: 'x', yref: 'y',
      text: '⚠ ' + runs[i] + ' corrida' + (runs[i] === 1 ? '' : 's'),
      showarrow: false, font: {size: 9, color: '#bd900b'}
    } : null).filter(Boolean)
  }, {responsive: true, displayModeBar: false});

  if (noteEl) {
    const flagged = sems.filter((s, i) => incompletos[i]);
    noteEl.textContent = flagged.length
      ? '⚠ Cobertura incompleta en ' + flagged.join(', ') + ': el monitor corrio muy pocas veces ese periodo (menos de ' +
        UMBRAL_RUNS + ' corridas), asi que las cifras no son comparables 1:1 con los demas periodos.'
      : '';
  }
}

let periodosFiltro = null;
const _snapAll = D.snapshot || [];

function _matchesPeriodo(v, filtro) {
  if (v === undefined || v === '') return false;
  const n = Math.round(parseFloat(String(v)));
  if (isNaN(n)) return false;
  return filtro === '12+' ? n >= 12 : n === parseInt(filtro, 10);
}
function _periodoLabel(filtro) {
  return filtro === '12+' ? '12 o más periodos' : filtro + ' periodos';
}

const _SNAP_COLS = ['CÓDIGO_SNIES_DEL_PROGRAMA','NOMBRE_DEL_PROGRAMA','NOMBRE_INSTITUCIÓN',
                    'SECTOR','MODALIDAD','DEPARTAMENTO_OFERTA_PROGRAMA','PERIODICIDAD'];
const _SNAP_HEAD = {
  'CÓDIGO_SNIES_DEL_PROGRAMA':'Cód. SNIES', 'NOMBRE_DEL_PROGRAMA':'Programa',
  'NOMBRE_INSTITUCIÓN':'Institución', 'SECTOR':'Sector', 'MODALIDAD':'Modalidad',
  'DEPARTAMENTO_OFERTA_PROGRAMA':'Departamento', 'PERIODICIDAD':'Periodicidad'
};

function _buildSnapTbl(data) {
  const cols = _SNAP_COLS.filter(c => !data.length || c in data[0]);
  let h = '<table><thead><tr>';
  cols.forEach(c => { h += '<th>' + (_SNAP_HEAD[c]||c) + '</th>'; });
  h += '</tr></thead><tbody>';
  if (!data.length) {
    h += '<tr><td colspan="'+cols.length+'" class="empty">Sin registros</td></tr>';
  } else {
    data.forEach(r => { h += '<tr>' + cols.map(c => '<td>'+(r[c]||'')+'</td>').join('') + '</tr>'; });
  }
  return h + '</tbody></table>';
}

function filterSnap() {
  const res = _snapAll.filter(r =>
    _matchesPeriodo(r['NÚMERO_PERIODOS_DE_DURACIÓN'], periodosFiltro) && passesGlobalFilter(r));
  document.getElementById('snap-tbl').innerHTML = _buildSnapTbl(res);
}

function _updateChipStyles() {
  document.querySelectorAll('[id^="pchip-"]').forEach(btn => {
    const active = btn.id === 'pchip-' + periodosFiltro;
    btn.style.background    = active ? '#2d5b9e' : '#f1f5f9';
    btn.style.color         = active ? '#fff'    : '';
    btn.style.borderColor   = active ? '#2d5b9e' : '#e2e8f0';
    btn.style.fontWeight    = active ? '600'     : '';
  });
}

function setPeriodosFilter(val) {
  val = String(val);
  if (periodosFiltro === val) { clearPeriodosFilter(); return; }
  periodosFiltro = val;
  const matching = _snapAll.filter(r =>
    _matchesPeriodo(r['NÚMERO_PERIODOS_DE_DURACIÓN'], val) && passesGlobalFilter(r));
  const label = _periodoLabel(val) + ' — ' + matching.length.toLocaleString('es-CO') + ' programas activos';
  document.getElementById('periodos-chip').style.display = 'flex';
  document.getElementById('periodos-chip-val').textContent = label;
  document.getElementById('snap-title').textContent = label;
  filterSnap();
  document.getElementById('snap-section').style.display = 'block';
  _updateChipStyles();
  document.getElementById('snap-section').scrollIntoView({behavior:'smooth'});
}

function clearPeriodosFilter() {
  periodosFiltro = null;
  document.getElementById('periodos-chip').style.display = 'none';
  document.getElementById('snap-section').style.display = 'none';
  _updateChipStyles();
}

const COLS = {
  nue: ['FECHA_OBTENCION','CÓDIGO_SNIES_DEL_PROGRAMA','NOMBRE_DEL_PROGRAMA',
        'NOMBRE_INSTITUCIÓN','MODALIDAD','DEPARTAMENTO_OFERTA_PROGRAMA','DIVISIÓN UNINORTE'],
  ina: ['FECHA_OBTENCION','CÓDIGO_SNIES_DEL_PROGRAMA','NOMBRE_DEL_PROGRAMA',
        'NOMBRE_INSTITUCIÓN','MODALIDAD','DEPARTAMENTO_OFERTA_PROGRAMA','DIVISIÓN UNINORTE'],
  mod: ['FECHA_OBTENCION','CÓDIGO_SNIES_DEL_PROGRAMA','NOMBRE_DEL_PROGRAMA',
        'NOMBRE_INSTITUCIÓN','QUE_CAMBIO']
};
const HEAD = {
  FECHA_OBTENCION:'Fecha', 'CÓDIGO_SNIES_DEL_PROGRAMA':'Cód.',
  NOMBRE_DEL_PROGRAMA:'Programa', 'NOMBRE_INSTITUCIÓN':'Institución',
  SECTOR:'Sector', MODALIDAD:'Modalidad',
  DEPARTAMENTO_OFERTA_PROGRAMA:'Dpto.', 'DIVISIÓN UNINORTE':'División',
  QUE_CAMBIO:'¿Qué cambió?'
};
const rows = {nue: D.nuevos || [], ina: D.inactivos || [], mod: D.modificados || []};
let filteredRows = {nue: rows.nue, ina: rows.ina, mod: rows.mod};
let sortDir = {};

// ── Filtro global (búsqueda + sector + departamento + modalidad) ─────────────
// La modalidad usa el mismo agrupamiento top6+Otras que las mini-tarjetas de
// evolución, así un mismo valor significa lo mismo en toda la página.
const MODALIDADES_TOP = ((D.historico_modalidad || {}).series || []).map(s => s.name);
const MODALIDADES_TOP_SET = new Set(MODALIDADES_TOP.filter(n => n !== 'Otras'));
function bucketModalidad(m) { return MODALIDADES_TOP_SET.has(m) ? m : 'Otras'; }

let FILTRO = {qTokens: [], sector: '', depto: '', modalidad: ''};
function passesGlobalFilter(r) {
  if (!_rowMatches(r, FILTRO.qTokens)) return false;
  if (FILTRO.sector && r['SECTOR'] !== FILTRO.sector) return false;
  if (FILTRO.depto && r['DEPARTAMENTO_OFERTA_PROGRAMA'] !== FILTRO.depto) return false;
  if (FILTRO.modalidad && bucketModalidad(r['MODALIDAD']) !== FILTRO.modalidad) return false;
  return true;
}

addOpts('gf-sector', uniq([..._snapAll, ...rows.nue, ...rows.ina, ...rows.mod].map(r => r['SECTOR'])));
addOpts('gf-depto', uniq([..._snapAll, ...rows.nue, ...rows.ina, ...rows.mod].map(r => r['DEPARTAMENTO_OFERTA_PROGRAMA'])));
addOpts('gf-modalidad', MODALIDADES_TOP);

function buildTbl(type, data) {
  const cols = COLS[type].filter(c => !data.length || c in data[0]);
  let h = '<table><thead><tr>';
  cols.forEach(c => {
    h += '<th onclick="sortTbl(\\'' + type + '\\',\\'' + c + '\\')">' + (HEAD[c]||c) + ' <span style="opacity:.4">↕</span></th>';
  });
  h += '</tr></thead><tbody>';
  if (!data.length) {
    h += '<tr><td colspan="' + cols.length + '" class="empty">Sin registros</td></tr>';
  } else {
    data.forEach(r => {
      h += '<tr>' + cols.map(c => '<td>' + (r[c]||'') + '</td>').join('') + '</tr>';
    });
  }
  return h + '</tbody></table>';
}

function render(type, data) {
  document.getElementById('tw-' + type).innerHTML = buildTbl(type, data);
}

function tab(id, btn) {
  document.querySelectorAll('.tab-pane').forEach(el => el.classList.remove('on'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('on'));
  document.getElementById('tp-' + id).classList.add('on');
  btn.classList.add('on');
}

function sortTbl(type, col) {
  const key = type + col;
  sortDir[key] = !sortDir[key];
  filteredRows[type] = [...filteredRows[type]].sort((a, b) => {
    const va = a[col] || '', vb = b[col] || '';
    return sortDir[key] ? va.localeCompare(vb, 'es') : vb.localeCompare(va, 'es');
  });
  render(type, filteredRows[type]);
}

function applyGlobalFilter() {
  FILTRO = {
    qTokens: _norm(gv('gf-q')).split(/\\s+/).filter(Boolean),
    sector: gv('gf-sector'),
    depto: gv('gf-depto'),
    modalidad: gv('gf-modalidad'),
  };

  ['nue','ina','mod'].forEach(t => {
    filteredRows[t] = rows[t].filter(passesGlobalFilter);
    sortDir = {};
    render(t, filteredRows[t]);
    document.getElementById('bn-' + t).textContent = filteredRows[t].length;
  });

  const filteredSnapshot = _snapAll.filter(passesGlobalFilter);
  renderSector(filteredSnapshot);
  renderDepto(filteredSnapshot);
  renderPeriodos(filteredSnapshot);
  renderFlujo();
  const agg = _historicoAgregado();
  renderHistorico(agg);
  renderModalidadMini(agg);

  const activo = FILTRO.qTokens.length || FILTRO.sector || FILTRO.depto || FILTRO.modalidad;
  const gc = document.getElementById('gf-count');
  if (gc) {
    gc.textContent = activo
      ? fmt(filteredSnapshot.length) + ' de ' + fmt(_snapAll.length) + ' programas activos'
      : fmt(_snapAll.length) + ' programas activos';
  }

  if (periodosFiltro !== null) filterSnap();
}

function resetGlobalFilter() {
  ['gf-q','gf-sector','gf-depto','gf-modalidad'].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = '';
  });
  applyGlobalFilter();
}

applyGlobalFilter();
</script>
</body>
</html>
"""

# ── Página de detalle (nuevos / inactivos / modificados) ─────────────────────
# Las 3 páginas comparten exactamente el mismo CSS/HTML/JS; sólo cambian el
# color de cabecera, el emoji/título, los enlaces extra (modificados) y qué
# tarjetas de gráficos se muestran en la sección central.

DETAIL_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITLE__</title>
__PLOTLY_CDN__
<style>
:root{
  --bg:#f1f5f9;--surface:#fff;--text:#0f172a;--muted:#64748b;
  --border:#e2e8f0;--blue:#2d5b9e;--radius:0.75rem;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);font-size:14px}
header{color:#fff;padding:1.2rem 2rem;display:flex;justify-content:space-between;align-items:center;gap:1rem}
header h1{font-size:1.25rem;font-weight:700}
header .sub{font-size:.77rem;opacity:.75;margin-top:.2rem}
.back-btn{display:inline-flex;align-items:center;gap:.3rem;background:rgba(255,255,255,.2);
  border:1px solid rgba(255,255,255,.35);color:#fff;text-decoration:none;padding:.38rem .85rem;
  border-radius:.4rem;font-size:.8rem;font-weight:500;white-space:nowrap;transition:background .15s}
.back-btn:hover{background:rgba(255,255,255,.32)}
.badge-update{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);
  padding:.45rem .9rem;border-radius:2rem;font-size:.74rem;text-align:right;white-space:nowrap}
.badge-update strong{display:block;font-size:.88rem}
/* filter bar */
.filter-bar{position:sticky;top:0;z-index:100;background:var(--surface);
  border-bottom:1px solid var(--border);padding:.6rem 2rem;
  display:flex;gap:.45rem;flex-wrap:wrap;align-items:center;
  box-shadow:0 2px 8px rgba(0,0,0,.06)}
.f-input{flex:1;min-width:200px;padding:.4rem .75rem;border:1px solid var(--border);
  border-radius:.4rem;font-size:.8rem;outline:none}
.f-input:focus{border-color:var(--blue)}
.f-sel{padding:.4rem .6rem;border:1px solid var(--border);border-radius:.4rem;
  font-size:.77rem;background:var(--surface);outline:none;cursor:pointer;max-width:175px}
.f-sel:focus{border-color:var(--blue)}
.f-btn{padding:.4rem .85rem;border:1px solid var(--border);border-radius:.4rem;
  font-size:.77rem;background:var(--surface);cursor:pointer;color:var(--muted);white-space:nowrap}
.f-btn:hover{background:var(--bg)}
.f-count{margin-left:auto;font-size:.82rem;font-weight:600;color:var(--blue);white-space:nowrap}
.ac-wrap{position:relative;flex:0 1 240px}
.ac-menu{position:absolute;top:calc(100% + 2px);left:0;right:0;z-index:300;
  background:var(--surface);border:1px solid var(--border);border-radius:.4rem;
  box-shadow:0 8px 20px rgba(0,0,0,.14);max-height:260px;overflow-y:auto;display:none}
.ac-menu.show{display:block}
.ac-item{padding:.45rem .75rem;font-size:.78rem;cursor:pointer;color:var(--text)}
.ac-item:hover{background:var(--bg)}
.ac-empty{padding:.45rem .75rem;font-size:.78rem;color:var(--muted)}
/* layout */
main{max-width:1380px;margin:0 auto;padding:1.4rem 2rem}
.card{background:var(--surface);border-radius:var(--radius);padding:1.2rem;
  box-shadow:0 1px 3px rgba(0,0,0,.07);margin-bottom:1rem}
.ct{font-size:.68rem;font-weight:600;text-transform:uppercase;letter-spacing:.06em;
  color:var(--muted);margin-bottom:.85rem}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1rem}
/* table */
.tbl-card{padding:0}
.tbl-card .ct{padding:1.1rem 1.2rem .5rem}
.tbl-wrap{max-height:520px;overflow-y:auto;border-top:1px solid var(--border)}
table{width:100%;border-collapse:collapse;font-size:.77rem}
th{background:var(--bg);padding:.6rem .85rem;text-align:left;font-size:.67rem;
  text-transform:uppercase;letter-spacing:.05em;color:var(--muted);cursor:pointer;
  user-select:none;position:sticky;top:0;z-index:1;white-space:nowrap}
th:hover{background:#e2e8f0}
td{padding:.6rem .85rem;border-bottom:1px solid var(--border);vertical-align:top;
  max-width:270px;word-break:break-word}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8fafc}
.empty{text-align:center;color:var(--muted);padding:2.5rem}
@media(max-width:900px){
  .g2{grid-template-columns:1fr}
  main{padding:1rem}
  header{flex-direction:column;gap:.6rem;text-align:center}
  .filter-bar{padding:.6rem 1rem}
  .f-count{margin-left:0}
}
</style>
</head>
<body>
<header style="background:__HDRGRAD__">
  <div style="display:flex;align-items:center;gap:.9rem">
    <a href="index.html" class="back-btn">← Dashboard</a>
    <div>
      <h1>__EMOJI__ __H1TEXT__</h1>
      <div class="sub">Programas de posgrado · Colombia</div>
    </div>
  </div>
  __BACKLINKS__
  <div class="badge-update">
    <span style="opacity:.7;font-size:.67rem">Total acumulado</span>
    <strong id="badge-total">–</strong>
  </div>
</header>

<div class="filter-bar">
  <input id="f-q" class="f-input" placeholder="Buscar por nombre, institución, código SNIES, departamento…" oninput="applyFilters()">
  <select id="f-sector"   class="f-sel" onchange="applyFilters()"><option value="">Todos los sectores</option></select>
  <select id="f-depto"    class="f-sel" onchange="applyFilters()"><option value="">Todos los departamentos</option></select>
  <div class="ac-wrap">
    <input id="f-institucion" class="f-sel" placeholder="Buscar institucion..." style="cursor:text;width:100%" oninput="applyFilters()">
    <div class="ac-menu" id="f-institucion-menu"></div>
  </div>
  <select id="f-division" class="f-sel" onchange="applyFilters()"><option value="">Todas las divisiones</option></select>
  __FILTRO_EXTRA__
  <select id="f-fecha"    class="f-sel" onchange="applyFilters()"><option value="">Todas las fechas</option></select>
  <button class="f-btn" onclick="resetFilters()">✕ Limpiar</button>
  <span class="f-count" id="f-count">–</span>
</div>

<main>
  <section>
__CHARTS_SECTION__
</section>

  <div id="per-det-chip" style="display:none;align-items:center;gap:.6rem;
      padding:.5rem 1.25rem;background:#eaf1f8;border:1px solid #c7d7ea;
      border-radius:.5rem;margin-bottom:.75rem">
    <span style="font-size:.78rem;color:#15284b">Filtrado por duración:
      <strong id="per-det-val"></strong></span>
    <button onclick="clearPeriodosDetalleFilter()"
      style="background:none;border:none;cursor:pointer;color:#2d5b9e;font-size:.82rem;padding:0">
      ✕ Limpiar filtro</button>
  </div>

  <div id="cine-det-chip" style="display:none;align-items:center;gap:.6rem;
      padding:.5rem 1.25rem;background:#fef6dc;border:1px solid #fce6a6;
      border-radius:.5rem;margin-bottom:.75rem">
    <span style="font-size:.78rem;color:#7a5d06">Filtrado por CINE:
      <strong id="cine-det-val"></strong></span>
    <button onclick="clearCineFiltro()"
      style="background:none;border:none;cursor:pointer;color:#bd900b;font-size:.82rem;padding:0">
      ✕ Limpiar filtro</button>
  </div>

  <section class="card tbl-card">
    <div class="ct">Registros</div>
    <div class="tbl-wrap" id="tbl-wrap"></div>
  </section>
</main>

<script>
const ROWS = __DATA__;
const CFG  = __CFG__;

const PC = {responsive:true, displayModeBar:false};
const fmt = n => (n ?? 0).toLocaleString('es-CO');
const _norm = s => String(s==null?'':s).normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').toLowerCase();
function _rowMatches(r, tokens) {
  if (!tokens.length) return true;
  const hay = _norm(Object.values(r).join(' '));
  return tokens.every(t => hay.includes(t));
}
const C   = CFG.color;
const CA  = CFG.colorAlpha;

let filtered = [...ROWS];
let sortDir  = {};

document.getElementById('badge-total').textContent = fmt(ROWS.length);

// ── Dropdowns ──────────────────────────────────────────────────────────────
function uniq(arr) {
  return [...new Set(arr.filter(v => v && String(v).trim() !== ''))].sort();
}
function addOpts(id, vals) {
  const el = document.getElementById(id);
  if (!el) return;
  vals.forEach(v => { const o = document.createElement('option'); o.value = o.textContent = v; el.appendChild(o); });
}
function initAutocomplete(inputId, menuId, options, onChange) {
  const inp = document.getElementById(inputId), menu = document.getElementById(menuId);
  if (!inp || !menu) return;
  function render() {
    const q = _norm(inp.value).trim();
    const matches = (q ? options.filter(o => _norm(o).includes(q)) : options).slice(0, 50);
    menu.innerHTML = matches.length
      ? matches.map(o => '<div class="ac-item">' + o.replace(/&/g,'&amp;').replace(/</g,'&lt;') + '</div>').join('')
      : '<div class="ac-empty">Sin coincidencias</div>';
    menu.classList.add('show');
  }
  inp.addEventListener('focus', render);
  inp.addEventListener('input', () => { render(); onChange(); });
  menu.addEventListener('mousedown', e => {
    const it = e.target.closest('.ac-item'); if (!it) return;
    e.preventDefault();
    inp.value = it.textContent;
    menu.classList.remove('show');
    onChange();
  });
  document.addEventListener('click', e => {
    if (e.target !== inp && !menu.contains(e.target)) menu.classList.remove('show');
  });
}

addOpts('f-sector',   uniq(ROWS.map(r => r['SECTOR'])));
addOpts('f-depto',    uniq(ROWS.map(r => r['DEPARTAMENTO_OFERTA_PROGRAMA'])));
addOpts('f-division', uniq(ROWS.map(r => r['DIVISIÓN UNINORTE'])));
initAutocomplete('f-institucion', 'f-institucion-menu', uniq(ROWS.map(r => r['NOMBRE_INSTITUCIÓN'])), applyFilters);

// Sort fechas newest-first (DD/MM/YYYY)
const parseFecha = s => { try { const [d,m,y]=s.split('/'); return new Date(+y,+m-1,+d); } catch(e){return new Date(0);} };
const sortedFechas = uniq(ROWS.map(r => r['FECHA_OBTENCION'])).sort((a,b) => parseFecha(b)-parseFecha(a));
addOpts('f-fecha', sortedFechas);

if (CFG.tipo === 'modificados') {
  const cs = new Set();
  ROWS.forEach(r => { (r['QUE_CAMBIO']||'').split(' | ').forEach(p => { const f=p.split(':')[0].trim(); if(f&&f!=='nan'&&f!=='') cs.add(f); }); });
  addOpts('f-tipo-cambio', [...cs].sort());
} else {
  addOpts('f-modalidad', uniq(ROWS.map(r => r['MODALIDAD'])));
  addOpts('f-cine', uniq(ROWS.map(r => (r['CINE_F_2013_AC_CAMPO_ESPECÍFIC']||'').trim())));
}

// ── Filters ────────────────────────────────────────────────────────────────
function gv(id) { const el=document.getElementById(id); return el?el.value:''; }

function applyFilters() {
  const qTokens = _norm(gv('f-q')).split(/\\s+/).filter(Boolean);
  const se = gv('f-sector'), de = gv('f-depto'), di = gv('f-division');
  const mo = gv('f-modalidad'), fe = gv('f-fecha'), tc = gv('f-tipo-cambio');
  const ci = gv('f-cine');
  const insTokens = _norm(gv('f-institucion')).split(/\\s+/).filter(Boolean);

  filtered = ROWS.filter(r => {
    if (!_rowMatches(r, qTokens)) return false;
    if (se && r['SECTOR'] !== se) return false;
    if (de && r['DEPARTAMENTO_OFERTA_PROGRAMA'] !== de) return false;
    if (di && r['DIVISIÓN UNINORTE'] !== di) return false;
    if (mo && r['MODALIDAD'] !== mo) return false;
    if (fe && r['FECHA_OBTENCION'] !== fe) return false;
    if (tc && !(r['QUE_CAMBIO']||'').includes(tc)) return false;
    if (ci && (r['CINE_F_2013_AC_CAMPO_ESPECÍFIC']||'').trim() !== ci) return false;
    if (insTokens.length) {
      const hayIns = _norm(r['NOMBRE_INSTITUCIÓN']);
      if (!insTokens.every(t => hayIns.includes(t))) return false;
    }
    return true;
  });
  renderAll(filtered);
}

function resetFilters() {
  ['f-q','f-sector','f-depto','f-institucion','f-division','f-modalidad','f-fecha','f-tipo-cambio','f-cine'].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = '';
  });
  _periodosDet=null;
  applyFilters();
}

// ── Distribution helpers ───────────────────────────────────────────────────
function countBy(rows, field, n) {
  const c={};
  rows.forEach(r => { const v=(r[field]||'Sin datos').toString().trim()||'Sin datos'; c[v]=(c[v]||0)+1; });
  return Object.entries(c).sort((a,b)=>b[1]-a[1]).slice(0,n||12);
}
// ── Shared date helper ─────────────────────────────────────────────────────
function getSem(s) {
  if(!s||!s.trim()) return null;
  let y,m;
  const iso=s.match(/^(\\d{4})-(\\d{2})/);
  if(iso){y=+iso[1];m=+iso[2];}
  else{const dmy=s.match(/^(\\d{2})\\/(\\d{2})\\/(\\d{4})/);if(dmy){y=+dmy[3];m=+dmy[2];}}
  if(!y||y<2014||y>2035) return null;
  return y+'-'+(m<=6?'1':'2');
}

// ── Charts ─────────────────────────────────────────────────────────────────
function _emptyChart(el, msg) {
  Plotly.purge(el);
  el.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#64748b;font-size:.82rem">'+(msg||'Sin datos para los filtros aplicados')+'</div>';
}

function plotDonut(id, rows, field) {
  const el=document.getElementById(id); if(!el) return;
  const d=countBy(rows,field,20); if(!d.length){ _emptyChart(el); return; }
  Plotly.react(id,[{labels:d.map(e=>e[0]),values:d.map(e=>e[1]),type:'pie',hole:.45,
    marker:{colors:['#2d5b9e','#fcc10e','#bd900b','#ae1e22','#6e91b9','#214174','#d56f18']},
    textinfo:'label+percent',hovertemplate:'%{label}<br><b>%{value:,}</b><extra></extra>'}],
    {margin:{t:10,r:10,b:35,l:10},showlegend:true,legend:{orientation:'h',y:-0.12,font:{size:10}},
    plot_bgcolor:'white',paper_bgcolor:'white'},PC);
}

function plotHBar(id, rows, field, color, n, maxLen) {
  const el=document.getElementById(id); if(!el) return;
  const d=[...countBy(rows,field,n||10)].reverse(); if(!d.length){ _emptyChart(el); return; }
  const trunc=s=>maxLen&&s.length>maxLen?s.slice(0,maxLen)+'…':s;
  const labels=d.map(e=>trunc(e[0]));
  const full=d.map(e=>e[0]);
  const lMargin=maxLen?Math.min(maxLen*6.5,200):210;
  Plotly.react(id,[{y:labels,x:d.map(e=>e[1]),customdata:full,type:'bar',orientation:'h',
    marker:{color,opacity:.85},
    hovertemplate:'%{customdata}<br><b>%{x:,}</b><extra></extra>'}],
    {margin:{t:10,r:20,b:30,l:lMargin},
    xaxis:{showgrid:true,gridcolor:'#e2e8f0',tickfont:{size:10}},
    yaxis:{tickfont:{size:10},automargin:false},
    plot_bgcolor:'white',paper_bgcolor:'white',bargap:.3},PC);
}

function plotVBar(id, rows, field, color) {
  const el=document.getElementById(id); if(!el) return;
  const d=countBy(rows,field,10); if(!d.length){ _emptyChart(el); return; }
  Plotly.react(id,[{x:d.map(e=>e[0]),y:d.map(e=>e[1]),type:'bar',
    marker:{color,opacity:.85},hovertemplate:'%{x}<br><b>%{y:,}</b><extra></extra>'}],
    {margin:{t:10,r:10,b:80,l:45},
    xaxis:{tickfont:{size:10},tickangle:-30},
    yaxis:{showgrid:true,gridcolor:'#e2e8f0'},
    plot_bgcolor:'white',paper_bgcolor:'white',bargap:.35},PC);
}

function plotTimeline(id, rows) {
  const el=document.getElementById(id); if(!el) return;
  const bySem={};
  rows.forEach(r=>{const s=getSem(r['FECHA_OBTENCION']); if(s) bySem[s]=(bySem[s]||0)+1;});
  const sems=Object.keys(bySem).sort();
  if(!sems.length){ _emptyChart(el, 'Sin datos de fecha de run'); return; }
  const x=sems, y=sems.map(s=>bySem[s]);
  Plotly.react(id,[{x,y,type:'scatter',mode:'lines+markers',
    line:{color:C,width:2.5},marker:{color:C,size:7},
    fill:'tozeroy',fillcolor:CA,
    hovertemplate:'%{x}<br><b>%{y:,}</b><extra></extra>'}],
    {margin:{t:10,r:20,b:40,l:50},
    xaxis:{showgrid:false,tickfont:{size:11},type:'category'},
    yaxis:{showgrid:true,gridcolor:'#e2e8f0',rangemode:'tozero'},
    plot_bgcolor:'white',paper_bgcolor:'white',hovermode:'x unified'},PC);
}

function plotAcumuladoModalidad(id, rows) {
  const el=document.getElementById(id); if(!el) return;
  if(!rows.length){ _emptyChart(el); return; }

  const semSet=new Set();
  rows.forEach(r=>{const s=getSem(r['FECHA_DE_REGISTRO_EN_SNIES']);if(s)semSet.add(s);});
  const sems=[...semSet].sort();

  if(!sems.length){
    _emptyChart(el, 'Sin datos de fecha de registro en SNIES');
    return;
  }

  const mods=[...new Set(rows.map(r=>r['MODALIDAD']).filter(v=>v&&v.trim()))].sort();
  const COLORS=['#2d5b9e','#fcc10e','#bd900b','#ae1e22','#6e91b9','#214174','#d56f18'];

  const traces=mods.map((mod,i)=>{
    const bySem={};
    rows.filter(r=>r['MODALIDAD']===mod).forEach(r=>{
      const s=getSem(r['FECHA_DE_REGISTRO_EN_SNIES']); if(s)bySem[s]=(bySem[s]||0)+1;
    });
    let cum=0; const x=[],y=[];
    sems.forEach(s=>{cum+=(bySem[s]||0);x.push(s);y.push(cum);});
    if(cum===0) return null;
    return{x,y,name:mod,type:'scatter',mode:'lines+markers',
      line:{color:COLORS[i%COLORS.length],width:2.5},
      marker:{color:COLORS[i%COLORS.length],size:6},
      hovertemplate:mod+'<br><b>%{x}</b><br>%{y:,} acumulados<extra></extra>'};
  }).filter(Boolean);

  if(!traces.length){ _emptyChart(el); return; }
  Plotly.react(id,traces,{
    margin:{t:10,r:20,b:65,l:55},
    xaxis:{showgrid:false,tickfont:{size:11},tickangle:-30,type:'category'},
    yaxis:{showgrid:true,gridcolor:'#e2e8f0',rangemode:'tozero',tickfont:{size:11}},
    plot_bgcolor:'white',paper_bgcolor:'white',
    hovermode:'x unified',
    legend:{orientation:'h',y:-0.28,font:{size:11}}
  },PC);
}

// ── CINE cumulative chart (nuevos/inactivos) ───────────────────────────────
const CINE_COLORS=['#2d5b9e','#fcc10e','#bd900b','#ae1e22','#6e91b9',
                   '#214174','#d56f18','#948e56','#15284b','#7a1518'];
const CINE_COL='CINE_F_2013_AC_CAMPO_ESPECÍFIC';
let _cineInit=false, _cineAll=[], _cineActive=[];

function initCINE() {
  const counts={};
  ROWS.forEach(r=>{
    const v=(r[CINE_COL]||'').trim()||'Sin clasificar';
    counts[v]=(counts[v]||0)+1;
  });
  _cineAll=Object.keys(counts).sort((a,b)=>counts[b]-counts[a]);
  _cineActive=[..._cineAll.slice(0,8)];
  const dl=document.getElementById('cine-list');
  if(dl){
    dl.innerHTML='';
    _cineAll.forEach(c=>{const o=document.createElement('option');o.value=c;dl.appendChild(o);});
  }
  _cineInit=true;
}

function renderCineChart(rows) {
  const el=document.getElementById('ch-timeline'); if(!el) return;
  const semSet=new Set();
  rows.forEach(r=>{const s=getSem(r['FECHA_DE_REGISTRO_EN_SNIES']);if(s)semSet.add(s);});
  const allSems=[...semSet].sort();
  const sems=allSems.filter(s=>s>='2023-2');

  const tagsEl=document.getElementById('cine-tags');
  if(tagsEl){
    tagsEl.innerHTML=_cineActive.map((cine,i)=>{
      const col=CINE_COLORS[i%CINE_COLORS.length];
      const safe=cine.replace(/\\\\/g,'\\\\\\\\').replace(/'/g,"\\\\'");
      return '<span title="Clic para filtrar tabla" onclick="setCineFiltro(\\''+safe+'\\')" style="display:inline-flex;align-items:center;gap:.3rem;'+
        'background:'+col+'22;border:1px solid '+col+'66;border-radius:2rem;'+
        'padding:.18rem .55rem;font-size:.72rem;color:'+col+';font-weight:500;cursor:pointer">'+
        cine+
        '<button onclick="event.stopPropagation();cineRemove(\\''+safe+'\\')" '+
        'style="background:none;border:none;cursor:pointer;color:'+col+';'+
        'font-size:.9rem;padding:0 0 0 .2rem;line-height:1;opacity:.75">\\xd7</button></span>';
    }).join('');
  }

  if(!sems.length){
    _emptyChart(el, 'Sin datos de fecha de registro');
    return;
  }

  const traces=_cineActive.map((cine,i)=>{
    const isUnclass=cine==='Sin clasificar';
    const bySem={};
    rows.filter(r=>{const v=(r[CINE_COL]||'').trim();return isUnclass?!v:v===cine;})
        .forEach(r=>{const s=getSem(r['FECHA_DE_REGISTRO_EN_SNIES']);if(s)bySem[s]=(bySem[s]||0)+1;});
    let cum=0;const x=[],y=[];
    allSems.forEach(s=>{cum+=(bySem[s]||0);if(s>='2023-2'){x.push(s);y.push(cum);}});
    if(cum===0) return null;
    const col=CINE_COLORS[i%CINE_COLORS.length];
    return{x,y,name:cine,type:'scatter',mode:'lines+markers',
      line:{color:col,width:2.5},marker:{color:col,size:6},
      hovertemplate:cine+'<br><b>%{x}</b><br>%{y:,} acumulados<extra></extra>'};
  }).filter(Boolean);

  if(!traces.length){ _emptyChart(el, 'Sin datos de fecha de registro'); return; }
  Plotly.react('ch-timeline',traces,{
    margin:{t:10,r:20,b:65,l:55},
    xaxis:{showgrid:false,tickfont:{size:11},tickangle:-30,type:'category'},
    yaxis:{showgrid:true,gridcolor:'#e2e8f0',rangemode:'tozero',tickfont:{size:11}},
    plot_bgcolor:'white',paper_bgcolor:'white',
    hovermode:'x unified',
    legend:{orientation:'h',y:-0.28,font:{size:11}}
  },PC).then(gd=>{
    gd.removeAllListeners('plotly_click');
    gd.removeAllListeners('plotly_legendclick');
    gd.on('plotly_click',ev=>{if(ev.points&&ev.points.length)setCineFiltro(ev.points[0].data.name);});
    gd.on('plotly_legendclick',data=>{setCineFiltro(gd.data[data.curveNumber].name);return false;});
  });
}

function plotAcumuladoCINE(id, rows) {
  if(!_cineInit) initCINE();
  renderCineChart(rows);
}

function cineAdd() {
  const inp=document.getElementById('cine-search'); if(!inp) return;
  const val=inp.value.trim();
  if(!val||_cineActive.includes(val)||!_cineAll.includes(val)){inp.value='';return;}
  _cineActive.push(val);
  inp.value='';
  renderCineChart(filtered);
}

function cineRemove(cine) {
  _cineActive=_cineActive.filter(c=>c!==cine);
  renderCineChart(filtered);
}

const PERIODOS_TOPE = 12;
function _bucketPeriodo(k) { return k >= PERIODOS_TOPE ? PERIODOS_TOPE + '+' : String(k); }
function _periodoSortKey(l) { return l === PERIODOS_TOPE + '+' ? Infinity : parseInt(l, 10); }
function _matchesPeriodoDet(v, filtro) {
  if (!v) return false;
  const n = Math.round(parseFloat(String(v)));
  if (isNaN(n)) return false;
  return filtro === PERIODOS_TOPE + '+' ? n >= PERIODOS_TOPE : n === parseInt(filtro, 10);
}

function plotPeriodos(id, rows) {
  const el=document.getElementById(id); if(!el) return;
  const COL='NÚMERO_PERIODOS_DE_DURACIÓN', PCOL='PERIODICIDAD';
  const PCOLORS=['#2d5b9e','#fcc10e','#bd900b','#ae1e22','#6e91b9',
                 '#214174','#d56f18','#948e56','#15284b','#7a1518'];
  const pivot={}, periSet=new Set();
  rows.forEach(r=>{
    const v=r[COL]; if(!v||String(v).trim()==='') return;
    const k0=Math.round(parseFloat(v)); if(isNaN(k0)) return;
    const k=_bucketPeriodo(k0);
    const p=(r[PCOL]||'Sin definir').trim()||'Sin definir';
    if(!pivot[k]) pivot[k]={};
    pivot[k][p]=(pivot[k][p]||0)+1;
    periSet.add(p);
  });
  const labels=Object.keys(pivot).sort((a,b)=>_periodoSortKey(a)-_periodoSortKey(b));
  if(!labels.length){ _emptyChart(el); return; }
  const peris=[...periSet].sort((a,b)=>{
    const ta=labels.reduce((s,l)=>s+(pivot[l][a]||0),0);
    const tb=labels.reduce((s,l)=>s+(pivot[l][b]||0),0);
    return tb-ta;
  });
  const traces=peris.map((p,i)=>({
    x:labels, y:labels.map(l=>pivot[l][p]||0), name:p, type:'bar',
    marker:{color:PCOLORS[i%PCOLORS.length],opacity:.85},
    hovertemplate:p+'<br>%{x} periodos — <b>%{y:,}</b> programas<extra></extra>'
  }));
  Plotly.react(id, traces, {
    barmode:'stack',
    margin:{t:10,r:20,b:90,l:60},
    xaxis:{title:'Periodos',type:'category',tickmode:'array',tickvals:labels,tickfont:{size:11}},
    yaxis:{title:'N. Programas',showgrid:true,gridcolor:'#e2e8f0',tickfont:{size:11}},
    plot_bgcolor:'white',paper_bgcolor:'white',bargap:.25,
    legend:{orientation:'h',y:-0.28,font:{size:11},entrywidth:120,entrywidthmode:'pixels'},
    hovermode:'x unified'
  }, PC).then(gd=>{
    gd.removeAllListeners('plotly_click');
    gd.on('plotly_click', ev=>setPeriodosDetalleFilter(ev.points[0].x));
  });
}

let _periodosDet=null;

function _applySubFilters(rows) {
  let r=rows;
  if(_periodosDet!==null){
    const COL='NÚMERO_PERIODOS_DE_DURACIÓN';
    r=r.filter(x=>_matchesPeriodoDet(x[COL], _periodosDet));
  }
  return r;
}

function setPeriodosDetalleFilter(val) {
  val = String(val);
  if(_periodosDet===val){clearPeriodosDetalleFilter();return;}
  _periodosDet=val;
  applyFilters();
  document.getElementById('tbl-wrap').scrollIntoView({behavior:'smooth'});
}

function clearPeriodosDetalleFilter() {
  _periodosDet=null;
  applyFilters();
}

function setCineFiltro(cine) {
  const el=document.getElementById('f-cine'); if(!el) return;
  if(el.value===cine){clearCineFiltro();return;}
  el.value=cine;
  applyFilters();
  document.getElementById('tbl-wrap').scrollIntoView({behavior:'smooth'});
}

function clearCineFiltro() {
  const el=document.getElementById('f-cine'); if(el) el.value='';
  applyFilters();
}

function plotTipoCambio(id, rows) {
  const el=document.getElementById(id); if(!el) return;
  const c={};
  rows.forEach(r => { (r['QUE_CAMBIO']||'').split(' | ').forEach(p => { const f=p.split(':')[0].trim(); if(f&&f!=='nan'&&f!=='') c[f]=(c[f]||0)+1; }); });
  const d=Object.entries(c).sort((a,b)=>b[1]-a[1]); if(!d.length){ _emptyChart(el); return; }
  Plotly.react(id,[{x:d.map(e=>e[0]),y:d.map(e=>e[1]),type:'bar',
    marker:{color:'#bd900b',opacity:.85},
    hovertemplate:'%{x}<br><b>%{y:,}</b> cambios<extra></extra>'}],
    {margin:{t:10,r:20,b:80,l:50},
    xaxis:{tickfont:{size:11},tickangle:-20},
    yaxis:{showgrid:true,gridcolor:'#e2e8f0'},
    plot_bgcolor:'white',paper_bgcolor:'white',bargap:.4},PC);
}

function plotScatter(id, rows) {
  const el=document.getElementById(id); if(!el) return;
  const pts=rows.map(r=>({
    x:parseFloat(r['NÚMERO_CRÉDITOS_ANTERIOR']),
    y:parseFloat(r['NÚMERO_CRÉDITOS']),
    t:r['NOMBRE_DEL_PROGRAMA']||''
  })).filter(p=>!isNaN(p.x)&&!isNaN(p.y)&&p.x!==p.y);
  if(!pts.length){ _emptyChart(el, 'Sin datos de creditos para comparar'); return; }
  const vals=pts.flatMap(p=>[p.x,p.y]);
  const mn=Math.min(...vals), mx=Math.max(...vals);
  Plotly.react(id,[
    {x:pts.map(p=>p.x),y:pts.map(p=>p.y),text:pts.map(p=>p.t),
     mode:'markers',type:'scatter',
     marker:{color:'#bd900b',size:8,opacity:.7},
     hovertemplate:'<b>%{text}</b><br>Antes: %{x} creditos<br>Despues: %{y} creditos<extra></extra>'},
    {x:[mn,mx],y:[mn,mx],mode:'lines',
     line:{color:'#9fb0c9',width:1,dash:'dot'},hoverinfo:'skip',showlegend:false}
  ],{margin:{t:10,r:20,b:50,l:60},
    xaxis:{title:'Creditos anteriores',showgrid:true,gridcolor:'#e2e8f0'},
    yaxis:{title:'Creditos actuales',showgrid:true,gridcolor:'#e2e8f0'},
    plot_bgcolor:'white',paper_bgcolor:'white'},PC);
}

// ── Table ──────────────────────────────────────────────────────────────────
const COL_HEAD = {
  'FECHA_OBTENCION':'Fecha','CÓDIGO_SNIES_DEL_PROGRAMA':'Cód. SNIES',
  'NOMBRE_DEL_PROGRAMA':'Programa','NOMBRE_INSTITUCIÓN':'Institución',
  'SECTOR':'Sector','MODALIDAD':'Modalidad',
  'DEPARTAMENTO_OFERTA_PROGRAMA':'Departamento','MUNICIPIO_OFERTA_PROGRAMA':'Municipio',
  'NÚMERO_CRÉDITOS':'Créditos','COSTO_MATRÍCULA_ESTUD_NUEVOS':'Costo matrícula',
  'PERIODICIDAD':'Periodicidad','DIVISIÓN UNINORTE':'División Uninorte',
  'QUE_CAMBIO':'¿Qué cambió?','NÚMERO_CRÉDITOS_ANTERIOR':'Créd. anteriores',
  'NIVEL_DE_FORMACIÓN':'Nivel de formación'
};
const TBL_COLS = CFG.cols;

function buildTbl(rows) {
  const cols = TBL_COLS.filter(c => !rows.length || c in rows[0]);
  let h = '<table><thead><tr>';
  cols.forEach(c => {
    h += '<th onclick="sortTbl(\\'' + c.replace(/'/g,"\\\\'") + '\\')">'+(COL_HEAD[c]||c)+' <span style="opacity:.4">↕</span></th>';
  });
  h += '</tr></thead><tbody>';
  if (!rows.length) {
    h += '<tr><td colspan="'+cols.length+'" class="empty">Sin registros para los filtros seleccionados</td></tr>';
  } else {
    rows.forEach(r => {
      h += '<tr>' + cols.map(c => '<td>'+(r[c]||'')+'</td>').join('') + '</tr>';
    });
  }
  return h + '</tbody></table>';
}

function renderTbl(rows) {
  document.getElementById('tbl-wrap').innerHTML = buildTbl(rows);
}

function sortTbl(col) {
  sortDir[col] = !sortDir[col];
  filtered = [...filtered].sort((a,b) => {
    const va=a[col]||'', vb=b[col]||'';
    const cmp = String(va).localeCompare(String(vb),'es',{numeric:true});
    return sortDir[col] ? cmp : -cmp;
  });
  renderTbl(_applySubFilters(filtered));
}

// ── Render all ─────────────────────────────────────────────────────────────
function renderAll(rows) {
  const subRows=_applySubFilters(rows);
  const cpd=document.getElementById('per-det-chip');
  const ccd=document.getElementById('cine-det-chip');
  if(cpd){
    cpd.style.display=_periodosDet!==null?'flex':'none';
    if(_periodosDet!==null) document.getElementById('per-det-val').textContent=
      (_periodosDet===PERIODOS_TOPE+'+' ? PERIODOS_TOPE+' o más periodos' : _periodosDet+' periodos')+
      ' — '+subRows.length.toLocaleString('es-CO')+' programas';
  }
  const cineFiltroVal=gv('f-cine');
  if(ccd){
    ccd.style.display=cineFiltroVal?'flex':'none';
    if(cineFiltroVal) document.getElementById('cine-det-val').textContent=
      cineFiltroVal+' — '+subRows.length.toLocaleString('es-CO')+' programas';
  }
  document.getElementById('f-count').textContent = fmt(subRows.length) + ' programas';

  plotDonut('ch-sector', rows, 'SECTOR');
  plotHBar('ch-instituciones', rows, 'NOMBRE_INSTITUCIÓN', C, 10, 32);
  plotHBar('ch-division',      rows, 'DIVISIÓN UNINORTE',  C, 12);
  plotHBar('ch-depto',         rows, 'DEPARTAMENTO_OFERTA_PROGRAMA', '#214174', 15);

  if (CFG.tipo === 'modificados') {
    plotTipoCambio('ch-tipo-cambio', rows);
    plotScatter('ch-scatter', rows);
    plotTimeline('ch-timeline', rows);
  } else if (CFG.tipo === 'nuevos' || CFG.tipo === 'inactivos') {
    plotVBar('ch-modalidad', rows, 'MODALIDAD', C);
    plotAcumuladoCINE('ch-timeline', rows);
    plotPeriodos('ch-periodos', rows);
  } else {
    plotVBar('ch-modalidad', rows, 'MODALIDAD', C);
    plotAcumuladoModalidad('ch-timeline', rows);
  }

  renderTbl(subRows);
}

// ── Init ───────────────────────────────────────────────────────────────────
applyFilters();
</script>
</body>
</html>
"""

CHARTS_NUEVOS_INACTIVOS = """<div class="g2">
  <div class="card"><div class="ct">Por sector</div><div id="ch-sector" style="height:260px"></div></div>
  <div class="card"><div class="ct">Top 10 instituciones</div><div id="ch-instituciones" style="height:260px"></div></div>
</div>
<div class="g2">
  <div class="card"><div class="ct">Por Division Uninorte</div><div id="ch-division" style="height:260px"></div></div>
  <div class="card"><div class="ct">Por modalidad</div><div id="ch-modalidad" style="height:260px"></div></div>
</div>
<div class="card"><div class="ct">Top 15 departamentos de oferta</div><div id="ch-depto" style="height:310px"></div></div>
<div class="card">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:.75rem;flex-wrap:wrap;gap:.5rem">
    <div class="ct" style="margin-bottom:0">Acumulado por campo CINE semestral (fecha de registro SNIES)</div>
    <div style="display:flex;gap:.5rem;align-items:center">
      <input id="cine-search" list="cine-list" placeholder="Buscar campo CINE…"
             style="padding:.35rem .7rem;border:1px solid #cbd5e1;border-radius:.4rem;font-size:.78rem;width:260px;outline:none"
             onkeydown="if(event.key==='Enter')cineAdd()">
      <datalist id="cine-list"></datalist>
      <button onclick="cineAdd()"
              style="padding:.35rem .8rem;background:#2d5b9e;color:#fff;border:none;border-radius:.4rem;font-size:.78rem;cursor:pointer;white-space:nowrap">+ Agregar</button>
    </div>
  </div>
  <div id="cine-tags" style="display:flex;flex-wrap:wrap;gap:.35rem;margin-bottom:.6rem;min-height:1.4rem"></div>
  <div id="ch-timeline" style="height:300px"></div>
</div>
<div class="card"><div class="ct">Distribución por duración (periodos requeridos)</div><div id="ch-periodos" style="height:260px"></div></div>"""

CHARTS_MODIFICADOS = """<div class="card"><div class="ct">Tipo de cambio detectado</div><div id="ch-tipo-cambio" style="height:260px"></div></div>
<div class="g2">
  <div class="card"><div class="ct">Top 10 instituciones modificadas</div><div id="ch-instituciones" style="height:260px"></div></div>
  <div class="card"><div class="ct">Por Division Uninorte</div><div id="ch-division" style="height:260px"></div></div>
</div>
<div class="g2">
  <div class="card"><div class="ct">Creditos: antes vs despues</div><div id="ch-scatter" style="height:300px"></div></div>
  <div class="card"><div class="ct">Por sector</div><div id="ch-sector" style="height:300px"></div></div>
</div>
<div class="card"><div class="ct">Top 15 departamentos afectados</div><div id="ch-depto" style="height:310px"></div></div>
<div class="card"><div class="ct">Modificaciones por periodo de run</div><div id="ch-timeline" style="height:200px"></div></div>"""


# ── Página de análisis: créditos ──────────────────────────────────────────────

CAMPO_CREDITOS_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Analisis de Creditos . SNIES Monitor</title>
__PLOTLY_CDN__
<style>
:root{
  --bg:#f1f5f9;--surface:#fff;--text:#0f172a;--muted:#64748b;
  --border:#e2e8f0;--blue:#2d5b9e;--green:#1a9e6b;--red:#ae1e22;--amber:#bd900b;
  --radius:0.75rem;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);font-size:14px}
header{background:linear-gradient(135deg,#15284b,#bd900b);color:#fff;
  padding:1.2rem 2rem;display:flex;justify-content:space-between;align-items:center;gap:1rem;flex-wrap:wrap}
header h1{font-size:1.25rem;font-weight:700}
header .sub{font-size:.77rem;opacity:.75;margin-top:.2rem}
.back-btn{display:inline-flex;align-items:center;gap:.3rem;background:rgba(255,255,255,.2);
  border:1px solid rgba(255,255,255,.35);color:#fff;text-decoration:none;padding:.38rem .85rem;
  border-radius:.4rem;font-size:.8rem;font-weight:500;white-space:nowrap;transition:background .15s}
.back-btn:hover{background:rgba(255,255,255,.32)}
.filter-bar{position:sticky;top:0;z-index:100;background:var(--surface);
  border-bottom:1px solid var(--border);padding:.6rem 2rem;
  display:flex;gap:.45rem;flex-wrap:wrap;align-items:center;
  box-shadow:0 2px 8px rgba(0,0,0,.06)}
.f-input{flex:1;min-width:200px;padding:.5rem .8rem;border:1px solid var(--border);
  border-radius:.4rem;font-size:.8rem;outline:none}
.f-input:focus{border-color:var(--blue)}
.f-sel{padding:.4rem .6rem;border:1px solid var(--border);border-radius:.4rem;
  font-size:.77rem;background:var(--surface);outline:none;cursor:pointer;max-width:200px}
.f-btn{padding:.4rem .85rem;border:1px solid var(--border);border-radius:.4rem;
  font-size:.77rem;background:var(--surface);cursor:pointer;color:var(--muted);white-space:nowrap}
.f-btn:hover{background:var(--bg)}
.f-count{margin-left:auto;font-size:.82rem;font-weight:600;color:var(--blue);white-space:nowrap}
.ac-wrap{position:relative;flex:0 1 240px}
.ac-menu{position:absolute;top:calc(100% + 2px);left:0;right:0;z-index:300;
  background:var(--surface);border:1px solid var(--border);border-radius:.4rem;
  box-shadow:0 8px 20px rgba(0,0,0,.14);max-height:260px;overflow-y:auto;display:none}
.ac-menu.show{display:block}
.ac-item{padding:.45rem .75rem;font-size:.78rem;cursor:pointer;color:var(--text)}
.ac-item:hover{background:var(--bg)}
.ac-empty{padding:.45rem .75rem;font-size:.78rem;color:var(--muted)}
main{max-width:1380px;margin:0 auto;padding:1.5rem 2rem}
section{margin-bottom:1.25rem}
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:1rem}
.kpi{background:var(--surface);border-radius:var(--radius);padding:1.1rem 1.4rem;
  box-shadow:0 1px 3px rgba(0,0,0,.07);border-left:4px solid var(--blue)}
.kpi.r{border-left-color:var(--red)}.kpi.g{border-left-color:var(--green)}.kpi.a{border-left-color:var(--amber)}
.kpi-label{font-size:.68rem;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:.4rem}
.kpi-val{font-size:1.8rem;font-weight:700;line-height:1}
.kpi-sub{font-size:.7rem;color:var(--muted);margin-top:.35rem}
.card{background:var(--surface);border-radius:var(--radius);padding:1.2rem;
  box-shadow:0 1px 3px rgba(0,0,0,.07);margin-bottom:1rem}
.ct{font-size:.68rem;font-weight:600;text-transform:uppercase;letter-spacing:.06em;
  color:var(--muted);margin-bottom:.85rem}
.ct-note{font-size:.74rem;color:var(--muted);margin-top:-.5rem;margin-bottom:.75rem;line-height:1.4}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:1rem}
.tbl-wrap{max-height:480px;overflow-y:auto;border:1px solid var(--border);border-radius:.5rem}
table{width:100%;border-collapse:collapse;font-size:.77rem}
th{background:var(--bg);padding:.6rem .85rem;text-align:left;font-size:.67rem;
  text-transform:uppercase;letter-spacing:.05em;color:var(--muted);cursor:pointer;
  user-select:none;position:sticky;top:0;z-index:1;white-space:nowrap}
th:hover{background:#e2e8f0}
td{padding:.6rem .85rem;border-bottom:1px solid var(--border);vertical-align:top;
  max-width:270px;word-break:break-word}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8fafc}
.delta-up{color:var(--green);font-weight:600}
.delta-down{color:var(--red);font-weight:600}
.empty{text-align:center;color:var(--muted);padding:2.5rem}
@media(max-width:900px){
  .kpi-grid{grid-template-columns:repeat(2,1fr)}
  .g2{grid-template-columns:1fr}
  main{padding:1rem}
  header{flex-direction:column;text-align:center}
  .f-count{margin-left:0}
}
</style>
</head>
<body>
<header>
  <div style="display:flex;align-items:center;gap:.9rem">
    <a href="modificados.html" class="back-btn">← Modificados</a>
    <div>
      <h1>📐 Analisis de Creditos</h1>
      <div class="sub">Que cambia cuando un programa modifica su numero de creditos</div>
    </div>
  </div>
  <a href="modificados_costos.html" class="back-btn">💰 Analisis de Costos →</a>
</header>

<div class="filter-bar">
  <input id="f-q" class="f-input" placeholder="Buscar por nombre, institucion, departamento... (filtra TODA la pagina)" oninput="applyFilters()">
  <select id="f-sector" class="f-sel" onchange="applyFilters()"><option value="">Todos los sectores</option></select>
  <select id="f-depto"  class="f-sel" onchange="applyFilters()"><option value="">Todos los departamentos</option></select>
  <div class="ac-wrap">
    <input id="f-institucion" class="f-sel" placeholder="Buscar institucion..." style="cursor:text;width:100%" oninput="applyFilters()">
    <div class="ac-menu" id="f-institucion-menu"></div>
  </div>
  <button class="f-btn" onclick="resetFilters()">✕ Limpiar</button>
  <span class="f-count" id="f-count">–</span>
</div>

<main>
  <section class="kpi-grid">
    <div class="kpi">
      <div class="kpi-label">Cambios de creditos detectados</div>
      <div class="kpi-val" id="k-total">–</div>
      <div class="kpi-sub">programas con creditos antes/despues distintos</div>
    </div>
    <div class="kpi g">
      <div class="kpi-label">delta promedio . Oficial</div>
      <div class="kpi-val" id="k-oficial">–</div>
      <div class="kpi-sub" id="k-oficial-sub">creditos por cambio</div>
    </div>
    <div class="kpi r">
      <div class="kpi-label">delta promedio . Privado</div>
      <div class="kpi-val" id="k-privado">–</div>
      <div class="kpi-sub" id="k-privado-sub">creditos por cambio</div>
    </div>
    <div class="kpi a">
      <div class="kpi-label">Tambien cambia la duracion</div>
      <div class="kpi-val" id="k-cocambio">–</div>
      <div class="kpi-sub" id="k-cocambio-sub">vs. tasa general</div>
    </div>
  </section>
  <div class="ct-note" style="margin:-.5rem 0 1rem">
    "delta promedio" = cambio promedio en numero de creditos (creditos despues - creditos antes) entre los programas
    modificados de ese grupo. <strong>Negativo</strong> significa que, en promedio, le quitan creditos al programa;
    <strong>positivo</strong>, que le agregan.
  </div>

  <section class="g2">
    <div class="card">
      <div class="ct">Sube vs. baja, por sector</div>
      <div id="ch-sector" style="height:280px"></div>
    </div>
    <div class="card">
      <div class="ct">Distribucion del cambio (creditos ganados/perdidos)</div>
      <div id="ch-hist" style="height:280px"></div>
    </div>
  </section>

  <section class="g2">
    <div class="card">
      <div class="ct">Top instituciones que mas cambian creditos</div>
      <div class="ct-note">Numero al lado de la barra: delta promedio de esa institucion (ver definicion arriba).</div>
      <div id="ch-instituciones" style="height:380px"></div>
    </div>
    <div class="card">
      <div class="ct">Top departamentos que mas cambian creditos</div>
      <div class="ct-note">Numero al lado de la barra: delta promedio de ese departamento (ver definicion arriba).</div>
      <div id="ch-departamentos" style="height:380px"></div>
    </div>
  </section>

  <section class="card">
    <div class="ct">Por campo CINE: que programas suben o bajan creditos</div>
    <div class="ct-note">Filtra por institucion arriba para ver el detalle de su oferta. Cada barra es un campo
      CINE; el valor es el balance neto de creditos ganados/perdidos entre los programas de ese campo
      (suma de los deltas). Verde = el campo gana creditos en neto; rojo = los pierde.</div>
    <div id="ch-cine-delta" style="height:420px"></div>
  </section>

  <section class="card">
    <div class="ct">Cuando cambian los creditos, que mas cambia?</div>
    <div class="ct-note">Barra ambar = entre los programas <strong>con</strong> cambio de creditos, que % tambien
      cambio ese campo. Barra gris = tasa base: que % cambia ese campo entre TODOS los programas modificados
      visibles con los filtros actuales (con o sin cambio de creditos) - es el punto de comparacion.</div>
    <div id="ch-cocambios" style="height:300px"></div>
  </section>

  <section class="card">
    <div class="ct">Creditos: antes -&gt; despues (coloreado por sector)</div>
    <div id="ch-scatter" style="height:380px"></div>
  </section>

  <section class="card">
    <div class="ct">Registros con cambio de creditos</div>
    <div class="tbl-wrap" id="tbl-wrap"></div>
  </section>
</main>

<script>
const D = __DATA__;
const PC = {responsive:true, displayModeBar:false};
const fmt = n => (n ?? 0).toLocaleString('es-CO');
const _norm = s => String(s==null?'':s).normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').toLowerCase();
function _rowMatches(r, tokens) {
  if (!tokens.length) return true;
  const hay = _norm(Object.values(r).join(' '));
  return tokens.every(t => hay.includes(t));
}

function _emptyChart(id, msg) {
  const el = document.getElementById(id); if (!el) return;
  Plotly.purge(el);
  el.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#64748b;font-size:.82rem">' + (msg || 'Sin datos para los filtros aplicados') + '</div>';
}

function agruparPorCampo(rows, field) {
  const m = new Map();
  rows.forEach(r => {
    const k = r[field]; if (!k) return;
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  });
  return m;
}

function resumenGrupo(rows) {
  const n = rows.length;
  const suben = rows.filter(r => r._delta > 0).length;
  const bajan = rows.filter(r => r._delta < 0).length;
  const promedio = n ? rows.reduce((s, r) => s + r._delta, 0) / n : 0;
  return {n, suben, bajan, promedio_delta: Math.round(promedio * 10) / 10};
}

function ranking(rows, field, topN) {
  const m = agruparPorCampo(rows, field);
  const out = [...m.entries()].map(([nombre, g]) => ({nombre, ...resumenGrupo(g)}));
  out.sort((a, b) => b.n - a.n);
  return out.slice(0, topN);
}

const BUCKETS_DELTA = [
  {min: -Infinity, max: -30,      label: '≤ -30'},
  {min: -29,       max: -10,      label: '-29 a -10'},
  {min: -9,        max: -1,       label: '-9 a -1'},
  {min: 0,         max: 0,        label: '0'},
  {min: 1,         max: 9,        label: '1 a 9'},
  {min: 10,        max: 29,       label: '10 a 29'},
  {min: 30,        max: Infinity, label: '≥ 30'},
];

function binDeltas(rows) {
  const counts = BUCKETS_DELTA.map(b => rows.filter(r => r._delta >= b.min && r._delta <= b.max).length);
  const colors = BUCKETS_DELTA.map(b => b.max < 0 ? '#ae1e22' : b.min > 0 ? '#1a9e6b' : '#94a3b8');
  return {labels: BUCKETS_DELTA.map(b => b.label), counts, colors};
}

const CAMPOS_COCAMBIO = [
  {flag: '_cambia_periodo',   label: 'Duracion (periodos)'},
  {flag: '_cambia_costo',     label: 'Costo de matricula'},
  {flag: '_cambia_modalidad', label: 'Modalidad'},
  {flag: '_cambia_municipio', label: 'Municipio'},
];

function calcularCoCambios(filtrados, creditRows) {
  return CAMPOS_COCAMBIO.map(({flag, label}) => {
    const conCredito = creditRows.filter(r => r[flag] !== null);
    const tasaConCredito = conCredito.length
      ? Math.round(1000 * conCredito.filter(r => r[flag] === true).length / conCredito.length) / 10
      : 0;
    const baseValidos = filtrados.filter(r => r[flag] !== null);
    const tasaBase = baseValidos.length
      ? Math.round(1000 * baseValidos.filter(r => r[flag] === true).length / baseValidos.length) / 10
      : 0;
    return {campo: label, n_con_dato: conCredito.length, tasa_con_cambio_credito: tasaConCredito, tasa_base: tasaBase};
  });
}

function renderSector(porSector) {
  if (!porSector.length) { _emptyChart('ch-sector'); return; }
  const labels = porSector.map(s => s.nombre);
  Plotly.react('ch-sector', [
    {x: labels, y: porSector.map(s => s.suben), name: 'Suben', type: 'bar',
     marker: {color: '#1a9e6b', opacity: .85}},
    {x: labels, y: porSector.map(s => s.bajan), name: 'Bajan', type: 'bar',
     marker: {color: '#ae1e22', opacity: .85}},
  ], {
    barmode: 'group', margin: {t:10,r:10,b:30,l:45},
    yaxis: {showgrid: true, gridcolor: '#e2e8f0'},
    plot_bgcolor: 'white', paper_bgcolor: 'white',
    legend: {orientation: 'h', y: -0.18}
  }, PC);
}

function renderHist(creditRows) {
  const h = binDeltas(creditRows);
  if (!h.counts.some(c => c > 0)) { _emptyChart('ch-hist'); return; }
  Plotly.react('ch-hist', [{
    x: h.labels, y: h.counts, type: 'bar',
    marker: {color: h.colors, opacity: .85},
    hovertemplate: 'Rango: %{x} creditos<br><b>%{y}</b> programas<extra></extra>'
  }], {
    margin: {t:10,r:10,b:45,l:45},
    xaxis: {title: 'Delta creditos (despues - antes)', tickfont: {size: 11}},
    yaxis: {title: 'N. programas', showgrid: true, gridcolor: '#e2e8f0'},
    plot_bgcolor: 'white', paper_bgcolor: 'white', bargap: .15
  }, PC);
}

function renderRanking(id, data) {
  if (!data.length) { _emptyChart(id); return; }
  const d = [...data].slice(0, 12).reverse();
  const trunc = s => s.length > 36 ? s.slice(0, 36) + '...' : s;
  const maxN = Math.max(...d.map(r => r.n));
  Plotly.react(id, [{
    y: d.map(r => trunc(r.nombre)), x: d.map(r => r.n), customdata: d.map(r => r.nombre),
    type: 'bar', orientation: 'h', marker: {color: '#bd900b', opacity: .85},
    text: d.map(r => `${r.promedio_delta > 0 ? '+' : ''}${r.promedio_delta}`),
    textposition: 'outside', cliponaxis: false,
    textfont: {size: 10, color: '#7a5d06'},
    hovertemplate: '%{customdata}<br><b>%{x}</b> cambios<extra></extra>'
  }], {
    margin: {t:10,r:60,b:30,l:230},
    xaxis: {showgrid: true, gridcolor: '#e2e8f0', range: [0, maxN * 1.35]},
    yaxis: {tickfont: {size: 10}},
    plot_bgcolor: 'white', paper_bgcolor: 'white', bargap: .3
  }, PC);
}

function renderCineDelta(creditRows) {
  const CINE_COL = 'CINE_F_2013_AC_CAMPO_ESPECÍFIC';
  const m = agruparPorCampo(creditRows, CINE_COL);
  const data = [...m.entries()].map(([campo, rows]) => ({
    campo,
    n: rows.length,
    suben: rows.filter(r => r._delta > 0).length,
    bajan: rows.filter(r => r._delta < 0).length,
    neto: rows.reduce((s, r) => s + r._delta, 0),
  }));
  data.sort((a, b) => Math.abs(b.neto) - Math.abs(a.neto));
  const top = data.slice(0, 20).reverse();
  if (!top.length) { _emptyChart('ch-cine-delta'); return; }
  const trunc = s => s.length > 42 ? s.slice(0, 42) + '...' : s;
  Plotly.react('ch-cine-delta', [{
    y: top.map(d => trunc(d.campo)), x: top.map(d => d.neto), customdata: top,
    type: 'bar', orientation: 'h',
    marker: {color: top.map(d => d.neto >= 0 ? '#1a9e6b' : '#ae1e22'), opacity: .85},
    text: top.map(d => (d.neto > 0 ? '+' : '') + d.neto),
    textposition: 'outside', cliponaxis: false, textfont: {size: 10},
    hovertemplate: '%{y}<br>Balance neto: <b>%{x}</b> creditos' +
      '<br>%{customdata.suben} suben / %{customdata.bajan} bajan (%{customdata.n} programas)<extra></extra>'
  }], {
    margin: {t:10,r:50,b:40,l:270},
    xaxis: {title: 'Balance neto de creditos (suma de deltas)', zeroline: true, zerolinecolor: '#94a3b8',
      showgrid: true, gridcolor: '#e2e8f0'},
    yaxis: {tickfont: {size: 10}},
    plot_bgcolor: 'white', paper_bgcolor: 'white', bargap: .3
  }, PC);
}

function renderCoCambios(coCambios) {
  if (!coCambios.length) { _emptyChart('ch-cocambios'); return; }
  const labels = coCambios.map(c => c.campo);
  Plotly.react('ch-cocambios', [
    {x: labels, y: coCambios.map(c => c.tasa_con_cambio_credito), name: 'Cuando cambian creditos',
     type: 'bar', marker: {color: '#bd900b', opacity: .9},
     text: coCambios.map(c => c.tasa_con_cambio_credito + '%'), textposition: 'outside', cliponaxis: false},
    {x: labels, y: coCambios.map(c => c.tasa_base), name: 'Tasa general (control)',
     type: 'bar', marker: {color: '#94a3b8', opacity: .9},
     text: coCambios.map(c => c.tasa_base + '%'), textposition: 'outside', cliponaxis: false},
  ], {
    barmode: 'group', margin: {t:30,r:10,b:50,l:50},
    yaxis: {title: '% de los casos', showgrid: true, gridcolor: '#e2e8f0', range: [0, 100]},
    plot_bgcolor: 'white', paper_bgcolor: 'white',
    legend: {orientation: 'h', y: -0.25}
  }, PC);
}

function renderScatter(creditRows) {
  if (!creditRows.length) { _emptyChart('ch-scatter'); return; }
  const porSector = {};
  creditRows.forEach(p => { (porSector[p.SECTOR] = porSector[p.SECTOR] || []).push(p); });
  const colores = {Oficial: '#2d5b9e', Privado: '#bd900b'};
  const vals = creditRows.flatMap(p => [p._cred_antes, p._cred_despues]);
  const mn = Math.min(...vals), mx = Math.max(...vals);
  const traces = Object.entries(porSector).map(([sector, pts]) => ({
    x: pts.map(p => p._cred_antes), y: pts.map(p => p._cred_despues),
    text: pts.map(p => p['NOMBRE_INSTITUCIÓN'] + ' - ' + p['NOMBRE_DEL_PROGRAMA']),
    mode: 'markers', type: 'scatter', name: sector,
    marker: {color: colores[sector] || '#64748b', size: 7, opacity: .65},
    hovertemplate: '<b>%{text}</b><br>Antes: %{x} creditos<br>Despues: %{y} creditos<extra></extra>'
  }));
  traces.push({x: [mn, mx], y: [mn, mx], mode: 'lines', line: {color: '#9fb0c9', width: 1, dash: 'dot'},
    hoverinfo: 'skip', showlegend: false});
  Plotly.react('ch-scatter', traces, {
    margin: {t:10,r:20,b:50,l:60},
    xaxis: {title: 'Creditos antes', showgrid: true, gridcolor: '#e2e8f0'},
    yaxis: {title: 'Creditos despues', showgrid: true, gridcolor: '#e2e8f0'},
    plot_bgcolor: 'white', paper_bgcolor: 'white',
    legend: {orientation: 'h', y: -0.2}
  }, PC);
}

const COL_HEAD = {
  FECHA_OBTENCION: 'Fecha', 'CODIGO_SNIES_DEL_PROGRAMA': 'Cod. SNIES', NOMBRE_DEL_PROGRAMA: 'Programa',
  'NOMBRE_INSTITUCIÓN': 'Institucion', SECTOR: 'Sector', DEPARTAMENTO_OFERTA_PROGRAMA: 'Departamento',
  'DIVISION UNINORTE': 'Division', _cred_antes: 'Cred. antes', _cred_despues: 'Cred. despues',
  _delta: 'Delta', QUE_CAMBIO: 'Que mas cambio?'
};
const TBL_COLS = ['FECHA_OBTENCION', 'NOMBRE_DEL_PROGRAMA', 'NOMBRE_INSTITUCIÓN', 'SECTOR',
  'DEPARTAMENTO_OFERTA_PROGRAMA', '_cred_antes', '_cred_despues', '_delta', 'QUE_CAMBIO'];
let sortDir = {};
let tablaRows = [];

function buildTbl(rows) {
  const cols = TBL_COLS.filter(c => !rows.length || c in rows[0]);
  let h = '<table><thead><tr>';
  cols.forEach(c => { h += '<th onclick="sortTbl(\\'' + c + '\\')">' + (COL_HEAD[c] || c) + ' <span style="opacity:.4">↕</span></th>'; });
  h += '</tr></thead><tbody>';
  if (!rows.length) {
    h += '<tr><td colspan="' + cols.length + '" class="empty">Sin registros para los filtros seleccionados</td></tr>';
  } else {
    rows.forEach(r => {
      h += '<tr>' + cols.map(c => {
        if (c === '_delta') {
          const v = r[c];
          const cls = v > 0 ? 'delta-up' : (v < 0 ? 'delta-down' : '');
          return '<td class="' + cls + '">' + (v > 0 ? '+' : '') + v + '</td>';
        }
        return '<td>' + (r[c] ?? '') + '</td>';
      }).join('') + '</tr>';
    });
  }
  return h + '</tbody></table>';
}

function renderTbl(rows) { tablaRows = rows; document.getElementById('tbl-wrap').innerHTML = buildTbl(rows); }

function sortTbl(col) {
  sortDir[col] = !sortDir[col];
  tablaRows = [...tablaRows].sort((a, b) => {
    const va = a[col] ?? '', vb = b[col] ?? '';
    const na = parseFloat(va), nb = parseFloat(vb);
    const cmp = (!isNaN(na) && !isNaN(nb)) ? na - nb : String(va).localeCompare(String(vb), 'es');
    return sortDir[col] ? cmp : -cmp;
  });
  document.getElementById('tbl-wrap').innerHTML = buildTbl(tablaRows);
}

function uniq(arr) { return [...new Set(arr.filter(v => v && String(v).trim() !== ''))].sort(); }
function addOpts(id, vals) {
  const el = document.getElementById(id); if (!el) return;
  vals.forEach(v => { const o = document.createElement('option'); o.value = o.textContent = v; el.appendChild(o); });
}
function initAutocomplete(inputId, menuId, options, onChange) {
  const inp = document.getElementById(inputId), menu = document.getElementById(menuId);
  if (!inp || !menu) return;
  function render() {
    const q = _norm(inp.value).trim();
    const matches = (q ? options.filter(o => _norm(o).includes(q)) : options).slice(0, 50);
    menu.innerHTML = matches.length
      ? matches.map(o => '<div class="ac-item">' + o.replace(/&/g,'&amp;').replace(/</g,'&lt;') + '</div>').join('')
      : '<div class="ac-empty">Sin coincidencias</div>';
    menu.classList.add('show');
  }
  inp.addEventListener('focus', render);
  inp.addEventListener('input', () => { render(); onChange(); });
  menu.addEventListener('mousedown', e => {
    const it = e.target.closest('.ac-item'); if (!it) return;
    e.preventDefault();
    inp.value = it.textContent;
    menu.classList.remove('show');
    onChange();
  });
  document.addEventListener('click', e => {
    if (e.target !== inp && !menu.contains(e.target)) menu.classList.remove('show');
  });
}
addOpts('f-sector', uniq(D.universo.map(r => r['SECTOR'])));
addOpts('f-depto',  uniq(D.universo.map(r => r['DEPARTAMENTO_OFERTA_PROGRAMA'])));
initAutocomplete('f-institucion', 'f-institucion-menu', uniq(D.universo.map(r => r['NOMBRE_INSTITUCIÓN'])), applyFilters);

function gv(id) { const el = document.getElementById(id); return el ? el.value : ''; }

function applyFilters() {
  const qTokens = _norm(gv('f-q')).split(/\\s+/).filter(Boolean);
  const se = gv('f-sector'), de = gv('f-depto');
  const insTokens = _norm(gv('f-institucion')).split(/\\s+/).filter(Boolean);

  const filtrados = D.universo.filter(r => {
    if (!_rowMatches(r, qTokens)) return false;
    if (se && r['SECTOR'] !== se) return false;
    if (de && r['DEPARTAMENTO_OFERTA_PROGRAMA'] !== de) return false;
    if (insTokens.length) {
      const hayIns = _norm(r['NOMBRE_INSTITUCIÓN']);
      if (!insTokens.every(t => hayIns.includes(t))) return false;
    }
    return true;
  });
  const creditRows = filtrados.filter(r => r._cambia_credito);

  document.getElementById('f-count').textContent = fmt(creditRows.length) + ' cambios de creditos (' + fmt(filtrados.length) + ' programas modificados en total)';
  document.getElementById('k-total').textContent = fmt(creditRows.length);

  const porSectorAgg = ranking(creditRows, 'SECTOR', 10).sort((a, b) => a.nombre.localeCompare(b.nombre));
  const sOficial = porSectorAgg.find(s => s.nombre === 'Oficial');
  const sPrivado = porSectorAgg.find(s => s.nombre === 'Privado');
  document.getElementById('k-oficial').textContent = sOficial ? (sOficial.promedio_delta > 0 ? '+' : '') + sOficial.promedio_delta : '-';
  document.getElementById('k-oficial-sub').textContent = sOficial ? `${sOficial.suben} suben / ${sOficial.bajan} bajan` : 'sin datos';
  document.getElementById('k-privado').textContent = sPrivado ? (sPrivado.promedio_delta > 0 ? '+' : '') + sPrivado.promedio_delta : '-';
  document.getElementById('k-privado-sub').textContent = sPrivado ? `${sPrivado.suben} suben / ${sPrivado.bajan} bajan` : 'sin datos';

  const coCambios = calcularCoCambios(filtrados, creditRows);
  const coDuracion = coCambios.find(c => c.campo.includes('Duracion'));
  document.getElementById('k-cocambio').textContent = coDuracion ? coDuracion.tasa_con_cambio_credito + '%' : '-';
  document.getElementById('k-cocambio-sub').textContent = coDuracion ? `vs. ${coDuracion.tasa_base}% tasa general` : '';

  renderSector(porSectorAgg);
  renderHist(creditRows);
  renderRanking('ch-instituciones', ranking(creditRows, 'NOMBRE_INSTITUCIÓN', 20));
  renderRanking('ch-departamentos', ranking(creditRows, 'DEPARTAMENTO_OFERTA_PROGRAMA', 20));
  renderCineDelta(creditRows);
  renderCoCambios(coCambios);
  renderScatter(creditRows);
  renderTbl(creditRows);
}

function resetFilters() {
  ['f-q', 'f-sector', 'f-depto', 'f-institucion'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  applyFilters();
}

applyFilters();
</script>
</body>
</html>
"""


# ── Página de análisis: costos de matrícula ───────────────────────────────────

CAMPO_COSTOS_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Analisis de Costos de Matricula . SNIES Monitor</title>
__PLOTLY_CDN__
<style>
:root{
  --bg:#f1f5f9;--surface:#fff;--text:#0f172a;--muted:#64748b;
  --border:#e2e8f0;--blue:#2d5b9e;--green:#1a9e6b;--red:#ae1e22;--amber:#bd900b;
  --radius:0.75rem;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);font-size:14px}
header{background:linear-gradient(135deg,#15284b,#1a9e6b);color:#fff;
  padding:1.2rem 2rem;display:flex;justify-content:space-between;align-items:center;gap:1rem;flex-wrap:wrap}
header h1{font-size:1.25rem;font-weight:700}
header .sub{font-size:.77rem;opacity:.75;margin-top:.2rem}
.back-btn{display:inline-flex;align-items:center;gap:.3rem;background:rgba(255,255,255,.2);
  border:1px solid rgba(255,255,255,.35);color:#fff;text-decoration:none;padding:.38rem .85rem;
  border-radius:.4rem;font-size:.8rem;font-weight:500;white-space:nowrap;transition:background .15s}
.back-btn:hover{background:rgba(255,255,255,.32)}
.filter-bar{position:sticky;top:0;z-index:100;background:var(--surface);
  border-bottom:1px solid var(--border);padding:.6rem 2rem;
  display:flex;gap:.45rem;flex-wrap:wrap;align-items:center;
  box-shadow:0 2px 8px rgba(0,0,0,.06)}
.f-input{flex:1;min-width:200px;padding:.5rem .8rem;border:1px solid var(--border);
  border-radius:.4rem;font-size:.8rem;outline:none}
.f-input:focus{border-color:var(--blue)}
.f-sel{padding:.4rem .6rem;border:1px solid var(--border);border-radius:.4rem;
  font-size:.77rem;background:var(--surface);outline:none;cursor:pointer;max-width:200px}
.f-btn{padding:.4rem .85rem;border:1px solid var(--border);border-radius:.4rem;
  font-size:.77rem;background:var(--surface);cursor:pointer;color:var(--muted);white-space:nowrap}
.f-btn:hover{background:var(--bg)}
.f-count{margin-left:auto;font-size:.82rem;font-weight:600;color:var(--blue);white-space:nowrap}
.ac-wrap{position:relative;flex:0 1 240px}
.ac-menu{position:absolute;top:calc(100% + 2px);left:0;right:0;z-index:300;
  background:var(--surface);border:1px solid var(--border);border-radius:.4rem;
  box-shadow:0 8px 20px rgba(0,0,0,.14);max-height:260px;overflow-y:auto;display:none}
.ac-menu.show{display:block}
.ac-item{padding:.45rem .75rem;font-size:.78rem;cursor:pointer;color:var(--text)}
.ac-item:hover{background:var(--bg)}
.ac-empty{padding:.45rem .75rem;font-size:.78rem;color:var(--muted)}
main{max-width:1380px;margin:0 auto;padding:1.5rem 2rem}
section{margin-bottom:1.25rem}
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:1rem}
.kpi{background:var(--surface);border-radius:var(--radius);padding:1.1rem 1.4rem;
  box-shadow:0 1px 3px rgba(0,0,0,.07);border-left:4px solid var(--blue)}
.kpi.r{border-left-color:var(--red)}.kpi.g{border-left-color:var(--green)}.kpi.a{border-left-color:var(--amber)}
.kpi-label{font-size:.68rem;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:.4rem}
.kpi-val{font-size:1.8rem;font-weight:700;line-height:1}
.kpi-sub{font-size:.7rem;color:var(--muted);margin-top:.35rem}
.card{background:var(--surface);border-radius:var(--radius);padding:1.2rem;
  box-shadow:0 1px 3px rgba(0,0,0,.07);margin-bottom:1rem}
.ct{font-size:.68rem;font-weight:600;text-transform:uppercase;letter-spacing:.06em;
  color:var(--muted);margin-bottom:.85rem}
.ct-note{font-size:.74rem;color:var(--muted);margin-top:-.5rem;margin-bottom:.75rem;line-height:1.4}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:1rem}
.tbl-wrap{max-height:480px;overflow-y:auto;border:1px solid var(--border);border-radius:.5rem}
table{width:100%;border-collapse:collapse;font-size:.77rem}
th{background:var(--bg);padding:.6rem .85rem;text-align:left;font-size:.67rem;
  text-transform:uppercase;letter-spacing:.05em;color:var(--muted);cursor:pointer;
  user-select:none;position:sticky;top:0;z-index:1;white-space:nowrap}
th:hover{background:#e2e8f0}
td{padding:.6rem .85rem;border-bottom:1px solid var(--border);vertical-align:top;
  max-width:270px;word-break:break-word}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8fafc}
.delta-up{color:var(--red);font-weight:600}
.delta-down{color:var(--green);font-weight:600}
.empty{text-align:center;color:var(--muted);padding:2.5rem}
@media(max-width:900px){
  .kpi-grid{grid-template-columns:repeat(2,1fr)}
  .g2{grid-template-columns:1fr}
  main{padding:1rem}
  header{flex-direction:column;text-align:center}
  .f-count{margin-left:0}
}
</style>
</head>
<body>
<header>
  <div style="display:flex;align-items:center;gap:.9rem">
    <a href="modificados.html" class="back-btn">← Modificados</a>
    <div>
      <h1>💰 Analisis de Costos de Matricula</h1>
      <div class="sub">Que cambia cuando un programa modifica el costo de matricula</div>
    </div>
  </div>
  <a href="modificados_creditos.html" class="back-btn">📐 Analisis de Creditos →</a>
</header>

<div class="filter-bar">
  <input id="f-q" class="f-input" placeholder="Buscar por nombre, institucion, departamento... (filtra TODA la pagina)" oninput="applyFilters()">
  <select id="f-sector" class="f-sel" onchange="applyFilters()"><option value="">Todos los sectores</option></select>
  <select id="f-depto"  class="f-sel" onchange="applyFilters()"><option value="">Todos los departamentos</option></select>
  <div class="ac-wrap">
    <input id="f-institucion" class="f-sel" placeholder="Buscar institucion..." style="cursor:text;width:100%" oninput="applyFilters()">
    <div class="ac-menu" id="f-institucion-menu"></div>
  </div>
  <button class="f-btn" onclick="resetFilters()">✕ Limpiar</button>
  <span class="f-count" id="f-count">–</span>
</div>

<main>
  <section class="kpi-grid">
    <div class="kpi">
      <div class="kpi-label">Cambios de costo detectados</div>
      <div class="kpi-val" id="k-total">–</div>
      <div class="kpi-sub">programas con costo antes/despues distinto</div>
    </div>
    <div class="kpi r">
      <div class="kpi-label">delta % promedio . Oficial</div>
      <div class="kpi-val" id="k-oficial">–</div>
      <div class="kpi-sub" id="k-oficial-sub">de costo por cambio</div>
    </div>
    <div class="kpi g">
      <div class="kpi-label">delta % promedio . Privado</div>
      <div class="kpi-val" id="k-privado">–</div>
      <div class="kpi-sub" id="k-privado-sub">de costo por cambio</div>
    </div>
    <div class="kpi a">
      <div class="kpi-label">Aumento maximo detectado</div>
      <div class="kpi-val" id="k-max">–</div>
      <div class="kpi-sub" id="k-max-sub">sin datos</div>
    </div>
  </section>
  <div class="ct-note" style="margin:-.5rem 0 1rem">
    "delta % promedio" = cambio promedio porcentual en el costo de matricula (costo despues vs. costo antes) entre
    los programas modificados de ese grupo. Se usa % y no pesos porque el costo varia en ordenes de magnitud
    distintos entre programas, asi que el porcentaje es lo unico comparable entre ellos.
    <strong>Negativo</strong> significa que, en promedio, baja la matricula; <strong>positivo</strong>, que sube.
  </div>

  <section class="g2">
    <div class="card">
      <div class="ct">Sube vs. baja, por sector</div>
      <div id="ch-sector" style="height:280px"></div>
    </div>
    <div class="card">
      <div class="ct">Distribucion del cambio (% de costo de matricula)</div>
      <div id="ch-hist" style="height:280px"></div>
    </div>
  </section>

  <section class="g2">
    <div class="card">
      <div class="ct">Top instituciones que mas cambian el costo</div>
      <div class="ct-note">Numero al lado de la barra: delta % promedio de esa institucion (ver definicion arriba).</div>
      <div id="ch-instituciones" style="height:380px"></div>
    </div>
    <div class="card">
      <div class="ct">Top departamentos que mas cambian el costo</div>
      <div class="ct-note">Numero al lado de la barra: delta % promedio de ese departamento (ver definicion arriba).</div>
      <div id="ch-departamentos" style="height:380px"></div>
    </div>
  </section>

  <section class="card">
    <div class="ct">Por campo CINE: que areas suben o bajan el costo de matricula</div>
    <div class="ct-note">Filtra por institucion arriba para ver el detalle de su oferta. Cada barra es un campo
      CINE; el valor es el cambio % promedio entre los programas de ese campo. Rojo = en promedio sube el costo;
      verde = en promedio baja.</div>
    <div id="ch-cine-delta" style="height:420px"></div>
  </section>

  <section class="card">
    <div class="ct">Costo de matricula: antes -&gt; despues (coloreado por sector, escala log)</div>
    <div id="ch-scatter" style="height:380px"></div>
  </section>

  <section class="card">
    <div class="ct">Registros con cambio de costo</div>
    <div class="tbl-wrap" id="tbl-wrap"></div>
  </section>
</main>

<script>
const D = __DATA__;
const PC = {responsive:true, displayModeBar:false};
const fmt = n => (n ?? 0).toLocaleString('es-CO');
const _norm = s => String(s==null?'':s).normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').toLowerCase();
function _rowMatches(r, tokens) {
  if (!tokens.length) return true;
  const hay = _norm(Object.values(r).join(' '));
  return tokens.every(t => hay.includes(t));
}

function _emptyChart(id, msg) {
  const el = document.getElementById(id); if (!el) return;
  Plotly.purge(el);
  el.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#64748b;font-size:.82rem">' + (msg || 'Sin datos para los filtros aplicados') + '</div>';
}

function agruparPorCampo(rows, field) {
  const m = new Map();
  rows.forEach(r => {
    const k = r[field]; if (!k) return;
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  });
  return m;
}

function resumenGrupo(rows) {
  const n = rows.length;
  const suben = rows.filter(r => r._delta_pct > 0).length;
  const bajan = rows.filter(r => r._delta_pct < 0).length;
  const promedio = n ? rows.reduce((s, r) => s + r._delta_pct, 0) / n : 0;
  return {n, suben, bajan, promedio_delta: Math.round(promedio * 10) / 10};
}

function ranking(rows, field, topN) {
  const m = agruparPorCampo(rows, field);
  const out = [...m.entries()].map(([nombre, g]) => ({nombre, ...resumenGrupo(g)}));
  out.sort((a, b) => b.n - a.n);
  return out.slice(0, topN);
}

const BUCKETS_DELTA = [
  {min: -Infinity, max: -20,      label: '≤ -20%'},
  {min: -19,       max: -10,      label: '-19 a -10%'},
  {min: -9,        max: -1,       label: '-9 a -1%'},
  {min: 0,         max: 0,        label: '0%'},
  {min: 1,         max: 9,        label: '1 a 9%'},
  {min: 10,        max: 19,       label: '10 a 19%'},
  {min: 20,        max: 49,       label: '20 a 49%'},
  {min: 50,        max: Infinity, label: '≥ 50%'},
];

function binDeltas(rows) {
  const counts = BUCKETS_DELTA.map(b => rows.filter(r => r._delta_pct >= b.min && r._delta_pct <= b.max).length);
  const colors = BUCKETS_DELTA.map(b => b.max < 0 ? '#1a9e6b' : b.min > 0 ? '#ae1e22' : '#94a3b8');
  return {labels: BUCKETS_DELTA.map(b => b.label), counts, colors};
}

function renderSector(porSector) {
  if (!porSector.length) { _emptyChart('ch-sector'); return; }
  const labels = porSector.map(s => s.nombre);
  Plotly.react('ch-sector', [
    {x: labels, y: porSector.map(s => s.suben), name: 'Suben', type: 'bar',
     marker: {color: '#ae1e22', opacity: .85}},
    {x: labels, y: porSector.map(s => s.bajan), name: 'Bajan', type: 'bar',
     marker: {color: '#1a9e6b', opacity: .85}},
  ], {
    barmode: 'group', margin: {t:10,r:10,b:30,l:45},
    yaxis: {showgrid: true, gridcolor: '#e2e8f0'},
    plot_bgcolor: 'white', paper_bgcolor: 'white',
    legend: {orientation: 'h', y: -0.18}
  }, PC);
}

function renderHist(costRows) {
  const h = binDeltas(costRows);
  if (!h.counts.some(c => c > 0)) { _emptyChart('ch-hist'); return; }
  Plotly.react('ch-hist', [{
    x: h.labels, y: h.counts, type: 'bar',
    marker: {color: h.colors, opacity: .85},
    hovertemplate: 'Rango: %{x}<br><b>%{y}</b> programas<extra></extra>'
  }], {
    margin: {t:10,r:10,b:75,l:45},
    xaxis: {title: {text: 'Cambio en costo de matricula (%)', standoff: 25}, tickangle: -30, tickfont: {size: 10}},
    yaxis: {title: 'N. programas', showgrid: true, gridcolor: '#e2e8f0'},
    plot_bgcolor: 'white', paper_bgcolor: 'white', bargap: .15
  }, PC);
}

function renderRanking(id, data) {
  if (!data.length) { _emptyChart(id); return; }
  const d = [...data].slice(0, 12).reverse();
  const trunc = s => s.length > 36 ? s.slice(0, 36) + '...' : s;
  const maxN = Math.max(...d.map(r => r.n));
  Plotly.react(id, [{
    y: d.map(r => trunc(r.nombre)), x: d.map(r => r.n), customdata: d.map(r => r.nombre),
    type: 'bar', orientation: 'h',
    marker: {color: d.map(r => r.promedio_delta >= 0 ? '#ae1e22' : '#1a9e6b'), opacity: .85},
    text: d.map(r => `${r.promedio_delta > 0 ? '+' : ''}${r.promedio_delta}%`),
    textposition: 'outside', cliponaxis: false,
    textfont: {size: 10, color: d.map(r => r.promedio_delta >= 0 ? '#7a1518' : '#0f6e49')},
    hovertemplate: '%{customdata}<br><b>%{x}</b> cambios<extra></extra>'
  }], {
    margin: {t:10,r:60,b:30,l:230},
    xaxis: {showgrid: true, gridcolor: '#e2e8f0', range: [0, maxN * 1.35]},
    yaxis: {tickfont: {size: 10}},
    plot_bgcolor: 'white', paper_bgcolor: 'white', bargap: .3
  }, PC);
}

function renderCineDelta(costRows) {
  const CINE_COL = 'CINE_F_2013_AC_CAMPO_ESPECÍFIC';
  const m = agruparPorCampo(costRows, CINE_COL);
  const data = [...m.entries()].map(([campo, rows]) => {
    const n = rows.length;
    const suben = rows.filter(r => r._delta_pct > 0).length;
    const bajan = rows.filter(r => r._delta_pct < 0).length;
    const promedio = n ? rows.reduce((s, r) => s + r._delta_pct, 0) / n : 0;
    return {campo, n, suben, bajan, promedio: Math.round(promedio * 10) / 10};
  });
  data.sort((a, b) => Math.abs(b.promedio) - Math.abs(a.promedio));
  const top = data.slice(0, 20).reverse();
  if (!top.length) { _emptyChart('ch-cine-delta'); return; }
  const trunc = s => s.length > 42 ? s.slice(0, 42) + '...' : s;
  Plotly.react('ch-cine-delta', [{
    y: top.map(d => trunc(d.campo)), x: top.map(d => d.promedio), customdata: top,
    type: 'bar', orientation: 'h',
    marker: {color: top.map(d => d.promedio >= 0 ? '#ae1e22' : '#1a9e6b'), opacity: .85},
    text: top.map(d => (d.promedio > 0 ? '+' : '') + d.promedio + '%'),
    textposition: 'outside', cliponaxis: false, textfont: {size: 10},
    hovertemplate: '%{y}<br>Cambio promedio: <b>%{x}%</b>' +
      '<br>%{customdata.suben} suben / %{customdata.bajan} bajan (%{customdata.n} programas)<extra></extra>'
  }], {
    margin: {t:10,r:50,b:40,l:270},
    xaxis: {title: 'Cambio promedio en costo de matricula (%)', zeroline: true, zerolinecolor: '#94a3b8',
      showgrid: true, gridcolor: '#e2e8f0'},
    yaxis: {tickfont: {size: 10}},
    plot_bgcolor: 'white', paper_bgcolor: 'white', bargap: .3
  }, PC);
}

function renderScatter(costRows) {
  if (!costRows.length) { _emptyChart('ch-scatter'); return; }
  const porSector = {};
  costRows.forEach(p => { (porSector[p.SECTOR] = porSector[p.SECTOR] || []).push(p); });
  const colores = {Oficial: '#2d5b9e', Privado: '#bd900b'};
  const vals = costRows.flatMap(p => [p._costo_antes, p._costo_despues]);
  const mn = Math.min(...vals), mx = Math.max(...vals);
  const traces = Object.entries(porSector).map(([sector, pts]) => ({
    x: pts.map(p => p._costo_antes), y: pts.map(p => p._costo_despues),
    text: pts.map(p => p['NOMBRE_INSTITUCIÓN'] + ' - ' + p['NOMBRE_DEL_PROGRAMA']),
    mode: 'markers', type: 'scatter', name: sector,
    marker: {color: colores[sector] || '#64748b', size: 7, opacity: .65},
    hovertemplate: '<b>%{text}</b><br>Antes: $%{x:,.0f}<br>Despues: $%{y:,.0f}<extra></extra>'
  }));
  traces.push({x: [mn, mx], y: [mn, mx], mode: 'lines', line: {color: '#9fb0c9', width: 1, dash: 'dot'},
    hoverinfo: 'skip', showlegend: false});
  Plotly.react('ch-scatter', traces, {
    margin: {t:10,r:20,b:50,l:70},
    xaxis: {title: 'Costo antes (COP)', type: 'log', showgrid: true, gridcolor: '#e2e8f0'},
    yaxis: {title: 'Costo despues (COP)', type: 'log', showgrid: true, gridcolor: '#e2e8f0'},
    plot_bgcolor: 'white', paper_bgcolor: 'white',
    legend: {orientation: 'h', y: -0.2}
  }, PC);
}

const COL_HEAD = {
  FECHA_OBTENCION: 'Fecha', 'CODIGO_SNIES_DEL_PROGRAMA': 'Cod. SNIES', NOMBRE_DEL_PROGRAMA: 'Programa',
  'NOMBRE_INSTITUCIÓN': 'Institucion', SECTOR: 'Sector', DEPARTAMENTO_OFERTA_PROGRAMA: 'Departamento',
  'DIVISION UNINORTE': 'Division', _costo_antes: 'Costo antes', _costo_despues: 'Costo despues',
  _delta_pct: 'Delta %', QUE_CAMBIO: 'Que mas cambio?'
};
const TBL_COLS = ['FECHA_OBTENCION', 'NOMBRE_DEL_PROGRAMA', 'NOMBRE_INSTITUCIÓN', 'SECTOR',
  'DEPARTAMENTO_OFERTA_PROGRAMA', '_costo_antes', '_costo_despues', '_delta_pct', 'QUE_CAMBIO'];
let sortDir = {};
let tablaRows = [];

function buildTbl(rows) {
  const cols = TBL_COLS.filter(c => !rows.length || c in rows[0]);
  let h = '<table><thead><tr>';
  cols.forEach(c => { h += '<th onclick="sortTbl(\\'' + c + '\\')">' + (COL_HEAD[c] || c) + ' <span style="opacity:.4">↕</span></th>'; });
  h += '</tr></thead><tbody>';
  if (!rows.length) {
    h += '<tr><td colspan="' + cols.length + '" class="empty">Sin registros para los filtros seleccionados</td></tr>';
  } else {
    rows.forEach(r => {
      h += '<tr>' + cols.map(c => {
        if (c === '_delta_pct') {
          const v = r[c];
          const cls = v > 0 ? 'delta-up' : (v < 0 ? 'delta-down' : '');
          return '<td class="' + cls + '">' + (v > 0 ? '+' : '') + v + '%</td>';
        }
        if (c === '_costo_antes' || c === '_costo_despues') {
          return '<td>$' + fmt(r[c]) + '</td>';
        }
        return '<td>' + (r[c] ?? '') + '</td>';
      }).join('') + '</tr>';
    });
  }
  return h + '</tbody></table>';
}

function renderTbl(rows) { tablaRows = rows; document.getElementById('tbl-wrap').innerHTML = buildTbl(rows); }

function sortTbl(col) {
  sortDir[col] = !sortDir[col];
  tablaRows = [...tablaRows].sort((a, b) => {
    const va = a[col] ?? '', vb = b[col] ?? '';
    const na = parseFloat(va), nb = parseFloat(vb);
    const cmp = (!isNaN(na) && !isNaN(nb)) ? na - nb : String(va).localeCompare(String(vb), 'es');
    return sortDir[col] ? cmp : -cmp;
  });
  document.getElementById('tbl-wrap').innerHTML = buildTbl(tablaRows);
}

function uniq(arr) { return [...new Set(arr.filter(v => v && String(v).trim() !== ''))].sort(); }
function addOpts(id, vals) {
  const el = document.getElementById(id); if (!el) return;
  vals.forEach(v => { const o = document.createElement('option'); o.value = o.textContent = v; el.appendChild(o); });
}
function initAutocomplete(inputId, menuId, options, onChange) {
  const inp = document.getElementById(inputId), menu = document.getElementById(menuId);
  if (!inp || !menu) return;
  function render() {
    const q = _norm(inp.value).trim();
    const matches = (q ? options.filter(o => _norm(o).includes(q)) : options).slice(0, 50);
    menu.innerHTML = matches.length
      ? matches.map(o => '<div class="ac-item">' + o.replace(/&/g,'&amp;').replace(/</g,'&lt;') + '</div>').join('')
      : '<div class="ac-empty">Sin coincidencias</div>';
    menu.classList.add('show');
  }
  inp.addEventListener('focus', render);
  inp.addEventListener('input', () => { render(); onChange(); });
  menu.addEventListener('mousedown', e => {
    const it = e.target.closest('.ac-item'); if (!it) return;
    e.preventDefault();
    inp.value = it.textContent;
    menu.classList.remove('show');
    onChange();
  });
  document.addEventListener('click', e => {
    if (e.target !== inp && !menu.contains(e.target)) menu.classList.remove('show');
  });
}
addOpts('f-sector', uniq(D.universo.map(r => r['SECTOR'])));
addOpts('f-depto',  uniq(D.universo.map(r => r['DEPARTAMENTO_OFERTA_PROGRAMA'])));
initAutocomplete('f-institucion', 'f-institucion-menu', uniq(D.universo.map(r => r['NOMBRE_INSTITUCIÓN'])), applyFilters);

function gv(id) { const el = document.getElementById(id); return el ? el.value : ''; }
const truncKpi = s => s.length > 55 ? s.slice(0, 55) + '…' : s;

function applyFilters() {
  const qTokens = _norm(gv('f-q')).split(/\\s+/).filter(Boolean);
  const se = gv('f-sector'), de = gv('f-depto');
  const insTokens = _norm(gv('f-institucion')).split(/\\s+/).filter(Boolean);

  const filtrados = D.universo.filter(r => {
    if (!_rowMatches(r, qTokens)) return false;
    if (se && r['SECTOR'] !== se) return false;
    if (de && r['DEPARTAMENTO_OFERTA_PROGRAMA'] !== de) return false;
    if (insTokens.length) {
      const hayIns = _norm(r['NOMBRE_INSTITUCIÓN']);
      if (!insTokens.every(t => hayIns.includes(t))) return false;
    }
    return true;
  });
  const costRows = filtrados.filter(r => r._cambia_costo);

  document.getElementById('f-count').textContent = fmt(costRows.length) + ' cambios de costo (' + fmt(filtrados.length) + ' programas modificados en total)';
  document.getElementById('k-total').textContent = fmt(costRows.length);

  const porSectorAgg = ranking(costRows, 'SECTOR', 10).sort((a, b) => a.nombre.localeCompare(b.nombre));
  const sOficial = porSectorAgg.find(s => s.nombre === 'Oficial');
  const sPrivado = porSectorAgg.find(s => s.nombre === 'Privado');
  document.getElementById('k-oficial').textContent = sOficial ? (sOficial.promedio_delta > 0 ? '+' : '') + sOficial.promedio_delta + '%' : '-';
  document.getElementById('k-oficial-sub').textContent = sOficial ? `${sOficial.suben} suben / ${sOficial.bajan} bajan` : 'sin datos';
  document.getElementById('k-privado').textContent = sPrivado ? (sPrivado.promedio_delta > 0 ? '+' : '') + sPrivado.promedio_delta + '%' : '-';
  document.getElementById('k-privado-sub').textContent = sPrivado ? `${sPrivado.suben} suben / ${sPrivado.bajan} bajan` : 'sin datos';

  const maxRow = costRows.reduce((best, r) => (!best || r._delta_pct > best._delta_pct) ? r : best, null);
  document.getElementById('k-max').textContent = maxRow ? (maxRow._delta_pct > 0 ? '+' : '') + maxRow._delta_pct + '%' : '-';
  document.getElementById('k-max-sub').textContent = maxRow ? truncKpi(maxRow['NOMBRE_INSTITUCIÓN'] + ' — ' + maxRow['NOMBRE_DEL_PROGRAMA']) : 'sin datos';

  renderSector(porSectorAgg);
  renderHist(costRows);
  renderRanking('ch-instituciones', ranking(costRows, 'NOMBRE_INSTITUCIÓN', 20));
  renderRanking('ch-departamentos', ranking(costRows, 'DEPARTAMENTO_OFERTA_PROGRAMA', 20));
  renderCineDelta(costRows);
  renderScatter(costRows);
  renderTbl(costRows);
}

function resetFilters() {
  ['f-q', 'f-sector', 'f-depto', 'f-institucion'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  applyFilters();
}

applyFilters();
</script>
</body>
</html>
"""


def _json_default(o):
    if pd.isna(o):
        return None
    return str(o)


def _clean_for_json(records):
    cleaned = []
    for r in records:
        rec = {}
        for k, v in r.items():
            if v is None:
                rec[k] = None
            elif isinstance(v, bool):
                rec[k] = v
            elif isinstance(v, float):
                rec[k] = None if pd.isna(v) else v
            else:
                rec[k] = v
        cleaned.append(rec)
    return cleaned


def _dumps(obj):
    return json.dumps(obj, ensure_ascii=False, default=_json_default).replace("</", "<\\/")


# ── Ensamblado de páginas ──────────────────────────────────────────────────────

def build_index(historico, nuevos_df, inactivos_df, mods_df, snapshot_df):
    kpis = {
        "total_activos":    historico[-1]["total"] if historico else 0,
        "nuevos_ultimo":    _count_last_run(nuevos_df),
        "inactivos_ultimo": _count_last_run(inactivos_df),
        "mods_ultimo":      _count_last_run(mods_df),
        "nuevos_total":     len(nuevos_df),
        "inactivos_total":  len(inactivos_df),
        "mods_total":       len(mods_df),
    }

    data = {
        "ultima_actualizacion": historico[-1]["fecha"] if historico else "N/A",
        "historico":     [{"fecha": h["fecha"], "total": h["total"]} for h in historico],
        "kpis":          kpis,
        "historico_modalidad": historico_por_modalidad(historico),
        "historico_cruzado": historico_cruzado(historico),
        "snapshot":      _to_records(snapshot_df, SNAPSHOT_COLS),
        "nuevos":        _to_records(nuevos_df,    COLS_IDX_PREVIEW),
        "inactivos":     _to_records(inactivos_df, COLS_IDX_PREVIEW),
        "modificados":   _to_records(mods_df,      COLS_IDX_PREVIEW_MOD),
        "n_nuevos":      len(nuevos_df),
        "n_inactivos":   len(inactivos_df),
        "n_modificados": len(mods_df),
    }

    html = INDEX_TEMPLATE.replace("__PLOTLY_CDN__", PLOTLY_CDN)
    html = html.replace("__DATA__", _dumps(data))
    return html


DETAIL_CFGS = {
    "nuevos": {"title": "Programas Nuevos", "emoji": "✅", "color": "#fcc10e",
               "colorAlpha": "rgba(252,193,14,0.08)", "hdrgrad": "linear-gradient(135deg,#15284b,#fcc10e)"},
    "inactivos": {"title": "Programas Inactivos", "emoji": "❌", "color": "#ae1e22",
                  "colorAlpha": "rgba(174,30,34,0.08)", "hdrgrad": "linear-gradient(135deg,#7a1518,#ae1e22)"},
    "modificados": {"title": "Programas Modificados", "emoji": "⚠️", "color": "#bd900b",
                    "colorAlpha": "rgba(189,144,11,0.08)", "hdrgrad": "linear-gradient(135deg,#15284b,#bd900b)"},
}


def build_detail_page(tipo, df, today):
    cfg_visual = DETAIL_CFGS[tipo]
    cols = COLS_MOD_DETAIL if tipo == "modificados" else COLS_DETAIL
    records = _to_records(df, cols)

    cfg = {
        "tipo": tipo,
        "title": cfg_visual["title"],
        "emoji": cfg_visual["emoji"],
        "color": cfg_visual["color"],
        "colorAlpha": cfg_visual["colorAlpha"],
        "cols": [c for c in cols if c in df.columns],
    }

    if tipo == "modificados":
        charts_section = CHARTS_MODIFICADOS
        backlinks = (
            '<a href="modificados_creditos.html" class="back-btn">📐 Análisis de Créditos →</a>'
            '<a href="modificados_costos.html" class="back-btn">💰 Análisis de Costos →</a>'
        )
    else:
        charts_section = CHARTS_NUEVOS_INACTIVOS
        backlinks = ""

    html = DETAIL_TEMPLATE.replace("__PLOTLY_CDN__", PLOTLY_CDN)
    html = html.replace("__TITLE__", f"{cfg_visual['title']} · SNIES Monitor")
    html = html.replace("__HDRGRAD__", cfg_visual["hdrgrad"])
    html = html.replace("__EMOJI__", cfg_visual["emoji"])
    html = html.replace("__H1TEXT__", cfg_visual["title"])
    html = html.replace("__BACKLINKS__", backlinks)
    html = html.replace(
        "__FILTRO_EXTRA__",
        '<select id="f-tipo-cambio" class="f-sel" onchange="applyFilters()"><option value="">Todos los cambios</option></select>'
        if tipo == "modificados" else
        '<select id="f-modalidad" class="f-sel" onchange="applyFilters()"><option value="">Todas las modalidades</option></select>\n'
        '  <select id="f-cine" class="f-sel" onchange="applyFilters()"><option value="">Todos los campos CINE</option></select>'
    )
    html = html.replace("__CHARTS_SECTION__", charts_section)
    html = html.replace("__DATA__", _dumps(records))
    html = html.replace("__CFG__", _dumps(cfg))
    return html


def build_creditos(df_m_enr, today):
    records = _universo_records(df_m_enr, [
        "_cred_antes", "_cred_despues", "_delta", "_cambia_credito",
        "_cambia_periodo", "_cambia_costo", "_cambia_modalidad", "_cambia_municipio",
    ])
    data = {"universo": _clean_for_json(records)}
    html = CAMPO_CREDITOS_TEMPLATE.replace("__PLOTLY_CDN__", PLOTLY_CDN)
    html = html.replace("__DATA__", _dumps(data))
    return html


def build_costos(df_m_enr, today):
    records = _universo_records(df_m_enr, [
        "_costo_antes", "_costo_despues", "_delta", "_cambia_costo",
        "_cambia_credito", "_cambia_periodo", "_cambia_modalidad", "_cambia_municipio",
    ])
    # delta_pct sólo tiene sentido si hay costo_antes > 0
    for r in records:
        ca, cd = r.get("_costo_antes"), r.get("_costo_despues")
        if ca and cd is not None and ca > 0:
            r["_delta_pct"] = round(100 * (cd - ca) / ca, 1)
        else:
            r["_delta_pct"] = None
    data = {"universo": _clean_for_json(records)}
    html = CAMPO_COSTOS_TEMPLATE.replace("__PLOTLY_CDN__", PLOTLY_CDN)
    html = html.replace("__DATA__", _dumps(data))
    return html


# ── main ───────────────────────────────────────────────────────────────────────

def main():
    today = datetime.today().strftime("%Y-%m-%d")
    print(f"[{today}] Leyendo novedades...")

    historico = leer_historico()
    nuevos_df    = leer_novedades("Nuevos_posgrado.xlsx")
    inactivos_df = leer_novedades("Inactivos_posgrado.xlsx")
    mods_df      = leer_novedades("Modificados_posgrado.xlsx")
    snapshot_df  = leer_snapshot_actual(historico)

    print(f"  Nuevos: {len(nuevos_df):,}  Inactivos: {len(inactivos_df):,}  Modificados: {len(mods_df):,}")
    print(f"  Snapshots historicos: {len(historico)}")

    mods_enr = enriquecer_modificados(mods_df) if not mods_df.empty else mods_df

    DOCS.mkdir(exist_ok=True)

    print("Generando index.html...")
    (DOCS / "index.html").write_text(
        build_index(historico, nuevos_df, inactivos_df, mods_df, snapshot_df),
        encoding="utf-8",
    )

    print("Generando nuevos.html...")
    (DOCS / "nuevos.html").write_text(build_detail_page("nuevos", nuevos_df, today), encoding="utf-8")

    print("Generando inactivos.html...")
    (DOCS / "inactivos.html").write_text(build_detail_page("inactivos", inactivos_df, today), encoding="utf-8")

    print("Generando modificados.html...")
    (DOCS / "modificados.html").write_text(build_detail_page("modificados", mods_df, today), encoding="utf-8")

    print("Generando modificados_creditos.html...")
    (DOCS / "modificados_creditos.html").write_text(build_creditos(mods_enr, today), encoding="utf-8")

    print("Generando modificados_costos.html...")
    (DOCS / "modificados_costos.html").write_text(build_costos(mods_enr, today), encoding="utf-8")

    print("OK Dashboard generado en docs/")
    paginas = ["index.html", "nuevos.html", "inactivos.html", "modificados.html",
               "modificados_creditos.html", "modificados_costos.html"]
    for f in paginas:
        size = (DOCS / f).stat().st_size
        print(f"  {f}: {size/1024:.0f} KB")


if __name__ == "__main__":
    main()
