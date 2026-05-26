"""
Energy Data Analysis Dashboard — Dash / Plotly
Fonti: OWID (Excel URL) · Energy Institute (upload .xlsx) · World Bank WDI (CSV locale)
"""
import io, base64, hashlib, warnings, requests
from pathlib import Path
import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, Input, Output, State, dash_table, no_update
import dash_bootstrap_components as dbc

warnings.filterwarnings("ignore")

# ── Configurazione fonti ─────────────────────────────────────────────────────────
DATA_SOURCES = {
    "OWID – Global Energy Data": {
        "type": "owid",
        "url": "https://nyc3.digitaloceanspaces.com/owid-public/data/energy/owid-energy-data.xlsx",
    },
    "Energy Institute – Stats Review": {
        "type": "ei",
        "url": "https://www.energyinst.org/__data/assets/excel_doc/0008/1656215/EI-Stats-Review-ALL-data.xlsx",
    },
    "World Bank – WDI (Energia & Ambiente)": {
        "type": "wdi",
        "local_data": str(Path(__file__).parent / "worldbank_wdi_bulk" / "WDICSV.csv"),
        "local_series": str(Path(__file__).parent / "worldbank_wdi_bulk" / "WDISeries.csv"),
    },
}
SOURCE_NAMES = list(DATA_SOURCES.keys())
CHART_TYPES  = ["Line", "Bar", "Scatter", "Area"]
WDI_PREFIXES = ("EG.", "EN.ATM", "EN.CO2", "EN.GHG", "CC.")

# ── Cache server-side (memoria) + disco ─────────────────────────────────────────
_CACHE: dict[str, pd.DataFrame] = {}
_DISK_CACHE = Path(__file__).parent / ".dashboard_cache"
_DISK_CACHE.mkdir(exist_ok=True)


def _disk_load(key: str) -> pd.DataFrame | None:
    p = _DISK_CACHE / f"{key}.pkl"
    if p.exists():
        df = pd.read_pickle(p)
        _CACHE[key] = df
        return df
    return None


def _disk_save(key: str, df: pd.DataFrame) -> None:
    df.to_pickle(_DISK_CACHE / f"{key}.pkl")


# ── Loaders ──────────────────────────────────────────────────────────────────────
def _fetch(url: str, verify: bool = True) -> bytes:
    h = {"User-Agent": "Mozilla/5.0 (compatible; EnergyDashboard/1.0)"}
    r = requests.get(url, timeout=120, verify=verify, headers=h)
    r.raise_for_status()
    return r.content


def load_owid(url: str) -> pd.DataFrame:
    if "owid" in _CACHE:
        return _CACHE["owid"]
    if (cached := _disk_load("owid")) is not None:
        return cached
    df = pd.read_excel(io.BytesIO(_fetch(url)), engine="openpyxl")
    df.columns = df.columns.str.strip()
    _CACHE["owid"] = df
    _disk_save("owid", df)
    return df


def ei_sheets(file_bytes: bytes) -> list[str]:
    return pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl").sheet_names


def load_ei(file_bytes: bytes, sheet: str) -> pd.DataFrame:
    sig = hashlib.md5(file_bytes[:512]).hexdigest()[:8]
    key = f"ei|{sheet}|{sig}"
    if key in _CACHE:
        return _CACHE[key]

    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet,
                        header=None, engine="openpyxl")
    # Auto-rileva riga header cercando la prima riga con ≥5 valori "anno"
    header_row = 0
    for i, row in raw.iterrows():
        if sum(1 for v in row if str(v).strip().isdigit()
               and 1900 < int(str(v).strip()) < 2100) >= 5:
            header_row = i
            break

    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet,
                       header=header_row, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    year_cols = [c for c in df.columns if c.isdigit() and 1900 < int(c) < 2100]
    if not year_cols:
        raise ValueError(f"Nessuna colonna anno nel foglio '{sheet}'")

    country_col = df.columns[0]
    df = df[[country_col] + year_cols].dropna(subset=[country_col])
    df = df[~df[country_col].astype(str).str.startswith(("Total", "of which", "Memo"))]

    df_long = df.melt(id_vars=[country_col], value_vars=year_cols,
                      var_name="year", value_name="value")
    df_long = df_long.rename(columns={country_col: "country"})
    df_long["year"]  = df_long["year"].astype(int)
    df_long["value"] = pd.to_numeric(df_long["value"], errors="coerce")
    df_long = df_long.dropna(subset=["value"])
    df_long["metric"] = sheet
    _CACHE[key] = df_long
    return df_long


def wdi_series_options() -> list[dict]:
    path = DATA_SOURCES["World Bank – WDI (Energia & Ambiente)"]["local_series"]
    s = pd.read_csv(path)
    s = s[s["Series Code"].str.startswith(WDI_PREFIXES, na=False)]
    s = s[["Series Code", "Indicator Name"]].dropna().sort_values("Indicator Name")
    return [{"label": r["Indicator Name"], "value": r["Series Code"]} for _, r in s.iterrows()]


def load_wdi(codes: tuple) -> pd.DataFrame:
    disk_key = "wdi_" + hashlib.md5("|".join(sorted(codes)).encode()).hexdigest()[:8]
    mem_key  = "wdi|" + disk_key[4:]
    if mem_key in _CACHE:
        return _CACHE[mem_key]
    if (cached := _disk_load(disk_key)) is not None:
        _CACHE[mem_key] = cached
        return cached

    path = DATA_SOURCES["World Bank – WDI (Energia & Ambiente)"]["local_data"]
    chunks = [c[c["Indicator Code"].isin(codes)]
              for c in pd.read_csv(path, chunksize=5000, low_memory=False)]
    chunks = [c for c in chunks if not c.empty]
    if not chunks:
        return pd.DataFrame()

    df = pd.concat(chunks, ignore_index=True)
    yr_cols = [c for c in df.columns if c.isdigit()]
    df_long = df.melt(
        id_vars=["Country Name", "Country Code", "Indicator Name", "Indicator Code"],
        value_vars=yr_cols, var_name="year", value_name="value",
    )
    df_long["year"]  = df_long["year"].astype(int)
    df_long["value"] = pd.to_numeric(df_long["value"], errors="coerce")
    df_long = df_long.dropna(subset=["value"])
    _CACHE[mem_key] = df_long
    _disk_save(disk_key, df_long)
    return df_long


# ── Helpers ──────────────────────────────────────────────────────────────────────
def _apply_filters(df: pd.DataFrame, meta: dict,
                   countries: list, years: list) -> pd.DataFrame:
    out = df.copy()
    cc, yc = meta["country_col"], meta["year_col"]
    if countries:
        out = out[out[cc].isin(countries)]
    if years and len(years) == 2:
        out = out[out[yc].between(years[0], years[1])]
    return out


def _make_fig(sub: pd.DataFrame, x: str, y: str, color: str,
              chart_type: str, title: str = "") -> object:
    kw = dict(data_frame=sub, x=x, y=y, color=color, title=title)
    return {
        "Line":    lambda: px.line(**kw),
        "Bar":     lambda: px.bar(**kw, barmode="group"),
        "Scatter": lambda: px.scatter(**kw),
        "Area":    lambda: px.area(**kw),
    }.get(chart_type, lambda: px.line(**kw))()


# WDI options loaded once at startup
try:
    _WDI_OPTS = wdi_series_options()
except Exception:
    _WDI_OPTS = []


# ── Layout ───────────────────────────────────────────────────────────────────────
_UPLOAD_STYLE = {
    "borderWidth": "1px", "borderStyle": "dashed", "borderRadius": "5px",
    "textAlign": "center", "padding": "10px", "cursor": "pointer",
    "color": "#6c757d", "fontSize": "0.85rem",
}

sidebar = dbc.Col(
    className="bg-light border-end p-3",
    style={"minHeight": "100vh", "overflowY": "auto"},
    width=3,
    children=[
        html.H5("Fonte dati", className="fw-bold"),
        dcc.Dropdown(
            id="source-select",
            options=[{"label": k, "value": k} for k in SOURCE_NAMES],
            value=SOURCE_NAMES[0],
            clearable=False,
            className="mb-2",
        ),

        html.Div(id="ei-options", style={"display": "none"}, children=[
            dcc.Upload(
                id="ei-upload",
                children=html.Div("Trascina o clicca per caricare il file .xlsx"),
                style=_UPLOAD_STYLE,
                multiple=False,
            ),
            dcc.Dropdown(id="ei-sheet", placeholder="Seleziona foglio…",
                         className="mt-2", clearable=False),
        ]),

        html.Div(id="wdi-options", style={"display": "none"}, children=[
            html.Label("Indicatori energetici:", className="small"),
            dcc.Dropdown(
                id="wdi-indicators",
                options=_WDI_OPTS,
                value=[o["value"] for o in _WDI_OPTS[:5]],
                multi=True,
                placeholder="Seleziona indicatori…",
            ),
        ]),

        dbc.Button("Carica / Aggiorna dati", id="btn-load",
                   color="primary", size="sm", className="w-100 my-3"),

        html.Hr(),
        html.H5("Filtri globali", className="fw-bold"),
        html.Label("Paesi / Regioni", className="small"),
        dcc.Dropdown(id="filter-countries", multi=True,
                     placeholder="Seleziona paesi…", className="mb-3"),
        html.Label("Intervallo anni", className="small"),
        dcc.RangeSlider(
            id="filter-years", step=1, marks=None,
            min=1960, max=2026, value=[2000, 2026],
            tooltip={"placement": "bottom", "always_visible": True},
        ),

        # ── Pannello Grafico ───────────────────────────────────────────────
        html.Div(id="panel-chart", style={"display": "none"}, children=[
            html.Hr(),
            html.H5("Grafico", className="fw-bold"),
            html.Label("Metrica", className="small"),
            dcc.Dropdown(id="chart-metric", className="mb-2", clearable=False),
            html.Div(id="chart-indicator-wrap", style={"display": "none"}, children=[
                html.Label("Indicatore", className="small"),
                dcc.Dropdown(id="chart-indicator", className="mb-2", clearable=False),
            ]),
            html.Label("Tipo grafico", className="small"),
            dcc.Dropdown(
                id="chart-type",
                options=[{"label": t, "value": t} for t in CHART_TYPES],
                value="Line", clearable=False,
            ),
        ]),

        # ── Pannello Confronto ─────────────────────────────────────────────
        html.Div(id="panel-compare", style={"display": "none"}, children=[
            html.Hr(),
            html.H5("Confronto", className="fw-bold"),

            html.Label("Modalità", className="small fw-semibold"),
            dcc.RadioItems(
                id="cmp-mode",
                options=[
                    {"label": " Serie temporale", "value": "temporal"},
                    {"label": " Snapshot anno",   "value": "snapshot"},
                    {"label": " Scatter (X vs Y)", "value": "scatter"},
                    {"label": " Radar multi-metrica", "value": "radar"},
                ],
                value="temporal",
                className="mb-2",
                inputStyle={"marginRight": "5px"},
                labelStyle={"display": "block", "marginBottom": "3px",
                            "fontSize": "0.85rem"},
            ),

            html.Label("Paesi da confrontare", className="small"),
            dcc.Dropdown(id="cmp-countries", multi=True,
                         placeholder="Tutti i filtrati…", className="mb-2"),

            html.Label("Metrica principale", className="small"),
            dcc.Dropdown(id="cmp-metric-a", className="mb-2", clearable=False),

            # Seconda metrica (solo Scatter)
            html.Div(id="cmp-metric-b-wrap", style={"display": "none"}, children=[
                html.Label("Metrica Y (Scatter)", className="small"),
                dcc.Dropdown(id="cmp-metric-b", className="mb-2", clearable=False),
            ]),

            # Metriche extra (Snapshot e Radar)
            html.Div(id="cmp-metrics-extra-wrap", style={"display": "none"}, children=[
                html.Label("Metriche aggiuntive", className="small"),
                dcc.Dropdown(id="cmp-metrics-extra", multi=True,
                             placeholder="Aggiungi metriche…", className="mb-2"),
            ]),

            # Anno (Snapshot, Scatter, Radar)
            html.Div(id="cmp-year-wrap", style={"display": "none"}, children=[
                html.Label("Anno di riferimento", className="small"),
                dcc.Dropdown(id="cmp-year", placeholder="Seleziona anno…",
                             className="mb-2", clearable=False),
            ]),
        ]),
    ],
)

main_area = dbc.Col(
    className="p-3",
    width=9,
    children=[
        dcc.Loading(type="circle", color="#2c3e50", children=html.Div(id="alert-area")),
        html.Div(id="summary-banner", className="mb-3"),
        dcc.Tabs(
            id="main-tabs", value="tab-table",
            children=[
                dcc.Tab(label="📋 Tabella dati",  value="tab-table"),
                dcc.Tab(label="📈 Grafico",        value="tab-chart"),
                dcc.Tab(label="⚖️  Confronto",     value="tab-compare"),
                dcc.Tab(label="📊 Statistiche",    value="tab-stats"),
            ],
        ),
        dcc.Loading(
            type="circle",
            color="#2c3e50",
            style={"marginTop": "60px"},
            children=html.Div(id="tab-content", className="mt-3"),
        ),
        dcc.Download(id="download-csv"),
        dcc.Store(id="store-key"),
        dcc.Store(id="store-meta"),
    ],
)

app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.FLATLY],
    suppress_callback_exceptions=True,
)
server = app.server  # esposto per gunicorn

app.layout = dbc.Container(fluid=True, children=[
    dbc.Row(dbc.Col(
        html.H3("⚡ Energy Data Analysis Platform",
                className="py-3 mb-0 border-bottom"),
    )),
    dbc.Row([sidebar, main_area]),
])


# ── Callbacks ────────────────────────────────────────────────────────────────────

# 1. Mostra/nascondi opzioni specifiche per fonte
@app.callback(
    Output("ei-options",  "style"),
    Output("wdi-options", "style"),
    Input("source-select", "value"),
)
def toggle_source_panels(source):
    stype = DATA_SOURCES[source]["type"]
    show, hide = {"display": "block"}, {"display": "none"}
    return (show if stype == "ei"  else hide,
            show if stype == "wdi" else hide)


# 2. Aggiorna lista fogli EI dopo upload
@app.callback(
    Output("ei-sheet", "options"),
    Output("ei-sheet", "value"),
    Input("ei-upload", "contents"),
    prevent_initial_call=True,
)
def update_ei_sheets(contents):
    if not contents:
        return [], None
    _, b64 = contents.split(",", 1)
    sheets = ei_sheets(base64.b64decode(b64))
    opts = [{"label": s, "value": s} for s in sheets]
    return opts, sheets[0] if sheets else None


# 3. Caricamento dati
@app.callback(
    Output("store-key",  "data"),
    Output("store-meta", "data"),
    Output("alert-area", "children"),
    Input("btn-load", "n_clicks"),
    State("source-select",   "value"),
    State("ei-upload",       "contents"),
    State("ei-sheet",        "value"),
    State("wdi-indicators",  "value"),
    prevent_initial_call=True,
)
def load_data(_, source, ei_contents, ei_sheet, wdi_codes):
    cfg   = DATA_SOURCES[source]
    stype = cfg["type"]
    try:
        # ── Calcola la chiave prima di qualsiasi I/O ───────────────────────────
        if stype == "owid":
            key  = "owid"
            meta = {"type": "owid", "format": "wide",
                    "country_col": "country", "year_col": "year",
                    "indicator_col": None}

        elif stype == "ei":
            if not ei_contents:
                return no_update, no_update, dbc.Alert(
                    "Carica prima il file .xlsx di EI Stats Review.", color="warning")
            if not ei_sheet:
                return no_update, no_update, dbc.Alert(
                    "Seleziona un foglio dalla lista.", color="warning")
            _, b64 = ei_contents.split(",", 1)
            fb  = base64.b64decode(b64)
            sig = hashlib.md5(fb[:512]).hexdigest()[:8]
            key  = f"ei|{ei_sheet}|{sig}"
            meta = {"type": "ei", "format": "long",
                    "country_col": "country", "year_col": "year",
                    "indicator_col": None, "sheet": ei_sheet}

        elif stype == "wdi":
            if not wdi_codes:
                return no_update, no_update, dbc.Alert(
                    "Seleziona almeno un indicatore WDI.", color="warning")
            key  = "wdi|" + hashlib.md5("|".join(sorted(wdi_codes)).encode()).hexdigest()[:8]
            meta = {"type": "wdi", "format": "long",
                    "country_col": "Country Name", "year_col": "year",
                    "indicator_col": "Indicator Name"}
        else:
            return no_update, no_update, no_update

        # ── Se già in memoria non fare nulla ───────────────────────────────────
        if key in _CACHE:
            df = _CACHE[key]
            alert = dbc.Alert(f"✅  Dati già in memoria: {len(df):,} righe.",
                              color="info", duration=3000, dismissable=True)
            return key, meta, alert

        # ── Scarica / legge solo se necessario ────────────────────────────────
        if stype == "owid":
            df = load_owid(cfg["url"])
        elif stype == "ei":
            df = load_ei(fb, ei_sheet)
        elif stype == "wdi":
            df = load_wdi(tuple(wdi_codes))

        alert = dbc.Alert(f"✅  Dati caricati: {len(df):,} righe.",
                          color="success", duration=4000, dismissable=True)
        return key, meta, alert

    except Exception as exc:
        return no_update, no_update, dbc.Alert(f"❌  Errore: {exc}", color="danger", dismissable=True)


# 4. Aggiorna controlli dopo caricamento
@app.callback(
    Output("filter-countries",     "options"),
    Output("filter-countries",     "value"),
    Output("filter-years",         "min"),
    Output("filter-years",         "max"),
    Output("filter-years",         "value"),
    Output("chart-metric",         "options"),
    Output("chart-metric",         "value"),
    Output("chart-indicator",      "options"),
    Output("chart-indicator",      "value"),
    Output("chart-indicator-wrap", "style"),
    Output("cmp-countries",        "options"),
    Output("cmp-metric-a",         "options"),
    Output("cmp-metric-a",         "value"),
    Output("cmp-metric-b",         "options"),
    Output("cmp-metric-b",         "value"),
    Output("cmp-metrics-extra",    "options"),
    Output("cmp-year",             "options"),
    Output("cmp-year",             "value"),
    Output("summary-banner",       "children"),
    Input("store-key",  "data"),
    State("store-meta", "data"),
)
def update_controls(key, meta):
    none14 = ([], [], 1960, 2026, [2000, 2026],
              [], None, [], None, {"display": "none"},
              [], [], None, [], None, [], [], None, html.Div())
    if not key or not meta or key not in _CACHE:
        return none14

    df  = _CACHE[key]
    cc  = meta["country_col"]
    yc  = meta["year_col"]
    fmt = meta["format"]
    ic  = meta.get("indicator_col")

    countries    = sorted(df[cc].dropna().unique())
    country_opts = [{"label": c, "value": c} for c in countries]

    y_min, y_max = int(df[yc].min()), int(df[yc].max())
    yr_val       = [max(y_min, y_max - 25), y_max]
    year_opts    = [{"label": str(y), "value": y}
                    for y in sorted(df[yc].dropna().unique(), reverse=True)]

    if fmt == "wide":
        num       = [c for c in df.select_dtypes(include="number").columns if c != yc]
        m_opts    = [{"label": c, "value": c} for c in num]
        m_val     = num[0] if num else None
        ind_opts  = []
        ind_val   = None
        ind_style = {"display": "none"}
    else:
        m_opts = [{"label": "value", "value": "value"}]
        m_val  = "value"
        if ic and ic in df.columns:
            inds     = sorted(df[ic].dropna().unique())
            ind_opts = [{"label": i, "value": i} for i in inds]
            ind_val  = inds[0] if inds else None
            ind_style = {"display": "block"}
            m_opts   = ind_opts
            m_val    = inds[0] if inds else None
        else:
            ind_opts, ind_val, ind_style = [], None, {"display": "none"}

    n_metrics = len(m_opts)
    banner = dbc.Row([
        dbc.Col(dbc.Card(dbc.CardBody([
            html.H6("Righe totali", className="card-subtitle text-muted"),
            html.H4(f"{len(df):,}", className="card-title"),
        ])), width=3),
        dbc.Col(dbc.Card(dbc.CardBody([
            html.H6("Paesi / Regioni", className="card-subtitle text-muted"),
            html.H4(str(len(countries)), className="card-title"),
        ])), width=3),
        dbc.Col(dbc.Card(dbc.CardBody([
            html.H6("Anni disponibili", className="card-subtitle text-muted"),
            html.H4(f"{y_min} – {y_max}", className="card-title"),
        ])), width=3),
        dbc.Col(dbc.Card(dbc.CardBody([
            html.H6("Metriche / Indicatori", className="card-subtitle text-muted"),
            html.H4(str(n_metrics), className="card-title"),
        ])), width=3),
    ], className="g-2")

    cmp_m_val = m_opts[0]["value"] if m_opts else None
    cmp_m_b   = m_opts[1]["value"] if len(m_opts) > 1 else cmp_m_val

    return (country_opts, [], y_min, y_max, yr_val,
            m_opts, m_val, ind_opts, ind_val, ind_style,
            country_opts, m_opts, cmp_m_val, m_opts, cmp_m_b,
            m_opts, year_opts, y_max, banner)


# 4b. Mostra/nascondi pannelli sidebar in base al tab attivo
@app.callback(
    Output("panel-chart",   "style"),
    Output("panel-compare", "style"),
    Input("main-tabs", "value"),
)
def toggle_sidebar_panels(tab):
    show, hide = {"display": "block"}, {"display": "none"}
    return (show if tab == "tab-chart"   else hide,
            show if tab == "tab-compare" else hide)


# 4c. Mostra/nascondi controlli confronto in base alla modalità
@app.callback(
    Output("cmp-metric-b-wrap",      "style"),
    Output("cmp-metrics-extra-wrap", "style"),
    Output("cmp-year-wrap",          "style"),
    Input("cmp-mode", "value"),
)
def toggle_compare_controls(mode):
    show, hide = {"display": "block"}, {"display": "none"}
    return (
        show if mode == "scatter"                       else hide,  # metrica Y
        show if mode in ("snapshot", "radar")           else hide,  # metriche extra
        show if mode in ("snapshot", "scatter", "radar") else hide, # anno
    )


# 5. Contenuto tab
@app.callback(
    Output("tab-content", "children"),
    Input("main-tabs",          "value"),
    Input("filter-countries",   "value"),
    Input("filter-years",       "value"),
    Input("chart-metric",       "value"),
    Input("chart-indicator",    "value"),
    Input("chart-type",         "value"),
    Input("cmp-mode",           "value"),
    Input("cmp-countries",      "value"),
    Input("cmp-metric-a",       "value"),
    Input("cmp-metric-b",       "value"),
    Input("cmp-metrics-extra",  "value"),
    Input("cmp-year",           "value"),
    State("store-key",  "data"),
    State("store-meta", "data"),
)
def render_tab(tab, countries, years, metric, indicator, chart_type,
               cmp_mode, cmp_countries, cmp_metric_a, cmp_metric_b,
               cmp_metrics_extra, cmp_year, key, meta):
    if not key or not meta or key not in _CACHE:
        return dbc.Alert("Carica i dati dalla barra laterale per iniziare.",
                         color="info")

    df       = _CACHE[key]
    filtered = _apply_filters(df, meta, countries, years)

    if filtered.empty:
        return dbc.Alert("Nessun dato con i filtri selezionati.", color="warning")

    cc  = meta["country_col"]
    yc  = meta["year_col"]
    fmt = meta["format"]
    ic  = meta.get("indicator_col")

    # ── Tabella ────────────────────────────────────────────────────────────────
    if tab == "tab-table":
        cols = [{"name": c, "id": c} for c in filtered.columns]
        data = filtered.to_dict("records")
        return html.Div([
            html.P(f"{len(filtered):,} righe totali",
                   className="text-muted small mb-1"),
            dash_table.DataTable(
                columns=cols,
                data=data,
                page_size=50,
                page_action="native",
                sort_action="native",
                filter_action="native",
                style_table={"overflowX": "auto"},
                style_cell={"fontSize": "0.82rem", "padding": "4px 8px"},
                style_header={"fontWeight": "bold", "backgroundColor": "#f8f9fa"},
            ),
            dbc.Button("⬇️  Scarica CSV", id="btn-download", color="secondary",
                       size="sm", className="mt-2"),
        ])

    # ── Grafico ────────────────────────────────────────────────────────────────
    if tab == "tab-chart":
        if fmt == "wide":
            if not metric:
                return dbc.Alert("Seleziona una metrica dalla sidebar.", color="info")
            sub = filtered[[cc, yc, metric]].dropna()
            fig = _make_fig(sub, yc, metric, cc, chart_type or "Line", metric)
        else:
            target_ind = indicator if ic else None
            if target_ind and ic in filtered.columns:
                sub = filtered[filtered[ic] == target_ind]
            else:
                sub = filtered
            if sub.empty:
                return dbc.Alert("Nessun dato per l'indicatore selezionato.", color="warning")
            title = target_ind or (meta.get("sheet") or "value")
            fig = _make_fig(sub, yc, "value", cc, chart_type or "Line", title)

        fig.update_layout(hovermode="x unified", legend_title_text="Paese",
                          margin={"t": 40})
        return dcc.Graph(figure=fig, style={"height": "520px"})

    # ── Confronto ──────────────────────────────────────────────────────────────
    if tab == "tab-compare":
        if not cmp_metric_a:
            return dbc.Alert("Seleziona almeno la metrica principale nella sidebar.",
                             color="info")

        # Paesi da usare: selezione specifica o tutti i filtrati
        cmp_cc = cmp_countries if cmp_countries else filtered[cc].dropna().unique().tolist()
        sub_base = filtered[filtered[cc].isin(cmp_cc)]
        if sub_base.empty:
            return dbc.Alert("Nessun dato per i paesi selezionati.", color="warning")

        def _get_series(df_in, metric):
            """Estrae (country, year, value) per una metrica, sia wide che long."""
            if fmt == "wide":
                return df_in[[cc, yc, metric]].dropna().rename(columns={metric: "value"})
            else:
                col = ic if ic and ic in df_in.columns else None
                if col:
                    return df_in[df_in[col] == metric][[cc, yc, "value"]].dropna()
                return df_in[[cc, yc, "value"]].dropna()

        def _snap(df_in, metric, year):
            """Snapshot anno: country vs value."""
            s = _get_series(df_in, metric)
            s = s[s[yc] == year][[cc, "value"]].dropna()
            return s.sort_values("value", ascending=False).head(40)

        graphs = []

        # ── Temporale ──────────────────────────────────────────────────────
        if cmp_mode == "temporal":
            metrics_plot = [cmp_metric_a] + (cmp_metrics_extra or [])
            for m in metrics_plot:
                s = _get_series(sub_base, m)
                if s.empty:
                    continue
                fig = px.line(data_frame=s, x=yc, y="value", color=cc,
                              title=str(m),
                              labels={yc: "Anno", "value": str(m), cc: "Paese"})
                fig.update_layout(hovermode="x unified", margin={"t": 40})
                graphs.append(dcc.Graph(figure=fig, style={"height": "420px"}))

        # ── Snapshot anno ──────────────────────────────────────────────────
        elif cmp_mode == "snapshot":
            year = cmp_year or int(sub_base[yc].max())
            metrics_plot = [cmp_metric_a] + (cmp_metrics_extra or [])

            if len(metrics_plot) == 1:
                s = _snap(sub_base, metrics_plot[0], year)
                if not s.empty:
                    fig = px.bar(data_frame=s, x=cc, y="value", color=cc,
                                 title=f"{metrics_plot[0]}  –  {year}",
                                 labels={"value": metrics_plot[0], cc: "Paese"})
                    fig.update_layout(showlegend=False, margin={"t": 40})
                    graphs.append(dcc.Graph(figure=fig, style={"height": "420px"}))
            else:
                # Costruisce DataFrame largo per bar grouped
                frames = []
                for m in metrics_plot:
                    s = _snap(sub_base, m, year)
                    s = s.assign(metrica=str(m))
                    frames.append(s)
                if frames:
                    combined = pd.concat(frames, ignore_index=True)
                    fig = px.bar(data_frame=combined, x=cc, y="value",
                                 color="metrica", barmode="group",
                                 title=f"Confronto metriche  –  {year}",
                                 labels={"value": "Valore", cc: "Paese",
                                         "metrica": "Metrica"})
                    fig.update_layout(margin={"t": 40})
                    graphs.append(dcc.Graph(figure=fig, style={"height": "460px"}))

        # ── Scatter X vs Y ─────────────────────────────────────────────────
        elif cmp_mode == "scatter":
            if not cmp_metric_b:
                return dbc.Alert("Seleziona la metrica Y per lo scatter.", color="warning")
            year = cmp_year or int(sub_base[yc].max())
            sx = _snap(sub_base, cmp_metric_a, year).rename(columns={"value": "x"})
            sy = _snap(sub_base, cmp_metric_b, year).rename(columns={"value": "y"})
            merged = sx.merge(sy, on=cc)
            if merged.empty:
                return dbc.Alert("Nessun dato in comune tra le due metriche.", color="warning")
            fig = px.scatter(data_frame=merged, x="x", y="y", color=cc, text=cc,
                             title=f"{cmp_metric_a}  vs  {cmp_metric_b}  –  {year}",
                             labels={"x": str(cmp_metric_a), "y": str(cmp_metric_b),
                                     cc: "Paese"})
            fig.update_traces(textposition="top center", marker_size=10)
            fig.update_layout(margin={"t": 40})
            graphs.append(dcc.Graph(figure=fig, style={"height": "500px"}))

        # ── Radar multi-metrica ────────────────────────────────────────────
        elif cmp_mode == "radar":
            import plotly.graph_objects as go
            metrics_plot = [cmp_metric_a] + (cmp_metrics_extra or [])
            if len(metrics_plot) < 3:
                return dbc.Alert("Seleziona almeno 3 metriche per il radar.", color="warning")
            year = cmp_year or int(sub_base[yc].max())

            radar_data: dict[str, list] = {}
            for m in metrics_plot:
                s = _snap(sub_base, m, year).set_index(cc)["value"]
                for country, val in s.items():
                    radar_data.setdefault(country, {})[m] = val

            # Normalizza 0-1 per metrica
            all_vals = {m: [radar_data[c].get(m, 0) for c in radar_data] for m in metrics_plot}
            norm = {}
            for m, vals in all_vals.items():
                mn, mx = min(vals), max(vals)
                rng = mx - mn or 1
                norm[m] = {c: (radar_data[c].get(m, 0) - mn) / rng for c in radar_data}

            fig = go.Figure()
            for country in list(radar_data.keys())[:8]:
                r_vals = [norm[m].get(country, 0) for m in metrics_plot]
                r_vals += [r_vals[0]]
                theta   = [str(m)[:30] for m in metrics_plot] + [str(metrics_plot[0])[:30]]
                fig.add_trace(go.Scatterpolar(r=r_vals, theta=theta,
                                              fill="toself", name=country))
            fig.update_layout(
                polar={"radialaxis": {"visible": True, "range": [0, 1]}},
                title=f"Radar  –  {year}  (valori normalizzati 0–1)",
                margin={"t": 60},
                height=520,
            )
            graphs.append(dcc.Graph(figure=fig))

        return html.Div(graphs) if graphs else dbc.Alert(
            "Nessun dato disponibile con i parametri selezionati.", color="warning")

    # ── Statistiche ────────────────────────────────────────────────────────────
    if tab == "tab-stats":
        if fmt == "wide":
            num_cols = [c for c in filtered.select_dtypes(include="number").columns if c != yc]
            if not num_cols:
                return dbc.Alert("Nessuna colonna numerica.", color="warning")

            desc = filtered[num_cols].describe().T.reset_index()
            desc.columns = ["Metrica"] + list(desc.columns[1:])

            corr_cols = num_cols[:20]
            corr      = filtered[corr_cols].corr()
            fig_corr  = px.imshow(
                corr, text_auto=".2f", color_continuous_scale="RdBu_r",
                zmin=-1, zmax=1, title="Correlazione tra metriche",
                height=500,
            )

            return html.Div([
                html.H5("Statistiche descrittive"),
                dash_table.DataTable(
                    columns=[{"name": c, "id": c} for c in desc.columns],
                    data=desc.round(4).to_dict("records"),
                    style_cell={"fontSize": "0.82rem", "padding": "4px 8px"},
                    style_header={"fontWeight": "bold", "backgroundColor": "#f8f9fa"},
                    style_table={"overflowX": "auto"},
                ),
                html.H5("Mappa di correlazione", className="mt-4"),
                dcc.Graph(figure=fig_corr),
            ])
        else:
            group = [ic] if ic and ic in filtered.columns else []
            if group:
                desc = (filtered.groupby(group)["value"]
                        .describe().reset_index().round(4))
            else:
                desc = filtered[["value"]].describe().T.reset_index().round(4)

            fig_box = px.box(
                filtered, x=ic if (ic and ic in filtered.columns) else None,
                y="value", color=cc,
                title="Distribuzione valori per indicatore",
                height=500,
            )
            return html.Div([
                html.H5("Statistiche descrittive"),
                dash_table.DataTable(
                    columns=[{"name": c, "id": c} for c in desc.columns],
                    data=desc.to_dict("records"),
                    style_cell={"fontSize": "0.82rem", "padding": "4px 8px"},
                    style_header={"fontWeight": "bold", "backgroundColor": "#f8f9fa"},
                    style_table={"overflowX": "auto"},
                ),
                html.H5("Box plot", className="mt-4"),
                dcc.Graph(figure=fig_box),
            ])

    return html.Div()


# 6. Download CSV
@app.callback(
    Output("download-csv", "data"),
    Input("btn-download", "n_clicks"),
    State("filter-countries", "value"),
    State("filter-years",     "value"),
    State("store-key",  "data"),
    State("store-meta", "data"),
    prevent_initial_call=True,
)
def download_csv(_, countries, years, key, meta):
    if not key or key not in _CACHE or not meta:
        return no_update
    df       = _CACHE[key]
    filtered = _apply_filters(df, meta, countries, years)
    return dcc.send_data_frame(filtered.to_csv, "energy_data.csv", index=False)


if __name__ == "__main__":
    app.run(debug=True, port=8050)
