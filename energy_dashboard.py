"""
Energy Data Analysis Dashboard — Dash / Plotly
Fonti: OWID (Excel URL) · Energy Institute (upload .xlsx) · World Bank WDI (CSV locale)
"""
import io, base64, hashlib, warnings, requests, threading
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
        "url": "https://raw.githubusercontent.com/owid/energy-data/master/owid-energy-data.csv",
    },
    "Energy Institute – Stats Review": {
        "type": "ei",
        "url": "https://www.energyinst.org/__data/assets/excel_doc/0008/1656215/EI-Stats-Review-ALL-data.xlsx",
        "local_file": str(Path(__file__).parent / "EI-Stats-Review-ALL-data.xlsx"),
    },
    "World Bank – WDI (Energia & Ambiente)": {
        "type": "wdi",
        "api_base": "https://api.worldbank.org/v2/country/all/indicator",
        "local_series": str(Path(__file__).parent / "worldbank_wdi_bulk" / "WDISeries.csv"),
    },
}
SOURCE_NAMES = list(DATA_SOURCES.keys())
CHART_TYPES  = ["Line", "Bar", "Scatter", "Area"]
WDI_PREFIXES = ("EG.", "EN.ATM", "EN.CO2", "EN.GHG", "CC.")

# ── Etichette leggibili per le colonne OWID ──────────────────────────────────────
OWID_LABELS: dict[str, str] = {
    # Energia primaria
    "primary_energy_consumption":   "Consumo energia primaria (TWh)",
    "energy_per_capita":            "Consumo energia pro capite (kWh/persona)",
    "energy_per_gdp":               "Intensità energetica (kWh/$)",
    # Elettricità
    "electricity_generation":       "Produzione elettricità totale (TWh)",
    "per_capita_electricity":       "Elettricità pro capite (kWh/persona)",
    "net_elec_imports":             "Import netto elettricità (TWh)",
    # Rinnovabili – produzione
    "renewables_electricity":       "Elettricità rinnovabile (TWh)",
    "solar_electricity":            "Solare (TWh)",
    "wind_electricity":             "Eolico (TWh)",
    "hydro_electricity":            "Idroelettrico (TWh)",
    "biofuel_electricity":          "Biomasse (TWh)",
    "other_renewables_electricity": "Altre rinnovabili (TWh)",
    "low_carbon_electricity":       "Elettricità low-carbon (TWh)",
    # Quote percentuali
    "renewables_share_elec":        "Rinnovabili su elettricità (%)",
    "low_carbon_share_elec":        "Low-carbon su elettricità (%)",
    "fossil_share_elec":            "Fossili su elettricità (%)",
    "fossil_share_energy":          "Fossili su energia primaria (%)",
    "renewables_share_energy":      "Rinnovabili su energia primaria (%)",
    "nuclear_share_elec":           "Nucleare su elettricità (%)",
    # Nucleare
    "nuclear_electricity":          "Nucleare (TWh)",
    # Fossili – consumo
    "coal_cons_twh":                "Consumo carbone (TWh)",
    "oil_cons_twh":                 "Consumo petrolio (TWh)",
    "gas_cons_twh":                 "Consumo gas naturale (TWh)",
    "fossil_electricity":           "Elettricità da fossili (TWh)",
    "coal_electricity":             "Elettricità da carbone (TWh)",
    "oil_electricity":              "Elettricità da petrolio (TWh)",
    "gas_electricity":              "Elettricità da gas (TWh)",
    # Fossili – produzione
    "coal_prod_twh":                "Produzione carbone (TWh)",
    "oil_prod_twh":                 "Produzione petrolio (TWh)",
    "gas_prod_twh":                 "Produzione gas naturale (TWh)",
    # Emissioni
    "greenhouse_gas_emissions":     "Emissioni GHG (Mt CO₂eq)",
    "co2_emissions":                "Emissioni CO₂ totali (Mt)",
    "co2_per_capita":               "CO₂ pro capite (t/persona)",
    "co2_per_gdp":                  "Intensità carbonica PIL (kg/$)",
    "carbon_intensity_elec":        "Intensità carbonica elettricità (gCO₂/kWh)",
    "methane_emissions":            "Emissioni metano (Mt CO₂eq)",
    "nitrous_oxide_emissions":      "Emissioni N₂O (Mt CO₂eq)",
}

# Ordine di priorità: queste metriche vengono mostrate per prime e usate come default
OWID_PRIORITY = [
    "primary_energy_consumption",
    "renewables_share_elec",
    "fossil_share_energy",
    "co2_per_capita",
    "energy_per_capita",
    "electricity_generation",
    "renewables_electricity",
    "solar_electricity",
    "wind_electricity",
    "hydro_electricity",
    "low_carbon_share_elec",
    "carbon_intensity_elec",
    "greenhouse_gas_emissions",
    "coal_cons_twh",
    "oil_cons_twh",
    "gas_cons_twh",
    "nuclear_electricity",
    "co2_emissions",
    "per_capita_electricity",
    "fossil_electricity",
    "fossil_share_elec",
    "renewables_share_energy",
]

# Indicatori WDI consigliati come selezione iniziale
WDI_DEFAULT_CODES = [
    "EG.USE.PCAP.KG.OE",   # Energy use per capita (kg oil equivalent)
    "EG.ELC.ACCS.ZS",      # Access to electricity (% population)
    "EG.FEC.RNEW.ZS",      # Renewable energy consumption (% total)
    "EN.ATM.CO2E.PC",      # CO2 emissions per capita (metric tons)
    "EG.USE.COMM.FO.ZS",   # Fossil fuel energy consumption (%)
]

# ── Cache server-side (memoria) + disco ─────────────────────────────────────────
_CACHE: dict[str, pd.DataFrame] = {}
_DISK_CACHE = Path(__file__).parent / ".dashboard_cache"
_DISK_CACHE.mkdir(exist_ok=True)
_cancel_event = threading.Event()


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
    df = pd.read_csv(io.BytesIO(_fetch(url)))
    df.columns = df.columns.str.strip()
    _CACHE["owid"] = df
    _disk_save("owid", df)
    return df


def ei_sheets(file_bytes: bytes) -> list[str]:
    return pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl").sheet_names


_EI_SHEET_CACHE: dict[str, list[str]] = {}  # sig → lista fogli validi
_EI_UNITS:       dict[str, str]       = {}  # cache-key → unità di misura


def ei_valid_sheets(file_bytes: bytes) -> list[str]:
    """Restituisce solo i fogli che hanno colonne anno (con cache per file).
    Apre il file una volta sola per velocità."""
    sig = hashlib.md5(file_bytes[:512]).hexdigest()[:8]
    if sig in _EI_SHEET_CACHE:
        return _EI_SHEET_CACHE[sig]
    xf = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    valid = []
    for sheet in xf.sheet_names:
        try:
            raw = xf.parse(sheet, header=None, nrows=10)
            if any(sum(1 for v in row if _is_year(v)) >= 3
                   for _, row in raw.iterrows()):
                valid.append(sheet)
        except Exception:
            pass
    _EI_SHEET_CACHE[sig] = valid
    return valid


def _is_year(v) -> bool:
    """True se v è un numero intero tra 1900 e 2100."""
    try:
        return 1900 < int(float(str(v))) < 2100 and float(str(v)) % 1 == 0
    except (ValueError, TypeError):
        return False


def load_ei(file_bytes: bytes, sheet: str) -> pd.DataFrame:
    sig = hashlib.md5(file_bytes[:512]).hexdigest()[:8]
    key = f"ei|{sheet}|{sig}"
    if key in _CACHE:
        return _CACHE[key]

    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet,
                        header=None, engine="openpyxl")

    # Trova la riga header: prima riga con ≥5 valori che sembrano anni
    header_row = 0
    for i, row in raw.iterrows():
        if sum(1 for v in row if _is_year(v)) >= 5:
            header_row = i
            break

    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet,
                       header=header_row, engine="openpyxl")

    # Normalizza nomi colonne: converti 1965.0 → "1965"
    new_cols = []
    for c in df.columns:
        if _is_year(c):
            new_cols.append(str(int(float(str(c)))))
        else:
            new_cols.append(str(c).strip())
    df.columns = new_cols

    year_cols = [c for c in df.columns if c.isdigit() and 1900 < int(c) < 2100]
    if not year_cols:
        raise ValueError(f"Nessuna colonna anno nel foglio '{sheet}'")

    country_col = df.columns[0]
    # L'intestazione della colonna paese contiene l'unità di misura
    # (es. "Exajoules", "Terawatt-hours", "Megawatts")
    unit_label = (country_col if country_col
                  and not country_col.lower().startswith(("unnamed", "country"))
                  else sheet)
    _EI_UNITS[key] = unit_label

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

    _cancel_event.clear()
    base = DATA_SOURCES["World Bank – WDI (Energia & Ambiente)"]["api_base"]
    frames = []
    for code in codes:
        if _cancel_event.is_set():
            raise InterruptedError("Caricamento interrotto dall'utente.")
        url = f"{base}/{code}?format=json&per_page=20000&date=1960:2026"
        try:
            r = requests.get(url, timeout=60)
            r.raise_for_status()
            meta, records = r.json()
            # Gestisce eventuale paginazione
            all_records = list(records)
            for page in range(2, meta["pages"] + 1):
                r2 = requests.get(url + f"&page={page}", timeout=60)
                all_records.extend(r2.json()[1])
            for rec in all_records:
                if rec["value"] is not None:
                    frames.append({
                        "Country Name":   rec["country"]["value"],
                        "Country Code":   rec["countryiso3code"],
                        "Indicator Name": rec["indicator"]["value"],
                        "Indicator Code": rec["indicator"]["id"],
                        "year":           int(rec["date"]),
                        "value":          float(rec["value"]),
                    })
        except Exception:
            continue

    if not frames:
        return pd.DataFrame()

    df = pd.DataFrame(frames)
    _CACHE[mem_key] = df
    _disk_save(disk_key, df)
    return df


# ── Helpers ──────────────────────────────────────────────────────────────────────
def _apply_filters(df: pd.DataFrame, meta: dict,
                   countries: list, years: list) -> pd.DataFrame:
    cc, yc = meta["country_col"], meta["year_col"]
    mask = pd.Series(True, index=df.index)
    if countries:
        mask &= df[cc].isin(countries)
    if years and len(years) == 2:
        mask &= df[yc].between(years[0], years[1])
    return df[mask]


def _make_fig(sub: pd.DataFrame, x: str, y: str, color: str,
              chart_type: str, title: str = "", y_label: str | None = None) -> object:
    kw = dict(data_frame=sub, x=x, y=y, color=color, title=title)
    if y_label:
        kw["labels"] = {y: y_label}
    return {
        "Line":    lambda: px.line(**kw),
        "Bar":     lambda: px.bar(**kw, barmode="group"),
        "Scatter": lambda: px.scatter(**kw),
        "Area":    lambda: px.area(**kw),
    }.get(chart_type, lambda: px.line(**kw))()


_MAX_CHART_COUNTRIES = 20   # massimo paesi in un grafico senza filtro esplicito


def _auto_limit_countries(df: pd.DataFrame, cc: str,
                           max_n: int = _MAX_CHART_COUNTRIES):
    """Limita a max_n paesi (quelli con più righe/dati) se ce ne sono troppi.
    Restituisce (df_limitato, banner_html | None)."""
    n = df[cc].nunique()
    if n <= max_n:
        return df, None
    keep = df[cc].value_counts().nlargest(max_n).index.tolist()
    banner = dbc.Alert(
        [html.Span(f"💡 Nessun paese selezionato — visualizzati i {max_n} paesi con "
                   "più dati disponibili. "),
         html.Span("Seleziona paesi specifici dalla sidebar per un confronto mirato.",
                   className="text-muted")],
        color="info", className="py-2 small mb-2",
    )
    return df[df[cc].isin(keep)], banner


def _no_data_alert(sub: pd.DataFrame, requested: list, cc: str) -> "dbc.Alert | None":
    """Restituisce un Alert se alcuni paesi richiesti non hanno valori numerici nel sub."""
    if not requested:
        return None
    present = set(sub[cc].dropna().unique()) if not sub.empty else set()
    missing = [c for c in requested if c not in present]
    if not missing:
        return None
    names = ", ".join(missing[:8]) + ("…" if len(missing) > 8 else "")
    return dbc.Alert(
        [html.Strong("Nessun dato disponibile per: "),
         names,
         html.Span(" — nessun valore numerico per la metrica selezionata.",
                   className="text-muted ms-1")],
        color="warning", className="py-2 small", dismissable=True,
    )


# WDI options loaded once at startup
try:
    _WDI_OPTS = wdi_series_options()
except Exception:
    _WDI_OPTS = []

def _ei_local_bytes() -> bytes | None:
    """Restituisce i bytes del file EI locale se esiste ed è un vero xlsx."""
    local_path = Path(DATA_SOURCES["Energy Institute – Stats Review"]["local_file"])
    if not local_path.exists():
        return None
    try:
        raw = local_path.read_bytes()
        pd.ExcelFile(io.BytesIO(raw), engine="openpyxl")
        return raw
    except Exception:
        return None

# Controlla solo l'esistenza del file all'avvio (evita di leggere 14 MB in fase di startup)
_EI_LOCAL_AVAILABLE = Path(DATA_SOURCES["Energy Institute – Stats Review"]["local_file"]).exists()


# ── Layout ───────────────────────────────────────────────────────────────────────
_UPLOAD_STYLE = {
    "borderWidth": "1px", "borderStyle": "dashed", "borderRadius": "5px",
    "textAlign": "center", "padding": "10px", "cursor": "pointer",
    "color": "#6c757d", "fontSize": "0.85rem",
}

sidebar = dbc.Col(
    id="sidebar-col",
    className="app-sidebar",
    width=3,
    children=[
        # Intestazione drawer (solo mobile)
        html.Div(
            className="d-flex justify-content-between align-items-center d-md-none mb-3",
            children=[
                html.Span("Filtri & Controlli",
                          style={"color": "#f1f5f9", "fontWeight": 600,
                                 "fontSize": "0.9rem", "letterSpacing": "0.02em"}),
                dbc.Button("✕", id="btn-sidebar-close", color="link", size="sm",
                           style={"color": "#94a3b8", "fontSize": "1.1rem",
                                  "padding": "0 4px", "lineHeight": "1"}),
            ],
        ),
        html.H5("Fonte dati", className="fw-bold"),
        dcc.Dropdown(
            id="source-select",
            options=[{"label": k, "value": k} for k in SOURCE_NAMES],
            value=SOURCE_NAMES[0],
            clearable=False,
            className="mb-2",
        ),

        html.Div(id="ei-options", style={"display": "none"}, children=[
            # Messaggio file locale trovato
            html.Div(
                id="ei-local-status",
                children=dbc.Alert("✅ File locale trovato — verrà caricato automaticamente.",
                                   color="success", className="py-2 small")
                         if _EI_LOCAL_AVAILABLE else
                         dbc.Alert(
                             ["File non trovato. Scaricalo da ",
                              html.A("energyinst.org", href="https://www.energyinst.org/statistical-review",
                                     target="_blank"),
                              " e mettilo nella cartella dell'app con il nome ",
                              html.Code("EI-Stats-Review-ALL-data.xlsx"),
                              ", oppure caricalo qui sotto."],
                             color="warning", className="py-2 small"),
            ),
            # Uploader (fallback se file locale non presente)
            html.Div(
                id="ei-upload-wrap",
                style={"display": "none"} if _EI_LOCAL_AVAILABLE else {"display": "block"},
                children=[
                    dcc.Upload(
                        id="ei-upload",
                        children=html.Div("Trascina o clicca per caricare il file .xlsx"),
                        style=_UPLOAD_STYLE,
                        multiple=False,
                    ),
                ],
            ),
        ]),

        html.Div(id="wdi-options", style={"display": "none"}, children=[
            html.Label("Indicatori energetici:", className="small"),
            dcc.Dropdown(
                id="wdi-indicators",
                options=_WDI_OPTS,
                value=[c for c in WDI_DEFAULT_CODES
                       if any(o["value"] == c for o in _WDI_OPTS)]
                      or [o["value"] for o in _WDI_OPTS[:5]],
                multi=True,
                placeholder="Seleziona indicatori…",
            ),
        ]),

        dbc.Button("Carica / Aggiorna dati", id="btn-load",
                   color="primary", size="sm", className="w-100 mt-3"),
        dbc.Button("⛔  Interrompi caricamento", id="btn-cancel",
                   color="danger", size="sm", className="w-100 mt-1 mb-3",
                   style={"display": "none"}),

        # Globe loader centrato nella sidebar (visibile durante il caricamento)
        html.Div(
            id="sidebar-spinner",
            style={"display": "none", "position": "fixed",
                   "top": "50%", "left": "12.5%",
                   "transform": "translate(-50%, -50%)", "zIndex": 1050},
            children=html.Div(
                className="globe-host",
                style={"--gl-size": "110", "--gl-ink": "#60a5fa", "--gl-bg": "#0f172a"},
            ),
        ),

        html.Hr(),
        html.H5("Filtri globali", className="fw-bold"),

        # Foglio EI — visibile solo quando la fonte è Energy Institute
        html.Div(id="ei-sheet-filter", style={"display": "none"}, children=[
            html.Label("Foglio / Metrica EI", className="small"),
            dcc.Dropdown(id="ei-sheet", placeholder="Seleziona foglio…",
                         className="mb-3", clearable=False),
        ]),

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
            # Metrica (solo OWID — nascosta per EI e WDI che hanno chart-indicator)
            html.Div(id="chart-metric-wrap", children=[
                html.Label("Metrica", className="small"),
                dcc.Dropdown(id="chart-metric", className="mb-2", clearable=False, multi=True),
            ]),
            html.Div(id="chart-indicator-wrap", style={"display": "none"}, children=[
                html.Label("Indicatore", className="small"),
                dcc.Dropdown(id="chart-indicator", className="mb-2", clearable=False, multi=True),
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
                    {"label": " Grafico a torta", "value": "pie"},
                    {"label": " Scatter (X vs Y)", "value": "scatter"},
                    {"label": " Radar multi-metrica", "value": "radar"},
                    {"label": " Tabella dati",     "value": "table"},
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

            html.Div(id="cmp-metric-a-wrap", children=[
                html.Label("Metrica principale", className="small"),
                dcc.Dropdown(id="cmp-metric-a", className="mb-2", clearable=False),
            ]),

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

            # Anno (Snapshot, Torta, Scatter, Radar)
            html.Div(id="cmp-year-wrap", style={"display": "none"}, children=[
                html.Label("Anno di riferimento", className="small"),
                dcc.Dropdown(id="cmp-year", placeholder="Seleziona anno…",
                             className="mb-2", clearable=False),
            ]),

            # Periodicità (solo Tabella dati)
            html.Div(id="cmp-period-wrap", style={"display": "none"}, children=[
                html.Label("Periodicità", className="small"),
                dcc.Dropdown(
                    id="cmp-period",
                    options=[
                        {"label": "Annuale",                    "value": 1},
                        {"label": "Biennale (ogni 2 anni)",     "value": 2},
                        {"label": "Quinquennale (ogni 5 anni)", "value": 5},
                    ],
                    value=1, clearable=False, className="mb-2",
                ),
            ]),
        ]),

        # ── Pannello Statistiche ───────────────────────────────────────────
        html.Div(id="panel-stats", style={"display": "none"}, children=[
            html.Hr(),
            html.H5("Statistiche", className="fw-bold"),
            html.Label("Metrica / Indicatore", className="small"),
            dcc.Dropdown(id="stats-metric", className="mb-2", clearable=False,
                         placeholder="Seleziona metrica…"),
        ]),
    ],
)

main_area = dbc.Col(
    id="main-col",
    className="app-main",
    width=9,
    children=[
        html.Div(id="alert-area"),
        html.Div(id="summary-banner", className="mb-3"),
        html.Div(className="tab-container", children=[
            dcc.Tabs(
                id="main-tabs", value="tab-table",
                children=[
                    dcc.Tab(label="📋 Tabella dati",  value="tab-table"),
                    dcc.Tab(label="📈 Grafico",        value="tab-chart"),
                    dcc.Tab(label="⚖️  Confronto",     value="tab-compare"),
                    dcc.Tab(label="📊 Statistiche",    value="tab-stats"),
                ],
            ),
            html.Div(id="tab-content", className="tab-content-area"),
        ]),
        html.Div(
            id="download-btn-wrap",
            style={"display": "none"},
            children=dbc.Button(
                "⬇️  Scarica CSV",
                id="btn-download",
                color="secondary",
                size="sm",
                className="mt-2",
            ),
        ),
        dcc.Download(id="download-csv"),
        dcc.Store(id="store-key"),
        dcc.Store(id="store-meta"),
    ],
)

app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.FLATLY,
        "https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap",
    ],
    suppress_callback_exceptions=True,
    meta_tags=[
        {"name": "viewport", "content": "width=device-width, initial-scale=1"},
        {"name": "mobile-web-app-capable", "content": "yes"},
        {"name": "apple-mobile-web-app-capable", "content": "yes"},
        {"name": "apple-mobile-web-app-status-bar-style", "content": "black-translucent"},
        {"name": "apple-mobile-web-app-title", "content": "EnergyApp"},
        {"name": "theme-color", "content": "#2563eb"},
    ],
)
server = app.server  # esposto per gunicorn

# PWA: collega manifest e registra service worker
app.index_string = """<!DOCTYPE html>
<html>
<head>
    {%metas%}
    <title>Energy Data Analysis Platform</title>
    <link rel="manifest" href="/assets/manifest.json">
    <link rel="apple-touch-icon" href="/assets/icon-192.png">
    {%favicon%}
    {%css%}
</head>
<body>
    {%app_entry%}
    <footer>
        {%config%}
        {%scripts%}
        {%renderer%}
    </footer>
    <script>
      if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('/assets/sw.js');
      }
    </script>
</body>
</html>"""

app.layout = dbc.Container(fluid=True, className="p-0", children=[
    dbc.Row(dbc.Col(
        html.Div(
            className="app-header",
            children=[
                # Hamburger (solo mobile)
                dbc.Button(
                    "☰",
                    id="btn-hamburger",
                    color="link",
                    className="d-md-none",
                    style={"color": "white", "fontSize": "1.4rem", "lineHeight": "1",
                           "padding": "0 0.75rem 0 0", "textDecoration": "none"},
                ),
                html.H3("⚡ Energy Data Analysis Platform"),
                # Pulsante collassa/espandi sidebar (solo desktop)
                dbc.Button(
                    "◀",
                    id="btn-sidebar-desktop",
                    color="link",
                    className="d-none d-md-inline-flex",
                ),
            ],
        ),
    ), className="g-0"),
    # Backdrop overlay (mobile)
    html.Div(id="sidebar-backdrop", className="sidebar-backdrop"),
    dbc.Row([sidebar, main_area], className="g-0"),
])


# ── Callbacks ────────────────────────────────────────────────────────────────────

# 0. Drawer sidebar mobile (hamburger apre, ✕ o backdrop chiude)
app.clientside_callback(
    """
    function(n_open, n_close, n_backdrop) {
        const ctx = dash_clientside.callback_context;
        if (!ctx.triggered.length) return [window.dash_clientside.no_update, window.dash_clientside.no_update];
        const tid = ctx.triggered[0].prop_id.split('.')[0];
        if (tid === 'btn-hamburger') {
            return ['app-sidebar sidebar-open', 'sidebar-backdrop show'];
        }
        return ['app-sidebar', 'sidebar-backdrop'];
    }
    """,
    Output("sidebar-col",      "className", allow_duplicate=True),
    Output("sidebar-backdrop", "className"),
    Input("btn-hamburger",     "n_clicks"),
    Input("btn-sidebar-close", "n_clicks"),
    Input("sidebar-backdrop",  "n_clicks"),
    prevent_initial_call=True,
)


# 0b. Toggle sidebar desktop (◀ collassa / ▶ espandi)
app.clientside_callback(
    """
    function(n, sidebar_cls, main_cls) {
        if (!n) return [window.dash_clientside.no_update,
                        window.dash_clientside.no_update,
                        window.dash_clientside.no_update];
        var collapsed = sidebar_cls && sidebar_cls.includes('sidebar-collapsed');
        if (collapsed) {
            return ['app-sidebar', 'app-main', '◀'];
        }
        return ['app-sidebar sidebar-collapsed', 'app-main sidebar-collapsed', '▶'];
    }
    """,
    Output("sidebar-col",       "className", allow_duplicate=True),
    Output("main-col",          "className"),
    Output("btn-sidebar-desktop", "children"),
    Input("btn-sidebar-desktop", "n_clicks"),
    State("sidebar-col",         "className"),
    State("main-col",            "className"),
    prevent_initial_call=True,
)


# 1. Mostra/nascondi opzioni specifiche per fonte
@app.callback(
    Output("ei-options",      "style"),
    Output("wdi-options",     "style"),
    Output("ei-sheet-filter", "style"),
    Input("source-select", "value"),
)
def toggle_source_panels(source):
    stype = DATA_SOURCES[source]["type"]
    show, hide = {"display": "block"}, {"display": "none"}
    return (show if stype == "ei"  else hide,
            show if stype == "wdi" else hide,
            show if stype == "ei"  else hide)


# 2. Aggiorna lista fogli EI (da file locale o upload)
@app.callback(
    Output("ei-sheet", "options"),
    Output("ei-sheet", "value"),
    Input("source-select", "value"),
    Input("ei-upload", "contents"),
)
def update_ei_sheets(source, contents):
    # Carica i fogli solo se è selezionata la fonte EI
    raw = None
    if DATA_SOURCES.get(source, {}).get("type") == "ei":
        raw = _ei_local_bytes()
    if raw is None and contents:
        _, b64 = contents.split(",", 1)
        raw = base64.b64decode(b64)
    if raw is None:
        return [], None
    valid = ei_valid_sheets(raw)
    if not valid:  # fallback: mostra tutti i fogli se il filtro non trova nulla
        valid = ei_sheets(raw)
    opts = [{"label": s, "value": s} for s in valid]
    return opts, valid[0] if valid else None


# 3. Caricamento dati
@app.callback(
    Output("store-key",  "data"),
    Output("store-meta", "data"),
    Output("alert-area", "children"),
    Input("btn-load",  "n_clicks"),
    Input("ei-sheet",  "value"),          # auto-aggiorna al cambio foglio EI
    State("source-select",   "value"),
    State("ei-upload",       "contents"),
    State("wdi-indicators",  "value"),
    prevent_initial_call=True,
)
def load_data(n_clicks, ei_sheet, source, ei_contents, wdi_codes):
    # Se il trigger è ei-sheet ma la fonte non è EI, ignora
    triggered = dash.callback_context.triggered[0]["prop_id"].split(".")[0] \
                if dash.callback_context.triggered else ""
    if triggered == "ei-sheet" and DATA_SOURCES.get(source, {}).get("type") != "ei":
        return no_update, no_update, no_update

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
            # Usa file locale se disponibile, altrimenti l'upload
            fb = _ei_local_bytes()
            if fb is None and ei_contents:
                _, b64 = ei_contents.split(",", 1)
                fb = base64.b64decode(b64)
            if fb is None:
                return no_update, no_update, dbc.Alert(
                    "File EI non trovato. Caricalo manualmente o metti il file nella cartella dell'app.",
                    color="warning")
            if not ei_sheet:
                return no_update, no_update, dbc.Alert(
                    "Seleziona un foglio dalla lista.", color="warning")
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

    except InterruptedError:
        return no_update, no_update, dbc.Alert(
            "⛔  Caricamento interrotto.", color="secondary", duration=4000, dismissable=True)
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
    Output("cmp-countries",        "value"),
    Output("cmp-metric-a",         "options"),
    Output("cmp-metric-a",         "value"),
    Output("cmp-metric-b",         "options"),
    Output("cmp-metric-b",         "value"),
    Output("cmp-metrics-extra",    "options"),
    Output("cmp-metrics-extra",    "value"),
    Output("cmp-year",             "options"),
    Output("cmp-year",             "value"),
    Output("cmp-metric-a-wrap",    "style"),
    Output("stats-metric",         "options"),
    Output("stats-metric",         "value"),
    Output("summary-banner",       "children"),
    Output("chart-metric-wrap",    "style"),
    Input("store-key",  "data"),
    State("store-meta", "data"),
    State("filter-countries", "value"),
)
def update_controls(key, meta, current_countries):
    none14 = ([], [], 1960, 2026, [2000, 2026],
              [], [], [], [], {"display": "none"},
              [], [], [], None, [], None, [], [], [], None, {"display": "none"},
              [], None, html.Div(), {"display": "block"})
    if not key or not meta or key not in _CACHE:
        return none14

    df  = _CACHE[key]
    cc  = meta["country_col"]
    yc  = meta["year_col"]
    fmt = meta["format"]
    ic  = meta.get("indicator_col")

    countries    = sorted(df[cc].dropna().unique())
    country_opts = [{"label": c, "value": c} for c in countries]

    # Conserva la selezione corrente se i paesi esistono ancora nel nuovo dataset
    countries_set    = set(countries)
    preserved_countries = [c for c in (current_countries or []) if c in countries_set]

    y_min, y_max = int(df[yc].min()), int(df[yc].max())
    yr_val       = [max(y_min, y_max - 25), y_max]
    year_opts    = [{"label": str(y), "value": y}
                    for y in sorted(df[yc].dropna().unique(), reverse=True)]

    if fmt == "wide":
        num = [c for c in df.select_dtypes(include="number").columns if c != yc]
        # Priorità: colonne note per prime, poi le restanti in ordine originale
        priority_present = [c for c in OWID_PRIORITY if c in num]
        other = [c for c in num if c not in set(OWID_PRIORITY)]
        ordered = priority_present + other
        m_opts    = [{"label": OWID_LABELS.get(c, c), "value": c} for c in ordered]
        m_val     = priority_present[0] if priority_present else (num[0] if num else None)
        ind_opts  = []
        ind_val   = None
        ind_style = {"display": "none"}
    else:
        m_opts = [{"label": "value", "value": "value"}]
        m_val  = "value"
        if ic and ic in df.columns:
            # WDI: indicatori come metriche
            inds     = sorted(df[ic].dropna().unique())
            ind_opts = [{"label": i, "value": i} for i in inds]
            ind_val  = inds[0] if inds else None
            ind_style = {"display": "block"}
            m_opts   = ind_opts
            m_val    = inds[0] if inds else None
        elif meta.get("type") == "ei":
            # EI: usa tutti i fogli validi come metriche confrontabili
            raw = _ei_local_bytes()
            if raw is not None:
                sheets = ei_valid_sheets(raw) or ei_sheets(raw)
                m_opts = [{"label": s, "value": s} for s in sheets]
                m_val  = meta.get("sheet", sheets[0] if sheets else None)
            ind_opts, ind_val, ind_style = [], None, {"display": "none"}
        else:
            ind_opts, ind_val, ind_style = [], None, {"display": "none"}

    n_metrics = len(m_opts)
    banner = dbc.Row([
        dbc.Col(dbc.Card(dbc.CardBody([
            html.H6("Righe totali", className="card-subtitle text-muted"),
            html.H4(f"{len(df):,}", className="card-title"),
        ])), xs=6, md=3),
        dbc.Col(dbc.Card(dbc.CardBody([
            html.H6("Paesi / Regioni", className="card-subtitle text-muted"),
            html.H4(str(len(countries)), className="card-title"),
        ])), xs=6, md=3),
        dbc.Col(dbc.Card(dbc.CardBody([
            html.H6("Anni disponibili", className="card-subtitle text-muted"),
            html.H4(f"{y_min} – {y_max}", className="card-title"),
        ])), xs=6, md=3),
        dbc.Col(dbc.Card(dbc.CardBody([
            html.H6("Metriche / Indicatori", className="card-subtitle text-muted"),
            html.H4(str(n_metrics), className="card-title"),
        ])), xs=6, md=3),
    ], className="g-2")

    if fmt == "wide":
        pp = [c for c in OWID_PRIORITY if any(o["value"] == c for o in m_opts)]
        cmp_m_val = pp[0] if pp else (m_opts[0]["value"] if m_opts else None)
        cmp_m_b   = pp[1] if len(pp) > 1 else (m_opts[1]["value"] if len(m_opts) > 1 else cmp_m_val)
    else:
        cmp_m_val = m_opts[0]["value"] if m_opts else None
        cmp_m_b   = m_opts[1]["value"] if len(m_opts) > 1 else cmp_m_val

    # Nasconde "Metrica principale" per fonti con una sola metrica (EI per foglio)
    metric_wrap_style = {"display": "block"} if len(m_opts) > 1 else {"display": "none"}

    # chart-metric-wrap: visibile solo per OWID (wide); per EI e WDI usa chart-indicator
    chart_metric_wrap_style = {"display": "block"} if fmt == "wide" else {"display": "none"}

    # stats-metric: per WDI usa indicatori, per gli altri usa m_opts
    stats_opts = ind_opts if (ic and ind_opts) else m_opts
    stats_val  = ind_val  if (ic and ind_opts) else m_val

    m_val_list   = [m_val]   if m_val   is not None else []
    ind_val_list = [ind_val] if ind_val is not None else []

    return (country_opts, preserved_countries, y_min, y_max, yr_val,
            m_opts, m_val_list, ind_opts, ind_val_list, ind_style,
            country_opts, [], m_opts, cmp_m_val, m_opts, cmp_m_b,
            m_opts, [], year_opts, y_max, metric_wrap_style,
            stats_opts, stats_val, banner, chart_metric_wrap_style)


# 4b. Mostra/nascondi pannelli sidebar in base al tab attivo
@app.callback(
    Output("panel-chart",   "style"),
    Output("panel-compare", "style"),
    Output("panel-stats",   "style"),
    Input("main-tabs", "value"),
)
def toggle_sidebar_panels(tab):
    show, hide = {"display": "block"}, {"display": "none"}
    return (show if tab == "tab-chart"   else hide,
            show if tab == "tab-compare" else hide,
            show if tab == "tab-stats"   else hide)


# 4c. Mostra/nascondi controlli confronto in base alla modalità
@app.callback(
    Output("cmp-metric-b-wrap",      "style"),
    Output("cmp-metrics-extra-wrap", "style"),
    Output("cmp-year-wrap",          "style"),
    Output("cmp-period-wrap",        "style"),
    Input("cmp-mode", "value"),
)
def toggle_compare_controls(mode):
    show, hide = {"display": "block"}, {"display": "none"}
    return (
        show if mode == "scatter"                                   else hide,  # metrica Y
        show if mode in ("temporal", "snapshot", "pie", "radar", "table") else hide,  # metriche extra
        show if mode in ("snapshot", "pie", "scatter", "radar")     else hide,  # anno
        show if mode == "table"                                     else hide,  # periodicità
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
    Input("cmp-period",         "value"),
    Input("stats-metric",       "value"),
    Input("store-key",          "data"),
    State("store-meta", "data"),
)
def render_tab(tab, countries, years, metric, indicator, chart_type,
               cmp_mode, cmp_countries, cmp_metric_a, cmp_metric_b,
               cmp_metrics_extra, cmp_year, cmp_period, stats_metric, key, meta):
    if not key or not meta or key not in _CACHE:
        return html.Div([
            html.Div(
                className="globe-host",
                style={"--gl-size": "140", "--gl-ink": "#2563eb", "--gl-bg": "#f1f5f9"},
            ),
            html.P("Seleziona una fonte e premi «Carica / Aggiorna dati» per iniziare.",
                   className="text-muted small mt-3"),
        ], style={"display": "flex", "flexDirection": "column",
                  "alignItems": "center", "justifyContent": "center",
                  "minHeight": "320px"})

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
        total_rows = len(filtered)
        ROW_LIMIT  = 10_000
        display_df = filtered.iloc[:ROW_LIMIT]
        cols = [{"name": c, "id": c} for c in display_df.columns]
        data = display_df.to_dict("records")
        notices = []
        if not countries:
            notices.append(dbc.Alert(
                "💡 Nessun paese selezionato — seleziona uno o più paesi dalla sidebar "
                "per filtrare i dati.",
                color="info", className="py-2 small mb-2",
            ))
        if total_rows > ROW_LIMIT:
            notices.append(dbc.Alert(
                f"⚠️ Dati troncati: vengono mostrate le prime {ROW_LIMIT:,} righe "
                f"su {total_rows:,} totali. Filtra per paese e/o intervallo anni per ridurre.",
                color="warning", className="py-2 small mb-2",
            ))
        return html.Div(notices + [
            html.P(f"{total_rows:,} righe totali  •  mostrando le prime "
                   f"{min(total_rows, ROW_LIMIT):,}",
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
        ])

    # ── Grafico ────────────────────────────────────────────────────────────────
    if tab == "tab-chart":
        try:
            auto_note = None
            single_country = len(countries or []) == 1
            legend_title = "Paese"

            if fmt == "wide":
                metrics_sel = metric if isinstance(metric, list) else ([metric] if metric else [])
                valid_metrics = [m for m in metrics_sel if m in filtered.columns]
                # Fallback: se nessuna metrica valida, usa la prima disponibile
                if not valid_metrics:
                    num_cols = [c for c in filtered.select_dtypes(include="number").columns
                                if c != yc]
                    priority_cols = [c for c in OWID_PRIORITY if c in num_cols]
                    fallback = (priority_cols[0] if priority_cols
                                else (num_cols[0] if num_cols else None))
                    valid_metrics = [fallback] if fallback else []
                if not valid_metrics:
                    return dbc.Alert("Nessuna metrica numerica disponibile nei dati.", color="warning")

                if single_country and len(valid_metrics) > 1:
                    # Un solo paese + più metriche: tutte sullo stesso grafico (color=metrica)
                    frames = []
                    for m in valid_metrics:
                        sub_m = (filtered[[cc, yc, m]].dropna()
                                 .rename(columns={m: "value"}))
                        sub_m["metrica"] = OWID_LABELS.get(m, m)
                        frames.append(sub_m)
                    sub = pd.concat(frames, ignore_index=True)
                    warn = _no_data_alert(sub, countries, cc)
                    if sub.empty:
                        return dbc.Alert("Nessun dato per le metriche selezionate con i filtri correnti.",
                                         color="warning")
                    fig = _make_fig(sub, yc, "value", "metrica", chart_type or "Line",
                                    f"{countries[0]} — Confronto metriche")
                    fig.update_layout(yaxis_title="Valore")
                    legend_title = "Metrica"
                else:
                    metric = valid_metrics[0]
                    sub = filtered[[cc, yc, metric]].dropna()
                    if not countries:
                        sub, auto_note = _auto_limit_countries(sub, cc)
                    metric_label = OWID_LABELS.get(metric, metric)
                    warn = _no_data_alert(sub, countries, cc)
                    if sub.empty:
                        return dbc.Alert("Nessun dato per la metrica selezionata con i filtri correnti.",
                                         color="warning")
                    fig = _make_fig(sub, yc, metric, cc, chart_type or "Line", metric_label,
                                    y_label=metric_label)
                    fig.update_layout(yaxis_title=metric_label)
            else:
                # Per WDI/EI: determina gli indicatori da usare
                inds_sel = indicator if isinstance(indicator, list) else ([indicator] if indicator else [])
                if ic and ic in filtered.columns:
                    avail_inds = set(filtered[ic].dropna().unique())
                    valid_inds = [i for i in inds_sel if i in avail_inds]
                    if not valid_inds:
                        avail = sorted(avail_inds)
                        valid_inds = [avail[0]] if avail else []
                else:
                    valid_inds = []

                if single_country and len(valid_inds) > 1:
                    # Un solo paese + più indicatori: tutti sullo stesso grafico (color=indicatore)
                    frames = []
                    for ind in valid_inds:
                        sub_i = (filtered[filtered[ic] == ind]
                                 .dropna(subset=["value"]).assign(metrica=str(ind)))
                        frames.append(sub_i)
                    sub = pd.concat(frames, ignore_index=True)
                    warn = _no_data_alert(sub, countries, cc)
                    if sub.empty:
                        return dbc.Alert("Nessun dato per gli indicatori selezionati.", color="warning")
                    unit_label = _EI_UNITS.get(key, "Valore") if meta.get("type") == "ei" else "Valore"
                    fig = _make_fig(sub, yc, "value", "metrica", chart_type or "Line",
                                    f"{countries[0]} — Confronto indicatori")
                    fig.update_layout(yaxis_title=unit_label)
                    legend_title = "Indicatore"
                else:
                    target_ind = valid_inds[0] if valid_inds else None
                    if target_ind and ic and ic in filtered.columns:
                        sub = filtered[filtered[ic] == target_ind]
                    else:
                        sub = filtered
                    sub = sub.dropna(subset=["value"])
                    if not countries:
                        sub, auto_note = _auto_limit_countries(sub, cc)
                    warn = _no_data_alert(sub, countries, cc)
                    if sub.empty:
                        return dbc.Alert("Nessun dato per l'indicatore selezionato.", color="warning")
                    title = target_ind or (meta.get("sheet") or "value")
                    unit_label = _EI_UNITS.get(key, "Valore") if meta.get("type") == "ei" else "Valore"
                    fig = _make_fig(sub, yc, "value", cc, chart_type or "Line", title,
                                    y_label=unit_label)
                    fig.update_layout(yaxis_title=unit_label)

            fig.update_layout(hovermode="x unified", legend_title_text=legend_title,
                              margin={"t": 40}, height=600)
            graph = dcc.Graph(figure=fig, style={"height": "600px"})
            extras = [x for x in [auto_note, warn] if x]
            return html.Div(extras + [graph]) if extras else graph

        except Exception as exc:
            return dbc.Alert(
                [html.Strong("Errore nella generazione del grafico: "),
                 html.Code(str(exc))],
                color="danger", className="small",
            )

    # ── Confronto ──────────────────────────────────────────────────────────────
    if tab == "tab-compare":
        if not cmp_metric_a:
            return dbc.Alert("Seleziona almeno la metrica principale nella sidebar.",
                             color="info")

        # Convalida che cmp_metric_a sia compatibile con la fonte corrente
        if fmt == "wide":
            num_cols_cmp = [c for c in filtered.select_dtypes(include="number").columns
                            if c != yc]
            if cmp_metric_a not in num_cols_cmp:
                priority_cols_cmp = [c for c in OWID_PRIORITY if c in num_cols_cmp]
                cmp_metric_a = (priority_cols_cmp[0] if priority_cols_cmp
                                else (num_cols_cmp[0] if num_cols_cmp else None))
            if not cmp_metric_a:
                return dbc.Alert("Nessuna metrica disponibile per il confronto.", color="warning")
        elif ic and ic in filtered.columns:
            avail_inds = set(filtered[ic].dropna().unique())
            if cmp_metric_a not in avail_inds:
                cmp_metric_a = next(iter(avail_inds), None)
            if not cmp_metric_a:
                return dbc.Alert("Nessun indicatore disponibile per il confronto.", color="warning")

        # Paesi da usare: selezione specifica o auto-limit ai top MAX_CHART_COUNTRIES
        if cmp_countries:
            cmp_cc = cmp_countries
            cmp_auto_note = None
        elif countries:
            cmp_cc = countries
            cmp_auto_note = None
        else:
            all_cmp = filtered[cc].dropna().unique().tolist()
            if len(all_cmp) > _MAX_CHART_COUNTRIES:
                cmp_cc = (filtered[cc].value_counts()
                          .nlargest(_MAX_CHART_COUNTRIES).index.tolist())
                cmp_auto_note = dbc.Alert(
                    [html.Span(f"💡 Nessun paese selezionato — confronto limitato ai "
                               f"{_MAX_CHART_COUNTRIES} paesi con più dati. "),
                     html.Span("Usa i filtri per scegliere paesi specifici.",
                               className="text-muted")],
                    color="info", className="py-2 small mb-2",
                )
            else:
                cmp_cc = all_cmp
                cmp_auto_note = None

        sub_base = filtered[filtered[cc].isin(cmp_cc)]
        if sub_base.empty:
            return dbc.Alert("Nessun dato per i paesi selezionati.", color="warning")

        is_ei = meta.get("type") == "ei"
        _ei_raw_bytes = _ei_local_bytes() if is_ei else None

        def _get_series(df_in, metric):
            """Estrae (country, year, value) per una metrica, sia wide che long."""
            if fmt == "wide":
                if metric not in df_in.columns:
                    return pd.DataFrame(columns=[cc, yc, "value"])
                return df_in[[cc, yc, metric]].dropna().rename(columns={metric: "value"})
            elif is_ei:
                # Per EI: metric = nome foglio; carica dinamicamente se diverso da quello corrente
                try:
                    sheet_df = load_ei(_ei_raw_bytes, metric) if _ei_raw_bytes else pd.DataFrame()
                except Exception:
                    return pd.DataFrame(columns=[cc, yc, "value"])
                filtered_sheet = _apply_filters(sheet_df, meta, countries, years)
                sub = filtered_sheet[filtered_sheet[cc].isin(cmp_cc)]
                return sub[[cc, yc, "value"]].dropna()
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

        def _period_years(years_series, period):
            """Anni selezionati a partire dal primo disponibile, ogni `period` anni."""
            years_sorted = sorted(years_series.dropna().unique())
            if not years_sorted or period <= 1:
                return years_sorted
            base = years_sorted[0]
            return [y for y in years_sorted if (y - base) % period == 0]

        graphs = []
        if cmp_auto_note:
            graphs.append(cmp_auto_note)

        def _mlabel(m):
            if fmt == "wide":
                return OWID_LABELS.get(m, m)
            if is_ei:
                ei_sig = hashlib.md5((_ei_raw_bytes or b"")[:512]).hexdigest()[:8]
                unit = _EI_UNITS.get(f"ei|{m}|{ei_sig}", "")
                return f"{m} ({unit})" if unit and unit != m else str(m)
            return str(m)

        def _unavailable(m):
            return dbc.Alert(
                [html.Strong("Nessun dato — "), _mlabel(m),
                 html.Span(" — nessun valore per i paesi/anni selezionati.",
                           className="text-muted ms-1")],
                color="warning", className="py-2 small", dismissable=True,
            )

        def _data_warnings(s, m):
            """Avvisi dati mancanti per una metrica: paesi assenti + serie vuota."""
            alerts = []
            if s.empty:
                alerts.append(_unavailable(m))
                return alerts
            # Paesi richiesti senza nessun dato per questa metrica
            present = set(s[cc].dropna().unique())
            missing = [p for p in cmp_cc if p not in present]
            if missing:
                names = ", ".join(missing[:8]) + ("…" if len(missing) > 8 else "")
                alerts.append(dbc.Alert(
                    [html.Strong(f"Dati assenti per «{_mlabel(m)}»: "),
                     names],
                    color="warning", className="py-2 small", dismissable=True,
                ))
            return alerts

        # ── Temporale ──────────────────────────────────────────────────────
        if cmp_mode == "temporal":
            metrics_plot = [cmp_metric_a] + (cmp_metrics_extra or [])
            for m in metrics_plot:
                s = _get_series(sub_base, m)
                graphs.extend(_data_warnings(s, m))
                if s.empty:
                    continue
                lbl = _mlabel(m)
                if is_ei:
                    ei_sig = hashlib.md5((_ei_raw_bytes or b"")[:512]).hexdigest()[:8]
                    y_lbl = _EI_UNITS.get(f"ei|{m}|{ei_sig}", m)
                else:
                    y_lbl = "Valore"
                fig = px.line(data_frame=s, x=yc, y="value", color=cc,
                              title=lbl,
                              labels={yc: "Anno", "value": y_lbl, cc: "Paese"})
                fig.update_layout(hovermode="x unified", margin={"t": 40}, height=580)
                graphs.append(dcc.Graph(figure=fig, style={"height": "580px"}))

        # ── Snapshot anno ──────────────────────────────────────────────────
        elif cmp_mode == "snapshot":
            ei_unit = _EI_UNITS.get(key, "Valore") if meta.get("type") == "ei" else None
            year = cmp_year or int(sub_base[yc].max())
            metrics_plot = [cmp_metric_a] + (cmp_metrics_extra or [])

            if len(metrics_plot) == 1:
                s = _snap(sub_base, metrics_plot[0], year)
                graphs.extend(_data_warnings(s, metrics_plot[0]))
                lbl0 = _mlabel(metrics_plot[0])
                y_label = ei_unit if ei_unit else lbl0
                if not s.empty:
                    fig = px.bar(data_frame=s, x=cc, y="value", color=cc,
                                 title=f"{lbl0}  –  {year}",
                                 labels={"value": y_label, cc: "Paese"})
                    fig.update_layout(showlegend=False, margin={"t": 40}, height=580)
                    graphs.append(dcc.Graph(figure=fig, style={"height": "580px"}))
            else:
                frames = []
                for m in metrics_plot:
                    s = _snap(sub_base, m, year)
                    graphs.extend(_data_warnings(s, m))
                    if not s.empty:
                        frames.append(s.assign(metrica=_mlabel(m)))
                if frames:
                    combined = pd.concat(frames, ignore_index=True)
                    fig = px.bar(data_frame=combined, x=cc, y="value",
                                 color="metrica", barmode="group",
                                 title=f"Confronto metriche  –  {year}",
                                 labels={"value": ei_unit or "Valore", cc: "Paese",
                                         "metrica": "Metrica"})
                    fig.update_layout(margin={"t": 40}, height=540)
                    graphs.append(dcc.Graph(figure=fig, style={"height": "540px"}))

        # ── Grafico a torta ───────────────────────────────────────────────
        elif cmp_mode == "pie":
            ei_unit = _EI_UNITS.get(key, "Valore") if meta.get("type") == "ei" else None
            year = cmp_year or int(sub_base[yc].max())
            metrics_plot = [cmp_metric_a] + (cmp_metrics_extra or [])

            if len(metrics_plot) == 1:
                # Una sola metrica → un'unica torta con i paesi come fette
                s = _snap(sub_base, cmp_metric_a, year)
                graphs.extend(_data_warnings(s, cmp_metric_a))
                lbl0 = _mlabel(cmp_metric_a)
                y_label = ei_unit if ei_unit else lbl0
                if not s.empty:
                    # Raggruppa le voci oltre le prime 11 in "Altri" per leggibilità
                    if len(s) > 12:
                        top = s.head(11)
                        altri_val = s.iloc[11:]["value"].sum()
                        s_pie = pd.concat(
                            [top, pd.DataFrame({cc: ["Altri"], "value": [altri_val]})],
                            ignore_index=True,
                        )
                    else:
                        s_pie = s
                    fig = px.pie(data_frame=s_pie, names=cc, values="value",
                                 title=f"{lbl0}  –  {year}",
                                 labels={"value": y_label, cc: "Paese"})
                    unit_suffix = f" {ei_unit}" if ei_unit else ""
                    fig.update_traces(textinfo="label+percent",
                                      hovertemplate="%{label}: %{value:,.2f}"
                                                    + unit_suffix + "<extra></extra>")
                    fig.update_layout(margin={"t": 40}, height=560)
                    graphs.append(dcc.Graph(figure=fig, style={"height": "560px"}))
            else:
                # Metriche multiple → una torta per paese con la composizione delle metriche
                def _country_metric_value(metric, country):
                    s = _get_series(sub_base, metric)
                    s = s[(s[cc] == country) & (s[yc] == year)]
                    return float(s["value"].iloc[0]) if not s.empty else None

                col_width = 12 if len(cmp_cc) == 1 else 6
                pie_cols = []
                for country in cmp_cc[:6]:
                    rows = [{"metrica": _mlabel(m), "value": v}
                            for m in metrics_plot
                            if (v := _country_metric_value(m, country)) is not None]
                    if not rows:
                        pie_cols.append(dbc.Col(dbc.Alert(
                            [html.Strong(f"{country}: "),
                             "nessun dato per le metriche selezionate."],
                            color="warning", className="py-2 small"), width=col_width))
                        continue
                    dfc = pd.DataFrame(rows)
                    fig = px.pie(data_frame=dfc, names="metrica", values="value",
                                 title=f"{country}  –  {year}")
                    fig.update_traces(textinfo="label+percent",
                                      hovertemplate="%{label}: %{value:,.2f}<extra></extra>")
                    fig.update_layout(margin={"t": 40}, height=460)
                    pie_cols.append(dbc.Col(
                        dcc.Graph(figure=fig, style={"height": "460px"}), width=col_width))
                if len(cmp_cc) > 6:
                    graphs.append(dbc.Alert(
                        f"Mostrati i primi 6 paesi su {len(cmp_cc)} selezionati.",
                        color="info", className="py-2 small"))
                graphs.append(dbc.Row(pie_cols, className="g-3"))

        # ── Scatter X vs Y ─────────────────────────────────────────────────
        elif cmp_mode == "scatter":
            if not cmp_metric_b:
                return dbc.Alert("Seleziona la metrica Y per lo scatter.", color="warning")
            year = cmp_year or int(sub_base[yc].max())
            sx = _snap(sub_base, cmp_metric_a, year).rename(columns={"value": "x"})
            sy = _snap(sub_base, cmp_metric_b, year).rename(columns={"value": "y"})
            graphs.extend(_data_warnings(sx.rename(columns={"x": "value"}), cmp_metric_a))
            graphs.extend(_data_warnings(sy.rename(columns={"y": "value"}), cmp_metric_b))
            if sx.empty or sy.empty:
                return html.Div(graphs)
            merged = sx.merge(sy, on=cc)
            if merged.empty:
                return dbc.Alert("Nessun dato in comune tra le due metriche per l'anno selezionato.",
                                 color="warning")
            fig = px.scatter(data_frame=merged, x="x", y="y", color=cc, text=cc,
                             title=f"{_mlabel(cmp_metric_a)}  vs  {_mlabel(cmp_metric_b)}  –  {year}",
                             labels={"x": _mlabel(cmp_metric_a), "y": _mlabel(cmp_metric_b),
                                     cc: "Paese"})
            fig.update_traces(textposition="top center", marker_size=10)
            fig.update_layout(margin={"t": 40}, height=580)
            graphs.append(dcc.Graph(figure=fig, style={"height": "580px"}))

        # ── Radar multi-metrica ────────────────────────────────────────────
        elif cmp_mode == "radar":
            import plotly.graph_objects as go
            metrics_plot = [cmp_metric_a] + (cmp_metrics_extra or [])
            if len(metrics_plot) < 3:
                return dbc.Alert("Seleziona almeno 3 metriche per il radar.", color="warning")
            year = cmp_year or int(sub_base[yc].max())

            radar_data: dict[str, list] = {}
            missing_radar = []
            for m in metrics_plot:
                s = _snap(sub_base, m, year).set_index(cc)["value"]
                if s.empty:
                    missing_radar.append(m)
                    continue
                for country, val in s.items():
                    radar_data.setdefault(country, {})[m] = val

            # Aggiorna metrics_plot escludendo le metriche senza dati
            metrics_plot = [m for m in metrics_plot if m not in missing_radar]
            if missing_radar:
                names = ", ".join(_mlabel(m) for m in missing_radar)
                graphs.append(dbc.Alert(
                    [html.Strong("Esclusi dal radar (nessun dato): "), names],
                    color="warning", className="py-2 small", dismissable=True,
                ))
            if len(metrics_plot) < 3:
                graphs.append(dbc.Alert(
                    "Dati insufficienti: servono almeno 3 metriche disponibili per il radar.",
                    color="warning"))
                return html.Div(graphs)

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
                theta   = [_mlabel(m)[:35] for m in metrics_plot] + [_mlabel(metrics_plot[0])[:35]]
                fig.add_trace(go.Scatterpolar(r=r_vals, theta=theta,
                                              fill="toself", name=country))
            fig.update_layout(
                polar={"radialaxis": {"visible": True, "range": [0, 1]}},
                title=f"Radar  –  {year}  (valori normalizzati 0–1)",
                margin={"t": 60},
                height=520,
            )
            graphs.append(dcc.Graph(figure=fig))

        # ── Tabella dati (periodicità) ──────────────────────────────────────
        elif cmp_mode == "table":
            period = int(cmp_period or 1)
            years_sel = _period_years(sub_base[yc], period)
            if not years_sel:
                return dbc.Alert("Nessun anno disponibile con i filtri selezionati.",
                                 color="warning")
            metrics_plot = [cmp_metric_a] + (cmp_metrics_extra or [])
            period_label = {1: "annuale", 2: "biennale", 5: "quinquennale"}.get(period, "")
            multi_country = len(cmp_cc) > 1

            combined = pd.DataFrame(index=years_sel)
            missing_notes = []
            for m in metrics_plot:
                s = _get_series(sub_base, m)
                s = s[s[yc].isin(years_sel)]
                if s.empty:
                    missing_notes.append(_mlabel(m))
                    continue
                lbl = _mlabel(m)
                for country in cmp_cc:
                    sc = s[s[cc] == country][[yc, "value"]].dropna()
                    if sc.empty:
                        continue
                    col_name = f"{country} — {lbl}" if multi_country else lbl
                    combined[col_name] = sc.set_index(yc)["value"]

            if missing_notes:
                graphs.append(dbc.Alert(
                    [html.Strong("Nessun dato per: "), ", ".join(missing_notes)],
                    color="warning", className="py-2 small", dismissable=True))

            if combined.empty or combined.dropna(how="all").empty:
                graphs.append(dbc.Alert("Nessun dato disponibile con i parametri selezionati.",
                                        color="warning"))
                return html.Div(graphs)

            combined = (combined.round(2).reset_index()
                        .rename(columns={"index": "Anno"}))
            graphs.append(html.H6(f"Tabella dati — vista {period_label}",
                                  className="mt-3 mb-2 fw-semibold"))
            n_cols = len(combined.columns)
            graphs.append(html.Div(className="mb-4", children=[
                dash_table.DataTable(
                    data=combined.to_dict("records"),
                    columns=[{"name": c, "id": c} for c in combined.columns],
                    style_table={"overflowX": "auto", "width": "100%",
                                 "tableLayout": "fixed" if n_cols > 3 else "auto"},
                    style_cell={"textAlign": "center", "padding": "6px",
                                "fontSize": "0.78rem",
                                "minWidth": "80px", "maxWidth": "180px",
                                "whiteSpace": "normal", "overflow": "hidden",
                                "textOverflow": "ellipsis"},
                    style_header={"fontWeight": "bold", "backgroundColor": "#f8fafc",
                                 "whiteSpace": "normal", "height": "auto",
                                 "lineHeight": "1.2"},
                    style_cell_conditional=[
                        {"if": {"column_id": "Anno"}, "minWidth": "70px",
                         "maxWidth": "70px", "fontWeight": "bold"},
                    ],
                    page_size=15,
                    export_format="csv",
                ),
            ]))

        return html.Div(graphs) if graphs else dbc.Alert(
            "Nessun dato disponibile con i parametri selezionati.", color="warning")

    # ── Statistiche ────────────────────────────────────────────────────────────
    if tab == "tab-stats":
        # ── Estrai serie per la metrica selezionata nel pannello statistiche ───
        if not stats_metric:
            return dbc.Alert("Seleziona una metrica nella sidebar.", color="info")
        if fmt == "wide":
            if stats_metric not in filtered.columns:
                return dbc.Alert("Metrica non disponibile nei dati correnti.", color="warning")
            serie = filtered[[cc, yc, stats_metric]].dropna().rename(columns={stats_metric: "value"})
            m_label = OWID_LABELS.get(stats_metric, stats_metric)
        else:
            if ic and ic in filtered.columns:
                serie = filtered[filtered[ic] == stats_metric][[cc, yc, "value"]].dropna()
            else:
                serie = filtered[[cc, yc, "value"]].dropna()
            if meta.get("type") == "ei":
                unit = _EI_UNITS.get(key, "")
                m_label = f"{stats_metric} ({unit})" if unit and unit != stats_metric else str(stats_metric)
            else:
                m_label = str(stats_metric)

        if serie.empty:
            return dbc.Alert("Nessun dato per la metrica selezionata con i filtri correnti.", color="warning")

        latest_year = int(serie[yc].max())
        earliest_year = int(serie[yc].min())
        ref_year = max(earliest_year, latest_year - 10)

        # ── 1. Ranking: ultimo valore disponibile per ciascun paese ───────────
        # Prende il valore dell'anno più recente disponibile per ogni paese
        last_per_country = (serie.sort_values(yc)
                            .groupby(cc)
                            .last()
                            .reset_index()[[cc, yc, "value"]])
        rank_df = (last_per_country
                   .sort_values("value", ascending=False)
                   .reset_index(drop=True))
        rank_df.index += 1
        rank_df.index.name = "Rank"
        rank_df = rank_df.reset_index()
        rank_df.columns = ["Rank", "Paese", "Anno dato", "Valore"]
        rank_df["Valore"] = rank_df["Valore"].round(2)

        # ── 2. Variazione % rispetto a ~10 anni fa ────────────────────────────
        val_now = last_per_country.groupby(cc)["value"].mean()
        val_ref = serie[serie[yc] == ref_year].groupby(cc)["value"].mean()
        delta = ((val_now - val_ref) / val_ref.abs() * 100).dropna()
        var_col = "Var % vs " + str(ref_year)
        rank_df[var_col] = rank_df["Paese"].map(delta).round(1)
        rank_df["_delta"] = rank_df[var_col]
        rank_df["Tendenza"] = rank_df[var_col].apply(
            lambda v: "▲" if (isinstance(v, float) and not pd.isna(v) and v > 0)
                      else ("▼" if (isinstance(v, float) and not pd.isna(v)) else "–")
        )

        # ── 3. Bar chart di tutti i paesi nel ranking ─────────────────────────
        n_paesi = len(rank_df)
        bar_h = max(460, n_paesi * 26)
        fig_bar = px.bar(
            rank_df, x="Valore", y="Paese", orientation="h",
            color="Valore", color_continuous_scale="Blues",
            title=f"{m_label} – ultimo dato disponibile",
            labels={"Valore": m_label, "Paese": ""},
        )
        fig_bar.update_layout(showlegend=False, margin={"t": 50},
                              coloraxis_showscale=False,
                              yaxis={"autorange": "reversed"}, height=bar_h)

        # ── 4. Evoluzione temporale (media aggregata dei paesi filtrati) ───────
        trend = serie.groupby(yc)["value"].agg(["mean", "min", "max"]).reset_index()
        fig_trend = px.line(trend, x=yc, y="mean",
                            title=f"Evoluzione media – {m_label}",
                            labels={yc: "Anno", "mean": m_label})
        fig_trend.add_scatter(x=trend[yc], y=trend["max"], mode="lines",
                              line={"dash": "dot", "color": "lightblue"},
                              name="Max", showlegend=True)
        fig_trend.add_scatter(x=trend[yc], y=trend["min"], mode="lines",
                              line={"dash": "dot", "color": "salmon"},
                              name="Min", showlegend=True)
        fig_trend.update_layout(hovermode="x unified", margin={"t": 50}, height=460)

        return html.Div([
            html.H5(f"Ranking paesi — {m_label}", className="mb-2"),
            dash_table.DataTable(
                columns=[{"name": c, "id": c} for c in rank_df.columns
                         if c != "_delta"],
                data=rank_df.to_dict("records"),
                page_size=len(rank_df),
                sort_action="native",
                style_cell={"fontSize": "0.82rem", "padding": "4px 10px"},
                style_header={"fontWeight": "bold", "backgroundColor": "#f8f9fa"},
                style_table={"overflowX": "auto"},
                style_data_conditional=[
                    {"if": {"filter_query": "{_delta} > 0",
                            "column_id": "Tendenza"},
                     "color": "#16a34a", "fontWeight": "bold"},
                    {"if": {"filter_query": "{_delta} < 0",
                            "column_id": "Tendenza"},
                     "color": "#dc2626", "fontWeight": "bold"},
                ],
            ),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_bar,  style={"height": f"{bar_h}px"}), width=6),
                dbc.Col(dcc.Graph(figure=fig_trend, style={"height": "460px"}), width=6),
            ], className="mt-4 g-3"),
        ])

    return html.Div()


# 5b. Sincronizza opzioni "Paesi da confrontare" con il filtro globale
@app.callback(
    Output("cmp-countries", "options", allow_duplicate=True),
    Output("cmp-countries", "value",   allow_duplicate=True),
    Input("filter-countries", "value"),
    State("store-key",  "data"),
    State("store-meta", "data"),
    prevent_initial_call=True,
)
def sync_cmp_countries(filter_countries, key, meta):
    if not key or key not in _CACHE or not meta:
        return no_update, no_update
    df  = _CACHE[key]
    cc  = meta["country_col"]
    pool = filter_countries if filter_countries else sorted(df[cc].dropna().unique())
    opts = [{"label": c, "value": c} for c in pool]
    return opts, []


# 6a. Mostra/nascondi bottone download
@app.callback(
    Output("download-btn-wrap", "style"),
    Input("main-tabs",  "value"),
    Input("store-key",  "data"),
)
def toggle_download_button(tab, key):
    if tab == "tab-table" and key and key in _CACHE:
        return {"display": "block"}
    return {"display": "none"}


# 6b. Download CSV
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


# 7. Mostra spinner + pulsante annulla quando parte il caricamento
app.clientside_callback(
    """
    function(n) {
        var show = {"display": "block"};
        var showFixed = {"display": "block", "position": "fixed",
                         "top": "50%", "left": "12.5%",
                         "transform": "translate(-50%, -50%)", "zIndex": 1050};
        return [showFixed, show];
    }
    """,
    Output("sidebar-spinner", "style"),
    Output("btn-cancel",      "style"),
    Input("btn-load", "n_clicks"),
    prevent_initial_call=True,
)

# Nasconde spinner + pulsante annulla quando arriva l'alert (caricamento terminato/interrotto)
app.clientside_callback(
    """
    function(alert) {
        var hide = {"display": "none"};
        return [hide, hide];
    }
    """,
    Output("sidebar-spinner", "style", allow_duplicate=True),
    Output("btn-cancel",      "style", allow_duplicate=True),
    Input("alert-area", "children"),
    prevent_initial_call=True,
)


# 8. Pulsante annulla: segnala l'interruzione al thread di caricamento
@app.callback(
    Output("alert-area", "children", allow_duplicate=True),
    Input("btn-cancel", "n_clicks"),
    prevent_initial_call=True,
)
def cancel_load(_):
    _cancel_event.set()
    return dbc.Alert("⛔  Caricamento interrotto dall'utente.",
                     color="secondary", duration=4000, dismissable=True)


if __name__ == "__main__":
    # host="0.0.0.0" rende l'app raggiungibile da altri dispositivi sulla stessa rete Wi-Fi
    app.run(debug=True, host="0.0.0.0", port=8050)
