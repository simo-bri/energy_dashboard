# Energy Data Analysis Platform

Web app interattiva per l'analisi di dati energetici globali (OWID, Energy Institute, World Bank WDI).

## Requisiti

- Python 3.10+
- Connessione internet (per scaricare dati OWID e WDI al primo avvio)

## Installazione

```bash
# 1. Clona o decomprimi il progetto
cd energy_dashboard

# 2. Crea un ambiente virtuale (consigliato)
python3 -m venv venv
source venv/bin/activate        # macOS/Linux
venv\Scripts\activate           # Windows

# 3. Installa le dipendenze
pip install -r requirements.txt

# 4. Avvia l'app
python3 energy_dashboard.py
```

Apri il browser su: **http://localhost:8050**

## Dati Energy Institute (opzionale)

Per usare la fonte "Energy Institute – Stats Review" scarica il file Excel ufficiale:

1. Vai su https://www.energyinst.org/statistical-review
2. Scarica `EI-Stats-Review-ALL-data.xlsx`
3. Mettilo nella cartella `energy_dashboard/` (stessa cartella di `energy_dashboard.py`)

Le altre due fonti (OWID e World Bank WDI) si scaricano automaticamente.

## Struttura cartella

```
energy_dashboard/
├── energy_dashboard.py      # App principale
├── requirements.txt         # Dipendenze Python
├── assets/
│   ├── style.css            # Stile UI
│   ├── globe-loader.js      # Animazione globo
│   ├── globe-mount.js       # Integrazione Web Component
│   ├── manifest.json        # PWA manifest
│   ├── sw.js                # Service worker
│   ├── icon-192.png         # Icona app
│   └── icon-512.png         # Icona app (alta risoluzione)
└── worldbank_wdi_bulk/      # Metadati indicatori WDI
```
