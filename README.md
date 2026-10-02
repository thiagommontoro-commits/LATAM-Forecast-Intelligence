# Forecast Lab — Scenario Dashboard (AGCO LATAM)

Interactive dashboard (IMR layout) showing 2024–2027 scenarios from every Forecast Lab model,
with ranking, backtest accuracy and variable importance by segment.

## Repository contents
| File | Purpose |
|---|---|
| `gerar_dashboard.py` | Reads the Forecast Lab workbook, recalculates all metrics and writes `index.html` (HTML template embedded) |
| `index.html` | The published dashboard (GitHub Pages serves this file) |
| `requirements.txt` | Python dependencies |
| `.gitignore` | Excludes Excel temp files and Python caches |
| `../Project Analytics/output_forecast_lab/Forecast_Lab_Ciclo_*.xlsx` | Forecast Lab output for the current cycle (default input) |

## Update every cycle
```bash
pip install -r requirements.txt     # first time only
python gerar_dashboard.py           # uses the most recent Excel in ../Project Analytics/output_forecast_lab
# or: python gerar_dashboard.py path/file.xlsx --out docs/index.html
git add . && git commit -m "Cycle 10+2" && git push
```

Automatic selection in `../Project Analytics/output_forecast_lab` uses the filename
`Forecast_Lab_Ciclo_<month>+<remaining>_<YYYYMMDD>_<HHMM>.xlsx`: newest year first,
then highest cycle (`11+1` after `10+2`, after `9+3`), then latest run timestamp.
File modification time breaks ties and is the fallback if no filename matches this pattern.
Excel temporary files (`~$`) are ignored; an explicit file path overrides automatic selection.

## Publish (first time)
GitHub → **Settings → Pages → Branch `main` / folder `/ (root)` → Save**. The dashboard URL is ready in 1–2 minutes.

## What the script does
- Reads the raw tabs (`Data_Monthly`, `Data_Backtest`, `Data_Annual_Backtest`, `Data_Importance`, `README`).
- **Recalculates in Python** every metric (WMAPE, annual WMAPE, MAPE, bias, R², MAE, RMSE), the weighted
  ranking using the README-tab weights, the Blend and the Product|Market roll-ups (in Excel these are
  formulas and arrive empty when read by Python).
- Recalculates `ENSEMBLE_MEDIANA` (median of the other models) for forecast months.
- Writes a self-contained `index.html` (data embedded; Chart.js loaded from CDN).

## Dashboard tabs
- **Consolidated View** — KPIs, annual volumes 2024–2027 by model with accuracy (total or by segment), scenario fan, segment × model accuracy map.
- **Segment Deep Dive** — monthly path, 2026/2027 volumes by model, accuracy vs. bias, full ranking, variable importance (with +1 SD effect of each driver).
- **Models Guide** — conceptual explanation of all 12 models with pros, cons and when each tends to win.
- **Methodology & Criteria** — backtest, metrics, ranking weights, Blend, agronomic filter, watch-outs.

To customize the layout without editing Python, save a `template_dashboard.html` next to the script (it overrides the embedded template).

Developed by Global Reporting & Analytics — AGCO (Thiago Montoro).
