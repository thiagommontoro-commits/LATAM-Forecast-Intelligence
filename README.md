# Forecast Lab — Scenario, Accuracy & Drivers Dashboard (AGCO LATAM)

Interactive dashboard (IMR layout) with every Forecast Lab model: volumes vs. previous years,
accuracy and ranking **per forecast year** (closing the current year, N+1, N+2), the Blend,
the track record of each model, and the factors (sales history vs. external drivers) behind each forecast.

## Repository contents
| File | Purpose |
|---|---|
| `gerar_dashboard.py` | Picks the most recent Forecast Lab workbook, reads the official rankings and writes `index.html` (layout embedded) |
| `index.html` | The published dashboard (served by GitHub Pages) |
| `atualizar_dashboard.bat` | One-click update on Windows: rebuild + commit + push |
| `requirements.txt` | Python dependencies |
| `.gitignore` | Keeps Excel files, temp files and caches out of the repository |

## Where the workbook comes from
The script always uses the **most recently modified** `Forecast_Lab*.xlsx` in:

```
C:\Users\tm75667\OneDrive - AGCO Corp\Área de Trabalho\AI Projects\Project Analytics\output_forecast_lab
```
(change `INPUT_DIR` at the top of `gerar_dashboard.py` if the folder moves, or use `--input-dir` / `--file`).

## Update every cycle
1. Run the Forecast Lab — the new workbook lands in `output_forecast_lab`.
2. Double-click **`atualizar_dashboard.bat`** (or run `python gerar_dashboard.py` and push).

First time only: `pip install -r requirements.txt`, then GitHub → **Settings → Pages → Branch `main` / `/ (root)` → Save**.

## How the numbers are built
- Rankings for each forecast year are read from the workbook (`Calc_Metrics`, `Calc_Metrics_NextYear`,
  `Calc_Metrics_YearAfter`). If a workbook comes without cached values, they are recalculated in Python
  from the raw backtest with the README weights.
- Blend: current year = winner of the year; following years = rank 1 if it leads rank 2 by at least the
  README threshold (WMAPE), otherwise the median of the top-3 — same rule as the workbook.
- Consolidated (Product | Market) rankings aggregate the segment backtests with the same weights.
- Years are detected automatically (e.g. next cycle 2027–2029 needs no code change).

## Dashboard tabs
- **Overview by Product & Market** — volumes vs. previous years, accuracy & rank per forecast year, track record, scenario fan, accuracy map.
- **Segment Deep Dive** — the same comparisons per segment + how the Blend was built + monthly path.
- **Factors & Drivers** — sales history vs. external drivers by model, top drivers, effect direction, drivers × models, by segment.
- **Models Guide** — every model in plain words, with pros, cons and results.
- **Methodology & Glossary** — backtest, ranking, Blend rule and its out-of-sample validation, glossary.

Developed by Global Reporting & Analytics — AGCO (Thiago Montoro).
