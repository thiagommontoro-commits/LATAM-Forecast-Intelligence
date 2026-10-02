# -*- coding: utf-8 -*-
"""
AGCO Forecast Lab — Dashboard generator (index.html)
----------------------------------------------------
Reads the Forecast Lab workbook (e.g. Forecast_Lab_Ciclo_9+3_*.xlsx), RECALCULATES in Python
all backtest metrics, the weighted ranking per segment (weights from the README tab), the Blend,
the Product|Market roll-ups and variable importances, and writes a self-contained
index.html (ready for GitHub Pages).

Usage:
  python gerar_dashboard.py                   # uses the most recent Excel in Project Analytics/output_forecast_lab
    python gerar_dashboard.py path/file.xlsx    # uses a specific file
    python gerar_dashboard.py file.xlsx --out docs/index.html

Dependencies: pandas, numpy, openpyxl
Author: Global Reporting & Analytics — AGCO (Thiago Montoro)
"""
import sys, os, re, glob, json, argparse
from datetime import datetime
import numpy as np
import pandas as pd

# --------------------------------------------------------------------------------------
# Rótulos
# --------------------------------------------------------------------------------------
PROD = {"TA": "Tractors", "CO": "Combines", "PA": "Planters", "PU": "Sprayers"}
PROD_ORDER = ["TA", "CO", "PA", "PU"]
REG = {"BRA": "Brazil", "ARG": "Argentina", "MEX": "Mexico", "OSA": "OSA"}
REG_ORDER = ["BRA", "ARG", "MEX", "OSA"]
MODEL_INFO = {
    "ETS_HOLT_WINTERS":   ("ETS Holt-Winters", "Time series", False),
    "SARIMA":             ("SARIMA", "Time series", False),
    "THETA":              ("Theta", "Time series", False),
    "STL_ETS":            ("STL + ETS", "Time series", False),
    "GLM_POISSON_SERIE":  ("GLM Poisson (segment)", "GLM", True),
    "GLM_POISSON_GLOBAL": ("GLM Poisson (global)", "GLM", True),
    "GLM_TWEEDIE_GLOBAL": ("GLM Tweedie (global)", "GLM", True),
    "XGBOOST":            ("XGBoost", "Machine learning", True),
    "XGBOOST_SEL":        ("XGBoost (selected)", "Machine learning", True),
    "LIGHTGBM_SEL":       ("LightGBM (selected)", "Machine learning", True),
    "RANDOM_FOREST_SEL":  ("Random Forest (selected)", "Machine learning", True),
    "ENSEMBLE_MEDIANA":   ("Ensemble (median)", "Ensemble", None),
}
TYPE_PT = {"History": "History (lags / moving avg.)", "Seasonality": "Seasonality", "Trend": "Trend", "Driver": "External driver"}
DEFAULT_WEIGHTS = {"wmape": 2, "ann_wmape": 2, "mape": 1, "abs_mbe": 1, "r2": 1, "mae": 1, "rmse": 1}
WEIGHT_LABELS = {"WMAPE (monthly)": "wmape", "Annual WMAPE (N and N+1)": "ann_wmape", "MAPE (monthly)": "mape",
                 "|MBE| bias (monthly)": "abs_mbe", "R² (monthly)": "r2", "MAE (monthly)": "mae", "RMSE (monthly)": "rmse"}


def seg_en(name):
    """Display label in English for segment names coming from the database."""
    n = str(name)
    n = n.replace("20 linhas ou +", "20+ rows").replace("0 < 20 linhas", "< 20 rows").replace("linhas", "rows")
    n = n.replace("Classe ", "Class ").replace(" & Over", "+")
    return n


def find_input(arg):
    if arg:
        return arg
    here = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(os.path.dirname(here), "Project Analytics", "output_forecast_lab")
    files = [f for f in glob.glob(os.path.join(input_dir, "*.xlsx")) if not os.path.basename(f).startswith("~$")]
    lab = [f for f in files if "forecast_lab" in os.path.basename(f).lower()] or files
    if not lab:
        sys.exit(f"No .xlsx found in {input_dir}. Pass the path: python gerar_dashboard.py file.xlsx")

    def round_order(path):
      match = re.fullmatch(r"Forecast_Lab_Ciclo_(\d+)\+(\d+)_(\d{8})_(\d{4})\.xlsx", os.path.basename(path), re.IGNORECASE)
      if match:
        cycle_month, remaining_months, date, time = match.groups()
        return (int(date[:4]), int(cycle_month), int(date + time), os.path.getmtime(path))
      return (0, 0, 0, os.path.getmtime(path))

    return max(lab, key=round_order)


def num(x):
    try:
        v = float(x)
        return None if np.isnan(v) else v
    except Exception:
        return None


# --------------------------------------------------------------------------------------
# Leitura
# --------------------------------------------------------------------------------------
def read_readme(xl):
    r = xl.parse("README", header=None)
    weights, tiebreak, run_info, cutoffs, log = dict(DEFAULT_WEIGHTS), 0.001, "", [], []
    for _, row in r.iterrows():
        vals = [v for v in row.tolist() if pd.notna(v)]
        if not vals:
            continue
        a = str(vals[0]).strip()
        if a.startswith("Run "):
            run_info = a
        if a in WEIGHT_LABELS and len(vals) >= 3 and num(vals[2]) is not None:
            weights[WEIGHT_LABELS[a]] = num(vals[2])
        if a.startswith("Tie-break") and len(vals) >= 2 and num(vals[-1]) is not None:
            tiebreak = num(vals[-1])
        if a in PROD and len(vals) >= 4 and str(vals[1]) in REG:
            cutoffs.append({"p": a, "r": str(vals[1]), "last": str(vals[2]), "cycle": str(vals[3])})
        if a == "•" and len(vals) >= 2:
            log.append(str(vals[1]))
    return weights, tiebreak, run_info, cutoffs, log


def read_monthly(xl):
    m = xl.parse("Data_Monthly", header=None)
    ym = m.iloc[0].tolist()
    mcols = [i for i, v in enumerate(ym) if num(v) is not None and 200001 <= num(v) <= 209912]
    months = [int(num(ym[i])) for i in mcols]
    body = m.iloc[3:].reset_index(drop=True)
    data = {}
    last_actual = {}
    for _, row in body.iterrows():
        p, r, seg, model, series = row[0], row[1], row[2], row[3], row[4]
        if pd.isna(series) or pd.isna(model):
            continue
        vals = [num(row[i]) for i in mcols]
        data.setdefault(series, {"p": p, "r": r, "s": seg, "models": {}})
        data[series]["models"][model] = vals
        if num(row[6]) is not None:
            last_actual[series] = int(num(row[6]))
    return months, data, last_actual


# --------------------------------------------------------------------------------------
# Cálculos
# --------------------------------------------------------------------------------------
def fill_ensemble(months, data, last_actual):
    """ENSEMBLE_MEDIANA no Excel é fórmula (MEDIAN) -> recalcula como mediana dos demais modelos."""
    for s, d in data.items():
        mods = d["models"]
        others = [k for k in mods if k != "ENSEMBLE_MEDIANA"]
        if "ENSEMBLE_MEDIANA" not in mods:
            continue
        cut = last_actual.get(s, 0)
        ens = list(mods["ENSEMBLE_MEDIANA"])
        for j, ym in enumerate(months):
            if ym > cut:
                vs = [mods[k][j] for k in others if mods[k][j] is not None]
                ens[j] = float(np.median(vs)) if vs else None
        mods["ENSEMBLE_MEDIANA"] = ens


def backtest_metrics(bt, ann):
    bt = bt[bt["Actual"].notna() & bt["Forecast"].notna()].copy()
    bt["err"] = bt["Forecast"] - bt["Actual"]
    bt["aerr"] = bt["err"].abs()
    out = {}
    for (model, series), g in bt.groupby(["Model", "Series"]):
        a, e = g["Actual"].values.astype(float), g["err"].values.astype(float)
        sa = a.sum()
        pos = a > 0
        sst = ((a - a.mean()) ** 2).sum()
        out[(model, series)] = {
            "wmape": np.abs(e).sum() / sa if sa > 0 else np.nan,
            "mape": (np.abs(e[pos]) / a[pos]).mean() if pos.any() else np.nan,
            "mbe_pct": e.sum() / sa if sa > 0 else np.nan,
            "r2": 1 - (e ** 2).sum() / sst if sst > 0 else np.nan,
            "mae": np.abs(e).mean(),
            "rmse": np.sqrt((e ** 2).mean()),
            "sum_aerr": np.abs(e).sum(), "sum_err": e.sum(), "sum_act": sa,
        }
    ann = ann[ann["Actual for year"].notna() & ann["Actual YTD + forecast"].notna()].copy()
    ann["aerr"] = (ann["Actual YTD + forecast"] - ann["Actual for year"]).abs()
    for (model, series), g in ann.groupby(["Model", "Series"]):
        sa = g["Actual for year"].sum()
        k = (model, series)
        if k in out:
            out[k]["ann_wmape"] = g["aerr"].sum() / sa if sa > 0 else np.nan
            out[k]["sum_ann_aerr"] = g["aerr"].sum()
            out[k]["sum_ann_act"] = sa
    for v in out.values():
        v["acc"] = 1 - v["wmape"]
        v["ann_acc"] = 1 - v.get("ann_wmape", np.nan)
        v["abs_mbe"] = abs(v["mbe_pct"])
    return out


def rank_models(df, weights, tiebreak, crit=None):
    """df: index=model, colunas = critérios. Retorna df com score e rank (1 = melhor)."""
    crit = crit or ["wmape", "ann_wmape", "mape", "abs_mbe", "r2", "mae", "rmse"]
    total_w, score = 0, 0
    for c in crit:
        w = weights.get(c, 0)
        if w == 0 or c not in df:
            continue
        asc = c != "r2"
        rk = df[c].rank(ascending=asc, method="min", na_option="bottom")
        score = score + w * rk
        total_w += w
    df = df.copy()
    df["score"] = score / total_w + tiebreak * df["wmape"].fillna(9)
    df["rank"] = df["score"].rank(method="first").astype(int)
    return df.sort_values("rank")


def year_totals(months, vals):
    t = {}
    for ym, v in zip(months, vals):
        y = ym // 100
        if v is not None:
            t[y] = t.get(y, 0) + v
    return t


def r1(x, n=4):
    return None if x is None or (isinstance(x, float) and np.isnan(x)) else round(float(x), n)


# --------------------------------------------------------------------------------------
def build(path):
    print(f"📂 Reading: {path}")
    xl = pd.ExcelFile(path)
    weights, tiebreak, run_info, cutoffs, log = read_readme(xl)
    months, data, last_actual = read_monthly(xl)
    fill_ensemble(months, data, last_actual)
    bt = xl.parse("Data_Backtest")
    ann = xl.parse("Data_Annual_Backtest")
    imp = xl.parse("Data_Importance")
    met = backtest_metrics(bt, ann)
    origins = sorted(bt["Origin"].astype(str).unique().tolist())

    mcyc = re.search(r"Ciclo[_ ]?(\d+\+\d+)", os.path.basename(path))
    cycle = mcyc.group(1) if mcyc else (cutoffs[0]["cycle"] if cutoffs else "")
    cut = max(last_actual.values()) if last_actual else None
    hist_years = sorted({ym // 100 for ym in months if cut and ym // 100 < cut // 100})
    fc_years = sorted({ym // 100 for ym in months if cut and ym // 100 >= cut // 100})
    models = list(MODEL_INFO.keys())

    # ---------------- importância por série × modelo
    imp = imp.rename(columns={"Importance (error increase after shuffling)": "imp", "Effect of +1 standard deviation": "eff"})
    importance = {}
    for (model, series), g in imp.groupby(["Model", "Series"]):
        g = g[g["imp"].notna()].copy()
        g["imp"] = g["imp"].clip(lower=0)
        tot = g["imp"].sum()
        dtot = g.loc[g["Type"] == "Driver", "imp"].sum()
        rows = []
        for _, r_ in g.sort_values("imp", ascending=False).iterrows():
            rows.append({
                "v": str(r_["Variable"]), "t": str(r_["Type"]),
                "sh": r1(r_["imp"] / tot if tot > 0 else 0),
                "dsh": r1(r_["imp"] / dtot) if (r_["Type"] == "Driver" and dtot > 0) else None,
                "eff": r1(num(r_["eff"])),
            })
        importance.setdefault(series, {})[model] = rows

    # ---------------- segmentos
    segments = {}
    for s, d in data.items():
        cutm = last_actual.get(s, cut)
        any_vals = next(iter(d["models"].values()))
        actual = [v if ym <= cutm else None for ym, v in zip(months, any_vals)]
        rows = []
        for mdl in models:
            if mdl not in d["models"] or (mdl, s) not in met:
                continue
            mm = met[(mdl, s)]
            rows.append(dict(model=mdl, **{k: mm.get(k, np.nan) for k in
                        ["wmape", "ann_wmape", "mape", "abs_mbe", "mbe_pct", "r2", "mae", "rmse", "acc", "ann_acc"]}))
        rdf = rank_models(pd.DataFrame(rows).set_index("model"), weights, tiebreak)
        ranking = []
        for mdl, rr in rdf.iterrows():
            vals = d["models"][mdl]
            yt = year_totals(months, vals)
            ytd = sum(v for ym, v in zip(months, vals) if ym <= cutm and ym // 100 == cutm // 100 and v is not None)
            ranking.append({
                "m": mdl, "rank": int(rr["rank"]), "score": r1(rr["score"], 3),
                "acc": r1(rr["acc"]), "ann_acc": r1(rr["ann_acc"]), "wmape": r1(rr["wmape"]), "mape": r1(rr["mape"]),
                "bias": r1(rr["mbe_pct"]), "r2": r1(rr["r2"]), "mae": r1(rr["mae"], 2), "rmse": r1(rr["rmse"], 2),
                "y": {str(y): r1(v, 2) for y, v in yt.items()}, "ytd": r1(ytd, 2),
                "imp": mdl in importance.get(s, {}),
            })
        segments[s] = {
            "p": d["p"], "r": d["r"], "s": seg_en(d["s"]), "cut": cutm,
            "actual": [r1(v, 2) for v in actual],
            "fc": {m: [r1(v, 3) for v in d["models"][m]] for m in models if m in d["models"]},
            "ranking": ranking, "winner": ranking[0]["m"],
        }

    # ---------------- consolidado Produto|Região (todos os modelos + BLEND)
    consolidated = {}
    combos = sorted({(v["p"], v["r"]) for v in segments.values()},
                    key=lambda x: (PROD_ORDER.index(x[0]) if x[0] in PROD_ORDER else 9, REG_ORDER.index(x[1]) if x[1] in REG_ORDER else 9))
    for p, r in combos:
        sids = [s for s, v in segments.items() if v["p"] == p and v["r"] == r]
        series_ag = {}
        stats = {}
        for mdl in models + ["BLEND"]:
            vec = np.zeros(len(months))
            sa_e = s_err = s_act = s_aae = s_aact = 0.0
            ok = True
            for s in sids:
                use = segments[s]["winner"] if mdl == "BLEND" else mdl
                if use not in segments[s]["fc"]:
                    ok = False
                    break
                vec += np.array([v or 0 for v in segments[s]["fc"][use]])
                mm = met[(use, s)]
                sa_e += mm["sum_aerr"]; s_err += mm["sum_err"]; s_act += mm["sum_act"]
                s_aae += mm.get("sum_ann_aerr", 0); s_aact += mm.get("sum_ann_act", 0)
            if not ok:
                continue
            series_ag[mdl] = vec
            stats[mdl] = {"wmape": sa_e / s_act if s_act else np.nan, "ann_wmape": s_aae / s_aact if s_aact else np.nan,
                          "mbe_pct": s_err / s_act if s_act else np.nan}
            stats[mdl]["abs_mbe"] = abs(stats[mdl]["mbe_pct"])
        sdf = pd.DataFrame(stats).T
        ranked = rank_models(sdf.drop(index="BLEND"), {"wmape": 2, "ann_wmape": 2, "abs_mbe": 1}, tiebreak,
                             crit=["wmape", "ann_wmape", "abs_mbe"])
        cutm = max(segments[s]["cut"] for s in sids)
        actual = np.zeros(len(months))
        for s in sids:
            actual += np.array([v or 0 for v in segments[s]["actual"]])
        actual_l = [r1(v, 2) if ym <= cutm else None for ym, v in zip(months, actual)]
        blend_y = year_totals(months, list(series_ag["BLEND"]))
        rows = []
        order = list(ranked.index) + ["BLEND"]
        for mdl in order:
            st = stats[mdl]
            yt = year_totals(months, list(series_ag[mdl]))
            won = sum(1 for s in sids if segments[s]["winner"] == mdl) if mdl != "BLEND" else len(sids)
            top5 = sum(1 for s in sids for x in segments[s]["ranking"] if x["m"] == mdl and x["rank"] <= 5) if mdl != "BLEND" else None
            rows.append({
                "m": mdl, "rank": int(ranked.loc[mdl, "rank"]) if mdl != "BLEND" else None,
                "acc": r1(1 - st["wmape"]), "ann_acc": r1(1 - st["ann_wmape"]), "bias": r1(st["mbe_pct"]),
                "y": {str(y): r1(v, 1) for y, v in yt.items()}, "won": won, "top5": top5,
                "dblend": r1((yt.get(fc_years[0], 0) / blend_y.get(fc_years[0], 1) - 1) if fc_years else None),
            })
        seg_rows = {}
        for sid in sids:
            seg_rows[sid] = [{"m": x["m"], "rank": x["rank"], "acc": x["acc"], "ann_acc": x["ann_acc"],
                              "bias": x["bias"], "y": {k: (round(v, 1) if v is not None else None) for k, v in x["y"].items()}}
                             for x in segments[sid]["ranking"]]
        consolidated[f"{p}|{r}"] = {
            "seg_rows": seg_rows,
            "p": p, "r": r, "segs": sids, "cut": cutm, "actual": actual_l,
            "fc": {m: [r1(v, 2) for v in series_ag[m]] for m in series_ag},
            "ranking": rows,
        }

    # ---------------- visão global (validação vs. aba Explanations)
    glob_acc = {}
    for mdl in models:
        e = sum(met[(mdl, s)]["sum_aerr"] for s in segments if (mdl, s) in met)
        a = sum(met[(mdl, s)]["sum_act"] for s in segments if (mdl, s) in met)
        glob_acc[mdl] = 1 - e / a
    e = sum(met[(segments[s]["winner"], s)]["sum_aerr"] for s in segments)
    a = sum(met[(segments[s]["winner"], s)]["sum_act"] for s in segments)
    be = sum(met[(segments[s]["winner"], s)]["sum_err"] for s in segments)
    best_single = max(glob_acc, key=glob_acc.get)
    fam_wins = {}
    for s in segments.values():
        f = MODEL_INFO[s["winner"]][1]
        fam_wins[f] = fam_wins.get(f, 0) + 1
    low_acc = sorted([(s, v["ranking"][0]["acc"]) for s, v in segments.items()], key=lambda x: x[1])
    summary = {
        "blend_acc": r1(1 - e / a), "blend_bias": r1(be / a), "best_single": best_single,
        "best_single_acc": r1(glob_acc[best_single]), "glob_acc": {k: r1(v) for k, v in glob_acc.items()},
        "fam_wins": fam_wins, "n_series": len(segments), "n_models": len(models),
        "low_acc": [[f'{PROD.get(segments[s]["p"], segments[s]["p"])} · {REG.get(segments[s]["r"], segments[s]["r"])} · {segments[s]["s"]}', acc]
                    for s, acc in low_acc if acc is not None and acc < 0.70],
    }
    print(f"✅ Blend monthly accuracy {summary['blend_acc']:.1%} | best single model: {best_single} {glob_acc[best_single]:.1%}")

    meta = {
        "cycle": cycle, "cut": cut, "run": run_info, "file": os.path.basename(path),
        "generated": datetime.now().strftime("%d/%m/%Y %H:%M"), "months": months,
        "hist_years": hist_years, "fc_years": fc_years, "origins": origins,
        "weights": weights, "tiebreak": tiebreak, "log": log, "cutoffs": cutoffs,
        "prod": PROD, "reg": REG, "prod_order": PROD_ORDER, "reg_order": REG_ORDER,
        "models": {k: {"n": v[0], "f": v[1], "d": v[2]} for k, v in MODEL_INFO.items()},
        "type_pt": TYPE_PT,
    }
    return {"meta": meta, "summary": summary, "segments": segments,
            "consolidated": consolidated, "importance": importance}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", nargs="?")
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    path = find_input(a.xlsx)
    payload = build(path)
    here = os.path.dirname(os.path.abspath(__file__))
    tpl_path = os.path.join(here, "template_dashboard.html")
    if os.path.exists(tpl_path):          # optional: customize the layout without touching Python
        with open(tpl_path, encoding="utf-8") as f:
            tpl = f.read()
    else:
        tpl = TEMPLATE
    js = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), allow_nan=False, default=lambda o: None)
    out = a.out or os.path.join(here, "index.html")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(tpl.replace("/*__DATA__*/null", js))
    print(f"🚀 Dashboard written: {out}  ({os.path.getsize(out)/1024:.0f} KB)")


# ======================================================================================
# HTML TEMPLATE (IMR layout) — embedded so the script works on its own
# ======================================================================================
TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Forecast Lab | AGCO LATAM — Model Scenarios & Accuracy</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script>if(typeof Chart==="undefined"){document.write('<script src="https://unpkg.com/chart.js@4.4.1/dist/chart.umd.min.js"><\/script>');}</script>
<style>
:root{--red:#C8102E;--red-dark:#8B0A1D;--red-soft:#FDE8EB;--ink:#0E1116;--ink-2:#2B303B;--muted:#6B7280;--bg:#EEF1F5;--card:#FFF;--line:#E6E9EF;
--up:#0B7A3B;--up-bg:#E4F7EC;--down:#C8102E;--down-bg:#FDE8EB;--flat:#8A93A2;--gold:#B8860B;--gold-bg:#FBF3DF;--teal:#0E7C86;--teal-bg:#E1F1F2;--purple:#7A3FB0;--purple-bg:#F3EEFA;
--shadow:0 1px 2px rgba(16,17,22,.04),0 6px 20px -8px rgba(16,17,22,.12);--shadow-lg:0 20px 45px -20px rgba(16,17,22,.28);--font:"Segoe UI",system-ui,-apple-system,Roboto,Arial,sans-serif}
*{margin:0;padding:0;box-sizing:border-box;font-family:var(--font)}
body{background:radial-gradient(1200px 500px at 100% -10%,rgba(200,16,46,.06),transparent 60%),var(--bg);color:var(--ink-2);padding:30px 22px 70px;line-height:1.5}
.container{max-width:1640px;margin:0 auto}
.hero{position:relative;overflow:hidden;border-radius:22px;padding:30px 36px;background:linear-gradient(135deg,#141821,#1d222d 55%,#241016);color:#fff;box-shadow:var(--shadow-lg);margin-bottom:22px}
.hero::after{content:"";position:absolute;right:-60px;top:-80px;width:340px;height:340px;background:radial-gradient(circle,rgba(200,16,46,.55),transparent 62%);filter:blur(6px)}
.hero-top{display:flex;justify-content:space-between;align-items:flex-start;gap:20px;position:relative;z-index:2;flex-wrap:wrap}
.brand-row{display:flex;align-items:center;gap:16px}
.cy-badge{display:flex;flex-direction:column;align-items:center;justify-content:center;min-width:82px;height:66px;border-radius:14px;background:linear-gradient(145deg,var(--red),var(--red-dark));box-shadow:0 8px 20px -6px rgba(200,16,46,.7);padding:0 12px}
.cy-badge b{font-size:1.55rem;font-weight:800;color:#fff;line-height:1}
.cy-badge span{font-size:.52rem;color:#FDE8EB;letter-spacing:.8px;margin-top:4px}
.hero h1{font-size:1.5rem;font-weight:800;letter-spacing:-.5px}
.hero .sub{color:#B7BECC;margin-top:4px;font-size:.92rem}
.hero .pill{display:inline-flex;align-items:center;gap:8px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.16);padding:9px 15px;border-radius:999px;font-size:.82rem;font-weight:600;color:#EAEDF3}
.pill .dot{width:8px;height:8px;border-radius:50%;background:#34d17f;box-shadow:0 0 0 4px rgba(52,209,127,.2)}
.hero-meta{display:flex;gap:30px;margin-top:22px;position:relative;z-index:2;flex-wrap:wrap}
.hero-meta .k{font-size:.7rem;text-transform:uppercase;letter-spacing:1.2px;color:#8B93A3}
.hero-meta .v{font-size:1.04rem;font-weight:700;color:#fff;margin-top:2px}
.hero-meta .v small{color:#FDE68A;font-weight:700}
.tabs{display:flex;gap:6px;padding:6px;background:linear-gradient(180deg,#1f1f1f,#0b0b0b);border-radius:14px;margin-bottom:16px;box-shadow:0 6px 16px -8px rgba(0,0,0,.45);flex-wrap:wrap}
.tab{padding:11px 22px;border:none;border-radius:10px;background:transparent;color:#C4C4C4;font-weight:700;font-size:.92rem;cursor:pointer;transition:.2s}
.tab:hover{color:#fff;background:rgba(255,255,255,.07)}
.tab.active{background:linear-gradient(135deg,var(--red),var(--red-dark));color:#fff;box-shadow:0 4px 12px rgba(200,16,46,.4)}
.filters{display:flex;gap:18px;margin-bottom:18px;flex-wrap:wrap;align-items:center;background:var(--card);border:1px solid var(--line);border-radius:16px;padding:12px 16px;box-shadow:var(--shadow)}
.fgroup{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.flabel{font-size:.66rem;font-weight:800;letter-spacing:.7px;text-transform:uppercase;color:var(--muted);margin-right:2px}
.fbtn{background:var(--card);color:var(--ink-2);border:1px solid var(--line);padding:7px 15px;border-radius:999px;font-size:.83rem;font-weight:600;cursor:pointer;transition:.18s}
.fbtn:hover{border-color:var(--red);color:var(--red)}
.fbtn.active{background:linear-gradient(135deg,var(--red),var(--red-dark));color:#fff;border-color:transparent}
.fbtn:disabled{opacity:.35;cursor:not-allowed}
.fbtn.sm{padding:5px 12px;font-size:.78rem}
.seg-bar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px}
.seg-btn{background:#fff;border:1px solid var(--line);border-left:4px solid var(--flat);padding:8px 14px;border-radius:10px;font-size:.82rem;font-weight:700;color:var(--ink);cursor:pointer;transition:.15s;text-align:left}
.seg-btn small{display:block;font-weight:600;color:var(--muted);font-size:.68rem}
.seg-btn:hover{box-shadow:var(--shadow)}
.seg-btn.active{background:#141821;color:#fff;border-color:#141821}
.seg-btn.active small{color:#FDE68A}
.kpi-wrap{display:grid;grid-template-columns:repeat(auto-fit,minmax(215px,1fr));gap:16px;margin-bottom:20px}
.kpi{background:var(--card);padding:18px 20px;border-radius:16px;box-shadow:var(--shadow);border:1px solid var(--line);position:relative;overflow:hidden;transition:.2s}
.kpi:hover{transform:translateY(-3px);box-shadow:var(--shadow-lg)}
.kpi::before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:var(--red)}
.kpi.win::before{background:var(--up)}.kpi.loss::before{background:var(--down)}.kpi.gold::before{background:var(--gold)}.kpi.purple::before{background:var(--purple)}.kpi.teal::before{background:var(--teal)}
.kpi h3{font-size:.68rem;color:var(--muted);text-transform:uppercase;letter-spacing:1px;font-weight:700;margin-bottom:6px}
.kpi .val{font-size:1.35rem;font-weight:800;color:var(--ink);line-height:1.15}
.kpi .val small{font-size:.82rem;font-weight:800}
.kpi .desc{margin-top:6px;color:var(--muted);font-size:.78rem}
.up{color:var(--up)}.down{color:var(--down)}.flat{color:var(--flat)}
.panel{background:var(--card);padding:22px 24px;border-radius:18px;box-shadow:var(--shadow);border:1px solid var(--line);margin-bottom:20px}
.ptitle{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px;gap:14px;flex-wrap:wrap}
.ptitle h2{font-size:1.12rem;color:var(--ink);font-weight:800}
.ptitle .hint{font-size:.8rem;color:var(--muted);margin-top:2px}
.ptools{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.exp{background:linear-gradient(135deg,#0B7A3B,#0a6c35);color:#fff;border:none;padding:8px 15px;border-radius:10px;font-size:.8rem;font-weight:700;cursor:pointer}
.grid2{display:grid;grid-template-columns:1.25fr 1fr;gap:20px}
@media(max-width:1100px){.grid2{grid-template-columns:1fr}}
.chartbox{position:relative;height:360px}
.chartbox.tall{height:420px}
.read{background:#F9FAFB;border:1px solid var(--line);border-left:4px solid var(--red);border-radius:10px;padding:14px 16px;font-size:.88rem;line-height:1.65;color:var(--ink-2);margin-bottom:16px}
.read b{color:var(--ink)}
.scroll{overflow-x:auto;border-radius:12px;border:1px solid var(--line)}
table{width:100%;border-collapse:separate;border-spacing:0}
th,td{padding:10px 12px;text-align:right;border-bottom:1px solid var(--line);font-size:.83rem;white-space:nowrap}
th.l,td.l{text-align:left}
thead th{background:#F7F8FA;color:var(--muted);font-weight:700;text-transform:uppercase;font-size:.66rem;letter-spacing:.5px;position:sticky;top:0;z-index:5}
th.grp{background:#F0F2F5;color:var(--ink);text-align:center;border-left:2px solid var(--line)}
th.fc,td.fc{background:#FCF3F4}
thead th.fc{color:var(--red)}
tbody tr:hover td{background:#FBFCFE}
tr.winner td{background:#F2FBF5}
tr.blend td{background:#141821!important;color:#fff;font-weight:800}
tr.blend td .mname{color:#fff}
.mname{font-weight:700;color:var(--ink)}
.fam{display:inline-block;font-size:.62rem;font-weight:800;padding:2px 8px;border-radius:999px;margin-left:6px;text-transform:uppercase;letter-spacing:.3px}
.fam.ts{background:var(--teal-bg);color:var(--teal)}.fam.glm{background:var(--gold-bg);color:var(--gold)}.fam.ml{background:var(--purple-bg);color:var(--purple)}.fam.ens{background:#EEF0F3;color:#4b5563}.fam.bl{background:var(--red);color:#fff}
.rk{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:8px;font-weight:800;font-size:.78rem;background:#F0F2F5;color:var(--ink)}
.rk.r1{background:linear-gradient(135deg,#0B7A3B,#0a6c35);color:#fff}.rk.r2{background:#CDEFD9;color:#0B7A3B}.rk.r3{background:#E4F7EC;color:#0B7A3B}
.accb{display:inline-block;min-width:58px;padding:3px 8px;border-radius:7px;font-weight:800;font-size:.78rem;text-align:center}
.dot-m{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:7px;vertical-align:middle}
.star{color:#0B7A3B;font-weight:900;margin-left:4px}
.heat td{font-size:.76rem;padding:7px 8px}
.heat td.cell{text-align:center;font-weight:700;cursor:pointer;min-width:62px}
.heat td.cell.best{outline:2px solid #0B7A3B;outline-offset:-2px}
.heat td.seg{text-align:left;font-weight:700;color:var(--ink);cursor:pointer;position:sticky;left:0;background:#fff;z-index:2}
.heat td.seg:hover{color:var(--red)}
.heat th.mh{writing-mode:vertical-rl;transform:rotate(180deg);height:185px;text-align:left;padding:8px 6px;font-size:.66rem}
.legend{display:flex;gap:14px;flex-wrap:wrap;margin-top:12px;font-size:.74rem;color:var(--muted);align-items:center}
.legend i{width:11px;height:11px;border-radius:3px;display:inline-block;margin-right:5px;vertical-align:middle}
.imp-models{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px;align-items:center}
.imp-row{display:grid;grid-template-columns:minmax(220px,1.3fr) 2fr 70px 150px;gap:12px;align-items:center;padding:8px 4px;border-bottom:1px dashed var(--line);font-size:.83rem}
.imp-row .vn{font-weight:700;color:var(--ink)}
.imp-row .vt{display:inline-block;font-size:.6rem;font-weight:800;padding:2px 7px;border-radius:999px;margin-left:6px;text-transform:uppercase}
.vt.Driver{background:var(--red-soft);color:var(--red)}.vt.History{background:#EEF0F3;color:#4b5563}.vt.Seasonality{background:var(--teal-bg);color:var(--teal)}.vt.Trend{background:var(--gold-bg);color:var(--gold)}
.bar{height:14px;border-radius:7px;background:#F0F2F5;overflow:hidden}
.bar>div{height:100%;border-radius:7px}
.imp-row .pct{text-align:right;font-weight:800;color:var(--ink)}
.eff{font-size:.78rem;font-weight:700}
.note{font-size:.8rem;color:var(--muted);background:#fff;border:1px dashed #D4D9E1;border-radius:8px;padding:10px 12px;margin-top:12px}
.note b{color:var(--ink-2)}
.dmx td.c{text-align:center;font-weight:700;font-size:.76rem}
.sec{margin-top:4px}
.step{font-size:.72rem;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);margin:22px 0 12px;display:flex;align-items:center;gap:9px}
.step .stepn{width:22px;height:22px;border-radius:50%;color:#fff;font-size:.74rem;font-weight:800;display:flex;align-items:center;justify-content:center;background:var(--red)}
.lead{background:#F9FAFB;border:1px solid var(--line);border-radius:10px;padding:14px 16px;font-size:.88rem;line-height:1.65;border-left:4px solid var(--red)}
.fx{margin-top:10px;font-size:.8rem;color:var(--muted);background:#fff;border:1px dashed #D4D9E1;border-radius:8px;padding:10px 12px}
.fx code{background:#141821;color:#EAEDF3;padding:2px 8px;border-radius:6px;font-size:.78rem}
.vlist{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}
@media(max-width:820px){.vlist{grid-template-columns:1fr}.imp-row{grid-template-columns:1fr}}
.vrow{display:flex;gap:11px;align-items:flex-start;background:#fff;border:1px solid var(--line);border-radius:10px;padding:12px 14px}
.vrow b{font-size:.86rem;color:var(--ink)}
.vrow span{display:block;font-size:.78rem;color:var(--muted);margin-top:2px;line-height:1.45}
.author{margin-top:22px;padding:16px 20px;border-radius:12px;background:linear-gradient(135deg,#141821,#241016);color:#EAEDF3;font-size:.85rem}
.author b{color:#FDE68A}
.foot{text-align:center;color:var(--muted);font-size:.76rem;margin-top:26px;line-height:1.7}
.hidden{display:none!important}
.tag{display:inline-block;padding:3px 10px;border-radius:999px;font-size:.72rem;font-weight:700;margin:2px 4px 2px 0}
.tag.n{background:var(--down-bg);color:var(--down)}.tag.p{background:var(--up-bg);color:var(--up)}

.mcard{background:#fff;border:1px solid var(--line);border-left:6px solid var(--red);border-radius:14px;padding:18px 20px;margin-bottom:14px;box-shadow:var(--shadow)}
.mc-head{display:flex;justify-content:space-between;align-items:flex-start;gap:14px;flex-wrap:wrap;margin-bottom:8px}
.mc-name{font-size:1.08rem;color:var(--ink)}
.mc-stats{display:flex;gap:10px;flex-wrap:wrap}
.mc-stats span{background:#F7F8FA;border:1px solid var(--line);border-radius:10px;padding:6px 10px;font-size:.7rem;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;text-align:center;min-width:92px}
.mc-stats b{display:block;font-size:1rem;color:var(--ink);text-transform:none;letter-spacing:0}
.mc-idea{font-size:.9rem;line-height:1.6;color:var(--ink-2);margin:6px 0}
.mc-analogy{font-size:.85rem;color:#521b25;background:#FEF6F7;border-left:3px solid var(--red);border-radius:8px;padding:8px 12px;margin:8px 0}
.pc{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:12px}
@media(max-width:820px){.pc{grid-template-columns:1fr}}
.pc h4{font-size:.78rem;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
.pc ul{padding-left:18px;font-size:.84rem;line-height:1.6}
.pros{background:var(--up-bg);border-radius:10px;padding:10px 14px}.pros h4{color:var(--up)}
.cons{background:var(--down-bg);border-radius:10px;padding:10px 14px}.cons h4{color:var(--down)}
.mc-wins{margin-top:10px;font-size:.84rem;color:var(--ink-2);background:var(--gold-bg);border-radius:8px;padding:8px 12px}
</style>
</head>
<body>
<div class="container">

  <div class="hero">
    <div class="hero-top">
      <div class="brand-row">
        <div class="cy-badge"><b id="hCycle">9+3</b><span>CYCLE</span></div>
        <div>
          <h1>Forecast Lab — Scenarios, Ranking &amp; Model Accuracy</h1>
          <div class="sub" id="hSub">Statistical models · historical backtest · weighted ranking · explainability by segment</div>
        </div>
      </div>
      <span class="pill"><span class="dot"></span><span id="hPill">—</span></span>
    </div>
    <div class="hero-meta" id="heroMeta"></div>
  </div>

  <div class="tabs">
    <button class="tab active" data-tab="cons">📊 Consolidated View</button>
    <button class="tab" data-tab="seg">🔍 Segment Deep Dive</button>
    <button class="tab" data-tab="mod">🧠 Models Guide</button>
    <button class="tab" data-tab="met">📐 Methodology &amp; Criteria</button>
  </div>

  <div class="filters" id="filters">
    <div class="fgroup"><span class="flabel">Product</span><span id="fProd"></span></div>
    <div class="fgroup"><span class="flabel">Market</span><span id="fReg"></span></div>
    <div class="fgroup"><span class="flabel">Models in chart</span>
      <button class="fbtn sm active" data-view="top5">Top 5 + Blend</button>
      <button class="fbtn sm" data-view="all">All models</button>
    </div>
  </div>

  <!-- ================== CONSOLIDATED ================== -->
  <section id="tab-cons">
    <div class="kpi-wrap" id="cKpis"></div>
    <div class="read" id="cRead"></div>
    <div class="panel">
      <div class="ptitle"><div><h2>Annual volumes by model — 2024 to 2027</h2><div class="hint" id="cRankHint"></div></div>
        <div class="ptools">
          <span class="flabel">Rows</span>
          <button class="fbtn sm active" data-rows="total">Total (all segments)</button>
          <button class="fbtn sm" data-rows="seg">By segment</button>
          <button class="exp" onclick="exportCSV('tCons','annual_volumes')">Export CSV</button>
        </div></div>
      <div class="scroll"><table id="tCons"></table></div>
      <div class="legend"><span><i style="background:#F2FBF5;border:1px solid #CDEFD9"></i>Ranking leader / segment winner</span><span><i style="background:#141821"></i>Blend = sum of each segment's winner</span><span><i style="background:#FCF3F4"></i>Forecast years (2026 = actuals to cutoff + forecast)</span></div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>Annual scenario fan</h2><div class="hint">2024–2025 actuals and 2026–2027 projection by model</div></div></div>
      <div class="chartbox"><canvas id="cAnnual"></canvas></div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>Accuracy map — segment × model</h2><div class="hint">Monthly backtest accuracy (1 − WMAPE). Green outline = segment winner. Click a segment to open the deep dive.</div></div></div>
      <div class="scroll"><table class="heat" id="tHeat"></table></div>
      <div class="legend"><span><i style="background:hsl(0,70%,88%)"></i>&lt; 50%</span><span><i style="background:hsl(40,80%,86%)"></i>50–70%</span><span><i style="background:hsl(90,55%,85%)"></i>70–85%</span><span><i style="background:hsl(135,50%,80%)"></i>&gt; 85%</span></div>
    </div>
  </section>

  <!-- ================== SEGMENT ================== -->
  <section id="tab-seg" class="hidden">
    <div class="seg-bar" id="segBar"></div>
    <div class="kpi-wrap" id="sKpis"></div>
    <div class="read" id="sRead"></div>
    <div class="panel">
      <div class="ptitle"><div><h2 id="sMonthlyT">Monthly path</h2><div class="hint">Jan/2024 – Dec/2027 · black = actuals · red = segment winner</div></div></div>
      <div class="chartbox tall"><canvas id="sMonthly"></canvas></div>
    </div>
    <div class="grid2">
      <div class="panel">
        <div class="ptitle"><div><h2>2026 and 2027 volumes by model</h2><div class="hint">Sorted by ranking · dashed line = 2025 actuals</div></div></div>
        <div class="chartbox"><canvas id="sAnnual"></canvas></div>
      </div>
      <div class="panel">
        <div class="ptitle"><div><h2>Accuracy vs. bias (backtest)</h2><div class="hint">Best = top of the chart, close to 0% bias</div></div></div>
        <div class="chartbox"><canvas id="sScatter"></canvas></div>
      </div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>Full segment ranking</h2><div class="hint">Score = weighted average of the ranks on each criterion (weights in the Methodology tab). Lowest score wins.</div></div>
        <div class="ptools"><button class="exp" onclick="exportCSV('tSeg','segment_ranking')">Export CSV</button></div></div>
      <div class="scroll"><table id="tSeg"></table></div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>What drives the model? — Variable importance</h2><div class="hint">Permutation importance: how much the error increases when the variable is shuffled. Bars = share of total explained.</div></div></div>
      <div class="imp-models" id="impModels"></div>
      <div id="impNote"></div>
      <div id="impList"></div>
      <div class="legend"><span><i style="background:var(--red)"></i>External driver</span><span><i style="background:#8A93A2"></i>History (lags / moving avg.)</span><span><i style="background:var(--teal)"></i>Seasonality</span><span><i style="background:var(--gold)"></i>Trend</span><span>▲/▼ = effect of +1 standard deviation of the driver on the forecast</span></div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>Drivers × models in this segment</h2><div class="hint">Driver weight among each model's drivers (%). Consistency across models = more robust evidence.</div></div></div>
      <div class="scroll"><table class="dmx" id="tDrv"></table></div>
    </div>
  </section>

  <!-- ================== MODELS GUIDE ================== -->
  <section id="tab-mod" class="hidden">
    <div class="panel" id="modPanel"></div>
  </section>

  <!-- ================== METHODOLOGY ================== -->
  <section id="tab-met" class="hidden">
    <div class="panel" id="metPanel"></div>
  </section>

  <div class="foot" id="foot"></div>
</div>
<script>
const D = /*__DATA__*/null;
(function(){
const M=D.meta, S=D.segments, C=D.consolidated, IMP=D.importance, SUM=D.summary;
const MODELS=Object.keys(M.models);
const COLORS={ETS_HOLT_WINTERS:"#0E7C86",SARIMA:"#2BA3AD",THETA:"#6FC6CC",STL_ETS:"#0A4F55",GLM_POISSON_SERIE:"#B8860B",GLM_POISSON_GLOBAL:"#E0A82E",GLM_TWEEDIE_GLOBAL:"#7A5806",XGBOOST:"#7A3FB0",XGBOOST_SEL:"#A57BDB",LIGHTGBM_SEL:"#4E2380",RANDOM_FOREST_SEL:"#C9A6F2",ENSEMBLE_MEDIANA:"#8A93A2",BLEND:"#C8102E",ACTUAL:"#0E1116"};
const FAMCLS={"Time series":"ts","GLM":"glm","Machine learning":"ml","Ensemble":"ens"};
const st={tab:"cons",p:null,r:null,seg:null,view:"top5",rows:"total",impModel:null,impMode:"all"};
const charts={};
const $=s=>document.querySelector(s);
const nf0=v=>v==null?"–":Math.round(v).toLocaleString("en-US");
const pct=(v,d=1)=>v==null?"–":(v*100).toLocaleString("en-US",{minimumFractionDigits:d,maximumFractionDigits:d})+"%";
const spct=(v,d=1)=>v==null?"–":(v>0?"+":"")+pct(v,d);
const cls=v=>v==null?"flat":v>0.005?"up":v<-0.005?"down":"flat";
const mn=m=>m==="BLEND"?"Blend (winners)":(M.models[m]?M.models[m].n:m);
const famTag=m=>{if(m==="BLEND")return '<span class="fam bl">Blend</span>';const f=M.models[m].f;return `<span class="fam ${FAMCLS[f]}">${f}</span>`;};
const ymLabel=ym=>{const mm=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][ym%100-1];return mm+"/"+String(Math.floor(ym/100)).slice(2);};
const heat=a=>{if(a==null)return"#F0F2F5";const x=Math.max(0,Math.min(1,(a-0.3)/0.65));return `hsl(${Math.round(x*135)},${55+15*(1-x)}%,${86-6*x}%)`;};
const yoy=(a,b)=>(a!=null&&b)?a/b-1:null;
const HY=M.hist_years.map(String), FY=M.fc_years.map(String), YEARS=HY.concat(FY);
const LASTH=HY[HY.length-1], F1=FY[0], F2=FY[1]||FY[0];
const segSort=(a,b)=>S[a].s.localeCompare(S[b].s,undefined,{numeric:true});

// ---------------- hero
$("#hCycle").textContent=M.cycle;
$("#hPill").textContent=`Cycle ${M.cycle} · last actual ${ymLabel(M.cut)} · generated ${M.generated}`;
$("#hSub").innerHTML=`${SUM.n_models} statistical models · ${SUM.n_series} series · backtest origins ${M.origins.join(", ")} · weighted ranking · explainability by segment`;
$("#heroMeta").innerHTML=[
 ["Coverage",`${SUM.n_series} series · ${SUM.n_models} models`],
 ["Horizon",`${HY[0]}–${LASTH} actuals · ${FY.join("–")} forecast`],
 ["Blend accuracy (backtest)",`${pct(SUM.blend_acc)} <small>bias ${spct(SUM.blend_bias)}</small>`],
 ["Best single model",`${mn(SUM.best_single)} <small>${pct(SUM.best_single_acc)}</small>`],
 ["Winners by family",Object.entries(SUM.fam_wins).map(([k,v])=>`${k} ${v}`).join(" · ")]
].map(([k,v])=>`<div><div class="k">${k}</div><div class="v">${v}</div></div>`).join("");
$("#foot").innerHTML=`<b>Forecast Lab · Cycle ${M.cycle}</b> · AGCO Global Reporting &amp; Analytics · Thiago Montoro<br>Source: ${M.file} · ${M.run}`;

// ---------------- filters
const combos=Object.keys(C);
function fillFilters(){
  const prods=M.prod_order.filter(p=>combos.some(k=>k.startsWith(p+"|")));
  if(!st.p)st.p=prods[0];
  $("#fProd").innerHTML=prods.map(p=>`<button class="fbtn ${p===st.p?"active":""}" data-p="${p}">${M.prod[p]}</button>`).join(" ");
  const regs=M.reg_order.filter(r=>combos.includes(st.p+"|"+r));
  if(!regs.includes(st.r))st.r=regs[0];
  $("#fReg").innerHTML=M.reg_order.map(r=>`<button class="fbtn ${r===st.r?"active":""}" data-r="${r}" ${regs.includes(r)?"":"disabled"}>${M.reg[r]}</button>`).join(" ");
  document.querySelectorAll("[data-p]").forEach(b=>b.onclick=()=>{st.p=b.dataset.p;st.seg=null;render();});
  document.querySelectorAll("[data-r]").forEach(b=>b.onclick=()=>{st.r=b.dataset.r;st.seg=null;render();});
}
document.querySelectorAll("[data-view]").forEach(b=>b.onclick=()=>{st.view=b.dataset.view;document.querySelectorAll("[data-view]").forEach(x=>x.classList.toggle("active",x===b));render();});
document.querySelectorAll("[data-rows]").forEach(b=>b.onclick=()=>{st.rows=b.dataset.rows;document.querySelectorAll("[data-rows]").forEach(x=>x.classList.toggle("active",x===b));render();});
document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>setTab(b.dataset.tab));
function setTab(t){st.tab=t;document.querySelectorAll(".tab").forEach(x=>x.classList.toggle("active",x.dataset.tab===t));
 ["cons","seg","mod","met"].forEach(k=>$("#tab-"+k).classList.toggle("hidden",k!==t));$("#filters").classList.toggle("hidden",t==="met"||t==="mod");render();window.scrollTo({top:0,behavior:"smooth"});}

// ---------------- chart helpers
const cutPlugin={id:"cut",beforeDatasetsDraw(ch,a,o){if(o.idx==null||o.idx<0)return;const{ctx,chartArea:ca,scales:{x}}=ch;const px=x.getPixelForValue(o.idx);ctx.save();ctx.fillStyle="rgba(200,16,46,.05)";ctx.fillRect(px,ca.top,ca.right-px,ca.bottom-ca.top);ctx.strokeStyle="rgba(200,16,46,.55)";ctx.setLineDash([5,4]);ctx.beginPath();ctx.moveTo(px,ca.top);ctx.lineTo(px,ca.bottom);ctx.stroke();ctx.setLineDash([]);ctx.fillStyle="#C8102E";ctx.font="700 11px Segoe UI";ctx.fillText("forecast →",px+6,ca.top+12);ctx.restore();}};
const HASCHART=typeof Chart!=="undefined";
if(HASCHART){Chart.register(cutPlugin);Chart.defaults.font.family='"Segoe UI",system-ui,Arial';Chart.defaults.color="#6B7280";}
function mk(id,cfg){const el=document.getElementById(id);if(!el)return;if(!HASCHART){el.parentNode.innerHTML='<div class="note">Chart unavailable: the Chart.js library did not load (check your connection). All tables remain complete.</div>';return;}
 if(charts[id])charts[id].destroy();charts[id]=new Chart(el,cfg);}
const tipNum={callbacks:{label:c=>`${c.dataset.label}: ${nf0(c.parsed.y)}`}};
function visibleModels(ranking){const ord=ranking.filter(r=>r.m!=="BLEND").sort((a,b)=>a.rank-b.rank).map(r=>r.m);return st.view==="top5"?ord.slice(0,5):ord;}
function fcLine(vals,months,cut){return vals.map((v,i)=>months[i]>=cut?v:null);}

// ---------------- CONSOLIDATED
function volHeader(extraLeft){
  let h=`<thead><tr>${extraLeft}<th class="l">#</th><th class="l">Model</th><th>Monthly accuracy</th><th>Annual accuracy</th><th>Bias</th>`;
  HY.forEach(y=>h+=`<th>${y} actual</th>`);
  FY.forEach(y=>h+=`<th class="fc">${y}</th><th class="fc">YoY ${y.slice(2)}</th>`);
  return h;
}
function volCells(r,isB){
  let h=`<td><span class="accb" style="background:${isB?"transparent":heat(r.acc)}">${pct(r.acc)}</span></td><td>${pct(r.ann_acc)}</td><td class="${isB?"":r.bias>0?"down":"up"}">${spct(r.bias)}</td>`;
  HY.forEach(y=>h+=`<td>${nf0(r.y[y])}</td>`);
  FY.forEach((y,i)=>{const pv=i===0?r.y[LASTH]:r.y[FY[i-1]];const g=yoy(r.y[y],pv);h+=`<td class="fc"><b>${nf0(r.y[y])}</b></td><td class="fc ${isB?"":cls(g)}">${spct(g)}</td>`;});
  return h;
}
function renderCons(){
  const key=st.p+"|"+st.r, c=C[key]; if(!c)return;
  const rk=c.ranking, bl=rk.find(r=>r.m==="BLEND"), lead=rk.find(r=>r.rank===1);
  const models=rk.filter(r=>r.m!=="BLEND");
  const v2=models.map(r=>r.y[F2]).filter(v=>v!=null);
  const mnV=Math.min(...v2), mxV=Math.max(...v2);
  const y1=yoy(bl.y[F1],bl.y[LASTH]), y2=yoy(bl.y[F2],bl.y[F1]);
  $("#cKpis").innerHTML=`
   <div class="kpi win"><h3>Leading model (consolidated)</h3><div class="val">${mn(lead.m)}</div><div class="desc">Accuracy ${pct(lead.acc)} monthly · ${pct(lead.ann_acc)} annual</div></div>
   <div class="kpi"><h3>Blend ${F1}</h3><div class="val">${nf0(bl.y[F1])} <small class="${cls(y1)}">${spct(y1)}</small></div><div class="desc">vs. ${LASTH} actuals: ${nf0(bl.y[LASTH])}</div></div>
   <div class="kpi"><h3>Blend ${F2}</h3><div class="val">${nf0(bl.y[F2])} <small class="${cls(y2)}">${spct(y2)}</small></div><div class="desc">vs. ${F1}</div></div>
   <div class="kpi gold"><h3>${F2} scenario range</h3><div class="val">${nf0(mnV)} – ${nf0(mxV)}</div><div class="desc">Spread of ±${pct((mxV-mnV)/2/bl.y[F2],1)} around the Blend</div></div>
   <div class="kpi teal"><h3>Blend accuracy (backtest)</h3><div class="val">${pct(bl.acc)}</div><div class="desc">Annual ${pct(bl.ann_acc)} · bias <span class="${bl.bias>0?"down":"up"}">${spct(bl.bias)}</span></div></div>`;
  const lowSegs=c.segs.map(s=>[s,S[s].ranking[0].acc]).filter(x=>x[1]<0.7);
  $("#cRead").innerHTML=`<b>Quick read — ${M.prod[st.p]} · ${M.reg[st.r]}:</b> the <b>Blend</b> (winning model of each of the ${c.segs.length} segments) projects <b>${nf0(bl.y[F1])}</b> units in ${F1} (<span class="${cls(y1)}">${spct(y1)}</span> vs. ${LASTH}) and <b>${nf0(bl.y[F2])}</b> in ${F2} (<span class="${cls(y2)}">${spct(y2)}</span>).
   Across the ${models.length} models, the ${F2} scenario ranges from <b>${nf0(mnV)}</b> to <b>${nf0(mxV)}</b>. The consolidated leader is <b>${mn(lead.m)}</b> (${pct(lead.acc)} monthly backtest accuracy).
   ${lowSegs.length?`<br>⚠ Segments with winner accuracy below 70% (review with Sales): ${lowSegs.map(x=>`<span class="tag n">${S[x[0]].s} ${pct(x[1],0)}</span>`).join("")}`:`<br>✅ All segments have winner accuracy ≥ 70%.`}`;

  // volume table
  if(st.rows==="total"){
    $("#cRankHint").innerHTML=`Volumes summed across all ${c.segs.length} segments of ${M.prod[st.p]} · ${M.reg[st.r]}. Consolidated ranking = weighted average of ranks on monthly WMAPE (2), annual WMAPE (2) and |bias| (1).`;
    let h=volHeader("")+`<th>Segments won</th><th>Δ vs Blend ${F1}</th></tr></thead><tbody>`;
    rk.forEach(r=>{const isB=r.m==="BLEND";
      h+=`<tr class="${isB?"blend":r.rank===1?"winner":""}"><td class="l">${isB?"★":`<span class="rk r${r.rank}">${r.rank}</span>`}</td><td class="l"><span class="dot-m" style="background:${COLORS[r.m]}"></span><span class="mname">${mn(r.m)}</span>${famTag(r.m)}</td>`+volCells(r,isB)+
      `<td>${r.won}</td><td class="${isB?"":cls(r.dblend)}">${isB?"–":spct(r.dblend)}</td></tr>`;});
    $("#tCons").innerHTML=h+"</tbody>";
  } else {
    $("#cRankHint").innerHTML=`One block per segment of ${M.prod[st.p]} · ${M.reg[st.r]}, models sorted by the segment ranking (rank 1 = winner, used in the Blend). Showing ${st.view==="top5"?"the Top 5 models per segment — switch to “All models” in the filter bar to see all":"all models"}.`;
    let h=volHeader('<th class="l">Segment</th>')+`</tr></thead><tbody>`;
    c.segs.slice().sort(segSort).forEach(s=>{
      let rows=c.seg_rows[s].slice().sort((a,b)=>a.rank-b.rank);
      if(st.view==="top5")rows=rows.slice(0,5);
      rows.forEach((r,i)=>{h+=`<tr class="${r.rank===1?"winner":""}">${i===0?`<td class="l" rowspan="${rows.length}" style="font-weight:800;color:var(--ink);background:#F7F8FA;vertical-align:top;border-right:2px solid var(--line)">${S[s].s}</td>`:""}<td class="l"><span class="rk r${r.rank}">${r.rank}</span></td><td class="l"><span class="dot-m" style="background:${COLORS[r.m]}"></span><span class="mname">${mn(r.m)}</span>${r.rank===1?'<span class="star">★</span>':""}${famTag(r.m)}</td>`+volCells(r,false)+`</tr>`;});
    });
    const blr=rk.find(r=>r.m==="BLEND");
    h+=`<tr class="blend"><td class="l">TOTAL</td><td class="l">★</td><td class="l"><span class="mname">Blend (sum of winners)</span></td>`+volCells(blr,true)+`</tr>`;
    $("#tCons").innerHTML=h+"</tbody>";
  }

  // annual fan
  const vis=visibleModels(rk);
  const ds=vis.map(m=>{const r=rk.find(x=>x.m===m);return{label:mn(m),data:YEARS.map(y=>r.y[y]),borderColor:COLORS[m],backgroundColor:COLORS[m],borderWidth:1.6,pointRadius:3,tension:.15};});
  ds.unshift({label:"Blend (winners)",data:YEARS.map(y=>bl.y[y]),borderColor:COLORS.BLEND,backgroundColor:COLORS.BLEND,borderWidth:4,pointRadius:5,tension:.15});
  mk("cAnnual",{type:"line",data:{labels:YEARS,datasets:ds},options:{maintainAspectRatio:false,interaction:{mode:"index",intersect:false},plugins:{cut:{idx:HY.length-1},legend:{position:"bottom",labels:{boxWidth:10,font:{size:10}}},tooltip:tipNum},scales:{y:{ticks:{callback:v=>nf0(v)}}}}});

  // heat
  const segs=c.segs.slice().sort(segSort);
  let hh=`<thead><tr><th class="l">Segment</th><th class="l">Winner</th>`+MODELS.map(m=>`<th class="mh">${mn(m)}</th>`).join("")+`</tr></thead><tbody>`;
  segs.forEach(s=>{const sg=S[s];hh+=`<tr><td class="seg" data-goto="${s}">${sg.s} ↗</td><td class="l"><span class="dot-m" style="background:${COLORS[sg.winner]}"></span>${mn(sg.winner)}</td>`;
   MODELS.forEach(m=>{const r=sg.ranking.find(x=>x.m===m);hh+=r?`<td class="cell ${r.rank===1?"best":""}" style="background:${heat(r.acc)}" title="${mn(m)} · rank ${r.rank} · accuracy ${pct(r.acc)}" data-goto="${s}">${pct(r.acc,0)}<br><small style="color:#6B7280">#${r.rank}</small></td>`:`<td class="cell">–</td>`;});hh+="</tr>";});
  $("#tHeat").innerHTML=hh+"</tbody>";
  document.querySelectorAll("#tHeat [data-goto]").forEach(el=>el.onclick=()=>{st.seg=el.dataset.goto;setTab("seg");});
}

// ---------------- SEGMENT
function renderSeg(){
  const key=st.p+"|"+st.r, c=C[key]; if(!c)return;
  const segs=c.segs.slice().sort(segSort);
  if(!segs.includes(st.seg)){st.seg=segs[0];st.impModel=null;}
  $("#segBar").innerHTML=segs.map(s=>{const w=S[s].ranking[0];return `<button class="seg-btn ${s===st.seg?"active":""}" data-s="${s}" style="border-left-color:${heat(w.acc)}">${S[s].s}<small>${mn(w.m)} · ${pct(w.acc,0)}</small></button>`;}).join("");
  document.querySelectorAll("[data-s]").forEach(b=>b.onclick=()=>{st.seg=b.dataset.s;st.impModel=null;renderSeg();});
  const sg=S[st.seg], rk=sg.ranking, w=rk[0];
  const y1=yoy(w.y[F1],w.y[LASTH]), y2=yoy(w.y[F2],w.y[F1]);
  const v2=rk.map(r=>r.y[F2]).filter(v=>v!=null);
  $("#sKpis").innerHTML=`
   <div class="kpi win"><h3>Winner (rank 1)</h3><div class="val">${mn(w.m)}</div><div class="desc">${M.models[w.m].f} · score ${w.score.toLocaleString("en-US")}</div></div>
   <div class="kpi teal"><h3>Accuracy (backtest)</h3><div class="val">${pct(w.acc)}</div><div class="desc">Annual ${pct(w.ann_acc)} · R² ${w.r2==null?"–":w.r2.toLocaleString("en-US",{maximumFractionDigits:2})}</div></div>
   <div class="kpi"><h3>${F1} total</h3><div class="val">${nf0(w.y[F1])} <small class="${cls(y1)}">${spct(y1)}</small></div><div class="desc">Actuals to ${ymLabel(sg.cut)}: ${nf0(w.ytd)} + forecast</div></div>
   <div class="kpi"><h3>${F2} total</h3><div class="val">${nf0(w.y[F2])} <small class="${cls(y2)}">${spct(y2)}</small></div><div class="desc">Scenarios: ${nf0(Math.min(...v2))} – ${nf0(Math.max(...v2))}</div></div>
   <div class="kpi ${Math.abs(w.bias)>0.1?"loss":"gold"}"><h3>Winner bias</h3><div class="val ${w.bias>0?"down":"up"}">${spct(w.bias)}</div><div class="desc">${w.bias>0?"Tends to over-forecast":"Tends to under-forecast"} in backtest</div></div>`;
  const second=rk[1];
  $("#sRead").innerHTML=`<b>Why did ${mn(w.m)} win in ${sg.s}?</b> It had the lowest weighted score (${w.score.toLocaleString("en-US")}) — monthly accuracy ${pct(w.acc)} vs. ${pct(second.acc)} for the runner-up (${mn(second.m)}), annual accuracy ${pct(w.ann_acc)} and bias ${spct(w.bias)}.
   ${w.acc<0.7?`<br>⚠ <b>Accuracy below 70%</b>: validate this scenario with Sales before the business review.`:""}
   ${Math.abs(w.bias)>0.15?`<br>⚠ Material bias (${spct(w.bias)}): consider a ${w.bias>0?"downward":"upward"} adjustment when reading this scenario.`:""}`;
  $("#sMonthlyT").textContent=`Monthly path — ${M.prod[sg.p]} · ${M.reg[sg.r]} · ${sg.s}`;
  const lab=M.months.map(ymLabel), ci=M.months.indexOf(sg.cut), vis=visibleModels(rk);
  const ds=[{label:"Actuals",data:sg.actual,borderColor:COLORS.ACTUAL,borderWidth:2.4,pointRadius:0},
    {label:`★ ${mn(w.m)} (winner)`,data:fcLine(sg.fc[w.m],M.months,sg.cut),borderColor:COLORS.BLEND,borderWidth:3.2,pointRadius:0}];
  vis.filter(m=>m!==w.m).forEach(m=>ds.push({label:mn(m),data:fcLine(sg.fc[m],M.months,sg.cut),borderColor:COLORS[m],borderWidth:1.3,pointRadius:0,borderDash:[4,3]}));
  mk("sMonthly",{type:"line",data:{labels:lab,datasets:ds},options:{maintainAspectRatio:false,interaction:{mode:"index",intersect:false},plugins:{cut:{idx:ci},legend:{position:"bottom",labels:{boxWidth:10,font:{size:10}}},tooltip:tipNum},scales:{x:{ticks:{maxTicksLimit:18,font:{size:10}}},y:{ticks:{callback:v=>nf0(v)}}}}});
  const ord=rk.slice(), ref=w.y[LASTH];
  mk("sAnnual",{type:"bar",data:{labels:ord.map(r=>(r.rank===1?"★ ":"")+mn(r.m)),datasets:[
    {label:F1,data:ord.map(r=>r.y[F1]),backgroundColor:ord.map(r=>COLORS[r.m]+"99"),borderColor:ord.map(r=>COLORS[r.m]),borderWidth:1},
    {label:F2,data:ord.map(r=>r.y[F2]),backgroundColor:ord.map(r=>COLORS[r.m]),borderWidth:0},
    {type:"line",label:`${LASTH} actuals`,data:ord.map(()=>ref),borderColor:"#0E1116",borderDash:[6,4],borderWidth:1.5,pointRadius:0}]},
    options:{maintainAspectRatio:false,plugins:{cut:{idx:null},legend:{position:"bottom",labels:{boxWidth:10}},tooltip:tipNum},scales:{x:{ticks:{font:{size:9},maxRotation:60,minRotation:45}},y:{ticks:{callback:v=>nf0(v)}}}}});
  mk("sScatter",{type:"bubble",data:{datasets:rk.map(r=>({label:mn(r.m),data:[{x:r.bias*100,y:r.acc*100,r:r.rank===1?12:Math.max(4,10-r.rank*0.5)}],backgroundColor:COLORS[r.m]+(r.rank===1?"":"B0"),borderColor:r.rank===1?"#0B7A3B":"#fff",borderWidth:r.rank===1?3:1}))},
    options:{maintainAspectRatio:false,plugins:{cut:{idx:null},legend:{position:"bottom",labels:{boxWidth:10,font:{size:10}}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: accuracy ${c.raw.y.toFixed(1)}% · bias ${c.raw.x>0?"+":""}${c.raw.x.toFixed(1)}%`}}},
    scales:{x:{title:{display:true,text:"Bias (MBE %) — + over-forecast / − under-forecast"},ticks:{callback:v=>v+"%"}},y:{title:{display:true,text:"Monthly accuracy (%)"},ticks:{callback:v=>v+"%"}}}}});
  let h=`<thead><tr><th class="l" rowspan="2">#</th><th class="l" rowspan="2">Model</th><th rowspan="2">Score</th><th class="grp" colspan="7">Backtest accuracy (${M.origins.length} origins)</th><th class="grp" colspan="${HY.length+1+FY.length*2}">Volumes (units)</th></tr><tr>
   <th>Monthly acc.</th><th>Annual acc.</th><th>WMAPE</th><th>MAPE</th><th>Bias</th><th>R²</th><th>MAE</th>`;
  HY.forEach(y=>h+=`<th>${y} actual</th>`);h+=`<th>${F1} YTD</th>`;FY.forEach(y=>h+=`<th class="fc">${y}</th><th class="fc">YoY</th>`);h+=`</tr></thead><tbody>`;
  rk.forEach(r=>{h+=`<tr class="${r.rank===1?"winner":""}"><td class="l"><span class="rk r${r.rank}">${r.rank}</span></td><td class="l"><span class="dot-m" style="background:${COLORS[r.m]}"></span><span class="mname">${mn(r.m)}</span>${r.rank===1?'<span class="star">★</span>':""}${famTag(r.m)}</td>
   <td>${r.score.toLocaleString("en-US")}</td><td><span class="accb" style="background:${heat(r.acc)}">${pct(r.acc)}</span></td><td>${pct(r.ann_acc)}</td><td>${pct(r.wmape)}</td><td>${pct(r.mape)}</td>
   <td class="${r.bias>0?"down":"up"}">${spct(r.bias)}</td><td>${r.r2==null?"–":r.r2.toLocaleString("en-US",{maximumFractionDigits:2})}</td><td>${nf0(r.mae)}</td>`;
   HY.forEach(y=>h+=`<td>${nf0(r.y[y])}</td>`);h+=`<td>${nf0(r.ytd)}</td>`;
   FY.forEach((y,i)=>{const pv=i===0?r.y[LASTH]:r.y[FY[i-1]];const g=yoy(r.y[y],pv);h+=`<td class="fc"><b>${nf0(r.y[y])}</b></td><td class="fc ${cls(g)}">${spct(g)}</td>`;});h+="</tr>";});
  $("#tSeg").innerHTML=h+"</tbody>";
  renderImp(sg);
}

function renderImp(sg){
  const imp=IMP[st.seg]||{}, rk=sg.ranking;
  const withImp=rk.filter(r=>imp[r.m]);
  if(!withImp.length){$("#impModels").innerHTML="";$("#impNote").innerHTML="";$("#impList").innerHTML=`<div class="note">No importance data for this segment.</div>`;$("#tDrv").innerHTML="";return;}
  const winnerHas=!!imp[rk[0].m];
  if(!st.impModel||!imp[st.impModel])st.impModel=winnerHas?rk[0].m:withImp[0].m;
  $("#impModels").innerHTML=withImp.map(r=>`<button class="fbtn sm ${r.m===st.impModel?"active":""}" data-im="${r.m}">#${r.rank} ${mn(r.m)}${r.rank===1?" ★":""}</button>`).join("")+
   `<span style="flex:1"></span><span class="flabel">Show</span><button class="fbtn sm ${st.impMode==="all"?"active":""}" data-imode="all">All variables</button><button class="fbtn sm ${st.impMode==="drv"?"active":""}" data-imode="drv">External drivers only</button>`;
  document.querySelectorAll("[data-im]").forEach(b=>b.onclick=()=>{st.impModel=b.dataset.im;renderImp(sg);});
  document.querySelectorAll("[data-imode]").forEach(b=>b.onclick=()=>{st.impMode=b.dataset.imode;renderImp(sg);});
  $("#impNote").innerHTML=winnerHas?"":`<div class="note">ℹ The winner <b>${mn(rk[0].m)}</b> is a time-series model and <b>does not use external drivers</b> (it explains volume only through its own history, trend and seasonality). As a reference, below is the best-ranked model that uses drivers: <b>${mn(withImp[0].m)}</b> (#${withImp[0].rank}).</div>`;
  const drvOnly=st.impMode==="drv";
  const rows=drvOnly?imp[st.impModel].filter(r=>r.t==="Driver").map(r=>Object.assign({},r,{sh:r.dsh})):imp[st.impModel];
  const TC={Driver:"var(--red)",History:"#8A93A2",Seasonality:"var(--teal)",Trend:"var(--gold)"};
  const maxSh=Math.max(...rows.map(r=>r.sh||0))||1;
  const drvShareAll=imp[st.impModel].filter(r=>r.t==="Driver").reduce((a,r)=>a+(r.sh||0),0);
  let h=`<div class="note" style="margin:0 0 10px"><b>${mn(st.impModel)}:</b> external drivers explain <b>${pct(drvShareAll)}</b> of total importance; the rest comes from history / seasonality / trend.${drvOnly?" Below, each driver's weight <b>among drivers only</b> (sums to 100%).":""}</div>`;
  if(!rows.length)h+=`<div class="note">This model did not select any external driver in this segment.</div>`;
  rows.forEach(r=>{const e=r.eff;const et=e==null?'<span class="flat">—</span>':`<span class="eff ${e>0?"up":"down"}">${e>0?"▲":"▼"} ${spct(e,1)} </span><small class="flat">per +1 SD</small>`;
    h+=`<div class="imp-row"><div><span class="vn">${r.v}</span><span class="vt ${r.t}">${M.type_pt[r.t]||r.t}</span></div><div class="bar"><div style="width:${(r.sh||0)/maxSh*100}%;background:${TC[r.t]||"#999"}"></div></div><div class="pct">${pct(r.sh)}</div><div>${et}</div></div>`;});
  $("#impList").innerHTML=h;
  const drv={};withImp.forEach(r=>imp[r.m].filter(x=>x.t==="Driver").forEach(x=>{drv[x.v]=drv[x.v]||{};drv[x.v][r.m]=x.dsh;}));
  const mx=n=>Math.max(...Object.values(drv[n]).map(v=>v||0));
  const allNames=Object.keys(drv).sort((a,b)=>Object.keys(drv[b]).length-Object.keys(drv[a]).length||mx(b)-mx(a));
  const names=allNames.slice(0,15), extra=allNames.length-names.length;
  if(!names.length){$("#tDrv").innerHTML=`<tbody><tr><td class="l">No external driver was selected by the models in this segment (agronomic filter + selection).</td></tr></tbody>`;return;}
  let t=`<thead><tr><th class="l">Driver</th><th># models</th>`+withImp.map(r=>`<th>#${r.rank} ${mn(r.m)}</th>`).join("")+`</tr></thead><tbody>`;
  names.forEach(n=>{t+=`<tr><td class="l"><b>${n}</b></td><td>${Object.keys(drv[n]).length}</td>`+withImp.map(r=>{const v=drv[n][r.m];return v==null?`<td class="c" style="color:#c3c8d0">–</td>`:`<td class="c" style="background:rgba(200,16,46,${0.08+0.6*v})">${pct(v,0)}</td>`;}).join("")+"</tr>";});
  if(extra>0)t+=`<tr><td class="l" colspan="${withImp.length+2}" style="color:#6B7280;font-style:italic">+ ${extra} more drivers with marginal weight (mostly in XGBoost, which uses every allowed driver).</td></tr>`;
  $("#tDrv").innerHTML=t+"</tbody>";
}

// ---------------- MODELS GUIDE
const GUIDE={
 ETS_HOLT_WINTERS:{idea:"Exponential smoothing that continuously updates three components — level, (damped) trend and a monthly seasonal factor — giving more weight to recent months.",
  analogy:"Like a moving average that “remembers” the seasonal shape of the year and slowly adapts to the current pace of sales.",
  eq:"level(t) = α(y − s) + (1 − α)(level + φ·trend) · trend(t) = β·Δlevel + (1 − β)·φ·trend · s(t) = γ(y − level) + (1 − γ)·s(t − 12)",
  params:"α, β, γ (smoothing) and φ (trend damping) estimated automatically; log scale.",
  pros:["Robust, industry standard for demand forecasting","Captures seasonality and recent momentum well","Fast and easy to explain"],
  cons:["Uses no external information (credit, prices, climate)","Reacts late to structural breaks","Can amplify recent shocks if the last months are atypical"],
  wins:"Long, regular series with stable seasonality."},
 SARIMA:{idea:"Seasonal ARIMA “airline” model: works on month-over-month and year-over-year differences and corrects with recent forecast errors.",
  analogy:"Asks: “how much does this month differ from last month and from the same month last year?” and learns that pattern.",
  eq:"(1 − B)(1 − B¹²) log(1 + y) = (1 + θB)(1 + ΘB¹²) ε",
  params:"Fixed order (0,1,1)(0,1,1)₁₂; θ and Θ estimated by maximum likelihood.",
  pros:["Classic, well-established statistical method","Strong when patterns are stable","Statistically grounded confidence intervals"],
  cons:["Sensitive to structural breaks and outliers","Needs long history (falls back to recent average on short series)","No external drivers"],
  wins:"Long, stable series with clear year-over-year patterns."},
 THETA:{idea:"Splits the deseasonalized series into a long-term trend line and short-term smoothing, then averages both and restores seasonality.",
  analogy:"Half “where is the long-term trend going?”, half “what happened recently?”.",
  eq:"ŷ = ½ · linear trend + ½ · exponential smoothing  (× seasonal factor)",
  params:"Seasonal period 12.",
  pros:["Very robust on short or noisy series","Simple and hard to break","Historically strong in forecasting competitions (M3)"],
  cons:["The trend line can extrapolate too far","Ignores external drivers","Limited ability to model turning points"],
  wins:"Noisy or short series."},
 STL_ETS:{idea:"First decomposes the series into Trend + Seasonality + Remainder with a robust method (STL), forecasts the seasonally adjusted part with ETS and then adds the seasonality back.",
  analogy:"Cleans the “noise” and seasonal effects first, forecasts the clean signal, then re-applies the calendar.",
  eq:"y = Trend + Seasonality + Remainder",
  params:"Period 12, robust STL (down-weights outliers).",
  pros:["Handles outliers and one-off spikes well","Flexible, evolving seasonality","Usually among the most accurate single models"],
  cons:["Needs at least ~3 years of data","No external drivers","Decomposition can be unstable at the end of the series"],
  wins:"Series with occasional spikes or changing seasonality."},
 GLM_POISSON_SERIE:{idea:"A Poisson regression fitted per segment: trend + month + up to 3 drivers that passed the agronomic filter and have the expected sign.",
  analogy:"A transparent “recipe”: each coefficient tells how many % volume changes when a driver moves.",
  eq:"ŷ = MIN(cap ; exp(β₀ + β_T·t + β_month + Σ β_k·driver_k))",
  params:"Tests 0–3 drivers, with/without trend, on the last 12 months; coefficients with the wrong sign are discarded; volume cap.",
  pros:["Highly interpretable (each β is a % effect)","Designed for count data (units)","Business logic enforced (sign consistency)"],
  cons:["Assumes stable multiplicative relationships","Limited history per segment","Few drivers per model"],
  wins:"Segments where 1–3 drivers clearly explain demand."},
 GLM_POISSON_GLOBAL:{idea:"One Poisson GLM trained on all segments at once: history (lags) + month + segment effect + selected drivers per segment; forecasts recursively.",
  analogy:"Learns shared behaviour across segments, so small segments borrow strength from large ones.",
  eq:"log(ŷ) = β₀ + Σ β_lag·log(1 + lag) + β_month + β_segment + Σ β_driver·driver",
  params:"α = 1.0 regularization; future drivers capped to the historical range.",
  pros:["Statistical and interpretable","Pools information across segments","Stable on small segments"],
  cons:["Linear relationships on the log scale","Recursive forecast can propagate errors","Assumes similar dynamics across segments"],
  wins:"Markets where segments move together."},
 GLM_TWEEDIE_GLOBAL:{idea:"Same structure as the global Poisson GLM, but with a Tweedie distribution that tolerates more dispersion and many zeros.",
  analogy:"A Poisson GLM made more “forgiving” for irregular, low-volume segments.",
  eq:"log(ŷ) = same equation as GLM global · Var(y) ∝ μ^1.5",
  params:"Tweedie power 1.5; α = 1.0.",
  pros:["Better for small or irregular segments","Handles zero-sales months","Same interpretability as the GLM"],
  cons:["Same limitations as the GLM (linearity on log scale)","Extra parameter (power) to justify","Less intuitive distribution to explain"],
  wins:"Low-volume or intermittent segments."},
 XGBOOST:{idea:"Gradient boosting: hundreds of small decision trees, each one correcting the error of the previous ones; uses every driver allowed by the agronomic filter and learns across segments.",
  analogy:"Like a golfer: the first shot goes toward the target, each following shot corrects the remaining distance.",
  eq:"ŷ = Σ tree_k(lags, month, segment, allowed drivers)",
  params:"400 trees, depth 5, learning rate 0.04; recursive forecast.",
  pros:["Captures non-linear relationships and interactions","Uses the full set of agronomic drivers","Strong on complex patterns"],
  cons:["Black box — needs importance analysis to explain","Can overfit with limited history","Does not extrapolate beyond the historical range"],
  wins:"Complex relationships between many drivers."},
 XGBOOST_SEL:{idea:"Same XGBoost engine, but each segment keeps only the K drivers that most improve validation accuracy (K = 0, 3, 5, 10 or all).",
  analogy:"XGBoost on a “diet”: keeps only the drivers that really earn their place.",
  eq:"ŷ = Σ tree_k(lags, month, segment, selected drivers)",
  params:"Same as XGBoost + per-segment driver selection.",
  pros:["Less noise than full XGBoost","More focused, easier to explain","Lower overfitting risk"],
  cons:["Selection can change between runs","Still a black box","Relies on a short validation window"],
  wins:"Segments where only a few drivers matter."},
 LIGHTGBM_SEL:{idea:"Alternative gradient-boosting engine that grows trees leaf-wise, with its own driver selection per segment.",
  analogy:"A faster, more aggressive cousin of XGBoost.",
  eq:"ŷ = Σ tree_k(lags, month, segment, selected drivers)",
  params:"400 trees, 15 leaves.",
  pros:["Fast training","Good accuracy on tabular data","Provides a second ML opinion"],
  cons:["Can overfit short series","Black box","Sensitive to hyper-parameters"],
  wins:"Alternative to XGBoost on medium-length series."},
 RANDOM_FOREST_SEL:{idea:"Averages 250 independent decision trees, each trained on a different sample of the data; uses the drivers selected by XGBoost_SEL.",
  analogy:"Asks 250 independent experts and takes the average opinion.",
  eq:"ŷ = average(tree₁ … tree₂₅₀)",
  params:"250 trees; drivers from XGBOOST_SEL.",
  pros:["Very stable, low variance","Robust to noise and outliers","Little tuning needed"],
  cons:["Does not extrapolate trends (flat beyond history)","Tends to under-react to strong growth or decline","Black box"],
  wins:"Noisy series without strong trends."},
 ENSEMBLE_MEDIANA:{idea:"For every month, takes the median of the forecasts of all competing models.",
  analogy:"The “wisdom of the crowd”: individual errors in opposite directions cancel out.",
  eq:"ŷ = MEDIAN(ŷ_model1, …, ŷ_modelN)",
  params:"All eligible models; recalculated in Python (MEDIAN formula in Excel).",
  pros:["Extremely stable and hard to beat on average","Protects against a single model going wrong","No extra assumptions"],
  cons:["Rarely the single best model","Not explainable by drivers","Inherits a shared bias if most models lean the same way"],
  wins:"Segments where no single model stands out."},
};
const FAMILY={
 "Time series":{icon:"📈",txt:"Learn only from the series' own past: level, trend and seasonality. No external information. Strong baselines, very robust, easy to defend — but blind to changes in credit, prices or climate."},
 "GLM":{icon:"📐",txt:"Generalized Linear Models: statistical regressions designed for counts (units sold). Each coefficient is an interpretable % effect of a driver, with business-logic sign checks."},
 "Machine learning":{icon:"🌲",txt:"Tree-based algorithms that capture non-linear relationships and interactions between many drivers. Highest flexibility; explained through permutation importance."},
 "Ensemble":{icon:"🤝",txt:"Combines the other models' forecasts. Not a model of the market itself but a stabilizer that reduces the risk of any single model going wrong."}
};
function renderMod(){
  const segAll=Object.values(S);
  const stat={};MODELS.forEach(m=>{const ranks=segAll.map(s=>s.ranking.find(x=>x.m===m)).filter(Boolean);
    stat[m]={won:ranks.filter(r=>r.rank===1).length,top3:ranks.filter(r=>r.rank<=3).length,avgRank:ranks.reduce((a,r)=>a+r.rank,0)/ranks.length};});
  const fams=["Time series","GLM","Machine learning","Ensemble"];
  let h=`<div class="ptitle"><div><h2>Models Guide — what each model does, with pros and cons</h2><div class="hint">Conceptual explanation of the ${MODELS.length} models competing in every segment · live statistics from cycle ${M.cycle}</div></div></div>
  <div class="lead">Every segment runs <b>${MODELS.length} models from four families</b>. Using different families on purpose makes the forecast <b>defensible</b>: if models with very different logic (pure history vs. drivers vs. machine learning) converge to similar volumes, confidence is high; if they diverge, the scenario range shows the uncertainty. The winner of each segment is chosen by backtest accuracy, not by preference.</div>
  <div class="vlist" style="margin-top:14px">${fams.map(f=>{const ms=MODELS.filter(m=>M.models[m].f===f);const won=ms.reduce((a,m)=>a+stat[m].won,0);
    return `<div class="vrow"><span style="font-size:1.4rem">${FAMILY[f].icon}</span><div><b>${f} <span class="fam ${FAMCLS[f]}">${ms.length} model${ms.length>1?"s":""} · ${won} segments won</span></b><span>${FAMILY[f].txt}</span></div></div>`;}).join("")}</div>`;
  fams.forEach(f=>{
    h+=`<div class="step" style="margin-top:28px"><span class="stepn">${FAMILY[f].icon}</span>${f}</div>`;
    MODELS.filter(m=>M.models[m].f===f).forEach(m=>{const g=GUIDE[m];const s=stat[m];
      h+=`<div class="mcard" style="border-left-color:${COLORS[m]}">
        <div class="mc-head"><div><span class="dot-m" style="background:${COLORS[m]}"></span><b class="mc-name">${mn(m)}</b>${famTag(m)}<span class="fam ens">${M.models[m].d===null?"Mixed drivers":M.models[m].d?"Uses drivers":"No drivers"}</span></div>
          <div class="mc-stats"><span><b>${pct(SUM.glob_acc[m])}</b>overall accuracy</span><span><b>${s.won}</b>segments won</span><span><b>${s.top3}</b>times in Top 3</span><span><b>${s.avgRank.toFixed(1)}</b>average rank</span></div></div>
        <p class="mc-idea">${g.idea}</p>
        <p class="mc-analogy">💡 ${g.analogy}</p>
        <div class="fx"><code>${g.eq}</code><br><small>${g.params}</small></div>
        <div class="pc"><div class="pros"><h4>✔ Pros</h4><ul>${g.pros.map(x=>`<li>${x}</li>`).join("")}</ul></div><div class="cons"><h4>✖ Cons</h4><ul>${g.cons.map(x=>`<li>${x}</li>`).join("")}</ul></div></div>
        <div class="mc-wins">🏆 <b>Tends to win when:</b> ${g.wins}</div>
      </div>`;});
  });
  h+=`<div class="step" style="margin-top:28px"><span class="stepn">⚖</span>Side-by-side comparison</div>
  <div class="scroll"><table><thead><tr><th class="l">Model</th><th class="l">Family</th><th class="l">External drivers</th><th class="l">Interpretability</th><th class="l">Handles trend breaks</th><th class="l">Data needed</th><th>Overall accuracy</th><th>Segments won</th></tr></thead><tbody>
  ${MODELS.slice().sort((a,b)=>SUM.glob_acc[b]-SUM.glob_acc[a]).map(m=>{const f=M.models[m].f;const interp=f==="GLM"?"High":f==="Time series"?"High":f==="Ensemble"?"Medium":"Low (via importance)";
    const brk={ETS_HOLT_WINTERS:"Medium",SARIMA:"Low",THETA:"Medium",STL_ETS:"Medium",GLM_POISSON_SERIE:"Medium",GLM_POISSON_GLOBAL:"Medium",GLM_TWEEDIE_GLOBAL:"Medium",XGBOOST:"High",XGBOOST_SEL:"High",LIGHTGBM_SEL:"High",RANDOM_FOREST_SEL:"Medium",ENSEMBLE_MEDIANA:"Medium"}[m];
    const data={SARIMA:"Long",STL_ETS:"≥ 3 years",THETA:"Short OK",RANDOM_FOREST_SEL:"Medium",GLM_POISSON_SERIE:"Medium"}[m]||(f==="Machine learning"?"Medium–long":"Medium");
    return `<tr><td class="l"><span class="dot-m" style="background:${COLORS[m]}"></span><b>${mn(m)}</b></td><td class="l">${famTag(m)}</td><td class="l">${M.models[m].d===null?"Mixed":M.models[m].d?"Yes":"No"}</td><td class="l">${interp}</td><td class="l">${brk}</td><td class="l">${data}</td><td><span class="accb" style="background:${heat(SUM.glob_acc[m])}">${pct(SUM.glob_acc[m])}</span></td><td>${stat[m].won}</td></tr>`;}).join("")}</tbody></table></div>
  <div class="author">Models Guide — <b>Forecast Lab · Cycle ${M.cycle}</b> · developed by <b>Global Reporting &amp; Analytics</b>, AGCO, by <b>Thiago Montoro</b>.</div>`;
  $("#modPanel").innerHTML=h;
}

// ---------------- METHODOLOGY
function renderMet(){
  const W=M.weights, wl={wmape:"Monthly WMAPE",ann_wmape:"Annual WMAPE (years N and N+1)",mape:"Monthly MAPE",abs_mbe:"|Bias| (MBE %)",r2:"Monthly R² (higher = better)",mae:"Monthly MAE",rmse:"Monthly RMSE"};
  $("#metPanel").innerHTML=`
  <div class="ptitle"><div><h2>Methodology — how scenarios are built and the winner is chosen</h2><div class="hint">Rationale to defend the forecast · Cycle ${M.cycle}</div></div></div>
  <div class="lead">The Forecast Lab trains <b>${SUM.n_models} models</b> from four families (time series, GLM, machine learning and ensemble) for each of the <b>${SUM.n_series} Product × Market × Segment series</b>. Each model is evaluated in a <b>backtest</b>: we go back to the same cutoff month in ${M.origins.length} previous years (${M.origins.join(", ")}), train only on past data and compare the forecast with actuals. The model with the best weighted performance wins the segment; the <b>Blend</b> adds up the winners. See the <b>Models Guide</b> tab for a conceptual description of each model.</div>
  <div class="step"><span class="stepn">1</span>Backtest (out-of-sample validation)</div>
  <div class="lead">For each origin, the model forecasts the rest of year N and the full year N+1 — exactly what the cycle requires: closing the current year and forecasting the next.
   <div class="fx"><code>${F1} total = actuals to ${ymLabel(M.cut)} + forecast of the remaining months</code> · <code>${F2} total = sum of the monthly forecast</code></div></div>
  <div class="step"><span class="stepn">2</span>Accuracy metrics</div>
  <div class="scroll"><table><thead><tr><th class="l">Metric</th><th class="l">Formula</th><th class="l">How to read</th></tr></thead><tbody>
   <tr><td class="l"><b>Accuracy</b></td><td class="l">1 − WMAPE</td><td class="l">100% = perfect. Primary metric.</td></tr>
   <tr><td class="l"><b>WMAPE</b></td><td class="l">Σ|Forecast − Actual| / Σ Actual</td><td class="l">Volume-weighted percentage error.</td></tr>
   <tr><td class="l"><b>Annual WMAPE</b></td><td class="l">Σ|Annual forecast total − Actual| / Σ Annual actual</td><td class="l">Error on the year close (N and N+1).</td></tr>
   <tr><td class="l"><b>MAPE</b></td><td class="l">mean of |error| / Actual</td><td class="l">Low-volume months weigh more.</td></tr>
   <tr><td class="l"><b>Bias (MBE %)</b></td><td class="l">Σ(Forecast − Actual) / Σ Actual</td><td class="l">+ over-forecast · − under-forecast.</td></tr>
   <tr><td class="l"><b>R²</b></td><td class="l">1 − Σerror² / Σ(Actual − mean)²</td><td class="l">1 = perfect · 0 = same as the mean · &lt;0 = worse than the mean.</td></tr>
   <tr><td class="l"><b>MAE / RMSE</b></td><td class="l">mean |error| / √ mean error²</td><td class="l">Error in units; RMSE penalizes large errors more.</td></tr></tbody></table></div>
  <div class="step"><span class="stepn">3</span>Weighted ranking (choosing the winner)</div>
  <div class="lead">In each segment, models get a rank on every criterion. <b>Score = weighted average of the ranks</b> (+ tie-break ${M.tiebreak} × WMAPE). Lowest score wins.
   <div class="vlist" style="margin-top:12px">${Object.entries(W).map(([k,v])=>`<div class="vrow"><div><b>${wl[k]||k}</b><span>Weight <b>${v}</b></span></div></div>`).join("")}</div></div>
  <div class="step"><span class="stepn">4</span>Blend and roll-up</div>
  <div class="lead"><b>Blend</b> = forecast of each segment's winner; Product|Market totals add up the winners. In the backtest, the Blend reaches <b>${pct(SUM.blend_acc)}</b> monthly accuracy vs. <b>${pct(SUM.best_single_acc)}</b> for the best single model (${mn(SUM.best_single)}). <i>Note on rigor: this gain is optimistic, because the winner is selected on the same backtest.</i> Aggregate Blend bias: ${spct(SUM.blend_bias)}.</div>
  <div class="step"><span class="stepn">5</span>Explainability and agronomic filter</div>
  <div class="lead"><b>Permutation importance</b> measures how much the model's error worsens when a variable is shuffled — the higher, the more the model depends on it. Only drivers allowed by the <b>agronomic filter</b> are used (each segment uses relevant groups, e.g. combines = grains + credit) and only with the expected effect direction; drivers with the opposite sign are removed and the model is retrained. The <b>effect of +1 standard deviation</b> shows the direction and size of each driver's impact on the forecast.</div>
  <div class="step"><span class="stepn">6</span>Watch-outs for this cycle</div>
  <div class="lead">Segments with winner accuracy below 70% (review with Sales):<br>${SUM.low_acc.map(([s,a])=>`<span class="tag n">${s} ${pct(a,0)}</span>`).join("")}
   ${M.log.length?`<div class="fx"><b>Run log:</b><br>${M.log.map(l=>"• "+l).join("<br>")}</div>`:""}</div>
  <div class="author">Dashboard <b>Forecast Lab · Cycle ${M.cycle}</b> developed by <b>Global Reporting &amp; Analytics</b>, AGCO, by <b>Thiago Montoro</b>. Metrics recalculated from the raw backtest data (${M.file}).</div>`;
}

// ---------------- export
window.exportCSV=function(id,name){const t=document.getElementById(id);if(!t)return;const rows=[...t.querySelectorAll("tr")].map(tr=>[...tr.children].map(td=>'"'+td.innerText.replace(/\s+/g," ").trim().replace(/"/g,'""')+'"').join(","));
 const blob=new Blob(["\ufeff"+rows.join("\n")],{type:"text/csv;charset=utf-8"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`ForecastLab_${M.cycle}_${name}_${st.p}_${st.r}.csv`;a.click();};

function render(){fillFilters();if(st.tab==="cons")renderCons();else if(st.tab==="seg")renderSeg();else if(st.tab==="mod")renderMod();else renderMet();}
render();
})();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
