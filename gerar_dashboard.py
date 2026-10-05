# -*- coding: utf-8 -*-
"""
AGCO Forecast Lab — Dashboard generator (index.html)
====================================================
Always picks the MOST RECENT Forecast Lab workbook (*.xlsx) from the output folder,
reads the official rankings of each forecast year (2026 / 2027 / 2028 …), rebuilds the
Blend with the workbook rules, consolidates by Product | Market, and writes a
self-contained index.html ready for GitHub Pages.

Usage
-----
    python gerar_dashboard.py                         # latest workbook in INPUT_DIR -> index.html next to this script
    python gerar_dashboard.py --input-dir "D:\\other\\folder"
    python gerar_dashboard.py --file path\\to\\Forecast_Lab_Ciclo_10+2_xxx.xlsx
    python gerar_dashboard.py --out docs\\index.html

Dependencies: pandas, numpy, openpyxl
Author: Global Reporting & Analytics — AGCO (Thiago Montoro)
"""
import os, re, sys, glob, json, argparse
from datetime import datetime
import numpy as np
import pandas as pd

# ------------------------------------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------------------------------------
INPUT_DIR = r"C:\Users\tm75667\OneDrive - AGCO Corp\Área de Trabalho\AI Projects\Project Analytics\output_forecast_lab"
FILE_PATTERN = "*.xlsx"

PROD = {"TA": "Tractors", "CO": "Combines", "PA": "Planters", "PU": "Sprayers"}
PROD_ORDER = ["TA", "CO", "PA", "PU"]
REG = {"BRA": "Brazil", "ARG": "Argentina", "MEX": "Mexico", "OSA": "OSA"}
REG_ORDER = ["BRA", "ARG", "MEX", "OSA"]
MODEL_INFO = {  # code: (display name, family, uses drivers)
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
TYPE_EN = {"History": "Own sales history", "Seasonality": "Seasonality (month)", "Trend": "Trend", "Driver": "External driver"}
CRITERIA = ["wmape", "ann_wmape", "mape", "abs_mbe", "r2", "mae", "rmse"]
DEFAULT_WEIGHTS = {"wmape": 2, "ann_wmape": 2, "mape": 1, "abs_mbe": 1, "r2": 1, "mae": 1, "rmse": 1}
WEIGHT_LABELS = {"WMAPE (monthly)": "wmape", "Annual WMAPE": "ann_wmape", "Annual WMAPE (N and N+1)": "ann_wmape",
                 "MAPE (monthly)": "mape", "|MBE| bias (monthly)": "abs_mbe", "R² (monthly)": "r2",
                 "MAE (monthly)": "mae", "RMSE (monthly)": "rmse"}
CALC_SHEETS = ["Calc_Metrics", "Calc_Metrics_NextYear", "Calc_Metrics_YearAfter"]  # horizon 0, 1, 2
HCODE = {0: "N", 1: "N1", 2: "N2"}


# ------------------------------------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------------------------------------
def num(x):
    try:
        v = float(x)
        return None if np.isnan(v) or np.isinf(v) else v
    except Exception:
        return None


def r(x, n=4):
    v = num(x)
    return None if v is None else round(v, n)


def seg_en(name):
    n = str(name)
    n = n.replace("20 linhas ou +", "20+ rows").replace("0 < 20 linhas", "< 20 rows").replace("linhas", "rows")
    return n.replace("Classe ", "Class ").replace(" & Over", "+")


def find_input(args):
    if args.file:
        if not os.path.exists(args.file):
            sys.exit(f"❌ File not found: {args.file}")
        return args.file
    here = os.path.dirname(os.path.abspath(__file__))
    folders = [args.input_dir] if args.input_dir else [INPUT_DIR, here]
    for folder in folders:
        if not folder or not os.path.isdir(folder):
            print(f"⚠️  Folder not found, skipping: {folder}")
            continue
        files = [f for f in glob.glob(os.path.join(folder, FILE_PATTERN)) if not os.path.basename(f).startswith("~$")]
        lab = [f for f in files if "forecast_lab" in os.path.basename(f).lower()] or files
        if lab:
            latest = max(lab, key=os.path.getmtime)
            print(f"📁 Folder: {folder}")
            print(f"   {len(lab)} workbook(s) found — using the most recent: {os.path.basename(latest)} "
                  f"(modified {datetime.fromtimestamp(os.path.getmtime(latest)):%d/%m/%Y %H:%M})")
            return latest
    sys.exit("❌ No Forecast Lab workbook found. Check INPUT_DIR or use --file.")


def score_rank(df, weights, tiebreak, criteria=CRITERIA):
    tot, sc = 0.0, 0.0
    for c in criteria:
        w = weights.get(c, 0)
        if not w or c not in df:
            continue
        rk = df[c].rank(ascending=(c != "r2"), method="min", na_option="bottom")
        sc = sc + w * rk
        tot += w
    out = df.copy()
    out["score"] = sc / tot + tiebreak * out["wmape"].fillna(9)
    out["rank"] = out["score"].rank(method="first").astype(int)
    return out


def metrics_from_sums(s):
    """s: dict with sum_act, sum_act2, sum_err, sum_aerr, sum_err2, sum_ape, n_ape, n, sum_ann_aerr, sum_ann_act"""
    sa, n = s["sum_act"], s["n"]
    m = {}
    m["wmape"] = s["sum_aerr"] / sa if sa else np.nan
    m["ann_wmape"] = s["sum_ann_aerr"] / s["sum_ann_act"] if s["sum_ann_act"] else np.nan
    m["mape"] = s["sum_ape"] / s["n_ape"] if s["n_ape"] else np.nan
    m["mbe_pct"] = s["sum_err"] / sa if sa else np.nan
    m["abs_mbe"] = abs(m["mbe_pct"]) if sa else np.nan
    sst = s["sum_act2"] - sa ** 2 / n if n else 0
    m["r2"] = 1 - s["sum_err2"] / sst if sst > 0 else np.nan
    m["mae"] = s["sum_aerr"] / n if n else np.nan
    m["rmse"] = np.sqrt(s["sum_err2"] / n) if n else np.nan
    m["acc"] = 1 - m["wmape"]
    m["ann_acc"] = 1 - m["ann_wmape"]
    return m


SUMKEYS = ["sum_act", "sum_act2", "sum_err", "sum_aerr", "sum_err2", "sum_ape", "n_ape", "n", "sum_ann_aerr", "sum_ann_act"]


# ------------------------------------------------------------------------------------------
# READ
# ------------------------------------------------------------------------------------------
def read_readme(xl):
    rr = xl.parse("README", header=None)
    weights, tiebreak, thr, run_info, cutoffs, log = dict(DEFAULT_WEIGHTS), 0.001, 0.05, "", [], []
    for _, row in rr.iterrows():
        v = [x for x in row.tolist() if pd.notna(x)]
        if not v:
            continue
        a = str(v[0]).strip()
        if a.startswith("Run "):
            run_info = a
        if a in WEIGHT_LABELS and len(v) >= 3 and num(v[2]) is not None:
            weights[WEIGHT_LABELS[a]] = num(v[2])
        if a.startswith("Tie-break") and num(v[-1]) is not None:
            tiebreak = num(v[-1])
        if "clear-winner threshold" in a and num(v[-1]) is not None:
            thr = num(v[-1])
        if a in PROD and len(v) >= 4 and str(v[1]) in REG:
            cutoffs.append({"p": a, "r": str(v[1]), "last": str(v[2]), "cycle": str(v[3])})
        if a == "•" and len(v) >= 2:
            log.append(str(v[1]))
    return weights, tiebreak, thr, run_info, cutoffs, log


def read_monthly(xl):
    m = xl.parse("Data_Monthly", header=None)
    ym = m.iloc[0].tolist()
    mcols = [i for i, v in enumerate(ym) if num(v) is not None and 200001 <= num(v) <= 209912]
    months = [int(num(ym[i])) for i in mcols]
    hdr = [str(x) for x in m.iloc[2].tolist()]
    i_last = hdr.index("Last actual month") if "Last actual month" in hdr else 6
    data, last_actual = {}, {}
    for _, row in m.iloc[3:].iterrows():
        series, model = row[4], row[3]
        if pd.isna(series) or pd.isna(model):
            continue
        data.setdefault(series, {"p": row[0], "r": row[1], "s": row[2], "models": {}})
        data[series]["models"][model] = [num(row[i]) for i in mcols]
        if num(row[i_last]) is not None:
            last_actual[series] = int(num(row[i_last]))
    # ENSEMBLE_MEDIANA is an Excel MEDIAN formula: rebuild it if values are missing
    for s, d in data.items():
        mods = d["models"]
        if "ENSEMBLE_MEDIANA" not in mods:
            continue
        cut = last_actual.get(s, 0)
        ens = list(mods["ENSEMBLE_MEDIANA"])
        for j, ymv in enumerate(months):
            if ymv > cut and ens[j] is None:
                vs = [mods[k][j] for k in mods if k != "ENSEMBLE_MEDIANA" and mods[k][j] is not None]
                ens[j] = float(np.median(vs)) if vs else None
        mods["ENSEMBLE_MEDIANA"] = ens
    return months, data, last_actual


def backtest_sums(bt, ann):
    """Sums per (horizon, model, series) and accuracy per (horizon, model, series, target year)."""
    bt = bt[bt["Actual"].notna() & bt["Forecast"].notna()].copy()
    bt["oy"] = bt["Origin"].astype(str).str[:4].astype(int)
    bt["h"] = bt["Year"] - bt["oy"]
    bt["e"] = bt["Forecast"] - bt["Actual"]
    bt["ae"] = bt["e"].abs()
    bt["e2"] = bt["e"] ** 2
    bt["a2"] = bt["Actual"] ** 2
    pos = bt["Actual"] > 0
    bt["ape"] = np.where(pos, bt["ae"] / bt["Actual"].where(pos, 1), 0.0)
    bt["nape"] = pos.astype(int)
    g = bt.groupby(["h", "Model", "Series"]).agg(sum_act=("Actual", "sum"), sum_act2=("a2", "sum"), sum_err=("e", "sum"),
                                                   sum_aerr=("ae", "sum"), sum_err2=("e2", "sum"), sum_ape=("ape", "sum"),
                                                   n_ape=("nape", "sum"), n=("Actual", "count"))
    ann = ann[ann["Actual for year"].notna() & ann["Actual YTD + forecast"].notna()].copy()
    ann["oy"] = ann["Origin"].astype(str).str[:4].astype(int)
    ann["h"] = ann["Year"] - ann["oy"]
    ann["ae"] = (ann["Actual YTD + forecast"] - ann["Actual for year"]).abs()
    ga = ann.groupby(["h", "Model", "Series"]).agg(sum_ann_aerr=("ae", "sum"), sum_ann_act=("Actual for year", "sum"))
    sums = g.join(ga, how="left").fillna({"sum_ann_aerr": 0, "sum_ann_act": 0})
    # track record: monthly sums by target year
    tr = bt.groupby(["h", "Model", "Series", "Year"]).agg(ae=("ae", "sum"), a=("Actual", "sum"), n=("Actual", "count"))
    # actual annual volumes of past years (target years of the annual backtest)
    hist = ann.drop_duplicates(["Series", "Year"]).set_index(["Series", "Year"])["Actual for year"]
    return sums, tr, hist


def read_calc(xl, sheet):
    """Official metrics and ranking for one horizon (cached Excel values). None if not available."""
    if sheet not in xl.sheet_names:
        return None
    c = xl.parse(sheet)
    need = {"Model", "Series", "Rank", "Score", "WMAPE"}
    if not need.issubset(c.columns) or c["Rank"].isna().all():
        return None
    ren = {"Σ Actual": "sum_act", "Σ Actual²": "sum_act2", "Σ Error": "sum_err", "Σ |Error|": "sum_aerr", "Σ Error²": "sum_err2",
           "Σ APE": "sum_ape", "N APE": "n_ape", "Number of months": "n", "Σ |Annual error|": "sum_ann_aerr",
           "Σ Annual actual": "sum_ann_act", "WMAPE": "wmape", "Annual WMAPE": "ann_wmape", "MAPE": "mape",
           "MBE %": "mbe_pct", "R²": "r2", "MAE": "mae", "RMSE": "rmse", "Monthly accuracy": "acc",
           "Annual accuracy": "ann_acc", "Score": "score", "Rank": "rank"}
    c = c.rename(columns=ren)
    c["abs_mbe"] = c["mbe_pct"].abs()
    return c.set_index(["Model", "Series"])


def read_blend_validation(xl):
    """Out-of-sample test tables of the Blend rule (tab 'Blend method', section A)."""
    if "Blend method" not in xl.sheet_names:
        return []
    b = xl.parse("Blend method", header=None)
    out, cur = [], None
    for _, row in b.iterrows():
        v = [x for x in row.tolist() if pd.notna(x)]
        if not v:
            cur = None
            continue
        a = str(v[0])
        if a.startswith("A ·"):
            cur = {"title": a, "header": None, "rows": []}
            out.append(cur)
        elif cur is not None and cur["header"] is None:
            cur["header"] = [str(x) for x in v]
        elif cur is not None:
            cur["rows"].append([str(v[0])] + [r(x) for x in v[1:]])
        if a.startswith("B ·"):
            break
    return out


# ------------------------------------------------------------------------------------------
# BUILD
# ------------------------------------------------------------------------------------------
def build(path):
    print(f"📂 Reading: {path}")
    xl = pd.ExcelFile(path)
    weights, tiebreak, thr, run_info, cutoffs, log = read_readme(xl)
    months, data, last_actual = read_monthly(xl)
    sums, track, hist_act = backtest_sums(xl.parse("Data_Backtest"), xl.parse("Data_Annual_Backtest"))
    imp_raw = xl.parse("Data_Importance")
    validation = read_blend_validation(xl)
    origins = sorted({str(o) for o in xl.parse("Data_Backtest", usecols=["Origin"])["Origin"].unique()})

    cut = max(last_actual.values())
    cy = cut // 100
    fc_years = sorted({m // 100 for m in months if m // 100 >= cy})
    h_avail = sorted(sums.index.get_level_values(0).unique())
    horizons = {}
    for y in fc_years:
        h = y - cy
        horizons[y] = {"h": h, "code": HCODE.get(h, f"N{h}"), "proxy": h not in h_avail,
                       "use_h": h if h in h_avail else max(h_avail)}
    past_years = sorted({int(y) for _, y in hist_act.index})
    month_years = sorted({m // 100 for m in months if m // 100 < cy})
    hist_years = sorted(set(past_years) | set(month_years))
    mcyc = re.search(r"Ciclo[_ ]?(\d+\+\d+)", os.path.basename(path))
    cycle = mcyc.group(1) if mcyc else (cutoffs[0]["cycle"] if cutoffs else "")
    models = [m for m in MODEL_INFO if any(m in d["models"] for d in data.values())]

    # ---------- official metrics per horizon (cached) or recalculated
    per_h, source = {}, {}
    for h in h_avail:
        calc = read_calc(xl, CALC_SHEETS[h]) if h < len(CALC_SHEETS) else None
        if calc is not None:
            per_h[h], source[h] = calc, "workbook"
            continue
        rows = []
        for (hh, model, series), s in sums.loc[[h]].iterrows():
            rows.append(dict(Model=model, Series=series, **s.to_dict(), **metrics_from_sums(s.to_dict())))
        df = pd.DataFrame(rows).set_index(["Model", "Series"])
        parts = [score_rank(g.droplevel(1).reset_index().set_index("Model"), weights, tiebreak)
                 .reset_index().assign(Series=sid) for sid, g in df.groupby(level=1)]
        per_h[h] = pd.concat(parts).set_index(["Model", "Series"])
        source[h] = "recalculated"
    print("🧮 Rankings: " + ", ".join(f"{y} ({source.get(v['use_h'])}{', proxy N+'+str(v['use_h']) if v['proxy'] else ''})"
                                      for y, v in horizons.items()))

    def year_totals(vals):
        t = {}
        for ymv, v in zip(months, vals):
            if v is not None:
                t[ymv // 100] = t.get(ymv // 100, 0.0) + v
        return t

    # ---------- segments
    segments = {}
    for sid, d in data.items():
        cutm = last_actual.get(sid, cut)
        anyv = next(iter(d["models"].values()))
        actual = [v if ymv <= cutm else None for ymv, v in zip(months, anyv)]
        hist = {}
        for y in hist_years:
            if (sid, y) in hist_act.index:
                hist[y] = float(hist_act.loc[(sid, y)])
        for y, v in year_totals(actual).items():
            if y < cy:
                hist[y] = v
        ytd = sum(v for ymv, v in zip(months, actual) if v is not None and ymv // 100 == cy)
        mods = {}
        for m in models:
            if m not in d["models"]:
                continue
            yt = year_totals(d["models"][m])
            entry = {"y": {str(y): r(yt.get(y), 2) for y in fc_years}, "yr": {}, "track": {}}
            for y, hz in horizons.items():
                k = (m, sid)
                df = per_h[hz["use_h"]]
                if k not in df.index:
                    continue
                q = df.loc[k]
                entry["yr"][str(y)] = {"rank": int(q["rank"]), "score": r(q["score"], 3), "acc": r(q["acc"]),
                                       "ann_acc": r(q["ann_acc"]), "wmape": r(q["wmape"]), "mape": r(q["mape"]),
                                       "bias": r(q["mbe_pct"]), "r2": r(q["r2"]), "mae": r(q["mae"], 2), "rmse": r(q["rmse"], 2)}
            for h in h_avail:
                try:
                    t = track.loc[(h, m, sid)]
                    entry["track"][str(h)] = {str(int(yy)): r(1 - row["ae"] / row["a"]) if row["a"] else None
                                              for yy, row in t.iterrows()}
                except KeyError:
                    pass
            mods[m] = entry
        # Blend per year (workbook rules)
        blend = {}
        for y, hz in horizons.items():
            ranked = sorted([m for m in mods if str(y) in mods[m]["yr"]], key=lambda m: mods[m]["yr"][str(y)]["rank"])
            if not ranked:
                continue
            r1 = ranked[0]
            if hz["h"] == 0:
                used, rule, lead = r1, "Winner of the year (rank 1)", None
            else:
                w1, w2 = mods[r1]["yr"][str(y)]["wmape"], mods[ranked[1]]["yr"][str(y)]["wmape"] if len(ranked) > 1 else None
                lead = (w2 - w1) / w1 if (w1 and w2 is not None) else None
                if lead is not None and lead >= thr:
                    used, rule = r1, f"Clear winner: rank 1 leads rank 2 by {lead:.0%} (≥ {thr:.0%})"
                else:
                    top3 = ranked[:3]
                    vals = sorted(top3, key=lambda m: mods[m]["y"][str(y)] or 0)
                    used = vals[len(vals) // 2]
                    rule = (f"Technical tie (lead {lead:.0%} < {thr:.0%}): median of the top-3"
                            if lead is not None else "Median of the top-3")
            blend[str(y)] = {"r1": r1, "used": used, "rule": rule, "lead": r(lead), "vol": mods[used]["y"][str(y)],
                             "top3": ranked[:3]}
        blend_monthly = []
        for j, ymv in enumerate(months):
            if ymv <= cutm:
                blend_monthly.append(r(actual[j], 2))
            else:
                b = blend.get(str(ymv // 100))
                blend_monthly.append(r(d["models"][b["used"]][j], 3) if b else None)
        segments[sid] = {"p": d["p"], "r": d["r"], "s": seg_en(d["s"]), "raw": d["s"], "cut": cutm,
                         "hist": {str(k): r(v, 1) for k, v in hist.items()}, "ytd": r(ytd, 1),
                         "actual": [r(v, 2) for v in actual], "blend_m": blend_monthly,
                         "fc": {m: [r(v, 3) for v in d["models"][m]] for m in mods}, "models": mods, "blend": blend}

    # ---------- consolidated by Product | Market
    consolidated = {}
    combos = sorted({(v["p"], v["r"]) for v in segments.values()},
                    key=lambda x: (PROD_ORDER.index(x[0]) if x[0] in PROD_ORDER else 9,
                                   REG_ORDER.index(x[1]) if x[1] in REG_ORDER else 9))
    for p, rg in combos:
        sids = [s for s, v in segments.items() if v["p"] == p and v["r"] == rg]
        cutm = max(segments[s]["cut"] for s in sids)
        hist = {}
        for s in sids:
            for y, v in segments[s]["hist"].items():
                if v is not None:
                    hist[y] = hist.get(y, 0) + v
        mods = {}
        for m in models:
            if not all(m in segments[s]["models"] for s in sids):
                continue
            yt = {str(y): r(sum(segments[s]["models"][m]["y"][str(y)] or 0 for s in sids), 1) for y in fc_years}
            mods[m] = {"y": yt, "yr": {}, "track": {}, "fc": list(np.round(np.sum(
                [[v or 0 for v in segments[s]["fc"][m]] for s in sids], axis=0), 2))}
        # metrics & ranking per year
        for y, hz in horizons.items():
            df = per_h[hz["use_h"]]
            rows = {}
            for m in mods:
                agg = {k: 0.0 for k in SUMKEYS}
                for s in sids:
                    if (m, s) in df.index:
                        q = df.loc[(m, s)]
                        for k in SUMKEYS:
                            agg[k] += num(q.get(k)) or 0
                rows[m] = {**agg, **metrics_from_sums(agg)}
            rdf = score_rank(pd.DataFrame(rows).T, weights, tiebreak)
            for m in mods:
                q = rdf.loc[m]
                won = sum(1 for s in sids if segments[s]["blend"][str(y)]["r1"] == m)
                mods[m]["yr"][str(y)] = {"rank": int(q["rank"]), "acc": r(q["acc"]), "ann_acc": r(q["ann_acc"]),
                                         "bias": r(q["mbe_pct"]), "wmape": r(q["wmape"]), "won": won}
        # track record (vectorized)
        trh = track.reset_index()
        trh = trh[trh["Series"].isin(sids)]
        for (h, m), g in trh.groupby(["h", "Model"]):
            if m in mods:
                gg = g.groupby("Year")[["ae", "a"]].sum()
                mods[m]["track"][str(h)] = {str(int(yy)): r(1 - row["ae"] / row["a"]) if row["a"] else None for yy, row in gg.iterrows()}
        # Blend
        bl = {"y": {}, "yr": {}, "track": {}, "used": {}}
        for y, hz in horizons.items():
            bl["y"][str(y)] = r(sum(segments[s]["blend"][str(y)]["vol"] or 0 for s in sids), 1)
            df = per_h[hz["use_h"]]
            agg = {k: 0.0 for k in SUMKEYS}
            for s in sids:
                q = df.loc[(segments[s]["blend"][str(y)]["r1"], s)]
                for k in SUMKEYS:
                    agg[k] += num(q.get(k)) or 0
            mt = metrics_from_sums(agg)
            bl["yr"][str(y)] = {"acc": r(mt["acc"]), "ann_acc": r(mt["ann_acc"]), "bias": r(mt["mbe_pct"]), "wmape": r(mt["wmape"])}
            cnt = {}
            for s in sids:
                u = segments[s]["blend"][str(y)]["used"]
                cnt[u] = cnt.get(u, 0) + 1
            bl["used"][str(y)] = cnt
        for h in h_avail:
            y = cy + h
            parts = {}
            for s in sids:
                r1 = segments[s]["blend"].get(str(y), segments[s]["blend"][str(max(fc_years))])["r1"] if str(y) in segments[s]["blend"] else None
                if r1 is None:
                    continue
                g = trh[(trh["h"] == h) & (trh["Model"] == r1) & (trh["Series"] == s)]
                for _, row in g.iterrows():
                    a = parts.setdefault(int(row["Year"]), [0.0, 0.0])
                    a[0] += row["ae"]; a[1] += row["a"]
            bl["track"][str(h)] = {str(yy): r(1 - v[0] / v[1]) if v[1] else None for yy, v in sorted(parts.items())}
        bl["fc"] = list(np.round(np.sum([[v or 0 for v in segments[s]["blend_m"]] for s in sids], axis=0), 2))
        actual = np.sum([[v or 0 for v in segments[s]["actual"]] for s in sids], axis=0)
        consolidated[f"{p}|{rg}"] = {"p": p, "r": rg, "segs": sids, "cut": cutm, "hist": {k: r(v, 1) for k, v in hist.items()},
                                     "ytd": r(sum(segments[s]["ytd"] or 0 for s in sids), 1),
                                     "actual": [r(v, 2) if ymv <= cutm else None for ymv, v in zip(months, actual)],
                                     "models": mods, "blend": bl}

    # ---------- importance
    imp = imp_raw.rename(columns={"Importance (error increase after shuffling)": "imp", "Effect of +1 standard deviation": "eff"})
    importance = {}
    for (m, sid), g in imp.groupby(["Model", "Series"]):
        g = g[g["imp"].notna()].copy()
        g["imp"] = g["imp"].clip(lower=0)
        tot, dtot = g["imp"].sum(), g.loc[g["Type"] == "Driver", "imp"].sum()
        importance.setdefault(sid, {})[m] = [
            {"v": str(x["Variable"]), "t": str(x["Type"]), "sh": r(x["imp"] / tot if tot > 0 else 0),
             "dsh": r(x["imp"] / dtot) if (x["Type"] == "Driver" and dtot > 0) else None, "eff": r(x["eff"])}
            for _, x in g.sort_values("imp", ascending=False).iterrows()]

    # ---------- global summary
    gs = {}
    for y, hz in horizons.items():
        df = per_h[hz["use_h"]]
        acc = {}
        for m in models:
            sub = df.loc[[k for k in df.index if k[0] == m]]
            acc[m] = r(1 - sub["sum_aerr"].sum() / sub["sum_act"].sum()) if "sum_aerr" in sub else None
        bl = [(segments[s]["blend"][str(y)]["r1"], s) for s in segments]
        e = sum(num(df.loc[k]["sum_aerr"]) or 0 for k in bl)
        a = sum(num(df.loc[k]["sum_act"]) or 0 for k in bl)
        fam = {}
        for s in segments.values():
            f = MODEL_INFO[s["blend"][str(y)]["r1"]][1]
            fam[f] = fam.get(f, 0) + 1
        low = sorted([(f'{PROD.get(v["p"], v["p"])} · {REG.get(v["r"], v["r"])} · {v["s"]}',
                       v["models"][v["blend"][str(y)]["r1"]]["yr"][str(y)]["acc"]) for v in segments.values()], key=lambda x: x[1] or 0)
        gs[str(y)] = {"model_acc": acc, "blend_acc": r(1 - e / a) if a else None,
                      "best": max(acc, key=lambda k: acc[k] or -9), "fam": fam,
                      "low": [[n, v] for n, v in low if v is not None and v < 0.7]}
    y0 = str(fc_years[0])
    print(f"✅ {len(segments)} series · {len(models)} models · Blend accuracy {y0}: {gs[y0]['blend_acc']:.1%}")

    meta = {"cycle": cycle, "cut": cut, "run": run_info, "file": os.path.basename(path),
            "file_time": datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d/%m/%Y %H:%M"),
            "generated": datetime.now().strftime("%d/%m/%Y %H:%M"), "months": months, "hist_years": hist_years,
            "fc_years": fc_years, "horizons": {str(k): v for k, v in horizons.items()}, "origins": origins,
            "weights": weights, "tiebreak": tiebreak, "threshold": thr, "log": log, "cutoffs": cutoffs,
            "source": {str(k): v for k, v in source.items()}, "prod": PROD, "reg": REG, "prod_order": PROD_ORDER,
            "reg_order": REG_ORDER, "models": {k: {"n": v[0], "f": v[1], "d": v[2]} for k, v in MODEL_INFO.items() if k in models},
            "type_en": TYPE_EN}
    return {"meta": meta, "summary": gs, "segments": segments, "consolidated": consolidated,
            "importance": importance, "validation": validation}


def main():
    ap = argparse.ArgumentParser(description="Build the Forecast Lab dashboard (index.html)")
    ap.add_argument("--file", help="specific workbook (default: most recent in INPUT_DIR)")
    ap.add_argument("--input-dir", help=f"folder with the Forecast Lab workbooks (default: {INPUT_DIR})")
    ap.add_argument("--out", help="output html (default: index.html next to this script)")
    a = ap.parse_args()
    path = find_input(a)
    payload = build(path)
    here = os.path.dirname(os.path.abspath(__file__))
    tpl_path = os.path.join(here, "template_dashboard.html")
    if os.path.exists(tpl_path):  # optional override of the embedded layout
        with open(tpl_path, encoding="utf-8") as f:
            tpl = f.read()
    else:
        tpl = TEMPLATE
    js = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), default=lambda o: None)
    js = js.replace("NaN", "null").replace("Infinity", "null")
    out = a.out or os.path.join(here, "index.html")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(tpl.replace("/*__DATA__*/null", js))
    print(f"🚀 Dashboard written: {out}  ({os.path.getsize(out) / 1024:.0f} KB)")


# ==========================================================================================
# HTML TEMPLATE (IMR layout) — embedded so this script works on its own
# ==========================================================================================
TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Forecast Lab | AGCO LATAM — Scenarios, Accuracy & Drivers</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script>if(typeof Chart==="undefined"){document.write('<script src="https://unpkg.com/chart.js@4.4.1/dist/chart.umd.min.js"><\/script>');}</script>
<style>
:root{--red:#C8102E;--red-dark:#8B0A1D;--red-soft:#FDE8EB;--ink:#0E1116;--ink-2:#2B303B;--muted:#6B7280;--bg:#EEF1F5;--card:#FFF;--line:#E6E9EF;
--up:#0B7A3B;--up-bg:#E4F7EC;--down:#C8102E;--down-bg:#FDE8EB;--flat:#8A93A2;--gold:#B8860B;--gold-bg:#FBF3DF;--teal:#0E7C86;--teal-bg:#E1F1F2;--purple:#7A3FB0;--purple-bg:#F3EEFA;
--y0:#0B7A3B;--y1:#D97706;--y2:#2563EB;--y3:#7A3FB0;
--shadow:0 1px 2px rgba(16,17,22,.04),0 6px 20px -8px rgba(16,17,22,.12);--shadow-lg:0 20px 45px -20px rgba(16,17,22,.28);--font:"Segoe UI",system-ui,-apple-system,Roboto,Arial,sans-serif}
*{margin:0;padding:0;box-sizing:border-box;font-family:var(--font)}
body{background:radial-gradient(1200px 500px at 100% -10%,rgba(200,16,46,.06),transparent 60%),var(--bg);color:var(--ink-2);padding:28px 22px 70px;line-height:1.5}
.container{max-width:1680px;margin:0 auto}
.hero{position:relative;overflow:hidden;border-radius:22px;padding:28px 34px;background:linear-gradient(135deg,#141821,#1d222d 55%,#241016);color:#fff;box-shadow:var(--shadow-lg);margin-bottom:20px}
.hero::after{content:"";position:absolute;right:-60px;top:-80px;width:340px;height:340px;background:radial-gradient(circle,rgba(200,16,46,.55),transparent 62%);filter:blur(6px)}
.hero-top{display:flex;justify-content:space-between;align-items:flex-start;gap:20px;position:relative;z-index:2;flex-wrap:wrap}
.brand-row{display:flex;align-items:center;gap:16px}
.cy-badge{display:flex;flex-direction:column;align-items:center;justify-content:center;min-width:82px;height:66px;border-radius:14px;background:linear-gradient(145deg,var(--red),var(--red-dark));box-shadow:0 8px 20px -6px rgba(200,16,46,.7);padding:0 12px}
.cy-badge b{font-size:1.55rem;font-weight:800;color:#fff;line-height:1}.cy-badge span{font-size:.52rem;color:#FDE8EB;letter-spacing:.8px;margin-top:4px}
.hero h1{font-size:1.48rem;font-weight:800;letter-spacing:-.5px}.hero .sub{color:#B7BECC;margin-top:4px;font-size:.9rem}
.hero .pill{display:inline-flex;align-items:center;gap:8px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.16);padding:9px 15px;border-radius:999px;font-size:.8rem;font-weight:600;color:#EAEDF3}
.pill .dot{width:8px;height:8px;border-radius:50%;background:#34d17f;box-shadow:0 0 0 4px rgba(52,209,127,.2)}
.hero-meta{display:flex;gap:30px;margin-top:20px;position:relative;z-index:2;flex-wrap:wrap}
.hero-meta .k{font-size:.68rem;text-transform:uppercase;letter-spacing:1.2px;color:#8B93A3}.hero-meta .v{font-size:1rem;font-weight:700;color:#fff;margin-top:2px}.hero-meta .v small{color:#FDE68A;font-weight:700}
.tabs{display:flex;gap:6px;padding:6px;background:linear-gradient(180deg,#1f1f1f,#0b0b0b);border-radius:14px;margin-bottom:14px;box-shadow:0 6px 16px -8px rgba(0,0,0,.45);flex-wrap:wrap}
.tab{padding:11px 20px;border:none;border-radius:10px;background:transparent;color:#C4C4C4;font-weight:700;font-size:.9rem;cursor:pointer;transition:.2s}
.tab:hover{color:#fff;background:rgba(255,255,255,.07)}.tab.active{background:linear-gradient(135deg,var(--red),var(--red-dark));color:#fff;box-shadow:0 4px 12px rgba(200,16,46,.4)}
.filters{display:flex;gap:18px;margin-bottom:16px;flex-wrap:wrap;align-items:center;background:var(--card);border:1px solid var(--line);border-radius:16px;padding:12px 16px;box-shadow:var(--shadow);position:sticky;top:8px;z-index:30}
.fgroup{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.flabel{font-size:.64rem;font-weight:800;letter-spacing:.7px;text-transform:uppercase;color:var(--muted);margin-right:3px}
.fbtn{background:var(--card);color:var(--ink-2);border:1px solid var(--line);padding:6px 14px;border-radius:999px;font-size:.82rem;font-weight:600;cursor:pointer;transition:.18s}
.fbtn:hover{border-color:var(--red);color:var(--red)}.fbtn.active{background:linear-gradient(135deg,var(--red),var(--red-dark));color:#fff;border-color:transparent}
.fbtn:disabled{opacity:.35;cursor:not-allowed}.fbtn.sm{padding:5px 11px;font-size:.77rem}
.ybtn{border:2px solid var(--line);background:#fff;padding:5px 13px;border-radius:999px;font-weight:800;font-size:.8rem;cursor:pointer}
.ybtn.active{color:#fff}
.seg-bar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px}
.seg-btn{background:#fff;border:1px solid var(--line);border-left:4px solid var(--flat);padding:8px 13px;border-radius:10px;font-size:.82rem;font-weight:700;color:var(--ink);cursor:pointer;transition:.15s;text-align:left}
.seg-btn small{display:block;font-weight:600;color:var(--muted);font-size:.67rem}.seg-btn:hover{box-shadow:var(--shadow)}
.seg-btn.active{background:#141821;color:#fff;border-color:#141821}.seg-btn.active small{color:#FDE68A}
.kpi-wrap{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin-bottom:16px}
.kpi{background:var(--card);padding:16px 18px;border-radius:16px;box-shadow:var(--shadow);border:1px solid var(--line);position:relative;overflow:hidden}
.kpi::before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:var(--kc,var(--red))}
.kpi h3{font-size:.66rem;color:var(--muted);text-transform:uppercase;letter-spacing:1px;font-weight:700;margin-bottom:5px}
.kpi .val{font-size:1.3rem;font-weight:800;color:var(--ink);line-height:1.15}.kpi .val small{font-size:.8rem;font-weight:800}
.kpi .desc{margin-top:6px;color:var(--muted);font-size:.77rem;line-height:1.45}
.up{color:var(--up)}.down{color:var(--down)}.flat{color:var(--flat)}
.panel{background:var(--card);padding:20px 22px;border-radius:18px;box-shadow:var(--shadow);border:1px solid var(--line);margin-bottom:18px}
.ptitle{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;gap:14px;flex-wrap:wrap}
.ptitle h2{font-size:1.1rem;color:var(--ink);font-weight:800}.ptitle .hint{font-size:.8rem;color:var(--muted);margin-top:2px;max-width:1000px}
.ptools{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.exp{background:linear-gradient(135deg,#0B7A3B,#0a6c35);color:#fff;border:none;padding:7px 14px;border-radius:10px;font-size:.78rem;font-weight:700;cursor:pointer}
.howto{background:#F4F7FB;border:1px solid #DCE5F0;border-left:4px solid var(--y2);border-radius:10px;padding:11px 14px;font-size:.83rem;line-height:1.6;color:var(--ink-2);margin-bottom:12px}
.howto b{color:var(--ink)}.howto summary{cursor:pointer;font-weight:800;color:var(--y2);list-style:none}.howto summary::before{content:"ⓘ  "}
.read{background:#F9FAFB;border:1px solid var(--line);border-left:4px solid var(--red);border-radius:10px;padding:13px 16px;font-size:.87rem;line-height:1.65;margin-bottom:16px}.read b{color:var(--ink)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px}@media(max-width:1100px){.grid2{grid-template-columns:1fr}}
.chartbox{position:relative;height:360px}.chartbox.tall{height:420px}.chartbox.xl{height:480px}
.scroll{overflow-x:auto;border-radius:12px;border:1px solid var(--line)}
table{width:100%;border-collapse:separate;border-spacing:0}
th,td{padding:9px 11px;text-align:right;border-bottom:1px solid var(--line);font-size:.82rem;white-space:nowrap}
th.l,td.l{text-align:left}
thead th{background:#F7F8FA;color:var(--muted);font-weight:700;text-transform:uppercase;font-size:.64rem;letter-spacing:.4px;position:sticky;top:0;z-index:5}
th.grp{color:#fff;text-align:center;font-size:.72rem;letter-spacing:.6px}
th.hgrp{background:#E9ECF1;color:var(--ink);text-align:center}
td.hist{color:#8A93A2;background:#FAFBFC}td.hist.act{color:var(--ink);font-weight:700}
td.sep,th.sep{border-left:2px solid #D4D9E1}
tbody tr:hover td{background:#FBFCFE}
tr.blend td{background:#141821!important;color:#fff;font-weight:800}tr.blend td.hist{color:#9AA3B2}
tr.actual td{background:#F0F2F5;font-weight:800;color:var(--ink)}
.mname{font-weight:700;color:var(--ink)}tr.blend .mname{color:#fff}
.fam{display:inline-block;font-size:.6rem;font-weight:800;padding:2px 7px;border-radius:999px;margin-left:6px;text-transform:uppercase;letter-spacing:.3px}
.fam.ts{background:var(--teal-bg);color:var(--teal)}.fam.glm{background:var(--gold-bg);color:var(--gold)}.fam.ml{background:var(--purple-bg);color:var(--purple)}.fam.ens{background:#EEF0F3;color:#4b5563}.fam.bl{background:var(--red);color:#fff}
.rk{display:inline-flex;align-items:center;justify-content:center;min-width:26px;height:24px;padding:0 5px;border-radius:7px;font-weight:800;font-size:.76rem;background:#F0F2F5;color:var(--ink)}
.rk.top{color:#fff}
.accb{display:inline-block;min-width:54px;padding:3px 7px;border-radius:7px;font-weight:800;font-size:.77rem;text-align:center}
.dot-m{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:7px;vertical-align:middle}
.heat td{font-size:.75rem;padding:6px 7px}.heat td.cell{text-align:center;font-weight:700;cursor:pointer;min-width:60px}
.heat td.cell.best{outline:2px solid var(--ink);outline-offset:-2px}
.heat td.seg{text-align:left;font-weight:700;color:var(--ink);cursor:pointer;position:sticky;left:0;background:#fff;z-index:2}.heat td.seg:hover{color:var(--red)}
.heat th.mh{writing-mode:vertical-rl;transform:rotate(180deg);height:180px;text-align:left;padding:8px 6px;font-size:.64rem}
.legend{display:flex;gap:14px;flex-wrap:wrap;margin-top:10px;font-size:.73rem;color:var(--muted);align-items:center}
.legend i{width:11px;height:11px;border-radius:3px;display:inline-block;margin-right:5px;vertical-align:middle}
.note{font-size:.8rem;color:var(--muted);background:#fff;border:1px dashed #D4D9E1;border-radius:8px;padding:10px 12px;margin:8px 0}.note b{color:var(--ink-2)}
.tag{display:inline-block;padding:3px 10px;border-radius:999px;font-size:.72rem;font-weight:700;margin:2px 4px 2px 0}.tag.n{background:var(--down-bg);color:var(--down)}.tag.p{background:var(--up-bg);color:var(--up)}.tag.u{background:#EEF0F3;color:#4b5563}
.imp-row{display:grid;grid-template-columns:minmax(230px,1.3fr) 2fr 64px 150px;gap:12px;align-items:center;padding:7px 4px;border-bottom:1px dashed var(--line);font-size:.83rem}
.imp-row .vn{font-weight:700;color:var(--ink)}.vt{display:inline-block;font-size:.58rem;font-weight:800;padding:2px 7px;border-radius:999px;margin-left:6px;text-transform:uppercase}
.vt.Driver{background:var(--red-soft);color:var(--red)}.vt.History{background:#EEF0F3;color:#4b5563}.vt.Seasonality{background:var(--teal-bg);color:var(--teal)}.vt.Trend{background:var(--gold-bg);color:var(--gold)}
.bar{height:14px;border-radius:7px;background:#F0F2F5;overflow:hidden}.bar>div{height:100%;border-radius:7px}
.imp-row .pct{text-align:right;font-weight:800;color:var(--ink)}.eff{font-size:.78rem;font-weight:700}
.stack{display:flex;height:22px;border-radius:7px;overflow:hidden;background:#F0F2F5;min-width:260px}.stack>div{height:100%}
.dmx td.c{text-align:center;font-weight:700;font-size:.76rem}
.step{font-size:.72rem;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);margin:22px 0 10px;display:flex;align-items:center;gap:9px}
.step .stepn{min-width:22px;height:22px;border-radius:50%;color:#fff;font-size:.74rem;font-weight:800;display:flex;align-items:center;justify-content:center;background:var(--red);padding:0 4px}
.lead{background:#F9FAFB;border:1px solid var(--line);border-radius:10px;padding:13px 16px;font-size:.87rem;line-height:1.65;border-left:4px solid var(--red)}
.fx{margin-top:10px;font-size:.8rem;color:var(--muted);background:#fff;border:1px dashed #D4D9E1;border-radius:8px;padding:10px 12px}
.fx code{background:#141821;color:#EAEDF3;padding:2px 8px;border-radius:6px;font-size:.77rem;white-space:normal}
.vlist{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}@media(max-width:820px){.vlist{grid-template-columns:1fr}.imp-row{grid-template-columns:1fr}}
.vrow{display:flex;gap:11px;align-items:flex-start;background:#fff;border:1px solid var(--line);border-radius:10px;padding:12px 14px}
.vrow b{font-size:.86rem;color:var(--ink)}.vrow span{display:block;font-size:.79rem;color:var(--muted);margin-top:2px;line-height:1.5}
.mcard{background:#fff;border:1px solid var(--line);border-left:6px solid var(--red);border-radius:14px;padding:18px 20px;margin-bottom:14px;box-shadow:var(--shadow)}
.mc-head{display:flex;justify-content:space-between;align-items:flex-start;gap:14px;flex-wrap:wrap;margin-bottom:8px}.mc-name{font-size:1.06rem;color:var(--ink)}
.mc-stats{display:flex;gap:8px;flex-wrap:wrap}.mc-stats span{background:#F7F8FA;border:1px solid var(--line);border-radius:10px;padding:6px 10px;font-size:.66rem;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;text-align:center;min-width:86px}
.mc-stats b{display:block;font-size:.98rem;color:var(--ink);text-transform:none;letter-spacing:0}
.mc-idea{font-size:.9rem;line-height:1.6;margin:6px 0}.mc-analogy{font-size:.85rem;color:#521b25;background:#FEF6F7;border-left:3px solid var(--red);border-radius:8px;padding:8px 12px;margin:8px 0}
.pc{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:10px}@media(max-width:820px){.pc{grid-template-columns:1fr}}
.pc h4{font-size:.76rem;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}.pc ul{padding-left:18px;font-size:.84rem;line-height:1.6}
.pros{background:var(--up-bg);border-radius:10px;padding:10px 14px}.pros h4{color:var(--up)}.cons{background:var(--down-bg);border-radius:10px;padding:10px 14px}.cons h4{color:var(--down)}
.mc-wins{margin-top:10px;font-size:.84rem;background:var(--gold-bg);border-radius:8px;padding:8px 12px}
.author{margin-top:20px;padding:15px 20px;border-radius:12px;background:linear-gradient(135deg,#141821,#241016);color:#EAEDF3;font-size:.84rem}.author b{color:#FDE68A}
.foot{text-align:center;color:var(--muted);font-size:.75rem;margin-top:24px;line-height:1.7}
.hidden{display:none!important}
.gloss{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}@media(max-width:1100px){.gloss{grid-template-columns:1fr}}
.gloss div{background:#fff;border:1px solid var(--line);border-radius:10px;padding:11px 13px;font-size:.8rem;color:var(--muted);line-height:1.5}.gloss b{display:block;color:var(--ink);font-size:.87rem;margin-bottom:2px}
</style>
</head>
<body>
<div class="container">
  <div class="hero">
    <div class="hero-top">
      <div class="brand-row">
        <div class="cy-badge"><b id="hCycle">—</b><span>CYCLE</span></div>
        <div><h1>Forecast Lab — Scenarios, Accuracy, Ranking &amp; Drivers</h1>
        <div class="sub" id="hSub"></div></div>
      </div>
      <span class="pill"><span class="dot"></span><span id="hPill">—</span></span>
    </div>
    <div class="hero-meta" id="heroMeta"></div>
  </div>

  <div class="tabs">
    <button class="tab active" data-tab="cons">📊 Overview by Product &amp; Market</button>
    <button class="tab" data-tab="seg">🔍 Segment Deep Dive</button>
    <button class="tab" data-tab="fac">🧩 Factors &amp; Drivers</button>
    <button class="tab" data-tab="mod">🧠 Models Guide</button>
    <button class="tab" data-tab="met">📐 Methodology &amp; Glossary</button>
  </div>

  <div class="filters" id="filters">
    <div class="fgroup"><span class="flabel">Product</span><span id="fProd"></span></div>
    <div class="fgroup"><span class="flabel">Market</span><span id="fReg"></span></div>
    <div class="fgroup" id="fYearG"><span class="flabel">Ranking year</span><span id="fYear"></span></div>
    <div class="fgroup" id="fViewG"><span class="flabel">Models in charts</span>
      <button class="fbtn sm active" data-view="top5">Top 5 + Blend</button><button class="fbtn sm" data-view="all">All models</button></div>
  </div>

  <section id="tab-cons">
    <div class="kpi-wrap" id="cKpis"></div>
    <div class="read" id="cRead"></div>
    <div class="panel">
      <div class="ptitle"><div><h2>Models vs. previous years — volume, accuracy and ranking by forecast year</h2><div class="hint" id="cHint"></div></div>
        <div class="ptools"><button class="exp" onclick="exportCSV('tCons','overview')">Export CSV</button></div></div>
      <details class="howto"><summary>How to read this table</summary>
        Each row is a forecasting model; the black row is the <b>Blend</b> (the final recommended number). Grey columns are <b>actual sales of previous years</b>, repeated on every row so each model can be compared line by line.
        For each forecast year you see: <b>Volume</b> (units), <b>vs. previous year</b> (growth), <b>vs. last actual year</b>, <b>Accuracy</b> (how close this model was to reality when it forecast the same horizon in past years — 100% = perfect) and <b>Rank</b> (1 = best model for that year).
        Each year has its own ranking because a model that is good at <i>closing the current year</i> is not always good at forecasting <i>one or two years ahead</i>. The coloured ★ marks the rank-1 model of each year. Use “Ranking year” in the filter bar to sort the table.
      </details>
      <div class="scroll"><table id="tCons"></table></div>
      <div class="legend" id="legYears"></div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>Track record — how accurate each model was in previous years</h2><div class="hint" id="cTrackHint"></div></div>
        <div class="ptools"><button class="exp" onclick="exportCSV('tTrack','track_record')">Export CSV</button></div></div>
      <details class="howto"><summary>How to read this table</summary>
        We went back in time: at the August cutoff of each past year the model was trained only with data available at that moment and asked to forecast. The columns show how accurate it was <b>for each past year</b>.
        A model that is accurate <b>every year</b> (high minimum, small spread) is more trustworthy than one that was excellent once and poor in other years. Change “Ranking year” to see the track record for closing the current year (N), one year ahead (N+1) or two years ahead (N+2).
      </details>
      <div class="scroll"><table id="tTrack"></table></div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>Scenario fan — actuals and forecasts by model</h2><div class="hint" id="cFanHint"></div></div></div>
      <div class="chartbox tall"><canvas id="cFan"></canvas></div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>Accuracy map — segment × model</h2><div class="hint" id="cHeatHint"></div></div></div>
      <div class="scroll"><table class="heat" id="tHeat"></table></div>
      <div class="legend"><span><i style="background:hsl(0,70%,88%)"></i>&lt; 50%</span><span><i style="background:hsl(40,80%,86%)"></i>50–70%</span><span><i style="background:hsl(90,55%,85%)"></i>70–85%</span><span><i style="background:hsl(135,50%,80%)"></i>&gt; 85%</span><span>Dark outline = rank 1 of the selected year · click a segment to open the deep dive</span></div>
    </div>
  </section>

  <section id="tab-seg" class="hidden">
    <div class="seg-bar" id="segBar"></div>
    <div class="kpi-wrap" id="sKpis"></div>
    <div class="read" id="sRead"></div>
    <div class="panel">
      <div class="ptitle"><div><h2>How the Blend was built for this segment</h2><div class="hint">The final number of each year and the rule that produced it</div></div></div>
      <div class="scroll"><table id="tBlend"></table></div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>Models vs. previous years — volume, accuracy and ranking by forecast year</h2><div class="hint" id="sHint"></div></div>
        <div class="ptools"><button class="exp" onclick="exportCSV('tSeg','segment')">Export CSV</button></div></div>
      <div class="scroll"><table id="tSeg"></table></div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>Track record — accuracy of each model in previous years</h2><div class="hint" id="sTrackHint"></div></div></div>
      <div class="scroll"><table id="tSTrack"></table></div>
    </div>
    <div class="grid2">
      <div class="panel"><div class="ptitle"><div><h2>Scenario fan (annual)</h2><div class="hint">Actual years and forecasts by model</div></div></div><div class="chartbox"><canvas id="sFan"></canvas></div></div>
      <div class="panel"><div class="ptitle"><div><h2>Monthly path</h2><div class="hint">Black = actuals · red = Blend · dashed = other models</div></div></div><div class="chartbox"><canvas id="sMonthly"></canvas></div></div>
    </div>
    <div class="panel"><div class="note" id="sFacLink"></div></div>
  </section>

  <section id="tab-fac" class="hidden">
    <div class="howto" style="border-left-color:var(--red)"><b>What is a “factor”?</b> Anything a model uses to predict sales. There are two big groups:
      <b>(1) Sales-based factors</b> — the segment's own sales history (recent months, moving averages), the usual month-by-month seasonality and the long-term trend; and
      <b>(2) External drivers</b> — agricultural and economic indicators (crop prices and production, interest rates, exchange rate, credit, delinquency…), only those allowed by the agronomic filter.
      Importance is measured by a simple test: <i>scramble one factor and see how much worse the forecast gets</i>. The bigger the damage, the more the model depends on that factor.
      Pure time-series models (ETS, SARIMA, Theta, STL) rely <b>100% on sales history by design</b>.</div>
    <div class="seg-bar" id="facSegBar"></div>
    <div class="kpi-wrap" id="fKpis"></div>
    <div class="panel">
      <div class="ptitle"><div><h2>What each model relies on — sales history vs. external drivers</h2><div class="hint" id="fMixHint"></div></div></div>
      <div class="scroll"><table id="tMix"></table></div>
      <div class="legend"><span><i style="background:#8A93A2"></i>Own sales history</span><span><i style="background:var(--teal)"></i>Seasonality</span><span><i style="background:var(--gold)"></i>Trend</span><span><i style="background:var(--red)"></i>External drivers</span></div>
    </div>
    <div class="grid2">
      <div class="panel"><div class="ptitle"><div><h2>Top external drivers (all models combined)</h2><div class="hint">Average weight among drivers across all model × segment combinations in this scope</div></div></div><div class="chartbox xl"><canvas id="fTop"></canvas></div><div class="legend"><span><i style="background:#0B7A3B"></i>Higher driver → higher forecast</span><span><i style="background:#C8102E"></i>Higher driver → lower forecast</span></div></div>
      <div class="panel"><div class="ptitle"><div><h2>Direction of the effect</h2><div class="hint">What happens to the forecast when the driver goes up (+1 standard deviation)</div></div></div><div class="scroll"><table id="tDir"></table></div></div>
    </div>
    <div class="panel">
      <div class="ptitle"><div><h2>External drivers × models</h2><div class="hint">Weight of each driver among each model's drivers (%). A driver chosen by several different models is stronger evidence.</div></div></div>
      <div class="scroll"><table class="dmx" id="tDxM"></table></div>
    </div>
    <div class="panel" id="fSegPanel">
      <div class="ptitle"><div><h2 id="fSegTitle">Key factors by segment</h2><div class="hint" id="fSegHint"></div></div></div>
      <div id="fSegBody"></div>
    </div>
  </section>

  <section id="tab-mod" class="hidden"><div class="panel" id="modPanel"></div></section>
  <section id="tab-met" class="hidden"><div class="panel" id="metPanel"></div></section>
  <div class="foot" id="foot"></div>
</div>

<script>
const D = /*__DATA__*/null;
(function(){
const M=D.meta,S=D.segments,C=D.consolidated,IMP=D.importance,SUM=D.summary,VAL=D.validation||[];
const MODELS=Object.keys(M.models);
const COLORS={ETS_HOLT_WINTERS:"#0E7C86",SARIMA:"#2BA3AD",THETA:"#6FC6CC",STL_ETS:"#0A4F55",GLM_POISSON_SERIE:"#B8860B",GLM_POISSON_GLOBAL:"#E0A82E",GLM_TWEEDIE_GLOBAL:"#7A5806",XGBOOST:"#7A3FB0",XGBOOST_SEL:"#A57BDB",LIGHTGBM_SEL:"#4E2380",RANDOM_FOREST_SEL:"#C9A6F2",ENSEMBLE_MEDIANA:"#8A93A2",BLEND:"#C8102E",ACTUAL:"#0E1116"};
const FAMCLS={"Time series":"ts","GLM":"glm","Machine learning":"ml","Ensemble":"ens"};
const HY=M.hist_years.map(String),FY=M.fc_years.map(String),LASTH=HY[HY.length-1];
const YCOL={};FY.forEach((y,i)=>YCOL[y]=["#0B7A3B","#D97706","#2563EB","#7A3FB0"][i]||"#444");
const HZ=y=>M.horizons[y];
const HLAB=y=>{const h=HZ(y).h;return h===0?`closing ${y} (horizon N)`:`${h} year${h>1?"s":""} ahead (horizon N+${h})`;};
const st={tab:"cons",p:null,r:null,seg:null,ry:FY[0],view:"top5",fseg:"ALL",fmodel:null};
const charts={};const $=s=>document.querySelector(s);
const nf0=v=>v==null?"–":Math.round(v).toLocaleString("en-US");
const pct=(v,d=1)=>v==null||isNaN(v)?"–":(v*100).toLocaleString("en-US",{minimumFractionDigits:d,maximumFractionDigits:d})+"%";
const spct=(v,d=1)=>v==null||isNaN(v)?"–":(v>0?"+":"")+pct(v,d);
const cls=v=>v==null?"flat":v>0.005?"up":v<-0.005?"down":"flat";
const yoy=(a,b)=>(a!=null&&b)?a/b-1:null;
const mn=m=>m==="BLEND"?"Blend (final number)":(M.models[m]?M.models[m].n:m);
const famTag=m=>m==="BLEND"?'<span class="fam bl">Blend</span>':`<span class="fam ${FAMCLS[M.models[m].f]}">${M.models[m].f}</span>`;
const ymLabel=ym=>["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][ym%100-1]+"/"+String(Math.floor(ym/100)).slice(2);
const heat=a=>{if(a==null)return"#F0F2F5";const x=Math.max(0,Math.min(1,(a-0.3)/0.65));return `hsl(${Math.round(x*135)},${55+15*(1-x)}%,${86-6*x}%)`;};
const segSort=(a,b)=>S[a].s.localeCompare(S[b].s,undefined,{numeric:true});
const rkCell=(q,y)=>!q?"–":`<span class="rk ${q.rank===1?"top":""}" style="${q.rank===1?`background:${YCOL[y]}`:""}">${q.rank===1?"★ ":""}${q.rank}</span>`;
const accCell=v=>`<span class="accb" style="background:${heat(v)}">${pct(v)}</span>`;

// ---------- hero
$("#hCycle").textContent=M.cycle;
$("#hPill").textContent=`Cycle ${M.cycle} · last actual ${ymLabel(M.cut)} · workbook ${M.file_time} · generated ${M.generated}`;
const nS=Object.keys(S).length;
$("#hSub").innerHTML=`${MODELS.length} statistical models · ${nS} series · backtest origins ${M.origins.join(", ")} · one ranking per forecast year`;
$("#heroMeta").innerHTML=[["Coverage",`${nS} series · ${MODELS.length} models`],["Actuals",`${HY[0]}–${LASTH} + ${ymLabel(M.cut)} YTD`],["Forecast",FY.join(" · ")]]
 .concat(FY.map(y=>[`Blend accuracy ${y}`,`${pct(SUM[y].blend_acc)} <small>${HZ(y).h===0?"closing":"N+"+HZ(y).h}</small>`]))
 .map(([k,v])=>`<div><div class="k">${k}</div><div class="v">${v}</div></div>`).join("");
$("#foot").innerHTML=`<b>Forecast Lab · Cycle ${M.cycle}</b> · AGCO Global Reporting &amp; Analytics · Thiago Montoro<br>Source: ${M.file} · ${M.run}`;
$("#legYears").innerHTML=FY.map(y=>`<span><i style="background:${YCOL[y]}"></i>★ rank 1 · ${y} (${HLAB(y)})</span>`).join("")+`<span><i style="background:#141821"></i>Blend = final recommended number</span><span><i style="background:#FAFBFC;border:1px solid #ddd"></i>Actual sales of previous years</span>`;

// ---------- filters
const combos=Object.keys(C);
function fillFilters(){
  const prods=M.prod_order.filter(p=>combos.some(k=>k.startsWith(p+"|")));if(!st.p)st.p=prods[0];
  $("#fProd").innerHTML=prods.map(p=>`<button class="fbtn ${p===st.p?"active":""}" data-p="${p}">${M.prod[p]}</button>`).join(" ");
  const regs=M.reg_order.filter(r=>combos.includes(st.p+"|"+r));if(!regs.includes(st.r))st.r=regs[0];
  $("#fReg").innerHTML=M.reg_order.map(r=>`<button class="fbtn ${r===st.r?"active":""}" data-r="${r}" ${regs.includes(r)?"":"disabled"}>${M.reg[r]}</button>`).join(" ");
  $("#fYear").innerHTML=FY.map(y=>`<button class="ybtn ${y===st.ry?"active":""}" data-y="${y}" style="border-color:${YCOL[y]};${y===st.ry?`background:${YCOL[y]}`:`color:${YCOL[y]}`}">${y}</button>`).join(" ");
  document.querySelectorAll("[data-p]").forEach(b=>b.onclick=()=>{st.p=b.dataset.p;st.seg=null;st.fseg="ALL";render();});
  document.querySelectorAll("[data-r]").forEach(b=>b.onclick=()=>{st.r=b.dataset.r;st.seg=null;st.fseg="ALL";render();});
  document.querySelectorAll("[data-y]").forEach(b=>b.onclick=()=>{st.ry=b.dataset.y;render();});
}
document.querySelectorAll("[data-view]").forEach(b=>b.onclick=()=>{st.view=b.dataset.view;document.querySelectorAll("[data-view]").forEach(x=>x.classList.toggle("active",x===b));render();});
document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>setTab(b.dataset.tab));
function setTab(t){st.tab=t;document.querySelectorAll(".tab").forEach(x=>x.classList.toggle("active",x.dataset.tab===t));
 ["cons","seg","fac","mod","met"].forEach(k=>$("#tab-"+k).classList.toggle("hidden",k!==t));
 $("#filters").classList.toggle("hidden",t==="mod"||t==="met");$("#fYearG").classList.toggle("hidden",t==="fac");$("#fViewG").classList.toggle("hidden",t==="fac");
 render();window.scrollTo({top:0,behavior:"smooth"});}

// ---------- charts
const HASCHART=typeof Chart!=="undefined";
const cutPlugin={id:"cut",beforeDatasetsDraw(ch,a,o){if(o.idx==null||o.idx<0)return;const{ctx,chartArea:ca,scales:{x}}=ch;const px=x.getPixelForValue(o.idx);ctx.save();ctx.fillStyle="rgba(200,16,46,.05)";ctx.fillRect(px,ca.top,ca.right-px,ca.bottom-ca.top);ctx.strokeStyle="rgba(200,16,46,.55)";ctx.setLineDash([5,4]);ctx.beginPath();ctx.moveTo(px,ca.top);ctx.lineTo(px,ca.bottom);ctx.stroke();ctx.setLineDash([]);ctx.fillStyle="#C8102E";ctx.font="700 11px Segoe UI";ctx.fillText("forecast →",px+6,ca.top+12);ctx.restore();}};
if(HASCHART){Chart.register(cutPlugin);Chart.defaults.font.family='"Segoe UI",system-ui,Arial';Chart.defaults.color="#6B7280";}
function mk(id,cfg){const el=document.getElementById(id);if(!el)return;if(!HASCHART){el.parentNode.innerHTML='<div class="note">Chart unavailable: Chart.js did not load (check your internet connection). All tables remain complete.</div>';return;}
 if(charts[id])charts[id].destroy();charts[id]=new Chart(el,cfg);}
const tipNum={callbacks:{label:c=>`${c.dataset.label}: ${nf0(c.parsed.y)}`}};

// ---------- common builders (obj = consolidated or segment)
function rankedModels(obj,y){return Object.keys(obj.models).filter(m=>obj.models[m].yr[y]).sort((a,b)=>obj.models[a].yr[y].rank-obj.models[b].yr[y].rank);}
function blendVol(obj,y){return obj.blend.y?obj.blend.y[y]:(obj.blend[y]?obj.blend[y].vol:null);}
function blendYr(obj,y,isSeg){if(!isSeg)return obj.blend.yr[y];const r1=obj.blend[y].r1;return obj.models[r1].yr[y];}
function yearVal(row,y){return HY.includes(y)?row.hist?row.hist[y]:null:row.y[y];}
function volTable(obj,isSeg){
  const ord=rankedModels(obj,st.ry);
  let h=`<thead><tr><th class="l" rowspan="2">Model</th><th class="hgrp" colspan="${HY.length}">Actual sales (previous years)</th>`;
  FY.forEach((y,i)=>h+=`<th class="grp sep" colspan="${i===0?4:5}" style="background:${YCOL[y]}">${y} · ${HLAB(y)}</th>`);
  h+=`</tr><tr>`+HY.map(y=>`<th>${y}</th>`).join("");
  FY.forEach((y,i)=>{h+=`<th class="sep">Volume</th><th>vs ${i===0?LASTH:FY[i-1]}</th>`+(i>0?`<th>vs ${LASTH}</th>`:"")+`<th>Accuracy</th><th>Rank</th>`;});
  h+=`</tr></thead><tbody>`;
  const hcells=(cl)=>HY.map((y,i)=>{const v=obj.hist[y],pv=i>0?obj.hist[HY[i-1]]:null,g=yoy(v,pv);return `<td class="hist ${cl}">${nf0(v)}${g!=null&&cl?` <small class="${cls(g)}">${spct(g,0)}</small>`:""}</td>`;}).join("");
  // actual row
  h+=`<tr class="actual"><td class="l">Actual sales</td>${hcells("act")}`+FY.map((y,i)=>i===0?`<td class="sep" colspan="4" style="text-align:left;font-weight:600;color:#6B7280">${y} actual YTD (to ${ymLabel(obj.cut)}): <b style="color:#0E1116">${nf0(obj.ytd)}</b></td>`:`<td class="sep" colspan="5"></td>`).join("")+`</tr>`;
  const fcCells=(get,yr,isB)=>FY.map((y,i)=>{const v=get(y),pv=i===0?obj.hist[LASTH]:get(FY[i-1]),g=yoy(v,pv),g2=yoy(v,obj.hist[LASTH]),q=yr(y);
    return `<td class="sep"><b>${nf0(v)}</b></td><td class="${isB?"":cls(g)}">${spct(g)}</td>`+(i>0?`<td class="${isB?"":cls(g2)}">${spct(g2)}</td>`:"")+`<td>${q?(isB?pct(q.acc):accCell(q.acc)):"–"}</td><td>${isB?"–":rkCell(q,y)}</td>`;}).join("");
  h+=`<tr class="blend"><td class="l"><span class="dot-m" style="background:${COLORS.BLEND}"></span><span class="mname">${mn("BLEND")}</span>${famTag("BLEND")}</td>${hcells("")}`+fcCells(y=>blendVol(obj,y),y=>blendYr(obj,y,isSeg),true)+`</tr>`;
  ord.forEach(m=>{const md=obj.models[m];
    h+=`<tr><td class="l"><span class="dot-m" style="background:${COLORS[m]}"></span><span class="mname">${mn(m)}</span>${famTag(m)}</td>${hcells("")}`+fcCells(y=>md.y[y],y=>md.yr[y],false)+`</tr>`;});
  return h+"</tbody>";
}
function trackTable(obj,isSeg){
  const h=String(HZ(st.ry).use_h);const ord=rankedModels(obj,st.ry);
  const yrs=[...new Set(ord.flatMap(m=>Object.keys(obj.models[m].track[h]||{})))].sort();
  const stats=t=>{const v=yrs.map(y=>t[y]).filter(x=>x!=null);if(!v.length)return[null,null,null];const avg=v.reduce((a,b)=>a+b,0)/v.length;return[avg,Math.min(...v),Math.max(...v)-Math.min(...v)];};
  let t=`<thead><tr><th class="l">Model</th><th>Rank ${st.ry}</th>`+yrs.map(y=>`<th>Forecasting ${y}${+y===+M.cut.toString().slice(0,4)?" (YTD)":""}</th>`).join("")+`<th class="sep">Average</th><th>Worst year</th><th>Spread</th></tr></thead><tbody>`;
  let bt=isSeg?obj.models[obj.blend[st.ry].r1].track[h]:obj.blend.track[h];
  if(bt){const[a,mi,sp]=stats(bt);t+=`<tr class="blend"><td class="l"><span class="mname">${isSeg?"Blend (rank-1 model of "+st.ry+")":"Blend (rank-1 models of "+st.ry+")"}</span></td><td>–</td>`+yrs.map(y=>`<td>${pct(bt[y])}</td>`).join("")+`<td class="sep">${pct(a)}</td><td>${pct(mi)}</td><td>${pct(sp)}</td></tr>`;}
  ord.forEach(m=>{const tr=obj.models[m].track[h]||{};const[a,mi,sp]=stats(tr);
    t+=`<tr><td class="l"><span class="dot-m" style="background:${COLORS[m]}"></span><span class="mname">${mn(m)}</span>${famTag(m)}</td><td>${rkCell(obj.models[m].yr[st.ry],st.ry)}</td>`+yrs.map(y=>`<td>${accCell(tr[y])}</td>`).join("")+`<td class="sep"><b>${pct(a)}</b></td><td>${pct(mi)}</td><td class="${sp>0.25?"down":""}">${pct(sp)}</td></tr>`;});
  return t+"</tbody>";
}
function fanChart(id,obj,isSeg){
  const ord=rankedModels(obj,st.ry),vis=st.view==="top5"?ord.slice(0,5):ord,L=HY.concat(FY);
  const act=L.map(y=>HY.includes(y)?obj.hist[y]:null);
  const ds=[{label:"Actual sales",data:act,borderColor:COLORS.ACTUAL,backgroundColor:COLORS.ACTUAL,borderWidth:3,pointRadius:4}];
  ds.push({label:"Blend (final number)",data:L.map(y=>y===LASTH?obj.hist[y]:FY.includes(y)?blendVol(obj,y):null),borderColor:COLORS.BLEND,backgroundColor:COLORS.BLEND,borderWidth:4,pointRadius:5});
  vis.forEach(m=>ds.push({label:mn(m),data:L.map(y=>y===LASTH?obj.hist[y]:FY.includes(y)?obj.models[m].y[y]:null),borderColor:COLORS[m],backgroundColor:COLORS[m],borderWidth:1.5,pointRadius:3,borderDash:[4,3]}));
  mk(id,{type:"line",data:{labels:L,datasets:ds},options:{maintainAspectRatio:false,interaction:{mode:"index",intersect:false},plugins:{cut:{idx:HY.length-1},legend:{position:"bottom",labels:{boxWidth:10,font:{size:10}}},tooltip:tipNum},scales:{y:{ticks:{callback:v=>nf0(v)}}}}});
}
function kpis(obj,isSeg){
  let h="";FY.forEach((y,i)=>{const v=blendVol(obj,y),pv=i===0?obj.hist[LASTH]:blendVol(obj,FY[i-1]),g=yoy(v,pv),q=blendYr(obj,y,isSeg),lead=rankedModels(obj,y)[0];
    h+=`<div class="kpi" style="--kc:${YCOL[y]}"><h3>Blend ${y} · ${HZ(y).h===0?"closing the year":"N+"+HZ(y).h}</h3><div class="val">${nf0(v)} <small class="${cls(g)}">${spct(g)}</small></div>
    <div class="desc">vs ${i===0?LASTH+" actual "+nf0(obj.hist[LASTH]):FY[i-1]} · vs ${LASTH}: <b class="${cls(yoy(v,obj.hist[LASTH]))}">${spct(yoy(v,obj.hist[LASTH]))}</b><br>Accuracy ${pct(q&&q.acc)} · rank 1: <b>${mn(lead)}</b></div></div>`;});
  const allv=FY.map(y=>Object.values(obj.models).map(m=>m.y[y]).filter(v=>v!=null));const ly=FY[FY.length-1],v=allv[allv.length-1];
  h+=`<div class="kpi" style="--kc:var(--gold)"><h3>${ly} scenario range (all models)</h3><div class="val">${nf0(Math.min(...v))} – ${nf0(Math.max(...v))}</div><div class="desc">The wider the range, the higher the uncertainty</div></div>`;
  return h;
}

// ---------- OVERVIEW
function renderCons(){
  const c=C[st.p+"|"+st.r];if(!c)return;
  $("#cKpis").innerHTML=kpis(c,false);
  const lows=c.segs.filter(s=>{const q=S[s].models[S[s].blend[st.ry].r1].yr[st.ry];return q&&q.acc<0.7;});
  const b=FY.map((y,i)=>`<b>${nf0(blendVol(c,y))}</b> in ${y} (<span class="${cls(yoy(blendVol(c,y),i===0?c.hist[LASTH]:blendVol(c,FY[i-1])))}">${spct(yoy(blendVol(c,y),i===0?c.hist[LASTH]:blendVol(c,FY[i-1])))}</span>)`).join(", ");
  $("#cRead").innerHTML=`<b>Quick read — ${M.prod[st.p]} · ${M.reg[st.r]} (${c.segs.length} segments):</b> after selling <b>${nf0(c.hist[LASTH])}</b> units in ${LASTH} (${spct(yoy(c.hist[LASTH],c.hist[HY[HY.length-2]]))} vs ${HY[HY.length-2]}), the Blend projects ${b}.
   Best models: ${FY.map(y=>`<b style="color:${YCOL[y]}">${y}</b> ${mn(rankedModels(c,y)[0])}`).join(" · ")}.
   ${lows.length?`<br>⚠ Segments whose best ${st.ry} model is below 70% accuracy (validate with Sales): ${lows.map(s=>`<span class="tag n">${S[s].s} ${pct(S[s].models[S[s].blend[st.ry].r1].yr[st.ry].acc,0)}</span>`).join("")}`:`<br>✅ All segments have a ${st.ry} best model with accuracy ≥ 70%.`}`;
  $("#cHint").innerHTML=`Volumes summed across the ${c.segs.length} segments. Sorted by the <b style="color:${YCOL[st.ry]}">${st.ry}</b> ranking. Rankings use the workbook weights on the backtest of each horizon (monthly WMAPE ×2, annual WMAPE ×2, MAPE, |bias|, R², MAE, RMSE ×1).`;
  $("#tCons").innerHTML=volTable(c,false);
  $("#cTrackHint").innerHTML=`Accuracy (1 − WMAPE) of each model for <b style="color:${YCOL[st.ry]}">${HLAB(st.ry)}</b>, by past year.${HZ(st.ry).proxy?" ⚠ No backtest exists for this horizon yet: showing the furthest horizon available.":""}`;
  $("#tTrack").innerHTML=trackTable(c,false);
  $("#cFanHint").innerHTML=`${HY[0]}–${LASTH} actuals and ${FY.join("–")} forecasts · ${st.view==="top5"?"Top 5 models of the "+st.ry+" ranking":"all models"} + Blend`;
  fanChart("cFan",c,false);
  $("#cHeatHint").innerHTML=`Accuracy of each model in each segment for <b style="color:${YCOL[st.ry]}">${HLAB(st.ry)}</b>. Change “Ranking year” to compare 2026, 2027 and 2028.`;
  const segs=c.segs.slice().sort(segSort);
  let hh=`<thead><tr><th class="l">Segment</th><th class="l">Rank 1 · ${st.ry}</th><th>Blend ${st.ry}</th>`+MODELS.map(m=>`<th class="mh">${mn(m)}</th>`).join("")+`</tr></thead><tbody>`;
  segs.forEach(s=>{const sg=S[s],r1=sg.blend[st.ry].r1;hh+=`<tr><td class="seg" data-goto="${s}">${sg.s} ↗</td><td class="l"><span class="dot-m" style="background:${COLORS[r1]}"></span>${mn(r1)}</td><td>${nf0(sg.blend[st.ry].vol)}</td>`;
   MODELS.forEach(m=>{const q=sg.models[m]&&sg.models[m].yr[st.ry];hh+=q?`<td class="cell ${q.rank===1?"best":""}" style="background:${heat(q.acc)}" title="${mn(m)} · rank ${q.rank} · accuracy ${pct(q.acc)}" data-goto="${s}">${pct(q.acc,0)}<br><small style="color:#6B7280">#${q.rank}</small></td>`:`<td class="cell">–</td>`;});hh+="</tr>";});
  $("#tHeat").innerHTML=hh+"</tbody>";
  document.querySelectorAll("#tHeat [data-goto]").forEach(el=>el.onclick=()=>{st.seg=el.dataset.goto;setTab("seg");});
}

// ---------- SEGMENT
function segChips(id,cur,withAll,onPick){
  const c=C[st.p+"|"+st.r],segs=c.segs.slice().sort(segSort);
  let h=withAll?`<button class="seg-btn ${cur==="ALL"?"active":""}" data-sp="ALL">All segments<small>${M.prod[st.p]} · ${M.reg[st.r]}</small></button>`:"";
  h+=segs.map(s=>{const r1=S[s].blend[st.ry].r1,q=S[s].models[r1].yr[st.ry];return `<button class="seg-btn ${s===cur?"active":""}" data-sp="${s}" style="border-left-color:${heat(q.acc)}">${S[s].s}<small>${mn(r1)} · ${pct(q.acc,0)}</small></button>`;}).join("");
  $(id).innerHTML=h;$(id).querySelectorAll("[data-sp]").forEach(b=>b.onclick=()=>onPick(b.dataset.sp));
  return segs;
}
function renderSeg(){
  const c=C[st.p+"|"+st.r];if(!c)return;
  const segs=c.segs.slice().sort(segSort);if(!segs.includes(st.seg))st.seg=segs[0];
  segChips("#segBar",st.seg,false,s=>{st.seg=s;renderSeg();});
  const sg=S[st.seg];
  $("#sKpis").innerHTML=kpis(sg,true);
  const r1=sg.blend[st.ry].r1,q=sg.models[r1].yr[st.ry],ord=rankedModels(sg,st.ry),q2=sg.models[ord[1]].yr[st.ry];
  $("#sRead").innerHTML=`<b>${M.prod[sg.p]} · ${M.reg[sg.r]} · ${sg.s} —</b> sold <b>${nf0(sg.hist[LASTH])}</b> units in ${LASTH}; ${nf0(sg.ytd)} until ${ymLabel(sg.cut)}.
   For <b style="color:${YCOL[st.ry]}">${st.ry}</b> the best model is <b>${mn(r1)}</b>: accuracy ${pct(q.acc)} vs ${pct(q2.acc)} for the runner-up (${mn(ord[1])}), bias ${spct(q.bias)} (${q.bias>0?"tends to over-forecast":"tends to under-forecast"}).
   ${q.acc<0.7?`<br>⚠ <b>Accuracy below 70%</b>: treat this number as a reference and validate with Sales.`:""}${Math.abs(q.bias)>0.15?`<br>⚠ Material bias (${spct(q.bias)}): consider a ${q.bias>0?"downward":"upward"} adjustment.`:""}`;
  let bt=`<thead><tr><th class="l">Year</th><th class="l">What is being forecast</th><th class="l">Rank 1 model</th><th>Accuracy rank 1</th><th>Lead over rank 2</th><th class="l">Rule applied</th><th class="l">Model used in the Blend</th><th>Blend volume</th><th>vs previous year</th></tr></thead><tbody>`;
  FY.forEach((y,i)=>{const b=sg.blend[y],pv=i===0?sg.hist[LASTH]:sg.blend[FY[i-1]].vol,g=yoy(b.vol,pv);
    bt+=`<tr><td class="l"><b style="color:${YCOL[y]}">${y}</b></td><td class="l">${HLAB(y)}</td><td class="l"><span class="dot-m" style="background:${COLORS[b.r1]}"></span>${mn(b.r1)}</td><td>${accCell(sg.models[b.r1].yr[y].acc)}</td><td>${b.lead==null?"–":spct(b.lead,0)}</td>
    <td class="l" style="white-space:normal;max-width:380px">${b.rule}${b.used!==b.r1?`<br><small style="color:#6B7280">Top-3: ${b.top3.map(m=>mn(m)+" "+nf0(sg.models[m].y[y])).join(" · ")}</small>`:""}</td><td class="l"><b>${mn(b.used)}</b></td><td><b>${nf0(b.vol)}</b></td><td class="${cls(g)}">${spct(g)}</td></tr>`;});
  $("#tBlend").innerHTML=bt+"</tbody>";
  $("#sHint").innerHTML=`Sorted by the <b style="color:${YCOL[st.ry]}">${st.ry}</b> ranking. Accuracy = 1 − WMAPE in the backtest of each horizon.`;
  $("#tSeg").innerHTML=volTable(sg,true);
  $("#sTrackHint").innerHTML=`Accuracy for <b style="color:${YCOL[st.ry]}">${HLAB(st.ry)}</b>, by past year (same logic as the overview).`;
  $("#tSTrack").innerHTML=trackTable(sg,true);
  fanChart("sFan",sg,true);
  const ci=M.months.indexOf(sg.cut),vis=(st.view==="top5"?ord.slice(0,5):ord);
  const ds=[{label:"Actual sales",data:sg.actual,borderColor:COLORS.ACTUAL,borderWidth:2.4,pointRadius:0},
    {label:"Blend (final number)",data:sg.blend_m.map((v,i)=>M.months[i]>=sg.cut?v:null),borderColor:COLORS.BLEND,borderWidth:3,pointRadius:0}];
  vis.forEach(m=>ds.push({label:mn(m),data:sg.fc[m].map((v,i)=>M.months[i]>=sg.cut?v:null),borderColor:COLORS[m],borderWidth:1.2,pointRadius:0,borderDash:[4,3]}));
  mk("sMonthly",{type:"line",data:{labels:M.months.map(ymLabel),datasets:ds},options:{maintainAspectRatio:false,interaction:{mode:"index",intersect:false},plugins:{cut:{idx:ci},legend:{position:"bottom",labels:{boxWidth:10,font:{size:10}}},tooltip:tipNum},scales:{x:{ticks:{maxTicksLimit:16,font:{size:10}}},y:{ticks:{callback:v=>nf0(v)}}}}});
  $("#sFacLink").innerHTML=`🧩 Want to know <b>what drives</b> these forecasts (sales history vs. external drivers such as crop prices, interest rates or credit)? <a href="#" id="goFac" style="color:var(--red);font-weight:800">Open the Factors &amp; Drivers tab for ${sg.s} →</a>`;
  $("#goFac").onclick=e=>{e.preventDefault();st.fseg=st.seg;setTab("fac");};
}

// ---------- FACTORS
const TC={History:"#8A93A2",Seasonality:"#0E7C86",Trend:"#B8860B",Driver:"#C8102E"};
function scopeSegs(){const c=C[st.p+"|"+st.r];return st.fseg==="ALL"?c.segs:[st.fseg];}
function mixFor(m,segs){const acc={History:0,Seasonality:0,Trend:0,Driver:0};let n=0;
  segs.forEach(s=>{const rows=(IMP[s]||{})[m];if(!rows)return;n++;rows.forEach(r=>acc[r.t]=(acc[r.t]||0)+(r.sh||0));});
  if(!n)return null;Object.keys(acc).forEach(k=>acc[k]/=n);return{mix:acc,n};}
function driverStats(segs){const d={};let pairs=0;
  segs.forEach(s=>Object.entries(IMP[s]||{}).forEach(([m,rows])=>{const dr=rows.filter(r=>r.t==="Driver");if(!dr.length)return;pairs++;
    dr.forEach(r=>{const o=d[r.v]=d[r.v]||{sum:0,n:0,models:{},segs:{},pos:0,neg:0,eff:[]};o.sum+=r.dsh||0;o.n++;o.models[m]=(o.models[m]||[]).concat(r.dsh||0);o.segs[s]=1;
      if(r.eff!=null){o.eff.push(r.eff);if(r.eff>0)o.pos++;else if(r.eff<0)o.neg++;}});}));
  Object.values(d).forEach(o=>{o.score=o.sum/Math.max(pairs,1);o.avgEff=o.eff.length?o.eff.reduce((a,b)=>a+b,0)/o.eff.length:null;});
  return{d,pairs};}
function renderFac(){
  const c=C[st.p+"|"+st.r];if(!c)return;
  if(st.fseg!=="ALL"&&!c.segs.includes(st.fseg))st.fseg="ALL";
  segChips("#facSegBar",st.fseg,true,s=>{st.fseg=s;st.fmodel=null;renderFac();});
  const segs=scopeSegs(),scope=st.fseg==="ALL"?`${M.prod[st.p]} · ${M.reg[st.r]} (all ${segs.length} segments)`:`${M.prod[st.p]} · ${M.reg[st.r]} · ${S[st.fseg].s}`;
  const{d,pairs}=driverStats(segs);const names=Object.keys(d).sort((a,b)=>d[b].score-d[a].score);
  const mixes=MODELS.map(m=>[m,mixFor(m,segs)]);
  const avgDrv=mixes.filter(x=>x[1]).map(x=>x[1].mix.Driver);const avgD=avgDrv.length?avgDrv.reduce((a,b)=>a+b,0)/avgDrv.length:0;
  const top=names[0];
  $("#fKpis").innerHTML=`<div class="kpi" style="--kc:#8A93A2"><h3>Scope</h3><div class="val" style="font-size:1rem">${scope}</div><div class="desc">${pairs} model × segment combinations with drivers</div></div>
   <div class="kpi" style="--kc:var(--red)"><h3>Share explained by external drivers</h3><div class="val">${pct(avgD)}</div><div class="desc">Average across driver-based models. The rest comes from sales history, seasonality and trend.</div></div>
   <div class="kpi" style="--kc:var(--gold)"><h3>Most relevant driver</h3><div class="val" style="font-size:1rem">${top||"–"}</div><div class="desc">${top?`Average weight ${pct(d[top].score)} · used by ${Object.keys(d[top].models).length} model(s)`:"No external driver selected"}</div></div>
   <div class="kpi" style="--kc:var(--teal)"><h3>Distinct drivers used</h3><div class="val">${names.length}</div><div class="desc">Only drivers allowed by the agronomic filter</div></div>`;
  $("#fMixHint").innerHTML=`${scope}. Share of the explanation coming from each type of factor (average across segments). Rank columns = position of the model in the ${FY.join(" / ")} rankings${st.fseg==="ALL"?" (consolidated)":""}.`;
  const obj=st.fseg==="ALL"?c:S[st.fseg];
  let t=`<thead><tr><th class="l">Model</th><th>Rank ${FY.join(" / ")}</th><th class="l">What the forecast relies on</th><th>Sales history</th><th>Seasonality</th><th>Trend</th><th>External drivers</th><th class="l" style="min-width:380px">Top 3 external drivers</th></tr></thead><tbody>`;
  rankedModels(obj,FY[0]).forEach(m=>{const mx=mixes.find(x=>x[0]===m)[1];const rks=FY.map(y=>rkCell(obj.models[m].yr[y],y)).join(" ");
    if(!mx){t+=`<tr><td class="l"><span class="dot-m" style="background:${COLORS[m]}"></span><span class="mname">${mn(m)}</span>${famTag(m)}</td><td>${rks}</td><td class="l"><div class="stack"><div style="width:100%;background:repeating-linear-gradient(45deg,#C9CED6,#C9CED6 6px,#DDE1E7 6px,#DDE1E7 12px)"></div></div></td><td colspan="4" class="l" style="color:#6B7280">${M.models[m].f==="Ensemble"?"Combination of the other models":"100% own sales history, seasonality and trend (by design — no external drivers)"}</td><td class="l">–</td></tr>`;return;}
    const x=mx.mix;const td=[];segs.forEach(s=>((IMP[s]||{})[m]||[]).filter(r=>r.t==="Driver").forEach(r=>{const o=td.find(z=>z.v===r.v);if(o)o.w+=r.dsh||0;else td.push({v:r.v,w:r.dsh||0});}));
    td.sort((a,b)=>b.w-a.w);
    t+=`<tr><td class="l"><span class="dot-m" style="background:${COLORS[m]}"></span><span class="mname">${mn(m)}</span>${famTag(m)}</td><td>${rks}</td><td class="l"><div class="stack">${["History","Seasonality","Trend","Driver"].map(k=>`<div style="width:${(x[k]||0)*100}%;background:${TC[k]}" title="${M.type_en[k]} ${pct(x[k])}"></div>`).join("")}</div></td>
    <td>${pct(x.History)}</td><td>${pct(x.Seasonality)}</td><td>${pct(x.Trend)}</td><td><b style="color:var(--red)">${pct(x.Driver)}</b></td><td class="l" style="white-space:normal;min-width:380px">${td.slice(0,3).map(z=>`<span class="tag n">${z.v}</span>`).join("")||"–"}</td></tr>`;});
  $("#tMix").innerHTML=t+"</tbody>";
  const tn=names.slice(0,15);
  mk("fTop",{type:"bar",data:{labels:tn,datasets:[{label:"Average weight among drivers",data:tn.map(n=>d[n].score*100),backgroundColor:tn.map(n=>d[n].avgEff==null?"#C8102E":d[n].avgEff>=0?"#0B7A3B":"#C8102E")}]},
   options:{indexAxis:"y",maintainAspectRatio:false,plugins:{cut:{idx:null},legend:{display:false},tooltip:{callbacks:{label:c=>`${c.parsed.x.toFixed(1)}% · used by ${Object.keys(d[tn[c.dataIndex]].models).length} models in ${Object.keys(d[tn[c.dataIndex]].segs).length} segment(s)`}}},scales:{x:{ticks:{callback:v=>v+"%"}},y:{ticks:{font:{size:10}}}}}});
  let dt=`<thead><tr><th class="l">Driver</th><th class="l">When it goes up, the forecast…</th><th>Avg. effect (+1 SD)</th><th>Models agreeing</th></tr></thead><tbody>`;
  tn.forEach(n=>{const o=d[n],tot=o.pos+o.neg,agree=tot?Math.max(o.pos,o.neg)/tot:null,dir=o.avgEff==null?"–":o.avgEff>0?'<span class="up">▲ goes up</span>':'<span class="down">▼ goes down</span>';
    dt+=`<tr><td class="l"><b>${n}</b></td><td class="l">${dir}</td><td class="${cls(o.avgEff)}">${o.avgEff==null?"–":spct(o.avgEff,1)}</td><td>${agree==null?"–":pct(agree,0)+" of "+tot}</td></tr>`;});
  $("#tDir").innerHTML=dt+"</tbody>";
  const dm=MODELS.filter(m=>segs.some(s=>(IMP[s]||{})[m]));
  let x=`<thead><tr><th class="l">Driver</th><th># models</th>`+dm.map(m=>`<th>${mn(m)}</th>`).join("")+`</tr></thead><tbody>`;
  names.slice(0,20).forEach(n=>{x+=`<tr><td class="l"><b>${n}</b></td><td>${Object.keys(d[n].models).length}</td>`+dm.map(m=>{const v=d[n].models[m];if(!v)return `<td class="c" style="color:#c3c8d0">–</td>`;const a=v.reduce((p,q)=>p+q,0)/v.length;return `<td class="c" style="background:rgba(200,16,46,${0.08+0.6*a})" title="${v.length} segment(s)">${pct(a,0)}</td>`;}).join("")+"</tr>";});
  if(names.length>20)x+=`<tr><td class="l" colspan="${dm.length+2}" style="color:#6B7280;font-style:italic">+ ${names.length-20} more drivers with marginal weight</td></tr>`;
  $("#tDxM").innerHTML=names.length?x+"</tbody>":`<tbody><tr><td class="l">No external driver selected in this scope.</td></tr></tbody>`;
  // segment level
  if(st.fseg==="ALL"){
    $("#fSegTitle").textContent="Key factors by segment";
    $("#fSegHint").innerHTML=`For each segment: the best model of each year, how much of its forecast comes from external drivers, and the top drivers across all driver-based models. Click a segment above to see the full detail.`;
    let s=`<thead><tr><th class="l">Segment</th>`+FY.map(y=>`<th class="l" style="color:${YCOL[y]}">Rank 1 · ${y}</th>`).join("")+`<th>Driver share (avg.)</th><th class="l">Top drivers (all models)</th><th class="l">Effect</th></tr></thead><tbody>`;
    c.segs.slice().sort(segSort).forEach(sid=>{const ds=driverStats([sid]),nm=Object.keys(ds.d).sort((a,b)=>ds.d[b].score-ds.d[a].score).slice(0,3);
      const mm=MODELS.map(m=>mixFor(m,[sid])).filter(Boolean),sh=mm.length?mm.reduce((a,b)=>a+b.mix.Driver,0)/mm.length:null;
      s+=`<tr><td class="l"><b>${S[sid].s}</b></td>`+FY.map(y=>{const r1=S[sid].blend[y].r1;return `<td class="l"><span class="dot-m" style="background:${COLORS[r1]}"></span>${mn(r1)}${M.models[r1].d?"":' <small style="color:#8A93A2">(no drivers)</small>'}</td>`;}).join("")+
      `<td>${pct(sh)}</td><td class="l" style="white-space:normal">${nm.map(n=>`<span class="tag n">${n} ${pct(ds.d[n].score,0)}</span>`).join("")||"–"}</td><td class="l">${nm.map(n=>ds.d[n].avgEff==null?"":ds.d[n].avgEff>0?'<span class="up">▲</span>':'<span class="down">▼</span>').join(" ")}</td></tr>`;});
    $("#fSegBody").innerHTML=`<div class="scroll"><table>${s}</tbody></table></div>`;
  } else {
    const sg=S[st.fseg],imp=IMP[st.fseg]||{},withImp=rankedModels(sg,st.ry).filter(m=>imp[m]);
    $("#fSegTitle").textContent=`Detailed factors — ${sg.s}`;
    $("#fSegHint").innerHTML=`Pick a model. Bars = share of the total explanation. ▲/▼ = what happens to the forecast when the driver rises by one standard deviation (a “typical” move).`;
    if(!withImp.length){$("#fSegBody").innerHTML=`<div class="note">No importance data for this segment.</div>`;return;}
    if(!st.fmodel||!imp[st.fmodel])st.fmodel=withImp[0];
    const rows=imp[st.fmodel],mx=Math.max(...rows.map(r=>r.sh||0))||1;
    let h=`<div class="ptools" style="margin-bottom:10px">`+withImp.map(m=>`<button class="fbtn sm ${m===st.fmodel?"active":""}" data-fm="${m}">#${sg.models[m].yr[st.ry].rank} ${mn(m)}</button>`).join("")+`</div>`;
    const r1=sg.blend[st.ry].r1;if(!imp[r1])h+=`<div class="note">ℹ The ${st.ry} rank-1 model <b>${mn(r1)}</b> uses no external drivers — it forecasts from the segment's own sales pattern. The models below show which external factors are relevant for this segment.</div>`;
    rows.forEach(r=>{const e=r.eff;h+=`<div class="imp-row"><div><span class="vn">${r.v}</span><span class="vt ${r.t}">${M.type_en[r.t]||r.t}</span></div><div class="bar"><div style="width:${(r.sh||0)/mx*100}%;background:${TC[r.t]}"></div></div><div class="pct">${pct(r.sh)}</div><div>${e==null?'<span class="flat">—</span>':`<span class="eff ${e>0?"up":"down"}">${e>0?"▲":"▼"} ${spct(e,1)}</span> <small class="flat">per +1 SD</small>`}</div></div>`;});
    $("#fSegBody").innerHTML=h;$("#fSegBody").querySelectorAll("[data-fm]").forEach(b=>b.onclick=()=>{st.fmodel=b.dataset.fm;renderFac();});
  }
}

// ---------- MODELS GUIDE
const GUIDE={
 ETS_HOLT_WINTERS:{idea:"Exponential smoothing: it keeps three running numbers — the current level of sales, the trend (damped so it does not run away) and a seasonal factor for each month — and updates them every month, giving more weight to recent months.",analogy:"Like an experienced planner who knows the shape of the year and adjusts to the current pace of sales.",eq:"level = α·(sales − seasonality) + (1 − α)·(previous level + φ·trend)",params:"α, β, γ (how fast it reacts) and φ (trend damping) are estimated automatically.",pros:["Industry standard for demand forecasting","Captures seasonality and recent momentum","Easy to explain and very stable"],cons:["Ignores external information (credit, prices, climate)","Reacts late to sudden market changes","An atypical last quarter can push the forecast too far"],wins:"Long, regular series with stable seasonality."},
 SARIMA:{idea:"Seasonal ARIMA ‘airline’ model: it looks at how this month differs from last month and from the same month last year, and learns how forecast errors carry over.",analogy:"Compares each month with the previous month and with the same month last year, then repeats the pattern.",eq:"(1 − B)(1 − B¹²)·log(sales) = (1 + θB)(1 + ΘB¹²)·error",params:"Fixed structure (0,1,1)(0,1,1)₁₂; θ and Θ estimated from the data.",pros:["Classic, well-documented statistical method","Strong when year-over-year patterns are stable","Good at the shape of the year"],cons:["Sensitive to breaks and outliers","Needs long history (falls back to a recent average on short series)","No external drivers"],wins:"Long, stable series with a clear yearly pattern."},
 THETA:{idea:"Removes seasonality, then averages two views — a long-term trend line and a short-term smoothed level — and puts the seasonality back.",analogy:"Half “where is the long-term trend going?”, half “what happened recently?”.",eq:"forecast = ½ · trend line + ½ · smoothed level  (× seasonal factor)",params:"Seasonal period 12.",pros:["Very robust on short or noisy series","Hard to break; few assumptions","Historically strong in forecasting competitions"],cons:["The trend line can extrapolate too far","Ignores external drivers","Poor at anticipating turning points"],wins:"Noisy or short series."},
 STL_ETS:{idea:"First splits sales into Trend + Seasonality + Remainder with a method that ignores one-off spikes (STL), forecasts the clean trend with ETS, then adds the seasonality back.",analogy:"Cleans the noise and the calendar effect first, forecasts the clean signal, then re-applies the calendar.",eq:"sales = trend + seasonality + remainder",params:"Period 12, robust decomposition.",pros:["Handles outliers and one-off spikes well","Seasonality can evolve over time","Usually among the most accurate single models"],cons:["Needs at least ~3 years of history","No external drivers","End-of-series decomposition can be unstable"],wins:"Series with occasional spikes or changing seasonality."},
 GLM_POISSON_SERIE:{idea:"A regression built for counts (units), fitted per segment: trend + month + up to 3 external drivers that pass the agronomic filter and push sales in the expected direction.",analogy:"A transparent recipe: each ingredient (driver) has a clear % effect on sales.",eq:"sales = exp(β₀ + trend + month + Σ β·driver)",params:"Tests 0–3 drivers, with/without trend; drivers with the wrong sign are discarded; volume cap.",pros:["Highly interpretable (each β is a % effect)","Designed for unit counts","Business logic enforced (correct direction)"],cons:["Assumes stable, proportional relationships","Limited history per segment","Few drivers per model"],wins:"Segments where 1–3 drivers clearly explain demand."},
 GLM_POISSON_GLOBAL:{idea:"One single count regression trained on all segments together: recent sales + month + segment + selected drivers. Small segments borrow information from large ones.",analogy:"One shared rulebook for the whole market, with an adjustment for each segment.",eq:"log(sales) = β₀ + recent sales + month + segment + Σ β·driver",params:"Regularization α = 1; future drivers capped to the historical range.",pros:["Statistical and interpretable","Pools information across segments","Stable on small segments"],cons:["Straight-line relationships (on the log scale)","Month-by-month chaining can carry errors forward","Assumes segments behave alike"],wins:"Markets where segments move together."},
 GLM_TWEEDIE_GLOBAL:{idea:"Same as the global GLM, but with a statistical distribution that tolerates more volatility and months with zero sales.",analogy:"A more forgiving version of the global GLM for irregular, low-volume segments.",eq:"log(sales) = same as the global GLM · variance ∝ mean^1.5",params:"Tweedie power 1.5; α = 1.",pros:["Better for small or irregular segments","Handles months with no sales","Same interpretability as the GLM"],cons:["Same limitations as the GLM","Extra parameter to justify","Harder to explain the distribution"],wins:"Low-volume or intermittent segments."},
 XGBOOST:{idea:"Gradient boosting: hundreds of small decision trees, each one correcting the errors of the previous ones. Uses every driver allowed by the agronomic filter and learns across segments.",analogy:"Like a golfer: the first shot goes toward the hole, every next shot corrects the remaining distance.",eq:"forecast = Σ of small decision trees (recent sales, month, segment, drivers)",params:"400 trees, depth 5, learning rate 0.04.",pros:["Captures non-linear effects and interactions","Uses the full set of agronomic drivers","Strong on complex patterns"],cons:["Black box — explained through importance","Can over-fit with limited history","Does not extrapolate beyond past ranges"],wins:"Complex relationships between many drivers."},
 XGBOOST_SEL:{idea:"Same XGBoost engine, but each segment keeps only the drivers that really improve accuracy (0, 3, 5, 10 or all).",analogy:"XGBoost on a diet: keeps only the drivers that earn their place.",eq:"forecast = Σ trees (recent sales, month, segment, selected drivers)",params:"Same as XGBoost + driver selection per segment.",pros:["Less noise than full XGBoost","More focused, easier to explain","Lower over-fitting risk"],cons:["Selection can change between runs","Still a black box","Relies on a short validation window"],wins:"Segments where only a few drivers matter."},
 LIGHTGBM_SEL:{idea:"An alternative boosting engine that grows trees in a different way (leaf by leaf), with its own driver selection.",analogy:"A faster, more aggressive cousin of XGBoost — a second machine-learning opinion.",eq:"forecast = Σ trees (recent sales, month, segment, selected drivers)",params:"400 trees, 15 leaves.",pros:["Fast","Good accuracy on tabular data","Independent second ML opinion"],cons:["Can over-fit short series","Black box","Sensitive to settings"],wins:"Alternative to XGBoost on medium-length series."},
 RANDOM_FOREST_SEL:{idea:"Averages 250 independent decision trees, each trained on a different sample of the data, using the drivers selected by XGBoost (selected).",analogy:"Asks 250 independent experts and takes the average opinion.",eq:"forecast = average(tree₁ … tree₂₅₀)",params:"250 trees.",pros:["Very stable","Robust to noise and outliers","Little tuning needed"],cons:["Cannot extrapolate trends (stays within past levels)","Under-reacts to strong growth or decline","Black box"],wins:"Noisy series without strong trends."},
 ENSEMBLE_MEDIANA:{idea:"For every month takes the median of the forecasts of all the other models.",analogy:"Wisdom of the crowd: individual errors in opposite directions cancel each other out.",eq:"forecast = median(all models)",params:"All eligible models.",pros:["Extremely stable; hard to beat on average","Protects against any single model going wrong","No extra assumptions"],cons:["Rarely the single best","Not explainable by drivers","Inherits a common bias if most models lean the same way"],wins:"Segments where no single model stands out."}};
const FAMILY={"Time series":["📈","Learn only from the segment's own sales: level, trend and seasonality. Strong, robust baselines that are easy to defend — but blind to credit, prices or climate."],
 "GLM":["📐","Statistical regressions designed for unit counts. Each external driver gets a clear % effect, with business-logic checks on the direction."],
 "Machine learning":["🌲","Tree-based algorithms that capture non-linear effects and interactions between many drivers. Most flexible; explained through factor importance."],
 "Ensemble":["🤝","Combines the other models. Not a market model itself, but a stabilizer against any single model going wrong."]};
function renderMod(){
  const segAll=Object.values(S),st2={};
  MODELS.forEach(m=>{st2[m]={};FY.forEach(y=>{const rr=segAll.map(s=>s.models[m]&&s.models[m].yr[y]).filter(Boolean);st2[m][y]={won:rr.filter(r=>r.rank===1).length,avg:rr.reduce((a,r)=>a+r.rank,0)/rr.length};});});
  const fams=["Time series","GLM","Machine learning","Ensemble"];
  let h=`<div class="ptitle"><div><h2>Models Guide — what each model does, in plain words</h2><div class="hint">${MODELS.length} models compete in every segment · statistics from cycle ${M.cycle}</div></div></div>
  <div class="lead">We deliberately run <b>models with very different logic</b>. If a model based only on sales history, a regression on agricultural drivers and a machine-learning model all point to similar volumes, confidence is high; if they diverge, the scenario range shows the uncertainty. <b>No model is chosen by preference</b>: the winner of each year in each segment is the one that would have been most accurate in past years.</div>
  <div class="vlist" style="margin-top:14px">${fams.map(f=>{const ms=MODELS.filter(m=>M.models[m].f===f);return `<div class="vrow"><span style="font-size:1.4rem">${FAMILY[f][0]}</span><div><b>${f} <span class="fam ${FAMCLS[f]}">${ms.length} model${ms.length>1?"s":""}</span></b><span>${FAMILY[f][1]}</span><span>Segments won: ${FY.map(y=>`<b style="color:${YCOL[y]}">${y}</b> ${ms.reduce((a,m)=>a+st2[m][y].won,0)}`).join(" · ")}</span></div></div>`;}).join("")}</div>`;
  fams.forEach(f=>{h+=`<div class="step" style="margin-top:26px"><span class="stepn">${FAMILY[f][0]}</span>${f}</div>`;
    MODELS.filter(m=>M.models[m].f===f).forEach(m=>{const g=GUIDE[m];
      h+=`<div class="mcard" style="border-left-color:${COLORS[m]}"><div class="mc-head"><div><span class="dot-m" style="background:${COLORS[m]}"></span><b class="mc-name">${mn(m)}</b>${famTag(m)}<span class="fam ens">${M.models[m].d===null?"Mixed":M.models[m].d?"Uses external drivers":"Sales history only"}</span></div>
       <div class="mc-stats">${FY.map(y=>`<span style="border-top:3px solid ${YCOL[y]}"><b>${pct(SUM[y].model_acc[m])}</b>accuracy ${y}</span><span style="border-top:3px solid ${YCOL[y]}"><b>${st2[m][y].won}</b>won ${y}</span>`).join("")}</div></div>
       <p class="mc-idea">${g.idea}</p><p class="mc-analogy">💡 ${g.analogy}</p><div class="fx"><code>${g.eq}</code><br><small>${g.params}</small></div>
       <div class="pc"><div class="pros"><h4>✔ Pros</h4><ul>${g.pros.map(x=>`<li>${x}</li>`).join("")}</ul></div><div class="cons"><h4>✖ Cons</h4><ul>${g.cons.map(x=>`<li>${x}</li>`).join("")}</ul></div></div>
       <div class="mc-wins">🏆 <b>Tends to win when:</b> ${g.wins}</div></div>`;});});
  h+=`<div class="step" style="margin-top:26px"><span class="stepn">⚖</span>Side-by-side</div><div class="scroll"><table><thead><tr><th class="l">Model</th><th class="l">Family</th><th class="l">External drivers</th><th class="l">How explainable</th>`+FY.map(y=>`<th style="color:${YCOL[y]}">Accuracy ${y}</th><th style="color:${YCOL[y]}">Won ${y}</th>`).join("")+`</tr></thead><tbody>`+
   MODELS.slice().sort((a,b)=>(SUM[FY[0]].model_acc[b]||0)-(SUM[FY[0]].model_acc[a]||0)).map(m=>{const f=M.models[m].f;return `<tr><td class="l"><span class="dot-m" style="background:${COLORS[m]}"></span><b>${mn(m)}</b></td><td class="l">${famTag(m)}</td><td class="l">${M.models[m].d===null?"Mixed":M.models[m].d?"Yes":"No"}</td><td class="l">${f==="Machine learning"?"Medium (via factor importance)":"High"}</td>`+FY.map(y=>`<td>${accCell(SUM[y].model_acc[m])}</td><td>${st2[m][y].won}</td>`).join("")+`</tr>`;}).join("")+`</tbody></table></div>
   <div class="author">Models Guide — <b>Forecast Lab · Cycle ${M.cycle}</b> · <b>Global Reporting &amp; Analytics</b>, AGCO · <b>Thiago Montoro</b>.</div>`;
  $("#modPanel").innerHTML=h;
}

// ---------- METHODOLOGY
function renderMet(){
  const W=M.weights,wl={wmape:"Monthly WMAPE",ann_wmape:"Annual WMAPE",mape:"Monthly MAPE",abs_mbe:"|Bias|",r2:"R²",mae:"MAE",rmse:"RMSE"};
  const val=VAL.map(v=>`<div class="step"><span class="stepn">✓</span>${v.title}</div><div class="scroll"><table><thead><tr>${v.header.map((x,i)=>`<th class="${i?"":"l"}">${x}</th>`).join("")}</tr></thead><tbody>${v.rows.map(r=>`<tr ${String(r[0]).includes("used")?'style="background:#F2FBF5;font-weight:700"':""}>${r.map((x,i)=>`<td class="${i?"":"l"}">${i===0?x:(x==null?"–":pct(x))}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`).join("");
  $("#metPanel").innerHTML=`<div class="ptitle"><div><h2>Methodology — how the numbers are produced and why they are defensible</h2><div class="hint">Cycle ${M.cycle} · written for business and technical readers, no data-science background needed</div></div></div>
  <div class="lead">For each of the <b>${Object.keys(S).length} Product × Market × Segment series</b>, <b>${MODELS.length} models</b> produce a monthly forecast until ${FY[FY.length-1]}. To decide which one to trust, we run a <b>backtest</b>: we go back to the same August cutoff in ${M.origins.length} past years (${M.origins.join(", ")}), pretend we do not know the future, forecast it, and compare with what really happened. The model that would have been most accurate wins — <b>separately for each forecast year</b>.</div>
  <div class="step"><span class="stepn">1</span>One ranking per forecast year</div>
  <div class="vlist">${FY.map(y=>`<div class="vrow" style="border-top:4px solid ${YCOL[y]}"><div><b style="color:${YCOL[y]}">${y} — ${HLAB(y)}</b><span>${HZ(y).h===0?`Question: which model best <b>closes ${y}</b>, knowing actuals through ${ymLabel(M.cut)}? Scored on the remaining months of the same year in each backtest (e.g. Sep–Dec) + the annual total.`:`Question: which model best forecasts a full year <b>${HZ(y).h} year${HZ(y).h>1?"s":""} ahead</b>? Scored on the 12 months that came ${HZ(y).h} year${HZ(y).h>1?"s":""} after each backtest cutoff + that year's total.`}${HZ(y).proxy?" ⚠ No backtest at this horizon yet — the furthest available horizon is used as a proxy.":""}</span></div></div>`).join("")}</div>
  <div class="step"><span class="stepn">2</span>How the winner is chosen (same rule every year)</div>
  <div class="lead">Each model is scored on 7 accuracy measures; it gets a position on each one, and the positions are averaged with these weights (lowest average wins; tie-break ${M.tiebreak} × WMAPE):
   <div class="vlist" style="margin-top:10px">${Object.entries(W).map(([k,v])=>`<div class="vrow"><div><b>${wl[k]||k}</b><span>Weight <b>${v}</b></span></div></div>`).join("")}</div></div>
  <div class="step"><span class="stepn">3</span>How the Blend (final number) is built</div>
  <div class="lead"><b>${FY[0]}:</b> forecast of the ${FY[0]} winner (actuals through ${ymLabel(M.cut)} + forecast of the remaining months).<br>
   <b>${FY.slice(1).join(", ")}:</b> the rank-1 model of the year is used <b>only if it clearly leads</b> — its error (WMAPE) must be at least <b>${pct(M.threshold,0)}</b> lower than rank 2. Otherwise the top models are technically tied and the Blend takes the <b>median of the top-3</b> (the middle of the three volumes), which avoids betting on a model that won by a whisker.
   The <b>Segment Deep Dive</b> shows, for every segment and year, which rule was applied and which model was used.</div>
  ${val?`<div class="step"><span class="stepn">4</span>Is the Blend rule validated? (honest out-of-sample test)</div><div class="lead">For each backtest year the ranking was rebuilt <b>without</b> that year and then tested on it — a fair test, because the model is judged on data it did not use to be chosen. Compare the rule in use (green) with “Rank 1 only”: differences of 1–2 points are within normal noise. Accuracy numbers elsewhere in the dashboard are slightly optimistic, because the winner is measured on the same data used to select it.</div>${val}`:""}
  <div class="step"><span class="stepn">5</span>Factors and the agronomic filter</div>
  <div class="lead">Models can only use external drivers that make agronomic sense for the segment (e.g. combines = grains + credit) and that move sales in the expected direction; drivers with the opposite effect are removed and the model is retrained. The <b>Factors &amp; Drivers</b> tab shows how much each model relies on sales history vs. external drivers, and which drivers matter most.</div>
  <div class="step"><span class="stepn">6</span>Glossary in plain words</div>
  <div class="gloss">
   <div><b>Accuracy</b>1 − WMAPE. “How close was the forecast to reality, weighting each month by its volume.” 100% = perfect; 80% = typical error of 20%.</div>
   <div><b>WMAPE</b>Total absolute error ÷ total actual sales. The main error measure.</div>
   <div><b>Annual accuracy</b>Accuracy on the yearly total only — monthly errors that cancel out within the year do not count.</div>
   <div><b>Bias</b>Does the model systematically forecast too high (+) or too low (−)? A good model is close to 0%.</div>
   <div><b>MAPE</b>Average % error per month — gives more weight to small months.</div>
   <div><b>R²</b>Does the model follow the ups and downs? 1 = perfectly, 0 = no better than a flat average.</div>
   <div><b>MAE / RMSE</b>Average error in units; RMSE punishes big misses more.</div>
   <div><b>Backtest</b>Replaying the past: forecasting years we already know, without peeking, to measure real accuracy.</div>
   <div><b>Horizon N / N+1 / N+2</b>Closing the current year / forecasting next year / the year after.</div>
   <div><b>Blend</b>The final recommended number: the best model of each segment and year (or the median of the top-3 when tied), summed.</div>
   <div><b>Factor importance</b>Scramble one factor and see how much worse the forecast gets. More damage = more important.</div>
   <div><b>+1 SD effect</b>Change in the forecast when the driver moves by a “typical” amount (one standard deviation).</div></div>
  <div class="step"><span class="stepn">7</span>Watch-outs for this cycle</div>
  <div class="lead">${FY.map(y=>`<b style="color:${YCOL[y]}">${y}</b> — segments whose best model is below 70% accuracy: ${SUM[y].low.length?SUM[y].low.map(([n,a])=>`<span class="tag n">${n} ${pct(a,0)}</span>`).join(""):"none"}`).join("<br>")}
   ${M.log.length?`<div class="fx"><b>Run log:</b><br>${M.log.map(l=>"• "+l).join("<br>")}</div>`:""}</div>
  <div class="author">Dashboard <b>Forecast Lab · Cycle ${M.cycle}</b> · <b>Global Reporting &amp; Analytics</b>, AGCO · <b>Thiago Montoro</b>. Source workbook: ${M.file} (rankings: ${Object.values(M.source).join(" / ")}).</div>`;
}

window.exportCSV=function(id,name){const t=document.getElementById(id);if(!t)return;const rows=[...t.querySelectorAll("tr")].map(tr=>[...tr.children].map(td=>'"'+td.innerText.replace(/\s+/g," ").trim().replace(/"/g,'""')+'"').join(","));
 const blob=new Blob(["\ufeff"+rows.join("\n")],{type:"text/csv;charset=utf-8"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`ForecastLab_${M.cycle}_${name}_${st.p}_${st.r}.csv`;a.click();};
function render(){fillFilters();({cons:renderCons,seg:renderSeg,fac:renderFac,mod:renderMod,met:renderMet})[st.tab]();}
render();
})();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
