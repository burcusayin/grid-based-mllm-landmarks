"""
Phase 3a — primary RQ1/RQ2/RQ3 analyses against consensus GT.

Mirrors the v1 analysis structure (RQ1 modality × strategy + per-landmark,
RQ2a strategy paired Wilcoxon at modality level + RQ2b per landmark, RQ3
reproducibility via Fleiss' κ + 3-way unanimous + ED spread, plus SDR with
Wilson CIs, NED, Shapiro-Wilk normality, area cross-rep Jaccard) — but uses
consensus_gt as the canonical reference and the canonical pipeline parser.

Reads:    results_consensus/full_run_records.pkl
Writes:   results_consensus/analysis.json   (RQ1 + RQ2)
          results_consensus/phase_b.json    (SDR/NED/κ/Shapiro/area-rel etc.)
          results_consensus/summary.json    (top-line counts + tokens)

Usage:    .venv/bin/python scripts/analyze_consensus_run.py
"""
from __future__ import annotations

import argparse
import json
import math
import pickle
import statistics
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import config  # noqa: E402

REFS_DEFAULT = "consensus_gt"  # canonical for v2

POINT_LANDMARKS_ORDER = [
    ("CEPHALOMETRIC", "Menton_Me"),
    ("CEPHALOMETRIC", "Nasion_N"),
    ("CEPHALOMETRIC", "Sella_S"),
    ("PERIAPICAL", "Tooth_36_Distal_Apex"),
    ("PERIAPICAL", "Tooth_36_Distal_CEJ"),
    ("PERIAPICAL", "Tooth_36_Mesial_CEJ"),
    ("PANORAMIC", "Mental_Foramen_L"),
    ("PANORAMIC", "Condylar_Head_R"),
    ("PANORAMIC", "Tooth_33_Apex"),
]
AREA_LANDMARKS = ["Mandibular_Canal_L", "Maxillary_Sinus_R", "External_Oblique_Ridge_R"]


# ── Bootstrap CI (percentile) ───────────────────────────────────────

def bootstrap_ci(values, stat_fn, *, n=10_000, seed=42, conf=0.95):
    """Percentile bootstrap CI for stat_fn(values). values: list of floats."""
    import random
    rng = random.Random(seed)
    n_v = len(values)
    if n_v == 0:
        return None, None
    samples = []
    for _ in range(n):
        s = [values[rng.randint(0, n_v - 1)] for _ in range(n_v)]
        samples.append(stat_fn(s))
    samples.sort()
    lo = samples[int((1 - conf) / 2 * n)]
    hi = samples[int((1 + conf) / 2 * n) - 1]
    return lo, hi


def wilson_ci(k, n, conf=0.95):
    if n == 0:
        return (0.0, 0.0)
    from scipy.stats import norm
    z = norm.ppf(1 - (1 - conf) / 2)
    p = k / n
    denom = 1 + z * z / n
    centre = (p + z * z / (2 * n)) / denom
    half = z * math.sqrt((p * (1 - p) / n) + z * z / (4 * n * n)) / denom
    return (max(0.0, centre - half), min(1.0, centre + half))


# ── Geometry helpers (same as recompute) ────────────────────────────

def grid_dims(modality):
    g = config.GRID_SPECS[modality]
    return g["cols"], g["rows"]


def diag(modality):
    cols, rows = grid_dims(modality)
    return math.hypot(cols - 1, rows - 1)


# ── Fleiss' κ ───────────────────────────────────────────────────────

def fleiss_kappa(item_rater_cat):
    """item_rater_cat: list of dicts {category: count}, each summing to R raters."""
    if not item_rater_cat:
        return None
    R = sum(item_rater_cat[0].values())
    N = len(item_rater_cat)
    cats = sorted({c for d in item_rater_cat for c in d})
    P_j = {c: sum(d.get(c, 0) for d in item_rater_cat) / (N * R) for c in cats}
    Pi = []
    for d in item_rater_cat:
        s = sum(n * (n - 1) for n in d.values())
        Pi.append(s / (R * (R - 1)))
    P_bar = sum(Pi) / N
    Pe_bar = sum(p ** 2 for p in P_j.values())
    if 1 - Pe_bar == 0:
        return 1.0
    return (P_bar - Pe_bar) / (1 - Pe_bar)


# ── Main ────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sandbox", default="results_consensus")
    ap.add_argument("--reference", default=REFS_DEFAULT,
                    help="Reference to score against (consensus_gt | omfr_1 | omfr_2 | student)")
    args = ap.parse_args()

    sandbox = (ROOT / args.sandbox).resolve()
    ref = args.reference

    records = pickle.load(open(sandbox / "full_run_records.pkl", "rb"))
    print(f"Loaded {len(records)} records from {sandbox}/full_run_records.pkl")
    print(f"Scoring against reference: {ref!r}")

    # Convenience accessor: per-rep ED/Jaccard/Dice/mean for the chosen reference
    def rec_metric(rec, key):
        return rec["metrics"][ref].get(key)

    # ────────────────────────────────────────────────────────────────
    # RQ1 — modality × strategy aggregates
    # ────────────────────────────────────────────────────────────────
    rq1_modstrat = {}
    for mod in ("CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"):
        for strat in ("zero_shot", "guided"):
            for ltype in ("point", "area"):
                if mod != "PANORAMIC" and ltype == "area":
                    continue
                vals = []
                for r in records:
                    if r["modality"] != mod or r["strategy"] != strat or r["landmark_type"] != ltype:
                        continue
                    v = rec_metric(r, "mean_ed" if ltype == "point" else "mean_jaccard")
                    if v is not None:
                        vals.append(v)
                if not vals:
                    continue
                mean = sum(vals) / len(vals)
                med = statistics.median(vals)
                ci_mean = bootstrap_ci(vals, lambda s: sum(s) / len(s))
                ci_med = bootstrap_ci(vals, statistics.median)
                rq1_modstrat[f"{mod}_{strat}_{ltype}"] = {
                    "n": len(vals), "mean": mean, "mean_ci": list(ci_mean),
                    "median": med, "median_ci": list(ci_med),
                }

    # RQ1 — per-landmark CIs (point landmarks)
    rq1_per_lm = {}
    for mod, lm in POINT_LANDMARKS_ORDER:
        for strat in ("zero_shot", "guided"):
            vals = [rec_metric(r, "mean_ed") for r in records
                    if r["modality"] == mod and r["structure"] == lm
                    and r["strategy"] == strat
                    and rec_metric(r, "mean_ed") is not None]
            if not vals:
                continue
            mean = sum(vals) / len(vals)
            med = statistics.median(vals)
            rq1_per_lm[f"{mod}/{lm}/{strat}"] = {
                "n": len(vals), "mean": mean,
                "mean_ci": list(bootstrap_ci(vals, lambda s: sum(s) / len(s))),
                "median": med,
                "median_ci": list(bootstrap_ci(vals, statistics.median)),
            }

    # ────────────────────────────────────────────────────────────────
    # RQ2 — paired Wilcoxon: zero_shot vs guided per (modality × landmark_type)
    # ────────────────────────────────────────────────────────────────
    from scipy import stats as scistats

    def paired_wilcoxon(deltas):
        """deltas: list of (zs - gd) values OR (gd - zs) for area (positive = guided better/worse).
        Returns dict with stat, p, mean_delta, median_delta, rank-biserial r and 95% CI."""
        nz = [d for d in deltas if abs(d) > 1e-12]
        n_total = len(deltas)
        n_nonzero = len(nz)
        if n_nonzero == 0:
            return {"n_total": n_total, "n_nonzero": 0, "stat": 0.0, "p": 1.0,
                    "mean_delta": 0.0, "median_delta": 0.0, "rank_biserial_r": 0.0,
                    "rank_biserial_ci_low": 0.0, "rank_biserial_ci_high": 0.0}
        try:
            res = scistats.wilcoxon(nz, zero_method="wilcox", alternative="two-sided", mode="auto")
            stat, p = float(res.statistic), float(res.pvalue)
        except Exception:
            stat, p = float("nan"), float("nan")
        ranks = scistats.rankdata([abs(d) for d in nz])
        Wp = sum(r for d, r in zip(nz, ranks) if d > 0)
        Wn = sum(r for d, r in zip(nz, ranks) if d < 0)
        rb = (Wp - Wn) / (Wp + Wn) if (Wp + Wn) > 0 else 0.0
        # Bootstrap CI for r
        import random
        rng = random.Random(42)
        n_b = 2000
        rs = []
        for _ in range(n_b):
            samp = [nz[rng.randint(0, n_nonzero - 1)] for _ in range(n_nonzero)]
            r2 = scistats.rankdata([abs(d) for d in samp])
            wp = sum(rk for d, rk in zip(samp, r2) if d > 0)
            wn = sum(rk for d, rk in zip(samp, r2) if d < 0)
            if wp + wn > 0:
                rs.append((wp - wn) / (wp + wn))
        rs.sort()
        lo = rs[int(0.025 * len(rs))] if rs else float("nan")
        hi = rs[int(0.975 * len(rs)) - 1] if rs else float("nan")
        return {"n_total": n_total, "n_nonzero": n_nonzero, "stat": stat, "p": p,
                "mean_delta": sum(deltas) / n_total,
                "median_delta": statistics.median(deltas),
                "rank_biserial_r": rb,
                "rank_biserial_ci_low": lo, "rank_biserial_ci_high": hi}

    # Build per-query strategy pairs
    by_qid = defaultdict(dict)  # qid → {strat: rec}
    for r in records:
        by_qid[r["query_id"]][r["strategy"]] = r

    # RQ2a — modality-level paired tests
    rq2a = {}
    for mod in ("CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"):
        for ltype in ("point", "area"):
            if mod != "PANORAMIC" and ltype == "area":
                continue
            deltas = []
            for qid, by_strat in by_qid.items():
                zs = by_strat.get("zero_shot")
                gd = by_strat.get("guided")
                if not zs or not gd:
                    continue
                if zs["modality"] != mod or zs["landmark_type"] != ltype:
                    continue
                if ltype == "point":
                    a = rec_metric(zs, "mean_ed")
                    b = rec_metric(gd, "mean_ed")
                    if a is None or b is None:
                        continue
                    deltas.append(a - b)  # pos = guided better (smaller ED)... actually pos = zero_shot worse
                    # Note: positive delta = zero_shot ED > guided ED → guided BETTER on that query.
                    # The aggregate sign convention is "zero_shot − guided", so positive overall = guided better.
                    # Hmm — for v1 the convention was "zero_shot − guided" as the delta, with positive meaning guided better.
                    # Let me match v1's convention: report mean_delta as zero_shot − guided.
                else:
                    a = rec_metric(zs, "mean_jaccard")
                    b = rec_metric(gd, "mean_jaccard")
                    if a is None or b is None:
                        continue
                    deltas.append(b - a)  # area: guided − zero_shot, positive = guided better
            if not deltas:
                continue
            rq2a[f"{mod}_{ltype}"] = paired_wilcoxon(deltas)

    # RQ2b — per-landmark paired tests (Bonferroni × 9 + 3)
    rq2b = {}
    for mod, lm in POINT_LANDMARKS_ORDER:
        deltas = []
        for qid, by_strat in by_qid.items():
            zs = by_strat.get("zero_shot")
            gd = by_strat.get("guided")
            if not zs or not gd:
                continue
            if zs["modality"] != mod or zs["structure"] != lm or zs["landmark_type"] != "point":
                continue
            a = rec_metric(zs, "mean_ed")
            b = rec_metric(gd, "mean_ed")
            if a is None or b is None:
                continue
            deltas.append(a - b)  # zero_shot − guided
        if deltas:
            rq2b[f"{mod}/{lm}"] = paired_wilcoxon(deltas)

    rq2b_area = {}
    for lm in AREA_LANDMARKS:
        deltas = []
        for qid, by_strat in by_qid.items():
            zs = by_strat.get("zero_shot")
            gd = by_strat.get("guided")
            if not zs or not gd:
                continue
            if zs["modality"] != "PANORAMIC" or zs["structure"] != lm or zs["landmark_type"] != "area":
                continue
            a = rec_metric(zs, "mean_jaccard")
            b = rec_metric(gd, "mean_jaccard")
            if a is None or b is None:
                continue
            deltas.append(b - a)  # guided − zero_shot
        if deltas:
            rq2b_area[f"PANORAMIC/{lm}/area"] = paired_wilcoxon(deltas)

    # ────────────────────────────────────────────────────────────────
    # SDR + Wilson CIs (point), grand NED, grand Jaccard
    # ────────────────────────────────────────────────────────────────
    THRESH = [(0, "SDR@0"), (1, "SDR@1"), (math.sqrt(2), "SDR@√2"), (2, "SDR@2")]

    sdr_modstrat = {}
    for mod in ("CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"):
        for strat in ("zero_shot", "guided"):
            eds = [rec_metric(r, "mean_ed") for r in records
                   if r["modality"] == mod and r["strategy"] == strat
                   and r["landmark_type"] == "point"
                   and rec_metric(r, "mean_ed") is not None]
            if not eds:
                continue
            n = len(eds)
            d = {"n": n, "mean_ed": sum(eds) / n}
            for thr, lbl in THRESH:
                k = sum(1 for e in eds if e <= thr + 1e-9)
                d[lbl] = k / n
                d[lbl + "_ci"] = list(wilson_ci(k, n))
            sdr_modstrat[f"{mod}_{strat}"] = d

    # Per-landmark SDR
    sdr_lm = {}
    for mod, lm in POINT_LANDMARKS_ORDER:
        for strat in ("zero_shot", "guided"):
            eds = [rec_metric(r, "mean_ed") for r in records
                   if r["modality"] == mod and r["structure"] == lm
                   and r["strategy"] == strat
                   and rec_metric(r, "mean_ed") is not None]
            if not eds:
                continue
            n = len(eds)
            d = {"n": n, "mean_ed": sum(eds) / n}
            for thr, lbl in THRESH:
                k = sum(1 for e in eds if e <= thr + 1e-9)
                d[lbl] = k / n
            sdr_lm[f"{mod}/{lm}/{strat}"] = d

    # NED per modality
    ned_modstrat = {}
    for mod in ("CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"):
        for strat in ("zero_shot", "guided"):
            vals = [rec_metric(r, "mean_ed") / diag(mod) for r in records
                    if r["modality"] == mod and r["strategy"] == strat
                    and r["landmark_type"] == "point"
                    and rec_metric(r, "mean_ed") is not None]
            if vals:
                ned_modstrat[f"{mod}_{strat}"] = {
                    "n": len(vals),
                    "mean_ned": sum(vals) / len(vals),
                    "median_ned": statistics.median(vals),
                }
    grand_ned = {}
    for strat in ("zero_shot", "guided"):
        vals = [rec_metric(r, "mean_ed") / diag(r["modality"]) for r in records
                if r["strategy"] == strat and r["landmark_type"] == "point"
                and rec_metric(r, "mean_ed") is not None]
        if vals:
            grand_ned[strat] = {"n": len(vals),
                                "mean_ned": sum(vals) / len(vals),
                                "median_ned": statistics.median(vals)}

    grand_jacc = {}
    for strat in ("zero_shot", "guided"):
        vals = [rec_metric(r, "mean_jaccard") for r in records
                if r["strategy"] == strat and r["landmark_type"] == "area"
                and rec_metric(r, "mean_jaccard") is not None]
        if vals:
            grand_jacc[strat] = {"n": len(vals),
                                 "mean_jacc": sum(vals) / len(vals),
                                 "median_jacc": statistics.median(vals)}

    # Per-area-landmark
    area_landmark_stats = {}
    for lm in AREA_LANDMARKS:
        for strat in ("zero_shot", "guided"):
            recs = [r for r in records
                    if r["structure"] == lm and r["strategy"] == strat
                    and rec_metric(r, "mean_jaccard") is not None]
            if not recs:
                continue
            js = [rec_metric(r, "mean_jaccard") for r in recs]
            ds = [rec_metric(r, "mean_dice") for r in recs]
            area_landmark_stats[f"PANORAMIC/{lm}/{strat}"] = {
                "n": len(recs),
                "mean_jaccard": sum(js) / len(js),
                "median_jaccard": statistics.median(js),
                "mean_dice": sum(ds) / len(ds),
                "median_dice": statistics.median(ds),
            }

    # ────────────────────────────────────────────────────────────────
    # Shapiro-Wilk on paired deltas (justifies non-parametric tests)
    # ────────────────────────────────────────────────────────────────
    paired_shapiro = {}
    for mod in ("CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"):
        deltas = []
        for qid, by_strat in by_qid.items():
            zs = by_strat.get("zero_shot")
            gd = by_strat.get("guided")
            if not zs or not gd:
                continue
            if zs["modality"] != mod or zs["landmark_type"] != "point":
                continue
            a = rec_metric(zs, "mean_ed")
            b = rec_metric(gd, "mean_ed")
            if a is None or b is None:
                continue
            deltas.append(a - b)
        if 3 <= len(deltas) <= 5000:
            stat, p = scistats.shapiro(deltas)
            paired_shapiro[f"{mod}_point_delta"] = {
                "n": len(deltas), "W": float(stat), "p": float(p),
                "normal_at_alpha_05": bool(p >= 0.05),
            }
    # Area
    deltas = []
    for qid, by_strat in by_qid.items():
        zs = by_strat.get("zero_shot")
        gd = by_strat.get("guided")
        if not zs or not gd:
            continue
        if zs["modality"] != "PANORAMIC" or zs["landmark_type"] != "area":
            continue
        a = rec_metric(zs, "mean_jaccard")
        b = rec_metric(gd, "mean_jaccard")
        if a is None or b is None:
            continue
        deltas.append(b - a)
    if deltas:
        stat, p = scistats.shapiro(deltas)
        paired_shapiro["PANORAMIC_area_delta"] = {
            "n": len(deltas), "W": float(stat), "p": float(p),
            "normal_at_alpha_05": bool(p >= 0.05),
        }

    # ────────────────────────────────────────────────────────────────
    # RQ3 — Reproducibility (Fleiss κ + 3-way unanimous + ED spread + area cross-rep)
    # ────────────────────────────────────────────────────────────────
    fleiss_per = {}
    for mod in ("PANORAMIC", "PERIAPICAL", "CEPHALOMETRIC"):
        for strat in ("zero_shot", "guided"):
            irc = []
            for r in records:
                if r["modality"] != mod or r["strategy"] != strat \
                        or r["landmark_type"] != "point":
                    continue
                cells = []
                for c in r["rep_pred_cells"]:
                    if c is None or len(c) == 0:
                        cells = None
                        break
                    cells.append(c[0])  # first parsed cell per rep for points
                if cells is None or len(cells) != 3:
                    continue
                cnt = Counter(tuple(c) for c in cells)
                irc.append(cnt)
            kappa = fleiss_kappa(irc)
            fleiss_per[f"{mod}_{strat}_point"] = {
                "n_items": len(irc), "kappa": kappa
            }

    # Overall κ across all point queries (mod-prefixed categories so they're disjoint)
    irc_all = []
    for r in records:
        if r["landmark_type"] != "point":
            continue
        cells = []
        for c in r["rep_pred_cells"]:
            if c is None or len(c) == 0:
                cells = None
                break
            cells.append(c[0])
        if cells is None or len(cells) != 3:
            continue
        cnt = Counter(f'{r["modality"]}_{r["strategy"]}_{c}' for c in cells)
        irc_all.append(cnt)
    fleiss_overall = fleiss_kappa(irc_all)

    # 3-way unanimous
    exact_3way = {}
    for mod in ("PANORAMIC", "PERIAPICAL", "CEPHALOMETRIC"):
        for strat in ("zero_shot", "guided"):
            items = [r for r in records
                     if r["modality"] == mod and r["strategy"] == strat
                     and r["landmark_type"] == "point"
                     and all(c is not None and len(c) > 0 for c in r["rep_pred_cells"])]
            n_unan = sum(1 for r in items
                         if len({tuple(c[0]) for c in r["rep_pred_cells"]}) == 1)
            exact_3way[f"{mod}_{strat}_point"] = {
                "n": len(items), "unanimous": n_unan,
                "rate": n_unan / len(items) if items else None,
            }
    overall_unan = sum(v["unanimous"] for v in exact_3way.values())
    overall_n = sum(v["n"] for v in exact_3way.values())

    # ED spread
    ed_spread = {}
    for mod in ("PANORAMIC", "PERIAPICAL", "CEPHALOMETRIC"):
        for strat in ("zero_shot", "guided"):
            spreads = []
            for r in records:
                if r["modality"] != mod or r["strategy"] != strat \
                        or r["landmark_type"] != "point":
                    continue
                eds = [v for v in r["metrics"][ref]["rep_ed"] if v is not None]
                if len(eds) == 3:
                    spreads.append(max(eds) - min(eds))
            if spreads:
                ed_spread[f"{mod}_{strat}"] = {
                    "n": len(spreads),
                    "mean_spread": sum(spreads) / len(spreads),
                    "median_spread": statistics.median(spreads),
                    "max_spread": max(spreads),
                    "min_spread": min(spreads),
                }

    # Area cross-rep mean pairwise Jaccard
    def pairwise_jacc(sets_):
        pairs = [(0, 1), (0, 2), (1, 2)]
        out = []
        for i, j in pairs:
            a, b = sets_[i], sets_[j]
            if a is None or b is None:
                continue
            if not a and not b:
                out.append(1.0)
            elif not a or not b:
                out.append(0.0)
            else:
                out.append(len(a & b) / len(a | b))
        return out

    area_reli = {}
    for strat in ("zero_shot", "guided"):
        per_q = []
        for r in records:
            if r["landmark_type"] != "area" or r["strategy"] != strat:
                continue
            sets_ = [(set(c) if c else None) for c in r["rep_pred_cells"]]
            js = pairwise_jacc(sets_)
            if js:
                per_q.append(sum(js) / len(js))
        if per_q:
            area_reli[f"PANORAMIC_{strat}_area"] = {
                "n": len(per_q),
                "mean_pairwise_jacc": sum(per_q) / len(per_q),
                "median": statistics.median(per_q),
            }

    # ────────────────────────────────────────────────────────────────
    # Token totals + summary
    # ────────────────────────────────────────────────────────────────
    anchor_root = ROOT / "results_full"  # tokens always counted from frozen run
    prompt_tok = compl_tok = 0
    for rep in (1, 2, 3):
        for jl in (anchor_root / f"run{rep}" / "responses").glob("*.jsonl"):
            for line in open(jl):
                obj = json.loads(line)
                u = obj.get("response", {}).get("body", {}).get("usage", {})
                prompt_tok += u.get("prompt_tokens", 0)
                compl_tok += u.get("completion_tokens", 0)

    n_calls = sum(1 for _ in range(len(records) * 3))  # 5400
    n_actual = n_calls
    n_failed = sum(r["n_failed"] for r in records)

    summary = {
        "n_queries": 900,
        "n_calls_per_rep": 1800,
        "n_reps": 3,
        "n_total_calls": n_calls,
        "n_actual_responses": n_actual,
        "n_failures": n_failed,
        "compliance_rate": (n_calls - n_failed) / n_calls,
        "prompt_tokens": prompt_tok,
        "completion_tokens": compl_tok,
        "naive_cost_usd": prompt_tok * 1.25e-6 + compl_tok * 7.5e-6,
        "reference_used": ref,
    }

    # Save outputs
    analysis_out = {
        "RQ1_modality_strategy": rq1_modstrat,
        "RQ1_per_landmark": rq1_per_lm,
        "RQ2a_strategy_per_modality": rq2a,
        "RQ2b_strategy_per_landmark": {**rq2b, **rq2b_area},
        "reference_used": ref,
    }
    phase_b_out = {
        "sdr_modality_with_ci": sdr_modstrat,
        "sdr_landmark": sdr_lm,
        "ned_modality": ned_modstrat,
        "grand_ned": grand_ned,
        "grand_jacc": grand_jacc,
        "area_landmark_stats": area_landmark_stats,
        "paired_shapiro": paired_shapiro,
        "fleiss_per_group": fleiss_per,
        "fleiss_overall_point": fleiss_overall,
        "exact_3way_unanimous": exact_3way,
        "overall_unanimous": {"unanimous": overall_unan, "total": overall_n,
                              "rate": overall_unan / overall_n if overall_n else None},
        "ed_spread_per_group": ed_spread,
        "area_reliability": area_reli,
        "tokens": {"prompt_tokens": prompt_tok, "completion_tokens": compl_tok,
                   "total_tokens": prompt_tok + compl_tok},
        "reference_used": ref,
    }

    out_dir = sandbox
    suffix = "" if ref == "consensus_gt" else f"_{ref}"
    (out_dir / f"analysis{suffix}.json").write_text(json.dumps(analysis_out, indent=2, default=str))
    (out_dir / f"phase_b{suffix}.json").write_text(json.dumps(phase_b_out, indent=2, default=str))
    (out_dir / f"summary{suffix}.json").write_text(json.dumps(summary, indent=2, default=str))
    print(f"Wrote analysis{suffix}.json, phase_b{suffix}.json, summary{suffix}.json")

    # Top-line print
    print(f"\n=== Top-line vs {ref} ===")
    print(f"  Compliance: {summary['compliance_rate']*100:.4f}%  "
          f"({summary['n_failures']} failures / {summary['n_total_calls']} calls)")
    print(f"  Tokens: prompt={prompt_tok:,}, completion={compl_tok:,}")
    for k in ("CEPHALOMETRIC_zero_shot_point", "CEPHALOMETRIC_guided_point",
              "PERIAPICAL_zero_shot_point", "PERIAPICAL_guided_point",
              "PANORAMIC_zero_shot_point", "PANORAMIC_guided_point",
              "PANORAMIC_zero_shot_area", "PANORAMIC_guided_area"):
        d = rq1_modstrat.get(k)
        if d:
            label = "ED" if "point" in k else "Jaccard"
            print(f"  {k}: n={d['n']}, mean {label}={d['mean']:.4f} "
                  f"[{d['mean_ci'][0]:.4f}, {d['mean_ci'][1]:.4f}]")
    print(f"  Fleiss κ overall (point): {float(fleiss_overall):.4f}")
    print(f"  3-way unanimous: {overall_unan}/{overall_n} = "
          f"{overall_unan/overall_n*100:.2f}%")


if __name__ == "__main__":
    main()
