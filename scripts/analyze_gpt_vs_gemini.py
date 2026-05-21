"""
Phase 4 — GPT-5.4 vs Gemini 3.1 Pro paired cross-model comparison.

For each (query_id, strategy) pair, we have:
  - GPT-5.4   mean ED/Jaccard/Dice across its 3 reps (results_consensus)
  - Gemini    mean ED/Jaccard/Dice across its 3 reps (results_full_gemini)

Both are scored against the SAME consensus_gt. The paired test compares
GPT_metric and GEMINI_metric on the same query → the difference is
attributable to the model itself (since prompts + images + GT are byte-
identical between the two runs).

Per (strategy × modality) and per (strategy × landmark), we compute:
  - Paired Wilcoxon signed-rank test on (GPT_metric - GEMINI_metric)
        for ED:      positive Δ → GPT worse / Gemini better
        for Jaccard: positive Δ → GPT better / Gemini worse
        (Jaccard is "higher = better", so we ALSO report Gemini-GPT as
         a sanity column.)
  - Mean Δ + median Δ + rank-biserial r with bootstrap 95% CI
  - Bland-Altman descriptive stats: bias + 95% limits of agreement
  - Effect size (Cohen's d for the paired differences)
  - n_paired, n_skipped (because one side failed)

Bonferroni correction applied across:
  - 4 (strategy × modality × type) families for the modality table
  - 12 (strategy × landmark) families for the per-landmark table

Reads:
    results_consensus/full_run_records.pkl       — GPT-5.4 v2 records
    results_full_gemini/full_run_records.pkl     — Gemini records
    results_consensus/query_index.json           — for landmark order
Writes:
    results_full_gemini/gpt_vs_gemini.json       — full paired-comparison object

Usage: .venv/bin/python scripts/analyze_gpt_vs_gemini.py
       [--gpt-sandbox results_consensus]
       [--gemini-sandbox results_full_gemini]
       [--reference consensus_gt]
"""
from __future__ import annotations

import argparse
import json
import math
import pickle
import random
import statistics
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from scipy import stats as scistats  # noqa: E402

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
STRATEGIES = ("zero_shot", "guided")


def paired_wilcoxon(deltas: list[float]) -> dict:
    """Two-sided Wilcoxon signed-rank on `deltas`, plus rank-biserial r with
    bootstrap (n=2000, seed=42) 95% CI."""
    if not deltas:
        return {"n_total": 0, "n_nonzero": 0, "stat": float("nan"),
                "p": float("nan"), "mean_delta": float("nan"),
                "median_delta": float("nan"), "rank_biserial_r": float("nan"),
                "rank_biserial_ci_low": float("nan"),
                "rank_biserial_ci_high": float("nan")}
    nz = [d for d in deltas if abs(d) > 1e-12]
    n_total = len(deltas)
    n_nonzero = len(nz)
    if n_nonzero == 0:
        return {"n_total": n_total, "n_nonzero": 0, "stat": 0.0, "p": 1.0,
                "mean_delta": 0.0, "median_delta": 0.0, "rank_biserial_r": 0.0,
                "rank_biserial_ci_low": 0.0, "rank_biserial_ci_high": 0.0}
    try:
        res = scistats.wilcoxon(nz, zero_method="wilcox", alternative="two-sided",
                                mode="auto")
        stat, p = float(res.statistic), float(res.pvalue)
    except Exception:
        stat, p = float("nan"), float("nan")
    ranks = scistats.rankdata([abs(d) for d in nz])
    Wp = sum(r for d, r in zip(nz, ranks) if d > 0)
    Wn = sum(r for d, r in zip(nz, ranks) if d < 0)
    rb = (Wp - Wn) / (Wp + Wn) if (Wp + Wn) > 0 else 0.0
    # Bootstrap CI for r — n=2000, seed=42 (matches the existing analyzers)
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
    return {"n_total": n_total, "n_nonzero": n_nonzero,
            "stat": stat, "p": p,
            "mean_delta": sum(deltas) / n_total,
            "median_delta": statistics.median(deltas),
            "rank_biserial_r": rb,
            "rank_biserial_ci_low": lo, "rank_biserial_ci_high": hi}


def bland_altman(values_a: list[float], values_b: list[float]) -> dict:
    """Descriptive Bland-Altman: bias + 95% limits of agreement.
    a = GPT, b = Gemini → bias = mean(a - b), positive = GPT larger."""
    diffs = [a - b for a, b in zip(values_a, values_b)]
    means = [(a + b) / 2 for a, b in zip(values_a, values_b)]
    if not diffs:
        return {"n": 0, "bias": float("nan"), "sd_diff": float("nan"),
                "loa_low": float("nan"), "loa_high": float("nan"),
                "mean_mean": float("nan")}
    bias = sum(diffs) / len(diffs)
    if len(diffs) > 1:
        sd = statistics.stdev(diffs)
    else:
        sd = 0.0
    return {"n": len(diffs), "bias": bias, "sd_diff": sd,
            "loa_low": bias - 1.96 * sd, "loa_high": bias + 1.96 * sd,
            "mean_mean": sum(means) / len(means)}


def cohen_d_paired(deltas: list[float]) -> float:
    """Cohen's d for paired differences = mean(delta) / sd(delta)."""
    if len(deltas) < 2:
        return float("nan")
    m = sum(deltas) / len(deltas)
    sd = statistics.stdev(deltas)
    return m / sd if sd > 0 else float("nan")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--gpt-sandbox", default="results_consensus")
    ap.add_argument("--gemini-sandbox", default="results_full_gemini")
    ap.add_argument("--reference", default="consensus_gt",
                    choices=["consensus_gt", "omfr_1", "omfr_2", "student"])
    args = ap.parse_args()

    gpt_path = (ROOT / args.gpt_sandbox / "full_run_records.pkl").resolve()
    gem_path = (ROOT / args.gemini_sandbox / "full_run_records.pkl").resolve()
    gpt_records = pickle.load(open(gpt_path, "rb"))
    gem_records = pickle.load(open(gem_path, "rb"))
    print(f"Loaded {len(gpt_records)} GPT records ← {gpt_path}")
    print(f"Loaded {len(gem_records)} Gemini records ← {gem_path}")
    print(f"Reference: {args.reference!r}")
    ref = args.reference

    # Index by (query_id, strategy)
    gpt_by = {(r["query_id"], r["strategy"]): r for r in gpt_records}
    gem_by = {(r["query_id"], r["strategy"]): r for r in gem_records}
    common_keys = sorted(set(gpt_by.keys()) & set(gem_by.keys()))
    print(f"  Common (query_id × strategy) keys: {len(common_keys)} "
          f"(expected 1800 = 900 q × 2 strat)")
    if len(common_keys) != 1800:
        print(f"  WARNING: missing pairs. GPT only={len(set(gpt_by)-set(gem_by))}, "
              f"Gemini only={len(set(gem_by)-set(gpt_by))}")

    # ── RQ-X1: Modality × strategy paired comparison ─────────────────
    rq_modality_strategy = {}
    for mod in ("CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"):
        for strat in STRATEGIES:
            for ltype in ("point", "area"):
                if mod != "PANORAMIC" and ltype == "area":
                    continue
                # For point: use mean_ed (lower=better)
                # For area: use mean_jaccard (higher=better)
                metric_key = "mean_ed" if ltype == "point" else "mean_jaccard"
                deltas: list[float] = []
                gpt_vals: list[float] = []
                gem_vals: list[float] = []
                skipped = 0
                qids_used: list[str] = []
                for k in common_keys:
                    g = gpt_by[k]
                    e = gem_by[k]
                    if (g["modality"] != mod or g["strategy"] != strat
                            or g["landmark_type"] != ltype):
                        continue
                    gv = g["metrics"][ref].get(metric_key)
                    ev = e["metrics"][ref].get(metric_key)
                    if gv is None or ev is None:
                        skipped += 1
                        continue
                    gpt_vals.append(gv)
                    gem_vals.append(ev)
                    deltas.append(gv - ev)  # ED: positive = Gemini better; Jaccard: positive = GPT better
                    qids_used.append(k[0])
                if not deltas:
                    continue
                w = paired_wilcoxon(deltas)
                ba = bland_altman(gpt_vals, gem_vals)
                d = cohen_d_paired(deltas)
                key = f"{mod}_{strat}_{ltype}"
                rq_modality_strategy[key] = {
                    "metric": metric_key,
                    "sign_convention": ("delta = GPT - Gemini; "
                                         "for ED, positive delta means Gemini better; "
                                         "for Jaccard, positive delta means GPT better"),
                    "n_paired": len(deltas),
                    "n_skipped_either_failed": skipped,
                    "gpt_mean": sum(gpt_vals) / len(gpt_vals),
                    "gpt_median": statistics.median(gpt_vals),
                    "gemini_mean": sum(gem_vals) / len(gem_vals),
                    "gemini_median": statistics.median(gem_vals),
                    **w,
                    "cohen_d_paired": d,
                    "bland_altman": ba,
                    # Bonferroni: 4 strategy×modality×type families (3 point-modalities
                    # + 1 area in PAN) × 2 strategies = 8 tests per family-row group
                    "bonferroni_n_tests_modality_family": 8,
                    "p_bonferroni_modality_family": min(1.0, w["p"] * 8) if not math.isnan(w["p"]) else float("nan"),
                }

    # ── RQ-X2: Per-landmark paired comparison ────────────────────────
    rq_per_landmark = {}
    n_landmark_tests = len(POINT_LANDMARKS_ORDER) + len(AREA_LANDMARKS)  # 12
    n_total_landmark_tests = n_landmark_tests * 2  # × 2 strategies = 24
    for mod, lm in POINT_LANDMARKS_ORDER + [("PANORAMIC", a) for a in AREA_LANDMARKS]:
        ltype = "area" if lm in AREA_LANDMARKS else "point"
        metric_key = "mean_ed" if ltype == "point" else "mean_jaccard"
        for strat in STRATEGIES:
            deltas = []
            gpt_vals = []
            gem_vals = []
            skipped = 0
            for k in common_keys:
                g = gpt_by[k]
                e = gem_by[k]
                if (g["modality"] != mod or g["structure"] != lm
                        or g["strategy"] != strat):
                    continue
                gv = g["metrics"][ref].get(metric_key)
                ev = e["metrics"][ref].get(metric_key)
                if gv is None or ev is None:
                    skipped += 1
                    continue
                gpt_vals.append(gv)
                gem_vals.append(ev)
                deltas.append(gv - ev)
            if not deltas:
                continue
            w = paired_wilcoxon(deltas)
            ba = bland_altman(gpt_vals, gem_vals)
            d = cohen_d_paired(deltas)
            key = f"{mod}/{lm}/{strat}"
            rq_per_landmark[key] = {
                "metric": metric_key,
                "landmark_type": ltype,
                "n_paired": len(deltas),
                "n_skipped_either_failed": skipped,
                "gpt_mean": sum(gpt_vals) / len(gpt_vals),
                "gpt_median": statistics.median(gpt_vals),
                "gemini_mean": sum(gem_vals) / len(gem_vals),
                "gemini_median": statistics.median(gem_vals),
                **w,
                "cohen_d_paired": d,
                "bland_altman": ba,
                "bonferroni_n_tests_landmark_family": n_total_landmark_tests,
                "p_bonferroni_landmark_family": min(1.0, w["p"] * n_total_landmark_tests) if not math.isnan(w["p"]) else float("nan"),
            }

    # ── F5/F6 attractor check: per-image cell collapse rate ──────────
    # Specifically for PANORAMIC Tooth_33_Apex (GPT-only failure mode in v4).
    f56_table = {}
    for strat in STRATEGIES:
        for label, recs_by_strat in [("GPT", gpt_by), ("Gemini", gem_by)]:
            cells_per_rep: list[str] = []
            gt_match_per_rep: list[bool] = []
            for k in common_keys:
                rec = recs_by_strat[k]
                if (rec["modality"] != "PANORAMIC" or rec["structure"] != "Tooth_33_Apex"
                        or rec["strategy"] != strat):
                    continue
                for rep_idx in range(3):
                    pred = rec["rep_pred_cells"][rep_idx]
                    if not pred:
                        continue
                    col, row = pred[0]
                    cell = f"{chr(ord('A') + (row - 1))}{col}"
                    cells_per_rep.append(cell)
                    # GT cell parsed from the consensus_gt string
                    gt_str = (rec.get("consensus_gt") or "").strip()
                    gt_match_per_rep.append(cell == gt_str)
            n = len(cells_per_rep)
            if n == 0:
                continue
            n_f56 = sum(1 for c in cells_per_rep if c in ("F5", "F6"))
            n_exact = sum(gt_match_per_rep)
            f56_table[f"{label}/Tooth_33_Apex/{strat}"] = {
                "n_predictions_with_valid_cell": n,
                "n_in_F5_or_F6": n_f56,
                "frac_F5_F6": n_f56 / n,
                "n_exact_match_consensus": n_exact,
                "frac_exact_match": n_exact / n,
            }

    output = {
        "reference_used": ref,
        "gpt_sandbox": str(gpt_path),
        "gemini_sandbox": str(gem_path),
        "n_common_keys": len(common_keys),
        "sign_convention_note": (
            "Delta is computed as GPT_metric - GEMINI_metric for ED and Jaccard. "
            "For ED (lower = better): positive delta → Gemini better. "
            "For Jaccard (higher = better): positive delta → GPT better. "
            "Reader should consult the 'metric' field of each row to interpret."
        ),
        "RQ_modality_strategy": rq_modality_strategy,
        "RQ_per_landmark": rq_per_landmark,
        "F5_F6_attractor": f56_table,
    }
    out_path = ROOT / args.gemini_sandbox / "gpt_vs_gemini.json"
    out_path.write_text(json.dumps(output, indent=2, default=str))
    print(f"Wrote {out_path}")

    # ── Summary print ────────────────────────────────────────────────
    print()
    print("=== Top-line modality×strategy paired Wilcoxon (GPT − Gemini, against consensus_gt) ===")
    print(f"{'Bucket':<40} {'metric':<13} {'n':>4} {'GPT mean':>10} {'Gem mean':>10} {'Δ mean':>10} {'p':>9} {'r':>6}")
    for k, v in sorted(rq_modality_strategy.items()):
        print(f"  {k:<40} {v['metric']:<13} {v['n_paired']:>4} "
              f"{v['gpt_mean']:>10.4f} {v['gemini_mean']:>10.4f} "
              f"{v['mean_delta']:>+10.4f} {v['p']:>9.2e} {v['rank_biserial_r']:>+6.3f}")

    print()
    print("=== F5/F6 attractor check (Tooth_33_Apex guided/zero_shot) ===")
    for k, v in sorted(f56_table.items()):
        print(f"  {k:<40} {v['n_in_F5_or_F6']}/{v['n_predictions_with_valid_cell']} "
              f"= {v['frac_F5_F6']*100:.1f}% on F5/F6  | "
              f"exact_match {v['frac_exact_match']*100:.1f}%")


if __name__ == "__main__":
    main()
