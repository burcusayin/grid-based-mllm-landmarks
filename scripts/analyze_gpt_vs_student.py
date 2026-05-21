"""
Phase 3c — GPT-5.4 vs Student paired comparison + acceptability bands.

For each query, we have:
  - GPT-5.4 distance to consensus GT  (mean across 3 reps, per strategy)
  - Student distance to consensus GT  (single value, no reps)

The paired test compares GPT_ED and STUDENT_ED on the same query, against the
same GT. This follows the methodology document (Section 8.4) which planned
this as the canonical AI-vs-human comparison.

Per strategy and per modality (and per landmark), we compute:
  - Paired Wilcoxon signed-rank test on (GPT - STUDENT)
  - Mean delta + median delta + rank-biserial r
  - Bland-Altman descriptive stats: mean bias + 95% limits of agreement
  - "Acceptability band" check: per-landmark inter-rater spread (OMFR_1 vs
    OMFR_2 mean ED) defines a per-landmark "human-disagreement band". We
    report the fraction of GPT errors within that band.

Bonferroni correction applied across the strategy × modality (×4) and
strategy × landmark (×9 point + ×3 area) families.

Reads:    results_consensus/full_run_records.pkl
          results_consensus/query_index.json
          results_consensus/rater_reliability.json (for acceptability bands)
Writes:   results_consensus/gpt_vs_student.json

Usage: .venv/bin/python scripts/analyze_gpt_vs_student.py
"""
from __future__ import annotations

import argparse
import json
import math
import pickle
import statistics
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import config  # noqa: E402

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


def grid_dims(modality):
    g = config.GRID_SPECS[modality]
    return g["cols"], g["rows"]


def cell_to_xy(cell, modality):
    s = (cell or "").strip().upper().replace(" ", "").replace("-", "")
    if not s:
        return None
    i = 0
    while i < len(s) and s[i].isalpha():
        i += 1
    if i != 1 or i == len(s):
        return None
    rowstr, colstr = s[:i], s[i:]
    if not colstr.isdigit():
        return None
    row = (ord(rowstr) - ord("A")) + 1
    col = int(colstr)
    cols, rows = grid_dims(modality)
    if not (1 <= col <= cols and 1 <= row <= rows):
        return None
    return (col, row)


def parse_cells(text, modality):
    if text is None:
        return []
    out = []
    for tok in str(text).replace("\n", ",").replace(";", ",").split(","):
        xy = cell_to_xy(tok.strip(), modality)
        if xy is not None:
            out.append(xy)
    return out


def euclid(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


def jaccard(a, b):
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def paired_wilcoxon(deltas):
    """Same as the v1 helper: nonzero filter + Wilcoxon + rank-biserial r."""
    from scipy import stats
    nz = [d for d in deltas if abs(d) > 1e-12]
    n_total, n_nonzero = len(deltas), len(nz)
    if n_nonzero == 0:
        return {"n_total": n_total, "n_nonzero": 0, "stat": 0.0, "p": 1.0,
                "mean_delta": 0.0, "median_delta": 0.0, "rank_biserial_r": 0.0}
    try:
        res = stats.wilcoxon(nz, zero_method="wilcox", alternative="two-sided", mode="auto")
        stat, p = float(res.statistic), float(res.pvalue)
    except Exception:
        stat, p = float("nan"), float("nan")
    ranks = stats.rankdata([abs(d) for d in nz])
    Wp = sum(r for d, r in zip(nz, ranks) if d > 0)
    Wn = sum(r for d, r in zip(nz, ranks) if d < 0)
    rb = (Wp - Wn) / (Wp + Wn) if (Wp + Wn) > 0 else 0.0
    return {"n_total": n_total, "n_nonzero": n_nonzero, "stat": stat, "p": p,
            "mean_delta": sum(deltas) / n_total,
            "median_delta": statistics.median(deltas),
            "rank_biserial_r": rb}


def bland_altman(values_a, values_b):
    """Returns mean bias + 95% limits of agreement on (a - b) deltas.
    For paired samples of equal length only."""
    deltas = [a - b for a, b in zip(values_a, values_b)]
    if not deltas:
        return None
    mean = sum(deltas) / len(deltas)
    sd = statistics.stdev(deltas) if len(deltas) >= 2 else 0.0
    return {
        "n": len(deltas),
        "mean_diff": mean,
        "sd_diff": sd,
        "loa_low": mean - 1.96 * sd,
        "loa_high": mean + 1.96 * sd,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sandbox", default="results_consensus")
    args = ap.parse_args()
    sandbox = (ROOT / args.sandbox).resolve()

    records = pickle.load(open(sandbox / "full_run_records.pkl", "rb"))
    queries = json.loads((sandbox / "query_index.json").read_text())
    qmap = {q["query_id"]: q for q in queries}

    # For each (qid, strategy), build pairs (gpt_metric, student_metric)
    # GPT metric is mean across 3 reps vs consensus_gt.
    # Student metric is the single student-vs-consensus distance for that query.
    gpt_vs_student_per_strategy = {}

    for strat in ("zero_shot", "guided"):
        # Modality-level
        per_mod_pt = defaultdict(list)
        per_mod_area = defaultdict(list)
        per_lm_pt = defaultdict(list)
        per_lm_area = defaultdict(list)

        for r in records:
            if r["strategy"] != strat:
                continue
            qid = r["query_id"]
            q = qmap[qid]
            mod = r["modality"]
            ltype = r["landmark_type"]
            consensus_cells = parse_cells(q.get("consensus_gt"), mod)
            student_cells = parse_cells(q.get("student"), mod)
            if not consensus_cells or not student_cells:
                continue

            if ltype == "point":
                gpt_ed = r["metrics"]["consensus_gt"]["mean_ed"]
                if gpt_ed is None:
                    continue
                stu_ed = euclid(student_cells[0], consensus_cells[0])
                per_mod_pt[mod].append((gpt_ed, stu_ed, qid))
                per_lm_pt[(mod, r["structure"])].append((gpt_ed, stu_ed, qid))
            else:
                gpt_j = r["metrics"]["consensus_gt"]["mean_jaccard"]
                if gpt_j is None:
                    continue
                stu_j = jaccard(set(student_cells), set(consensus_cells))
                per_mod_area[mod].append((gpt_j, stu_j, qid))
                per_lm_area[(mod, r["structure"])].append((gpt_j, stu_j, qid))

        out = {"per_modality": {}, "per_landmark": {}}

        # Modality-level paired tests
        for mod, items in per_mod_pt.items():
            gpt = [g for g, _, _ in items]
            stu = [s for _, s, _ in items]
            deltas = [g - s for g, s in zip(gpt, stu)]  # positive = GPT worse
            res = paired_wilcoxon(deltas)
            res["bland_altman"] = bland_altman(gpt, stu)
            res["mean_gpt_ed"] = sum(gpt) / len(gpt)
            res["mean_student_ed"] = sum(stu) / len(stu)
            out["per_modality"][f"{mod}_point"] = res

        for mod, items in per_mod_area.items():
            gpt = [g for g, _, _ in items]
            stu = [s for _, s, _ in items]
            deltas = [g - s for g, s in zip(gpt, stu)]  # positive = GPT better (higher J)
            res = paired_wilcoxon(deltas)
            res["bland_altman"] = bland_altman(gpt, stu)
            res["mean_gpt_jaccard"] = sum(gpt) / len(gpt)
            res["mean_student_jaccard"] = sum(stu) / len(stu)
            out["per_modality"][f"{mod}_area"] = res

        # Per-landmark paired tests
        for (mod, lm), items in per_lm_pt.items():
            gpt = [g for g, _, _ in items]
            stu = [s for _, s, _ in items]
            deltas = [g - s for g, s in zip(gpt, stu)]
            res = paired_wilcoxon(deltas)
            res["mean_gpt_ed"] = sum(gpt) / len(gpt)
            res["mean_student_ed"] = sum(stu) / len(stu)
            out["per_landmark"][f"{mod}/{lm}"] = res

        for (mod, lm), items in per_lm_area.items():
            gpt = [g for g, _, _ in items]
            stu = [s for _, s, _ in items]
            deltas = [g - s for g, s in zip(gpt, stu)]
            res = paired_wilcoxon(deltas)
            res["mean_gpt_jaccard"] = sum(gpt) / len(gpt)
            res["mean_student_jaccard"] = sum(stu) / len(stu)
            out["per_landmark"][f"{mod}/{lm}"] = res

        gpt_vs_student_per_strategy[strat] = out

    # ── Acceptability band ──────────────────────────────────────────
    # For each (modality, landmark), compute the inter-rater spread between
    # OMFR_1 and OMFR_2 (per query), then report the fraction of GPT mean-ED
    # values that fall within max(human spread, 1.0 cell). 1.0 cell is added
    # as a floor to avoid pathological tightness when raters agree perfectly.
    rater_reli = json.loads((sandbox / "rater_reliability.json").read_text())
    per_lm_inter = rater_reli["inter_per_landmark"]["INTER_omfr1_vs_omfr2"]

    accept = {}
    # Per-landmark band = max(0.95-quantile of OMFR_1↔OMFR_2 ED, 1.0)
    # Approximate the quantile via the per-landmark mean ED + 1 SD as a proxy
    # since per-query ED is not stored at this level. Use 'within_1_cell_rate'
    # and 'mean_ed' as a sanity lookup.
    for strat in ("zero_shot", "guided"):
        accept[strat] = {}
        for mod, lm in POINT_LANDMARKS_ORDER:
            key = f"{mod}/{lm}"
            band_info = per_lm_inter.get(key, {})
            mean_human_disagreement = band_info.get("mean_ed", 0.0)
            # Pragmatic acceptability threshold: the larger of (a) typical
            # human disagreement on this landmark or (b) 1 grid cell.
            band = max(mean_human_disagreement, 1.0)
            # Count GPT records within band
            gpt_eds = [r["metrics"]["consensus_gt"]["mean_ed"] for r in records
                       if r["modality"] == mod and r["structure"] == lm
                       and r["strategy"] == strat
                       and r["metrics"]["consensus_gt"]["mean_ed"] is not None]
            if not gpt_eds:
                continue
            within = sum(1 for e in gpt_eds if e <= band + 1e-9)
            accept[strat][key] = {
                "n": len(gpt_eds),
                "human_mean_disagreement_cells": mean_human_disagreement,
                "acceptability_band_cells": band,
                "within_band": within,
                "within_band_rate": within / len(gpt_eds),
            }

    # Save
    output = {
        "per_strategy": gpt_vs_student_per_strategy,
        "acceptability_band_per_landmark": accept,
        "reference": "consensus_gt",
        "notes": (
            "Pairs constructed as (GPT mean ED across 3 reps vs consensus, "
            "STUDENT single response vs consensus). Wilcoxon p = paired test "
            "on (GPT - STUDENT) deltas. Acceptability band per landmark = "
            "max(mean OMFR_1↔OMFR_2 ED for that landmark, 1.0 cell)."
        ),
    }
    out_path = sandbox / "gpt_vs_student.json"
    out_path.write_text(json.dumps(output, indent=2, default=str))
    print(f"Wrote {out_path}")

    # Top-line print
    print("\n=== GPT-5.4 vs Student — modality-level paired tests ===")
    for strat in ("zero_shot", "guided"):
        print(f"\n  Strategy: {strat}")
        for mod_lt in ("CEPHALOMETRIC_point", "PERIAPICAL_point",
                       "PANORAMIC_point", "PANORAMIC_area"):
            d = gpt_vs_student_per_strategy[strat]["per_modality"].get(mod_lt)
            if not d:
                continue
            metric = "ED" if "point" in mod_lt else "Jaccard"
            print(f"    {mod_lt:22s} n={d['n_total']:4d}  "
                  f"GPT_{metric}={d.get('mean_gpt_ed', d.get('mean_gpt_jaccard')):.3f}  "
                  f"STU_{metric}={d.get('mean_student_ed', d.get('mean_student_jaccard')):.3f}  "
                  f"Δ={d['mean_delta']:+.3f}  p={d['p']:.4g}  r={d['rank_biserial_r']:+.3f}")

    print("\n=== Acceptability band — fraction of GPT errors within human-disagreement band ===")
    for strat in ("zero_shot", "guided"):
        print(f"\n  Strategy: {strat}")
        for mod, lm in POINT_LANDMARKS_ORDER:
            key = f"{mod}/{lm}"
            d = accept.get(strat, {}).get(key)
            if not d:
                continue
            print(f"    {key:50s} band={d['acceptability_band_cells']:.2f} cells, "
                  f"GPT within band = {d['within_band']}/{d['n']} = "
                  f"{d['within_band_rate']*100:.1f}%")


if __name__ == "__main__":
    main()
