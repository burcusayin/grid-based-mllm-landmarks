"""
Phase 3c-gem — Gemini 3.1 Pro vs Student paired comparison + acceptability bands.

Mirror of scripts/analyze_gpt_vs_student.py but operating on the Gemini
records produced by scripts/recompute_gemini.py.

For each query, we have:
  - Gemini distance to consensus GT  (mean across 3 reps, per strategy)
  - Student distance to consensus GT  (single value, no reps)

The paired test compares GEMINI_ED and STUDENT_ED on the same query, against
the same consensus GT. The methodology mirrors gpt_vs_student.json exactly
so the Gemini-vs-student comparison is statistically apples-to-apples with
the GPT-vs-student comparison reported in §RQ5 of the paper.

Per strategy and per modality (and per landmark), we compute:
  - Paired Wilcoxon signed-rank test on (GEMINI - STUDENT)
  - Mean delta + median delta + rank-biserial r
  - Bland-Altman descriptive stats: mean bias + 95% limits of agreement
  - "Acceptability band" per landmark = max(human disagreement, 1.0 cell)

Bonferroni correction is applied downstream (×4 strategy×modality;
×9 point + ×3 area for per-landmark) at report-generation time.

Reads:
    results_full_gemini/full_run_records.pkl       — Gemini per-query records
    results_consensus/query_index.json             — canonical (has consensus_gt + student)
    results_consensus/rater_reliability.json       — for acceptability bands

Writes:
    results_full_gemini/gemini_vs_student.json

Usage:
    .venv/bin/python scripts/analyze_gemini_vs_student.py
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
AREA_LANDMARKS = ["Mandibular_Canal_L", "Maxillary_Sinus_R",
                   "External_Oblique_Ridge_R"]


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
    """Mirror of analyze_gpt_vs_student.py.paired_wilcoxon"""
    from scipy import stats
    nz = [d for d in deltas if abs(d) > 1e-12]
    n_total, n_nonzero = len(deltas), len(nz)
    if n_nonzero == 0:
        return {"n_total": n_total, "n_nonzero": 0, "stat": 0.0, "p": 1.0,
                "mean_delta": 0.0, "median_delta": 0.0,
                "rank_biserial_r": 0.0}
    try:
        res = stats.wilcoxon(nz, zero_method="wilcox",
                              alternative="two-sided", mode="auto")
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
    ap.add_argument("--gemini-sandbox", default="results_full_gemini")
    ap.add_argument("--query-index",
                    default="results_consensus/query_index.json",
                    help="Canonical query index that holds consensus_gt + student")
    ap.add_argument("--rater-reliability",
                    default="results_consensus/rater_reliability.json")
    args = ap.parse_args()

    sandbox = (ROOT / args.gemini_sandbox).resolve()
    records = pickle.load(open(sandbox / "full_run_records.pkl", "rb"))
    queries = json.loads((ROOT / args.query_index).read_text())
    qmap = {q["query_id"]: q for q in queries}

    gemini_vs_student_per_strategy = {}

    for strat in ("zero_shot", "guided"):
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
                gem_ed = r["metrics"]["consensus_gt"].get("mean_ed")
                if gem_ed is None:
                    continue
                stu_ed = euclid(student_cells[0], consensus_cells[0])
                per_mod_pt[mod].append((gem_ed, stu_ed, qid))
                per_lm_pt[(mod, r["structure"])].append((gem_ed, stu_ed, qid))
            else:
                gem_j = r["metrics"]["consensus_gt"].get("mean_jaccard")
                if gem_j is None:
                    continue
                stu_j = jaccard(set(student_cells), set(consensus_cells))
                per_mod_area[mod].append((gem_j, stu_j, qid))
                per_lm_area[(mod, r["structure"])].append((gem_j, stu_j, qid))

        out = {"per_modality": {}, "per_landmark": {}}

        for mod, items in per_mod_pt.items():
            gem = [g for g, _, _ in items]
            stu = [s for _, s, _ in items]
            deltas = [g - s for g, s in zip(gem, stu)]
            res = paired_wilcoxon(deltas)
            res["bland_altman"] = bland_altman(gem, stu)
            res["mean_gemini_ed"] = sum(gem) / len(gem)
            res["mean_student_ed"] = sum(stu) / len(stu)
            out["per_modality"][f"{mod}_point"] = res

        for mod, items in per_mod_area.items():
            gem = [g for g, _, _ in items]
            stu = [s for _, s, _ in items]
            deltas = [g - s for g, s in zip(gem, stu)]
            res = paired_wilcoxon(deltas)
            res["bland_altman"] = bland_altman(gem, stu)
            res["mean_gemini_jaccard"] = sum(gem) / len(gem)
            res["mean_student_jaccard"] = sum(stu) / len(stu)
            out["per_modality"][f"{mod}_area"] = res

        for (mod, lm), items in per_lm_pt.items():
            gem = [g for g, _, _ in items]
            stu = [s for _, s, _ in items]
            deltas = [g - s for g, s in zip(gem, stu)]
            res = paired_wilcoxon(deltas)
            res["mean_gemini_ed"] = sum(gem) / len(gem)
            res["mean_student_ed"] = sum(stu) / len(stu)
            out["per_landmark"][f"{mod}/{lm}"] = res

        for (mod, lm), items in per_lm_area.items():
            gem = [g for g, _, _ in items]
            stu = [s for _, s, _ in items]
            deltas = [g - s for g, s in zip(gem, stu)]
            res = paired_wilcoxon(deltas)
            res["mean_gemini_jaccard"] = sum(gem) / len(gem)
            res["mean_student_jaccard"] = sum(stu) / len(stu)
            out["per_landmark"][f"{mod}/{lm}"] = res

        gemini_vs_student_per_strategy[strat] = out

    # ── Acceptability band ─────────────────────────────────────────
    rater_reli = json.loads((ROOT / args.rater_reliability).read_text())
    per_lm_inter = rater_reli["inter_per_landmark"]["INTER_omfr1_vs_omfr2"]

    accept = {}
    for strat in ("zero_shot", "guided"):
        accept[strat] = {}
        for mod, lm in POINT_LANDMARKS_ORDER:
            key = f"{mod}/{lm}"
            band_info = per_lm_inter.get(key, {})
            mean_human_disagreement = band_info.get("mean_ed", 0.0)
            band = max(mean_human_disagreement, 1.0)
            gem_eds = [r["metrics"]["consensus_gt"]["mean_ed"] for r in records
                       if r["modality"] == mod and r["structure"] == lm
                       and r["strategy"] == strat
                       and r["metrics"]["consensus_gt"].get("mean_ed") is not None]
            if not gem_eds:
                continue
            within = sum(1 for e in gem_eds if e <= band + 1e-9)
            accept[strat][key] = {
                "n": len(gem_eds),
                "human_mean_disagreement_cells": mean_human_disagreement,
                "acceptability_band_cells": band,
                "within_band": within,
                "within_band_rate": within / len(gem_eds),
            }

    output = {
        "per_strategy": gemini_vs_student_per_strategy,
        "acceptability_band_per_landmark": accept,
        "reference": "consensus_gt",
        "model": "gemini-3.1-pro",
        "notes": (
            "Pairs constructed as (Gemini mean ED across 3 reps vs "
            "consensus, STUDENT single response vs consensus). Wilcoxon p "
            "= paired test on (GEMINI - STUDENT) deltas. Bonferroni "
            "correction applied downstream (×4 strategy×modality; "
            "×9 point + ×3 area for per-landmark) by the report generator. "
            "Acceptability band per landmark = max(mean OMFR_1↔OMFR_2 ED "
            "for that landmark, 1.0 cell)."
        ),
    }
    out_path = sandbox / "gemini_vs_student.json"
    out_path.write_text(json.dumps(output, indent=2, default=str))
    print(f"Wrote {out_path}")

    # Top-line print
    print("\n=== Gemini 3.1 Pro vs Student — modality-level paired tests ===")
    for strat in ("zero_shot", "guided"):
        print(f"\n  Strategy: {strat}")
        for mod_lt in ("CEPHALOMETRIC_point", "PERIAPICAL_point",
                       "PANORAMIC_point", "PANORAMIC_area"):
            d = gemini_vs_student_per_strategy[strat]["per_modality"].get(mod_lt)
            if not d:
                continue
            metric = "ED" if "point" in mod_lt else "Jaccard"
            mean_gem = d.get("mean_gemini_ed", d.get("mean_gemini_jaccard"))
            mean_stu = d.get("mean_student_ed", d.get("mean_student_jaccard"))
            print(f"    {mod_lt:22s} n={d['n_total']:4d}  "
                  f"GEM_{metric}={mean_gem:.3f}  "
                  f"STU_{metric}={mean_stu:.3f}  "
                  f"Δ={d['mean_delta']:+.3f}  "
                  f"p={d['p']:.4g}  "
                  f"r={d['rank_biserial_r']:+.3f}")

    print("\n=== Per-landmark Gemini vs Student (point landmarks only) ===")
    for strat in ("zero_shot", "guided"):
        print(f"\n  Strategy: {strat}")
        for mod, lm in POINT_LANDMARKS_ORDER:
            key = f"{mod}/{lm}"
            d = gemini_vs_student_per_strategy[strat]["per_landmark"].get(key)
            if not d:
                continue
            print(f"    {key:45s} n={d['n_total']:3d}  "
                  f"GEM={d.get('mean_gemini_ed', 0):.3f}  "
                  f"STU={d.get('mean_student_ed', 0):.3f}  "
                  f"Δ={d['mean_delta']:+.3f}  p={d['p']:.4g}  "
                  f"r={d['rank_biserial_r']:+.3f}")


if __name__ == "__main__":
    main()
