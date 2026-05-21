"""
Phase 3b — inter- and intra-rater reliability for the GT validation section.

Inter-rater pairs (n = 900 each):
  - OMFR_1 vs OMFR_2
  - OMFR_1 vs CONSENSUS
  - OMFR_2 vs CONSENSUS
  - STUDENT vs CONSENSUS
Intra-rater pairs (n = 180 each):
  - OMFR_1 vs OMFR_1_Second
  - OMFR_2 vs OMFR_2_Second

For point landmarks (n = 600 = 300 PAN + 150 PA + 150 CEPH per pair):
  - Mean Euclidean distance between raters
  - Cohen's κ (linear/quadratic if cells encoded as ordinal? unweighted here on cell-id)
  - SDR@1 (rate of within-1-cell agreement)
For area landmarks (PAN only, n = 300):
  - Mean Jaccard index
  - Mean Dice coefficient
  - "Strict equal" rate (cell sets identical)

All output is computed from results_consensus/query_index.json. Reads only.
Writes results_consensus/rater_reliability.json.

Usage: .venv/bin/python scripts/analyze_rater_reliability.py
"""
from __future__ import annotations

import argparse
import json
import math
import statistics
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import config  # noqa: E402


# ── Geometry ────────────────────────────────────────────────────────

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
    try:
        col = int(colstr)
    except ValueError:
        return None
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


def dice(a, b):
    n = len(a) + len(b)
    return (2 * len(a & b)) / n if n else 1.0


# ── Cohen's κ on cell categories (point landmarks) ─────────────────

def cohens_kappa(items):
    """items: list of (cat_a, cat_b) where cats are hashable.
    Returns unweighted Cohen's κ."""
    n = len(items)
    if n == 0:
        return None
    cats = sorted({c for pair in items for c in pair})
    if len(cats) <= 1:
        return 1.0  # all raters agree on the only category
    cnt_a = defaultdict(int)
    cnt_b = defaultdict(int)
    agree = 0
    for a, b in items:
        cnt_a[a] += 1
        cnt_b[b] += 1
        if a == b:
            agree += 1
    P_o = agree / n
    P_e = sum((cnt_a[c] / n) * (cnt_b[c] / n) for c in cats)
    if 1 - P_e == 0:
        return 1.0
    return (P_o - P_e) / (1 - P_e)


# ── Pairwise reliability computation ────────────────────────────────

def reliability_for_pair(queries, field_a, field_b, *, restrict_to=None):
    """Compute reliability metrics between two reference fields across queries.

    restrict_to: optional callable q→bool to filter queries (e.g. q for which
    field_b is non-null, used for intra-rater subsets).

    Returns dict keyed by (modality, landmark_type) and an "ALL" aggregate.
    """
    by_grp = defaultdict(list)  # (mod, ltype) → list of records
    for q in queries:
        if restrict_to is not None and not restrict_to(q):
            continue
        a_text = q.get(field_a)
        b_text = q.get(field_b)
        if a_text is None or b_text is None:
            continue
        mod = q["sheet"]
        ltype = q["landmark_type"]
        a_cells = parse_cells(a_text, mod)
        b_cells = parse_cells(b_text, mod)
        if not a_cells or not b_cells:
            continue
        by_grp[(mod, ltype)].append((q, a_cells, b_cells))

    result = {}
    all_point_eds = []
    all_area_jaccs = []
    all_area_dices = []
    all_point_pairs = []  # for grand κ
    all_point_within1 = []  # SDR@1 across pairs
    n_total = 0

    for (mod, ltype), items in by_grp.items():
        n = len(items)
        n_total += n
        if ltype == "point":
            eds = [euclid(a[0], b[0]) for _, a, b in items]
            pairs = [(a[0], b[0]) for _, a, b in items]
            kappa = cohens_kappa(pairs)
            within1 = sum(1 for e in eds if e <= 1 + 1e-9) / n
            exact = sum(1 for e in eds if e == 0) / n
            result[f"{mod}_{ltype}"] = {
                "n": n,
                "mean_ed": sum(eds) / n,
                "median_ed": statistics.median(eds),
                "max_ed": max(eds),
                "exact_match_rate": exact,
                "within_1_cell_rate": within1,
                "cohens_kappa": kappa,
            }
            all_point_eds.extend(eds)
            all_point_pairs.extend(pairs)
            all_point_within1.extend(eds)
        else:  # area
            jaccs = [jaccard(set(a), set(b)) for _, a, b in items]
            dices = [dice(set(a), set(b)) for _, a, b in items]
            strict_eq = sum(1 for j in jaccs if j == 1.0) / n
            result[f"{mod}_{ltype}"] = {
                "n": n,
                "mean_jaccard": sum(jaccs) / n,
                "median_jaccard": statistics.median(jaccs),
                "mean_dice": sum(dices) / n,
                "median_dice": statistics.median(dices),
                "strict_equal_rate": strict_eq,
            }
            all_area_jaccs.extend(jaccs)
            all_area_dices.extend(dices)

    # ALL-point aggregate (using Cohen's κ across modalities — categories are mod-prefixed)
    if all_point_pairs:
        prefixed = []
        for (q, a_cells, b_cells), (a, b) in zip(
            ((items[0], items[1], items[2]) for items in []), all_point_pairs):
            pass  # placeholder so type-checker is happy; below uses items.

    # Easier: rebuild with mod-prefix for the grand κ
    grand_pairs_prefixed = []
    grand_eds_norm = []  # NED across modalities
    for (mod, ltype), items in by_grp.items():
        if ltype != "point":
            continue
        cols, rows = grid_dims(mod)
        diag = math.hypot(cols - 1, rows - 1)
        for q, a, b in items:
            grand_pairs_prefixed.append((f"{mod}_{a[0]}", f"{mod}_{b[0]}"))
            grand_eds_norm.append(euclid(a[0], b[0]) / diag)

    if grand_pairs_prefixed:
        result["ALL_point"] = {
            "n": len(grand_pairs_prefixed),
            "mean_ed_cells": sum(all_point_eds) / len(all_point_eds),
            "mean_ned": sum(grand_eds_norm) / len(grand_eds_norm),
            "exact_match_rate": sum(1 for e in all_point_eds if e == 0) / len(all_point_eds),
            "within_1_cell_rate": sum(1 for e in all_point_eds if e <= 1 + 1e-9) / len(all_point_eds),
            "cohens_kappa": cohens_kappa(grand_pairs_prefixed),
        }
    if all_area_jaccs:
        result["ALL_area"] = {
            "n": len(all_area_jaccs),
            "mean_jaccard": sum(all_area_jaccs) / len(all_area_jaccs),
            "mean_dice": sum(all_area_dices) / len(all_area_dices),
            "strict_equal_rate": sum(1 for j in all_area_jaccs if j == 1.0) / len(all_area_jaccs),
        }

    return result


def per_landmark_reliability(queries, field_a, field_b):
    """Same as reliability_for_pair but bucketed by (modality, structure)."""
    by_lm = defaultdict(list)
    for q in queries:
        a_text = q.get(field_a)
        b_text = q.get(field_b)
        if a_text is None or b_text is None:
            continue
        mod = q["sheet"]
        ltype = q["landmark_type"]
        a_cells = parse_cells(a_text, mod)
        b_cells = parse_cells(b_text, mod)
        if not a_cells or not b_cells:
            continue
        by_lm[(mod, q["structure"], ltype)].append((a_cells, b_cells))
    result = {}
    for (mod, lm, ltype), items in by_lm.items():
        n = len(items)
        if ltype == "point":
            eds = [euclid(a[0], b[0]) for a, b in items]
            within1 = sum(1 for e in eds if e <= 1 + 1e-9) / n
            result[f"{mod}/{lm}"] = {
                "n": n, "type": "point",
                "mean_ed": sum(eds) / n,
                "exact_match_rate": sum(1 for e in eds if e == 0) / n,
                "within_1_cell_rate": within1,
            }
        else:
            jaccs = [jaccard(set(a), set(b)) for a, b in items]
            dices = [dice(set(a), set(b)) for a, b in items]
            result[f"{mod}/{lm}"] = {
                "n": n, "type": "area",
                "mean_jaccard": sum(jaccs) / n,
                "mean_dice": sum(dices) / n,
            }
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sandbox", default="results_consensus")
    args = ap.parse_args()
    sandbox = (ROOT / args.sandbox).resolve()

    queries = json.loads((sandbox / "query_index.json").read_text())
    print(f"Loaded {len(queries)} queries")

    pairs_inter = [
        ("omfr_1", "omfr_2", "INTER_omfr1_vs_omfr2"),
        ("omfr_1", "consensus_gt", "INTER_omfr1_vs_consensus"),
        ("omfr_2", "consensus_gt", "INTER_omfr2_vs_consensus"),
        ("student", "consensus_gt", "STUDENT_vs_consensus"),
    ]
    pairs_intra = [
        ("omfr_1", "omfr_1_second", "INTRA_omfr1"),
        ("omfr_2", "omfr_2_second", "INTRA_omfr2"),
    ]

    output = {"inter_rater": {}, "intra_rater": {},
              "inter_per_landmark": {}, "intra_per_landmark": {}}

    for a, b, label in pairs_inter:
        output["inter_rater"][label] = reliability_for_pair(queries, a, b)
        output["inter_per_landmark"][label] = per_landmark_reliability(queries, a, b)
    for a, b, label in pairs_intra:
        # Intra-rater: only on rows where field_b is populated
        output["intra_rater"][label] = reliability_for_pair(
            queries, a, b, restrict_to=lambda q, fb=b: q.get(fb) is not None)
        output["intra_per_landmark"][label] = per_landmark_reliability(queries, a, b)

    out_path = sandbox / "rater_reliability.json"
    out_path.write_text(json.dumps(output, indent=2, default=str))
    print(f"Wrote {out_path}")

    # Print top-line summary
    print("\n=== Inter-rater reliability summary ===")
    for label in ("INTER_omfr1_vs_omfr2", "INTER_omfr1_vs_consensus",
                   "INTER_omfr2_vs_consensus", "STUDENT_vs_consensus"):
        r = output["inter_rater"].get(label, {})
        ap_ = r.get("ALL_point", {})
        aa = r.get("ALL_area", {})
        print(f"\n  {label}:")
        if ap_:
            print(f"    POINT n={ap_['n']:4d}  mean_ED={ap_['mean_ed_cells']:.3f}  "
                  f"NED={ap_['mean_ned']:.4f}  SDR@1={ap_['within_1_cell_rate']*100:.1f}%  "
                  f"κ={ap_['cohens_kappa']:.3f}")
        if aa:
            print(f"    AREA  n={aa['n']:4d}  mean_J={aa['mean_jaccard']:.3f}  "
                  f"mean_D={aa['mean_dice']:.3f}  strict_eq={aa['strict_equal_rate']*100:.1f}%")

    print("\n=== Intra-rater reliability summary (180 doubly-rated queries) ===")
    for label in ("INTRA_omfr1", "INTRA_omfr2"):
        r = output["intra_rater"].get(label, {})
        ap_ = r.get("ALL_point", {})
        aa = r.get("ALL_area", {})
        print(f"\n  {label}:")
        if ap_:
            print(f"    POINT n={ap_['n']:4d}  mean_ED={ap_['mean_ed_cells']:.3f}  "
                  f"NED={ap_['mean_ned']:.4f}  SDR@1={ap_['within_1_cell_rate']*100:.1f}%  "
                  f"κ={ap_['cohens_kappa']:.3f}")
        if aa:
            print(f"    AREA  n={aa['n']:4d}  mean_J={aa['mean_jaccard']:.3f}  "
                  f"mean_D={aa['mean_dice']:.3f}  strict_eq={aa['strict_equal_rate']*100:.1f}%")


if __name__ == "__main__":
    main()
