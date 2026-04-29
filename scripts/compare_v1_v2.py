"""
Paired v1 vs v2 comparison report.

Reads parsed_responses.json from results_pilot/run{1,2,3} (v1) and
results_pilot_v2/run{1,2,3} (v2). Computes per-query mean ED and
Jaccard across the 3 reps for each version, then performs a paired
analysis.

Usage:
    .venv/bin/python scripts/compare_v1_v2.py
    .venv/bin/python scripts/compare_v1_v2.py --self-test  # v1 vs v1 (sanity)

The --self-test mode runs the comparison with both halves pointing at
v1 (using runs 1+2 as "old" and run 3 as "new"). Used to catch bugs in
the comparison code before the real v2 run completes.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
from collections import defaultdict
from pathlib import Path
from statistics import mean, stdev

ROOT = Path(__file__).resolve().parent.parent


def coord_num(c: str) -> tuple[int, int] | None:
    m = re.match(r'([A-H])(\d+)', c.upper())
    return (ord(m.group(1)) - ord('A') + 1, int(m.group(2))) if m else None


def euclidean(c1: str, c2: str) -> float | None:
    a, b = coord_num(c1), coord_num(c2)
    if not a or not b:
        return None
    return math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2)


def jaccard(pred_cells: list[str] | None, gt_str: str) -> float | None:
    if not pred_cells:
        return 0.0
    gt_set = {c.strip().upper() for c in gt_str.split(",")}
    pred_set = {c.upper() for c in pred_cells}
    union = len(gt_set | pred_set)
    return (len(gt_set & pred_set) / union) if union else 0.0


def load_runs(base_dir: Path) -> dict[int, list[dict]]:
    out: dict[int, list[dict]] = {}
    for r in (1, 2, 3):
        path = base_dir / f"run{r}" / "parsed_responses.json"
        if path.exists():
            out[r] = json.loads(path.read_text())
    return out


def per_query_metric(runs: dict[int, list[dict]], qi: dict[str, dict]) -> dict[tuple, dict]:
    """
    Aggregate across runs: for each (query_id, strategy), compute the
    mean ED (point) or Jaccard (area) across the available runs.

    Returns {(query_id, strategy): {"metric": "ed"|"jaccard", "values": [...],
                                    "mean": float, "fail_count": int}}
    """
    bucket: dict[tuple, dict] = defaultdict(lambda: {"values": [], "fail_count": 0,
                                                      "raw_responses": []})
    for r, parsed in runs.items():
        for p in parsed:
            qid = p["query_id"]
            strat = p["strategy"]
            q = qi[qid]
            key = (qid, strat)
            bucket[key]["query_type"] = q["landmark_type"]
            bucket[key]["sheet"] = q["sheet"]
            bucket[key]["raw_responses"].append(p.get("raw_response", ""))
            if p.get("failure_category"):
                bucket[key]["fail_count"] += 1
                continue
            parsed_coords = p.get("parsed_coordinates")
            if not parsed_coords:
                bucket[key]["fail_count"] += 1
                continue
            if q["landmark_type"] == "point":
                d = euclidean(parsed_coords[0], q["omfr_1"].split(",")[0].strip())
                if d is not None:
                    bucket[key]["values"].append(d)
                else:
                    bucket[key]["fail_count"] += 1
            else:
                j = jaccard(parsed_coords, q["omfr_1"])
                bucket[key]["values"].append(j)

    for k, v in bucket.items():
        v["metric"] = "ed" if v["query_type"] == "point" else "jaccard"
        v["mean"] = mean(v["values"]) if v["values"] else None
        v["n"] = len(v["values"])

    return bucket


def paired_wilcoxon(deltas: list[float]) -> dict | None:
    """Two-sided paired Wilcoxon signed-rank on a list of v2-v1 differences."""
    nz = [d for d in deltas if d != 0]
    if len(nz) < 5:
        return None
    try:
        from scipy.stats import wilcoxon
        stat, p = wilcoxon(nz, zero_method="wilcox")
        return {"n_nonzero": len(nz), "n_total": len(deltas),
                "median_delta": sorted(deltas)[len(deltas) // 2],
                "mean_delta": sum(deltas) / len(deltas),
                "stat": float(stat), "p_value": float(p)}
    except ImportError:
        return None


def compare(v1: dict[tuple, dict], v2: dict[tuple, dict],
            qi: dict[str, dict]) -> dict:
    """Build the paired comparison."""
    common = set(v1) & set(v2)
    rows = []
    for key in sorted(common):
        v1_data = v1[key]
        v2_data = v2[key]
        qid, strat = key
        sheet = v1_data["sheet"]
        ltype = v1_data["query_type"]

        # Did this query's prompt change between v1 and v2?
        affected = (sheet == "PANORAMIC" and strat == "guided")

        v1_mean = v1_data["mean"]
        v2_mean = v2_data["mean"]
        delta = (v2_mean - v1_mean) if (v1_mean is not None and v2_mean is not None) else None

        # Distinct response check
        v1_resps = sorted(set(r for r in v1_data["raw_responses"] if r))
        v2_resps = sorted(set(r for r in v2_data["raw_responses"] if r))
        identical = v1_resps == v2_resps

        rows.append({
            "query_id": qid,
            "strategy": strat,
            "sheet": sheet,
            "type": ltype,
            "affected_by_change": affected,
            "metric": v1_data["metric"],
            "v1_mean": v1_mean,
            "v2_mean": v2_mean,
            "delta": delta,
            "v1_responses": v1_resps,
            "v2_responses": v2_resps,
            "responses_identical": identical,
            "v1_fails": v1_data["fail_count"],
            "v2_fails": v2_data["fail_count"],
        })

    # Group analyses
    out = {"per_query": rows, "groups": {}, "regression_check": {}}
    groups = {
        "PANORAMIC_guided_point":   ("PANORAMIC", "guided", "point", True),
        "PANORAMIC_guided_area":    ("PANORAMIC", "guided", "area", True),
        "PANORAMIC_zeroshot_point": ("PANORAMIC", "zero_shot", "point", False),
        "PANORAMIC_zeroshot_area":  ("PANORAMIC", "zero_shot", "area", False),
        "PERIAPICAL_guided":        ("PERIAPICAL", "guided", None, False),
        "PERIAPICAL_zeroshot":      ("PERIAPICAL", "zero_shot", None, False),
        "CEPHALOMETRIC_guided":     ("CEPHALOMETRIC", "guided", None, False),
        "CEPHALOMETRIC_zeroshot":   ("CEPHALOMETRIC", "zero_shot", None, False),
    }
    for name, (sheet, strat, ltype, expected_change) in groups.items():
        deltas = []
        for r in rows:
            if r["sheet"] != sheet or r["strategy"] != strat:
                continue
            if ltype is not None and r["type"] != ltype:
                continue
            if r["delta"] is None:
                continue
            deltas.append(r["delta"])
        wil = paired_wilcoxon(deltas) if deltas else None
        out["groups"][name] = {
            "sheet": sheet, "strategy": strat, "type_filter": ltype,
            "expected_change": expected_change,
            "n": len(deltas),
            "mean_delta": (sum(deltas) / len(deltas)) if deltas else None,
            "wilcoxon": wil,
        }

    # Regression check: groups not expected to change should have ~zero deltas
    for name, g in out["groups"].items():
        if g["expected_change"]: continue
        # Flag if any query in this group changed
        n_changed = sum(1 for r in rows
                        if not r["responses_identical"]
                        and r["sheet"] == g["sheet"]
                        and r["strategy"] == g["strategy"]
                        and (g["type_filter"] is None or r["type"] == g["type_filter"]))
        out["regression_check"][name] = {"n_responses_changed": n_changed,
                                          "warning": n_changed > 0}

    # Lateralization check: count "right side" landmarks where v2 vs v1 sided
    out["lateralization_check"] = analyze_lateralization(rows, qi)

    return out


def analyze_lateralization(rows: list[dict], qi: dict[str, dict]) -> dict:
    """For PANORAMIC guided point landmarks with explicit L/R in the structure
    name, check whether v2 responses are on the correct side of the grid more
    often than v1. The 'correct' side is where the ground-truth column is."""
    lat_data = []
    for r in rows:
        if not r["affected_by_change"]: continue
        if r["type"] != "point": continue
        q = qi[r["query_id"]]
        struct = q["structure"]
        if not (struct.endswith("_R") or struct.endswith("_L")):
            continue
        gt_coord = q["omfr_1"].split(",")[0].strip()
        gt_col = coord_num(gt_coord)
        if not gt_col: continue
        gt_col_num = gt_col[1]
        midpoint = 8  # PAN has 16 cols → midpoint between col 8 and 9

        def sided(resps: list[str]) -> dict:
            same_side = 0
            opp_side = 0
            for resp in resps:
                m = re.match(r'([A-H])(\d+)', resp.upper())
                if not m: continue
                resp_col = int(m.group(2))
                if (gt_col_num <= midpoint) == (resp_col <= midpoint):
                    same_side += 1
                else:
                    opp_side += 1
            return {"same_side": same_side, "opp_side": opp_side}

        lat_data.append({
            "query_id": r["query_id"],
            "structure": struct,
            "gt_side": "left" if gt_col_num <= midpoint else "right",
            "v1": sided(r["v1_responses"]),
            "v2": sided(r["v2_responses"]),
        })
    # Aggregate
    v1_correct = sum(d["v1"]["same_side"] for d in lat_data)
    v1_total = sum(d["v1"]["same_side"] + d["v1"]["opp_side"] for d in lat_data)
    v2_correct = sum(d["v2"]["same_side"] for d in lat_data)
    v2_total = sum(d["v2"]["same_side"] + d["v2"]["opp_side"] for d in lat_data)
    return {
        "n_queries_with_lateralization": len(lat_data),
        "v1_correct_side_rate": v1_correct / v1_total if v1_total else None,
        "v2_correct_side_rate": v2_correct / v2_total if v2_total else None,
        "v1_correct": v1_correct, "v1_total": v1_total,
        "v2_correct": v2_correct, "v2_total": v2_total,
        "per_query": lat_data,
    }


def print_report(result: dict) -> None:
    print("\n" + "=" * 78)
    print("v1 vs v2 COMPARISON REPORT")
    print("=" * 78)

    # 1. Group-level results
    print("\n--- Per-group mean ED/Jaccard delta (v2 - v1) ---")
    print(f"  {'Group':<30s} {'n':>4s} {'mean Δ':>10s} {'p-value':>10s} {'expected'}")
    for name, g in result["groups"].items():
        wil = g.get("wilcoxon")
        p_str = f"{wil['p_value']:.3f}" if wil else "n/a"
        delta_str = f"{g['mean_delta']:+.3f}" if g['mean_delta'] is not None else "n/a"
        ec = "→ should improve" if g["expected_change"] else "should ≈ 0 (regression)"
        print(f"  {name:<30s} {g['n']:>4d} {delta_str:>10s} {p_str:>10s}   {ec}")

    # 2. Regression checks
    print("\n--- Regression checks (these groups should be ~unchanged) ---")
    for name, rc in result["regression_check"].items():
        marker = "⚠ " if rc["warning"] else "✓ "
        print(f"  {marker}{name}: {rc['n_responses_changed']} queries with changed responses")

    # 3. Lateralization
    lat = result["lateralization_check"]
    print(f"\n--- Lateralization check (PANORAMIC guided, L/R landmarks) ---")
    print(f"  queries analyzed: {lat['n_queries_with_lateralization']}")
    if lat["v1_total"]:
        print(f"  v1 correct-side rate: {lat['v1_correct_side_rate']:.3f} "
              f"({lat['v1_correct']}/{lat['v1_total']})")
        print(f"  v2 correct-side rate: {lat['v2_correct_side_rate']:.3f} "
              f"({lat['v2_correct']}/{lat['v2_total']})")
        if lat["v2_correct_side_rate"] is not None and lat["v1_correct_side_rate"] is not None:
            improvement = lat["v2_correct_side_rate"] - lat["v1_correct_side_rate"]
            print(f"  improvement: {improvement:+.3f}")

    # 4. Sample of biggest deltas
    print("\n--- Largest deltas (top 10 by |delta|) ---")
    rows = [r for r in result["per_query"] if r["delta"] is not None]
    rows.sort(key=lambda r: -abs(r["delta"]))
    print(f"  {'query_id':<42s} {'strat':<10s} {'metric':<8s} "
          f"{'v1':>7s} {'v2':>7s} {'delta':>9s} {'affected'}")
    for r in rows[:10]:
        marker = "★" if r["affected_by_change"] else " "
        print(f"  {marker} {r['query_id']:<40s} {r['strategy']:<10s} {r['metric']:<8s} "
              f"{r['v1_mean']:>7.2f} {r['v2_mean']:>7.2f} {r['delta']:>+9.2f}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--v1", default="results_pilot",
                    help="Path to v1 sandbox (default results_pilot)")
    ap.add_argument("--v2", default="results_pilot_v2",
                    help="Path to v2 sandbox (default results_pilot_v2)")
    ap.add_argument("--out", default=None,
                    help="Optional path to write JSON report")
    ap.add_argument("--self-test", action="store_true",
                    help="Use v1 runs 1+2 as 'old' and run 3 as 'new' "
                         "(sanity check that comparison code works)")
    args = ap.parse_args()

    v1_path = ROOT / args.v1
    v2_path = ROOT / args.v2

    qi_path = v1_path / "query_index.json"
    qi = {q["query_id"]: q for q in json.loads(qi_path.read_text())}

    v1_runs = load_runs(v1_path)
    if args.self_test:
        # Split v1 runs: 1+2 vs 3 — should show near-zero deltas everywhere
        old_runs = {1: v1_runs[1], 2: v1_runs[2]}
        new_runs = {3: v1_runs[3]}
        print(f"SELF-TEST MODE: comparing v1 runs 1+2 vs run 3 (no v2 needed)")
    else:
        v2_runs = load_runs(v2_path)
        if not v2_runs:
            print(f"v2 has no runs yet at {v2_path} — has the v2 pilot completed?")
            return 1
        old_runs = v1_runs
        new_runs = v2_runs

    v1_metrics = per_query_metric(old_runs, qi)
    v2_metrics = per_query_metric(new_runs, qi)

    result = compare(v1_metrics, v2_metrics, qi)
    print_report(result)

    if args.out:
        Path(args.out).write_text(json.dumps(result, indent=2, default=str))
        print(f"\nFull report saved to {args.out}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
