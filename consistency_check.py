"""
Dental MLLM Benchmark — Consistency Check (Issue 1)

Runs Fleiss' kappa across N repeated runs (typically 3) of the same batch at
temperature = 0, to verify that the models produce deterministic output.
A kappa of 1.0 justifies reporting single-run evaluation in Methods.

Usage:
    python3 consistency_check.py <run1.json> <run2.json> <run3.json> [...]

Each input file is a parsed_responses.json produced by `pipeline.py parse`
from a separate submission (same queries, temperature = 0, different seed
runs). The new parsed_responses format is a list of records, each with
explicit query_id, strategy, and model_key fields. This script pairs
responses across runs by (query_id, strategy, model_key) and computes
Fleiss' kappa over the categorical outcome (parsed coordinate signature,
or a sentinel for unparseable / invalidated responses).
"""

import argparse
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path


def load_run(path):
    with open(path) as f:
        data = json.load(f)
    if isinstance(data, dict):
        raise SystemExit(
            f"ERROR: {path} is in the legacy dict format. Re-run "
            f"'pipeline.py parse' to regenerate the list-of-records format."
        )
    return data


def coord_signature(parsed, failure_category=None):
    """
    Convert parsed_coordinates into a canonical category string.

    Valid coordinates → "<coord>" or "<coord1>|<coord2>|..." (sorted for
    order-independent comparison on area landmarks).

    Invalid / unparseable responses → a sentinel that includes the failure
    category so different modes of failure do not falsely agree with one
    another. E.g., a "refusal" in one run and an "ambiguous" in another are
    NOT treated as identical.
    """
    if not parsed:
        return f"__FAIL__{failure_category or 'unknown'}"
    return "|".join(sorted(parsed))


def fleiss_kappa(category_matrix):
    """
    Fleiss' kappa for N items rated by n raters into k mutually exclusive categories.
    category_matrix: list of dicts {category: count}, one per item.
    Each item must have the same total count (n raters).
    """
    N = len(category_matrix)
    if N == 0:
        return float("nan")

    # Collect all categories
    all_categories = set()
    for row in category_matrix:
        all_categories.update(row.keys())
    categories = sorted(all_categories)
    k = len(categories)

    if k <= 1:
        # All items in one category → perfect agreement by construction
        return 1.0

    # Per-item number of raters (assumed constant)
    n_raters = sum(category_matrix[0].values())
    if n_raters < 2:
        return float("nan")

    # P_i: extent to which raters agree for item i
    P_items = []
    for row in category_matrix:
        total = sum(row.values())
        if total != n_raters:
            raise ValueError(
                f"Inconsistent rater count: expected {n_raters}, got {total}"
            )
        s = sum(c * (c - 1) for c in row.values())
        P_items.append(s / (n_raters * (n_raters - 1)))

    P_bar = sum(P_items) / N

    # p_j: proportion of all assignments in category j
    p_js = []
    total_assignments = N * n_raters
    for cat in categories:
        cj = sum(row.get(cat, 0) for row in category_matrix)
        p_js.append(cj / total_assignments)
    P_e = sum(p * p for p in p_js)

    if P_e == 1.0:
        return 1.0
    return (P_bar - P_e) / (1 - P_e)


def main():
    parser = argparse.ArgumentParser(
        description="Compute Fleiss' kappa across repeated parsed_responses.json runs"
    )
    parser.add_argument("runs", nargs="+", help="parsed_responses.json files (>=2)")
    parser.add_argument("--output", default=None,
                        help="Optional path to save JSON results")
    args = parser.parse_args()

    if len(args.runs) < 2:
        print("ERROR: need at least 2 runs to compute agreement.", file=sys.stderr)
        sys.exit(1)

    runs = [load_run(Path(p)) for p in args.runs]
    n_raters = len(runs)
    print(f"Loaded {n_raters} runs with "
          f"{', '.join(str(len(r)) for r in runs)} entries respectively.")

    # Index each run by (query_id, strategy, model_key). Warn on collisions
    # within a single run (should not happen in well-formed data).
    indexed_runs = []
    for i, run in enumerate(runs):
        idx = {}
        for rec in run:
            key = (rec.get("query_id"), rec.get("strategy"), rec.get("model_key"))
            if None in key:
                continue
            if key in idx:
                print(f"WARNING: run {i} has duplicate {key}", file=sys.stderr)
            idx[key] = rec
        indexed_runs.append(idx)

    common_keys = set(indexed_runs[0].keys())
    for idx in indexed_runs[1:]:
        common_keys &= set(idx.keys())
    common_keys = sorted(common_keys)
    print(f"Common (query_id, strategy, model_key) tuples: {len(common_keys)}")

    if not common_keys:
        print("ERROR: no overlapping items between runs.", file=sys.stderr)
        sys.exit(1)

    # Build category matrix: one row per item, counting categories across raters
    category_matrix = []
    exact_match_count = 0
    disagreements = []

    for key in common_keys:
        row = Counter()
        sigs = []
        for idx in indexed_runs:
            rec = idx[key]
            sig = coord_signature(
                rec.get("parsed_coordinates"),
                rec.get("failure_category"),
            )
            row[sig] += 1
            sigs.append(sig)
        category_matrix.append(dict(row))
        if len(set(sigs)) == 1:
            exact_match_count += 1
        else:
            disagreements.append({"key": list(key), "signatures": sigs})

    kappa = fleiss_kappa(category_matrix)
    exact_rate = exact_match_count / len(common_keys)

    print()
    print(f"Fleiss' kappa:         {kappa:.6f}")
    print(f"Exact agreement rate:  {exact_rate:.4f} "
          f"({exact_match_count}/{len(common_keys)} items)")
    print(f"Items with disagreement: {len(disagreements)}")

    if disagreements:
        print("\nFirst 10 disagreements:")
        for d in disagreements[:10]:
            print(f"  {d['key']}: {d['signatures']}")

    if abs(kappa - 1.0) < 1e-9:
        print("\n✅ kappa = 1.0 — single-run evaluation is justified.")
    else:
        print("\n⚠️  kappa < 1.0 — investigate non-determinism before single-run "
              "evaluation.")

    if args.output:
        result = {
            "n_runs": n_raters,
            "n_items": len(common_keys),
            "fleiss_kappa": kappa,
            "exact_match_count": exact_match_count,
            "exact_match_rate": exact_rate,
            "disagreements": disagreements,
        }
        with open(args.output, "w") as f:
            json.dump(result, f, indent=2)
        print(f"\nSaved: {args.output}")


if __name__ == "__main__":
    main()
