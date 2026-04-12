"""
Dental MLLM Benchmark — Analysis Pipeline

Computes metrics, runs statistical tests, and generates visualizations.

Usage:
    python3 analysis.py metrics     # Compute all metrics
    python3 analysis.py stats       # Run statistical tests
    python3 analysis.py visualize   # Generate plots
    python3 analysis.py report      # Generate summary tables (CSV/Excel)
    python3 analysis.py all         # Run everything
"""

import argparse
import json
import math
import os
import re
import statistics
import sys
from collections import defaultdict
from pathlib import Path

import config

# ============================================================
# Grid Coordinate Utilities
# ============================================================

def coord_to_numeric(coord):
    """
    Convert grid coordinate string to (row, col) numeric tuple.
    "B3" → (2, 3), "A1" → (1, 1), "H16" → (8, 16)
    """
    if not coord:
        return None
    coord = coord.strip().upper()
    match = re.match(r'^([A-H])(\d{1,2})$', coord)
    if not match:
        return None
    row = ord(match.group(1)) - ord('A') + 1
    col = int(match.group(2))
    return (row, col)


def parse_cell_set(coord_list):
    """
    Parse a list of coordinate strings into a set of (row, col) tuples.
    ["B3", "C4", "D5"] → {(2,3), (3,4), (4,5)}
    """
    cells = set()
    for coord in coord_list:
        num = coord_to_numeric(coord)
        if num:
            cells.add(num)
    return cells


def parse_ground_truth(gt_string, landmark_type):
    """
    Parse ground truth from the Excel OMFR column.
    Point: "G11" → ["G11"]
    Area: "D15, E14, E15, F14" → ["D15", "E14", "E15", "F14"]
    """
    if not gt_string:
        return None
    gt_string = str(gt_string).strip()
    parts = [p.strip() for p in gt_string.replace(";", ",").split(",")]
    parts = [p for p in parts if p]
    return parts


# ============================================================
# Metric Computation
# ============================================================

def euclidean_distance(pred_coord, gt_coord):
    """
    Compute Euclidean distance between two grid coordinates.
    Returns distance in grid cell units.
    """
    pred = coord_to_numeric(pred_coord)
    gt = coord_to_numeric(gt_coord)
    if not pred or not gt:
        return None
    return math.sqrt((pred[0] - gt[0])**2 + (pred[1] - gt[1])**2)


def normalized_euclidean_distance(ed, modality):
    """Normalize ED by grid diagonal for cross-modality comparison."""
    diagonal = config.GRID_SPECS[modality]["diagonal"]
    return ed / diagonal


def jaccard_index(pred_set, gt_set):
    """
    Compute Jaccard Index (IoU) between two cell sets.

    Returns None when the ground-truth set is empty (undefined denominator;
    should not occur in valid data but we refuse to invent a value).
    Returns 0.0 when the prediction is empty but the GT is not
    (a total miss).
    """
    if not gt_set:
        return None
    if not pred_set:
        return 0.0
    intersection = len(pred_set & gt_set)
    union = len(pred_set | gt_set)
    return intersection / union


def dice_coefficient(pred_set, gt_set):
    """
    Compute Dice Coefficient (F1) between two cell sets.

    Returns None when the ground-truth set is empty (undefined denominator).
    Returns 0.0 when prediction is empty but GT is not.
    """
    if not gt_set:
        return None
    if not pred_set:
        return 0.0
    intersection = len(pred_set & gt_set)
    total = len(pred_set) + len(gt_set)
    return 2 * intersection / total


def success_detection_rate(distances, threshold):
    """Compute SDR: % of distances ≤ threshold."""
    if not distances:
        return 0.0
    successes = sum(1 for d in distances if d is not None and d <= threshold)
    valid = sum(1 for d in distances if d is not None)
    return (successes / valid * 100) if valid > 0 else 0.0


# ============================================================
# Load Data
# ============================================================

def load_query_index():
    """Load the query index from prepare step."""
    index_path = config.RESULTS_DIR / "query_index.json"
    if not index_path.exists():
        print("ERROR: query_index.json not found. Run 'pipeline.py prepare' first.")
        sys.exit(1)
    with open(index_path) as f:
        return json.load(f)


def load_parsed_responses():
    """
    Load parsed API responses.

    New format: a JSON list of records, each with fields
      query_id, strategy, model_key, custom_id, raw_response,
      parsed_coordinates, failure_category, out_of_range, modality,
      landmark_type, source_file.
    """
    resp_path = config.RESULTS_DIR / "parsed_responses.json"
    if not resp_path.exists():
        print("ERROR: parsed_responses.json not found. Run 'pipeline.py parse' first.")
        sys.exit(1)
    with open(resp_path) as f:
        data = json.load(f)
    if isinstance(data, dict):
        # Legacy format — the old dict keyed by custom_id has known issues
        # (custom_id collision across providers). Refuse to use it.
        print("ERROR: parsed_responses.json is in the legacy dict format. "
              "Delete it and re-run 'pipeline.py parse' after downloading "
              "responses, to regenerate the list-of-records format.")
        sys.exit(1)
    return data


def load_ground_truth(queries):
    """
    Build a dict of query_id -> ground-truth coordinates.

    Preference order (per query): consensus_gt -> omfr_1 -> None.
    Reports how many queries fell back to omfr_1 so that the user notices
    when consensus annotations are incomplete.
    """
    gt = {}
    fallback_to_omfr1 = 0
    missing = 0
    for q in queries:
        raw = q.get("consensus_gt")
        source = "consensus"
        if raw is None or (isinstance(raw, str) and not raw.strip()):
            raw = q.get("omfr_1")
            source = "omfr_1"
            if raw is not None and not (isinstance(raw, str) and not raw.strip()):
                fallback_to_omfr1 += 1
        if raw is None or (isinstance(raw, str) and not raw.strip()):
            missing += 1
            continue
        parsed = parse_ground_truth(raw, q["landmark_type"])
        if parsed:
            gt[q["query_id"]] = {"coords": parsed, "source": source}
    if fallback_to_omfr1:
        print(f"NOTE: {fallback_to_omfr1}/{len(queries)} queries fall back to "
              f"omfr_1 (consensus_gt missing).")
    if missing:
        print(f"WARNING: {missing}/{len(queries)} queries have no ground truth "
              f"in either consensus_gt or omfr_1 — excluded from metrics.")
    return gt


def _stats_summary(values):
    """Return dict with n, mean, std (sample), median, min, max.
    Uses statistics module (correct median for even n)."""
    n = len(values)
    if n == 0:
        return {"n": 0, "mean": None, "std": None, "median": None,
                "min": None, "max": None}
    mean = statistics.fmean(values)
    std = statistics.stdev(values) if n > 1 else 0.0
    med = statistics.median(values)
    return {
        "n": n,
        "mean": round(mean, 4),
        "std": round(std, 4),
        "median": round(med, 4),
        "min": round(min(values), 4),
        "max": round(max(values), 4),
    }


# ============================================================
# Metrics Command
# ============================================================

def cmd_metrics(args):
    """Compute all metrics for all model responses against ground truth."""
    queries = load_query_index()
    responses = load_parsed_responses()
    ground_truth = load_ground_truth(queries)

    # Build query lookup
    query_lookup = {q["query_id"]: q for q in queries}

    # Group responses by (model_key, strategy) — the new format gives us
    # these fields directly, no fragile string parsing required.
    grouped = defaultdict(list)
    for rec in responses:
        model_key = rec.get("model_key")
        strategy = rec.get("strategy")
        if not model_key or not strategy:
            continue
        grouped[(model_key, strategy)].append(rec)

    all_metrics = {}

    for (model_key, strategy), recs in sorted(grouped.items()):
        batch_key = f"{model_key}_{strategy}"
        print(f"\nComputing metrics for: {batch_key} ({len(recs)} responses)")

        metrics_list = []

        for rec in recs:
            query_id = rec["query_id"]
            query = query_lookup.get(query_id)
            gt_entry = ground_truth.get(query_id)
            if not query or not gt_entry:
                continue
            gt_coords = gt_entry["coords"]

            modality = query["sheet"]
            lm_type = query["landmark_type"]
            pred_coords = rec.get("parsed_coordinates")  # pre-validated in parse step

            metric = {
                "query_id": query_id,
                "image_id": query["image_id"],
                "modality": modality,
                "structure": query["structure"],
                "landmark_type": lm_type,
                "ground_truth": gt_coords,
                "gt_source": gt_entry["source"],
                "prediction": pred_coords,
                "parseable": pred_coords is not None,
                "failure_category": rec.get("failure_category"),
            }

            if lm_type == "point":
                gt_coord = gt_coords[0] if gt_coords else None
                pred_coord = pred_coords[0] if pred_coords else None
                if gt_coord and pred_coord:
                    ed = euclidean_distance(pred_coord, gt_coord)
                    metric["euclidean_distance"] = ed
                    if ed is not None:
                        metric["normalized_ed"] = normalized_euclidean_distance(ed, modality)
                        metric["exact_match"] = ed == 0.0
                        metric["within_1"] = ed <= 1.0
                        metric["within_sqrt2"] = ed <= math.sqrt(2)
                        metric["within_2"] = ed <= 2.0
                else:
                    metric["euclidean_distance"] = None

            elif lm_type == "area":
                gt_set = parse_cell_set(gt_coords)
                pred_set = parse_cell_set(pred_coords) if pred_coords else set()
                metric["jaccard"] = jaccard_index(pred_set, gt_set)
                metric["dice"] = dice_coefficient(pred_set, gt_set)
                metric["gt_cells"] = len(gt_set)
                metric["pred_cells"] = len(pred_set)
                metric["intersection_cells"] = len(pred_set & gt_set)

            metrics_list.append(metric)

        all_metrics[batch_key] = metrics_list

        # Print summary
        point_metrics = [m for m in metrics_list if m["landmark_type"] == "point"]
        area_metrics = [m for m in metrics_list if m["landmark_type"] == "area"]

        if point_metrics:
            eds = [m["euclidean_distance"] for m in point_metrics if m.get("euclidean_distance") is not None]
            if eds:
                mean_ed = sum(eds) / len(eds)
                sdr0 = success_detection_rate(eds, 0)
                sdr1 = success_detection_rate(eds, 1)
                sdr2 = success_detection_rate(eds, 2)
                print(f"  Point landmarks ({len(eds)} queries):")
                print(f"    Mean ED: {mean_ed:.3f} | SDR@0: {sdr0:.1f}% | SDR@1: {sdr1:.1f}% | SDR@2: {sdr2:.1f}%")

        if area_metrics:
            jaccards = [m["jaccard"] for m in area_metrics if m.get("jaccard") is not None]
            dices = [m["dice"] for m in area_metrics if m.get("dice") is not None]
            if jaccards:
                print(f"  Area landmarks ({len(jaccards)} queries):")
                print(f"    Mean Jaccard: {statistics.fmean(jaccards):.3f} | "
                      f"Mean Dice: {statistics.fmean(dices):.3f}")

    # Save all metrics
    config.ANALYSIS_DIR.mkdir(parents=True, exist_ok=True)
    out_path = config.ANALYSIS_DIR / "all_metrics.json"
    with open(out_path, "w") as f:
        json.dump(all_metrics, f, indent=2)
    print(f"\nMetrics saved: {out_path}")

    return all_metrics


# ============================================================
# Inter-Rater Reliability (Issue H4)
# ============================================================

def icc_2_1(ratings):
    """
    Intraclass Correlation Coefficient ICC(2,1): two-way random effects,
    single measurement, absolute agreement.

    Input: ratings — a list of rows, each row being [rater_1, rater_2, ...].
    Returns a float in [-inf, 1] or None if the sample is too small.

    Formula (Shrout & Fleiss 1979):
        ICC(2,1) = (BMS - EMS) / (BMS + (k-1)*EMS + k*(JMS - EMS)/n)
    where
        BMS = between-subjects mean square
        JMS = between-raters (columns) mean square
        EMS = residual mean square
        n   = number of subjects (rows)
        k   = number of raters (cols)
    """
    n = len(ratings)
    if n < 2:
        return None
    k = len(ratings[0])
    if k < 2 or any(len(r) != k for r in ratings):
        return None

    grand_mean = sum(sum(r) for r in ratings) / (n * k)
    row_means = [sum(r) / k for r in ratings]
    col_means = [sum(ratings[i][j] for i in range(n)) / n for j in range(k)]

    # Sum of squares
    ss_rows = sum(k * (rm - grand_mean) ** 2 for rm in row_means)
    ss_cols = sum(n * (cm - grand_mean) ** 2 for cm in col_means)
    ss_total = sum((ratings[i][j] - grand_mean) ** 2 for i in range(n) for j in range(k))
    ss_error = ss_total - ss_rows - ss_cols

    # Degrees of freedom
    df_rows = n - 1
    df_cols = k - 1
    df_error = (n - 1) * (k - 1)
    if df_rows == 0 or df_cols == 0 or df_error == 0:
        return None

    bms = ss_rows / df_rows
    jms = ss_cols / df_cols
    ems = ss_error / df_error

    denom = bms + (k - 1) * ems + k * (jms - ems) / n
    if denom == 0:
        return None
    return (bms - ems) / denom


def cohen_kappa_binary(pairs):
    """
    Cohen's kappa for paired binary annotations.
    Input: list of (a, b) pairs where a,b ∈ {0,1}.
    """
    n = len(pairs)
    if n == 0:
        return None
    a00 = sum(1 for a, b in pairs if a == 0 and b == 0)
    a01 = sum(1 for a, b in pairs if a == 0 and b == 1)
    a10 = sum(1 for a, b in pairs if a == 1 and b == 0)
    a11 = sum(1 for a, b in pairs if a == 1 and b == 1)
    po = (a00 + a11) / n
    p_a1 = (a10 + a11) / n
    p_b1 = (a01 + a11) / n
    pe = p_a1 * p_b1 + (1 - p_a1) * (1 - p_b1)
    if pe == 1.0:
        return 1.0
    return (po - pe) / (1 - pe)


def compute_inter_rater_reliability(queries):
    """
    Compute ICC(2,1) for point-landmark coordinates and Cohen's κ for
    area-landmark cell sets between the two OMFR specialists.

    The technical report specifies a two-way mixed-effects, absolute-agreement,
    single-measurement model. Numerically this is identical to Shrout & Fleiss
    ICC(2,1), so that is what we implement (and verify against the published
    Shrout & Fleiss 1979 worked example).

    Point landmarks
    ---------------
    Each query yields a numeric (row, col) pair from each rater. We compute
    ICC(2,1) separately for the row component and the column component (the
    standard 1-D ICC). In addition, we compute the mean inter-rater Euclidean
    distance across all queries — a complementary, intuitive number that
    reviewers familiar with landmark detection literature will recognise.

    Area landmarks
    --------------
    For each query, iterate over the FULL modality grid (e.g. 128 cells for
    panoramic). For each cell, record (in_rater1, in_rater2) as a binary pair.
    Cohen's κ is computed over the concatenated pairs across all area queries
    in a modality. Using the union-of-selected cells alone collapses κ to ~0
    because both rater base rates become ~1 — that's a mathematical artefact
    of conditioning on at-least-one-positive.

    Returns
    -------
    A dict with per-modality and overall results, plus a `notes` field
    describing how many queries were excluded for missing data.
    """
    point_rows_pairs = defaultdict(list)    # modality -> list[(r1,r2)]
    point_cols_pairs = defaultdict(list)    # modality -> list[(c1,c2)]
    point_eds = defaultdict(list)           # modality -> list[float] inter-rater EDs
    area_pairs = defaultdict(list)          # modality -> list[(in1, in2)]

    missing_point = 0
    missing_area = 0
    point_total = 0
    area_total = 0

    for q in queries:
        lm_type = q["landmark_type"]
        modality = q["sheet"]
        g1_raw = q.get("omfr_1")
        g2_raw = q.get("omfr_2")

        if not g1_raw or not g2_raw:
            if lm_type == "point":
                missing_point += 1
            else:
                missing_area += 1
            continue

        g1 = parse_ground_truth(g1_raw, lm_type)
        g2 = parse_ground_truth(g2_raw, lm_type)
        if not g1 or not g2:
            if lm_type == "point":
                missing_point += 1
            else:
                missing_area += 1
            continue

        if lm_type == "point":
            point_total += 1
            c1 = coord_to_numeric(g1[0])
            c2 = coord_to_numeric(g2[0])
            if c1 is None or c2 is None:
                missing_point += 1
                continue
            point_rows_pairs[modality].append((c1[0], c2[0]))
            point_cols_pairs[modality].append((c1[1], c2[1]))
            ed = math.sqrt((c1[0] - c2[0]) ** 2 + (c1[1] - c2[1]) ** 2)
            point_eds[modality].append(ed)
        else:  # area
            area_total += 1
            s1 = parse_cell_set(g1)
            s2 = parse_cell_set(g2)
            if not s1 and not s2:
                missing_area += 1
                continue
            # Cohen's κ on area landmarks requires a fixed cell universe —
            # the full modality grid. Using only the union of the two sets
            # excludes all (0,0) pairs and makes κ collapse to ~0 because
            # both rater base rates become ~1.
            grid = config.GRID_SPECS.get(modality)
            if not grid:
                continue
            rows = grid["rows"]
            cols = grid["cols"]
            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    cell = (r, c)
                    area_pairs[modality].append((
                        1 if cell in s1 else 0,
                        1 if cell in s2 else 0,
                    ))

    result = {
        "point_landmarks": {},
        "area_landmarks": {},
        "overall": {},
        "notes": {
            "point_total": point_total,
            "point_missing": missing_point,
            "area_total": area_total,
            "area_missing": missing_area,
        },
    }

    # ICC and Euclidean distance summary per modality for point landmarks
    all_row_pairs = []
    all_col_pairs = []
    all_eds = []
    for modality in sorted(set(point_rows_pairs.keys()) | set(point_cols_pairs.keys())):
        rps = point_rows_pairs.get(modality, [])
        cps = point_cols_pairs.get(modality, [])
        eds = point_eds.get(modality, [])
        if not rps:
            continue
        icc_row = icc_2_1([list(p) for p in rps])
        icc_col = icc_2_1([list(p) for p in cps])
        mean_ed = statistics.fmean(eds) if eds else None
        sd_ed = statistics.stdev(eds) if len(eds) > 1 else (0.0 if eds else None)
        result["point_landmarks"][modality] = {
            "n_queries": len(rps),
            "icc_row": round(icc_row, 4) if icc_row is not None else None,
            "icc_col": round(icc_col, 4) if icc_col is not None else None,
            "inter_rater_mean_ed": round(mean_ed, 4) if mean_ed is not None else None,
            "inter_rater_sd_ed": round(sd_ed, 4) if sd_ed is not None else None,
        }
        all_row_pairs.extend(rps)
        all_col_pairs.extend(cps)
        all_eds.extend(eds)

    if all_row_pairs:
        icc_row_all = icc_2_1([list(p) for p in all_row_pairs])
        icc_col_all = icc_2_1([list(p) for p in all_col_pairs])
        result["overall"]["point_icc_row"] = round(icc_row_all, 4) if icc_row_all is not None else None
        result["overall"]["point_icc_col"] = round(icc_col_all, 4) if icc_col_all is not None else None
        if all_eds:
            result["overall"]["point_inter_rater_mean_ed"] = round(statistics.fmean(all_eds), 4)
            result["overall"]["point_inter_rater_sd_ed"] = round(
                statistics.stdev(all_eds) if len(all_eds) > 1 else 0.0, 4
            )

    # Cohen's κ per modality for area landmarks
    all_area_pairs = []
    for modality in sorted(area_pairs.keys()):
        pairs = area_pairs[modality]
        if not pairs:
            continue
        kappa = cohen_kappa_binary(pairs)
        result["area_landmarks"][modality] = {
            "n_cells": len(pairs),
            "cohen_kappa": round(kappa, 4) if kappa is not None else None,
        }
        all_area_pairs.extend(pairs)

    if all_area_pairs:
        result["overall"]["area_cohen_kappa"] = round(
            cohen_kappa_binary(all_area_pairs), 4
        )

    return result


# ============================================================
# Student Paired Tests (Issue H5)
# ============================================================

def compute_student_per_query_metrics(queries, ground_truth):
    """
    For each query, compute the mean Euclidean distance of the 8 (or fewer)
    non-empty student responses from the ground truth.

    Returns a dict: {query_id: {"mean_ed": float, "n": int, "std_ed": float}}.
    Area-landmark queries and queries without GT or any student data are
    skipped.
    """
    student_metrics = {}
    for q in queries:
        if q.get("landmark_type") != "point":
            continue
        gt_entry = ground_truth.get(q["query_id"])
        if not gt_entry:
            continue
        gt_coord = gt_entry["coords"][0] if gt_entry["coords"] else None
        if gt_coord is None:
            continue
        eds = []
        for s in q.get("students", []) or []:
            if not s:
                continue
            parsed = parse_ground_truth(s, "point")
            if not parsed:
                continue
            ed = euclidean_distance(parsed[0], gt_coord)
            if ed is not None:
                eds.append(ed)
        if not eds:
            continue
        student_metrics[q["query_id"]] = {
            "n": len(eds),
            "mean_ed": statistics.fmean(eds),
            "std_ed": statistics.stdev(eds) if len(eds) > 1 else 0.0,
        }
    return student_metrics


def _matched_pairs_rank_biserial(diffs):
    """
    Rank-biserial correlation effect size for a Wilcoxon signed-rank test.
    r = (W+ - W-) / (W+ + W-)
    where W+ and W- are the sums of positive and negative ranks computed on
    the absolute differences, ignoring zero diffs. Returns None if all
    diffs are zero (no effect size computable).
    """
    non_zero = [d for d in diffs if d != 0]
    if not non_zero:
        return None
    abs_d = sorted(range(len(non_zero)), key=lambda i: abs(non_zero[i]))
    # Assign ranks with ties averaged
    ranks = [0.0] * len(non_zero)
    i = 0
    while i < len(non_zero):
        j = i
        while j + 1 < len(non_zero) and abs(non_zero[abs_d[j + 1]]) == abs(non_zero[abs_d[i]]):
            j += 1
        avg_rank = (i + j) / 2 + 1  # ranks are 1-based
        for k in range(i, j + 1):
            ranks[abs_d[k]] = avg_rank
        i = j + 1
    w_pos = sum(r for d, r in zip(non_zero, ranks) if d > 0)
    w_neg = sum(r for d, r in zip(non_zero, ranks) if d < 0)
    if w_pos + w_neg == 0:
        return None
    return (w_pos - w_neg) / (w_pos + w_neg)


def run_student_paired_tests(all_metrics, student_metrics, scipy_stats):
    """
    Paired Wilcoxon signed-rank test comparing each (model, strategy)
    per-query ED against the per-query student mean ED, with Bonferroni
    correction across the family of (model, strategy) batches.

    Family-wise error rate rationale
    --------------------------------
    Each batch tests the same null hypothesis ("this model's per-query error
    distribution is equal to the student mean error distribution") on the
    same student reference group. Running k=|batches| uncorrected tests
    inflates FWER; e.g. for 4 batches at α=0.05, FWER ≤ 1 - (1 - 0.05)^4 ≈
    0.185. Bonferroni correction (multiply each p by k, cap at 1.0) keeps
    the FWER at α overall. This is consistent with the technical report's
    statistical philosophy of Bonferroni-corrected pairwise post-hoc tests.

    Returns one result dict per (model, strategy) batch that has at least
    5 paired observations.
    """
    raw_results = []
    for batch_key, metrics_list in all_metrics.items():
        pairs = []
        for m in metrics_list:
            if m.get("landmark_type") != "point":
                continue
            ed = m.get("euclidean_distance")
            if ed is None:
                continue
            sm = student_metrics.get(m["query_id"])
            if not sm:
                continue
            pairs.append((ed, sm["mean_ed"]))
        if len(pairs) < 5:
            continue
        model_eds = [p[0] for p in pairs]
        student_eds = [p[1] for p in pairs]
        diffs = [a - b for a, b in zip(model_eds, student_eds)]
        try:
            wstat, wp = scipy_stats.wilcoxon(model_eds, student_eds, alternative="two-sided")
        except Exception:
            wstat, wp = None, None
        mean_diff = statistics.fmean(diffs)
        n_pos = sum(1 for d in diffs if d > 0)
        n_neg = sum(1 for d in diffs if d < 0)
        n_ties = sum(1 for d in diffs if d == 0)
        rb = _matched_pairs_rank_biserial(diffs)
        raw_results.append({
            "batch": batch_key,
            "n_paired": len(pairs),
            "model_mean_ed": round(statistics.fmean(model_eds), 4),
            "student_mean_ed": round(statistics.fmean(student_eds), 4),
            "mean_diff_model_minus_student": round(mean_diff, 4),
            "wilcoxon_stat": None if wstat is None else round(float(wstat), 4),
            "wilcoxon_p_raw": None if wp is None else round(float(wp), 6),
            "rank_biserial_r": None if rb is None else round(rb, 4),
            "pair_sign": {"pos": n_pos, "neg": n_neg, "ties": n_ties},
        })

    # Bonferroni correction across the family of batches
    k = len(raw_results)
    for res in raw_results:
        p_raw = res["wilcoxon_p_raw"]
        if p_raw is None or k == 0:
            res["wilcoxon_p_bonferroni"] = None
            res["significant_bonferroni"] = False
        else:
            adj = min(p_raw * k, 1.0)
            res["wilcoxon_p_bonferroni"] = round(adj, 6)
            res["significant_bonferroni"] = bool(adj < 0.05)
    return raw_results


# ============================================================
# Statistics Command
# ============================================================

def cmd_stats(args):
    """Run statistical tests on computed metrics."""
    metrics_path = config.ANALYSIS_DIR / "all_metrics.json"
    if not metrics_path.exists():
        print("ERROR: Metrics not computed yet. Run 'analysis.py metrics' first.")
        sys.exit(1)

    with open(metrics_path) as f:
        all_metrics = json.load(f)

    try:
        from scipy import stats as scipy_stats
        HAS_SCIPY = True
    except ImportError:
        HAS_SCIPY = False
        print("WARNING: scipy not installed. Install with: pip3 install scipy")
        print("         Statistical tests will be skipped.\n")

    try:
        import scikit_posthocs as sp
        HAS_POSTHOCS = True
    except ImportError:
        HAS_POSTHOCS = False
        print("WARNING: scikit-posthocs not installed. Install with: pip3 install scikit-posthocs")
        print("         Dunn's post-hoc test will be unavailable; falling back to pairwise MWU.\n")

    stat_results = {}

    # Collect ED values per model for each reporting level
    # Level 1: Per landmark type
    # Level 2: Per modality
    # Level 3: Grand aggregate

    for level_name, group_key in [
        ("per_structure", "structure"),
        ("per_modality", "modality"),
        ("grand", None),
    ]:
        print(f"\n{'='*60}")
        print(f"Level: {level_name}")
        print(f"{'='*60}")

        # Collect groups
        groups = defaultdict(lambda: defaultdict(list))

        for batch_key, metrics_list in all_metrics.items():
            for m in metrics_list:
                if m["landmark_type"] == "point" and m.get("euclidean_distance") is not None:
                    if group_key:
                        gk = m[group_key]
                    else:
                        gk = "all"
                    groups[gk][batch_key].append(m["euclidean_distance"])

                elif m["landmark_type"] == "area" and m.get("jaccard") is not None:
                    if group_key:
                        gk = m[group_key]
                    else:
                        gk = "all"
                    groups[gk][f"{batch_key}_jaccard"] = groups[gk].get(f"{batch_key}_jaccard", [])
                    groups[gk][f"{batch_key}_jaccard"].append(m["jaccard"])

        level_results = {}

        for group_name, model_data in sorted(groups.items()):
            print(f"\n  Group: {group_name}")

            group_results = {}

            for batch_key, values in sorted(model_data.items()):
                if not values:
                    continue
                summary = _stats_summary(values)
                group_results[batch_key] = summary
                print(f"    {batch_key}: n={summary['n']}, "
                      f"mean={summary['mean']:.3f} ± {summary['std']:.3f}, "
                      f"median={summary['median']:.3f}")

                # Normality test (Shapiro-Wilk requires n >= 3 and non-zero variance)
                if HAS_SCIPY and summary['n'] >= 3:
                    if max(values) == min(values):
                        # All values identical — degenerate case (perfect scores).
                        # Treat as non-normal to force the non-parametric path;
                        # KW will handle ties cleanly.
                        group_results[batch_key]["shapiro_p"] = None
                        group_results[batch_key]["zero_variance"] = True
                        print(f"      Shapiro-Wilk: skipped (zero variance)")
                    else:
                        import warnings as _warnings
                        with _warnings.catch_warnings():
                            _warnings.simplefilter("ignore")
                            _, p = scipy_stats.shapiro(values)
                        p = float(p)
                        group_results[batch_key]["shapiro_p"] = round(p, 6)
                        normal = "yes" if p > 0.05 else "no"
                        print(f"      Shapiro-Wilk: p={p:.4f} (normal: {normal})")

            # Group comparison (if multiple models / strategies in this group)
            ed_groups = {k: v for k, v in model_data.items() if "_jaccard" not in k}
            if HAS_SCIPY and len(ed_groups) >= 2:
                group_values = list(ed_groups.values())
                group_labels = list(ed_groups.keys())

                if not all(len(g) >= 5 for g in group_values):
                    level_results[group_name] = group_results
                    continue

                # Decide parametric vs non-parametric based on per-group
                # Shapiro-Wilk results stored above. If ALL groups pass
                # normality (p > 0.05) AND all have n >= 3, use ANOVA + Tukey.
                # A shapiro_p of None signals a degenerate zero-variance group;
                # route those to the non-parametric path which handles ties.
                def _is_normal(lbl):
                    sp = group_results.get(lbl, {}).get("shapiro_p")
                    return sp is not None and sp > 0.05
                all_normal = all(_is_normal(lbl) for lbl in group_labels)

                if all_normal:
                    # Parametric path: one-way ANOVA
                    f_stat, anova_p = scipy_stats.f_oneway(*group_values)
                    f_stat_val = float(f_stat) if not math.isnan(float(f_stat)) else None
                    anova_p_val = float(anova_p) if not math.isnan(float(anova_p)) else None
                    is_sig = bool(anova_p_val is not None and anova_p_val < 0.05)
                    group_results["anova"] = {
                        "f_statistic": round(f_stat_val, 4) if f_stat_val is not None else None,
                        "p_value": round(anova_p_val, 6) if anova_p_val is not None else None,
                        "significant": is_sig,
                    }
                    label = "not computable (zero variance)" if anova_p_val is None \
                            else ("significant" if is_sig else "not significant")
                    print(f"\n    ANOVA: F={f_stat}, p={anova_p} ({label})")
                    if not is_sig:
                        # Skip post-hoc path when the omnibus isn't significant
                        # or is undefined.
                        anova_p = 1.0
                    else:
                        anova_p = anova_p_val

                    if anova_p < 0.05:
                        # Tukey HSD via scipy if available, else pairwise t + Bonferroni
                        try:
                            from scipy.stats import tukey_hsd
                            res = tukey_hsd(*group_values)
                            pairwise = {}
                            for i in range(len(group_labels)):
                                for j in range(i + 1, len(group_labels)):
                                    adj_p = float(res.pvalue[i, j])
                                    mean_i = statistics.fmean(group_values[i])
                                    mean_j = statistics.fmean(group_values[j])
                                    pooled_sd = math.sqrt(
                                        (statistics.variance(group_values[i]) +
                                         statistics.variance(group_values[j])) / 2
                                    ) if len(group_values[i]) > 1 and len(group_values[j]) > 1 else 0
                                    d = (mean_i - mean_j) / pooled_sd if pooled_sd else 0
                                    pair_key = f"{group_labels[i]} vs {group_labels[j]}"
                                    pairwise[pair_key] = {
                                        "test": "tukey_hsd",
                                        "adjusted_p": round(adj_p, 6),
                                        "cohen_d": round(d, 4),
                                        "significant": bool(adj_p < 0.05),
                                    }
                                    sig = "***" if adj_p < 0.001 else "**" if adj_p < 0.01 else "*" if adj_p < 0.05 else "ns"
                                    print(f"    {pair_key}: Tukey p_adj={adj_p:.4f} {sig} (d={d:.3f})")
                            group_results["pairwise"] = pairwise
                            group_results["posthoc_method"] = "tukey_hsd"
                        except ImportError:
                            print("    (scipy.stats.tukey_hsd unavailable — skipping post-hoc)")
                else:
                    # Non-parametric path: Kruskal-Wallis + Dunn's with Bonferroni
                    # Guard against degenerate inputs: if every group is
                    # constant AND every group has the SAME constant, KW is
                    # mathematically undefined and scipy returns NaN. Detect
                    # this up-front and store a clean "not_computable" record
                    # so the JSON remains RFC-compliant.
                    flat = [v for g in group_values for v in g]
                    if max(flat) == min(flat):
                        group_results["kruskal_wallis"] = {
                            "h_statistic": None,
                            "p_value": None,
                            "significant": False,
                            "note": "all observations identical — test undefined",
                        }
                        print(f"\n    Kruskal-Wallis: skipped "
                              f"(all observations identical, test undefined)")
                        level_results[group_name] = group_results
                        continue

                    import warnings as _warnings
                    with _warnings.catch_warnings():
                        _warnings.simplefilter("ignore", RuntimeWarning)
                        h_stat, kw_p = scipy_stats.kruskal(*group_values)
                    h_stat = float(h_stat)
                    kw_p = float(kw_p)
                    if math.isnan(h_stat) or math.isnan(kw_p):
                        # Should not normally occur after the guard above,
                        # but be defensive in case of pathological inputs.
                        group_results["kruskal_wallis"] = {
                            "h_statistic": None,
                            "p_value": None,
                            "significant": False,
                            "note": "scipy returned NaN — test undefined",
                        }
                        print(f"\n    Kruskal-Wallis: skipped (scipy returned NaN)")
                        level_results[group_name] = group_results
                        continue
                    group_results["kruskal_wallis"] = {
                        "h_statistic": round(h_stat, 4),
                        "p_value": round(kw_p, 6),
                        "significant": bool(kw_p < 0.05),
                    }
                    print(f"\n    Kruskal-Wallis: H={h_stat:.3f}, p={kw_p:.4f} "
                          f"({'significant' if kw_p < 0.05 else 'not significant'})")

                    if kw_p < 0.05 and HAS_POSTHOCS:
                        pairwise = {}
                        # scikit-posthocs returns a k×k DataFrame with integer
                        # labels 1..k (1-based). Use positional .iat for safety
                        # and also assert the shape.
                        dunn_df = sp.posthoc_dunn(group_values, p_adjust="bonferroni")
                        assert dunn_df.shape == (len(group_labels), len(group_labels)), \
                            f"Dunn DF shape mismatch: {dunn_df.shape}"
                        for i in range(len(group_labels)):
                            for j in range(i + 1, len(group_labels)):
                                adj_p = float(dunn_df.iat[i, j])
                                u_stat, _ = scipy_stats.mannwhitneyu(
                                    group_values[i], group_values[j],
                                    alternative="two-sided"
                                )
                                n1, n2 = len(group_values[i]), len(group_values[j])
                                r = 1 - (2 * u_stat) / (n1 * n2)
                                pair_key = f"{group_labels[i]} vs {group_labels[j]}"
                                pairwise[pair_key] = {
                                    "test": "dunn_bonferroni",
                                    "adjusted_p": round(adj_p, 6),
                                    "effect_size_r": round(r, 4),
                                    "significant": bool(adj_p < 0.05),
                                }
                                sig = "***" if adj_p < 0.001 else "**" if adj_p < 0.01 else "*" if adj_p < 0.05 else "ns"
                                print(f"    {pair_key}: Dunn p_adj={adj_p:.4f} {sig} (r={r:.3f})")
                        group_results["pairwise"] = pairwise
                        group_results["posthoc_method"] = "dunn_bonferroni"
                    elif kw_p < 0.05:
                        # Fallback if scikit-posthocs unavailable
                        pairwise = {}
                        n_comparisons = len(group_labels) * (len(group_labels) - 1) // 2
                        for i in range(len(group_labels)):
                            for j in range(i + 1, len(group_labels)):
                                u_stat, mwu_p = scipy_stats.mannwhitneyu(
                                    group_values[i], group_values[j],
                                    alternative="two-sided"
                                )
                                adj_p = min(mwu_p * n_comparisons, 1.0)
                                n1, n2 = len(group_values[i]), len(group_values[j])
                                r = 1 - (2 * u_stat) / (n1 * n2)
                                pair_key = f"{group_labels[i]} vs {group_labels[j]}"
                                pairwise[pair_key] = {
                                    "test": "mannwhitneyu_bonferroni_fallback",
                                    "u_statistic": round(u_stat, 4),
                                    "p_value": round(mwu_p, 6),
                                    "adjusted_p": round(adj_p, 6),
                                    "effect_size_r": round(r, 4),
                                    "significant": bool(adj_p < 0.05),
                                }
                                sig = "***" if adj_p < 0.001 else "**" if adj_p < 0.01 else "*" if adj_p < 0.05 else "ns"
                                print(f"    {pair_key}: MWU p_adj={adj_p:.4f} {sig} (r={r:.3f}) [fallback]")
                        group_results["pairwise"] = pairwise
                        group_results["posthoc_method"] = "mannwhitneyu_bonferroni_fallback"

            level_results[group_name] = group_results

        stat_results[level_name] = level_results

    # Save statistical results
    out_path = config.ANALYSIS_DIR / "statistical_results.json"
    with open(out_path, "w") as f:
        # allow_nan=False makes json.dump raise on NaN/Inf instead of
        # silently emitting the non-RFC tokens NaN/Infinity. Catching the
        # error is preferable to producing invalid JSON downstream.
        json.dump(stat_results, f, indent=2, allow_nan=False)
    print(f"\nStatistical results saved: {out_path}")

    # ----- Inter-rater reliability (H4) -----
    print(f"\n{'='*60}")
    print("Inter-Rater Reliability (OMFR_1 vs OMFR_2)")
    print(f"{'='*60}")
    queries_for_irr = load_query_index()
    irr = compute_inter_rater_reliability(queries_for_irr)
    irr_path = config.ANALYSIS_DIR / "inter_rater_reliability.json"
    with open(irr_path, "w") as f:
        json.dump(irr, f, indent=2)
    if irr["notes"]["point_total"] == 0 and irr["notes"]["area_total"] == 0:
        print("  (no OMFR_2 data found; skipping.)")
    else:
        print(f"  Point: {irr['notes']['point_total']} queries, "
              f"{irr['notes']['point_missing']} missing")
        for mod, d in irr["point_landmarks"].items():
            print(f"    {mod}: ICC_row={d['icc_row']}, ICC_col={d['icc_col']}, n={d['n_queries']}")
        if "point_icc_row" in irr["overall"]:
            print(f"    Overall: ICC_row={irr['overall']['point_icc_row']}, "
                  f"ICC_col={irr['overall']['point_icc_col']}")
        print(f"  Area: {irr['notes']['area_total']} queries, "
              f"{irr['notes']['area_missing']} missing")
        for mod, d in irr["area_landmarks"].items():
            print(f"    {mod}: Cohen's κ={d['cohen_kappa']} ({d['n_cells']} cells)")
        if "area_cohen_kappa" in irr["overall"]:
            print(f"    Overall: Cohen's κ={irr['overall']['area_cohen_kappa']}")
    print(f"\nInter-rater results saved: {irr_path}")

    # ----- Student paired tests (H5) -----
    print(f"\n{'='*60}")
    print("Student Paired Tests (Wilcoxon signed-rank vs student mean)")
    print(f"{'='*60}")
    ground_truth = load_ground_truth(queries_for_irr)
    student_metrics_map = compute_student_per_query_metrics(queries_for_irr, ground_truth)
    # Persist per-query student mean ED for Bland-Altman plotting (Issue 6)
    student_metrics_path = config.RESULTS_DIR / "student_metrics.json"
    serialised_sm = {
        qid: {"mean_ed": round(v["mean_ed"], 6),
              "std_ed": round(v["std_ed"], 6),
              "n": v["n"]}
        for qid, v in student_metrics_map.items()
    }
    with open(student_metrics_path, "w") as f:
        json.dump(serialised_sm, f, indent=2)
    print(f"  Student per-query metrics: {len(student_metrics_map)} queries with data")
    print(f"  Saved: {student_metrics_path}")

    if HAS_SCIPY and student_metrics_map:
        paired_results = run_student_paired_tests(
            all_metrics, student_metrics_map, scipy_stats
        )
        paired_path = config.ANALYSIS_DIR / "student_paired_tests.json"
        with open(paired_path, "w") as f:
            json.dump(paired_results, f, indent=2)
        for res in paired_results:
            p_raw = res["wilcoxon_p_raw"]
            p_adj = res["wilcoxon_p_bonferroni"]
            sig = "*" if res["significant_bonferroni"] else "ns"
            r = res.get("rank_biserial_r")
            print(f"  {res['batch']}: n={res['n_paired']}, "
                  f"model={res['model_mean_ed']}, student={res['student_mean_ed']}, "
                  f"diff={res['mean_diff_model_minus_student']:+}, "
                  f"Wilcoxon p_raw={p_raw}, p_adj={p_adj} ({sig}), "
                  f"r={r}")
        if not paired_results:
            print("  (no batches with enough paired data to run Wilcoxon)")
        print(f"\nStudent paired-test results saved: {paired_path}")
    else:
        print("  (skipped: student data empty or scipy unavailable)")


# ============================================================
# Visualization Command
# ============================================================

def cmd_visualize(args):
    """Generate visualization plots."""
    metrics_path = config.ANALYSIS_DIR / "all_metrics.json"
    if not metrics_path.exists():
        print("ERROR: Metrics not computed yet. Run 'analysis.py metrics' first.")
        sys.exit(1)

    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
    except ImportError:
        print("ERROR: matplotlib not installed. Install with: pip3 install matplotlib")
        sys.exit(1)

    with open(metrics_path) as f:
        all_metrics = json.load(f)

    plots_dir = config.ANALYSIS_DIR / "plots"
    plots_dir.mkdir(parents=True, exist_ok=True)

    # ---- Plot 1: Box plot of ED per model ----
    ed_data = {}
    for batch_key, metrics_list in all_metrics.items():
        eds = [m["euclidean_distance"] for m in metrics_list
               if m["landmark_type"] == "point" and m.get("euclidean_distance") is not None]
        if eds:
            ed_data[batch_key] = eds

    if ed_data:
        fig, ax = plt.subplots(figsize=(10, 6))
        labels = list(ed_data.keys())
        data = [ed_data[k] for k in labels]
        bp = ax.boxplot(data, tick_labels=[l.replace("_", "\n") for l in labels],
                        patch_artist=True, showmeans=True)
        colors = plt.cm.Set2.colors
        for i, patch in enumerate(bp["boxes"]):
            patch.set_facecolor(colors[i % len(colors)])
        ax.set_ylabel("Euclidean Distance (grid cells)")
        ax.set_title("Point Landmark Accuracy — Euclidean Distance Distribution")
        ax.axhline(y=1, color="red", linestyle="--", alpha=0.5, label="1-cell threshold")
        ax.axhline(y=2, color="orange", linestyle="--", alpha=0.5, label="2-cell threshold")
        ax.legend()
        plt.tight_layout()
        plt.savefig(plots_dir / "boxplot_ed_by_model.png", dpi=150)
        plt.close()
        print(f"  Saved: boxplot_ed_by_model.png")

    # ---- Plot 2: SDR bar chart ----
    if ed_data:
        fig, ax = plt.subplots(figsize=(10, 6))
        thresholds = [0, 1, math.sqrt(2), 2]
        threshold_labels = ["SDR@0\n(exact)", "SDR@1\n(±1 cell)", "SDR@√2\n(±1 diag)", "SDR@2\n(±2 cells)"]
        x = range(len(thresholds))
        width = 0.8 / len(ed_data)

        for i, (batch_key, eds) in enumerate(ed_data.items()):
            sdrs = [success_detection_rate(eds, t) for t in thresholds]
            offset = (i - len(ed_data)/2 + 0.5) * width
            bars = ax.bar([xi + offset for xi in x], sdrs, width,
                          label=batch_key.replace("_", " "), color=colors[i % len(colors)])
            for bar, sdr in zip(bars, sdrs):
                if sdr > 0:
                    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                            f"{sdr:.0f}%", ha="center", va="bottom", fontsize=7)

        ax.set_xticks(list(x))
        ax.set_xticklabels(threshold_labels)
        ax.set_ylabel("Success Detection Rate (%)")
        ax.set_title("SDR at Different Thresholds")
        ax.set_ylim(0, 110)
        ax.legend(loc="upper left", fontsize=8)
        plt.tight_layout()
        plt.savefig(plots_dir / "sdr_bar_chart.png", dpi=150)
        plt.close()
        print(f"  Saved: sdr_bar_chart.png")

    # ---- Plot 3: ED heatmap per landmark per model ----
    structures_point = []
    for mod in ["PANORAMIC", "PERIAPICAL", "CEPHALOMETRIC"]:
        for lm in config.LANDMARKS[mod]:
            if lm["type"] == "point":
                structures_point.append(lm["structure"])

    if ed_data:
        model_keys = list(ed_data.keys())
        heatmap_data = []

        for struct in structures_point:
            row = []
            for batch_key, metrics_list in all_metrics.items():
                eds = [m["euclidean_distance"] for m in metrics_list
                       if m["structure"] == struct and m.get("euclidean_distance") is not None]
                mean_ed = sum(eds) / len(eds) if eds else float("nan")
                row.append(mean_ed)
            heatmap_data.append(row)

        if heatmap_data and model_keys:
            fig, ax = plt.subplots(figsize=(max(8, len(model_keys)*2.5), max(6, len(structures_point)*0.6)))
            im = ax.imshow(heatmap_data, cmap="RdYlGn_r", aspect="auto")
            ax.set_xticks(range(len(model_keys)))
            ax.set_xticklabels([k.replace("_", "\n") for k in model_keys], fontsize=8)
            ax.set_yticks(range(len(structures_point)))
            ax.set_yticklabels(structures_point, fontsize=9)
            ax.set_title("Mean ED per Landmark × Model")
            plt.colorbar(im, ax=ax, label="Mean ED (lower = better)")

            # Add text annotations
            for i in range(len(structures_point)):
                for j in range(len(model_keys)):
                    val = heatmap_data[i][j]
                    if not math.isnan(val):
                        ax.text(j, i, f"{val:.2f}", ha="center", va="center",
                                fontsize=8, color="black" if val < 3 else "white")

            plt.tight_layout()
            plt.savefig(plots_dir / "heatmap_ed_by_landmark.png", dpi=150)
            plt.close()
            print(f"  Saved: heatmap_ed_by_landmark.png")

    # ---- Plot 4: Jaccard/Dice scatter for area landmarks ----
    area_data = {}
    for batch_key, metrics_list in all_metrics.items():
        jd = [(m["jaccard"], m["dice"]) for m in metrics_list
              if m["landmark_type"] == "area" and m.get("jaccard") is not None]
        if jd:
            area_data[batch_key] = jd

    if area_data:
        fig, ax = plt.subplots(figsize=(8, 8))
        for i, (batch_key, jd_pairs) in enumerate(area_data.items()):
            jaccards = [p[0] for p in jd_pairs]
            dices = [p[1] for p in jd_pairs]
            ax.scatter(jaccards, dices, alpha=0.5, label=batch_key.replace("_", " "),
                       color=colors[i % len(colors)], s=30)
        ax.plot([0, 1], [0, 1], "k--", alpha=0.3, label="J = D (never happens)")
        ax.set_xlabel("Jaccard Index")
        ax.set_ylabel("Dice Coefficient")
        ax.set_title("Area Landmark Overlap — Jaccard vs Dice")
        ax.set_xlim(-0.05, 1.05)
        ax.set_ylim(-0.05, 1.05)
        ax.legend(fontsize=8)
        ax.set_aspect("equal")
        plt.tight_layout()
        plt.savefig(plots_dir / "scatter_jaccard_dice.png", dpi=150)
        plt.close()
        print(f"  Saved: scatter_jaccard_dice.png")

    # ---- Plot 5: Per-modality comparison ----
    if ed_data:
        modalities = ["PANORAMIC", "PERIAPICAL", "CEPHALOMETRIC"]
        fig, axes = plt.subplots(1, 3, figsize=(15, 5), sharey=True)

        for idx, mod in enumerate(modalities):
            ax = axes[idx]
            mod_data = {}
            for batch_key, metrics_list in all_metrics.items():
                eds = [m["euclidean_distance"] for m in metrics_list
                       if m["modality"] == mod and m["landmark_type"] == "point"
                       and m.get("euclidean_distance") is not None]
                if eds:
                    mod_data[batch_key] = eds

            if mod_data:
                labels = list(mod_data.keys())
                data = [mod_data[k] for k in labels]
                bp = ax.boxplot(data, tick_labels=[l.replace("_", "\n") for l in labels],
                                patch_artist=True, showmeans=True)
                for i, patch in enumerate(bp["boxes"]):
                    patch.set_facecolor(colors[i % len(colors)])

            ax.set_title(mod.capitalize())
            ax.axhline(y=1, color="red", linestyle="--", alpha=0.5)
            if idx == 0:
                ax.set_ylabel("Euclidean Distance (grid cells)")

        plt.suptitle("ED Distribution by Modality", fontsize=14)
        plt.tight_layout()
        plt.savefig(plots_dir / "boxplot_ed_by_modality.png", dpi=150)
        plt.close()
        print(f"  Saved: boxplot_ed_by_modality.png")

    # ---- Plot 6: Radar chart — model performance across landmark types ----
    # Computes mean NED per point structure and mean Jaccard per area structure for each model
    all_structures = []
    for mod in ["PANORAMIC", "PERIAPICAL", "CEPHALOMETRIC"]:
        for lm in config.LANDMARKS[mod]:
            all_structures.append(lm["structure"])

    model_keys = list(all_metrics.keys())
    if model_keys and all_structures:
        radar_data = {}  # {model_key: [score_per_structure]}
        for batch_key, metrics_list in all_metrics.items():
            scores = []
            for struct in all_structures:
                struct_metrics = [m for m in metrics_list if m["structure"] == struct]
                if not struct_metrics:
                    scores.append(0)
                    continue
                lm_type = struct_metrics[0]["landmark_type"]
                if lm_type == "point":
                    neds = [m["normalized_ed"] for m in struct_metrics
                            if m.get("normalized_ed") is not None]
                    # Convert NED to accuracy score (1 - NED), clamped to [0, 1]
                    score = max(0, 1 - (sum(neds) / len(neds))) if neds else 0
                else:
                    jaccards = [m["jaccard"] for m in struct_metrics
                                if m.get("jaccard") is not None]
                    score = sum(jaccards) / len(jaccards) if jaccards else 0
                scores.append(score)
            radar_data[batch_key] = scores

        if radar_data:
            n_structs = len(all_structures)
            angles = [i / n_structs * 2 * math.pi for i in range(n_structs)]
            angles += angles[:1]  # close the polygon

            fig, ax = plt.subplots(figsize=(9, 9), subplot_kw=dict(polar=True))
            for i, (batch_key, scores) in enumerate(radar_data.items()):
                values = scores + scores[:1]
                ax.plot(angles, values, linewidth=2,
                        label=batch_key.replace("_", " "),
                        color=colors[i % len(colors)])
                ax.fill(angles, values, alpha=0.1, color=colors[i % len(colors)])

            # Shorten structure names for readability
            short_names = [s.replace("_", "\n") for s in all_structures]
            ax.set_xticks(angles[:-1])
            ax.set_xticklabels(short_names, fontsize=7)
            ax.set_ylim(0, 1)
            ax.set_title("Model Performance Profile Across Landmarks\n"
                         "(Point: 1−NED, Area: Jaccard)", fontsize=12, pad=20)
            ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.1), fontsize=8)
            plt.tight_layout()
            plt.savefig(plots_dir / "radar_landmark_profile.png", dpi=150,
                        bbox_inches="tight")
            plt.close()
            print(f"  Saved: radar_landmark_profile.png")

    # ---- Plot 7: Bland-Altman plots (Issue 6) ----
    # Agreement analysis between model predictions and student mean responses.
    # Requires a precomputed student metrics file at results/student_metrics.json
    # with the shape: { "<query_id>": { "mean_ed": <float> }, ... }
    #
    # Methodological note: the classical Bland-Altman limits of agreement
    # (mean ± 1.96·SD) assume the differences are approximately normally
    # distributed. For each modality panel we run a Shapiro-Wilk test on
    # the differences and annotate the subplot with the p-value so the
    # reader can judge whether the parametric LoA is appropriate. When the
    # differences are non-normal, the LoA should be interpreted as a rough
    # guide rather than an exact 95% limit.
    student_metrics_path = config.RESULTS_DIR / "student_metrics.json"
    try:
        from scipy import stats as _scipy_stats
        _HAS_SCIPY = True
    except ImportError:
        _HAS_SCIPY = False

    if not student_metrics_path.exists():
        print("  Skipping Bland-Altman: results/student_metrics.json not found "
              "(student data ingestion pending).")
    else:
        with open(student_metrics_path) as f:
            student_metrics = json.load(f)

        ba_diagnostics = {}   # batch_key -> per-modality diagnostics
        for batch_key, metrics_list in all_metrics.items():
            # Pair per-query model ED with student mean ED (point landmarks only)
            paired = []  # list of (model_ed, student_mean_ed, modality)
            for m in metrics_list:
                if m["landmark_type"] != "point":
                    continue
                if m.get("euclidean_distance") is None:
                    continue
                s = student_metrics.get(m["query_id"])
                if not s or s.get("mean_ed") is None:
                    continue
                paired.append((m["euclidean_distance"], s["mean_ed"], m["modality"]))

            if len(paired) < 5:
                continue

            modalities = ["PANORAMIC", "PERIAPICAL", "CEPHALOMETRIC"]
            fig, axes = plt.subplots(1, 3, figsize=(15, 5), sharey=True)
            ba_diagnostics[batch_key] = {}
            for idx, mod in enumerate(modalities):
                ax = axes[idx]
                mod_pairs = [p for p in paired if p[2] == mod]
                if len(mod_pairs) < 2:
                    ax.set_title(f"{mod.capitalize()} (n/a)")
                    ax.axis("off")
                    continue

                model_eds = [p[0] for p in mod_pairs]
                student_eds = [p[1] for p in mod_pairs]
                means = [(a + b) / 2 for a, b in zip(model_eds, student_eds)]
                diffs = [a - b for a, b in zip(model_eds, student_eds)]

                mean_diff = sum(diffs) / len(diffs)
                var_diff = sum((d - mean_diff) ** 2 for d in diffs) / (len(diffs) - 1) if len(diffs) > 1 else 0.0
                sd_diff = math.sqrt(var_diff)
                loa_upper = mean_diff + 1.96 * sd_diff
                loa_lower = mean_diff - 1.96 * sd_diff

                # Normality diagnostic for the differences
                shapiro_p = None
                if _HAS_SCIPY and len(diffs) >= 3 and max(diffs) != min(diffs):
                    import warnings as _w
                    with _w.catch_warnings():
                        _w.simplefilter("ignore")
                        _, shapiro_p = _scipy_stats.shapiro(diffs)
                    shapiro_p = float(shapiro_p)
                ba_diagnostics[batch_key][mod] = {
                    "n": len(mod_pairs),
                    "bias": round(mean_diff, 4),
                    "sd_diff": round(sd_diff, 4),
                    "loa_upper": round(loa_upper, 4),
                    "loa_lower": round(loa_lower, 4),
                    "shapiro_p_diffs": round(shapiro_p, 6) if shapiro_p is not None else None,
                    "diffs_approximately_normal": bool(
                        shapiro_p is not None and shapiro_p > 0.05
                    ),
                }

                ax.scatter(means, diffs, alpha=0.5, s=25, color=colors[0])
                ax.axhline(mean_diff, color="black", linestyle="-", linewidth=1.2,
                           label=f"Bias = {mean_diff:.2f}")
                ax.axhline(loa_upper, color="red", linestyle="--", linewidth=1,
                           label=f"+1.96 SD = {loa_upper:.2f}")
                ax.axhline(loa_lower, color="red", linestyle="--", linewidth=1,
                           label=f"-1.96 SD = {loa_lower:.2f}")
                ax.axhline(0, color="gray", linestyle=":", alpha=0.5)
                ax.set_xlabel("Mean of Model & Student ED")
                if idx == 0:
                    ax.set_ylabel("Model ED − Student Mean ED")
                title = f"{mod.capitalize()} (n={len(mod_pairs)})"
                if shapiro_p is not None:
                    note = "normal" if shapiro_p > 0.05 else "non-normal"
                    title += f"\nShapiro p={shapiro_p:.3f} ({note})"
                ax.set_title(title, fontsize=10)
                ax.legend(fontsize=7, loc="upper right")

            plt.suptitle(f"Bland-Altman: {batch_key.replace('_', ' ')} vs Students",
                         fontsize=13)
            plt.tight_layout()
            fname = f"bland_altman_{batch_key}.png"
            plt.savefig(plots_dir / fname, dpi=150)
            plt.close()
            print(f"  Saved: {fname}")

        # Persist the numeric diagnostics alongside the plots
        if ba_diagnostics:
            ba_path = config.ANALYSIS_DIR / "bland_altman_diagnostics.json"
            with open(ba_path, "w") as f:
                json.dump(ba_diagnostics, f, indent=2, allow_nan=False)
            print(f"  Bland-Altman diagnostics saved: {ba_path}")

    print(f"\nAll plots saved to: {plots_dir}")


# ============================================================
# Report Command
# ============================================================

def cmd_report(args):
    """Generate summary CSV/Excel tables."""
    metrics_path = config.ANALYSIS_DIR / "all_metrics.json"
    if not metrics_path.exists():
        print("ERROR: Run 'analysis.py metrics' first.")
        sys.exit(1)

    with open(metrics_path) as f:
        all_metrics = json.load(f)

    # Build summary table: Model × Strategy × Modality → metrics
    rows = []

    for batch_key, metrics_list in all_metrics.items():
        parts = batch_key.split("_", 1)
        model_key = parts[0] if len(parts) > 0 else batch_key

        # Try to determine model and strategy from batch_key
        for mk in config.MODELS:
            if batch_key.startswith(mk):
                model_key = mk
                strategy = batch_key[len(mk)+1:]
                break
        else:
            model_key = batch_key
            strategy = "unknown"

        for modality in ["PANORAMIC", "PERIAPICAL", "CEPHALOMETRIC", "ALL"]:
            if modality == "ALL":
                point_m = [m for m in metrics_list if m["landmark_type"] == "point"]
                area_m = [m for m in metrics_list if m["landmark_type"] == "area"]
            else:
                point_m = [m for m in metrics_list if m["landmark_type"] == "point" and m["modality"] == modality]
                area_m = [m for m in metrics_list if m["landmark_type"] == "area" and m["modality"] == modality]

            # Point metrics
            eds = [m["euclidean_distance"] for m in point_m if m.get("euclidean_distance") is not None]
            neds = [m["normalized_ed"] for m in point_m if m.get("normalized_ed") is not None]

            row = {
                "model": config.MODELS.get(model_key, {}).get("display_name", model_key),
                "strategy": strategy,
                "modality": modality,
                "n_point": len(eds),
                "mean_ed": round(statistics.fmean(eds), 3) if eds else None,
                "std_ed": round(statistics.stdev(eds), 3) if len(eds) > 1 else None,
                "median_ed": round(statistics.median(eds), 3) if eds else None,
                "mean_ned": round(statistics.fmean(neds), 4) if neds else None,
                "sdr_0": round(success_detection_rate(eds, 0), 1) if eds else None,
                "sdr_1": round(success_detection_rate(eds, 1), 1) if eds else None,
                "sdr_sqrt2": round(success_detection_rate(eds, math.sqrt(2)), 1) if eds else None,
                "sdr_2": round(success_detection_rate(eds, 2), 1) if eds else None,
            }

            # Area metrics — None entries are excluded (undefined GT)
            jaccards = [m["jaccard"] for m in area_m if m.get("jaccard") is not None]
            dices = [m["dice"] for m in area_m if m.get("dice") is not None]

            row["n_area"] = len(jaccards)
            row["mean_jaccard"] = round(statistics.fmean(jaccards), 3) if jaccards else None
            row["mean_dice"] = round(statistics.fmean(dices), 3) if dices else None

            # Compliance rate: parseable AND no failure override
            total = len(point_m) + len(area_m)
            compliant = sum(
                1 for m in point_m + area_m
                if m.get("parseable", False) and m.get("failure_category") is None
            )
            row["compliance_rate"] = round(compliant/total*100, 1) if total > 0 else None

            rows.append(row)

    # Save as CSV
    import csv
    out_csv = config.ANALYSIS_DIR / "summary_table.csv"
    if rows:
        fieldnames = rows[0].keys()
        with open(out_csv, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
        print(f"Summary table saved: {out_csv}")

    # Also save detailed per-query results as CSV
    detail_rows = []
    for batch_key, metrics_list in all_metrics.items():
        for m in metrics_list:
            detail_rows.append({
                "batch": batch_key,
                "query_id": m["query_id"],
                "image_id": m["image_id"],
                "modality": m["modality"],
                "structure": m["structure"],
                "type": m["landmark_type"],
                "ground_truth": ",".join(m.get("ground_truth", [])),
                "prediction": ",".join(m.get("prediction", [])) if m.get("prediction") else "",
                "parseable": m.get("parseable", False),
                "ed": m.get("euclidean_distance"),
                "ned": m.get("normalized_ed"),
                "jaccard": m.get("jaccard"),
                "dice": m.get("dice"),
            })

    out_detail = config.ANALYSIS_DIR / "detailed_results.csv"
    if detail_rows:
        with open(out_detail, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=detail_rows[0].keys())
            writer.writeheader()
            writer.writerows(detail_rows)
        print(f"Detailed results saved: {out_detail}")


# ============================================================
# All Command
# ============================================================

def cmd_all(args):
    """Run the complete analysis pipeline."""
    print("Step 1/4: Computing metrics...")
    cmd_metrics(args)
    print("\nStep 2/4: Running statistical tests...")
    cmd_stats(args)
    print("\nStep 3/4: Generating visualizations...")
    cmd_visualize(args)
    print("\nStep 4/4: Generating summary tables...")
    cmd_report(args)
    print("\nAnalysis complete!")


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Dental MLLM Benchmark — Analysis Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Commands:
  metrics     Compute ED, NED, Jaccard, Dice, SDR for all responses
  stats       Run Shapiro-Wilk, Kruskal-Wallis, pairwise tests
  visualize   Generate box plots, heatmaps, bar charts
  report      Generate summary CSV tables
  all         Run everything in sequence
        """,
    )
    parser.add_argument("command", choices=["metrics", "stats", "visualize", "report", "all"])
    args = parser.parse_args()

    commands = {
        "metrics": cmd_metrics,
        "stats": cmd_stats,
        "visualize": cmd_visualize,
        "report": cmd_report,
        "all": cmd_all,
    }
    commands[args.command](args)


if __name__ == "__main__":
    main()
