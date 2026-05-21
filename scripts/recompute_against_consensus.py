"""
Re-evaluate the existing GPT-5.4 batch outputs against the consensus GT.

Reads:
    results_consensus/query_index.json      — extended schema (consensus_gt, omfr_1/2, student)
    results_consensus/reanalysis_anchor.json — SHA-256 anchor for raw JSONLs
    results_full/run{N}/responses/*.jsonl   — frozen raw OpenAI batch outputs (READ-ONLY)

Writes (atomic):
    results_consensus/run{N}/parsed_responses.json   — per-(custom_id, rep) record
    results_consensus/run{N}/compliance_stats.json   — per-rep compliance summary
    results_consensus/full_run_records.pkl           — per-(query, strategy) record with
                                                       predictions and metrics computed
                                                       against consensus_gt, omfr_1, omfr_2,
                                                       and student.

No API calls. No writes outside results_consensus/. Cryptographically refuses to
proceed if any anchored JSONL has drifted from the SHA recorded at prepare time.

Usage:
    .venv/bin/python scripts/recompute_against_consensus.py \
        [--anchor-root results_full] [--sandbox results_consensus]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import pickle
import sys
from pathlib import Path
from collections import defaultdict, Counter

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import config  # noqa: E402
import pipeline  # noqa: E402  (parse_grid_coordinate, atomic writes)


# ── Geometry ────────────────────────────────────────────────────────

def grid_dims(modality):
    g = config.GRID_SPECS[modality]
    return g["cols"], g["rows"]


def cell_to_xy(cell, modality):
    """Normalised cell string (e.g. 'C5') → (col_int, row_int) 1-indexed.
    Returns None if the cell is malformed or out of range for the modality."""
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


def parse_reference_cells(text, modality):
    """Reference (GT/OMFR/student) is a comma-separated list of cells.
    Returns list[(col,row)] of valid cells; tolerates extra whitespace."""
    if text is None:
        return []
    out = []
    for tok in str(text).replace("\n", ",").replace(";", ",").split(","):
        xy = cell_to_xy(tok.strip(), modality)
        if xy is not None:
            out.append(xy)
    return out


def parse_predicted_cells(raw_response, modality):
    """Use the project's canonical parser to extract grid coordinates from the
    model's raw output. Returns list[(col,row)] of valid cells or [] if none."""
    parsed = pipeline.parse_grid_coordinate(raw_response or "")
    if not parsed:
        return []
    out = []
    for tok in parsed:
        xy = cell_to_xy(tok, modality)
        if xy is not None:
            out.append(xy)
    return out


# ── Metrics ─────────────────────────────────────────────────────────

def euclid(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


def jaccard(set_a, set_b):
    if not set_a and not set_b:
        return 1.0
    if not set_a or not set_b:
        return 0.0
    return len(set_a & set_b) / len(set_a | set_b)


def dice(set_a, set_b):
    n = len(set_a) + len(set_b)
    if n == 0:
        return 1.0
    return (2 * len(set_a & set_b)) / n


def metric_pair(pred_cells, ref_cells, landmark_type):
    """Return (ed, jacc, dice) — each float or None — for a single rep."""
    if not pred_cells or not ref_cells:
        return None, None, None
    if landmark_type == "point":
        # Use the first valid predicted cell against the first valid ref cell.
        return euclid(pred_cells[0], ref_cells[0]), None, None
    # Area: set-overlap metrics
    a, b = set(pred_cells), set(ref_cells)
    return None, jaccard(a, b), dice(a, b)


def mean_or_none(values):
    vals = [v for v in values if v is not None]
    if not vals:
        return None
    return sum(vals) / len(vals)


# ── Anchor verification ─────────────────────────────────────────────

def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def verify_anchor(anchor_path, anchor_root):
    """Refuse to proceed if any frozen JSONL has drifted from its prepare-time SHA."""
    anchor = json.loads(anchor_path.read_text())
    declared_root = Path(anchor["anchor_root"]).resolve()
    if declared_root != anchor_root.resolve():
        raise RuntimeError(
            f"Anchor file declares root {declared_root}, but --anchor-root is {anchor_root}. "
            f"Refusing to proceed."
        )
    drift = []
    for rel, info in sorted(anchor["files"].items()):
        path = anchor_root / rel
        if not path.exists():
            drift.append(f"missing: {rel}")
            continue
        actual = sha256_file(path)
        if actual != info["sha256"]:
            drift.append(f"sha mismatch: {rel} (declared {info['sha256'][:12]}…, "
                         f"actual {actual[:12]}…)")
    upstream_qi = anchor.get("upstream_query_index")
    if upstream_qi:
        path = anchor_root / upstream_qi["path"]
        if path.exists():
            if sha256_file(path) != upstream_qi["sha256"]:
                drift.append(f"sha mismatch: {upstream_qi['path']} (upstream query index)")
        else:
            drift.append(f"missing: {upstream_qi['path']}")
    if drift:
        raise RuntimeError(
            f"Anchor verification FAILED — {len(drift)} drift(s) detected. "
            f"First few:\n  " + "\n  ".join(drift[:5])
        )
    print(f"Anchor verified: {len(anchor['files'])} JSONL files + upstream query_index "
          f"intact at {anchor_root}.")


# ── Main ────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sandbox", default="results_consensus",
                    help="v2 sandbox directory containing query_index.json + "
                         "reanalysis_anchor.json (default: results_consensus)")
    ap.add_argument("--anchor-root", default="results_full",
                    help="Frozen sandbox holding the raw JSONLs to re-evaluate "
                         "(default: results_full)")
    args = ap.parse_args()

    sandbox = (ROOT / args.sandbox).resolve()
    anchor_root = (ROOT / args.anchor_root).resolve()
    if sandbox == anchor_root:
        print(f"ERROR: sandbox and anchor-root must differ.", file=sys.stderr)
        sys.exit(1)

    anchor_path = sandbox / "reanalysis_anchor.json"
    qi_path = sandbox / "query_index.json"
    if not anchor_path.exists() or not qi_path.exists():
        print(f"ERROR: run pipeline.py prepare_v2 first; "
              f"missing {anchor_path} or {qi_path}.", file=sys.stderr)
        sys.exit(1)

    verify_anchor(anchor_path, anchor_root)

    # Build canonical (query_id → query) lookup
    qi = json.loads(qi_path.read_text())
    qmap = {q["query_id"]: q for q in qi}
    print(f"Loaded {len(qi)} queries from {qi_path}")

    # Pre-parse all reference cell sets (consensus, omfr_1, omfr_2, student) ONCE per query
    # and store on the query record for fast lookup downstream.
    for q in qi:
        mod = q["sheet"]
        q["_ref_cells"] = {
            "consensus_gt": parse_reference_cells(q.get("consensus_gt"), mod),
            "omfr_1":       parse_reference_cells(q.get("omfr_1"), mod),
            "omfr_2":       parse_reference_cells(q.get("omfr_2"), mod),
            "student":      parse_reference_cells(q.get("student"), mod),
        }

    # ── Walk reps ────────────────────────────────────────────────────
    REFS = ("consensus_gt", "omfr_1", "omfr_2", "student")
    # records keyed by (query_id, strategy)
    records = {}
    for q in qi:
        for strat in ("zero_shot", "guided"):
            key = (q["query_id"], strat)
            records[key] = {
                "query_id": q["query_id"],
                "strategy": strat,
                "modality": q["sheet"],
                "image_id": q["image_id"],
                "structure": q["structure"],
                "landmark_type": q["landmark_type"],
                "uses_fdi": q.get("uses_fdi", False),
                # all reference annotations (raw strings) — preserved for the appendix
                "consensus_gt":  q.get("consensus_gt"),
                "omfr_1":        q.get("omfr_1"),
                "omfr_1_second": q.get("omfr_1_second"),
                "omfr_2":        q.get("omfr_2"),
                "omfr_2_second": q.get("omfr_2_second"),
                "student":       q.get("student"),
                # rep-level fields
                "rep_raw": [None, None, None],
                "rep_pred_cells": [None, None, None],
                "rep_failure": [None, None, None],
                # metrics per reference computed below
                "metrics": {ref: {"rep_ed": [None]*3, "rep_jaccard": [None]*3,
                                  "rep_dice": [None]*3} for ref in REFS},
            }

    parsed_per_rep = {1: [], 2: [], 3: []}

    for rep in (1, 2, 3):
        rep_dir = anchor_root / f"run{rep}" / "responses"
        if not rep_dir.exists():
            print(f"ERROR: missing {rep_dir}", file=sys.stderr)
            sys.exit(1)
        n_lines = 0
        n_parse_fail = 0
        for jl in sorted(rep_dir.glob("*.jsonl")):
            for line in open(jl):
                line = line.rstrip("\n")
                if not line:
                    continue
                obj = json.loads(line)
                cid = obj["custom_id"]
                if cid.endswith("_zero_shot"):
                    qid, strat = cid[:-len("_zero_shot")], "zero_shot"
                elif cid.endswith("_guided"):
                    qid, strat = cid[:-len("_guided")], "guided"
                else:
                    raise ValueError(f"Unrecognised custom_id format: {cid}")

                content = (obj.get("response", {}).get("body", {})
                              .get("choices", [{}])[0].get("message", {})
                              .get("content", "") or "").strip()

                q = qmap.get(qid)
                if q is None:
                    raise ValueError(f"custom_id {cid} has no matching query in v2 query_index")
                modality = q["sheet"]
                pred_cells = parse_predicted_cells(content, modality)

                rec = records[(qid, strat)]
                idx = rep - 1
                rec["rep_raw"][idx] = content
                if pred_cells:
                    rec["rep_pred_cells"][idx] = pred_cells
                else:
                    rec["rep_failure"][idx] = pipeline.categorise_unparseable_response(content)
                    n_parse_fail += 1

                # Compute per-rep metrics against each reference
                for ref in REFS:
                    ref_cells = q["_ref_cells"][ref]
                    ed, j, d = metric_pair(pred_cells, ref_cells, q["landmark_type"])
                    rec["metrics"][ref]["rep_ed"][idx] = ed
                    rec["metrics"][ref]["rep_jaccard"][idx] = j
                    rec["metrics"][ref]["rep_dice"][idx] = d

                # Also keep per-rep parsed_responses for compatibility
                parsed_per_rep[rep].append({
                    "query_id": qid, "strategy": strat, "model_key": "gpt-5.4",
                    "custom_id": cid, "raw_response": content,
                    "parsed_coordinates": pipeline.parse_grid_coordinate(content) or [],
                    "modality": modality, "landmark_type": q["landmark_type"],
                    "structure": q["structure"], "image_id": q["image_id"],
                    "out_of_range": [],
                    "failure_category": rec["rep_failure"][idx],
                    "consensus_gt": q.get("consensus_gt"),
                    "source_file": jl.name,
                })
                n_lines += 1
        print(f"Rep {rep}: {n_lines} responses ({n_parse_fail} parse failures)")

    # Aggregate per-record metrics: mean across the 3 reps for each reference
    for rec in records.values():
        rec["n_failed"] = sum(1 for f in rec["rep_failure"] if f is not None)
        rec["n_valid"] = 3 - rec["n_failed"]
        for ref in REFS:
            m = rec["metrics"][ref]
            m["mean_ed"]      = mean_or_none(m["rep_ed"])
            m["mean_jaccard"] = mean_or_none(m["rep_jaccard"])
            m["mean_dice"]    = mean_or_none(m["rep_dice"])

    # ── Persist outputs ──────────────────────────────────────────────
    record_list = sorted(records.values(),
                         key=lambda r: (r["modality"], r["query_id"], r["strategy"]))

    # Strip transient parser cache from queries before any future use
    for q in qi:
        q.pop("_ref_cells", None)

    # Per-rep parsed_responses + compliance
    for rep in (1, 2, 3):
        rep_dir = sandbox / f"run{rep}"
        rep_dir.mkdir(parents=True, exist_ok=True)
        pipeline.atomic_write_json(rep_dir / "parsed_responses.json", parsed_per_rep[rep])
        compliance = pipeline.compute_compliance_stats(parsed_per_rep[rep])
        pipeline.atomic_write_json(rep_dir / "compliance_stats.json", compliance)
        n_total = len(parsed_per_rep[rep])
        n_ok = sum(1 for r in parsed_per_rep[rep] if r["parsed_coordinates"])
        print(f"  results_consensus/run{rep}/: {n_total} parsed_responses written "
              f"({n_ok} successful, {n_total - n_ok} parse-failures)")

    # Combined records pickle
    pkl_path = sandbox / "full_run_records.pkl"
    with open(pkl_path, "wb") as f:
        pickle.dump(record_list, f, protocol=pickle.HIGHEST_PROTOCOL)
    print(f"  {pkl_path}: {len(record_list)} records")

    # Top-line summary so we can see the recompute landed where expected
    n_pairs_total = len(record_list)
    n_calls_total = n_pairs_total * 3
    n_failures_total = sum(r["n_failed"] for r in record_list)
    print(f"\nSummary:")
    print(f"  pairs (query × strategy):   {n_pairs_total}")
    print(f"  rep-level calls:            {n_calls_total}")
    print(f"  parse failures (strict):    {n_failures_total}")
    print(f"  strict compliance:          "
          f"{(n_calls_total - n_failures_total) / n_calls_total * 100:.4f}%")

    # Quick mean ED summary against consensus_gt for points (sanity)
    print(f"\n  Mean ED vs consensus_gt (point landmarks, per modality × strategy):")
    by_grp = defaultdict(list)
    for r in record_list:
        if r["landmark_type"] != "point":
            continue
        v = r["metrics"]["consensus_gt"]["mean_ed"]
        if v is not None:
            by_grp[(r["modality"], r["strategy"])].append(v)
    for (mod, strat), vals in sorted(by_grp.items()):
        print(f"    {mod:13s} {strat:9s} n={len(vals):3d}  mean_ed={sum(vals)/len(vals):.4f}")

    print(f"\nDone.")


if __name__ == "__main__":
    main()
