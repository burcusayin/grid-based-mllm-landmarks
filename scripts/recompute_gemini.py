"""
Re-evaluate the Gemini 3.1 Pro full-run outputs against consensus_gt, OMFR_1,
OMFR_2 and student references. Mirrors recompute_against_consensus.py (the GPT
equivalent) so analyze_consensus_run.py works with --sandbox results_full_gemini.

Reads:
    results_consensus/query_index.json                           — extended schema
                                                                    (consensus_gt,
                                                                    omfr_1/2, student)
    results_full_gemini/run{1,2,3}/responses/gemini-3.1-pro_{strategy}_chunk{C:03d}.json
                                                                 — frozen raw Gemini
                                                                    chunk outputs
    results_full_gemini/run1/responses/gemini-3.1-pro_requeries.jsonl
                                                                 — 78 rep-1 re-queries
                                                                    that recovered
                                                                    MAX_TOKENS truncations
Writes (atomic):
    results_full_gemini/full_run_records.pkl                     — per-(query, strategy)
                                                                    record with rep_raw,
                                                                    rep_pred_cells,
                                                                    rep_failure, and
                                                                    metrics (ED, Jaccard,
                                                                    Dice) against each
                                                                    reference
    results_full_gemini/gemini_recompute_anchor.json             — SHA-256 of every
                                                                    chunk file + the
                                                                    re-queries jsonl
                                                                    consumed by this
                                                                    recompute, so any
                                                                    future re-run can
                                                                    detect drift.

No API calls. No writes outside results_full_gemini/.

Usage:
    .venv/bin/python scripts/recompute_gemini.py
        [--sandbox results_full_gemini]
        [--query-index results_consensus/query_index.json]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import pickle
import sys
from pathlib import Path
from collections import defaultdict

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import config  # noqa: E402
import pipeline  # noqa: E402

# Import the geometry + metric helpers from the GPT recompute script
# (they are pure functions with no GPT-specific state).
sys.path.insert(0, str(ROOT / "scripts"))
from recompute_against_consensus import (  # noqa: E402
    parse_reference_cells,
    parse_predicted_cells,
    metric_pair,
    mean_or_none,
    sha256_file,
)


MODEL_KEY = "gemini-3.1-pro"
STRATEGIES = ("zero_shot", "guided")
REFS = ("consensus_gt", "omfr_1", "omfr_2", "student")


def load_chunks_for_rep(rep_dir: Path) -> dict[str, str]:
    """Return {custom_id → raw_text} from all Google chunk JSON files in a rep dir.

    Each chunk file is a JSON list of {"custom_id": ..., "response": <GCContent>}.
    The raw text lives at response.candidates[0].content.parts[0].text. If the
    response is missing (e.g., per-request error), we record empty string —
    matching cmd_parse's behaviour.

    Also fills in re-queries from gemini-3.1-pro_requeries.jsonl if present
    (rep 1 only). Re-queries OVERWRITE any chunk-derived text for the same
    custom_id, mirroring the source-of-truth precedence used by the
    operational pipeline that produced parsed_responses.json.
    """
    cid_to_text: dict[str, str | None] = {}
    responses_dir = rep_dir / "responses"

    # ── Phase 1: read every chunk JSON file ────────────────────────
    for chunk_path in sorted(responses_dir.glob(f"{MODEL_KEY}_*_chunk*.json")):
        data = json.loads(chunk_path.read_text())
        for entry in data:
            cid = entry.get("custom_id", "")
            resp = entry.get("response") or {}
            cands = resp.get("candidates", []) or []
            # When the response failed at Google's side (e.g., deadline
            # expired, _error block present, no candidates), the operational
            # cmd_parse stores raw_response=None. We match that exactly so
            # the recomputed parsed_responses is byte-identical.
            text: str | None = None
            if cands and isinstance(cands[0], dict):
                content = cands[0].get("content", {}) or {}
                parts = content.get("parts", []) or []
                if parts and isinstance(parts[0], dict):
                    text = parts[0].get("text", "") or ""
            cid_to_text[cid] = text

    # ── Phase 2: merge re-queries (rep 1 only — file absent for others) ─
    requeries_path = responses_dir / f"{MODEL_KEY}_requeries.jsonl"
    n_requeries = 0
    if requeries_path.exists():
        for line in open(requeries_path):
            line = line.strip()
            if not line:
                continue
            obj = json.loads(line)
            cid = obj["custom_id"]
            cid_to_text[cid] = obj.get("raw_response", "") or ""
            n_requeries += 1
    return cid_to_text, n_requeries


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sandbox", default="results_full_gemini",
                    help="Gemini sandbox dir (default: results_full_gemini)")
    ap.add_argument("--query-index", default="results_consensus/query_index.json",
                    help="Extended query_index.json with consensus_gt + omfr_1/2 + student")
    args = ap.parse_args()

    sandbox = (ROOT / args.sandbox).resolve()
    qi_path = (ROOT / args.query_index).resolve()
    if not sandbox.exists():
        print(f"ERROR: sandbox {sandbox} not found", file=sys.stderr)
        sys.exit(1)
    if not qi_path.exists():
        print(f"ERROR: query_index {qi_path} not found", file=sys.stderr)
        sys.exit(1)

    qi = json.loads(qi_path.read_text())
    qmap = {q["query_id"]: q for q in qi}
    print(f"Loaded {len(qi)} queries from {qi_path}")

    # Pre-parse all reference cell sets once per query
    for q in qi:
        mod = q["sheet"]
        q["_ref_cells"] = {
            "consensus_gt": parse_reference_cells(q.get("consensus_gt"), mod),
            "omfr_1":       parse_reference_cells(q.get("omfr_1"), mod),
            "omfr_2":       parse_reference_cells(q.get("omfr_2"), mod),
            "student":      parse_reference_cells(q.get("student"), mod),
        }

    # Initialise the (query_id × strategy) record table — mirrors GPT schema
    records: dict[tuple[str, str], dict] = {}
    for q in qi:
        for strat in STRATEGIES:
            records[(q["query_id"], strat)] = {
                "query_id": q["query_id"],
                "strategy": strat,
                "modality": q["sheet"],
                "image_id": q["image_id"],
                "structure": q["structure"],
                "landmark_type": q["landmark_type"],
                "uses_fdi": q.get("uses_fdi", False),
                "consensus_gt":  q.get("consensus_gt"),
                "omfr_1":        q.get("omfr_1"),
                "omfr_1_second": q.get("omfr_1_second"),
                "omfr_2":        q.get("omfr_2"),
                "omfr_2_second": q.get("omfr_2_second"),
                "student":       q.get("student"),
                "rep_raw":        [None, None, None],
                "rep_pred_cells": [None, None, None],
                "rep_failure":    [None, None, None],
                "metrics": {ref: {"rep_ed":      [None]*3,
                                   "rep_jaccard": [None]*3,
                                   "rep_dice":    [None]*3} for ref in REFS},
            }

    # Anchor: SHA-256 every input file we consume, for reproducibility
    anchor = {
        "query_index_sha256": sha256_file(qi_path),
        "query_index_path": str(qi_path),
        "reps": {},
    }

    parsed_per_rep = {1: [], 2: [], 3: []}
    total_lines = 0
    total_failures = 0

    for rep in (1, 2, 3):
        rep_dir = sandbox / f"run{rep}"
        if not rep_dir.exists():
            print(f"ERROR: missing {rep_dir}", file=sys.stderr)
            sys.exit(1)

        # Anchor this rep's chunks + requeries
        responses_dir = rep_dir / "responses"
        rep_anchor = {"chunks": {}}
        for chunk_path in sorted(responses_dir.glob(f"{MODEL_KEY}_*_chunk*.json")):
            rep_anchor["chunks"][chunk_path.name] = {
                "sha256": sha256_file(chunk_path),
                "size": chunk_path.stat().st_size,
            }
        rq_path = responses_dir / f"{MODEL_KEY}_requeries.jsonl"
        if rq_path.exists():
            rep_anchor["requeries_jsonl"] = {
                "sha256": sha256_file(rq_path),
                "size": rq_path.stat().st_size,
            }
        anchor["reps"][f"run{rep}"] = rep_anchor

        # Load raw text per cid, with re-queries overlaid
        cid_to_text, n_rq = load_chunks_for_rep(rep_dir)
        if n_rq:
            print(f"Rep {rep}: merged {n_rq} re-queries (rep 1 expected)")

        # Precompute the set of cids that came from the re-queries file
        # (used to label source_file in recomputed parsed_responses; cheap
        # one-time pass instead of per-record re-open).
        rq_cids: set[str] = set()
        if rq_path.exists():
            for L in open(rq_path):
                L = L.strip()
                if L:
                    rq_cids.add(json.loads(L)["custom_id"])

        # Build records
        rep_failures = 0
        for cid, raw_text in cid_to_text.items():
            if cid.endswith("_zero_shot"):
                qid, strat = cid[:-len("_zero_shot")], "zero_shot"
            elif cid.endswith("_guided"):
                qid, strat = cid[:-len("_guided")], "guided"
            else:
                raise ValueError(f"Unrecognised custom_id format: {cid!r}")

            q = qmap.get(qid)
            if q is None:
                raise ValueError(f"custom_id {cid} has no matching query in {qi_path.name}")
            modality = q["sheet"]
            lm_type = q["landmark_type"]
            pred_cells = parse_predicted_cells(raw_text, modality)

            # Enforce _finalise_record's strict rules so the recompute matches
            # the operational pipeline exactly:
            #
            #   Point landmark + raw parser returned >1 cell → "ambiguous"
            #     (instruction violation; we don't second-guess which cell
            #     was the model's real answer).
            #   Single parsed cell that is out-of-range → "out_of_range"
            #     (the model meant a coordinate, but outside the grid).
            #   No cells parsed at all → derive from raw text
            #     (refusal / ambiguous / verbose / no_engage).
            raw_parsed = pipeline.parse_grid_coordinate(raw_text or "")
            failure_cat: str | None = None
            if not raw_parsed:
                # Branch 1: nothing parsed at all
                failure_cat = pipeline.categorise_unparseable_response(raw_text)
                pred_cells = []
            elif lm_type == "point" and len(raw_parsed) > 1:
                # Branch 3a: multi-cell response on point landmark
                failure_cat = "ambiguous"
                pred_cells = []
            elif lm_type == "point" and not pred_cells:
                # Branch 3b: single parsed cell, but it failed range validation
                failure_cat = "out_of_range"
                pred_cells = []
            elif lm_type == "area" and not pred_cells:
                # Branch 4: all area cells were out of range
                failure_cat = "out_of_range"
                pred_cells = []

            rec = records[(qid, strat)]
            idx = rep - 1
            rec["rep_raw"][idx] = raw_text
            if failure_cat is None:
                rec["rep_pred_cells"][idx] = pred_cells
            else:
                rec["rep_failure"][idx] = failure_cat
                rep_failures += 1
                total_failures += 1

            # Per-rep metrics for every reference. metric_pair returns
            # (None, None, None) when pred_cells is empty — so failures
            # automatically contribute no metric (which is what we want).
            for ref in REFS:
                ref_cells = q["_ref_cells"][ref]
                ed, j, d = metric_pair(pred_cells, ref_cells, lm_type)
                rec["metrics"][ref]["rep_ed"][idx] = ed
                rec["metrics"][ref]["rep_jaccard"][idx] = j
                rec["metrics"][ref]["rep_dice"][idx] = d

            # Per-rep parsed_responses (for parity with the operational
            # pipeline output: parsed_coordinates is set ONLY when the
            # record passed all of _finalise_record's rules — i.e., when
            # failure_category is None).
            if failure_cat is None:
                # Convert (col, row) back to "X{n}" strings for parity
                parsed_strs = [
                    f"{chr(ord('A') + (c[1] - 1))}{c[0]}" for c in (pred_cells or [])
                ]
            else:
                parsed_strs = []
            parsed_per_rep[rep].append({
                "query_id": qid, "strategy": strat, "model_key": MODEL_KEY,
                "custom_id": cid, "raw_response": raw_text,
                "parsed_coordinates": parsed_strs,
                "modality": modality, "landmark_type": q["landmark_type"],
                "structure": q["structure"], "image_id": q["image_id"],
                "out_of_range": [],
                "failure_category": failure_cat,
                "consensus_gt": q.get("consensus_gt"),
                "source_file": (
                    f"{MODEL_KEY}_requeries.jsonl" if cid in rq_cids
                    else "chunk"
                ),
            })
            total_lines += 1

        print(f"Rep {rep}: {len(cid_to_text)} responses ({rep_failures} parse-failures)")

    # Aggregate per-record metrics: mean across the 3 reps for each reference
    for rec in records.values():
        rec["n_failed"] = sum(1 for f in rec["rep_failure"] if f is not None)
        rec["n_valid"] = 3 - rec["n_failed"]
        for ref in REFS:
            m = rec["metrics"][ref]
            m["mean_ed"]      = mean_or_none(m["rep_ed"])
            m["mean_jaccard"] = mean_or_none(m["rep_jaccard"])
            m["mean_dice"]    = mean_or_none(m["rep_dice"])

    # Persist outputs
    record_list = sorted(records.values(),
                         key=lambda r: (r["modality"], r["query_id"], r["strategy"]))

    # Strip transient parser cache from queries before any future use
    for q in qi:
        q.pop("_ref_cells", None)

    # Write anchor
    pipeline.atomic_write_json(sandbox / "gemini_recompute_anchor.json", anchor)
    n_chunks_anchored = sum(len(rep["chunks"]) for rep in anchor["reps"].values())
    print(f"  Anchored {n_chunks_anchored} chunk files + 1 re-queries jsonl "
          f"→ {sandbox / 'gemini_recompute_anchor.json'}")

    # Per-rep parsed_responses (these match what the operational pipeline
    # already produced, but we DO NOT overwrite operational files — write
    # alongside as parsed_responses.recomputed.json so the comparison is
    # auditable without disturbing the original).
    for rep in (1, 2, 3):
        rep_dir = sandbox / f"run{rep}"
        out_path = rep_dir / "parsed_responses.recomputed.json"
        pipeline.atomic_write_json(out_path, parsed_per_rep[rep])
        n = len(parsed_per_rep[rep])
        n_ok = sum(1 for r in parsed_per_rep[rep] if r["parsed_coordinates"])
        print(f"  {rep_dir.name}/parsed_responses.recomputed.json: {n} records "
              f"({n_ok} parsed OK, {n - n_ok} failures)")

    # Combined records pickle (the input to analyze_consensus_run.py)
    pkl_path = sandbox / "full_run_records.pkl"
    with open(pkl_path, "wb") as f:
        pickle.dump(record_list, f, protocol=pickle.HIGHEST_PROTOCOL)
    print(f"  {pkl_path}: {len(record_list)} records (model={MODEL_KEY})")

    # Quick summary
    print()
    print(f"Total raw responses processed: {total_lines}")
    print(f"Total parse failures: {total_failures}  (expected: ~3 across 5400)")
    n_complete = sum(1 for r in records.values() if r["n_valid"] == 3)
    print(f"Records with all 3 reps valid: {n_complete}/{len(records)}")


if __name__ == "__main__":
    main()
