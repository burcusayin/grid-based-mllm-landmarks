"""
Snapshot the v1 and v2 rendered prompts for every query in the pilot subset.

Writes results_pilot_v2/v1_prompts_snapshot.json and v2_prompts_snapshot.json,
plus a diff summary that proves the only difference between the two is the
L-R clause on PANORAMIC guided prompts.

Read-only: makes no API calls, never modifies query_index files.
"""
from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import config  # noqa: E402
import pipeline  # noqa: E402

V2_DIR = ROOT / "results_pilot_v2"

# v1 GUIDED_SYSTEM_ADDITION as it was before commit 3c80bf5 (no
# {modality_clause} placeholder). Reconstructed from git history. The
# rest of config (templates, landmark descriptions, etc.) is unchanged
# between v1 and v2.
V1_GUIDED_SYSTEM_TEMPLATE = (
    "The image has a grid overlay with columns numbered 1 through {cols} from left to right, "
    "and rows labeled A through {max_row} from top to bottom. "
    "Each cell is identified by its row letter followed by its column number, "
    "with no space or punctuation between them. "
    "The grid lines are drawn in cyan and labels are in yellow. "
    "For point-based questions, respond with exactly one cell coordinate. "
    "For area-based questions, list all cells the structure occupies, separated by commas."
)


def render_v1_prompt(query: dict, strategy: str) -> tuple[str, str]:
    """Render the v1 prompt for one query+strategy combination."""
    sheet = query["sheet"]
    grid = config.GRID_SPECS[sheet]
    lm_type = query["landmark_type"]
    desc = query["landmark_description_en"]
    uses_fdi = query.get("uses_fdi", False)
    fdi_prefix = "Using the FDI two-digit numbering system, " if uses_fdi else ""
    identify = "identify" if uses_fdi else "Identify"

    if strategy == "zero_shot":
        # Zero-shot is unchanged between v1 and v2
        system_prompt = config.SYSTEM_PROMPT
        if lm_type == "point":
            user_prompt = config.ZERO_SHOT_POINT_TEMPLATE.format(
                cols=grid["cols"], rows=grid["rows"],
                landmark_description=desc,
                fdi_prefix=fdi_prefix, Identify=identify,
            )
        else:
            user_prompt = config.ZERO_SHOT_AREA_TEMPLATE.format(
                cols=grid["cols"], rows=grid["rows"],
                landmark_description=desc,
                fdi_prefix=fdi_prefix, Identify=identify,
            )
    elif strategy == "guided":
        # v1: no modality clause, no L-R sentence
        v1_grid_explanation = V1_GUIDED_SYSTEM_TEMPLATE.format(
            cols=grid["cols"], max_row=grid["max_row_letter"])
        system_prompt = config.SYSTEM_PROMPT + "\n\n" + v1_grid_explanation
        if lm_type == "point":
            user_prompt = config.GUIDED_POINT_TEMPLATE.format(
                landmark_description=desc,
                fdi_prefix=fdi_prefix, Identify=identify,
            )
        else:
            user_prompt = config.GUIDED_AREA_TEMPLATE.format(
                landmark_description=desc,
                fdi_prefix=fdi_prefix, Identify=identify,
            )
    else:
        raise ValueError(f"Unknown strategy: {strategy}")
    return system_prompt, user_prompt


def snapshot(version: str, render_fn) -> dict:
    """Render every prompt for every (query, strategy) and return a snapshot."""
    queries = json.loads((V2_DIR / "query_index.json").read_text())
    entries = []
    for q in queries:
        for strat in ("zero_shot", "guided"):
            sys_p, usr_p = render_fn(q, strat)
            full = sys_p + "\n---USER---\n" + usr_p
            entries.append({
                "query_id": q["query_id"],
                "strategy": strat,
                "sheet": q["sheet"],
                "landmark_type": q["landmark_type"],
                "system_prompt": sys_p,
                "user_prompt": usr_p,
                "system_sha256": hashlib.sha256(sys_p.encode()).hexdigest(),
                "user_sha256": hashlib.sha256(usr_p.encode()).hexdigest(),
                "full_sha256": hashlib.sha256(full.encode()).hexdigest(),
            })
    out = {
        "version": version,
        "n_queries": len(queries),
        "n_prompt_instances": len(entries),
        "entries": entries,
    }
    return out


def main() -> int:
    print("Rendering v1 prompts (reconstructed from git history)...")
    v1 = snapshot("v1", render_v1_prompt)
    print(f"  {len(v1['entries'])} prompt instances rendered")

    print("\nRendering v2 prompts (current config.py)...")
    v2 = snapshot("v2", pipeline.generate_prompt)
    print(f"  {len(v2['entries'])} prompt instances rendered")

    # Build diff summary
    print("\nComputing v1 vs v2 diff...")
    v1_by_key = {(e["query_id"], e["strategy"]): e for e in v1["entries"]}
    v2_by_key = {(e["query_id"], e["strategy"]): e for e in v2["entries"]}

    same = 0
    diff = 0
    diffs_per_bucket = {}
    diff_details = []
    for key in sorted(v1_by_key.keys()):
        e1 = v1_by_key[key]
        e2 = v2_by_key[key]
        bucket = f"{e1['sheet']}_{e1['strategy']}"
        if e1["full_sha256"] == e2["full_sha256"]:
            same += 1
        else:
            diff += 1
            diffs_per_bucket[bucket] = diffs_per_bucket.get(bucket, 0) + 1
            char_delta = len(e2["system_prompt"]) - len(e1["system_prompt"])
            diff_details.append({"query_id": key[0], "strategy": key[1],
                                  "bucket": bucket,
                                  "system_char_delta": char_delta})

    print(f"  identical: {same}/{len(v1_by_key)}")
    print(f"  different: {diff}/{len(v1_by_key)}")
    print(f"  diffs per bucket: {diffs_per_bucket}")
    if diff_details:
        deltas = sorted({d["system_char_delta"] for d in diff_details})
        print(f"  unique char-deltas in system prompt: {deltas}")

    # Write the three files
    v1_path = V2_DIR / "v1_prompts_snapshot.json"
    v2_path = V2_DIR / "v2_prompts_snapshot.json"
    diff_path = V2_DIR / "v1_v2_prompt_diff.json"

    v1_path.write_text(json.dumps(v1, indent=2))
    v2_path.write_text(json.dumps(v2, indent=2))

    diff_summary = {
        "n_total": len(v1_by_key),
        "n_identical": same,
        "n_different": diff,
        "diffs_per_bucket": diffs_per_bucket,
        "unique_system_char_deltas": sorted({d["system_char_delta"] for d in diff_details}) if diff_details else [],
        "expected_diffs": "30 PANORAMIC_guided prompts, all with +158 char delta (L-R sentence)",
        "details": diff_details,
    }
    diff_path.write_text(json.dumps(diff_summary, indent=2))

    # Aggregate hashes for the snapshot files themselves
    print(f"\nFile hashes (anchor for the audit trail):")
    for p in (v1_path, v2_path, diff_path):
        h = hashlib.sha256(p.read_bytes()).hexdigest()
        print(f"  {h}  {p.name}")

    # Sanity assertions
    assert same + diff == 120, f"expected 120 prompt instances, got {same + diff}"
    expected_diffs = {"PANORAMIC_guided": 30}
    assert diffs_per_bucket == expected_diffs, \
        f"unexpected diff buckets: got {diffs_per_bucket} expected {expected_diffs}"
    assert sorted({d['system_char_delta'] for d in diff_details}) == [158], \
        f"unexpected char-delta values"

    print(f"\n✓ Snapshot complete. Audit trail written to {V2_DIR}/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
