"""Export a sanitised copy of the 900 rendered prompts to the public repo.

The source `results_full/prompts_used.json` (and Gemini equivalent) contains:
  query_id, image_id, structure, sheet, landmark_type, consensus_gt, omfr_1,
  by_strategy.{zero_shot, guided}.{system_prompt, user_prompt}

`consensus_gt` and `omfr_1` are ground-truth annotations — part of the
institutional dataset that we do NOT publish. Everything else (query
metadata + rendered prompts) is safe.

This script writes `prompts_all900.json` to the repo root, containing only
the safe fields. The rendered prompts are byte-identical to what GPT-5.4 and
Gemini 3.1 Pro actually received during the benchmark runs (verified at
orchestration time via Stage 2 byte-identity gates).

Reads:
    results_full/prompts_used.json       (GPT-5.4 main run; canonical)
Writes:
    prompts_all900.json                  (repo root, will be tracked by git)
"""
from __future__ import annotations
import json
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "results_full" / "prompts_used.json"
OUT    = ROOT / "prompts_all900.json"

SAFE_TOP_KEYS = ("sandbox", "n_queries", "strategies", "snapshot_note")
SAFE_PROMPT_KEYS = (
    "query_id", "image_id", "structure", "sheet", "landmark_type",
    "by_strategy",
)
SAFE_STRATEGY_KEYS = ("system_prompt", "user_prompt")

# Fields we explicitly drop (ground-truth annotations)
DROPPED_KEYS = ("consensus_gt", "omfr_1", "omfr_2", "student",
                "system_prompt_len", "user_prompt_len")


def sanitise_prompt_entry(entry: dict) -> dict:
    out = {}
    for k in SAFE_PROMPT_KEYS:
        if k == "by_strategy":
            strats = {}
            for strat, content in entry.get("by_strategy", {}).items():
                strats[strat] = {sk: content[sk] for sk in SAFE_STRATEGY_KEYS
                                  if sk in content}
            out["by_strategy"] = strats
        elif k in entry:
            out[k] = entry[k]
    return out


def main() -> None:
    src = json.loads(SOURCE.read_text())
    safe = {k: src[k] for k in SAFE_TOP_KEYS if k in src}
    safe["prompts"] = [sanitise_prompt_entry(p) for p in src["prompts"]]
    safe["note"] = (
        "This file lists the EXACT system + user prompts that were issued "
        "to GPT-5.4 (Responses API) and Gemini 3.1 Pro (Generative Language "
        "API v1beta) for every (query, strategy) pair in the published "
        "benchmark. The two models received byte-identical prompts; this is "
        "verified at orchestration time by the Stage 2 prompt-identity gate "
        "in scripts/run_full_run_gemini.py before any API spend. Ground-truth "
        "annotations (consensus_gt, omfr_1, etc.) are NOT included here: those "
        "fields are part of the institutional clinical dataset and are "
        "available from the corresponding author on request. The query_id "
        "and image_id fields permit cross-referencing this file with that "
        "dataset on a per-query basis."
    )

    # Schema audit: enumerate every key we are about to write and refuse
    # to proceed if any entry still contains a sensitive field.
    leaked = []
    for entry in safe["prompts"]:
        for k in entry:
            if k in DROPPED_KEYS:
                leaked.append((entry.get("query_id", "?"), k))
    if leaked:
        print("ABORT — sensitive fields detected in sanitised output:")
        for q, k in leaked[:5]:
            print(f"  {q}: {k}")
        raise SystemExit(1)

    OUT.write_text(json.dumps(safe, indent=2))
    print(f"Wrote {OUT}")
    print(f"  Entries: {len(safe['prompts'])}")
    print(f"  Strategies: {safe['strategies']}")
    print(f"  File size: {OUT.stat().st_size:,} bytes")
    print(f"  SHA-256:   {hashlib.sha256(OUT.read_bytes()).hexdigest()}")

    # Sanity sample
    print()
    print("=== Sample entry (first query) ===")
    ex = safe["prompts"][0]
    print(f"  query_id:      {ex['query_id']}")
    print(f"  image_id:      {ex['image_id']}")
    print(f"  structure:     {ex['structure']}")
    print(f"  sheet:         {ex['sheet']}")
    print(f"  landmark_type: {ex['landmark_type']}")
    print(f"  by_strategy keys: {list(ex['by_strategy'].keys())}")
    sp_len = len(ex['by_strategy']['guided']['system_prompt'])
    up_len = len(ex['by_strategy']['guided']['user_prompt'])
    print(f"  guided system_prompt: {sp_len} chars")
    print(f"  guided user_prompt:   {up_len} chars")
    print(f"  Fields NOT present (dropped): consensus_gt, omfr_1, omfr_2, student")
    # Negative-check
    for forbidden in DROPPED_KEYS:
        assert forbidden not in ex, f"Forbidden field present: {forbidden}"
    print(f"  ✓ Confirmed no sensitive fields")


if __name__ == "__main__":
    main()
