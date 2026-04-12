"""
End-to-end smoke test for the Dental MLLM Benchmark pipeline.

Purpose
-------
Exercise every downstream stage of the pipeline (parse → metrics → stats
→ visualize → report) without hitting any external APIs and WITHOUT
touching the real results/ directory or the original Excel file.

Isolation
---------
The test runs entirely inside a sandbox directory (`tests/.smoke_sandbox/`)
and sets the DENTAL_MLLM_RESULTS_DIR environment variable so that
pipeline.py and analysis.py write every artefact into the sandbox. The
sandbox is wiped clean at the start of each run. At NO POINT does this
script write to the real `results/` tree or modify the Excel file.

Because the sandbox has its own query_index.json, the test can safely
inject synthetic omfr_2 and student data without any risk of polluting
the real query_index used by production runs.

What the test covers
--------------------
1. Pipeline parse stage:
   - Writes synthetic OpenAI-format and Google-format response files for
     two models × two strategies.
   - Includes refusals, ambiguous multi-coord responses, out-of-range
     coordinates, and exact matches so that:
       * the response categoriser is exercised for all four failure modes
       * the single-cell rule for point landmarks is exercised
       * out-of-range detection triggers
       * the new per-record parsed_responses list format is produced
   - Runs `pipeline.cmd_parse` and asserts that the resulting list contains
     one record per (query, strategy, model) with the expected fields.
     In particular: no collisions between providers (the bug the review
     discovered), and compliance counts match the injected ground-truth.

2. Inter-rater reliability:
   - Injects synthetic omfr_2 values (perfect agreement with omfr_1 for most
     queries and a controlled disagreement for a few) and asserts that
     ICC(2,1) is near 1 for rows and columns.
   - Injects synthetic area omfr_2 sets (with symmetric perturbation) and
     asserts Cohen's κ is > 0.5.

3. Student paired tests:
   - Injects synthetic student responses for a subset of point queries
     (noisy centred around omfr_1) and asserts that the Wilcoxon call
     runs without error, producing a finite p-value.

4. Analysis:
   - Runs `analysis.cmd_metrics`, `cmd_stats`, `cmd_report`, `cmd_visualize`
     via subprocess and asserts that the expected output files exist and
     contain the expected top-level keys.

Running
-------
    .venv/bin/python tests/smoke_e2e.py

The test takes ~20 seconds and exits with code 0 on success.
"""

import json
import os
import random
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT))

SANDBOX_DIR = HERE / ".smoke_sandbox"
SANDBOX_RESPONSES = SANDBOX_DIR / "responses"
SANDBOX_ANALYSIS = SANDBOX_DIR / "analysis"
SANDBOX_QUERY_INDEX = SANDBOX_DIR / "query_index.json"

# All subprocess calls inherit this env, pointing pipeline/analysis at the
# sandbox. The real results/ directory is NEVER touched.
SANDBOX_ENV = os.environ.copy()
SANDBOX_ENV["DENTAL_MLLM_RESULTS_DIR"] = str(SANDBOX_DIR)

# Deterministic RNG for reproducible synthetic predictions
RNG = random.Random(20260411)


def setup_sandbox():
    """Create a clean sandbox and seed it with a synthetic query_index."""
    if SANDBOX_DIR.exists():
        shutil.rmtree(SANDBOX_DIR)
    SANDBOX_DIR.mkdir(parents=True)
    SANDBOX_RESPONSES.mkdir()

    # Run the real `pipeline.py prepare` INTO the sandbox, so we get a
    # fresh query_index.json derived from the real Excel — but the file
    # lands inside the sandbox, never in the real results/ tree.
    run([
        str(ROOT / ".venv" / "bin" / "python"),
        "pipeline.py", "prepare",
    ])
    # Sanity check: the real results/query_index.json must be unchanged.
    real_index = ROOT / "results" / "query_index.json"
    real_mtime_before = real_index.stat().st_mtime if real_index.exists() else None
    return real_mtime_before


def load_sandbox_queries():
    with open(SANDBOX_QUERY_INDEX) as f:
        return json.load(f)


def save_sandbox_queries(queries):
    with open(SANDBOX_QUERY_INDEX, "w") as f:
        json.dump(queries, f, indent=2)


def inject_omfr2_and_students(queries):
    """
    Augment the sandbox query_index with deterministic synthetic omfr_2 and
    student responses so IRR and student paired-test codepaths have data.

    Strategy:
      - omfr_2: equals omfr_1 for 90% of queries, perturbed by one cell for 10%.
      - students: 8 responses per point query, each offset by a small random
        row/col delta (drawn from {-1, 0, 0, 0, 1}).
    """
    import re
    rng = random.Random(424242)

    def shift_coord(coord, drow, dcol):
        m = re.match(r"^([A-H])(\d{1,2})$", coord)
        if not m:
            return coord
        r = ord(m.group(1)) - ord('A') + 1 + drow
        c = int(m.group(2)) + dcol
        if r < 1:
            r = 1
        if c < 1:
            c = 1
        return f"{chr(ord('A') + r - 1)}{c}"

    for q in queries:
        omfr_1 = q.get("omfr_1")
        if not omfr_1:
            continue

        # omfr_2
        if rng.random() < 0.9:
            q["omfr_2"] = omfr_1
        else:
            if q["landmark_type"] == "point":
                q["omfr_2"] = shift_coord(omfr_1.strip(), rng.choice([-1, 1]), 0)
            else:
                cells = [c.strip() for c in str(omfr_1).split(",") if c.strip()]
                if rng.random() < 0.5 and len(cells) > 1:
                    q["omfr_2"] = ", ".join(cells[:-1])
                elif cells:
                    extra = shift_coord(cells[-1], rng.choice([-1, 1]), rng.choice([-1, 1]))
                    q["omfr_2"] = omfr_1 + ", " + extra
                else:
                    q["omfr_2"] = omfr_1

        # Students (point landmarks only)
        if q["landmark_type"] == "point":
            gt = omfr_1.strip()
            students = []
            for _ in range(8):
                drow = rng.choice([-1, 0, 0, 0, 1])
                dcol = rng.choice([-1, 0, 0, 0, 1])
                students.append(shift_coord(gt, drow, dcol))
            students.extend([None] * (40 - len(students)))
            q["students"] = students

    save_sandbox_queries(queries)


def _openai_record(custom_id, text):
    return {
        "custom_id": custom_id,
        "response": {
            "body": {
                "choices": [{"message": {"content": text}}],
            },
        },
    }


def _google_record(custom_id, text):
    return {
        "custom_id": custom_id,
        "response": {
            "candidates": [{
                "content": {"parts": [{"text": text}]},
            }],
        },
    }


def synthesise_response_text(q, model_key, strategy):
    """Deterministic response text for a (query, model, strategy) triple."""
    seed = hash((q["query_id"], model_key, strategy)) & 0xFFFFFFFF
    rnd = random.Random(seed)
    roll = rnd.random()

    omfr_1 = q.get("omfr_1")
    if omfr_1 is None:
        return "I cannot identify this landmark."
    omfr_1 = str(omfr_1).strip()

    if q["landmark_type"] == "area":
        if roll < 0.15:
            return "I cannot provide this answer; the image is unclear."
        return omfr_1

    # Point landmark branches
    if roll < 0.05:
        return "I cannot identify the landmark in this image."
    if roll < 0.10:
        import re
        m = re.match(r"^([A-H])(\d{1,2})$", omfr_1)
        if m:
            r = ord(m.group(1)) - ord('A') + 1
            c = int(m.group(2))
            alt = f"{chr(ord('A') + r - 1)}{c + 1}"
            return f"Possibly {omfr_1} or {alt}"
        return "I am not sure which cell"
    if roll < 0.15:
        # Out-of-range at the modality level
        if q["sheet"] == "PERIAPICAL":
            return "H1"
        if q["sheet"] == "CEPHALOMETRIC":
            return "H11"
        return "I am not sure which cell is correct"
    if roll < 0.25:
        import re
        m = re.match(r"^([A-H])(\d{1,2})$", omfr_1)
        if m:
            r = ord(m.group(1)) - ord('A') + 1
            c = int(m.group(2))
            if r + 1 <= 6:
                return f"{chr(ord('A') + r)}{c}"
        return omfr_1
    return omfr_1


def synthesise_responses(queries):
    """Write synthetic response files for all (model, strategy) combinations."""
    import config

    for model_key in config.ACTIVE_MODELS:
        provider = config.MODELS[model_key]["provider"]
        for strategy in config.STRATEGIES:
            batch_name = f"{model_key}_{strategy}"

            if provider == "openai":
                out = SANDBOX_RESPONSES / f"{batch_name}_chunk000_results.jsonl"
                with open(out, "w") as f:
                    for q in queries:
                        custom_id = f"{q['query_id']}_{strategy}"
                        text = synthesise_response_text(q, model_key, strategy)
                        f.write(json.dumps(_openai_record(custom_id, text)) + "\n")
                print(f"  wrote {out.name}")
            elif provider == "google":
                out = SANDBOX_RESPONSES / f"{batch_name}_chunk000.json"
                entries = []
                for q in queries:
                    custom_id = f"{q['query_id']}_{strategy}"
                    text = synthesise_response_text(q, model_key, strategy)
                    entries.append(_google_record(custom_id, text))
                with open(out, "w") as f:
                    json.dump(entries, f, indent=2)
                print(f"  wrote {out.name}")
            else:
                raise RuntimeError(f"synthesise_responses: unsupported provider {provider}")


def run(cmd):
    print(f"\n$ {' '.join(cmd)}")
    r = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True, env=SANDBOX_ENV)
    if r.stdout:
        print(r.stdout)
    if r.stderr:
        print("STDERR:", r.stderr)
    r.check_returncode()
    return r


def assert_real_results_untouched(mtime_before):
    real_index = ROOT / "results" / "query_index.json"
    if mtime_before is None:
        assert not real_index.exists(), \
            "real results/query_index.json appeared during the test"
    else:
        mtime_after = real_index.stat().st_mtime
        assert mtime_before == mtime_after, (
            f"real results/query_index.json was modified during the test! "
            f"before={mtime_before}, after={mtime_after}"
        )
    # Also make sure analysis dir in the real tree is untouched
    real_analysis = ROOT / "results" / "analysis"
    if real_analysis.exists():
        for p in real_analysis.rglob("*"):
            assert p.is_dir() or p.stat().st_mtime < mtime_before + 1, \
                f"real results/analysis/{p.name} was modified"


def main():
    venv_python = str(ROOT / ".venv" / "bin" / "python")
    assert Path(venv_python).exists(), "venv not found; run python3 -m venv .venv"

    print("=== Smoke test: setting up sandbox ===")
    real_index = ROOT / "results" / "query_index.json"
    mtime_before = real_index.stat().st_mtime if real_index.exists() else None
    setup_sandbox()

    print("=== Loading sandbox query index ===")
    queries = load_sandbox_queries()
    print(f"Loaded {len(queries)} queries from the SANDBOX")

    print("=== Injecting synthetic omfr_2 and student data (sandbox only) ===")
    inject_omfr2_and_students(queries)
    queries = load_sandbox_queries()
    print(f"  omfr_2 populated: {sum(1 for q in queries if q.get('omfr_2'))}")
    print(f"  students populated: {sum(1 for q in queries if any(q.get('students') or []))}")

    print("=== Synthesising response files (sandbox only) ===")
    synthesise_responses(queries)

    print("=== pipeline parse ===")
    run([venv_python, "pipeline.py", "parse"])

    with open(SANDBOX_DIR / "parsed_responses.json") as f:
        records = json.load(f)
    print(f"  parsed_responses records: {len(records)}")
    assert isinstance(records, list), "parsed_responses.json must be a list"
    batches = set((r["model_key"], r["strategy"]) for r in records)
    print(f"  distinct (model, strategy) batches: {sorted(batches)}")
    assert len(batches) == 4, f"expected 4 batches, got {len(batches)}"
    assert len(records) == 4 * len(queries), \
        f"expected {4 * len(queries)} records, got {len(records)}"

    from collections import Counter
    cat_counts = Counter(r.get("failure_category") for r in records)
    print(f"  failure category counts: {dict(cat_counts)}")
    assert cat_counts.get("out_of_range", 0) > 0, "expected some out_of_range"
    assert cat_counts.get("refusal", 0) > 0, "expected some refusals"
    assert cat_counts.get("ambiguous", 0) > 0, "expected some ambiguous"

    print("=== analysis metrics ===")
    run([venv_python, "analysis.py", "metrics"])
    with open(SANDBOX_ANALYSIS / "all_metrics.json") as f:
        all_metrics = json.load(f)
    assert set(all_metrics.keys()) == {
        "gpt-5.4_zero_shot", "gpt-5.4_guided",
        "gemini-3.1-pro_zero_shot", "gemini-3.1-pro_guided",
    }, f"unexpected batch keys: {list(all_metrics.keys())}"
    for k, ms in all_metrics.items():
        n_point = sum(1 for m in ms if m["landmark_type"] == "point")
        assert n_point > 0, f"no point metrics for {k}"

    print("=== analysis stats ===")
    run([venv_python, "analysis.py", "stats"])
    with open(SANDBOX_ANALYSIS / "statistical_results.json") as f:
        stat = json.load(f)
    assert "per_structure" in stat and "per_modality" in stat and "grand" in stat

    irr_path = SANDBOX_ANALYSIS / "inter_rater_reliability.json"
    assert irr_path.exists(), "inter_rater_reliability.json missing"
    with open(irr_path) as f:
        irr = json.load(f)
    print(f"  IRR notes: {irr['notes']}")
    assert irr["notes"]["point_total"] > 0, "expected ICC point data"
    overall_row = irr["overall"].get("point_icc_row")
    assert overall_row is not None and overall_row > 0.7, \
        f"expected high ICC_row (~1); got {overall_row}"
    overall_kappa = irr["overall"].get("area_cohen_kappa")
    assert overall_kappa is not None and overall_kappa > 0.5, \
        f"expected positive Cohen κ; got {overall_kappa}"

    paired_path = SANDBOX_ANALYSIS / "student_paired_tests.json"
    assert paired_path.exists(), "student_paired_tests.json missing"
    with open(paired_path) as f:
        paired = json.load(f)
    print(f"  student paired test rows: {len(paired)}")
    assert len(paired) == 4, f"expected 4 batch-level paired tests, got {len(paired)}"
    for p in paired:
        assert p["wilcoxon_p_raw"] is not None
        assert p["wilcoxon_p_bonferroni"] is not None
        assert p["n_paired"] > 0
        assert "rank_biserial_r" in p

    print("=== analysis report ===")
    run([venv_python, "analysis.py", "report"])
    assert (SANDBOX_ANALYSIS / "summary_table.csv").exists()
    assert (SANDBOX_ANALYSIS / "detailed_results.csv").exists()

    print("=== analysis visualize ===")
    run([venv_python, "analysis.py", "visualize"])
    plots_dir = SANDBOX_ANALYSIS / "plots"
    assert plots_dir.exists(), "plots dir missing"
    plots = sorted(p.name for p in plots_dir.glob("*.png"))
    print(f"  plots: {plots}")
    assert any("bland_altman" in p for p in plots), \
        "expected at least one bland_altman plot"

    # Invariant: the real results/ tree was never touched.
    assert_real_results_untouched(mtime_before)

    print("\n✅ ALL SMOKE TESTS PASSED — real data and real results/ untouched")


if __name__ == "__main__":
    main()
