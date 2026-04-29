"""
Dental MLLM Benchmark — Pilot Orchestrator

Purpose
-------
Run a small, low-cost preliminary experiment that exercises every dimension
of the full pipeline (prompt format compliance, grid methodology, model
comparison, and temperature=0 determinism) while keeping the API spend to
a few dollars. The orchestrator also computes Fleiss' kappa across 3
repetitions to formally resolve Issue 1 (consistency verification) from
the pre-experiment review.

Design goals
------------
1.  Every real API call is gated behind a printed cost estimate and an
    explicit yes/no confirmation. There is no way to accidentally spend
    money.
2.  The orchestrator is IDEMPOTENT. Re-running after an interruption picks
    up where the previous run left off (the tracking file + already-
    downloaded responses are respected).
3.  Every artefact (query index, batch tracking, responses, parsed results,
    analysis) lives under a sandbox directory (`results_pilot/` by default).
    The real `results/` tree is never touched.
4.  Pre-flight checks: API keys present, image files valid, enough budget,
    smoke test can be run (optional), etc. If any check fails the
    orchestrator exits BEFORE any real API call.

Usage
-----
    # Default: 5 images per modality, GPT-only, 3 repetitions, sandbox at
    # results_pilot/ in the project root.
    .venv/bin/python scripts/run_pilot.py

    # Override: 3 repetitions, 3 images per modality, change sandbox dir:
    .venv/bin/python scripts/run_pilot.py --n-per-modality 3 --repetitions 3 --sandbox results_pilot_v2

    # Bigger pilot with both models:
    .venv/bin/python scripts/run_pilot.py --models gpt-5.4,gemini-3.1-pro

    # Dry run (no real API calls, useful for testing the orchestrator):
    .venv/bin/python scripts/run_pilot.py --dry-run

Stages
------
  1. setup         : create sandbox tree and a shared subset query_index.json
  2. estimate_cost : compute upper-bound cost and prompt for confirmation
  3. submit        : submit each repetition to its own sandbox subdir
  4. wait          : poll batch status until all repetitions are complete
  5. download      : download responses for every repetition
  6. parse         : parse responses into the list-of-records format
  7. consistency   : run consistency_check.py across the 3 parsed outputs
  8. analyse       : run analysis.py all on repetition 1 (metrics/stats/plots)

Any stage can be rerun individually via --only, and --skip lets you turn
off stages that have already completed.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VENV_PY = ROOT / ".venv" / "bin" / "python"

# Per-query cost upper-bound in USD (using the rates baked into config.py
# and a generous token budget — these are sanity caps, not accurate quotes).
# batch APIs are assumed (50% off standard).
PER_QUERY_USD = {
    "gpt-5.4":         0.0040,
    "gemini-3.1-pro":  0.0040,
    "claude-sonnet-4.6": 0.0060,
}


class PilotError(RuntimeError):
    pass


def log(msg: str, *, level: str = "info") -> None:
    prefix = {"info": "ℹ", "ok": "✓", "warn": "⚠", "err": "✗", "step": "»"}[level]
    print(f"{prefix} {msg}", flush=True)


def run_pipeline(cmd: list[str], sandbox: Path, *, check: bool = True,
                 capture: bool = False) -> subprocess.CompletedProcess:
    """
    Run a subprocess command with the DENTAL_MLLM_RESULTS_DIR env var
    pointing at `sandbox`. Real data is never touched.
    """
    env = os.environ.copy()
    env["DENTAL_MLLM_RESULTS_DIR"] = str(sandbox)
    r = subprocess.run(
        cmd, cwd=str(ROOT), env=env,
        capture_output=capture, text=True,
    )
    # Surface non-empty stderr from captured subprocesses so transient failures
    # (e.g. per-chunk OpenAI errors during cmd_status) are visible in the
    # orchestrator log instead of being silently discarded.
    if capture and r.stderr and r.stderr.strip():
        log(f"subprocess stderr from {cmd[-1] if cmd else '?'}:", level="warn")
        for line in r.stderr.rstrip().splitlines():
            print(f"    {line}")
    if check and r.returncode != 0:
        if capture:
            print(r.stdout)
            print(r.stderr, file=sys.stderr)
        raise PilotError(f"command failed: {' '.join(cmd)} (exit {r.returncode})")
    return r


def load_dotenv(path: Path) -> int:
    """
    Minimal .env loader (no dependency on python-dotenv).
    Parses KEY=VALUE lines, ignores blanks and comments, strips matched
    surrounding quotes. Existing os.environ values are NOT overwritten,
    so a value already exported in the shell wins.
    Returns the number of variables loaded.
    """
    if not path.exists():
        return 0
    n = 0
    for raw in path.read_text().splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export "):]
        if "=" not in line:
            continue
        key, _, val = line.partition("=")
        key = key.strip()
        val = val.strip()
        if len(val) >= 2 and val[0] == val[-1] and val[0] in ("'", '"'):
            val = val[1:-1]
        if key and key not in os.environ:
            os.environ[key] = val
            n += 1
    return n


def preflight(args: argparse.Namespace) -> None:
    """Fail early if anything is misconfigured."""
    log("Pre-flight checks", level="step")

    # Safety gate: .api_lock prevents accidental API spend
    api_lock = ROOT / ".api_lock"
    if api_lock.exists() and not args.dry_run:
        raise PilotError(
            "API calls are locked. Delete .api_lock to unlock, or use --dry-run."
        )

    assert VENV_PY.exists(), f"missing venv at {VENV_PY}; run `python3 -m venv .venv` first"

    loaded = load_dotenv(ROOT / ".env")
    if loaded:
        log(f"Loaded {loaded} variable(s) from .env", level="ok")

    # Parse models + validate
    import sys as _sys
    _sys.path.insert(0, str(ROOT))
    import config  # type: ignore

    models = [m.strip() for m in args.models.split(",") if m.strip()]
    unknown = [m for m in models if m not in config.MODELS]
    if unknown:
        raise PilotError(f"unknown models in --models: {unknown}; "
                         f"valid: {sorted(config.MODELS.keys())}")

    # API keys (skip on --dry-run)
    if not args.dry_run:
        required_providers = {config.MODELS[m]["provider"] for m in models}
        env_map = {"openai": "OPENAI_API_KEY", "google": "GOOGLE_API_KEY",
                   "anthropic": "ANTHROPIC_API_KEY"}
        missing = [env_map[p] for p in required_providers if not os.environ.get(env_map[p])]
        if missing:
            raise PilotError(f"missing API key env var(s): {missing}. "
                             f"Set them before running.")
        log(f"API keys present for providers: {sorted(required_providers)}", level="ok")
    else:
        log("--dry-run: skipping API key check", level="warn")

    # Sandbox sanity: either nonexistent or a directory owned by us
    sandbox = Path(args.sandbox).resolve()
    if sandbox.exists() and not sandbox.is_dir():
        raise PilotError(f"sandbox path {sandbox} exists and is not a directory")
    if sandbox == (ROOT / "results").resolve():
        raise PilotError("REFUSING to use the real results/ directory as the sandbox. "
                         "Pick a different --sandbox (e.g., results_pilot/).")

    # Venv check
    log(f"Using venv python: {VENV_PY}", level="ok")
    log(f"Sandbox: {sandbox}", level="ok")
    log(f"Models: {models}", level="ok")
    log(f"Repetitions: {args.repetitions}", level="ok")
    log(f"Images per modality: {args.n_per_modality}", level="ok")

    args._models = models  # stash resolved list on the Namespace


def stage_setup(args: argparse.Namespace) -> Path:
    """Create the sandbox tree and a SHARED master query_index.json."""
    log(f"Stage: setup", level="step")
    sandbox = Path(args.sandbox).resolve()
    sandbox.mkdir(parents=True, exist_ok=True)

    master_index = sandbox / "query_index.json"
    # Sandboxes that contain a v1_reference.json are paired-comparison sandboxes
    # whose query_index is cryptographically anchored to a sibling pilot run.
    # Regenerating it would invalidate the pairing silently. Refuse outright.
    v1_ref = sandbox / "v1_reference.json"
    if v1_ref.exists() and args.force_prepare:
        raise PilotError(
            f"REFUSING --force-prepare on {sandbox}: this sandbox is a "
            f"paired-comparison run (v1_reference.json present) and its "
            f"query_index.json is hash-anchored to a sibling pilot. "
            f"Regenerating it would silently break the pairing. Remove "
            f"v1_reference.json first if you really mean to."
        )
    if master_index.exists() and not args.force_prepare:
        log(f"Master query_index already exists at {master_index} — reusing "
            f"(use --force-prepare to regenerate)", level="ok")
    else:
        run_pipeline(
            [str(VENV_PY), "pipeline.py", "prepare",
             "--n-per-modality", str(args.n_per_modality)],
            sandbox=sandbox, capture=True,
        )
        log(f"Master query_index written to {master_index}", level="ok")

    # Sanity: verify the subset count matches expectations
    with open(master_index) as f:
        queries = json.load(f)
    n_images = len({q["image_id"] for q in queries})
    log(f"  {len(queries)} queries from {n_images} images", level="ok")
    return sandbox


def stage_estimate_cost(args: argparse.Namespace, sandbox: Path) -> float:
    """Compute an upper-bound cost and prompt for explicit confirmation."""
    log("Stage: cost estimate", level="step")
    with open(sandbox / "query_index.json") as f:
        queries = json.load(f)
    n_queries = len(queries)
    n_strategies = 2  # zero_shot + guided
    total_calls_per_rep = n_queries * n_strategies * len(args._models)
    total_calls = total_calls_per_rep * args.repetitions

    per_call = max(PER_QUERY_USD[m] for m in args._models)
    total_cost = total_calls * per_call

    print()
    print("  Cost estimate (upper bound)")
    print("  ----------------------------")
    print(f"  queries              : {n_queries}")
    print(f"  strategies           : {n_strategies} (zero_shot + guided)")
    print(f"  models               : {args._models}")
    print(f"  repetitions          : {args.repetitions}")
    print(f"  total API calls      : {total_calls}")
    print(f"  per-call ceiling     : ${per_call:.4f}")
    print(f"  estimated max cost   : ${total_cost:.2f}")
    print()

    if args.dry_run:
        log("--dry-run: skipping confirmation prompt", level="warn")
        return total_cost

    if args.yes:
        log("--yes: skipping confirmation prompt", level="warn")
        return total_cost

    answer = input("  Proceed? Type 'yes' exactly to continue: ").strip()
    if answer != "yes":
        raise PilotError("user did not confirm — aborting before any API call")
    return total_cost


def stage_submit(args: argparse.Namespace, sandbox: Path) -> list[Path]:
    """
    For each repetition, create a run dir, copy the master query_index into
    it, and submit. Idempotent: skips runs whose batch_tracking.json already
    records a non-failed status for every (model, strategy).
    """
    log("Stage: submit", level="step")
    run_dirs: list[Path] = []
    master_index = sandbox / "query_index.json"

    for i in range(1, args.repetitions + 1):
        run_dir = sandbox / f"run{i}"
        run_dir.mkdir(parents=True, exist_ok=True)
        run_index = run_dir / "query_index.json"
        if not run_index.exists():
            shutil.copyfile(master_index, run_index)

        if args.dry_run:
            log(f"  run{i}: --dry-run → synthesising fake responses", level="warn")
            _synthesise_dry_run_responses(run_dir, args._models)
            run_dirs.append(run_dir)
            continue

        # Safety: refuse to launch a real run on top of leftover response files
        # (e.g., from a previous --dry-run). The download stage skips files that
        # already exist, which would silently feed FAKE data into parse/analysis.
        # Only allow leftovers if a real batch_tracking.json exists for them
        # (idempotent resume of a real run).
        responses_dir = run_dir / "responses"
        leftover = sorted(responses_dir.glob("*")) if responses_dir.exists() else []
        tracking_path = run_dir / "batch_tracking.json"
        if leftover and not tracking_path.exists():
            raise PilotError(
                f"REFUSING to submit run{i}: {len(leftover)} response file(s) "
                f"already exist in {responses_dir} but there is no "
                f"batch_tracking.json. These are likely dry-run leftovers and "
                f"would be silently consumed by the download stage. "
                f"Delete {run_dir} (or the entire sandbox) and re-run."
            )

        tracking_path = run_dir / "batch_tracking.json"
        if tracking_path.exists():
            with open(tracking_path) as f:
                tracking = json.load(f)
            expected_batches = {f"{m}_{s}"
                                for m in args._models
                                for s in ("zero_shot", "guided")}
            completed = {k for k, v in tracking.items()
                         if v.get("status") not in (None, "failed")}
            if expected_batches.issubset(completed):
                log(f"  run{i}: already fully submitted — skipping", level="ok")
                run_dirs.append(run_dir)
                continue

        log(f"  run{i}: submitting {len(args._models)} model(s) × 2 strategies", level="step")
        run_pipeline(
            [str(VENV_PY), "pipeline.py", "submit",
             "--models", ",".join(args._models)],
            sandbox=run_dir,
        )
        run_dirs.append(run_dir)

    return run_dirs


def _synthesise_dry_run_responses(run_dir: Path, models: list[str]) -> None:
    """
    In --dry-run mode, write fabricated OpenAI/Google-format response files
    straight into the sandbox responses/ directory so the rest of the
    orchestrator (status/download/parse/analyse) can be exercised end-to-end
    without any real API call.
    """
    import random
    import sys as _sys
    _sys.path.insert(0, str(ROOT))
    import config  # type: ignore
    responses_dir = run_dir / "responses"
    responses_dir.mkdir(parents=True, exist_ok=True)
    with open(run_dir / "query_index.json") as f:
        queries = json.load(f)

    rng = random.Random(hash(run_dir.name))
    for model_key in models:
        provider = config.MODELS[model_key]["provider"]
        for strategy in ("zero_shot", "guided"):
            batch_name = f"{model_key}_{strategy}"
            if provider == "openai":
                out = responses_dir / f"{batch_name}_chunk000_results.jsonl"
                with open(out, "w") as f:
                    for q in queries:
                        gt = (q.get("omfr_1") or "B3").strip().split(",")[0].strip()
                        # Mostly correct, occasionally off-by-one
                        text = gt if rng.random() < 0.8 else "I cannot identify"
                        f.write(json.dumps({
                            "custom_id": f"{q['query_id']}_{strategy}",
                            "response": {"body": {"choices": [{"message": {"content": text}}]}},
                        }) + "\n")
            elif provider == "google":
                out = responses_dir / f"{batch_name}_chunk000.json"
                entries = []
                for q in queries:
                    gt = (q.get("omfr_1") or "B3").strip().split(",")[0].strip()
                    text = gt if rng.random() < 0.8 else "I cannot identify"
                    entries.append({
                        "custom_id": f"{q['query_id']}_{strategy}",
                        "response": {"candidates": [{
                            "content": {"parts": [{"text": text}]},
                        }]},
                    })
                with open(out, "w") as f:
                    json.dump(entries, f)
            else:
                raise PilotError(f"dry-run unsupported provider: {provider}")


def stage_wait(args: argparse.Namespace, run_dirs: list[Path]) -> None:
    """Poll `pipeline.py status` on each run until all batches complete."""
    if args.dry_run:
        log("--dry-run: skipping wait", level="warn")
        return
    log("Stage: wait for completion", level="step")

    deadline = time.time() + args.max_wait_seconds
    poll_interval = args.poll_interval
    while True:
        all_done = True
        for run_dir in run_dirs:
            tracking_path = run_dir / "batch_tracking.json"
            if not tracking_path.exists():
                all_done = False
                break
            # Let pipeline.py status refresh its own tracking file
            run_pipeline([str(VENV_PY), "pipeline.py", "status"],
                         sandbox=run_dir, capture=True)
            with open(tracking_path) as f:
                tracking = json.load(f)
            for batch_name, info in tracking.items():
                status = info.get("status")
                if status in ("completed", "ended"):
                    continue
                if info.get("provider") == "google" and status == "completed":
                    continue
                # Anything else is still pending
                all_done = False
                log(f"  {run_dir.name}/{batch_name}: {status}", level="info")
        if all_done:
            log("All batches complete", level="ok")
            return
        if time.time() > deadline:
            raise PilotError(
                f"Batches did not complete within {args.max_wait_seconds}s. "
                f"You can re-run the orchestrator to resume from this point."
            )
        log(f"  polling again in {poll_interval}s...", level="info")
        time.sleep(poll_interval)


def stage_download(args: argparse.Namespace, run_dirs: list[Path]) -> None:
    if args.dry_run:
        log("--dry-run: skipping download (responses already synthesised)", level="warn")
        return
    log("Stage: download", level="step")
    for run_dir in run_dirs:
        run_pipeline([str(VENV_PY), "pipeline.py", "download"],
                     sandbox=run_dir)


def stage_parse(args: argparse.Namespace, run_dirs: list[Path]) -> None:
    log("Stage: parse", level="step")
    for run_dir in run_dirs:
        run_pipeline([str(VENV_PY), "pipeline.py", "parse"],
                     sandbox=run_dir)


def stage_consistency(args: argparse.Namespace, run_dirs: list[Path]) -> None:
    log("Stage: consistency check (Fleiss' kappa)", level="step")
    if len(run_dirs) < 2:
        log("need at least 2 runs for consistency — skipping", level="warn")
        return
    paths = [str(run_dir / "parsed_responses.json") for run_dir in run_dirs]
    out = run_dirs[0].parent / "consistency_check.json"
    subprocess.run(
        [str(VENV_PY), "consistency_check.py", *paths, "--output", str(out)],
        cwd=str(ROOT), check=True,
    )
    log(f"Consistency results written to {out}", level="ok")


def stage_analyse(args: argparse.Namespace, run_dirs: list[Path]) -> None:
    log("Stage: analyse run1", level="step")
    run_pipeline([str(VENV_PY), "analysis.py", "metrics"], sandbox=run_dirs[0])
    run_pipeline([str(VENV_PY), "analysis.py", "stats"], sandbox=run_dirs[0])
    run_pipeline([str(VENV_PY), "analysis.py", "report"], sandbox=run_dirs[0])
    run_pipeline([str(VENV_PY), "analysis.py", "visualize"], sandbox=run_dirs[0])
    log(f"Analysis written to {run_dirs[0] / 'analysis'}", level="ok")


def stage_summary(args: argparse.Namespace, sandbox: Path, run_dirs: list[Path]) -> None:
    log("Stage: summary", level="step")
    print()
    print("=" * 70)
    print("PILOT SUMMARY")
    print("=" * 70)

    # Compliance per run
    for run_dir in run_dirs:
        cs_path = run_dir / "compliance_stats.json"
        if cs_path.exists():
            with open(cs_path) as f:
                cs = json.load(f)
            print(f"\n{run_dir.name} compliance: "
                  f"{cs['compliant']}/{cs['total']} ({cs['compliance_rate']:.1%})")
            print(f"  failure modes: {cs['failure_modes']}")

    # Consistency
    cc_path = sandbox / "consistency_check.json"
    if cc_path.exists():
        with open(cc_path) as f:
            cc = json.load(f)
        print(f"\nFleiss' kappa across {cc['n_runs']} runs: {cc['fleiss_kappa']:.4f}")
        print(f"Exact match rate:  {cc['exact_match_rate']:.4f} "
              f"({cc['exact_match_count']}/{cc['n_items']})")
        if abs(cc["fleiss_kappa"] - 1.0) < 1e-6:
            print("  ✓ Perfect determinism — single-run evaluation justified.")
        else:
            print("  ⚠ Non-deterministic — investigate before single-run eval.")

    # Summary table path
    summary_csv = run_dirs[0] / "analysis" / "summary_table.csv"
    if summary_csv.exists():
        print(f"\nAnalysis outputs in {run_dirs[0] / 'analysis'}:")
        print(f"  summary:  {summary_csv}")
        print(f"  plots:    {run_dirs[0] / 'analysis' / 'plots'}")

    print()
    print("Next step: inspect the summary table and plots, then decide whether")
    print("to commit to the full 900-query experiment.")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Dental MLLM pilot orchestrator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--sandbox", default="results_pilot",
                        help="Directory to hold all pilot artefacts (default: results_pilot)")
    parser.add_argument("--n-per-modality", type=int, default=5,
                        help="Images per modality (default: 5)")
    parser.add_argument("--repetitions", type=int, default=3,
                        help="Number of repeated runs for Fleiss' kappa (default: 3)")
    parser.add_argument("--models", default="gpt-5.4",
                        help="Comma-separated model_keys (default: gpt-5.4)")
    parser.add_argument("--force-prepare", action="store_true",
                        help="Regenerate master query_index.json even if it exists")
    parser.add_argument("--yes", action="store_true",
                        help="Skip the cost-confirmation prompt (dangerous)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Never call real APIs; fabricate responses locally")
    parser.add_argument("--max-wait-seconds", type=int, default=24 * 3600,
                        help="Max time to poll batch status before giving up")
    parser.add_argument("--poll-interval", type=int, default=60,
                        help="Seconds between status polls")
    args = parser.parse_args()

    try:
        preflight(args)
        sandbox = stage_setup(args)
        stage_estimate_cost(args, sandbox)
        run_dirs = stage_submit(args, sandbox)
        stage_wait(args, run_dirs)
        stage_download(args, run_dirs)
        stage_parse(args, run_dirs)
        stage_consistency(args, run_dirs)
        stage_analyse(args, run_dirs)
        stage_summary(args, sandbox, run_dirs)
    except PilotError as e:
        log(str(e), level="err")
        sys.exit(2)
    except KeyboardInterrupt:
        log("interrupted — rerun the script to resume", level="warn")
        sys.exit(130)


if __name__ == "__main__":
    main()
