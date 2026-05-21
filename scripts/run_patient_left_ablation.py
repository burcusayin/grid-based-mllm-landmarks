"""
Orchestrate the patient-left ablation experiment (Stage 1).

Submits the `guided_patient_left` strategy on the 100 panoramic Tooth_33_Apex
queries, three repetitions, against gpt-5.4. Uses the existing pipeline
infrastructure (atomic writes, per-chunk persistence, .api_lock,
terminal-state handling) — this orchestrator is just a tightly-scoped
sequencer that:

  1. Refuses to launch unless preflight passes
  2. Refuses if .api_lock is present
  3. Issues a single live API call (NOT batch) as a final auth/cost check
  4. Submits the batch in 3 reps × 6 chunks each
  5. Polls until each batch reaches a terminal state
  6. Downloads + parses each rep's responses into the sandbox
  7. Auto-relocks .api_lock in finally

No work is done outside results_ablation_patient_left/. The frozen results_full/
and results_consensus/ are anchored read-only and never modified.

USAGE:
    .venv/bin/python scripts/run_patient_left_ablation.py
        [--sandbox results_ablation_patient_left]
        [--repetitions 3]
        [--no-live-test]   (skip the single-call live test; not recommended)
        [--dry-run]        (synthesise fake responses for pipeline rehearsal)

The orchestrator mirrors run_pilot.py's structure but is scoped to this one
ablation.
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


class AblationError(Exception):
    pass


def log(msg: str, *, level: str = "info") -> None:
    prefix = {"info": "ℹ", "ok": "✓", "warn": "⚠", "err": "✗", "step": "»"}[level]
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {prefix} {msg}", flush=True)


def run_pipeline(cmd, sandbox, *, check=True, capture=False):
    env = os.environ.copy()
    env["DENTAL_MLLM_RESULTS_DIR"] = str(sandbox)
    r = subprocess.run(cmd, cwd=str(ROOT), env=env, capture_output=capture, text=True)
    if capture and r.stderr and r.stderr.strip():
        log(f"subprocess stderr:", level="warn")
        for line in r.stderr.rstrip().splitlines():
            print(f"    {line}")
    if check and r.returncode != 0:
        if capture:
            print(r.stdout); print(r.stderr, file=sys.stderr)
        raise AblationError(f"command failed: {' '.join(cmd)} (exit {r.returncode})")
    return r


def load_dotenv(path: Path) -> None:
    if not path.exists():
        return
    for raw in path.read_text().splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        if line.startswith("export "):
            line = line[len("export "):]
        k, _, v = line.partition("=")
        v = v.strip().strip('"').strip("'")
        os.environ.setdefault(k.strip(), v)


# ── Stage 0: preflight ─────────────────────────────────────────────

def preflight(args) -> None:
    log("Running preflight gate", level="step")
    r = subprocess.run(
        [str(VENV_PY), "scripts/verify_patient_left_ablation_setup.py",
         "--sandbox", args.sandbox],
        cwd=str(ROOT),
    )
    if r.returncode != 0:
        raise AblationError(
            "Preflight failed — refusing to launch. "
            "Fix the reported issues, then re-run."
        )
    # .api_lock check
    if (ROOT / ".api_lock").exists():
        raise AblationError(
            ".api_lock present. Remove it manually with `rm .api_lock` "
            "AFTER you have reviewed everything; this orchestrator refuses "
            "to remove it for you."
        )


# ── Stage 1: live single-call test ────────────────────────────────

def stage_live_test(args, sandbox: Path) -> None:
    """Single non-batch API call to verify auth + model + prompt + parsing
    before committing to a 300-call batch. Cost: ~$0.001 (one call)."""
    if args.no_live_test:
        log("Skipping live test (--no-live-test)", level="warn")
        return
    if args.dry_run:
        log("Dry-run: skipping live test", level="warn")
        return

    log("Stage: live single-call API test", level="step")
    sys.path.insert(0, str(ROOT))
    import pipeline  # type: ignore
    import config    # type: ignore

    queries = json.loads((sandbox / "query_index.json").read_text())
    q = queries[0]  # PAN_001_Tooth_33_Apex
    sys_p, usr_p = pipeline.generate_prompt(q, "guided_patient_left")
    log(f"  query: {q['query_id']}, GT: {q['consensus_gt']}", level="info")
    log(f"  prompt length: system={len(sys_p)} chars, user={len(usr_p)} chars",
        level="info")

    api_key = os.environ.get("OPENAI_API_KEY", "")
    if not api_key.startswith("sk-"):
        raise AblationError("OPENAI_API_KEY missing or malformed for live test")

    # Use OpenAI's Responses API directly (not batch) — cheapest single call
    import urllib.request
    import urllib.error
    import base64

    img_b64 = base64.b64encode(Path(q["image_path"]).read_bytes()).decode()
    body = {
        "model": "gpt-5.4",
        "temperature": 0,
        "seed": 42,
        "max_completion_tokens": 50,
        "messages": [
            {"role": "system", "content": sys_p},
            {"role": "user", "content": [
                {"type": "image_url",
                 "image_url": {"url": f"data:image/png;base64,{img_b64}",
                               "detail": "high"}},
                {"type": "text", "text": usr_p},
            ]},
        ],
    }
    req = urllib.request.Request(
        "https://api.openai.com/v1/chat/completions",
        method="POST",
        data=json.dumps(body).encode(),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            payload = json.loads(resp.read())
    except urllib.error.HTTPError as e:
        msg = e.read().decode(errors="replace")[:400]
        raise AblationError(f"Live test FAILED ({e.code}): {msg}")
    except Exception as e:
        raise AblationError(f"Live test FAILED: {e}")

    content = (payload.get("choices", [{}])[0]
                      .get("message", {})
                      .get("content", "")).strip()
    usage = payload.get("usage", {})
    parsed = pipeline.parse_grid_coordinate(content)
    log(f"  response: {content!r}", level="info")
    log(f"  parsed: {parsed}", level="info")
    log(f"  tokens: prompt={usage.get('prompt_tokens')}, "
        f"completion={usage.get('completion_tokens')}", level="info")
    if not parsed:
        raise AblationError(
            f"Live test parse failure — model returned {content!r} which "
            f"could not be parsed as a grid coordinate. Refusing to launch batch."
        )
    log("  ✓ live test passed: model authenticated, response parses correctly",
        level="ok")
    # Persist as a sanity record
    (sandbox / "live_test_record.json").write_text(json.dumps({
        "query_id": q["query_id"],
        "raw_response": content,
        "parsed_coordinates": parsed,
        "usage": usage,
    }, indent=2))


# ── Stage 2: submit batches ────────────────────────────────────────

def stage_submit(args, sandbox: Path) -> list[Path]:
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
            log(f"  run{i}: --dry-run → fake responses", level="warn")
            _synthesise_fake(run_dir)
            run_dirs.append(run_dir)
            continue

        # Refuse if real run on top of leftover responses with no tracking
        responses_dir = run_dir / "responses"
        leftover = sorted(responses_dir.glob("*")) if responses_dir.exists() else []
        tracking_path = run_dir / "batch_tracking.json"
        if leftover and not tracking_path.exists():
            raise AblationError(
                f"REFUSING run{i}: {len(leftover)} stale response file(s) "
                f"in {responses_dir} but no batch_tracking.json. Delete and retry."
            )

        if tracking_path.exists():
            tracking = json.loads(tracking_path.read_text())
            expected = {"gpt-5.4_guided_patient_left"}
            completed = {k for k, v in tracking.items()
                         if v.get("status") not in (None, "failed")}
            if expected.issubset(completed):
                log(f"  run{i}: already submitted — skipping", level="ok")
                run_dirs.append(run_dir)
                continue

        log(f"  run{i}: submitting gpt-5.4 × guided_patient_left (100 queries)",
            level="step")
        run_pipeline(
            [str(VENV_PY), "pipeline.py", "submit",
             "--models", "gpt-5.4",
             "--strategies", "guided_patient_left"],
            sandbox=run_dir,
        )
        run_dirs.append(run_dir)
    return run_dirs


def _synthesise_fake(run_dir: Path) -> None:
    """Dry-run: emit fabricated batch outputs so downstream stages can run."""
    import random
    sys.path.insert(0, str(ROOT))
    import config  # noqa: F401  (kept for parity)
    rng = random.Random(42)
    responses_dir = run_dir / "responses"
    responses_dir.mkdir(parents=True, exist_ok=True)
    queries = json.loads((run_dir / "query_index.json").read_text())
    out = responses_dir / "gpt-5.4_guided_patient_left_chunk000_results.jsonl"
    with open(out, "w") as f:
        for q in queries:
            cell = chr(ord("A") + rng.randint(0, 7)) + str(rng.randint(1, 16))
            obj = {
                "id": f"batch_req_{q['query_id']}",
                "custom_id": f"{q['query_id']}_guided_patient_left",
                "response": {"status_code": 200, "request_id": "dry_run",
                             "body": {"choices": [{"message": {"role": "assistant",
                                                                "content": cell}}],
                                      "usage": {"prompt_tokens": 3000,
                                                "completion_tokens": 5}}},
                "error": None,
            }
            f.write(json.dumps(obj) + "\n")
    # Minimal tracking
    (run_dir / "batch_tracking.json").write_text(json.dumps({
        "gpt-5.4_guided_patient_left": {
            "provider": "openai",
            "batch_ids": ["dry_run"],
            "status": "completed",
            "model": "gpt-5.4",
            "strategy": "guided_patient_left",
            "output_file_ids": ["dry_run"],
        }
    }, indent=2))


# ── Stage 3: wait for terminal state ───────────────────────────────

def stage_wait(args, run_dirs: list[Path]) -> None:
    if args.dry_run:
        return
    log("Stage: wait for batch completion", level="step")
    poll_seconds = args.poll_seconds
    for rd in run_dirs:
        log(f"  polling {rd.name}", level="step")
        while True:
            r = run_pipeline([str(VENV_PY), "pipeline.py", "status"],
                             sandbox=rd, capture=True)
            output = r.stdout
            tracking = json.loads((rd / "batch_tracking.json").read_text())
            ok = True
            for name, info in tracking.items():
                st = info.get("status")
                if st in ("completed", "completed_with_failures", "failed"):
                    continue
                ok = False
            if ok:
                log(f"    {rd.name} reached terminal state", level="ok")
                break
            log(f"    not yet terminal; sleeping {poll_seconds}s", level="info")
            time.sleep(poll_seconds)


# ── Stage 4: download + parse ──────────────────────────────────────

def stage_download(args, run_dirs: list[Path]) -> None:
    if args.dry_run:
        return
    log("Stage: download", level="step")
    for rd in run_dirs:
        run_pipeline([str(VENV_PY), "pipeline.py", "download"], sandbox=rd)


def stage_parse(args, run_dirs: list[Path]) -> None:
    log("Stage: parse", level="step")
    for rd in run_dirs:
        run_pipeline([str(VENV_PY), "pipeline.py", "parse"], sandbox=rd)


# ── Stage 5: top-line summary ──────────────────────────────────────

def stage_summary(args, sandbox: Path, run_dirs: list[Path]) -> None:
    log("Stage: summary", level="step")
    total_calls = 0; total_failures = 0
    for rd in run_dirs:
        cs = rd / "compliance_stats.json"
        if not cs.exists():
            log(f"  {rd.name}: missing compliance_stats.json", level="warn")
            continue
        stats = json.loads(cs.read_text())
        # Compliance stats schema mirrors v2: presumably has total + parse_failures
        n = stats.get("total", stats.get("n_total", 0))
        fail = stats.get("parse_failures",
                         stats.get("n_unparseable", 0))
        log(f"  {rd.name}: {n} responses, {fail} parse failures", level="info")
        total_calls += n
        total_failures += fail
    log(f"  TOTAL: {total_calls} responses, {total_failures} parse failures",
        level="ok")
    if total_calls != 300:
        log(f"  WARN: expected 300 calls (100 × 1 strategy × 3 reps), got {total_calls}",
            level="warn")


# ── Main ───────────────────────────────────────────────────────────

def relock(reason: str) -> None:
    lock = ROOT / ".api_lock"
    if not lock.exists():
        lock.write_text(f"Re-locked after patient-left ablation: {reason}\n"
                        f"Time: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        log(f"Re-locked .api_lock ({reason})", level="ok")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sandbox", default="results_ablation_patient_left")
    ap.add_argument("--repetitions", type=int, default=3)
    ap.add_argument("--poll-seconds", type=int, default=120)
    ap.add_argument("--no-live-test", action="store_true",
                    help="Skip the live single-call test (NOT recommended)")
    ap.add_argument("--dry-run", action="store_true",
                    help="Synthesise fake responses (no API calls)")
    args = ap.parse_args()

    load_dotenv(ROOT / ".env")
    sandbox = (ROOT / args.sandbox).resolve()

    # Refuse if sandbox is a frozen/canonical directory
    for forbidden in ("results_full", "results_consensus"):
        p = (ROOT / forbidden).resolve()
        if sandbox == p or p in sandbox.parents:
            raise AblationError(
                f"REFUSING to use {sandbox} — it is or lives inside {p}.")

    if not (sandbox / "query_index.json").exists():
        raise AblationError(
            f"{sandbox}/query_index.json missing. Run prepare_ablation first:\n"
            f"  DENTAL_MLLM_RESULTS_DIR={args.sandbox} "
            f".venv/bin/python pipeline.py prepare_ablation "
            f"--structures Tooth_33_Apex --expected-count 100 "
            f"--anchor-to results_full")

    # Confirm prompt — explicit yes required
    print("\nThis will submit 300 OpenAI Batch API calls (gpt-5.4 × "
          "guided_patient_left × 100 queries × 3 reps).")
    print("Estimated cost: ~$1–2.")
    print("\nType 'yes' to continue: ", end="", flush=True)
    confirm = sys.stdin.readline().strip().lower()
    if confirm != "yes":
        log("Aborted by user", level="warn")
        return

    try:
        preflight(args)
        stage_live_test(args, sandbox)
        run_dirs = stage_submit(args, sandbox)
        stage_wait(args, run_dirs)
        stage_download(args, run_dirs)
        stage_parse(args, run_dirs)
        stage_summary(args, sandbox, run_dirs)
        log("patient-left ablation Stage 1 complete", level="ok")
    finally:
        relock("orchestrator finally")


if __name__ == "__main__":
    main()
