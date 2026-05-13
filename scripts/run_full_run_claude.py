"""
Orchestrate the CLAUDE SONNET 4.6 full 900-query run.

Submits all 900 queries × 2 strategies (zero_shot, guided) × 3 repetitions
= 5,400 API calls to Anthropic's Message Batches API. Uses the same canonical
prompts that the GPT-5.4 v2 main run received (byte-identity asserted by the
preflight). Stores everything under results_full_claude/ in complete isolation
from every other sandbox.

Pipeline (mirrors run_full_run_gemini.py; Anthropic-specific deltas only):

  Stage 0  preflight       run scripts/verify_full_run_claude_setup.py
  Stage 1  prepare_sandbox  create sandbox; copy v2 query_index; manifests
  Stage 2  snapshot_prompts write sandbox/prompts_used.json BEFORE any API
                            call; verify byte-identity vs GPT-5.4 v2
  Stage 3  live_test        one non-batch call to Anthropic Messages API
                            to verify auth + request format + parsing
  Stage 4  submit           three reps × two strategies via
                            pipeline.cmd_submit with --models claude-sonnet-4.6
  Stage 5  wait             ANTHROPIC IS ASYNC — poll batch status until
                            terminal state (typically 5-30 minutes per
                            batch). Calls pipeline.cmd_status repeatedly.
  Stage 6  download         download Anthropic batch result files
  Stage 7  parse            parse responses → parsed_responses.json
  Stage 8  summary          aggregate compliance / call counts across reps

  finally  relock           recreate .api_lock

USAGE:
    .venv/bin/python scripts/run_full_run_claude.py
        [--sandbox results_full_claude]
        [--repetitions 3]
        [--poll-seconds 120]  (how often to poll Anthropic batch status)
        [--no-live-test]      (skip the live single-call test; NOT recommended)
        [--dry-run]           (synthesise fake responses for pipeline rehearsal)

Cost
----
Naive upper-bound projection (5,400 calls × Claude batch pricing
$1.50 / $7.50 per 1M tokens): ~$23.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import random
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VENV_PY = ROOT / ".venv" / "bin" / "python"
SANDBOX_DEFAULT = "results_full_claude"
MODEL_KEY = "claude-sonnet-4.6"
STRATEGIES = ("zero_shot", "guided")
EXPECTED_N_QUERIES = 900


class FullRunError(Exception):
    pass


def log(msg: str, *, level: str = "info") -> None:
    prefix = {"info": "ℹ", "ok": "✓", "warn": "⚠", "err": "✗", "step": "»"}[level]
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {prefix} {msg}", flush=True)


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def run_pipeline(cmd, sandbox: Path, *, check: bool = True, capture: bool = False):
    env = os.environ.copy()
    env["DENTAL_MLLM_RESULTS_DIR"] = str(sandbox)
    r = subprocess.run(cmd, cwd=str(ROOT), env=env, capture_output=capture, text=True)
    if capture and r.stderr and r.stderr.strip():
        log("subprocess stderr:", level="warn")
        for line in r.stderr.rstrip().splitlines():
            print(f"    {line}")
    if check and r.returncode != 0:
        if capture:
            print(r.stdout)
            print(r.stderr, file=sys.stderr)
        raise FullRunError(f"command failed: {' '.join(cmd)} (exit {r.returncode})")
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
    log("Stage 0: preflight gate", level="step")
    if args.dry_run:
        log("  --dry-run: skipping API-key + .api_lock checks "
            "(dry-run never touches the network)", level="warn")
        return
    r = subprocess.run(
        [str(VENV_PY), "scripts/verify_full_run_claude_setup.py",
         "--sandbox", args.sandbox],
        cwd=str(ROOT),
    )
    if r.returncode != 0:
        raise FullRunError(
            "Preflight failed — refusing to launch. "
            "Fix the reported issues, then re-run."
        )
    if (ROOT / ".api_lock").exists():
        raise FullRunError(
            ".api_lock present. Remove it manually with `rm .api_lock` "
            "AFTER you have reviewed everything; this orchestrator refuses "
            "to remove it for you."
        )


# ── Stage 1: prepare sandbox ───────────────────────────────────────

def stage_prepare_sandbox(args, sandbox: Path) -> None:
    """Create sandbox dir, copy v2 query_index, write manifests + anchors.

    Same shape as run_full_run_gemini's prepare stage.
    """
    log("Stage 1: prepare sandbox (copy v2 query_index, write manifests)",
        level="step")
    sandbox.mkdir(parents=True, exist_ok=True)

    v2_qi = ROOT / "results_consensus" / "query_index.json"
    v2_qi_sha = sha256_file(v2_qi)
    sandbox_qi = sandbox / "query_index.json"

    if sandbox_qi.exists():
        existing_sha = sha256_file(sandbox_qi)
        if existing_sha != v2_qi_sha:
            raise FullRunError(
                f"REFUSING: {sandbox_qi} already exists with a DIFFERENT SHA "
                f"({existing_sha[:16]}…) than the canonical v2 query_index "
                f"({v2_qi_sha[:16]}…). Inspect and remove the stale file "
                f"before re-running.")
        log("  query_index.json already in place (matches v2 byte-for-byte)",
            level="ok")
    else:
        shutil.copyfile(v2_qi, sandbox_qi)
        log(f"  copied {v2_qi} → {sandbox_qi}", level="ok")
        assert sha256_file(sandbox_qi) == v2_qi_sha, \
            "post-copy SHA mismatch — disk write may have been truncated"

    final_xlsx = ROOT / "data" / "Final_Dental_MLLM_Benchmark_Data.xlsx"
    excel_sha = sha256_file(final_xlsx)
    v2_manifest = json.loads((ROOT / "results_consensus" / "v2_manifest.json").read_text())

    anchored_files: dict[str, dict] = {}
    for relrun in ("run1", "run2", "run3"):
        for f in sorted((ROOT / "results_full" / relrun / "responses").glob("*.jsonl")):
            rel = f.relative_to(ROOT / "results_full")
            anchored_files[str(rel)] = {
                "sha256": sha256_file(f),
                "size": f.stat().st_size,
            }

    # Anchor every IMAGE file referenced by query_index. If any image is
    # modified before/during the run, this SHA chain detects it; this is
    # essential for "Claude saw the same image bytes as GPT-5.4".
    image_paths = sorted({Path(q["image_path"]) for q in
                          json.loads(sandbox_qi.read_text())})
    log(f"  computing SHA-256 for {len(image_paths)} image files...",
        level="info")
    anchored_images: dict[str, dict] = {}
    for p in image_paths:
        rel = p.relative_to(ROOT / "data")
        anchored_images[str(rel)] = {
            "sha256": sha256_file(p),
            "size": p.stat().st_size,
        }

    anchor = {
        "anchor_root": str((ROOT / "results_full").resolve()),
        "purpose": (
            "Anchor the GPT-5.4 v2 main-run raw JSONLs AND the image files "
            "while the Claude run is in flight. If any of these change, the "
            "Claude-vs-GPT comparison in the final report would be against "
            "a different baseline than the one in scope."
        ),
        "files": anchored_files,
        "source_excel": {
            "path": str(final_xlsx),
            "sha256": excel_sha,
        },
        "v2_query_index": {
            "path": str(v2_qi),
            "sha256": v2_qi_sha,
        },
        "v2_prompts_used": {
            "path": str(ROOT / "results_full" / "prompts_used.json"),
            "sha256": sha256_file(ROOT / "results_full" / "prompts_used.json"),
        },
        "images_root": str((ROOT / "data").resolve()),
        "images": anchored_images,
    }
    (sandbox / "full_run_anchor.json").write_text(json.dumps(anchor, indent=2))
    log(f"  wrote full_run_anchor.json "
        f"({len(anchored_files)} anchored JSONLs + "
        f"{len(anchored_images)} image SHAs)",
        level="ok")

    git_sha = subprocess.check_output(
        ["git", "-C", str(ROOT), "rev-parse", "HEAD"]).decode().strip()
    manifest = {
        "sandbox": str(sandbox),
        "model_key": MODEL_KEY,
        "model_id": "claude-sonnet-4-6-20250514",
        "provider": "anthropic",
        "strategies": list(STRATEGIES),
        "n_queries": EXPECTED_N_QUERIES,
        "n_reps": args.repetitions,
        "expected_total_calls": EXPECTED_N_QUERIES * len(STRATEGIES) * args.repetitions,
        "inference_settings": {
            "temperature": 0,
            "seed": None,  # Anthropic batches don't accept seed
            "max_output_tokens": 256,
            "note": ("Anthropic Messages Batches API does not accept a "
                     "seed parameter. Reproducibility relies solely on "
                     "temperature=0 plus rep-averaging across 3 runs."),
        },
        "source_excel_sha256": excel_sha,
        "source_excel_matches_v2_manifest": (
            v2_manifest.get("source_excel_sha256") == excel_sha),
        "v2_query_index_sha256": v2_qi_sha,
        "git_commit": git_sha,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "comparison_baseline": "results_full/ (GPT-5.4 v2 main run)",
    }
    (sandbox / "full_run_manifest.json").write_text(json.dumps(manifest, indent=2))
    log("  wrote full_run_manifest.json", level="ok")


# ── Stage 2: snapshot prompts ─────────────────────────────────────

def stage_snapshot_prompts(args, sandbox: Path) -> None:
    """Persist EXACT system + user prompts. Verify byte-identity vs GPT-5.4 v2.

    Writes: sandbox/prompts_used.json
    """
    log("Stage 2: snapshot prompts + verify byte-identity vs GPT-5.4 v2",
        level="step")
    sys.path.insert(0, str(ROOT))
    import pipeline  # type: ignore

    queries = json.loads((sandbox / "query_index.json").read_text())
    v2_pu = json.loads((ROOT / "results_full" / "prompts_used.json").read_text())
    v2_by_qid = {p["query_id"]: p for p in v2_pu["prompts"]}

    out = {
        "sandbox": str(sandbox),
        "model_key": MODEL_KEY,
        "n_queries": len(queries),
        "strategies": list(STRATEGIES),
        "snapshot_note": (
            "Exact system+user prompts issued to Claude Sonnet 4.6 for every "
            "(query × strategy) pair. Verified byte-identical to "
            "results_full/prompts_used.json (GPT-5.4 v2 main run) — Claude "
            "received the same prompts as GPT-5.4 so that any performance "
            "difference is attributable to the model itself, not to "
            "prompt drift between runs."
        ),
        "prompts": [],
    }
    n_drift = 0
    for q in queries:
        entry = {
            "query_id": q["query_id"],
            "image_id": q["image_id"],
            "structure": q["structure"],
            "sheet": q["sheet"],
            "landmark_type": q["landmark_type"],
            "consensus_gt": q.get("consensus_gt"),
            "by_strategy": {},
        }
        v2_entry = v2_by_qid.get(q["query_id"])
        for strat in STRATEGIES:
            sys_p, usr_p = pipeline.generate_prompt(q, strat)
            entry["by_strategy"][strat] = {
                "system_prompt": sys_p,
                "user_prompt": usr_p,
                "system_prompt_len": len(sys_p),
                "user_prompt_len": len(usr_p),
            }
            if v2_entry:
                if sys_p != v2_entry["by_strategy"][strat]["system_prompt"]:
                    n_drift += 1
                if usr_p != v2_entry["by_strategy"][strat]["user_prompt"]:
                    n_drift += 1
        out["prompts"].append(entry)

    if n_drift > 0:
        raise FullRunError(
            f"REFUSING: {n_drift} prompt drift(s) vs GPT-5.4 v2 main-run "
            f"detected. Claude would receive different prompts than GPT-5.4, "
            f"breaking the cross-model comparability. Stop and inspect "
            f"pipeline.generate_prompt vs results_full/prompts_used.json.")

    snapshot_path = sandbox / "prompts_used.json"
    import tempfile
    fd, tmp = tempfile.mkstemp(prefix=".prompts_", dir=str(sandbox))
    with os.fdopen(fd, "w") as f:
        json.dump(out, f, indent=2)
    os.replace(tmp, snapshot_path)
    log(f"  prompts_used.json: {len(queries)} queries × {len(STRATEGIES)} strategies "
        f"({snapshot_path.stat().st_size:,} bytes); 0 drifts vs GPT-5.4 v2 ✓",
        level="ok")


# ── Stage 3: live single-call test ─────────────────────────────────

def stage_live_test(args, sandbox: Path) -> None:
    """Single non-batch call to Anthropic's /v1/messages API.

    Cost: ~$0.01 (one call at non-batch standard pricing).
    """
    if args.no_live_test:
        log("Skipping live test (--no-live-test)", level="warn")
        return
    if args.dry_run:
        log("Dry-run: skipping live test", level="warn")
        return

    log("Stage 3: live single-call API test (one query, Claude)", level="step")
    sys.path.insert(0, str(ROOT))
    import pipeline  # type: ignore
    import config    # type: ignore
    import base64
    import urllib.request
    import urllib.error

    queries = json.loads((sandbox / "query_index.json").read_text())
    q = queries[0]
    sys_p, usr_p = pipeline.generate_prompt(q, "guided")
    log(f"  query: {q['query_id']}, GT: {q['consensus_gt']}", level="info")
    log(f"  prompt: system={len(sys_p)} chars, user={len(usr_p)} chars",
        level="info")

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key.startswith("sk-ant-"):
        raise FullRunError(
            "ANTHROPIC_API_KEY missing or malformed for live test "
            "(expected to start with 'sk-ant-')")

    img_b64 = base64.b64encode(Path(q["image_path"]).read_bytes()).decode()
    model_cfg = config.MODELS[MODEL_KEY]
    body = {
        "model": model_cfg["model_id"],
        "max_tokens": model_cfg["max_output_tokens"],
        "temperature": 0,
        "system": sys_p,
        "messages": [{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": img_b64,
                    },
                },
                {"type": "text", "text": usr_p},
            ],
        }],
    }
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps(body).encode(),
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            payload = json.loads(resp.read())
    except urllib.error.HTTPError as e:
        msg = e.read().decode(errors="replace")[:400]
        raise FullRunError(f"Live test FAILED ({e.code}): {msg}")
    except Exception as e:
        raise FullRunError(f"Live test FAILED: {e}")

    content_blocks = payload.get("content", [])
    if not content_blocks or content_blocks[0].get("type") != "text":
        raise FullRunError(
            f"Live test FAILED: Claude returned no text content. Payload: "
            f"{json.dumps(payload)[:400]}")
    content = content_blocks[0].get("text", "").strip()
    usage = payload.get("usage", {})
    parsed = pipeline.parse_grid_coordinate(content)
    log(f"  response: {content!r}", level="info")
    log(f"  parsed: {parsed}", level="info")
    log(f"  usage: input_tokens={usage.get('input_tokens')}, "
        f"output_tokens={usage.get('output_tokens')}", level="info")
    if not parsed:
        raise FullRunError(
            f"Live test parse failure — Claude returned {content!r} which "
            f"could not be parsed as a grid coordinate. Refusing to launch batch.")
    log("  ✓ live test passed: Claude authenticated, response parses correctly",
        level="ok")
    (sandbox / "live_test_record.json").write_text(json.dumps({
        "query_id": q["query_id"],
        "raw_response": content,
        "parsed_coordinates": parsed,
        "usage": usage,
        "model_id": model_cfg["model_id"],
    }, indent=2))


# ── Stage 4: submit ────────────────────────────────────────────────

def stage_submit(args, sandbox: Path) -> list[Path]:
    """Submit each repetition into its own subdirectory. For Anthropic each
    (model × strategy) submit creates ONE batch with all 900 queries —
    Anthropic handles chunking server-side. cmd_submit will record the
    batch_id in batch_tracking.json."""
    log("Stage 4: submit (Anthropic Messages Batches — async, 1 batch per strategy/rep)",
        level="step")
    run_dirs: list[Path] = []
    master_index = sandbox / "query_index.json"

    for i in range(1, args.repetitions + 1):
        run_dir = sandbox / f"run{i}"
        run_dir.mkdir(parents=True, exist_ok=True)
        run_index = run_dir / "query_index.json"
        if not run_index.exists():
            shutil.copyfile(master_index, run_index)
            assert sha256_file(run_index) == sha256_file(master_index)

        if args.dry_run:
            log(f"  run{i}: --dry-run → fake responses", level="warn")
            _synthesise_fake(run_dir)
            run_dirs.append(run_dir)
            continue

        responses_dir = run_dir / "responses"
        leftover = sorted(responses_dir.glob("*")) if responses_dir.exists() else []
        tracking_path = run_dir / "batch_tracking.json"
        if leftover and not tracking_path.exists():
            raise FullRunError(
                f"REFUSING run{i}: {len(leftover)} stale response file(s) "
                f"in {responses_dir} but no batch_tracking.json. "
                f"Delete and retry.")

        if tracking_path.exists():
            tracking = json.loads(tracking_path.read_text())
            expected = {f"{MODEL_KEY}_{s}" for s in STRATEGIES}
            done_states = ("ended", "completed", "completed_with_failures")
            completed = {
                k for k, v in tracking.items() if v.get("status") in done_states
            }
            if expected.issubset(completed):
                log(f"  run{i}: already submitted+ended — skipping", level="ok")
                run_dirs.append(run_dir)
                continue

        log(f"  run{i}: submitting {MODEL_KEY} × {STRATEGIES} "
            f"({EXPECTED_N_QUERIES} queries × {len(STRATEGIES)} strategies "
            f"= {EXPECTED_N_QUERIES * len(STRATEGIES)} calls)",
            level="step")
        run_pipeline(
            [str(VENV_PY), "pipeline.py", "submit",
             "--models", MODEL_KEY,
             "--strategies", ",".join(STRATEGIES)],
            sandbox=run_dir,
        )
        run_dirs.append(run_dir)
    return run_dirs


def _synthesise_fake(run_dir: Path) -> None:
    """Dry-run: emit fabricated Anthropic-format JSONL outputs so downstream
    stages (download, parse, summary) can run without ever calling the API.

    Anthropic batch results JSONL format:
      {"custom_id": "...", "result": {"type": "succeeded",
                                       "message": {"content": [{"type": "text",
                                                                  "text": "B3"}]}}}
    """
    rng = random.Random(42)
    responses_dir = run_dir / "responses"
    responses_dir.mkdir(parents=True, exist_ok=True)
    queries = json.loads((run_dir / "query_index.json").read_text())

    tracking: dict = {}
    for strat in STRATEGIES:
        out_path = responses_dir / f"{MODEL_KEY}_{strat}_results.jsonl"
        with open(out_path, "w") as f:
            for q in queries:
                cell = chr(ord("A") + rng.randint(0, 7)) + str(rng.randint(1, 16))
                obj = {
                    "custom_id": f"{q['query_id']}_{strat}",
                    "result": {
                        "type": "succeeded",
                        "message": {
                            "content": [{"type": "text", "text": cell}],
                            "usage": {"input_tokens": 2500, "output_tokens": 5},
                        },
                    },
                }
                f.write(json.dumps(obj) + "\n")
        tracking[f"{MODEL_KEY}_{strat}"] = {
            "provider": "anthropic",
            "batch_ids": ["dry_run"],
            "status": "ended",
            "model": MODEL_KEY,
            "strategy": strat,
        }
    (run_dir / "batch_tracking.json").write_text(json.dumps(tracking, indent=2))


# ── Stage 5: wait ──────────────────────────────────────────────────

def stage_wait(args, run_dirs: list[Path]) -> None:
    if args.dry_run:
        return
    log("Stage 5: wait — poll Anthropic batch status until terminal state",
        level="step")
    poll_seconds = args.poll_seconds
    for rd in run_dirs:
        log(f"  polling {rd.name}", level="step")
        while True:
            run_pipeline([str(VENV_PY), "pipeline.py", "status"],
                         sandbox=rd, capture=True)
            tracking = json.loads((rd / "batch_tracking.json").read_text())
            all_ended = True
            for name, info in tracking.items():
                st = info.get("status")
                # Anthropic terminal status set by cmd_status is "ended".
                if st not in ("ended", "completed", "completed_with_failures",
                              "failed"):
                    all_ended = False
                    break
            if all_ended:
                log(f"    {rd.name} reached terminal state", level="ok")
                break
            log(f"    not yet terminal; sleeping {poll_seconds}s", level="info")
            time.sleep(poll_seconds)


# ── Stage 6: download ──────────────────────────────────────────────

def stage_download(args, run_dirs: list[Path]) -> None:
    if args.dry_run:
        return
    log("Stage 6: download Anthropic batch result files", level="step")
    for rd in run_dirs:
        run_pipeline([str(VENV_PY), "pipeline.py", "download"], sandbox=rd)


# ── Stage 7: parse ─────────────────────────────────────────────────

def stage_parse(args, run_dirs: list[Path]) -> None:
    log("Stage 7: parse responses → per-run parsed_responses.json",
        level="step")
    for rd in run_dirs:
        run_pipeline([str(VENV_PY), "pipeline.py", "parse"], sandbox=rd)


# ── Stage 8: summary ───────────────────────────────────────────────

def stage_summary(args, sandbox: Path, run_dirs: list[Path]) -> None:
    log("Stage 8: summary", level="step")
    total_calls = 0
    total_failures = 0
    for rd in run_dirs:
        cs = rd / "compliance_stats.json"
        if not cs.exists():
            log(f"  {rd.name}: missing compliance_stats.json", level="warn")
            continue
        stats = json.loads(cs.read_text())
        n = stats.get("total", 0)
        fail = stats.get("non_compliant", 0)
        rate = stats.get("compliance_rate", 0)
        log(f"  {rd.name}: {n} responses, "
            f"{fail} non-compliant ({rate*100:.3f}% strict compliance)",
            level="info")
        total_calls += n
        total_failures += fail
    log(f"  TOTAL: {total_calls} responses, {total_failures} non-compliant",
        level="ok")
    expected = EXPECTED_N_QUERIES * len(STRATEGIES) * args.repetitions
    if total_calls != expected:
        log(f"  WARN: expected {expected} calls "
            f"(900 × {len(STRATEGIES)} strat × {args.repetitions} reps), "
            f"got {total_calls}",
            level="warn")


# ── Finalisation ───────────────────────────────────────────────────

def relock(reason: str) -> None:
    lock = ROOT / ".api_lock"
    if not lock.exists():
        lock.write_text(
            f"Re-locked after Claude Sonnet 4.6 full run: {reason}\n"
            f"Time: {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
        )
        log(f"Re-locked .api_lock ({reason})", level="ok")


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--sandbox", default=SANDBOX_DEFAULT)
    ap.add_argument("--repetitions", type=int, default=3)
    ap.add_argument("--poll-seconds", type=int, default=120,
                    help="Anthropic batch poll interval (default: 120s)")
    ap.add_argument("--no-live-test", action="store_true",
                    help="Skip the live single-call test (NOT recommended)")
    ap.add_argument("--dry-run", action="store_true",
                    help="Synthesise fake responses (no API calls)")
    ap.add_argument("--only-prepare", action="store_true",
                    help="Run ONLY Stage 1 (prepare_sandbox): create the "
                         "folder, copy query_index, write manifests + anchors. "
                         "No API calls, no .api_lock changes, no cost prompt.")
    args = ap.parse_args()

    load_dotenv(ROOT / ".env")
    sandbox = (ROOT / args.sandbox).resolve()

    FORBIDDEN = [
        "results_full", "results_consensus", "data",
        "results_ablation_no_tooth_num", "results_ablation_patient_left",
        "results_ablation_no_LR", "results_pilot", "results_pilot_v2",
        "results_full_gemini",
    ]
    for forbidden in FORBIDDEN:
        p = (ROOT / forbidden).resolve()
        if sandbox == p or p in sandbox.parents:
            raise FullRunError(
                f"REFUSING to use {sandbox} — it is or lives inside {p}. "
                f"Pick a sandbox name that starts with 'results_full_claude'.")

    # --only-prepare: pure file-ops + SHA computation, no API surface.
    # Skip cost confirmation, preflight subprocess, and finally:relock.
    if args.only_prepare:
        log("--only-prepare: running Stage 1 only (no API calls, .api_lock untouched)",
            level="step")
        stage_prepare_sandbox(args, sandbox)
        log("Stage 1 complete. Sandbox prepared at " + str(sandbox), level="ok")
        return

    expected_cost = 24.0
    n_calls = EXPECTED_N_QUERIES * len(STRATEGIES) * args.repetitions
    if not args.dry_run:
        print("\n" + "=" * 70)
        print(f"  CLAUDE SONNET 4.6 FULL RUN — COST CONFIRMATION")
        print("=" * 70)
        print(f"  Sandbox:            {sandbox}")
        print(f"  Queries:            {EXPECTED_N_QUERIES}")
        print(f"  Strategies:         {list(STRATEGIES)}")
        print(f"  Repetitions:        {args.repetitions}")
        print(f"  Total API calls:    {n_calls:,}")
        print(f"  Naive cost est.:    ~${expected_cost:.0f}")
        print(f"  Live single-call:   +~$0.01")
        print()
        print(f"  Type 'yes' to proceed: ", end="", flush=True)
        confirm = sys.stdin.readline().strip().lower()
        if confirm != "yes":
            log("Aborted by user", level="warn")
            return

    try:
        preflight(args)
        stage_prepare_sandbox(args, sandbox)
        stage_snapshot_prompts(args, sandbox)
        stage_live_test(args, sandbox)
        run_dirs = stage_submit(args, sandbox)
        stage_wait(args, run_dirs)
        stage_download(args, run_dirs)
        stage_parse(args, run_dirs)
        stage_summary(args, sandbox, run_dirs)
        log("Claude Sonnet 4.6 full run complete", level="ok")
    finally:
        relock("orchestrator finally")


if __name__ == "__main__":
    main()
