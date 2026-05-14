"""
Orchestrate the GEMINI 3.1 PRO full 900-query run.

Submits all 900 queries × 2 strategies (zero_shot, guided) × 3 repetitions
= 5,400 GPT-5.4-equivalent API calls to Google's Gemini batchGenerateContent
API. Uses the same canonical prompts that the GPT-5.4 v2 main run received
(byte-identity asserted by the preflight). Stores everything under
results_full_gemini/ in complete isolation from every other sandbox.

Pipeline (mirrors the ablation orchestrator pattern, scaled to a full run):

  Stage 0  preflight       run scripts/verify_full_run_gemini_setup.py
                            and refuse to launch unless it passes
  Stage 1  prepare_sandbox  create sandbox dir; copy results_consensus/
                            query_index.json into it; write
                            full_run_manifest.json + full_run_anchor.json
  Stage 2  snapshot_prompts write sandbox/prompts_used.json BEFORE any API
                            call. Lists all 1,800 (query × strategy) prompts
                            with byte-identity verification against
                            results_full/prompts_used.json (GPT-5.4 v2 main
                            run) — so we know Gemini received the same
                            prompts as GPT-5.4.
  Stage 3  live_test        one non-batch call to Gemini Generative
                            Language API to verify auth, request format,
                            and response parsing before committing to the
                            5,400-call batch
  Stage 4  submit           three reps × two strategies × ~2 file-batches
                            each (one per GOOGLE_FILE_CHUNK_SIZE chunk).
                            ASYNC — submit returns once every batch is
                            CREATED with state BATCH_STATE_PENDING. Each rep
                            writes into sandbox/run{N}/ via the
                            DENTAL_MLLM_RESULTS_DIR env var.
  Stage 5  wait             poll pipeline.cmd_status until every batch
                            reaches BATCH_STATE_SUCCEEDED / FAILED /
                            CANCELLED / EXPIRED. Google's batch SLA is
                            "within 24h" with no fast-path; expect 30 min -
                            several hours per rep.
  Stage 6  download         run pipeline.cmd_download per rep — fetches the
                            JSONL result file behind each succeeded batch
                            and converts it to chunk-JSON for the parser.
  Stage 7  parse            run pipeline.cmd_parse per rep, producing
                            sandbox/run{N}/parsed_responses.json
  Stage 8  summary          aggregate compliance / call counts across reps

  finally  relock          recreate .api_lock so accidental future runs cannot
                            spend API money without an explicit unlock

USAGE:
    .venv/bin/python scripts/run_full_run_gemini.py
        [--sandbox results_full_gemini]
        [--repetitions 3]
        [--poll-seconds 120]  (how often to poll Google batch state)
        [--no-live-test]      (skip the live single-call test; NOT recommended)
        [--dry-run]           (synthesise fake responses for pipeline rehearsal)

Cost
----
Naive upper-bound projection (5,400 calls × HIGH image res ~2200 tok/image
× Gemini batch pricing $1.00 / $6.00 per 1M): ~$15.
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
SANDBOX_DEFAULT = "results_full_gemini"
MODEL_KEY = "gemini-3.1-pro"
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
    """Run a pipeline.py sub-command in the given sandbox via DENTAL_MLLM_RESULTS_DIR.

    The env-var redirection is what makes the sandbox isolation work: every
    pipeline.py command writes to RESULTS_DIR derived from this variable.
    """
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
        [str(VENV_PY), "scripts/verify_full_run_gemini_setup.py",
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

    Refuses to overwrite an existing query_index.json — once written, the
    sandbox is considered prepared and the orchestrator can be re-invoked
    safely (it picks up where it left off).
    """
    log("Stage 1: prepare sandbox (copy v2 query_index, write manifests)",
        level="step")
    sys.path.insert(0, str(ROOT))
    import config  # type: ignore
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

    # Source data anchors
    final_xlsx = ROOT / "data" / "Final_Dental_MLLM_Benchmark_Data.xlsx"
    excel_sha = sha256_file(final_xlsx)
    v2_manifest = json.loads((ROOT / "results_consensus" / "v2_manifest.json").read_text())

    # Anchor every results_full/ raw JSONL so that if a colleague ever wonders
    # whether the GPT-5.4 baseline drifted while Gemini was running, we can
    # cryptographically prove it didn't.
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
    # essential for "Gemini saw the same image bytes as GPT-5.4".
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
            "while the Gemini run is in flight. If any of these change, the "
            "Gemini-vs-GPT comparison in the final report would be against "
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
        "model_id": "gemini-3.1-pro-preview",
        "provider": "google",
        "strategies": list(STRATEGIES),
        "n_queries": EXPECTED_N_QUERIES,
        "n_reps": args.repetitions,
        "expected_total_calls": EXPECTED_N_QUERIES * len(STRATEGIES) * args.repetitions,
        "inference_settings": {
            "temperature": 0,
            "seed": 42,
            "max_output_tokens": config.MODELS[MODEL_KEY]["max_output_tokens"],
            "media_resolution": config.MODELS[MODEL_KEY].get(
                "media_resolution", "MEDIA_RESOLUTION_HIGH"),
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
    """Persist the EXACT system + user prompts that will be issued to Gemini,
    for every (query, strategy) pair. Verifies byte-identity vs the GPT-5.4
    v2 main-run prompts_used.json before writing.

    Writes: sandbox/prompts_used.json
    Reads:  sandbox/query_index.json, pipeline.generate_prompt,
            results_full/prompts_used.json (canonical reference)
    No API calls.
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
            "Exact system+user prompts issued to Gemini 3.1 Pro for every "
            "(query × strategy) pair. Verified byte-identical to "
            "results_full/prompts_used.json (GPT-5.4 v2 main run) — Gemini "
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
            f"detected. Gemini would receive different prompts than GPT-5.4, "
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
    """Single non-batch call to Gemini's generativeLanguage API to verify
    auth + model_id + request format + response parsing before committing
    to a 5,400-call batch. Cost: ~$0.005 (one call)."""
    if args.no_live_test:
        log("Skipping live test (--no-live-test)", level="warn")
        return
    if args.dry_run:
        log("Dry-run: skipping live test", level="warn")
        return

    log("Stage 3: live single-call API test (one query, Gemini)", level="step")
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

    api_key = os.environ.get("GOOGLE_API_KEY", "")
    if not api_key.startswith("AIza"):
        raise FullRunError(
            "GOOGLE_API_KEY missing or malformed for live test "
            "(expected to start with 'AIza')")

    img_b64 = base64.b64encode(Path(q["image_path"]).read_bytes()).decode()
    model_cfg = config.MODELS[MODEL_KEY]
    body = {
        "systemInstruction": {"parts": [{"text": sys_p}]},
        "contents": [{
            "role": "user",
            "parts": [
                {"inlineData": {"mimeType": "image/png", "data": img_b64}},
                {"text": usr_p},
            ],
        }],
        "generationConfig": {
            "temperature": 0,
            # Match the actual batch settings so the live test exercises the
            # same configuration the batch will use (not a stale 256 default).
            "maxOutputTokens": model_cfg["max_output_tokens"],
            "seed": 42,
            "mediaResolution": model_cfg.get("media_resolution",
                                              "MEDIA_RESOLUTION_HIGH"),
        },
    }
    model_id = model_cfg["model_id"]
    url = (f"https://generativelanguage.googleapis.com/v1beta/"
           f"models/{model_id}:generateContent?key={api_key}")
    req = urllib.request.Request(
        url, data=json.dumps(body).encode(),
        headers={"Content-Type": "application/json"},
    )
    try:
        # 5 min timeout: Gemini 3.1 Pro thinking can run 30-60s for a single
        # query (especially area landmarks). 60s was too tight once
        # max_output_tokens was bumped to 2048 to fit thinking + answer.
        with urllib.request.urlopen(req, timeout=300) as resp:
            payload = json.loads(resp.read())
    except urllib.error.HTTPError as e:
        msg = e.read().decode(errors="replace")[:400]
        raise FullRunError(f"Live test FAILED ({e.code}): {msg}")
    except Exception as e:
        raise FullRunError(f"Live test FAILED: {e}")

    candidates = payload.get("candidates", [])
    if not candidates:
        raise FullRunError(
            f"Live test FAILED: Gemini returned no candidates. Payload: "
            f"{json.dumps(payload)[:400]}")
    parts = candidates[0].get("content", {}).get("parts", [])
    content = (parts[0].get("text") if parts else "").strip()
    usage = payload.get("usageMetadata", {})
    parsed = pipeline.parse_grid_coordinate(content)
    log(f"  response: {content!r}", level="info")
    log(f"  parsed: {parsed}", level="info")
    log(f"  usage: promptTokenCount={usage.get('promptTokenCount')}, "
        f"candidatesTokenCount={usage.get('candidatesTokenCount')}", level="info")
    if not parsed:
        raise FullRunError(
            f"Live test parse failure — Gemini returned {content!r} which "
            f"could not be parsed as a grid coordinate. Refusing to launch batch.")
    log("  ✓ live test passed: Gemini authenticated, response parses correctly",
        level="ok")
    (sandbox / "live_test_record.json").write_text(json.dumps({
        "query_id": q["query_id"],
        "raw_response": content,
        "parsed_coordinates": parsed,
        "usage": usage,
        "model_id": model_id,
    }, indent=2))


# ── Stage 4: submit ────────────────────────────────────────────────

def stage_submit(args, sandbox: Path) -> list[Path]:
    """Submit each repetition into its own subdirectory so the three reps
    cannot stomp on each other. Each rep's submission goes through
    pipeline.cmd_submit with --models gemini-3.1-pro, redirected via the
    DENTAL_MLLM_RESULTS_DIR env var.

    Gemini submits are ASYNC (file-based batch). cmd_submit returns once
    every batch is CREATED with state BATCH_STATE_PENDING; stage_wait then
    polls until each batch reaches a terminal state.
    """
    log("Stage 4: submit (Gemini async file-based batch — uploads JSONL, creates batches)",
        level="step")
    run_dirs: list[Path] = []
    master_index = sandbox / "query_index.json"

    # Terminal states that mean the bundle does NOT need re-submission.
    # We mirror the Anthropic orchestrator's done_states pattern:
    #  - "succeeded" / "completed_with_failures" → outer skips, work is done
    #  - "submitted" → NOT included; outer re-calls cmd_submit which inner-
    #    skips because the bundle is mid-flight (status != failed/partial)
    #  - "failed" → NOT included; we WANT to retry failed bundles
    #  - "completed" (legacy from the pre-fix sync attempt) → NOT included;
    #    those bundles never actually produced data and should be retried
    DONE_STATES = ("succeeded", "completed_with_failures")

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
            already = {
                k for k, v in tracking.items()
                if v.get("status") in DONE_STATES
            }
            if expected.issubset(already):
                log(f"  run{i}: already submitted+terminal — skipping", level="ok")
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
    """Dry-run: emit fabricated Gemini chunk-JSON files matching the format
    cmd_download writes for real runs, so downstream stages (parse, summary)
    can run without ever calling the API.

    Real flow: cmd_submit returns batch IDs (async). cmd_status polls them to
    BATCH_STATE_SUCCEEDED. cmd_download fetches the result JSONL and converts
    it to chunk-JSON files. The fake responses below mirror what
    cmd_download produces, with status='succeeded' in tracking.
    """
    rng = random.Random(42)
    responses_dir = run_dir / "responses"
    responses_dir.mkdir(parents=True, exist_ok=True)
    queries = json.loads((run_dir / "query_index.json").read_text())

    # Mirror pipeline.GOOGLE_FILE_CHUNK_SIZE (450) so the per-batch counts
    # match what the real flow would produce.
    CHUNK = 450

    tracking: dict = {}
    for strat in STRATEGIES:
        batch_ids = []
        batch_progress = {}
        for chunk_idx, start in enumerate(range(0, len(queries), CHUNK)):
            chunk = queries[start:start + CHUNK]
            paired = []
            for q in chunk:
                cell = chr(ord("A") + rng.randint(0, 7)) + str(rng.randint(1, 16))
                paired.append({
                    "custom_id": f"{q['query_id']}_{strat}",
                    "response": {
                        "candidates": [{
                            "content": {"parts": [{"text": cell}]},
                        }],
                        "usageMetadata": {
                            "promptTokenCount": 2400,
                            "candidatesTokenCount": 5,
                        },
                    },
                })
            chunk_path = responses_dir / f"{MODEL_KEY}_{strat}_chunk{chunk_idx:03d}.json"
            chunk_path.write_text(json.dumps(paired))
            fake_bid = f"batches/dry-run-{strat}-{chunk_idx:03d}"
            batch_ids.append(fake_bid)
            batch_progress[fake_bid] = {
                "state": "BATCH_STATE_SUCCEEDED",
                "batchStats": {"requestCount": str(len(chunk)),
                                "completedRequestCount": str(len(chunk))},
                "output_file": f"files/dry-run-{strat}-{chunk_idx:03d}",
            }
        tracking[f"{MODEL_KEY}_{strat}"] = {
            "provider": "google",
            "batch_ids": batch_ids,
            "batch_progress": batch_progress,
            "status": "succeeded",
            "model": MODEL_KEY,
            "strategy": strat,
        }
    (run_dir / "batch_tracking.json").write_text(json.dumps(tracking, indent=2))


# ── Stage 5: wait (poll Google batch states until terminal) ───────

def stage_wait(args, run_dirs: list[Path]) -> None:
    """Poll cmd_status repeatedly until every batch in every bundle reaches a
    Google-terminal state (BATCH_STATE_SUCCEEDED / FAILED / CANCELLED / EXPIRED).

    Gemini batches are async with no SLA on speed; this loop typically runs
    for 30 min - several hours per rep.
    """
    if args.dry_run:
        return
    log("Stage 5: wait — poll Google batch state until terminal",
        level="step")
    poll_seconds = getattr(args, "poll_seconds", 120)
    DONE_BUNDLE_STATES = ("succeeded", "completed_with_failures", "failed")
    for rd in run_dirs:
        log(f"  polling {rd.name}", level="step")
        while True:
            run_pipeline([str(VENV_PY), "pipeline.py", "status"],
                         sandbox=rd, capture=True)
            tracking = json.loads((rd / "batch_tracking.json").read_text())
            all_done = True
            for name, info in tracking.items():
                if info.get("status") not in DONE_BUNDLE_STATES:
                    all_done = False
                    break
            if all_done:
                log(f"    {rd.name} all bundles terminal", level="ok")
                # Refuse to proceed if every bundle failed — saves the
                # operator from wasting time on a fully-broken run.
                problems = [
                    f"{name}: {info.get('status')}"
                    for name, info in tracking.items()
                    if info.get("status") == "failed"
                ]
                if problems:
                    raise FullRunError(
                        f"{rd.name}: every batch in some bundle terminal-failed:\n"
                        + "\n".join(f"    {p}" for p in problems))
                break
            log(f"    not yet terminal; sleeping {poll_seconds}s", level="info")
            time.sleep(poll_seconds)


# ── Stage 6: download (fetch result files + convert to chunk JSON) ─

def stage_download(args, run_dirs: list[Path]) -> None:
    """Download the result file behind every SUCCEEDED Google batch and
    convert it to the chunk-JSON format the parser already understands."""
    if args.dry_run:
        return
    log("Stage 6: download Google batch result files", level="step")
    for rd in run_dirs:
        run_pipeline([str(VENV_PY), "pipeline.py", "download"], sandbox=rd)


# ── Stage 7: parse ─────────────────────────────────────────────────

def stage_parse(args, run_dirs: list[Path]) -> None:
    """Parse downloaded chunk files → parsed_responses.json per rep.

    Idempotency guard: if parsed_responses.json already exists for a rep,
    skip re-parsing. This protects post-processing artifacts (e.g., merged
    re-queries) from being clobbered when the orchestrator is re-run with
    --repetitions > 1 to add more reps. To force re-parse, delete the
    rep's parsed_responses.json before re-running.
    """
    log("Stage 7: parse responses → per-run parsed_responses.json",
        level="step")
    for rd in run_dirs:
        pr = rd / "parsed_responses.json"
        if pr.exists():
            log(f"  {rd.name}: parsed_responses.json already exists — "
                f"skipping re-parse (delete to force re-parse)", level="ok")
            continue
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
            f"Re-locked after Gemini 3.1 Pro full run: {reason}\n"
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
                    help="Google batch poll interval (default: 120s)")
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

    # Refuse if sandbox is a frozen / canonical / other-model directory
    FORBIDDEN = [
        "results_full", "results_consensus", "data",
        "results_ablation_no_tooth_num", "results_ablation_patient_left",
        "results_ablation_no_LR", "results_pilot", "results_pilot_v2",
        "results_full_claude",
    ]
    for forbidden in FORBIDDEN:
        p = (ROOT / forbidden).resolve()
        if sandbox == p or p in sandbox.parents:
            raise FullRunError(
                f"REFUSING to use {sandbox} — it is or lives inside {p}. "
                f"Pick a sandbox name that starts with 'results_full_gemini'.")

    # --only-prepare: pure file-ops + SHA computation, no API surface.
    # Skip cost confirmation, preflight subprocess, and finally:relock.
    if args.only_prepare:
        log("--only-prepare: running Stage 1 only (no API calls, .api_lock untouched)",
            level="step")
        stage_prepare_sandbox(args, sandbox)
        log("Stage 1 complete. Sandbox prepared at " + str(sandbox), level="ok")
        return

    # Cost confirmation prompt unless dry-run
    expected_cost = 16.0  # conservative estimate from preflight
    n_calls = EXPECTED_N_QUERIES * len(STRATEGIES) * args.repetitions
    if not args.dry_run:
        print("\n" + "=" * 70)
        print(f"  GEMINI 3.1 PRO FULL RUN — COST CONFIRMATION")
        print("=" * 70)
        print(f"  Sandbox:            {sandbox}")
        print(f"  Queries:            {EXPECTED_N_QUERIES}")
        print(f"  Strategies:         {list(STRATEGIES)}")
        print(f"  Repetitions:        {args.repetitions}")
        print(f"  Total API calls:    {n_calls:,}")
        print(f"  Naive cost est.:    ~${expected_cost:.0f}")
        print(f"  Live single-call:   +~$0.005")
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
        log("Gemini 3.1 Pro full run complete", level="ok")
    finally:
        relock("orchestrator finally")


if __name__ == "__main__":
    main()
