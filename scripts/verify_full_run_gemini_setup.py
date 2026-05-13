"""
Pre-flight verification for the GEMINI 3.1 PRO full 900-query run.

Read-only. Makes ZERO API calls. Refuses to declare ready unless every gate
passes. Run BEFORE removing .api_lock and submitting.

Run:
    .venv/bin/python scripts/verify_full_run_gemini_setup.py \\
        [--sandbox results_full_gemini]

Design goals
------------
- ISOLATION: the sandbox must be a NEW directory, outside results_full/,
  results_consensus/, results_ablation_*/, results_pilot*/, results_full_claude/.
- PROMPT IDENTITY: every (query × strategy) prompt that Gemini receives must
  be BYTE-IDENTICAL to what GPT-5.4 received in the v2 main run. We assert
  this against results_full/prompts_used.json (the retroactive snapshot of the
  GPT-5.4 v2 main run) for all 900 queries × 2 strategies = 1,800 prompts.
- COST CAP: refuse if the projected cost exceeds the budget cap (default $30).
- SHA ANCHORING: the source Excel and the v2 query_index must match the SHAs
  recorded in results_consensus/v2_manifest.json — otherwise we'd be running
  Gemini on a different dataset than GPT-5.4 saw.
- PROVIDER SANITY: build a Gemini request live from pipeline.build_google_request,
  inspect the resulting JSON structure, confirm temperature=0, seed=42,
  model_id correct, image-bytes embedded.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

errors: list[str] = []
warnings: list[str] = []


def check(label: str, cond: bool, detail: str = "") -> None:
    mark = "✓" if cond else "✗"
    if cond:
        print(f"  {mark} {label}")
    else:
        line = f"  {mark} {label}"
        if detail:
            line += f" — {detail}"
        print(line)
        errors.append(label + (f" — {detail}" if detail else ""))


def warn(msg: str) -> None:
    warnings.append(msg)
    print(f"  ⚠ {msg}")


def section(title: str) -> None:
    print(f"\n{title}")
    print("-" * len(title))


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def _summarise() -> int:
    print("\n" + "=" * 70)
    if errors:
        print(f"✗ {len(errors)} CHECK(S) FAILED — DO NOT LAUNCH")
        for e in errors:
            print(f"  • {e}")
        return 1
    print(f"✓ ALL CHECKS PASSED")
    if warnings:
        print(f"\n⚠ {len(warnings)} warning(s):")
        for w in warnings:
            print(f"  • {w}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--sandbox", default="results_full_gemini",
                    help="Sandbox directory (default: results_full_gemini)")
    ap.add_argument("--budget-cap-usd", type=float, default=30.0,
                    help="Hard cost cap in USD (default: $30)")
    ap.add_argument("--n-reps", type=int, default=3,
                    help="Number of repetitions planned (default: 3)")
    args = ap.parse_args()

    sys.path.insert(0, str(ROOT))
    import config        # type: ignore  # noqa: E402
    import pipeline      # type: ignore  # noqa: E402

    sandbox = (ROOT / args.sandbox).resolve()
    print(f"Gemini full-run sandbox: {sandbox}")

    # ── 1. Sandbox isolation ───────────────────────────────────────
    section("1. Sandbox isolation")
    FORBIDDEN = [
        "results_full",
        "results_consensus",
        "results_ablation_no_tooth_num",
        "results_ablation_patient_left",
        "results_ablation_no_LR",
        "results_pilot",
        "results_pilot_v2",
        "results_full_claude",
        "data",
    ]
    forbidden_dirs = []
    for name in FORBIDDEN:
        p = (ROOT / name).resolve()
        if sandbox == p or p in sandbox.parents:
            forbidden_dirs.append(str(p))
    check("sandbox is OUTSIDE every frozen / canonical / other-model dir",
          not forbidden_dirs,
          f"sandbox is inside {forbidden_dirs}" if forbidden_dirs else "")
    check(f"sandbox name starts with 'results_full_gemini'",
          sandbox.name.startswith("results_full_gemini"))

    # ── 2. Environment ─────────────────────────────────────────────
    section("2. Environment")
    check(".venv exists", (ROOT / ".venv" / "bin" / "python").exists())
    check(".env file exists", (ROOT / ".env").exists())
    if (ROOT / ".env").exists():
        for raw in (ROOT / ".env").read_text().splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            if line.startswith("export "):
                line = line[len("export "):]
            k, _, v = line.partition("=")
            v = v.strip().strip('"').strip("'")
            os.environ.setdefault(k.strip(), v)
    api_key = os.environ.get("GOOGLE_API_KEY", "")
    check("GOOGLE_API_KEY loaded", bool(api_key))
    # Google API keys typically start with "AIza" and are ~39 chars
    check("GOOGLE_API_KEY looks valid (AIza, length≥30)",
          api_key.startswith("AIza") and len(api_key) >= 30,
          f"length={len(api_key)} prefix={api_key[:4] if api_key else ''!r}")

    # ── 3. .api_lock ────────────────────────────────────────────────
    section("3. API lock")
    lock = ROOT / ".api_lock"
    if lock.exists():
        print(f"  ℹ .api_lock IS present — remove it (`rm .api_lock`) just "
              f"before you launch the orchestrator.")
        warn("Reminder: remove .api_lock before invoking run_full_run_gemini.py")
    else:
        print(f"  ℹ .api_lock is NOT present — orchestrator will re-lock "
              f"automatically after the run completes.")

    # ── 4. SHA anchors — same dataset as GPT-5.4 v2 ────────────────
    section("4. Cryptographic anchors")
    final_xlsx = ROOT / "data" / "Final_Dental_MLLM_Benchmark_Data.xlsx"
    check("Final Excel exists", final_xlsx.exists())
    excel_sha = ""
    if final_xlsx.exists():
        excel_sha = sha256_file(final_xlsx)
        print(f"    Final Excel SHA-256: {excel_sha[:16]}…")
    v2_manifest = ROOT / "results_consensus" / "v2_manifest.json"
    check("v2_manifest.json exists", v2_manifest.exists())
    if v2_manifest.exists() and excel_sha:
        mf = json.loads(v2_manifest.read_text())
        check("Final Excel SHA matches v2 manifest (same dataset as GPT-5.4)",
              mf.get("source_excel_sha256") == excel_sha,
              f"manifest {mf.get('source_excel_sha256','')[:16]}… "
              f"vs disk {excel_sha[:16]}…")
    v2_qi = ROOT / "results_consensus" / "query_index.json"
    check("v2 query_index exists", v2_qi.exists())
    v2_qi_sha = sha256_file(v2_qi) if v2_qi.exists() else ""
    print(f"    v2 query_index SHA: {v2_qi_sha[:16]}…")

    # ── 5. Sandbox query_index integrity (if sandbox was already prepared) ─
    section("5. Sandbox query_index")
    qi_path = sandbox / "query_index.json"
    if not qi_path.exists():
        warn(f"{qi_path} does not exist yet — orchestrator will create the "
             f"sandbox at launch time by copying results_consensus/query_index.json. "
             f"This is expected on a fresh preflight.")
        queries = json.loads(v2_qi.read_text()) if v2_qi.exists() else []
    else:
        queries = json.loads(qi_path.read_text())
        check("sandbox query_index has 900 queries", len(queries) == 900,
              f"got {len(queries)}")
        check("sandbox query_index matches v2 byte-for-byte",
              sha256_file(qi_path) == v2_qi_sha,
              f"sandbox SHA {sha256_file(qi_path)[:16]}… vs v2 SHA {v2_qi_sha[:16]}…")
    if queries:
        n_per_mod = {}
        for q in queries:
            n_per_mod[q["sheet"]] = n_per_mod.get(q["sheet"], 0) + 1
        check("modality split PAN=600, PA=150, CEPH=150",
              n_per_mod == {"PANORAMIC": 600, "PERIAPICAL": 150, "CEPHALOMETRIC": 150},
              str(n_per_mod))
        check("every query has consensus_gt",
              all(q.get("consensus_gt") for q in queries))
        check("every query has image_path that exists",
              all(Path(q["image_path"]).exists() for q in queries))
        check("no model answers in query_index (input-only)",
              not any(any(k in q for k in ("response", "parsed_coordinates"))
                      for q in queries))

    # ── 6. Prompt byte-identity to GPT-5.4 v2 main-run ─────────────
    section("6. Prompt byte-identity vs GPT-5.4 v2 main-run")
    v2_prompts_path = ROOT / "results_full" / "prompts_used.json"
    if not v2_prompts_path.exists():
        check("results_full/prompts_used.json exists (canonical reference)", False,
              "this file is the retroactive snapshot of the GPT-5.4 v2 main-run prompts")
    elif not queries:
        warn("Skipping prompt-identity check until sandbox query_index is in place")
    else:
        v2_pu = json.loads(v2_prompts_path.read_text())
        v2_by_qid = {p["query_id"]: p for p in v2_pu["prompts"]}
        n_drift_sys = 0
        n_drift_usr = 0
        n_missing = 0
        for q in queries:
            v2_entry = v2_by_qid.get(q["query_id"])
            if v2_entry is None:
                n_missing += 1
                continue
            for strat in ("zero_shot", "guided"):
                sys_p, usr_p = pipeline.generate_prompt(q, strat)
                v2_sys = v2_entry["by_strategy"][strat]["system_prompt"]
                v2_usr = v2_entry["by_strategy"][strat]["user_prompt"]
                if sys_p != v2_sys:
                    n_drift_sys += 1
                if usr_p != v2_usr:
                    n_drift_usr += 1
        check(f"0 queries missing from v2 prompts_used.json (over {len(queries)} queries)",
              n_missing == 0, f"{n_missing} missing")
        check(f"0 SYSTEM-prompt drifts vs GPT-5.4 v2 main-run "
              f"(over {len(queries) * 2} (query × strategy) renderings)",
              n_drift_sys == 0,
              f"{n_drift_sys} drifts — Gemini would see DIFFERENT prompts than GPT-5.4!")
        check(f"0 USER-prompt drifts vs GPT-5.4 v2 main-run",
              n_drift_usr == 0,
              f"{n_drift_usr} drifts — Gemini would see DIFFERENT prompts than GPT-5.4!")

    # ── 7. Strategy coverage ───────────────────────────────────────
    section("7. Strategy coverage")
    expected_strategies = ["zero_shot", "guided"]
    check(f"config.STRATEGIES == {expected_strategies}",
          config.STRATEGIES == expected_strategies,
          f"got {config.STRATEGIES}")

    # ── 8. Model config sanity ─────────────────────────────────────
    section("8. Gemini model config")
    check("'gemini-3.1-pro' in config.MODELS", "gemini-3.1-pro" in config.MODELS)
    if "gemini-3.1-pro" in config.MODELS:
        mc = config.MODELS["gemini-3.1-pro"]
        check("provider == 'google'", mc.get("provider") == "google")
        check("temperature == 0", mc.get("temperature") == 0)
        check("seed == 42 (reproducibility)", mc.get("seed") == 42)
        check("batch_api == True", mc.get("batch_api") is True)
        check("model_id present", bool(mc.get("model_id")))
        check("max_output_tokens >= 50 (enough for grid coord)",
              mc.get("max_output_tokens", 0) >= 50,
              f"got {mc.get('max_output_tokens')}")
        print(f"    model_id: {mc.get('model_id')}")
        print(f"    max_output_tokens: {mc.get('max_output_tokens')}")
        print(f"    media_resolution: {mc.get('media_resolution')}")

    # ── 9. Sample Gemini request structure ──────────────────────────
    section("9. Sample Gemini request structure")
    if queries:
        q = queries[0]
        try:
            img_b64 = base64.b64encode(Path(q["image_path"]).read_bytes()).decode()
            req = pipeline.build_google_request(
                q, "guided", config.MODELS["gemini-3.1-pro"], img_b64)
            inner = req["request"]
            check("request has 'model' field", "model" in inner)
            check("request has 'systemInstruction'", "systemInstruction" in inner)
            check("request has 'contents'", "contents" in inner)
            check("temperature == 0",
                  inner["generationConfig"].get("temperature") == 0)
            check("seed == 42",
                  inner["generationConfig"].get("seed") == 42)
            check("custom_id format correct",
                  req["custom_id"] == f"{q['query_id']}_guided")
            check("system prompt embedded",
                  inner["systemInstruction"]["parts"][0]["text"].startswith(
                      "You are an expert Oral and Maxillofacial Radiologist"))
            # First content part: image inlineData with base64 data
            parts = inner["contents"][0]["parts"]
            check("image embedded as inlineData (base64 PNG)",
                  parts[0].get("inlineData", {}).get("mimeType") == "image/png"
                  and parts[0]["inlineData"]["data"] == img_b64)
            check("user prompt is second part",
                  "text" in parts[1])
        except Exception as e:
            check(f"build_google_request did not raise (got {type(e).__name__})", False,
                  str(e))

    # ── 10. Cost projection ─────────────────────────────────────────
    section("10. Cost projection")
    # Use observed GPT-5.4 v2 main-run token rates as the proxy for prompt-token
    # count (Gemini will tokenise the same text slightly differently, but the
    # input size is the dominant cost driver and tokenisation gives ±10%).
    # For Gemini specifically, the image token count is fixed by the API
    # configuration (mediaResolution = MEDIA_RESOLUTION_HIGH → ~2200
    # tokens per image; this is the max enum value accepted by v1beta), so
    # the dominant factor is the per-call image cost.
    prompt_tok_per_call_est = None
    if queries:
        # Use observed v2 PAN/guided averages as a reasonable proxy
        v2_run1 = ROOT / "results_full" / "run1" / "responses"
        if v2_run1.exists():
            n = 0; tot_p = 0
            for f in v2_run1.glob("*.jsonl"):
                for line in open(f):
                    obj = json.loads(line)
                    u = obj.get("response", {}).get("body", {}).get("usage", {})
                    pt = u.get("prompt_tokens", 0)
                    if pt > 0:
                        tot_p += pt; n += 1
            if n > 0:
                prompt_tok_per_call_est = tot_p / n
                print(f"    Observed GPT-5.4 prompt-tok mean: {prompt_tok_per_call_est:.0f} "
                      f"(proxy — Gemini will be close but not identical)")

    mc = config.MODELS.get("gemini-3.1-pro", {})
    in_price = mc.get("batch_input_price_per_1m", 1.00)
    out_price = mc.get("batch_output_price_per_1m", 6.00)
    n_calls_planned = len(queries) * len(expected_strategies) * args.n_reps  # 5400
    # Gemini HIGH image: ~2200 tokens; text: ~200 tokens; total ~2400 prompt-tok per call.
    # Use whichever is larger of (observed-from-v2-GPT) or (Gemini-image-tokens estimate).
    gemini_input_tok_est = (prompt_tok_per_call_est or 2200) + 240  # text-only diff
    avg_out_tok = 10  # 5-20 typical
    est_in_cost = gemini_input_tok_est * in_price * n_calls_planned / 1e6
    est_out_cost = avg_out_tok * out_price * n_calls_planned / 1e6
    naive_total = est_in_cost + est_out_cost
    print(f"    Planned: {n_calls_planned:,} calls "
          f"(900 q × {len(expected_strategies)} strat × {args.n_reps} reps)")
    print(f"    Est. input  tok/call: {gemini_input_tok_est:.0f} @ ${in_price}/1M → ${est_in_cost:.2f}")
    print(f"    Est. output tok/call: {avg_out_tok} @ ${out_price}/1M → ${est_out_cost:.2f}")
    print(f"    Naive total estimate: ${naive_total:.2f}")
    check(f"naive cost ≤ ${args.budget_cap_usd:.0f} budget cap",
          naive_total <= args.budget_cap_usd,
          f"${naive_total:.2f} > ${args.budget_cap_usd:.0f}")

    # ── 11. Disk space ─────────────────────────────────────────────
    section("11. Disk space")
    free = shutil.disk_usage(str(ROOT)).free / 1e9
    check(f"≥1 GB free", free >= 1.0, f"{free:.1f} GB")

    # ── 12. Git state ──────────────────────────────────────────────
    section("12. Git state")
    git_sha = subprocess.check_output(
        ["git", "-C", str(ROOT), "rev-parse", "HEAD"]).decode().strip()
    print(f"    Current commit: {git_sha[:12]}")
    git_dirty = subprocess.check_output(
        ["git", "-C", str(ROOT), "status", "--porcelain"]).decode().strip()
    if git_dirty:
        warn("git tree dirty — uncommitted changes will not be in reproducibility manifest")
        for ln in git_dirty.splitlines()[:5]:
            print(f"      {ln}")

    return _summarise()


if __name__ == "__main__":
    sys.exit(main())
