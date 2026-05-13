"""
Pre-flight verification for the CLAUDE SONNET 4.6 full 900-query run.

Read-only. Makes ZERO API calls. Refuses to declare ready unless every gate
passes. Run BEFORE removing .api_lock and submitting.

Run:
    .venv/bin/python scripts/verify_full_run_claude_setup.py \\
        [--sandbox results_full_claude]

Design goals (same as the Gemini preflight; Anthropic-specific deltas only):
- ISOLATION: sandbox must be a NEW directory, outside results_full/,
  results_consensus/, results_ablation_*/, results_pilot*/, results_full_gemini/.
- PROMPT IDENTITY: every (query × strategy) prompt that Claude receives must
  be BYTE-IDENTICAL to what GPT-5.4 received in the v2 main run. We assert
  this against results_full/prompts_used.json for all 900 × 2 = 1,800 prompts.
- COST CAP: refuse if projected cost exceeds the budget cap (default $30).
- SHA ANCHORING: source Excel and v2 query_index must match v2_manifest.json.
- PROVIDER SANITY: build an Anthropic request via pipeline.build_anthropic_request,
  inspect resulting JSON, confirm temperature=0, model_id correct, image-bytes
  embedded as base64 image block, no `seed` (Anthropic batches don't accept it).
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
    ap.add_argument("--sandbox", default="results_full_claude",
                    help="Sandbox directory (default: results_full_claude)")
    ap.add_argument("--budget-cap-usd", type=float, default=30.0,
                    help="Hard cost cap in USD (default: $30)")
    ap.add_argument("--n-reps", type=int, default=3,
                    help="Number of repetitions planned (default: 3)")
    args = ap.parse_args()

    sys.path.insert(0, str(ROOT))
    import config        # type: ignore  # noqa: E402
    import pipeline      # type: ignore  # noqa: E402

    sandbox = (ROOT / args.sandbox).resolve()
    print(f"Claude full-run sandbox: {sandbox}")

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
        "results_full_gemini",
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
    check(f"sandbox name starts with 'results_full_claude'",
          sandbox.name.startswith("results_full_claude"))

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
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    check("ANTHROPIC_API_KEY loaded", bool(api_key))
    # Anthropic API keys start with "sk-ant-" and are ~100+ chars
    check("ANTHROPIC_API_KEY looks valid (sk-ant-, length≥80)",
          api_key.startswith("sk-ant-") and len(api_key) >= 80,
          f"length={len(api_key)} prefix={api_key[:8] if api_key else ''!r}")

    # ── 3. .api_lock ────────────────────────────────────────────────
    section("3. API lock")
    lock = ROOT / ".api_lock"
    if lock.exists():
        print(f"  ℹ .api_lock IS present — remove it (`rm .api_lock`) just "
              f"before you launch the orchestrator.")
        warn("Reminder: remove .api_lock before invoking run_full_run_claude.py")
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

    # ── 5. Sandbox query_index integrity ───────────────────────────
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
              f"{n_drift_sys} drifts — Claude would see DIFFERENT prompts than GPT-5.4!")
        check(f"0 USER-prompt drifts vs GPT-5.4 v2 main-run",
              n_drift_usr == 0,
              f"{n_drift_usr} drifts — Claude would see DIFFERENT prompts than GPT-5.4!")

    # ── 7. Strategy coverage ───────────────────────────────────────
    section("7. Strategy coverage")
    expected_strategies = ["zero_shot", "guided"]
    check(f"config.STRATEGIES == {expected_strategies}",
          config.STRATEGIES == expected_strategies,
          f"got {config.STRATEGIES}")

    # ── 8. Model config sanity ─────────────────────────────────────
    section("8. Claude model config")
    check("'claude-sonnet-4.6' in config.MODELS", "claude-sonnet-4.6" in config.MODELS)
    if "claude-sonnet-4.6" in config.MODELS:
        mc = config.MODELS["claude-sonnet-4.6"]
        check("provider == 'anthropic'", mc.get("provider") == "anthropic")
        check("temperature == 0", mc.get("temperature") == 0)
        check("seed is NOT set (Anthropic batches don't support seed)",
              "seed" not in mc,
              f"unexpected seed={mc.get('seed')}")
        check("batch_api == True", mc.get("batch_api") is True)
        check("model_id present", bool(mc.get("model_id")))
        check("max_output_tokens >= 50 (enough for grid coord)",
              mc.get("max_output_tokens", 0) >= 50,
              f"got {mc.get('max_output_tokens')}")
        print(f"    model_id: {mc.get('model_id')}")
        print(f"    max_output_tokens: {mc.get('max_output_tokens')}")

    # ── 9. Sample Anthropic request structure ───────────────────────
    section("9. Sample Anthropic request structure")
    if queries:
        q = queries[0]
        try:
            img_b64 = base64.b64encode(Path(q["image_path"]).read_bytes()).decode()
            req = pipeline.build_anthropic_request(
                q, "guided", config.MODELS["claude-sonnet-4.6"], img_b64)
            params = req["params"]
            check("request has 'model' field", "model" in params)
            check("request has 'system' (system prompt as string)",
                  "system" in params and isinstance(params["system"], str))
            check("request has 'messages' (list)",
                  "messages" in params and isinstance(params["messages"], list))
            check("temperature == 0", params.get("temperature") == 0)
            check("max_tokens set (Anthropic uses 'max_tokens', not 'max_completion_tokens')",
                  "max_tokens" in params)
            check("no 'seed' in params (Anthropic batches don't accept it)",
                  "seed" not in params)
            check("custom_id format correct",
                  req["custom_id"] == f"{q['query_id']}_guided")
            check("system prompt is canonical",
                  params["system"].startswith(
                      "You are an expert Oral and Maxillofacial Radiologist"))
            content = params["messages"][0]["content"]
            check("image embedded as base64 PNG block",
                  content[0].get("type") == "image"
                  and content[0]["source"].get("type") == "base64"
                  and content[0]["source"].get("media_type") == "image/png"
                  and content[0]["source"]["data"] == img_b64)
            check("user-prompt text is second content block",
                  content[1].get("type") == "text"
                  and content[1].get("text"))
        except Exception as e:
            check(f"build_anthropic_request did not raise (got {type(e).__name__})", False,
                  str(e))

    # ── 10. Cost projection ─────────────────────────────────────────
    section("10. Cost projection")
    # Use observed GPT-5.4 v2 main-run token rates as the prompt-token-count proxy.
    # Claude's tokeniser differs from GPT's but the order-of-magnitude is the same;
    # the image token contribution (~1,200-2,500 for high-res) dominates.
    prompt_tok_per_call_est = None
    if queries:
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
                      f"(proxy — Claude will be close but not identical)")

    mc = config.MODELS.get("claude-sonnet-4.6", {})
    in_price = mc.get("batch_input_price_per_1m", 1.50)
    out_price = mc.get("batch_output_price_per_1m", 7.50)
    n_calls_planned = len(queries) * len(expected_strategies) * args.n_reps  # 5400
    # Use observed token rates; add 10% safety margin for tokeniser differences.
    claude_input_tok_est = (prompt_tok_per_call_est or 2200) * 1.10
    avg_out_tok = 10  # 5-20 typical
    est_in_cost = claude_input_tok_est * in_price * n_calls_planned / 1e6
    est_out_cost = avg_out_tok * out_price * n_calls_planned / 1e6
    naive_total = est_in_cost + est_out_cost
    print(f"    Planned: {n_calls_planned:,} calls "
          f"(900 q × {len(expected_strategies)} strat × {args.n_reps} reps)")
    print(f"    Est. input  tok/call: {claude_input_tok_est:.0f} @ ${in_price}/1M → ${est_in_cost:.2f}")
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
