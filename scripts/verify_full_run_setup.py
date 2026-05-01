"""
Pre-flight verification for the FULL 900-query run (GPT-5.4 only).

Read-only. Makes ZERO API calls. Verifies every prerequisite for a multi-hour
production run before money is spent.

Run:
    .venv/bin/python scripts/verify_full_run_setup.py [--sandbox results_full]
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
    """Print a check; show the detail only when the check fails."""
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


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--sandbox", default="results",
                    help="Sandbox directory (default: results)")
    ap.add_argument("--n-reps", type=int, default=3,
                    help="Planned number of repetitions")
    ap.add_argument("--budget-cap-usd", type=float, default=20.0,
                    help="Maximum acceptable upper-bound cost (USD)")
    args = ap.parse_args()

    sys.path.insert(0, str(ROOT))
    import config  # type: ignore
    import pipeline  # type: ignore

    sandbox = ROOT / args.sandbox
    print(f"Sandbox: {sandbox}")

    # ── 1. Environment ─────────────────────────────────────────────
    section("1. Environment")
    check(".venv exists", (ROOT / ".venv" / "bin" / "python").exists())
    check(".env file exists", (ROOT / ".env").exists())
    # Load .env for the rest of this script's checks
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
    api_key = os.environ.get("OPENAI_API_KEY", "")
    check("OPENAI_API_KEY loaded", bool(api_key))
    check("OPENAI_API_KEY looks valid (sk-, length>=80)",
          api_key.startswith("sk-") and len(api_key) >= 80,
          f"length={len(api_key)}")

    # ── 2. .api_lock ───────────────────────────────────────────────
    section("2. API lock")
    lock = ROOT / ".api_lock"
    check(".api_lock present (must be removed manually before launch)",
          lock.exists())
    if lock.exists():
        warn("Remove with: rm .api_lock — only AFTER you've confirmed everything")

    # ── 3. Raw data integrity ──────────────────────────────────────
    section("3. Raw data integrity")
    excel = ROOT / "data" / "Dental_MLLM_Benchmark_Data.xlsx"
    check("Excel exists", excel.exists())
    excel_sha = hashlib.sha256(excel.read_bytes()).hexdigest() if excel.exists() else ""
    print(f"    Excel SHA-256: {excel_sha[:16]}...")

    # Image counts per modality folder
    expected_imgs = {"Panoramic_Grid": 100, "Periapical_Grid": 50, "Cephalometric_Grid": 50}
    img_total = 0
    bad_imgs = 0
    for sub, expected in expected_imgs.items():
        folder = ROOT / "data" / "images" / sub
        if not folder.exists():
            check(f"{sub} folder exists", False)
            continue
        pngs = sorted(p for p in folder.iterdir() if p.suffix == ".png" and p.is_file())
        check(f"{sub}: {expected} PNGs", len(pngs) == expected,
              f"got {len(pngs)}")
        for p in pngs:
            raw = p.read_bytes()
            if not raw.startswith(b"\x89PNG\r\n\x1a\n"):
                bad_imgs += 1
            elif len(raw) < 1000:
                bad_imgs += 1
        img_total += len(pngs)
    check(f"all {img_total} images valid PNGs (no zero-byte / corrupt)", bad_imgs == 0,
          f"{bad_imgs} bad" if bad_imgs else "")

    # ── 4. Query index ─────────────────────────────────────────────
    section("4. Query index")
    qi_path = sandbox / "query_index.json"
    if not qi_path.exists():
        check(f"{qi_path} exists", False)
        print(f"    Run: DENTAL_MLLM_RESULTS_DIR={args.sandbox} "
              f".venv/bin/python pipeline.py prepare")
        return 1
    qi = json.loads(qi_path.read_text())
    check(f"query_index has 900 queries", len(qi) == 900, f"got {len(qi)}")
    check(f"query_index has 200 unique images",
          len({q['image_id'] for q in qi}) == 200,
          f"got {len({q['image_id'] for q in qi})}")
    mods = {}
    for q in qi:
        mods[q["sheet"]] = mods.get(q["sheet"], 0) + 1
    check("modality split: PAN=600, PA=150, CEPH=150",
          mods == {"PANORAMIC": 600, "PERIAPICAL": 150, "CEPHALOMETRIC": 150},
          str(mods))
    check("every query has omfr_1 (ground truth)",
          all(q.get("omfr_1") for q in qi))
    check("no model answers in query_index (input-only)",
          not any(any(k in q for k in ("response", "parsed_coordinates"))
                  for q in qi))

    # ── 5. Prompt rendering — no GT contamination ──────────────────
    section("5. Prompt rendering (full set)")
    leaks = 0
    for q in qi:
        gt_norm = (q.get("omfr_1") or "").upper().replace(" ", "")
        cells = {c.strip() for c in gt_norm.split(",") if c.strip()}
        for strat in ("zero_shot", "guided"):
            sys_p, usr_p = pipeline.generate_prompt(q, strat)
            full = sys_p + "\n" + usr_p
            for c in cells:
                if c and re.search(r'\b' + re.escape(c) + r'\b', full):
                    leaks += 1
                    break
    check(f"0 GT leaks across {len(qi)*2} prompt instances", leaks == 0,
          f"{leaks} leaks" if leaks else "")

    # PAN guided has L-R clause; others don't
    lr_correct = 0
    lr_wrong = 0
    for q in qi:
        sys_p, _ = pipeline.generate_prompt(q, "guided")
        has_lr = "patient's right side appears on the left" in sys_p
        if (q["sheet"] == "PANORAMIC") == has_lr:
            lr_correct += 1
        else:
            lr_wrong += 1
    check("L-R clause in all PANORAMIC guided prompts (and only those)",
          lr_wrong == 0,
          f"{lr_wrong} wrong" if lr_wrong else "")

    # ── 6. Build a sample request to verify it's valid ──────────────
    section("6. Sample request structure")
    q = qi[0]
    img_b64 = base64.b64encode(Path(q["image_path"]).read_bytes()).decode()
    req = pipeline.build_openai_request(q, "guided",
                                          config.MODELS["gpt-5.4"], img_b64)
    body = req["body"]
    check("OpenAI request has max_completion_tokens (NOT max_tokens)",
          "max_completion_tokens" in body and "max_tokens" not in body)
    check("temperature == 0", body.get("temperature") == 0)
    check("seed == 42", body.get("seed") == 42)
    check("model == gpt-5.4", body.get("model") == "gpt-5.4")
    check("custom_id format correct",
          req["custom_id"] == f"{q['query_id']}_guided")
    check("image embedded as data URL with detail=high",
          body["messages"][1]["content"][0]["image_url"]["detail"] == "high")

    # ── 7. Cost projection ─────────────────────────────────────────
    section("7. Cost projection")
    # Use observed token rates from the v2 pilot
    pilot_path = ROOT / "results_pilot_v2"
    if pilot_path.exists():
        prompt_tok = compl_tok = 0
        n = 0
        for run in (1, 2, 3):
            for f in (pilot_path / f"run{run}" / "responses").glob("*.jsonl"):
                for line in open(f):
                    obj = json.loads(line)
                    u = obj.get("response", {}).get("body", {}).get("usage", {})
                    prompt_tok += u.get("prompt_tokens", 0)
                    compl_tok += u.get("completion_tokens", 0)
                    n += 1
        if n > 0:
            avg_p = prompt_tok / n
            avg_c = compl_tok / n
            full_calls = 900 * 2 * args.n_reps
            full_p = avg_p * full_calls
            full_c = avg_c * full_calls
            # Naive (no caching): upper bound
            naive = full_p * 1.25e-6 + full_c * 7.5e-6
            # Optimistic (similar cache benefit as pilot): observed rate
            observed_per_call = 0.65 / 360  # user reported $0.65/pilot
            optimistic = full_calls * observed_per_call
            print(f"    Pilot avg per call: prompt={avg_p:.0f} tok, compl={avg_c:.1f} tok")
            print(f"    Full run: {full_calls:,} calls")
            print(f"    Conservative (naive): ${naive:.2f}")
            print(f"    Optimistic (observed cached rate): ${optimistic:.2f}")
            check(f"naive cost ≤ ${args.budget_cap_usd:.0f} budget cap",
                  naive <= args.budget_cap_usd,
                  f"${naive:.2f} > ${args.budget_cap_usd:.0f}")
        else:
            warn("v2 pilot has no responses; using token-based projection only")

    # ── 8. Disk space ──────────────────────────────────────────────
    section("8. Disk space")
    free = shutil.disk_usage(str(ROOT)).free / 1e9
    check(f"≥1 GB free in {ROOT.parent}", free >= 1.0, f"{free:.1f} GB")

    # ── 9. Git state ───────────────────────────────────────────────
    section("9. Git state")
    git_sha = subprocess.check_output(
        ["git", "-C", str(ROOT), "rev-parse", "HEAD"]).decode().strip()
    print(f"    Current commit: {git_sha[:12]}")
    git_dirty = subprocess.check_output(
        ["git", "-C", str(ROOT), "status", "--porcelain"]).decode().strip()
    if git_dirty:
        warn("git tree dirty — uncommitted changes will not be in reproducibility manifest")
        for ln in git_dirty.splitlines()[:5]:
            print(f"      {ln}")

    # ── 10. Sandbox write protection ───────────────────────────────
    section("10. Sandbox safety")
    if sandbox.resolve() == (ROOT / "data").resolve():
        check("sandbox != data/ (would corrupt raw data)", False)
    else:
        check(f"sandbox ({sandbox.name}) is outside data/", True)

    # ── Summary ────────────────────────────────────────────────────
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
    print("\nTo launch:")
    print("  rm .api_lock")
    print(f"  echo yes | .venv/bin/python scripts/run_pilot.py "
          f"--sandbox {args.sandbox} --repetitions {args.n_reps}")
    print(f"\n(After completion, the .api_lock will be re-created automatically.)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
