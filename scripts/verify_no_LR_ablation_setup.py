"""
Pre-flight verification for the FDI ablation experiment (Stage 1).

Read-only. Makes ZERO API calls. Refuses to declare ready unless every gate
passes. Run BEFORE removing .api_lock and submitting.

Run:
    .venv/bin/python scripts/verify_fdi_ablation_setup.py \
        [--sandbox results_ablation_no_LR]
"""
from __future__ import annotations

import argparse
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


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--sandbox", default="results_ablation_no_LR",
                    help="Ablation sandbox directory")
    ap.add_argument("--budget-cap-usd", type=float, default=5.0,
                    help="Hard cost cap (USD)")
    args = ap.parse_args()

    sys.path.insert(0, str(ROOT))
    import config        # type: ignore  # noqa: E402
    import pipeline      # type: ignore  # noqa: E402

    sandbox = (ROOT / args.sandbox).resolve()
    print(f"Ablation sandbox: {sandbox}")

    # ── 1. Sandbox isolation ────────────────────────────────────────
    section("1. Sandbox isolation")
    forbidden_dirs = []
    for name in ("results_full", "results_consensus", "data"):
        p = (ROOT / name).resolve()
        if sandbox == p or p in sandbox.parents:
            forbidden_dirs.append(str(p))
    check("sandbox is OUTSIDE results_full/, results_consensus/, and data/",
          not forbidden_dirs,
          f"sandbox is inside {forbidden_dirs}" if forbidden_dirs else "")
    check(f"sandbox name 'results_ablation_no_LR' (or similar)",
          sandbox.name.startswith("results_ablation"))

    # ── 2. Environment ──────────────────────────────────────────────
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
    api_key = os.environ.get("OPENAI_API_KEY", "")
    check("OPENAI_API_KEY loaded", bool(api_key))
    check("OPENAI_API_KEY looks valid (sk-, length≥80)",
          api_key.startswith("sk-") and len(api_key) >= 80,
          f"length={len(api_key)}")

    # ── 3. .api_lock ────────────────────────────────────────────────
    section("3. API lock")
    # Informational only — the orchestrator does its own enforcement.
    # Stand-alone preflight runs (before launch) expect the lock to be
    # present; orchestrator-driven runs expect it to have been removed
    # immediately before invoking. Either is fine here.
    lock = ROOT / ".api_lock"
    if lock.exists():
        print(f"  ℹ .api_lock IS present — remove it (`rm .api_lock`) just "
              f"before you launch the orchestrator.")
        warn("Reminder: remove .api_lock before invoking run_no_LR_ablation.py")
    else:
        print(f"  ℹ .api_lock is NOT present — orchestrator will re-lock "
              f"automatically after the run completes.")

    # ── 4. SHA anchors ──────────────────────────────────────────────
    section("4. Cryptographic anchors")
    final_xlsx = ROOT / "data" / "Final_Dental_MLLM_Benchmark_Data.xlsx"
    check("Final Excel exists", final_xlsx.exists())
    if final_xlsx.exists():
        excel_sha = sha256_file(final_xlsx)
        print(f"    Final Excel SHA-256: {excel_sha[:16]}…")
        # Compare to v2 manifest
        v2_manifest = ROOT / "results_consensus" / "v2_manifest.json"
        if v2_manifest.exists():
            mf = json.loads(v2_manifest.read_text())
            check("Final Excel SHA matches v2 manifest",
                  mf["source_excel_sha256"] == excel_sha,
                  f"manifest declared {mf['source_excel_sha256'][:16]}… "
                  f"actual {excel_sha[:16]}…")

    # v2 query_index SHA
    v2_qi = ROOT / "results_consensus" / "query_index.json"
    if v2_qi.exists():
        v2_qi_sha = sha256_file(v2_qi)
        print(f"    v2 query_index SHA: {v2_qi_sha[:16]}…")

    # ── 5. Filtered query_index integrity ──────────────────────────
    section("5. Filtered query_index")
    qi_path = sandbox / "query_index.json"
    if not qi_path.exists():
        check(f"{qi_path} exists", False)
        print(f"    Run: DENTAL_MLLM_RESULTS_DIR={args.sandbox} "
              f".venv/bin/python pipeline.py prepare_ablation "
              f"--structures Tooth_33_Apex --expected-count 100 "
              f"--anchor-to results_full")
        # Dump errors so far and exit early
        return _summarise()

    queries = json.loads(qi_path.read_text())
    check("filter has exactly 100 queries", len(queries) == 100,
          f"got {len(queries)}")
    structures = sorted({q["structure"] for q in queries})
    check("only Tooth_33_Apex queries present",
          structures == ["Tooth_33_Apex"],
          f"got {structures}")
    sheets = sorted({q["sheet"] for q in queries})
    check("all queries are PANORAMIC", sheets == ["PANORAMIC"], f"got {sheets}")
    n_uses_fdi = sum(1 for q in queries if q.get("uses_fdi"))
    check("all queries are uses_fdi=True", n_uses_fdi == len(queries),
          f"got {n_uses_fdi}/{len(queries)}")
    n_with_no_fdi_desc = sum(
        1 for q in queries if q.get("landmark_description_en_no_fdi"))
    check("all queries have landmark_description_en_no_fdi populated",
          n_with_no_fdi_desc == len(queries),
          f"got {n_with_no_fdi_desc}/{len(queries)}")
    n_with_consensus = sum(1 for q in queries if q.get("consensus_gt"))
    check("all queries have consensus_gt populated",
          n_with_consensus == len(queries))

    # ── 6. Prompt-difference verification ──────────────────────────
    section("6. Prompt-difference verification (L–R clause removed; system-prompt-only change)")
    #
    # Expected change for each Tooth_33_Apex query (PANORAMIC, FDI-flagged):
    #   - User prompt: byte-identical to `guided` (the canonical user prompt,
    #     including the "tooth #33 (lower left canine)" phrase, is preserved).
    #   - System prompt: the panoramic L–R inversion clause is REMOVED. The
    #     remainder of the system prompt is byte-identical.
    # i.e., the diagnostic test isolates the L–R clause as the single
    # variable. If guided_no_LR recovers performance to zero-shot levels,
    # the clause itself was the proximate cause of the regression. If not,
    # the cause lies in something else (model-level prior, image features).
    LR_CLAUSE = ("In panoramic radiographs, the patient's right side appears "
                 "on the left side of the image, and the patient's left side "
                 "appears on the right side of the image.")
    sample_idxs = [0, 25, 50, 75, 99]  # 5 queries

    n_user_identical = 0
    n_system_differs = 0
    n_LR_clause_absent = 0
    n_LR_clause_present_in_guided = 0
    n_isolated_removal = 0    # removing the clause from guided yields guided_no_LR exactly

    for idx in sample_idxs:
        q = queries[idx]
        gd_sys, gd_usr = pipeline.generate_prompt(q, "guided")
        nl_sys, nl_usr = pipeline.generate_prompt(q, "guided_no_LR")

        # User prompts MUST be identical
        if nl_usr == gd_usr:
            n_user_identical += 1

        # System prompts MUST differ (L-R clause removed from one side)
        if nl_sys != gd_sys:
            n_system_differs += 1

        # L-R clause MUST be absent from guided_no_LR
        if LR_CLAUSE not in nl_sys:
            n_LR_clause_absent += 1

        # L-R clause MUST be present in canonical guided (sanity check)
        if LR_CLAUSE in gd_sys:
            n_LR_clause_present_in_guided += 1

        # Isolated removal: deleting the clause (plus the trailing space) from
        # guided's system prompt must yield guided_no_LR's system prompt exactly.
        if gd_sys.replace(LR_CLAUSE + " ", "") == nl_sys:
            n_isolated_removal += 1

    check(f"guided_no_LR user prompt == guided user prompt for all "
          f"{len(sample_idxs)} samples",
          n_user_identical == len(sample_idxs))
    check(f"guided_no_LR system prompt DIFFERS from guided system prompt "
          f"for all {len(sample_idxs)} samples",
          n_system_differs == len(sample_idxs))
    check(f"L–R clause ABSENT from guided_no_LR for all "
          f"{len(sample_idxs)} samples",
          n_LR_clause_absent == len(sample_idxs))
    check(f"L–R clause PRESENT in canonical guided (sanity) for all "
          f"{len(sample_idxs)} samples",
          n_LR_clause_present_in_guided == len(sample_idxs))
    check(f"removal is ISOLATED — deleting the L–R clause from guided "
          f"yields guided_no_LR exactly, for all {len(sample_idxs)} samples",
          n_isolated_removal == len(sample_idxs))

    # Show first sample for visual inspection — note that for this variant
    # it is the SYSTEM prompt that changes, not the user prompt.
    q0 = queries[0]
    print(f"\n    Sample prompt rendering for {q0['query_id']}:")
    nl_sys, nl_usr = pipeline.generate_prompt(q0, "guided_no_LR")
    gd_sys, gd_usr = pipeline.generate_prompt(q0, "guided")
    print(f"    --- guided SYSTEM (length {len(gd_sys)}) ---")
    print(f"    {gd_sys}")
    print(f"    --- guided_no_LR SYSTEM (length {len(nl_sys)}) ---")
    print(f"    {nl_sys}")
    print(f"    --- user prompts identical: "
          f"{'YES (correct)' if gd_usr == nl_usr else 'NO — BUG'} ---")

    # ── 7. GT contamination ────────────────────────────────────────
    section("7. GT contamination check")
    leaks = 0
    for q in queries:
        gt_norm = (q.get("consensus_gt") or "").upper().replace(" ", "")
        cells = {c.strip() for c in gt_norm.split(",") if c.strip()}
        for strat in ("zero_shot", "guided", "guided_no_LR"):
            sys_p, usr_p = pipeline.generate_prompt(q, strat)
            full = sys_p + "\n" + usr_p
            for c in cells:
                if c and re.search(r"\b" + re.escape(c) + r"\b", full):
                    leaks += 1
                    break
    check(f"0 GT leaks across {len(queries)} queries × 3 strategies "
          f"({len(queries)*3} prompt instances)",
          leaks == 0,
          f"{leaks} leaks" if leaks else "")

    # ── 8. Cost projection ─────────────────────────────────────────
    section("8. Cost projection")
    # Use observed token rates from the v2 PAN-Tooth_33 subset for accuracy
    pilot_path = ROOT / "results_full"
    n_calls_planned = len(queries) * 1 * 3  # 100 × 1 strategy × 3 reps
    if pilot_path.exists():
        prompt_tok = compl_tok = 0
        n_observed = 0
        for run in (1, 2, 3):
            for f in (pilot_path / f"run{run}" / "responses").glob("*.jsonl"):
                for line in open(f):
                    obj = json.loads(line)
                    cid = obj.get("custom_id", "")
                    # Only count PAN guided Tooth_33 calls (closest analogue
                    # to what we'll submit, since prompt length is similar)
                    if "Tooth_33_Apex_guided" not in cid:
                        continue
                    u = obj.get("response", {}).get("body", {}).get("usage", {})
                    prompt_tok += u.get("prompt_tokens", 0)
                    compl_tok += u.get("completion_tokens", 0)
                    n_observed += 1
        if n_observed > 0:
            avg_p = prompt_tok / n_observed
            avg_c = compl_tok / n_observed
            est_p = avg_p * n_calls_planned
            est_c = avg_c * n_calls_planned
            naive = est_p * 1.25e-6 + est_c * 7.5e-6
            print(f"    Observed avg per call: prompt={avg_p:.0f} tok, compl={avg_c:.1f} tok "
                  f"(from {n_observed} v2 PAN guided Tooth_33 calls)")
            print(f"    Planned: {n_calls_planned} calls (100 queries × 1 strategy × 3 reps)")
            print(f"    Naive cost estimate: ${naive:.2f}")
            check(f"naive cost ≤ ${args.budget_cap_usd:.0f} budget cap",
                  naive <= args.budget_cap_usd,
                  f"${naive:.2f} > ${args.budget_cap_usd:.0f}")

    # ── 9. Anchor verification ─────────────────────────────────────
    section("9. Anchor verification")
    anchor_path = sandbox / "ablation_anchor.json"
    if anchor_path.exists():
        anchor = json.loads(anchor_path.read_text())
        if "anchor_root" in anchor:
            anchor_root = Path(anchor["anchor_root"])
            n_drift = 0
            for rel, info in anchor["files"].items():
                p = anchor_root / rel
                if not p.exists():
                    n_drift += 1
                    continue
                if sha256_file(p) != info["sha256"]:
                    n_drift += 1
            check(f"all {len(anchor['files'])} anchored JSONLs intact",
                  n_drift == 0,
                  f"{n_drift} drifts" if n_drift else "")
        if "v2_query_index" in anchor:
            qi_p = Path(anchor["v2_query_index"]["path"])
            if qi_p.exists():
                check("v2 query_index SHA intact",
                      sha256_file(qi_p) == anchor["v2_query_index"]["sha256"])

    # ── 10. Disk space ─────────────────────────────────────────────
    section("10. Disk space")
    free = shutil.disk_usage(str(ROOT)).free / 1e9
    check(f"≥0.5 GB free in {ROOT.parent}", free >= 0.5, f"{free:.1f} GB")

    # ── 11. Git state ──────────────────────────────────────────────
    section("11. Git state")
    git_sha = subprocess.check_output(
        ["git", "-C", str(ROOT), "rev-parse", "HEAD"]).decode().strip()
    print(f"    Current commit: {git_sha[:12]}")
    git_dirty = subprocess.check_output(
        ["git", "-C", str(ROOT), "status", "--porcelain"]).decode().strip()
    if git_dirty:
        warn("git tree dirty — uncommitted changes will not be in reproducibility manifest")
        for ln in git_dirty.splitlines()[:5]:
            print(f"      {ln}")

    # ── 12. Strategies override sanity ─────────────────────────────
    section("12. Submit configuration")
    print(f"    Will submit: --strategies guided_no_LR --models gpt-5.4")
    print(f"    config.STRATEGIES (canonical, untouched): {config.STRATEGIES}")
    print(f"    config.ACTIVE_MODELS: {config.ACTIVE_MODELS}")
    check("guided_no_LR is renderable by generate_prompt",
          True)  # we already verified above

    return _summarise()


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
    print("\nTo launch (after explicit user approval):")
    print("  rm .api_lock")
    print("  echo yes | .venv/bin/python scripts/run_no_LR_ablation.py")
    print("\n(.api_lock will auto-recreate on completion.)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
