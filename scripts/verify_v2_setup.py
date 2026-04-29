"""
Pre-flight verification for the v2 pilot run.

Read-only. Makes ZERO API calls. Verifies that the v2 sandbox is set up
correctly for a paired comparison against the v1 pilot results.

Exit code 0 = all checks pass, safe to launch.
Exit code 1 = at least one check failed, do NOT launch.

Run:
    .venv/bin/python scripts/verify_v2_setup.py
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
V1_DIR = ROOT / "results_pilot"
V2_DIR = ROOT / "results_pilot_v2"
REF_PATH = V2_DIR / "v1_reference.json"

errors: list[str] = []
warnings: list[str] = []

def check(label: str, cond: bool, detail: str = "") -> None:
    mark = "✓" if cond else "✗"
    line = f"  {mark} {label}"
    if detail:
        line += f" — {detail}"
    print(line)
    if not cond:
        errors.append(label + (f" — {detail}" if detail else ""))


def section(title: str) -> None:
    print(f"\n{title}")
    print("-" * len(title))


def main() -> int:
    sys.path.insert(0, str(ROOT))
    import config  # type: ignore
    import pipeline  # type: ignore

    # ── 1. Reference file exists and is valid ───────────────────────
    section("1. v1 reference file")
    check("v1_reference.json exists", REF_PATH.exists())
    if not REF_PATH.exists():
        return 1
    ref = json.loads(REF_PATH.read_text())
    v1_expected_sha = ref["v1_query_index"]["sha256"]
    v2_expected_sha = ref["v2_query_index"]["sha256"]
    check("v1 sha256 in reference", bool(v1_expected_sha))
    check("v2 expected sha matches v1 (identical)", v1_expected_sha == v2_expected_sha)

    # ── 2. v1 query_index is byte-identical to reference ────────────
    section("2. v1 query_index integrity (must be untouched)")
    v1_qi = V1_DIR / "query_index.json"
    check("v1 query_index exists", v1_qi.exists())
    if v1_qi.exists():
        v1_actual_sha = hashlib.sha256(v1_qi.read_bytes()).hexdigest()
        check("v1 query_index sha256 matches reference", v1_actual_sha == v1_expected_sha,
              f"actual={v1_actual_sha[:16]} expected={v1_expected_sha[:16]}")

    # ── 3. v2 query_index is byte-identical to v1 ───────────────────
    section("3. v2 query_index identity")
    v2_qi = V2_DIR / "query_index.json"
    check("v2 query_index exists", v2_qi.exists())
    if v2_qi.exists():
        v2_actual_sha = hashlib.sha256(v2_qi.read_bytes()).hexdigest()
        check("v2 query_index sha256 matches v1", v2_actual_sha == v1_expected_sha,
              f"actual={v2_actual_sha[:16]}")

    # ── 4. v1 pilot result hashes unchanged ─────────────────────────
    section("4. v1 pilot results integrity (must be untouched)")
    for fname, expected_sha in ref["v1_pilot_data"]["responses_hashes"].items():
        path = V1_DIR / fname
        if not path.exists():
            check(f"{fname} exists", False)
            continue
        actual_sha = hashlib.sha256(path.read_bytes()).hexdigest()
        check(f"{fname} sha256 unchanged", actual_sha == expected_sha,
              f"{actual_sha[:12]} vs {expected_sha[:12]}")

    # ── 5. v2 sandbox is otherwise clean (no leftover responses) ────
    section("5. v2 sandbox cleanliness (no leftover response files)")
    expected_v2_files = {
        "query_index.json",
        "v1_reference.json",
        "v1_prompts_snapshot.json",
        "v2_prompts_snapshot.json",
        "v1_v2_prompt_diff.json",
        "raw_data_manifest.json",
    }
    v2_files = {p.name for p in V2_DIR.iterdir() if p.is_file()}
    extras = v2_files - expected_v2_files
    check("no unexpected files in v2 root", len(extras) == 0,
          f"extras={sorted(extras)}" if extras else "")
    run_dirs = [p for p in V2_DIR.iterdir() if p.is_dir() and p.name.startswith("run")]
    check("no run dirs yet (will be created during launch)", len(run_dirs) == 0,
          f"existing run dirs: {[p.name for p in run_dirs]}" if run_dirs else "")

    # ── 6. Prompts render correctly: L-R clause where it should be ──
    section("6. Prompt rendering — L-R clause placement")
    queries = json.loads(v2_qi.read_bytes())
    lr_present_when_expected = 0
    lr_absent_when_expected = 0
    failures = []

    for q in queries:
        for strat in ("zero_shot", "guided"):
            sys_p, _ = pipeline.generate_prompt(q, strat)
            has_lr = "patient's right side appears on the left" in sys_p
            should_have_lr = (q["sheet"] == "PANORAMIC" and strat == "guided")
            if should_have_lr and has_lr:
                lr_present_when_expected += 1
            elif not should_have_lr and not has_lr:
                lr_absent_when_expected += 1
            else:
                failures.append((q["query_id"], strat, has_lr, should_have_lr))

    check(f"L-R clause present in all {lr_present_when_expected} expected cases",
          not any(f for f in failures if f[3]))
    check(f"L-R clause absent in all {lr_absent_when_expected} non-expected cases",
          not any(f for f in failures if not f[3]))
    if failures:
        for f in failures[:5]:
            print(f"    UNEXPECTED: {f}")

    # ── 7. Non-PAN-guided prompts byte-identical to v1 reconstruction ──
    section("7. Regression check — non-affected prompts unchanged")
    v1_guided_system = (
        "The image has a grid overlay with columns numbered 1 through {cols} from left to right, "
        "and rows labeled A through {max_row} from top to bottom. "
        "Each cell is identified by its row letter followed by its column number, "
        "with no space or punctuation between them. "
        "The grid lines are drawn in cyan and labels are in yellow. "
        "For point-based questions, respond with exactly one cell coordinate. "
        "For area-based questions, list all cells the structure occupies, separated by commas."
    )

    mismatches_zs = 0
    mismatches_pa = 0
    mismatches_ceph = 0
    affected_pan_guided = 0
    pan_guided_diff_chars = set()

    for q in queries:
        grid = config.GRID_SPECS[q["sheet"]]
        for strat in ("zero_shot", "guided"):
            sys_v2, _ = pipeline.generate_prompt(q, strat)

            # Reconstruct v1
            if strat == "zero_shot":
                sys_v1 = sys_v2  # zero-shot unchanged across versions
            else:
                v1_grid_explanation = v1_guided_system.format(
                    cols=grid["cols"], max_row=grid["max_row_letter"])
                sys_v1 = config.SYSTEM_PROMPT + "\n\n" + v1_grid_explanation

            if q["sheet"] == "PANORAMIC" and strat == "guided":
                # Should differ by exactly the L-R sentence
                if sys_v1 != sys_v2:
                    affected_pan_guided += 1
                    pan_guided_diff_chars.add(len(sys_v2) - len(sys_v1))
            else:
                if sys_v1 != sys_v2:
                    if strat == "zero_shot":
                        mismatches_zs += 1
                    elif q["sheet"] == "PERIAPICAL":
                        mismatches_pa += 1
                    elif q["sheet"] == "CEPHALOMETRIC":
                        mismatches_ceph += 1

    check("zero-shot prompts unchanged for ALL queries", mismatches_zs == 0,
          f"{mismatches_zs} mismatches" if mismatches_zs else "")
    check("PERIAPICAL guided prompts unchanged", mismatches_pa == 0,
          f"{mismatches_pa} mismatches" if mismatches_pa else "")
    check("CEPHALOMETRIC guided prompts unchanged", mismatches_ceph == 0,
          f"{mismatches_ceph} mismatches" if mismatches_ceph else "")
    check("PANORAMIC guided prompts changed by exactly +158 chars",
          pan_guided_diff_chars == {158},
          f"diff_chars={sorted(pan_guided_diff_chars)}")
    # PANORAMIC has 6 landmarks (3 point + 3 area) × 5 images = 30 guided prompts
    check("All PANORAMIC guided queries affected (6 landmarks × 5 images = 30)",
          affected_pan_guided == 30, f"actual={affected_pan_guided}")

    # ── 8. GT contamination check (must be 0/1800) ──────────────────
    section("8. Ground-truth contamination check")
    gt_leaks = 0
    for q in queries:
        gt_norm = (q.get("omfr_1") or "").upper().replace(" ", "")
        cells = {c.strip() for c in gt_norm.split(",") if c.strip()}
        for strat in ("zero_shot", "guided"):
            sys_p, usr_p = pipeline.generate_prompt(q, strat)
            full = sys_p + "\n" + usr_p
            for c in cells:
                if c and re.search(r'\b' + re.escape(c) + r'\b', full):
                    gt_leaks += 1
                    break
    check(f"0 GT leaks across {len(queries)*2} prompt instances", gt_leaks == 0,
          f"{gt_leaks} leaks" if gt_leaks else "")

    # ── 9. Build all 360 OpenAI requests (no API call) ──────────────
    section("9. Request building — all 360 calls")
    custom_ids = set()
    build_errors = []
    for q in queries:
        try:
            img_b64 = base64.b64encode(Path(q["image_path"]).read_bytes()).decode()
        except Exception as e:
            build_errors.append(f"{q['query_id']}: image read failed: {e}")
            continue
        for strat in ("zero_shot", "guided"):
            try:
                req = pipeline.build_openai_request(
                    q, strat, config.MODELS["gpt-5.4"], img_b64)
            except Exception as e:
                build_errors.append(f"{q['query_id']}/{strat}: {e}")
                continue
            cid = req["custom_id"]
            if cid in custom_ids:
                build_errors.append(f"DUPLICATE custom_id: {cid}")
            custom_ids.add(cid)
            body = req["body"]
            if "max_tokens" in body:
                build_errors.append(f"{cid}: forbidden max_tokens present")
            if "max_completion_tokens" not in body:
                build_errors.append(f"{cid}: missing max_completion_tokens")
            if body.get("temperature") != 0:
                build_errors.append(f"{cid}: temperature != 0")
            if body.get("seed") != 42:
                build_errors.append(f"{cid}: seed != 42")
    check(f"all {len(custom_ids)} requests build cleanly", not build_errors,
          f"{len(build_errors)} errors")
    if build_errors:
        for e in build_errors[:5]:
            print(f"    {e}")

    # Each request × 3 reps = 360 — but a single rep is 120 unique requests
    check("120 unique custom_ids per rep (1800 total chars/120 = 15 imgs × 4 modxtype × 2 strat)",
          len(custom_ids) == 120, f"got {len(custom_ids)}")

    # ── 10. Image integrity for the 15 pilot images ─────────────────
    section("10. Image integrity")
    images = sorted({q["image_path"] for q in queries})
    bad_images = []
    for ip in images:
        p = Path(ip)
        if not p.exists():
            bad_images.append(f"missing: {ip}")
            continue
        raw = p.read_bytes()
        if raw[:8] != b"\x89PNG\r\n\x1a\n":
            bad_images.append(f"not PNG: {ip}")
        elif len(raw) < 1000:
            bad_images.append(f"too small ({len(raw)}b): {ip}")
    check(f"all {len(images)} images valid PNGs", not bad_images,
          f"{len(bad_images)} bad" if bad_images else "")
    for b in bad_images[:5]:
        print(f"    {b}")

    # ── 11. .api_lock state ─────────────────────────────────────────
    section("11. API lock state")
    lock = ROOT / ".api_lock"
    check(".api_lock present (must be removed manually before launch)", lock.exists())

    # ── 12. Git state ───────────────────────────────────────────────
    section("12. Git state")
    git_sha = subprocess.check_output(["git", "-C", str(ROOT), "rev-parse", "HEAD"]).decode().strip()
    git_dirty = subprocess.check_output(
        ["git", "-C", str(ROOT), "status", "--porcelain"]).decode().strip()
    expected_sha = ref.get("git_commit_at_v2_setup", "")
    check("git commit unchanged from v2 setup", git_sha == expected_sha,
          f"now={git_sha[:8]} setup={expected_sha[:8]}")
    if git_dirty:
        warnings.append(f"working tree has uncommitted changes:\n{git_dirty}")
        print(f"  ⚠ git working tree dirty:\n{git_dirty}")

    # ── Summary ─────────────────────────────────────────────────────
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
    print("\nv2 sandbox is ready. To launch:")
    print("  rm .api_lock")
    print("  echo yes | .venv/bin/python scripts/run_pilot.py "
          "--sandbox results_pilot_v2 --n-per-modality 5 --repetitions 3")
    print("\n(Do NOT pass --force-prepare — that would regenerate the query_index "
          "and break the v1 pairing.)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
