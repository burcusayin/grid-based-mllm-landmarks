"""
Three-way analysis of the patient-left ablation experiment (Stage 1).

Compares for each of the 100 panoramic Tooth_33_Apex queries:
  - zero_shot         (existing v2 data, scored against consensus_gt)
  - guided            (existing v2 data, scored against consensus_gt)
  - guided_patient_left     (new data from this ablation)

Pre-specified primary outcome:
    correct-side rate = % of predictions on the image right (cols ≥ 9),
    where the GT lives for tooth 33.

Pre-specified decision criteria (locked before run):
  - guided_patient_left correct-side rate ≥ 90% → hypothesis CONFIRMED
  - 30%–90%                              → hypothesis PARTIALLY supported
  - < 30% (similar to guided's 5%)       → hypothesis REJECTED

Also reports: mean ED, SDR thresholds, three-way paired Wilcoxon, per-image
detail table, Fleiss κ (within-strategy reproducibility for guided_patient_left).

Reads:
    results_ablation_patient_left/run{1,2,3}/parsed_responses.json   (new)
    results_ablation_patient_left/ablation_anchor.json               (verifies frozen anchor)
    results_full/run{1,2,3}/responses/*.jsonl                   (existing zero_shot/guided, READ-ONLY)
    results_consensus/query_index.json                          (consensus_gt)

Writes:
    results_ablation_patient_left/ablation_analysis.json
    results_ablation_patient_left/Ablation_Report.docx (optional, can be added later)
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import statistics
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import config    # noqa: E402
import pipeline  # noqa: E402

TARGET_STRUCTURE = "Tooth_33_Apex"


# ── geometry ────────────────────────────────────────────────────────

def grid_dims(modality):
    g = config.GRID_SPECS[modality]; return g["cols"], g["rows"]


def cell_to_xy(cell, mod):
    s = (cell or "").strip().upper().replace(" ", "").replace("-", "")
    if not s:
        return None
    i = 0
    while i < len(s) and s[i].isalpha():
        i += 1
    if i != 1 or i == len(s) or not s[i:].isdigit():
        return None
    row = ord(s[:i]) - ord("A") + 1
    col = int(s[i:])
    cols, rows = grid_dims(mod)
    return (col, row) if (1 <= col <= cols and 1 <= row <= rows) else None


def parse_predicted_first_cell(text, mod):
    parsed = pipeline.parse_grid_coordinate(text or "")
    if not parsed:
        return None
    for tok in parsed:
        xy = cell_to_xy(tok, mod)
        if xy is not None:
            return xy
    return None


def euclid(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


# ── anchor verification ────────────────────────────────────────────

def verify_anchor(sandbox: Path) -> None:
    anchor_path = sandbox / "ablation_anchor.json"
    if not anchor_path.exists():
        raise RuntimeError(f"missing {anchor_path}")
    anchor = json.loads(anchor_path.read_text())
    drift = []
    if "anchor_root" in anchor:
        anchor_root = Path(anchor["anchor_root"])
        for rel, info in anchor["files"].items():
            p = anchor_root / rel
            if not p.exists():
                drift.append(f"missing {rel}")
                continue
            if sha256_file(p) != info["sha256"]:
                drift.append(f"sha mismatch {rel}")
    if "v2_query_index" in anchor:
        p = Path(anchor["v2_query_index"]["path"])
        if p.exists() and sha256_file(p) != anchor["v2_query_index"]["sha256"]:
            drift.append("v2 query_index sha mismatch")
    if drift:
        raise RuntimeError(f"Anchor drift: {drift[:5]}")
    print(f"Anchor verified: {len(anchor.get('files', {}))} JSONLs intact "
          f"+ v2 query_index intact")


# ── Wilcoxon helpers ──────────────────────────────────────────────

def paired_wilcoxon(deltas):
    from scipy import stats as scistats
    nz = [d for d in deltas if abs(d) > 1e-12]
    n_total, n_nonzero = len(deltas), len(nz)
    if n_nonzero == 0:
        return {"n_total": n_total, "n_nonzero": 0, "stat": 0.0, "p": 1.0,
                "mean_delta": 0.0, "median_delta": 0.0,
                "rank_biserial_r": 0.0}
    try:
        res = scistats.wilcoxon(nz, zero_method="wilcox",
                                alternative="two-sided", mode="auto")
        stat, p = float(res.statistic), float(res.pvalue)
    except Exception:
        stat, p = float("nan"), float("nan")
    ranks = scistats.rankdata([abs(d) for d in nz])
    Wp = sum(r for d, r in zip(nz, ranks) if d > 0)
    Wn = sum(r for d, r in zip(nz, ranks) if d < 0)
    rb = (Wp - Wn) / (Wp + Wn) if (Wp + Wn) > 0 else 0.0
    return {"n_total": n_total, "n_nonzero": n_nonzero, "stat": stat, "p": p,
            "mean_delta": sum(deltas) / n_total,
            "median_delta": statistics.median(deltas),
            "rank_biserial_r": rb}


# ── Fleiss kappa ──────────────────────────────────────────────────

def fleiss_kappa(item_rater_cat):
    if not item_rater_cat:
        return None
    R = sum(item_rater_cat[0].values())
    N = len(item_rater_cat)
    cats = sorted({c for d in item_rater_cat for c in d})
    P_j = {c: sum(d.get(c, 0) for d in item_rater_cat) / (N * R) for c in cats}
    Pi = []
    for d in item_rater_cat:
        s = sum(n * (n - 1) for n in d.values())
        Pi.append(s / (R * (R - 1)))
    P_bar = sum(Pi) / N
    Pe_bar = sum(p ** 2 for p in P_j.values())
    if 1 - Pe_bar == 0:
        return 1.0
    return (P_bar - Pe_bar) / (1 - Pe_bar)


# ── main ──────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sandbox", default="results_ablation_patient_left")
    ap.add_argument("--anchor-root", default="results_full")
    ap.add_argument("--v2-sandbox", default="results_consensus")
    args = ap.parse_args()

    sandbox = (ROOT / args.sandbox).resolve()
    anchor_root = (ROOT / args.anchor_root).resolve()
    v2_sandbox = (ROOT / args.v2_sandbox).resolve()

    verify_anchor(sandbox)

    # Build qmap from v2 query_index (consensus_gt is the canonical reference)
    queries_v2 = json.loads((v2_sandbox / "query_index.json").read_text())
    qmap_v2 = {q["query_id"]: q for q in queries_v2}
    target_qids = {q["query_id"] for q in queries_v2 if q["structure"] == TARGET_STRUCTURE}
    print(f"Target query_ids: {len(target_qids)}")

    # ── Collect predictions per (qid, strategy, rep) ───────────────
    preds = {}  # (qid, strategy, rep) -> raw_response_text
    # 1) Existing zero_shot + guided from results_full
    for rep in (1, 2, 3):
        rep_dir = anchor_root / f"run{rep}" / "responses"
        for jl in sorted(rep_dir.glob("*.jsonl")):
            for line in open(jl):
                if not line.strip():
                    continue
                obj = json.loads(line)
                cid = obj["custom_id"]
                if cid.endswith("_zero_shot"):
                    qid, strat = cid[:-len("_zero_shot")], "zero_shot"
                elif cid.endswith("_guided"):
                    qid, strat = cid[:-len("_guided")], "guided"
                else:
                    continue
                if qid not in target_qids:
                    continue
                content = (obj.get("response", {}).get("body", {})
                              .get("choices", [{}])[0].get("message", {})
                              .get("content", "") or "").strip()
                preds[(qid, strat, rep)] = content
    # 2) New guided_patient_left from sandbox
    for rep in (1, 2, 3):
        rep_dir = sandbox / f"run{rep}" / "responses"
        if not rep_dir.exists():
            print(f"WARNING: {rep_dir} missing; ablation may not be complete")
            continue
        for jl in sorted(rep_dir.glob("*.jsonl")):
            for line in open(jl):
                if not line.strip():
                    continue
                obj = json.loads(line)
                cid = obj["custom_id"]
                if not cid.endswith("_guided_patient_left"):
                    continue
                qid = cid[:-len("_guided_patient_left")]
                if qid not in target_qids:
                    continue
                content = (obj.get("response", {}).get("body", {})
                              .get("choices", [{}])[0].get("message", {})
                              .get("content", "") or "").strip()
                preds[(qid, "guided_patient_left", rep)] = content

    # Verify completeness
    expected = len(target_qids) * 3 * 3   # 100 × 3 strategies × 3 reps = 900 cells
    actual = len(preds)
    print(f"Predictions collected: {actual} / {expected} cells")
    missing = []
    for qid in target_qids:
        for strat in ("zero_shot", "guided", "guided_patient_left"):
            for rep in (1, 2, 3):
                if (qid, strat, rep) not in preds:
                    missing.append((qid, strat, rep))
    if missing:
        print(f"WARNING: {len(missing)} missing cells; first 5: {missing[:5]}")

    # ── Build per-(qid,strategy) records with metrics vs consensus_gt ────
    records = {}
    for qid in target_qids:
        q = qmap_v2[qid]
        modality = q["sheet"]
        gt_cell = cell_to_xy(q["consensus_gt"], modality)
        for strat in ("zero_shot", "guided", "guided_patient_left"):
            rep_eds = [None, None, None]
            rep_pred_cells = [None, None, None]
            rep_raws = [None, None, None]
            for rep in (1, 2, 3):
                content = preds.get((qid, strat, rep))
                rep_raws[rep - 1] = content
                if content is None:
                    continue
                pred = parse_predicted_first_cell(content, modality)
                if pred is None:
                    continue
                rep_pred_cells[rep - 1] = pred
                if gt_cell is not None:
                    rep_eds[rep - 1] = euclid(pred, gt_cell)
            valid_eds = [e for e in rep_eds if e is not None]
            mean_ed = sum(valid_eds) / len(valid_eds) if valid_eds else None
            records[(qid, strat)] = {
                "query_id": qid, "strategy": strat,
                "consensus_gt": q["consensus_gt"], "gt_cell": gt_cell,
                "rep_raws": rep_raws, "rep_pred_cells": rep_pred_cells,
                "rep_eds": rep_eds, "mean_ed": mean_ed,
                "n_failed": sum(1 for r in rep_pred_cells if r is None),
            }

    # ── Per-strategy aggregates ────────────────────────────────────
    print("\n=== Per-strategy aggregates ===")
    THRESHOLDS = [(0, "SDR@0"), (1, "SDR@1"), (math.sqrt(2), "SDR@√2"),
                  (2, "SDR@2"), (4, "SDR@4")]
    summary_per_strat = {}
    for strat in ("zero_shot", "guided", "guided_patient_left"):
        eds = [r["mean_ed"] for r in records.values()
               if r["strategy"] == strat and r["mean_ed"] is not None]
        # Correct-side rate (image right = cols ≥ 9 for tooth 33 GT)
        # Compute by averaging the predicted column across 3 reps
        correct_side_count = 0
        n_with_pred = 0
        for r in records.values():
            if r["strategy"] != strat:
                continue
            cols = [c[0] for c in r["rep_pred_cells"] if c is not None]
            if not cols:
                continue
            n_with_pred += 1
            mean_col = sum(cols) / len(cols)
            if mean_col >= 9:
                correct_side_count += 1
        summary_per_strat[strat] = {
            "n_with_metric": len(eds),
            "mean_ed": sum(eds) / len(eds) if eds else None,
            "median_ed": statistics.median(eds) if eds else None,
            "correct_side_rate": correct_side_count / n_with_pred if n_with_pred else None,
            "n_correct_side": correct_side_count,
            "n_with_prediction": n_with_pred,
        }
        for thr, lbl in THRESHOLDS:
            hits = sum(1 for e in eds if e <= thr + 1e-9)
            summary_per_strat[strat][lbl] = hits / len(eds) if eds else None

        d = summary_per_strat[strat]
        print(f"\n  {strat}:")
        print(f"    n with mean_ed: {d['n_with_metric']}")
        print(f"    mean ED: {d['mean_ed']:.3f} cells" if d['mean_ed'] is not None else "    mean ED: -")
        print(f"    median ED: {d['median_ed']:.3f}" if d['median_ed'] is not None else "    median ED: -")
        print(f"    correct-side rate: {d['n_correct_side']}/{d['n_with_prediction']} = "
              f"{(d['correct_side_rate'] or 0)*100:.1f}%")
        for thr, lbl in THRESHOLDS:
            print(f"    {lbl}: {(d.get(lbl) or 0)*100:.1f}%")

    # ── Three-way paired Wilcoxon ──────────────────────────────────
    print("\n=== Pre-specified paired tests ===")
    pair_tests = {}
    for a, b in [("guided_patient_left", "guided"),
                 ("guided_patient_left", "zero_shot"),
                 ("guided", "zero_shot")]:
        deltas = []
        for qid in target_qids:
            ra = records.get((qid, a))
            rb = records.get((qid, b))
            if not ra or not rb: continue
            if ra["mean_ed"] is None or rb["mean_ed"] is None: continue
            deltas.append(ra["mean_ed"] - rb["mean_ed"])  # positive = a worse
        if deltas:
            res = paired_wilcoxon(deltas)
            pair_tests[f"{a}_vs_{b}"] = res
            print(f"  {a} vs {b}: n={res['n_total']}, "
                  f"meanΔ={res['mean_delta']:+.3f}, p={res['p']:.4g}, "
                  f"r={res['rank_biserial_r']:+.3f}")

    # ── Reproducibility for guided_patient_left ──────────────────────────
    print("\n=== Reproducibility (Fleiss κ across 3 reps) ===")
    irc = []
    for qid in target_qids:
        r = records.get((qid, "guided_patient_left"))
        if not r: continue
        cells = [c for c in r["rep_pred_cells"] if c is not None]
        if len(cells) != 3: continue
        cnt = Counter(tuple(c) for c in cells)
        irc.append(cnt)
    kappa = fleiss_kappa(irc)
    n_unanimous = sum(1 for c in irc if max(c.values()) == 3)
    print(f"  n items with 3 valid reps: {len(irc)}")
    print(f"  Fleiss κ: {kappa:.4f}" if kappa is not None else "  Fleiss κ: -")
    print(f"  3-way unanimous: {n_unanimous}/{len(irc)} = "
          f"{n_unanimous/len(irc)*100:.1f}%" if irc else "  3-way unanimous: -")

    # ── Pre-specified decision ─────────────────────────────────────
    print("\n=== PRE-SPECIFIED DECISION ===")
    cs_rate = summary_per_strat["guided_patient_left"]["correct_side_rate"] or 0
    if cs_rate >= 0.90:
        verdict = "CONFIRMED"
        msg = ("explicit patient-frame disambiguation of 'left' fixes the "
               "Tooth_33_Apex guided regression; recommend canonicalising "
               "guided_patient_left (apply the same patient-frame rewording "
               "to Tooth_36_* in the periapical FDI landmarks for consistency)")
    elif cs_rate >= 0.30:
        verdict = "PARTIALLY supported"
        msg = ("patient-frame disambiguation contributes but is not sufficient; "
               "investigate residual sources before any canonical change. "
               "Likely next-step: combine with L–R-clause removal to test "
               "Variant C")
    else:
        verdict = "REJECTED"
        msg = ("patient-frame 'left' disambiguation is NOT the active "
               "ingredient; the Tooth_33_Apex regression is driven by "
               "something else — most plausibly the panoramic L–R inversion "
               "clause itself (Variant C diagnostic) rather than landmark-name "
               "wording")
    print(f"  guided_patient_left correct-side rate = {cs_rate*100:.1f}%")
    print(f"  Verdict: {verdict}")
    print(f"  Implication: {msg}")

    # ── Persist ────────────────────────────────────────────────────
    output = {
        "target_structure": TARGET_STRUCTURE,
        "n_target_queries": len(target_qids),
        "summary_per_strategy": summary_per_strat,
        "paired_wilcoxon": pair_tests,
        "reproducibility": {
            "fleiss_kappa": kappa, "n_items": len(irc),
            "n_unanimous_3way": n_unanimous,
        },
        "verdict": {
            "guided_patient_left_correct_side_rate": cs_rate,
            "verdict": verdict, "implication": msg,
        },
        "n_predictions": len(preds),
        "n_missing_cells": len(missing),
    }
    out_path = sandbox / "ablation_analysis.json"
    out_path.write_text(json.dumps(output, indent=2, default=str))
    print(f"\nWrote {out_path}")

    # Per-image debug table for spot-checking
    debug = []
    for qid in sorted(target_qids):
        row = {"query_id": qid, "consensus_gt": qmap_v2[qid]["consensus_gt"]}
        for strat in ("zero_shot", "guided", "guided_patient_left"):
            r = records.get((qid, strat))
            if r:
                row[f"{strat}_pred_cells"] = [c for c in r["rep_pred_cells"]]
                row[f"{strat}_mean_ed"] = r["mean_ed"]
                row[f"{strat}_raw"] = r["rep_raws"]
        debug.append(row)
    (sandbox / "ablation_per_image.json").write_text(
        json.dumps(debug, indent=2, default=str))


if __name__ == "__main__":
    main()
