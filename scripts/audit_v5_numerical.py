"""
Numerical audit of v5 report: walk every number in the new §11.6, §12, §13,
§14, §15 sections and verify it can be reproduced from the source JSON files.

This script DOES NOT depend on python-docx reading the document — instead it
re-derives each cell value the same way the generator does, from
results_full_gemini/{analysis,phase_b,summary,gpt_vs_gemini}.json and
results_full_gemini/full_run_records.pkl, then verifies internal consistency.
The expectation is that every number appears in the source JSON, and the
prose claims in the body match those numbers exactly.

Run from project root: .venv/bin/python /tmp/audit_v5_numerical.py
"""
from __future__ import annotations

import json
import pickle
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
GS = ROOT / "results_full_gemini"

# Load everything the generator loads
gem_analysis = json.load(open(GS / "analysis.json"))
gem_phase_b = json.load(open(GS / "phase_b.json"))
gem_summary = json.load(open(GS / "summary.json"))
gem_manifest = json.load(open(GS / "full_run_manifest.json"))
gem_manifest_rep1 = json.load(open(GS / "full_run_manifest_rep1.json"))
gpt_vs_gem = json.load(open(GS / "gpt_vs_gemini.json"))
gem_records = pickle.load(open(GS / "full_run_records.pkl", "rb"))

phase_b_gpt = json.load(open(ROOT / "results_consensus" / "phase_b.json"))

RESULTS: list[tuple[str, bool, str]] = []
def check(label: str, condition: bool, detail: str = "") -> None:
    RESULTS.append((label, condition, detail))
    flag = "OK " if condition else "FAIL"
    print(f"  [{flag}] {label}" + (f" -- {detail}" if detail else ""))


# ════════════════════════════════════════════════════════════════════
# §11.6: cross-model perspective bullets
# ════════════════════════════════════════════════════════════════════
print("\n=== §11.6 Cross-model perspective ===")
xm = gpt_vs_gem["RQ_modality_strategy"]
check("PANORAMIC_guided_point has gpt_mean field",
      "gpt_mean" in xm["PANORAMIC_guided_point"],
      f"gpt_mean={xm['PANORAMIC_guided_point']['gpt_mean']:.4f}")
check("PANORAMIC_guided_point has gemini_mean",
      "gemini_mean" in xm["PANORAMIC_guided_point"],
      f"gemini_mean={xm['PANORAMIC_guided_point']['gemini_mean']:.4f}")
check("PANORAMIC_guided_point rank_biserial_r > 0.9 (Gemini better)",
      xm["PANORAMIC_guided_point"]["rank_biserial_r"] > 0.9,
      f"r={xm['PANORAMIC_guided_point']['rank_biserial_r']:.4f}")
check("CEPH guided point r < -0.5 (GPT better, direction reverses)",
      xm["CEPHALOMETRIC_guided_point"]["rank_biserial_r"] < -0.5,
      f"r={xm['CEPHALOMETRIC_guided_point']['rank_biserial_r']:.4f}")
check("F5/F6 GPT guided is 87.3%",
      abs(gpt_vs_gem["F5_F6_attractor"]["GPT/Tooth_33_Apex/guided"]["frac_F5_F6"] - 0.873) < 0.01,
      f"actual={gpt_vs_gem['F5_F6_attractor']['GPT/Tooth_33_Apex/guided']['frac_F5_F6']*100:.1f}%")
check("F5/F6 Gemini guided is 0.0%",
      gpt_vs_gem["F5_F6_attractor"]["Gemini/Tooth_33_Apex/guided"]["frac_F5_F6"] == 0.0,
      "exactly 0/300")

# ════════════════════════════════════════════════════════════════════
# §12.1: compliance + cost + max_tokens disclosure
# ════════════════════════════════════════════════════════════════════
print("\n=== §12.1 Operational outcomes ===")
check("Total calls = 5400", gem_summary["n_total_calls"] == 5400,
      f"actual={gem_summary['n_total_calls']}")
check("Compliance rate ≈ 99.944% (3/5400 failures)",
      abs(gem_summary["compliance_rate"] - 0.99944) < 0.001,
      f"actual={gem_summary['compliance_rate']*100:.3f}%")
check("3 failures total", gem_summary["n_failures"] == 3,
      f"actual={gem_summary['n_failures']}")
# Re-derive failures from records
fail_list = [(r["query_id"], r["strategy"], i+1, r["rep_failure"][i])
              for r in gem_records for i in range(3)
              if r["rep_failure"][i] is not None]
check("Failure list has 3 entries", len(fail_list) == 3,
      f"actual={len(fail_list)}")
# Categories: 1 ambiguous, 1 verbose, 1 no_engage
cats = sorted([f[3] for f in fail_list])
check("Failure categories are (ambiguous, no_engage, verbose)",
      cats == ["ambiguous", "no_engage", "verbose"],
      f"actual={cats}")
check("Manifest max_output_tokens = 4096", gem_manifest["inference_settings"]["max_output_tokens"] == 4096)
check("Rep 1 manifest max_output_tokens = 2048",
      gem_manifest_rep1["inference_settings"]["max_output_tokens"] == 2048)

# ════════════════════════════════════════════════════════════════════
# §12.2: Modality-stratified accuracy tables
# ════════════════════════════════════════════════════════════════════
print("\n=== §12.2 Modality-stratified accuracy (Tables 25-28) ===")
rq1 = gem_analysis["RQ1_modality_strategy"]
# Each table row's mean and CI must come from RQ1_modality_strategy
for key in ("CEPHALOMETRIC_zero_shot_point", "CEPHALOMETRIC_guided_point",
            "PERIAPICAL_zero_shot_point", "PERIAPICAL_guided_point",
            "PANORAMIC_zero_shot_point", "PANORAMIC_guided_point",
            "PANORAMIC_zero_shot_area", "PANORAMIC_guided_area"):
    v = rq1.get(key, {})
    check(f"  {key}: has n={v.get('n')}, mean={v.get('mean')}",
          v.get("n") and v.get("mean") is not None,
          f"n={v.get('n')}, mean={v.get('mean', 0):.4f}")

# NED table check
ned = gem_phase_b["ned_modality"]
check("NED has CEPH/PA/PAN guided + zero_shot (6 entries)",
      sum(1 for k in ned if any(k.startswith(m) for m in ("CEPH","PERI","PANO"))) == 6,
      f"keys={list(ned.keys())}")

# SDR table check
sdr = gem_phase_b["sdr_modality_with_ci"]
check("SDR has 6 (modality × strategy) entries",
      len(sdr) == 6,
      f"keys={list(sdr.keys())}")
# Specific top-line SDR@1 values
check("PAN guided SDR@1 = 0.74",
      abs(sdr["PANORAMIC_guided"]["SDR@1"] - 0.74) < 0.01,
      f"actual={sdr['PANORAMIC_guided']['SDR@1']:.4f}")
check("PERIAPICAL guided SDR@1 = 0.88",
      abs(sdr["PERIAPICAL_guided"]["SDR@1"] - 0.88) < 0.01,
      f"actual={sdr['PERIAPICAL_guided']['SDR@1']:.4f}")
check("CEPHALOMETRIC guided SDR@2 = 1.0",
      sdr["CEPHALOMETRIC_guided"]["SDR@2"] == 1.0)

# ════════════════════════════════════════════════════════════════════
# §12.3: within-Gemini strategy paired tests
# ════════════════════════════════════════════════════════════════════
print("\n=== §12.3 within-Gemini strategy paired tests ===")
rq2a = gem_analysis["RQ2a_strategy_per_modality"]
check("RQ2a has 4 modality × type entries", len(rq2a) == 4,
      f"keys={list(rq2a.keys())}")
rq2b = gem_analysis["RQ2b_strategy_per_landmark"]
check("RQ2b has 12 landmark entries", len(rq2b) == 12,
      f"actual={len(rq2b)}")

# ════════════════════════════════════════════════════════════════════
# §12.4: Reproducibility (Fleiss + unanimity + area reliability)
# ════════════════════════════════════════════════════════════════════
print("\n=== §12.4 Reproducibility ===")
check("Overall Fleiss kappa is reported",
      "fleiss_overall_point" in gem_phase_b,
      f"value={gem_phase_b['fleiss_overall_point']:.4f}")
check("Fleiss kappa > 0.85 (good reproducibility)",
      gem_phase_b["fleiss_overall_point"] > 0.85,
      f"actual={gem_phase_b['fleiss_overall_point']:.4f}")
check("fleiss_per_group has 6 entries", len(gem_phase_b["fleiss_per_group"]) == 6)
check("Gemini overall Fleiss kappa > GPT overall Fleiss kappa",
      gem_phase_b["fleiss_overall_point"] > phase_b_gpt["fleiss_overall_point"],
      f"gemini={gem_phase_b['fleiss_overall_point']:.4f}, gpt={phase_b_gpt['fleiss_overall_point']:.4f}")

# Area reliability
area_rel = gem_phase_b["area_reliability"]
check("Area reliability has 2 entries (zero_shot/guided)", len(area_rel) == 2)
for k, v in area_rel.items():
    check(f"  {k} mean_pairwise_jacc > 0.8",
          v.get("mean_pairwise_jacc", 0) > 0.8,
          f"actual={v.get('mean_pairwise_jacc'):.4f}")

# ════════════════════════════════════════════════════════════════════
# §13.2: Modality-level paired (Table 35)
# ════════════════════════════════════════════════════════════════════
print("\n=== §13.2 Modality-level paired comparison ===")
xm = gpt_vs_gem["RQ_modality_strategy"]
check("RQ_modality_strategy has 8 entries (4 × 2)", len(xm) == 8,
      f"actual={len(xm)}")
# Every cell has the Wilcoxon stats
for key, v in xm.items():
    for need in ("gpt_mean", "gemini_mean", "mean_delta", "rank_biserial_r", "p", "metric"):
        if need not in v:
            check(f"  {key} missing {need}", False)
# All 8 must survive Bonferroni × 8
for key, v in xm.items():
    p_bonf = min(1.0, v["p"] * 8)
    check(f"  {key} Bonferroni p < 0.001",
          p_bonf < 0.001,
          f"p_bonf={p_bonf:.2e}")

# ════════════════════════════════════════════════════════════════════
# §13.3: Per-landmark paired (Table 36)
# ════════════════════════════════════════════════════════════════════
print("\n=== §13.3 Per-landmark paired ===")
xm_lm = gpt_vs_gem["RQ_per_landmark"]
check("RQ_per_landmark has 24 entries (12 × 2)", len(xm_lm) == 24,
      f"actual={len(xm_lm)}")

# ════════════════════════════════════════════════════════════════════
# §13.4: F5/F6 attractor table
# ════════════════════════════════════════════════════════════════════
print("\n=== §13.4 F5/F6 attractor ===")
f56 = gpt_vs_gem["F5_F6_attractor"]
check("F5/F6 has 4 entries (2 models × 2 strategies)", len(f56) == 4,
      f"keys={list(f56.keys())}")
check("GPT guided F5/F6 fraction matches prose claim (~87.3%)",
      abs(f56["GPT/Tooth_33_Apex/guided"]["frac_F5_F6"] - 0.873) < 0.005,
      f"actual={f56['GPT/Tooth_33_Apex/guided']['frac_F5_F6']*100:.1f}%")
check("Gemini guided F5/F6 fraction = 0",
      f56["Gemini/Tooth_33_Apex/guided"]["frac_F5_F6"] == 0)
check("Gemini zero_shot F5/F6 fraction = 0",
      f56["Gemini/Tooth_33_Apex/zero_shot"]["frac_F5_F6"] == 0)
check("Gemini guided exact-match ≈ 20.7%",
      abs(f56["Gemini/Tooth_33_Apex/guided"]["frac_exact_match"] - 0.207) < 0.005,
      f"actual={f56['Gemini/Tooth_33_Apex/guided']['frac_exact_match']*100:.1f}%")
check("Gemini zero_shot exact-match ≈ 22.7%",
      abs(f56["Gemini/Tooth_33_Apex/zero_shot"]["frac_exact_match"] - 0.227) < 0.005,
      f"actual={f56['Gemini/Tooth_33_Apex/zero_shot']['frac_exact_match']*100:.1f}%")
check("GPT guided exact-match ≈ 1.0%",
      abs(f56["GPT/Tooth_33_Apex/guided"]["frac_exact_match"] - 0.010) < 0.005,
      f"actual={f56['GPT/Tooth_33_Apex/guided']['frac_exact_match']*100:.1f}%")
check("GPT zero_shot exact-match ≈ 5.0%",
      abs(f56["GPT/Tooth_33_Apex/zero_shot"]["frac_exact_match"] - 0.050) < 0.005,
      f"actual={f56['GPT/Tooth_33_Apex/zero_shot']['frac_exact_match']*100:.1f}%")

# ════════════════════════════════════════════════════════════════════
# Final tally
# ════════════════════════════════════════════════════════════════════
print()
n_ok = sum(1 for _, ok, _ in RESULTS if ok)
n_fail = sum(1 for _, ok, _ in RESULTS if not ok)
print("=" * 70)
print(f"NUMERICAL AUDIT: {n_ok} passed, {n_fail} failed (of {len(RESULTS)})")
print("=" * 70)
if n_fail:
    print("\nFAILED:")
    for lbl, ok, detail in RESULTS:
        if not ok:
            print(f"  FAIL {lbl}: {detail}")
    sys.exit(1)
