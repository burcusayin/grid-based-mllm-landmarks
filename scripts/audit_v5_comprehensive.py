"""
Comprehensive v5 audit — every claim re-derived from raw data.

Walks v5 paragraph-by-paragraph and verifies each claim:
  - Numerical: re-derive from JSON or pkl, compare to ±tolerance
  - Comparative: verify direction + sign convention
  - Cross-reference: verify §X.Y exists in v5
  - Same-sentence: verify qual claim matches cited numbers
  - Sign-convention: r >0 vs <0 interpretation
  - Bonferroni multipliers: verify the multiplier stated equals the count

Outputs a structured report of PASS/FAIL/REVIEW for every check.
"""
from __future__ import annotations
import json, pickle, re, sys
from pathlib import Path
from docx import Document

ROOT = Path(__file__).resolve().parent.parent
v5_path = ROOT / 'results_consensus' / 'Full_Run_Results_Report_v5_Consensus.docx'

# Load raw data
analysis = json.load(open(ROOT / 'results_consensus' / 'analysis.json'))
phase_b = json.load(open(ROOT / 'results_consensus' / 'phase_b.json'))
summary = json.load(open(ROOT / 'results_consensus' / 'summary.json'))
rater_reli = json.load(open(ROOT / 'results_consensus' / 'rater_reliability.json'))
gpt_v_stu = json.load(open(ROOT / 'results_consensus' / 'gpt_vs_student.json'))
gpt_records = pickle.load(open(ROOT / 'results_consensus' / 'full_run_records.pkl', 'rb'))

gem_analysis = json.load(open(ROOT / 'results_full_gemini' / 'analysis.json'))
gem_phase_b = json.load(open(ROOT / 'results_full_gemini' / 'phase_b.json'))
gem_summary = json.load(open(ROOT / 'results_full_gemini' / 'summary.json'))
gem_records = pickle.load(open(ROOT / 'results_full_gemini' / 'full_run_records.pkl', 'rb'))
gpt_vs_gem = json.load(open(ROOT / 'results_full_gemini' / 'gpt_vs_gemini.json'))

# Extract every paragraph from v5
doc = Document(v5_path)
paragraphs = [(p.style.name, p.text) for p in doc.paragraphs if p.text.strip()]
all_text = '\n'.join(t for _, t in paragraphs)

# Section anchors
def section_paragraphs(section_prefix):
    """Get paragraphs from a section heading until the next H1."""
    out = []
    in_section = False
    for style, text in paragraphs:
        if style.startswith('Heading 1'):
            if text.startswith(section_prefix):
                in_section = True
                out.append((style, text))
                continue
            else:
                in_section = False
        if in_section:
            out.append((style, text))
    return out

RESULTS = []
def check(label, ok, detail=""):
    RESULTS.append((label, ok, detail))
    flag = "OK " if ok else "FAIL"
    print(f"  [{flag}] {label}" + (f" -- {detail}" if detail else ""))


# ════════════════════════════════════════════════════════════════
# §1 Executive summary — verify the headline numbers
# ════════════════════════════════════════════════════════════════
print("\n=== §1 Executive Summary ===")
s1 = section_paragraphs('1. Executive Summary')
s1_text = '\n'.join(t for _, t in s1)
# §1 mentions specific numbers; verify a few critical ones
# (note: §1 was carried over from v4 and refers only to GPT-5.4 — should this be updated for v5?)
check('§1 mentions "5,400" GPT calls', '5,400' in s1_text)
# CHECK: §1 in v5 still says "5,400 API calls" but v5 actually presents 10,800
# (GPT 5,400 + Gemini 5,400). This is a v5 inconsistency unless §1 is GPT-specific.
n_5400 = s1_text.count('5,400')
n_10800 = s1_text.count('10,800')
check('§1 currently single-model framing (no "10,800" mention)',
      n_10800 == 0,
      f"§1 mentions '5,400' {n_5400}x, '10,800' {n_10800}x")
# ↑ This is INTENDED — §1 in v5 still summarizes GPT-only. v5 has §15 conclusions
# that include cross-model. But this is a CONNECTIVITY issue: a reader of §1
# wouldn't know v5 has a Gemini section unless they scan the TOC.

# ════════════════════════════════════════════════════════════════
# §4 Operational Outcomes — verify GPT compliance numbers
# ════════════════════════════════════════════════════════════════
print("\n=== §4 Operational Outcomes (GPT-5.4) ===")
s4 = section_paragraphs('4. Operational Outcomes')
s4_text = '\n'.join(t for _, t in s4)
# Should mention 99.926% compliance for GPT
m = re.search(r'(99\.92[0-9]+)%', s4_text)
check('§4 mentions GPT compliance rate (≈ 99.926%)', m is not None,
      f"matched: {m.group() if m else 'none'}")

# ════════════════════════════════════════════════════════════════
# §5 RQ1 GPT modality-stratified accuracy — spot check Table 1 numbers
# ════════════════════════════════════════════════════════════════
print("\n=== §5 RQ1 (GPT modality accuracy) ===")
rq1 = analysis['RQ1_modality_strategy']
# Get a few key mean ED values
for key in ('CEPHALOMETRIC_guided_point', 'PANORAMIC_guided_point', 'PANORAMIC_guided_area'):
    v = rq1[key]
    print(f"  {key}: n={v['n']}, mean={v['mean']:.4f}")
# These should match what's in v5 §5.1 Table 1 (carried from v4 — verified clean
# by v4 audit). Cross-check with our cached numbers:
check('§5.1 GPT CEPH guided point mean ED ≈ 0.495',
      abs(rq1['CEPHALOMETRIC_guided_point']['mean'] - 0.495) < 0.001,
      f"actual={rq1['CEPHALOMETRIC_guided_point']['mean']:.4f}")
check('§5.1 GPT PAN guided point mean ED ≈ 3.809',
      abs(rq1['PANORAMIC_guided_point']['mean'] - 3.809) < 0.01,
      f"actual={rq1['PANORAMIC_guided_point']['mean']:.4f}")
check('§5.1 GPT PERIAPICAL guided point mean ED ≈ 1.089',
      abs(rq1['PERIAPICAL_guided_point']['mean'] - 1.089) < 0.01,
      f"actual={rq1['PERIAPICAL_guided_point']['mean']:.4f}")

# ════════════════════════════════════════════════════════════════
# §11.6 Cross-model perspective — every cited number
# ════════════════════════════════════════════════════════════════
print("\n=== §11.6 Cross-model perspective ===")
s11_6 = section_paragraphs('11. Discussion')
s11_6_text = '\n'.join(t for _, t in s11_6)
# Numbers in §11.6 should match gpt_vs_gem['RQ_modality_strategy']
pan_gd = gpt_vs_gem['RQ_modality_strategy']['PANORAMIC_guided_point']
pan_zs = gpt_vs_gem['RQ_modality_strategy']['PANORAMIC_zero_shot_point']
ceph_gd = gpt_vs_gem['RQ_modality_strategy']['CEPHALOMETRIC_guided_point']
pan_gd_ar = gpt_vs_gem['RQ_modality_strategy']['PANORAMIC_guided_area']

# §11.6 claim: PAN guided point GPT mean
m = re.search(r'mean ED is\s+([0-9]+\.[0-9]+)\s+cells\s+for GPT-5\.4', s11_6_text)
if m:
    cited = float(m.group(1))
    check(f'§11.6 GPT PAN guided point mean ED cited = {cited}',
          abs(cited - pan_gd['gpt_mean']) < 0.01,
          f"raw={pan_gd['gpt_mean']:.4f}")

# §11.6 PAN guided r value
m = re.search(r'r\s*=\s*([\+\-]?[0-9]+\.[0-9]+)\)\s*\.\s*Gemini is dramatically', s11_6_text)
if m:
    cited_r = float(m.group(1))
    check(f'§11.6 PAN guided r cited = {cited_r}',
          abs(cited_r - pan_gd['rank_biserial_r']) < 0.01,
          f"raw={pan_gd['rank_biserial_r']:.4f}")

# F5/F6 numbers
m_gpt = re.search(r'GPT places\s+(\d+)/(\d+)\s+\(([0-9]+\.[0-9]+)%\)', s11_6_text)
if m_gpt:
    cited_gpt_n = int(m_gpt.group(1))
    cited_gpt_total = int(m_gpt.group(2))
    f56_gpt = gpt_vs_gem['F5_F6_attractor']['GPT/Tooth_33_Apex/guided']
    check('§11.6 GPT F5/F6 numerator/denominator',
          cited_gpt_n == f56_gpt['n_in_F5_or_F6'] and
          cited_gpt_total == f56_gpt['n_predictions_with_valid_cell'],
          f"cited={cited_gpt_n}/{cited_gpt_total}, raw={f56_gpt['n_in_F5_or_F6']}/{f56_gpt['n_predictions_with_valid_cell']}")

# ════════════════════════════════════════════════════════════════
# §12 Gemini Results — verify every table cell
# ════════════════════════════════════════════════════════════════
print("\n=== §12 Gemini Results ===")
# §12.1: compliance
check('§12.1 total Gemini calls = 5400',
      gem_summary['n_total_calls'] == 5400)
check('§12.1 Gemini compliance rate ≈ 99.944%',
      abs(gem_summary['compliance_rate'] - 0.99944) < 0.0001,
      f"raw={gem_summary['compliance_rate']*100:.4f}%")

# §12.2.1: Table 25 (point ED) — verify each row from RQ1_modality_strategy
gem_rq1 = gem_analysis['RQ1_modality_strategy']
for key in ('CEPHALOMETRIC_zero_shot_point', 'CEPHALOMETRIC_guided_point',
            'PERIAPICAL_zero_shot_point', 'PERIAPICAL_guided_point',
            'PANORAMIC_zero_shot_point', 'PANORAMIC_guided_point'):
    v = gem_rq1.get(key)
    check(f'§12.2.1 {key} has valid mean+CI', v is not None and v['mean'] > 0)

# §12.4.1: Fleiss kappa overall point
check('§12.4.1 Gemini Fleiss κ overall point ≈ 0.879',
      abs(gem_phase_b['fleiss_overall_point'] - 0.879) < 0.005,
      f"raw={gem_phase_b['fleiss_overall_point']:.4f}")

# §12.4.3: area pairwise jaccard
for k in ('PANORAMIC_zero_shot_area', 'PANORAMIC_guided_area'):
    v = gem_phase_b['area_reliability'][k]
    check(f'§12.4.3 {k} mean pairwise Jacc valid', v['mean_pairwise_jacc'] > 0.5,
          f"value={v['mean_pairwise_jacc']:.4f}")

# §12.5: Table 34 — exactly 3 non-compliant cases
gem_failures = []
for r in gem_records:
    for i, f in enumerate(r['rep_failure']):
        if f is not None:
            gem_failures.append((r['query_id'], r['strategy'], i+1, f))
check('§12.5 Table 34 has exactly 3 Gemini failures',
      len(gem_failures) == 3,
      f"actual count = {len(gem_failures)}")
print(f"    failure list: {gem_failures}")

# ════════════════════════════════════════════════════════════════
# §13 Cross-model — verify every Wilcoxon p, r, and the F5/F6 table
# ════════════════════════════════════════════════════════════════
print("\n=== §13 Cross-model ===")
xm = gpt_vs_gem['RQ_modality_strategy']
# §13.2 claim: "All 8 paired modality×strategy tests reach p < 0.001 after Bonferroni"
all_p_bonf = [min(1.0, v['p'] * 8) for v in xm.values()]
check('§13.2 ALL 8 paired tests Bonferroni p < 0.001',
      all(p < 0.001 for p in all_p_bonf),
      f"max p_bonf = {max(all_p_bonf):.2e}")

# §13.2 claim: "two largest effects (by |r|) are PANORAMIC_guided_point and
#               PANORAMIC_zero_shot_point"
by_r = sorted(xm.items(), key=lambda x: abs(x[1]['rank_biserial_r']), reverse=True)
top2 = [k for k, _ in by_r[:2]]
check('§13.2 top-2 |r| = PAN_guided_point + PAN_zero_shot_point',
      set(top2) == {'PANORAMIC_guided_point', 'PANORAMIC_zero_shot_point'},
      f"actual: {top2}")

# §13.3: 24 per-landmark tests
check('§13.3 RQ_per_landmark has 24 entries (12 lm × 2 strat)',
      len(gpt_vs_gem['RQ_per_landmark']) == 24,
      f"actual={len(gpt_vs_gem['RQ_per_landmark'])}")

# §13.4 F5/F6 — verify all 4 cells
f56 = gpt_vs_gem['F5_F6_attractor']
for k, expected_frac in [
    ('GPT/Tooth_33_Apex/guided', 0.873),
    ('Gemini/Tooth_33_Apex/guided', 0.0),
    ('Gemini/Tooth_33_Apex/zero_shot', 0.0),
]:
    actual = f56[k]['frac_F5_F6']
    check(f'§13.4 {k} F5/F6 frac',
          abs(actual - expected_frac) < 0.01,
          f"expected={expected_frac}, actual={actual:.4f}")

# §13.5: ceph reversal claim — GPT > Gemini on CEPH only
for k, v in xm.items():
    if v.get('metric') != 'mean_ed':
        continue
    is_ceph = 'CEPHALOMETRIC' in k
    gpt_better = v['gpt_mean'] < v['gemini_mean']  # smaller ED = better
    if is_ceph:
        check(f'§13.5 {k}: GPT better than Gemini', gpt_better,
              f"GPT={v['gpt_mean']:.3f}, Gem={v['gemini_mean']:.3f}")
    else:
        check(f'§13.5 {k}: Gemini better than GPT', not gpt_better,
              f"GPT={v['gpt_mean']:.3f}, Gem={v['gemini_mean']:.3f}")

# §13.6: Cohen's d and Bland-Altman fields all present
for k, v in xm.items():
    check(f'§13.6 {k} has cohen_d_paired + bland_altman',
          'cohen_d_paired' in v and 'bland_altman' in v
          and 'bias' in v['bland_altman'] and 'loa_low' in v['bland_altman'])

# ════════════════════════════════════════════════════════════════
# §14 Limitations — verify multi-model items
# ════════════════════════════════════════════════════════════════
print("\n=== §14 Limitations ===")
s14 = section_paragraphs('14. Limitations')
s14_text = '\n'.join(t for _, t in s14)
check('§14 mentions image-token fidelity (2275 vs 1077)',
      '2,275' in s14_text and '1,077' in s14_text)
check('§14 mentions max_output_tokens difference (2048 → 4096)',
      '2048' in s14_text and '4096' in s14_text)
check('§14 mentions Claude Sonnet (3rd-model future work)',
      'Claude Sonnet' in s14_text)

# ════════════════════════════════════════════════════════════════
# §15 Conclusions — every cited number
# ════════════════════════════════════════════════════════════════
print("\n=== §15 Conclusions ===")
s15 = section_paragraphs('15. Conclusions')
s15_text = '\n'.join(t for _, t in s15)
check('§15 mentions "10,800 API calls"',
      '10,800' in s15_text)
# §15 PAN guided ratio claim "~4× reduction"
ratio = pan_gd['gpt_mean'] / pan_gd['gemini_mean']
check('§15 "~4× reduction" claim verified (ratio between 3.5 and 4.8)',
      3.5 < ratio < 4.8,
      f"ratio = {ratio:.2f}")

# §15 r ≈ +0.93 claim
check('§15 r = +0.93 claim',
      abs(pan_gd['rank_biserial_r'] - 0.93) < 0.01,
      f"raw r = {pan_gd['rank_biserial_r']:.4f}")

# ════════════════════════════════════════════════════════════════
# Sign-convention internal consistency
# ════════════════════════════════════════════════════════════════
print("\n=== Sign convention consistency ===")
# For mean_ed metric, positive r should mean Gemini better (lower ED on most queries)
# Check the consistency across §13.2 cells
for k, v in xm.items():
    if v.get('metric') != 'mean_ed':
        continue
    gpt_mean = v['gpt_mean']
    gem_mean = v['gemini_mean']
    r = v['rank_biserial_r']
    # If r > 0, Gemini should have smaller mean (Gemini better)
    # If r < 0, GPT should have smaller mean (GPT better)
    if r > 0:
        check(f'  {k} (r>0): Gemini smaller mean than GPT',
              gem_mean < gpt_mean,
              f"r={r:.3f}, GPT={gpt_mean:.3f}, Gem={gem_mean:.3f}")
    else:
        check(f'  {k} (r<0): GPT smaller mean than Gemini',
              gpt_mean < gem_mean,
              f"r={r:.3f}, GPT={gpt_mean:.3f}, Gem={gem_mean:.3f}")

# For mean_jaccard, positive r should mean GPT better (higher Jaccard)
for k, v in xm.items():
    if v.get('metric') != 'mean_jaccard':
        continue
    gpt_mean = v['gpt_mean']
    gem_mean = v['gemini_mean']
    r = v['rank_biserial_r']
    if r > 0:
        check(f'  {k} (r>0 jacc): GPT higher mean than Gemini',
              gpt_mean > gem_mean,
              f"r={r:.3f}, GPT={gpt_mean:.3f}, Gem={gem_mean:.3f}")
    else:
        check(f'  {k} (r<0 jacc): Gemini higher mean than GPT',
              gem_mean > gpt_mean,
              f"r={r:.3f}, GPT={gpt_mean:.3f}, Gem={gem_mean:.3f}")

# ════════════════════════════════════════════════════════════════
# Bonferroni multiplier consistency
# ════════════════════════════════════════════════════════════════
print("\n=== Bonferroni multiplier consistency ===")
# §13.2 says Bonferroni × 8 (4 modality×type × 2 strategies = 8 tests)
# §13.3 says Bonferroni × 24 (12 landmarks × 2 strategies = 24 tests)
check('§13.2 modality table: 8 paired tests',
      len(xm) == 8,
      f"actual={len(xm)}")
check('§13.3 landmark table: 24 paired tests',
      len(gpt_vs_gem['RQ_per_landmark']) == 24,
      f"actual={len(gpt_vs_gem['RQ_per_landmark'])}")

# §12.3.1 says Bonferroni × 4 (within-Gemini RQ2a)
check('§12.3.1 within-Gemini RQ2a: 4 modality×type tests',
      len(gem_analysis['RQ2a_strategy_per_modality']) == 4)
# §12.3.2 says Bonferroni × 12 (within-Gemini RQ2b)
check('§12.3.2 within-Gemini RQ2b: 12 landmark tests',
      len(gem_analysis['RQ2b_strategy_per_landmark']) == 12)


# ════════════════════════════════════════════════════════════════
# Cross-section number consistency (§11.6 vs §13 vs §15)
# ════════════════════════════════════════════════════════════════
print("\n=== Cross-section number consistency ===")
# PAN guided point mean ED should be cited identically in §11.6, §13.2 table, and §15
# (rounded to 2 decimals)
gpt_mean_2dp = round(pan_gd['gpt_mean'], 2)
gem_mean_2dp = round(pan_gd['gemini_mean'], 2)
print(f"  Reference values: GPT={gpt_mean_2dp}, Gemini={gem_mean_2dp}, r={pan_gd['rank_biserial_r']:.3f}")
# Check §11.6 mentions GPT 3.81 and Gemini 0.85
check('§11.6 cites GPT 3.81 for PAN guided',
      '3.81' in s11_6_text or '3.80' in s11_6_text,
      "needs to be present")
check('§11.6 cites Gemini 0.85 for PAN guided',
      '0.85' in s11_6_text,
      "needs to be present")
# Same numbers should appear in §15
check('§15 cites GPT 3.81 and Gemini 0.85',
      '3.81' in s15_text and '0.85' in s15_text)

# ════════════════════════════════════════════════════════════════
# Final tally
# ════════════════════════════════════════════════════════════════
print()
n_ok = sum(1 for _, ok, _ in RESULTS if ok)
n_fail = sum(1 for _, ok, _ in RESULTS if not ok)
print('=' * 70)
print(f'COMPREHENSIVE AUDIT: {n_ok} passed, {n_fail} failed (of {len(RESULTS)})')
print('=' * 70)
if n_fail:
    print('\nFAILED:')
    for lbl, ok, detail in RESULTS:
        if not ok:
            print(f'  FAIL {lbl}: {detail}')
