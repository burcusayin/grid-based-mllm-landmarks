# Companion code for the paper titled "*Comparative Analysis of Multimodal Large Language Models and Dental Students in Radiographic Anatomic Landmark Identification: A Novel Grid-Based Assessment Model*"

This repository contains the complete code, statistical-analysis scripts, prompt templates and reproducibility infrastructure used in the manuscript (currently under review). It is intended both as supplementary material for reviewers and as a reusable benchmark framework for future MLLM spatial-reasoning studies in dental radiography.

> **Citation (placeholder — will be replaced on acceptance):**
> Aydemir M.E. et al. *Comparative Analysis of Multimodal Large Language Models and Dental Students in Radiographic Anatomic Landmark Identification: A Novel Grid-Based Assessment Model.* (under review), 2026.

## What this code does

A 900-query benchmark on 200 dental radiographs (100 panoramic, 50 periapical, 50 cephalometric) and 12 landmarks (9 points, 3 areas), evaluating two frontier MLLMs (GPT-5.4 and Gemini 3.1 Pro) under two prompting strategies (zero-shot, guided) across three independent repetitions — **10,800 API calls in total** — against a two-rater adjudicated specialist consensus ground truth and a team-adjudicated fourth-year dental student consensus (n = 40).

Key methodological contributions implemented here:

1. **Grid-based coordinate-response framework.** Each radiograph is overlaid with a per-modality grid (panoramic 16×8; periapical 8×6; cephalometric 10×8) rendered programmatically with cyan grid lines and yellow alphanumeric labels. Models and human raters return cell coordinates (e.g., `G10`). This collapses the localisation problem into a discrete classification task with built-in spatial discretisation, enabling apples-to-apples comparison across modalities and across MLLM/human respondents.
2. **Provider-max image fidelity, byte-identical inputs.** GPT-5.4 (`detail = high`, ~2,275 image tokens) and Gemini 3.1 Pro (`mediaResolution = HIGH`, ~1,077 image tokens) receive the **same prompts** (byte-identity verified at orchestration time) on the **same image PNGs** (SHA-256-anchored). Any cross-model difference is attributable to the model itself, not to prompt or input drift.
3. **Cryptographic reproducibility anchoring.** Every input file (200 PNG images, the source Excel, every raw JSONL/chunk output) is SHA-256-anchored. A `reproducibility_manifest.json` is regenerable from the source via `scripts/reproducibility_manifest.py` and verifies that nothing has drifted since the original run.
4. **Pre-registered prompt-level ablations.** Three targeted ablations (FDI tooth-number removal; patient-frame disambiguation; lateralisation-clause removal) test wording-level explanations for a catastrophic Tooth_33_Apex regression; each had pre-specified decision criteria and 300 fresh API calls.
5. **Cross-model paired Wilcoxon framework with Bonferroni correction.** All cross-model and AI-vs-human comparisons use paired non-parametric tests (justified by Shapiro–Wilk non-normality of paired differences, p < 10⁻⁶), with rank-biserial r as effect size, bootstrap 95% CIs, and Bonferroni correction within each comparison family.

## Repository layout

```
.
├── README.md                                  ← this file
├── LICENSE                                    ← MIT (code)
├── requirements.txt                           ← pinned dependencies
├── SETUP_AND_COSTS.md                         ← API-key + cost guide
├── .env.example                               ← template for API keys (no real keys committed)
├── .gitignore                                 ← blocks raw data + secrets + .api_lock
│
├── config.py                                  ← prompts, grid dims, model configs, seeds
├── pipeline.py                                ← batch submit / status / download / parse
├── analysis.py                                ← legacy v1 metrics + statistics helpers
├── consistency_check.py                       ← Fleiss' κ across repetitions
├── generate_sota_report.py                    ← SOTA-comparison reporting
│
├── scripts/
│   ├── reproducibility_manifest.py            ← regenerate full SHA + version manifest
│   ├── snapshot_prompts.py                    ← persist exact prompts before any run
│   ├── export_sanitized_prompts.py            ← regenerate prompts_all900.json (no GT fields)
│   │
│   ├── run_pilot.py                           ← legacy cost-capped GPT-only pilot
│   ├── run_full_run_gemini.py                 ← Gemini 3.1 Pro full-run orchestrator (file-based async batch)
│   ├── run_full_run_claude.py                 ← Claude Sonnet 4.6 orchestrator (wired but inactive in current paper)
│   ├── run_fdi_ablation.py                    ← Ablation A: FDI tooth-number removal
│   ├── run_patient_left_ablation.py           ← Ablation B: patient-frame disambiguation
│   ├── run_no_LR_ablation.py                  ← Ablation C: lateralisation-clause removal
│   │
│   ├── verify_full_run_setup.py               ← preflight gate before any real API call
│   ├── verify_v2_setup.py                     ← preflight: v2 prompt revision
│   ├── verify_full_run_gemini_setup.py        ← preflight: Gemini full run
│   ├── verify_full_run_claude_setup.py        ← preflight: Claude full run
│   ├── verify_fdi_ablation_setup.py           ← preflight: Ablation A
│   ├── verify_patient_left_ablation_setup.py  ← preflight: Ablation B
│   ├── verify_no_LR_ablation_setup.py         ← preflight: Ablation C
│   │
│   ├── recompute_against_consensus.py         ← re-evaluate GPT records vs consensus GT (no API calls)
│   ├── recompute_gemini.py                    ← re-evaluate Gemini records vs consensus GT (no API calls)
│   │
│   ├── analyze_consensus_run.py               ← RQ1/RQ2/RQ3: GPT modality-stratified + strategy + reproducibility
│   ├── analyze_rater_reliability.py           ← RQ4: inter/intra-rater κ (OMFR_1↔OMFR_2)
│   ├── analyze_gpt_vs_student.py              ← RQ5: GPT vs students, paired Wilcoxon + Bland-Altman
│   ├── analyze_gpt_vs_gemini.py               ← RQ7: cross-model paired Wilcoxon + F5/F6 attractor
│   ├── analyze_gemini_vs_student.py           ← RQ8: Gemini vs students, paired Wilcoxon
│   ├── analyze_fdi_ablation.py                ← Ablation A analysis + pre-registered verdict
│   ├── analyze_patient_left_ablation.py       ← Ablation B analysis
│   ├── analyze_no_LR_ablation.py              ← Ablation C analysis
│   │
│   ├── v4_canonical_stats.py                  ← Canonical statistical pipeline (Table 2 / 3 / 4 / 5 from raw, with bootstrap 95 % CIs + Wilcoxon W⁺)
│   ├── compute_S2_cis.py                      ← Bootstrap 95 % CIs for inter/intra-rater Cohen's κ and mean Jaccard / Dice
│   ├── regen_fig3_v4.py                       ← Figure 3 (SDR@1 bar chart by modality)
│   ├── regen_fig4_v4.py                       ← Figure 4 (Tooth_33_Apex prediction heatmap, GPT vs Gemini)
│   └── regen_fig5_v4.py                       ← Figure 5 (GPT-5.4 per-landmark zero-shot vs guided rank-biserial r)
│
└── tests/
    └── smoke_e2e.py                           ← regression gate (parse → metrics → plots)
```

## Privacy: what is **not** in this repository

The benchmark dataset itself (radiograph PNGs, expert annotations, student responses) is institutional clinical data and is **not** redistributed via this repository. Specifically, the following directories are `.gitignore`d:

* `data/` — the source Excel (`Final_Dental_MLLM_Benchmark_Data.xlsx`) and all 200 anonymised PNG images.
* `docs/` — the manuscript drafts, cover letter, title page, and colleague correspondence.
* `results_full/`, `results_full_gemini/`, `results_consensus/`, `results_ablation_*/` — the raw API responses (≥ 1.4 GB of JSONL chunks per repetition) and all derived analysis JSONs.
* `.env`, `.api_lock`, any `*.key` / `*.pem` — secrets and the safety lock.

**Reviewers and replication seekers**:  Image dataset use is governed by the institutional ethics committee (Decision No: GO 2026/3048; Burdur Mehmet Akif Ersoy University). The code in this repository will operate on any compatible Excel + image set if the schema documented in `config.py` is followed.

## Reproducing the results (with your own data)

### 1. Install

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Python 3.11+ is required (the original runs used 3.14). Dependencies: `openpyxl`, `scipy`, `numpy`, `matplotlib`, `python-docx`. No GPU is needed; all model inference is via cloud Batch APIs.

### 2. Configure data

Place your radiographs and Excel under `data/`:

```
data/
  Final_Dental_MLLM_Benchmark_Data.xlsx
  images/
    Panoramic_Grid/    PAN_001.png ... PAN_100.png
    Periapical_Grid/   PA_001.png ... PA_050.png
    Cephalometric_Grid/CEPH_001.png ... CEPH_050.png
```

The Excel must contain three sheets (`PANORAMIC`, `PERIAPICAL`, `CEPHALOMETRIC`) with columns including `Image_ID`, `Category`, `Target_Structure`, `Question_Prompt`, `OMFR_1`, `OMFR_2`, `CONSENSUS_Ground_Truth`, and `Student_*` annotations. The exact schema is validated by `pipeline.py prepare_v2` and documented in `config.py`.

### 3. API keys

```bash
cp .env.example .env       # then paste your OPENAI_API_KEY, GOOGLE_API_KEY into .env
```

`.env` is gitignored. The pipeline reads keys from `os.environ`; `.env` is loaded automatically by all orchestrators (`scripts/run_full_run_gemini.py`, `scripts/run_full_run_claude.py`, `pipeline.py`).

See [`SETUP_AND_COSTS.md`](SETUP_AND_COSTS.md) for full cost guidance (≈ US$18 GPT-5.4 + US$55 Gemini 3.1 Pro for 3 repetitions each + the 3 GPT-only ablations ≈ US$5; ≈ US$80 total).

### 4. Run an end-to-end full run (Gemini example)

The orchestrators are idempotent, preflight-gated, cost-confirmation-prompted, and re-create `.api_lock` in a `finally:` block on every exit (success or failure).

```bash
# Stage 1 only — no API spend, prepares sandbox and computes SHA anchors
.venv/bin/python scripts/run_full_run_gemini.py --only-prepare

# Full run: remove .api_lock manually first, then launch
rm .api_lock
.venv/bin/python scripts/run_full_run_gemini.py --repetitions 3
# Type 'yes' at the cost-confirmation prompt to authorise spending.
```

The orchestrator's 7-stage pipeline:

| Stage | Action |
|------:|---|
| 0 | Preflight (auth, lock, SHA, prompt-byte-identity check vs GPT-5.4 baseline) |
| 1 | Prepare sandbox (copy query_index, compute SHA anchors over all 200 images + raw JSONLs) |
| 2 | Snapshot exact prompts — verifies byte-identity vs the GPT-5.4 baseline before any spend |
| 3 | Live single-call canary (≈ US$0.005) — refuses to proceed if parse fails |
| 4 | File-based async batch submit (uploads JSONL to Files API, creates batches) |
| 5 | Wait — poll Google batch state every 120 s until terminal |
| 6 | Download — fetch each batch's result JSONL, convert to chunk JSON for parser |
| 7 | Parse + per-rep compliance summary |

Total wall time for a 3-rep Gemini run: ≈ 1–4 h (Google's batch SLA is "within 24 h" with no fast-path).

### 5. Analysis (no API calls)

```bash
# Re-derive per-query records (ED / Jaccard / Dice vs every reference)
.venv/bin/python scripts/recompute_against_consensus.py    # GPT
.venv/bin/python scripts/recompute_gemini.py               # Gemini

# Statistical analyses (RQ1–RQ8)
.venv/bin/python scripts/analyze_consensus_run.py          # RQ1/2/3 GPT
.venv/bin/python scripts/analyze_rater_reliability.py      # RQ4 OMFR_1↔OMFR_2
.venv/bin/python scripts/analyze_gpt_vs_student.py         # RQ5
.venv/bin/python scripts/analyze_gpt_vs_gemini.py          # RQ7 cross-model
.venv/bin/python scripts/analyze_gemini_vs_student.py      # RQ8

# Canonical recomputation (Tables 2/3/4/5 + bootstrap CIs + Wilcoxon W⁺)
.venv/bin/python scripts/v4_canonical_stats.py
.venv/bin/python scripts/compute_S2_cis.py
```

All analyses are deterministic: bootstrap CIs use `random.Random(42)` so two clean runs produce byte-identical JSONs.

### 6. Regenerate the published figures and statistical tables

```bash
# All four scripts are deterministic (seed = 42 wired in)
.venv/bin/python scripts/v4_canonical_stats.py     # → results_v4_canonical.json
.venv/bin/python scripts/compute_S2_cis.py         # → results_v4_S2_cis.json
.venv/bin/python scripts/regen_fig3_v4.py          # → fig3.png (SDR@1 bar chart)
.venv/bin/python scripts/regen_fig4_v4.py          # → fig4.png (Tooth_33_Apex heatmap)
.venv/bin/python scripts/regen_fig5_v4.py          # → fig5.png (GPT per-landmark r)
```

`v4_canonical_stats.py` is the single canonical source for every numerical
claim in the manuscript's Tables 2 (mean ED + SDR), 3 (GPT zero-shot vs
guided per-landmark), 4 (model vs student) and 5 (cross-model). It loads
raw `parsed_responses.json` from each repetition and the consensus pickle,
recomputes every per-query metric, runs `scipy.stats.wilcoxon`, derives
the rank-biserial r from W⁺/W⁻, and bootstraps 95 % CIs (10,000 resamples
for means and Δ, 2,000 for r).

### 7. Smoke / regression gate

```bash
.venv/bin/python tests/smoke_e2e.py
```

Exercises the full pipeline on a sandbox (`tests/.smoke_sandbox/`) and verifies that real `results_*/` directories are untouched.

## Reproducibility manifest

```bash
.venv/bin/python scripts/reproducibility_manifest.py
```

Captures: current git SHA, Python version, `pip freeze`, SHA-256 of the Excel + every image + every source file, plus the rendered prompt templates and config parameters. Output is deterministic (no timestamps) so two clean runs on the same state produce byte-identical manifests.

`config.RANDOM_SEED = 42` is wired into OpenAI and Gemini batch requests. Anthropic Batch API does not accept a seed and is intentionally omitted.

## Safety: cost gates

Every API-spending entry point checks for a sentinel file `.api_lock` in the repo root; if present, the orchestrator exits with `ERROR: API calls are locked. Delete .api_lock to unlock.` and spends $0. Orchestrators re-create `.api_lock` in their `finally:` block on every exit (success or failure), so the default state after every run is "locked".

Before any real-API operation:

1. Always run with `--only-prepare` (or `--dry-run`) first to verify Stage 0 + Stage 1.
2. Run the preflight (`scripts/verify_full_run_*.py`) — it refuses if `.api_lock` is present.
3. Read the cost projection printed at the prompt before typing `yes`.

## Ethics & data governance

Ethical approval was granted by the Non-Interventional Clinical Research Ethics Committee of Burdur Mehmet Akif Ersoy University (Decision No: GO 2026/3048; Meeting Date: 13 May 2026). All radiographs were retrieved retrospectively from institutional archives with all personal identifiers removed, in full accordance with the Declaration of Helsinki. The 40-student cohort participated as the human reference group under the standard institutional informed-consent process.

## License

MIT — see [LICENSE](LICENSE). Code is reusable for academic and commercial purposes. The dataset itself is not redistributed and is not covered by this license.

## Contact

Burcu Sayın — burcusayin.w@gmail.com for inquiries about the code.
