# Grid-Based MLLM Landmark Benchmark

Benchmark comparing multimodal LLMs (GPT-5.4, Gemini 3.1 Pro) against dental students on anatomic landmark identification in dental radiographs, using a coarse grid-cell methodology across panoramic, periapical, and cephalometric modalities.

## Overview

- **Task**: identify anatomic landmarks (points and areas) on dental radiographs by naming the grid cell(s) containing them.
- **Modalities**: panoramic (8×16 grid), periapical (6×8), cephalometric (8×10).
- **Models**: GPT-5.4 and Gemini 3.1 Pro via Batch APIs (Anthropic Claude wired but inactive by default).
- **Prompt strategies**: `zero_shot` and `guided` (anatomic-context primer), exactly two.
- **Metrics**:
  - Point landmarks: Euclidean distance (grid-cell units), Successful Detection Rate (SDR), failure categorization
  - Area landmarks: Jaccard, Dice
  - Statistical: Shapiro-Wilk → ANOVA + Tukey + Cohen's d *or* Kruskal-Wallis + Dunn-Bonferroni + rank-biserial; paired Wilcoxon (Bonferroni-corrected) for model-vs-student comparison; Bland-Altman with normality diagnostic; ICC(2,1) Shrout-Fleiss and Cohen's κ for inter-rater reliability.

## Repository layout

```
config.py                  # Models, prompts, grid dims, paths, seed
pipeline.py                # Batch submission + collection + parsing
analysis.py                # Metrics + statistics + Bland-Altman + ICC/κ
consistency_check.py       # Fleiss' κ across pilot repetitions
generate_sota_report.py    # SOTA comparison report
scripts/run_pilot.py       # End-to-end pilot orchestrator (cost-gated, idempotent)
scripts/reproducibility_manifest.py
tests/smoke_e2e.py         # Regression gate
requirements.txt           # Pinned dependencies
SETUP_AND_COSTS.md         # API key + cost guide
data/                      # NOT in git — see "Data" below
results/                   # NOT in git — pipeline output
```

The `data/` and `docs/` directories are intentionally gitignored.

## Requirements

- Python 3.14 (project was tested on `/opt/homebrew/opt/python@3.14`)
- A virtual environment is recommended

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Data

The benchmark Excel and image files are not in this repository. To reproduce, you need:

```
data/
  Dental_MLLM_Benchmark_Data.xlsx
  images/
    Panoramic_Grid/    PAN_001.png ... PAN_100.png
    Periapical_Grid/   PA_001.png ... PA_050.png
    Cephalometric_Grid/CEPH_001.png ... CEPH_050.png
```

The Excel must contain three sheets (`PANORAMIC`, `PERIAPICAL`, `CEPHALOMETRIC`) with columns including `Image_ID`, `Category`, `Target_Structure`, `Question_Prompt`, `OMFR_1` (ground-truth grid cells), and placeholder columns for second-rater and student annotations.

Schema is documented in `config.py` and validated by `pipeline.py` at parse time.

## API keys

See [SETUP_AND_COSTS.md](SETUP_AND_COSTS.md) for full instructions on creating accounts, enabling Batch APIs, and estimated cost.

The pipeline reads keys from `os.environ`. The recommended way is a `.env` file at the project root — both `pipeline.py` and `scripts/run_pilot.py` load it automatically (no `python-dotenv` dependency, just stdlib). `.env` is gitignored, so your keys never leave the machine.

```bash
cp .env.example .env
# then open .env in your editor and paste your real keys
```

Minimum required for the default pilot is `OPENAI_API_KEY`. If you prefer the classic shell-export route, that still works and takes precedence over `.env`:

```bash
export OPENAI_API_KEY=sk-...
```

## Run the pilot (recommended first)

The pilot orchestrator runs a small, cost-capped experiment (~$1.44 max) covering every dimension of the full pipeline. Every API call is gated behind an explicit confirmation prompt, the run is idempotent, and all artefacts go to a sandbox (`results_pilot/`) so the real `results/` tree is never touched.

```bash
.venv/bin/python scripts/run_pilot.py
```

Use `--dry-run` to print what would happen without spending anything, and `--yes` to skip the confirmation prompt (only after a dry run).

## Run the full pipeline

```bash
# 1. Build the query index from the Excel (no API calls)
.venv/bin/python pipeline.py prepare

# 2. Submit batch jobs (requires API keys; prompts for confirmation)
.venv/bin/python pipeline.py submit

# 3. Poll batch status
.venv/bin/python pipeline.py status

# 4. Download finished batches
.venv/bin/python pipeline.py download

# 5. Parse provider responses into unified results
.venv/bin/python pipeline.py parse
```

## Analysis

```bash
.venv/bin/python analysis.py metrics      # Per-query metrics
.venv/bin/python analysis.py stats        # ANOVA / KW + post-hocs
.venv/bin/python analysis.py visualize    # Bland-Altman + heatmaps
.venv/bin/python analysis.py interrater   # ICC(2,1) + Cohen's κ
.venv/bin/python generate_sota_report.py  # SOTA comparison table
```

## Smoke / regression test

```bash
.venv/bin/python tests/smoke_e2e.py
```

This is the regression gate — it exercises every critical path (parse → metrics → stats → visualize → interrater) on a sandbox copy of the data and verifies the real `results/` tree is untouched (mtime check). Run it after any code change.

## Reproducibility

```bash
.venv/bin/python scripts/reproducibility_manifest.py
```

Captures the current git SHA, Python version, `pip freeze`, SHA-256 of the Excel, all images, and every source file, plus the active prompt templates and config parameters. Output is deterministic (no timestamps) so two clean runs on the same state produce identical manifests.

The random seed (`config.RANDOM_SEED = 42`) is wired into OpenAI and Gemini batch requests. Anthropic Batch API does not accept a seed and is intentionally omitted.

## Status

The codebase passed a final deep audit on 2026-04-12 (codebase + data both verified clean, zero CRITICAL bugs). During the first pilot launch attempt, 4 bugs were found and fixed (commit 933b3d4). A subsequent contamination audit discovered that 91/900 queries (10.1%) have ground-truth values that coincidentally match the fixed example coordinates embedded in prompt templates, creating a potential bias. A proposed fix replacing concrete examples with abstract format descriptions has been drafted but requires colleague approval before implementation. The experiment is paused pending this sign-off. The full experiment remains gated on the pilot's success and on populating the `OMFR_2` and `Student_*` columns in the Excel for inter-rater and student-comparison analyses.
