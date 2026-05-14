"""
Configuration for the Dental MLLM Benchmark Study.

Defines landmarks, grid specifications, prompt templates, model configurations,
and all constants used across the pipeline.
"""

import math
import os
from pathlib import Path

# ============================================================
# Paths
# ============================================================
# PROJECT_ROOT always points at the repository root (where this file lives).
# DATA_DIR / IMAGES_DIR / EXCEL_PATH are ALWAYS tied to the real data and must
# NEVER be overridden. Only RESULTS_DIR can be redirected via the
# DENTAL_MLLM_RESULTS_DIR environment variable so that tests and dry runs can
# sandbox their outputs without any chance of contaminating the real derived
# artefacts (e.g. results/query_index.json).
PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"
IMAGES_DIR = DATA_DIR / "images"
EXCEL_PATH = DATA_DIR / "Dental_MLLM_Benchmark_Data.xlsx"

_RESULTS_ENV = os.environ.get("DENTAL_MLLM_RESULTS_DIR")
if _RESULTS_ENV:
    RESULTS_DIR = Path(_RESULTS_ENV).resolve()
else:
    RESULTS_DIR = PROJECT_ROOT / "results"

# Hard refusal: RESULTS_DIR must NEVER coincide with or live inside DATA_DIR.
# This protects the raw benchmark data (Excel + images) from any possible
# write contamination, no matter what the env var or CLI flags say. Triggers
# at import time so every entry point (pipeline.py, analysis.py, run_pilot.py,
# tests, scripts) is covered uniformly.
_data_resolved = DATA_DIR.resolve()
_results_resolved = RESULTS_DIR.resolve()
if _results_resolved == _data_resolved or _data_resolved in _results_resolved.parents:
    raise RuntimeError(
        f"REFUSING to use {_results_resolved} as RESULTS_DIR — it is the raw "
        f"benchmark data directory ({_data_resolved}) or lives inside it. "
        f"Pick a sandbox outside data/ (e.g. results/ or results_pilot_v2/)."
    )

BATCH_DIR = RESULTS_DIR / "batches"
RESPONSES_DIR = RESULTS_DIR / "responses"
ANALYSIS_DIR = RESULTS_DIR / "analysis"

# ============================================================
# Grid Specifications per Modality
# ============================================================
GRID_SPECS = {
    "PANORAMIC": {
        "cols": 16,
        "rows": 8,
        "row_labels": "A-H",
        "col_labels": "1-16",
        "max_row_letter": "H",
        "diagonal": math.sqrt(15 ** 2 + 7 ** 2),  # sqrt(274) ≈ 16.5529
        "image_prefix": "PAN",
        "image_dir": IMAGES_DIR / "Panoramic_Grid",
        "landmarks_per_image": 6,
    },
    "PERIAPICAL": {
        "cols": 8,
        "rows": 6,
        "row_labels": "A-F",
        "col_labels": "1-8",
        "max_row_letter": "F",
        "diagonal": math.sqrt(7 ** 2 + 5 ** 2),  # sqrt(74) ≈ 8.6023
        "image_prefix": "PA",
        "image_dir": IMAGES_DIR / "Periapical_Grid",
        "landmarks_per_image": 3,
    },
    "CEPHALOMETRIC": {
        "cols": 10,
        "rows": 8,
        "row_labels": "A-H",
        "col_labels": "1-10",
        "max_row_letter": "H",
        "diagonal": math.sqrt(9 ** 2 + 7 ** 2),  # sqrt(130) ≈ 11.4018
        "image_prefix": "CEPH",
        "image_dir": IMAGES_DIR / "Cephalometric_Grid",
        "landmarks_per_image": 3,
    },
}

# ============================================================
# Landmark Definitions
# ============================================================
LANDMARKS = {
    "PANORAMIC": [
        {
            "structure": "Mental_Foramen_L",
            "type": "point",
            "description_en": "the center of the left mental foramen",
            "description_tr": "Sol Mental Foramenin tam orta noktası",
        },
        {
            "structure": "Condylar_Head_R",
            "type": "point",
            "description_en": "the most superior (highest) point of the right condylar head",
            "description_tr": "Sağ kondilin en üst (superior) noktası",
        },
        {
            "structure": "Tooth_33_Apex",
            "type": "point",
            "description_en": "the root apex of tooth #33 (lower left canine)",
            "description_en_no_fdi": "the root apex of the lower left canine",
            "description_en_patient_frame": "the root apex of tooth #33 (patient's left lower canine)",
            "description_tr": "33 numaralı dişin kök ucu noktası",
            "uses_fdi": True,
        },
        {
            "structure": "Mandibular_Canal_L",
            "type": "area",
            "description_en": "all grid cells through which the left mandibular canal passes",
            "description_tr": "Sol Mandibular Kanal hattının geçtiği tüm kareler",
        },
        {
            "structure": "Maxillary_Sinus_R",
            "type": "area",
            "description_en": "all grid cells within the boundaries of the right maxillary sinus",
            "description_tr": "Sağ Maksiller Sinüs boşluğunun sınırları içindeki kareler",
        },
        {
            "structure": "External_Oblique_Ridge_R",
            "type": "area",
            "description_en": "all grid cells along the line traced by the right external oblique ridge (linea obliqua externa)",
            "description_tr": "Sağ eksternal oblik sırtın (linea obliqua externa) izlediği hat üzerindeki tüm kareler",
        },
    ],
    "PERIAPICAL": [
        {
            "structure": "Tooth_36_Distal_Apex",
            "type": "point",
            "description_en": "the most apical (tip) point of the distal root of tooth #36 (lower left first molar)",
            "description_en_no_fdi": "the most apical (tip) point of the distal root of the lower left first molar",
            "description_en_patient_frame": "the most apical (tip) point of the distal root of tooth #36 (patient's left lower first molar)",
            "description_tr": "36 numaralı dişin distal kök ucunun bittiği en uç nokta",
            "uses_fdi": True,
        },
        {
            "structure": "Tooth_36_Mesial_CEJ",
            "type": "point",
            "description_en": "the cemento-enamel junction (CEJ) on the mesial side of tooth #36 (lower left first molar)",
            "description_en_no_fdi": "the cemento-enamel junction (CEJ) on the mesial side of the lower left first molar",
            "description_en_patient_frame": "the cemento-enamel junction (CEJ) on the mesial side of tooth #36 (patient's left lower first molar)",
            "description_tr": "36 numaralı dişin mezial servikal bölgesinde, mine ile kök yüzeyinin birleştiği nokta",
            "uses_fdi": True,
        },
        {
            "structure": "Tooth_36_Distal_CEJ",
            "type": "point",
            "description_en": "the cemento-enamel junction (CEJ) on the distal side of tooth #36 (lower left first molar)",
            "description_en_no_fdi": "the cemento-enamel junction (CEJ) on the distal side of the lower left first molar",
            "description_en_patient_frame": "the cemento-enamel junction (CEJ) on the distal side of tooth #36 (patient's left lower first molar)",
            "description_tr": "36 numaralı dişin distal servikal bölgesinde, mine ile kök yüzeyinin birleştiği nokta",
            "uses_fdi": True,
        },
    ],
    "CEPHALOMETRIC": [
        {
            "structure": "Sella_S",
            "type": "point",
            "description_en": "the geometric center of the sella turcica (Sella point)",
            "description_tr": "Sella turcica'nın tam orta noktası",
        },
        {
            "structure": "Nasion_N",
            "type": "point",
            "description_en": "the most anterior point of the nasofrontal suture (Nasion)",
            "description_tr": "Alın ile burun kökünün birleştiği sütürdaki en ön nokta",
        },
        {
            "structure": "Menton_Me",
            "type": "point",
            "description_en": "the lowermost point of the mandibular symphysis (Menton)",
            "description_tr": "Alt çene ucu kemiğinin (Simfiz) en alt noktası",
        },
    ],
}

# ============================================================
# Prompt Templates
# ============================================================
SYSTEM_PROMPT = (
    "You are an expert Oral and Maxillofacial Radiologist with extensive "
    "experience in identifying anatomic landmarks on dental radiographs."
)

# Strategy 1: Zero-Shot Baseline
ZERO_SHOT_POINT_TEMPLATE = (
    "This image is overlaid with a {cols}x{rows} square grid. "
    "{fdi_prefix}{Identify} the coordinate of the cell containing {landmark_description}. "
    "Respond with a single coordinate formatted as a row letter followed by a column number "
    "(no spaces, no punctuation). "
    "Provide only the coordinate in your response."
)

ZERO_SHOT_AREA_TEMPLATE = (
    "This image is overlaid with a {cols}x{rows} square grid. "
    "{fdi_prefix}{Identify} {landmark_description}. "
    "List all cells separated by commas, each formatted as a row letter followed by a column number. "
    "Provide only the coordinates in your response."
)

# Strategy 2: Guided (Zero-Shot with Detailed Grid Explanation)
GUIDED_SYSTEM_ADDITION = (
    "The image has a grid overlay with columns numbered 1 through {cols} from left to right, "
    "and rows labeled A through {max_row} from top to bottom. "
    "Each cell is identified by its row letter followed by its column number, "
    "with no space or punctuation between them. "
    "The grid lines are drawn in cyan and labels are in yellow. "
    "{modality_clause}"
    "For point-based questions, respond with exactly one cell coordinate. "
    "For area-based questions, list all cells the structure occupies, separated by commas."
)

# Modality-specific clinical clauses injected into GUIDED_SYSTEM_ADDITION via
# {modality_clause}. Empty string for modalities without a specific note.
GUIDED_MODALITY_CLAUSES = {
    "PANORAMIC": (
        "In panoramic radiographs, the patient's right side appears on the left side of "
        "the image, and the patient's left side appears on the right side of the image. "
    ),
    "PERIAPICAL": "",
    "CEPHALOMETRIC": "",
}

GUIDED_POINT_TEMPLATE = (
    "{fdi_prefix}{Identify} the coordinate of the cell containing {landmark_description}. "
    "Provide only the coordinate in your response."
)

GUIDED_AREA_TEMPLATE = (
    "{fdi_prefix}{Identify} {landmark_description}. "
    "List all cells separated by commas. "
    "Provide only the coordinates in your response."
)

# ============================================================
# Model Configurations
# ============================================================
# Fixed reproducibility seed. Passed to providers that accept a `seed`
# parameter (OpenAI, Google Gemini). Anthropic does not expose a seed on
# the Messages Batches API, so for Anthropic we rely solely on temperature=0.
# Change this only when you intentionally want a fresh reproducibility cohort.
RANDOM_SEED = 42

MODELS = {
    "gpt-5.4": {
        "provider": "openai",
        "model_id": "gpt-5.4",
        "display_name": "GPT-5.4",
        "batch_api": True,
        "batch_input_price_per_1m": 1.25,   # 50% off standard $2.50
        "batch_output_price_per_1m": 7.50,   # 50% off standard $15.00
        "max_output_tokens": 256,
        "temperature": 0,
        "seed": RANDOM_SEED,                 # best-effort reproducibility
    },
    "gemini-3.1-pro": {
        "provider": "google",
        "model_id": "gemini-3.1-pro-preview",
        "display_name": "Gemini 3.1 Pro",
        "batch_api": True,
        "batch_input_price_per_1m": 1.00,   # 50% off standard $2.00
        "batch_output_price_per_1m": 6.00,   # 50% off standard $12.00 (≤200K context)
        # Gemini 3.1 Pro uses thinking mode internally. The 2026-05-14
        # mini-canary observed 244 thinking tokens + answer tokens fighting for
        # the same budget. At 256 max, area-landmark answers got truncated
        # (finishReason=MAX_TOKENS) and most batch requests were cancelled.
        # Raised to 2048 to leave generous headroom for thinking + answer;
        # Gemini bills actual usage, so no cost penalty for unused budget.
        "max_output_tokens": 2048,
        "temperature": 0,
        "seed": RANDOM_SEED,                 # Gemini generationConfig.seed
        "media_resolution": "MEDIA_RESOLUTION_HIGH",  # ~2200 tokens/image (max in v1beta)
    },
    "claude-sonnet-4.6": {
        "provider": "anthropic",
        "model_id": "claude-sonnet-4-6-20250514",
        "display_name": "Claude Sonnet 4.6",
        "batch_api": True,
        "batch_input_price_per_1m": 1.50,   # 50% off standard $3.00
        "batch_output_price_per_1m": 7.50,   # 50% off standard $15.00
        "max_output_tokens": 256,
        "temperature": 0,
        # No "seed" key — Anthropic does not accept seed on batches.
    },
}

# Which models to run (toggle here)
ACTIVE_MODELS = ["gpt-5.4", "gemini-3.1-pro"]
# ACTIVE_MODELS = ["gpt-5.4", "gemini-3.1-pro", "claude-sonnet-4.6"]  # with optional 3rd

# ============================================================
# Prompting Strategies
# ============================================================
# STRATEGIES is the "active" list — strategies that get submitted by default
# when pipeline.py submit is run without --strategies override. It is the
# canonical pair (zero_shot, guided) that drives every main run.
STRATEGIES = ["zero_shot", "guided"]

# ALL_STRATEGIES is the full registry of every strategy name that
# pipeline.generate_prompt knows how to render. It MUST include every key
# that generate_prompt has a branch for — i.e. STRATEGIES plus every ablation
# variant. Parser code in pipeline.parse_response_filename uses this list
# (NOT STRATEGIES) so that ablation response files like
# `gpt-5.4_guided_no_tooth_num_chunk000_results.jsonl` correctly attribute
# their `strategy` field to "guided_no_tooth_num" rather than collapsing to
# "guided" (which would silently mis-label every ablation record).
#
# When you add a new ablation strategy:
#   1. Add a branch in pipeline.generate_prompt(query, strategy)
#   2. Append the strategy name HERE so the response parser recognises it
ALL_STRATEGIES = [
    "zero_shot",
    "guided",
    # Section-10 ablations
    "guided_no_tooth_num",
    "guided_patient_left",
    "guided_no_LR",
]

# ============================================================
# API Configuration
# ============================================================
# API keys are loaded from environment variables
# Set these before running:
#   export OPENAI_API_KEY="sk-..."
#   export GOOGLE_API_KEY="..."
#   export ANTHROPIC_API_KEY="sk-ant-..."

def get_api_key(provider):
    """Get API key for a provider from environment variables."""
    env_vars = {
        "openai": "OPENAI_API_KEY",
        "google": "GOOGLE_API_KEY",
        "anthropic": "ANTHROPIC_API_KEY",
    }
    key = os.environ.get(env_vars.get(provider, ""))
    if not key:
        raise ValueError(
            f"API key not set for {provider}. "
            f"Set {env_vars[provider]} environment variable."
        )
    return key
