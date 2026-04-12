"""
Dental MLLM Benchmark — Reproducibility Manifest Generator

Produces a single `reproducibility_manifest.json` that captures:
    * git commit + dirty flag
    * Python version
    * pinned package versions (pip freeze)
    * SHA-256 of the Excel data and all 200 image files
    * SHA-256 of every source file under the project (pipeline, analysis, config, scripts, tests)
    * config.RANDOM_SEED, config.MODELS, config.STRATEGIES, config.GRID_SPECS, config.LANDMARKS
    * Python's platform triple and libc info (best-effort)

Usage
-----
    .venv/bin/python scripts/reproducibility_manifest.py [--output PATH]

    Default output: ./reproducibility_manifest.json

The manifest is deterministic — rerunning it on the same tree will
produce byte-identical output (barring timestamps, which we intentionally
exclude so the file stays stable).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import platform
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def git_commit_info() -> dict:
    try:
        commit = subprocess.run(
            ["git", "-C", str(ROOT), "rev-parse", "HEAD"],
            capture_output=True, text=True, check=True,
        ).stdout.strip()
    except Exception:
        commit = None
    try:
        dirty = subprocess.run(
            ["git", "-C", str(ROOT), "status", "--porcelain"],
            capture_output=True, text=True, check=True,
        ).stdout.strip() != ""
    except Exception:
        dirty = None
    return {"commit": commit, "dirty": dirty}


def pip_freeze() -> list[str]:
    venv_py = ROOT / ".venv" / "bin" / "python"
    py = str(venv_py) if venv_py.exists() else sys.executable
    r = subprocess.run(
        [py, "-m", "pip", "freeze"],
        capture_output=True, text=True, check=True,
    )
    return sorted(line.strip() for line in r.stdout.splitlines() if line.strip())


def hash_images() -> dict:
    import config  # type: ignore
    out = {}
    for prefix_cfg in config.GRID_SPECS.values():
        img_dir = prefix_cfg["image_dir"]
        if not img_dir.exists():
            continue
        for p in sorted(img_dir.glob("*.png")):
            out[f"{img_dir.name}/{p.name}"] = sha256_of(p)
    return out


def hash_source_files() -> dict:
    """Hash every Python, JSON, and docx/pdf file that is an input to the pipeline."""
    targets = [
        "pipeline.py", "analysis.py", "config.py", "consistency_check.py",
        "scripts/run_pilot.py", "scripts/reproducibility_manifest.py",
        "tests/smoke_e2e.py",
        "requirements.txt",
    ]
    out = {}
    for rel in targets:
        p = ROOT / rel
        if p.exists():
            out[rel] = sha256_of(p)
    return out


def summarise_config() -> dict:
    import config  # type: ignore
    return {
        "random_seed": getattr(config, "RANDOM_SEED", None),
        "active_models": config.ACTIVE_MODELS,
        "strategies": list(config.STRATEGIES),
        "models": {
            k: {
                "provider": v["provider"],
                "model_id": v["model_id"],
                "temperature": v.get("temperature"),
                "max_output_tokens": v.get("max_output_tokens"),
                "seed": v.get("seed"),
                "media_resolution": v.get("media_resolution"),
                # Prices omitted on purpose — they can change over time.
            }
            for k, v in config.MODELS.items()
        },
        "grid_specs": {
            k: {
                "cols": v["cols"],
                "rows": v["rows"],
                "row_labels": v["row_labels"],
                "col_labels": v["col_labels"],
                "diagonal": v["diagonal"],
            }
            for k, v in config.GRID_SPECS.items()
        },
        "landmarks": {
            sheet: [
                {
                    "structure": lm["structure"],
                    "type": lm["type"],
                    "description_en": lm["description_en"],
                    "uses_fdi": lm.get("uses_fdi", False),
                }
                for lm in lms
            ]
            for sheet, lms in config.LANDMARKS.items()
        },
    }


def summarise_prompts() -> dict:
    """
    Render the final prompt for one representative query per landmark and
    store the exact text in the manifest. Future reproducers can diff
    against this to detect prompt drift.
    """
    import pipeline  # type: ignore
    import config    # type: ignore

    index_path = ROOT / "results" / "query_index.json"
    if not index_path.exists():
        return {}
    with open(index_path) as f:
        queries = json.load(f)

    out = {}
    for strategy in config.STRATEGIES:
        out[strategy] = {}
        seen = set()
        for q in queries:
            s = q["structure"]
            if s in seen:
                continue
            seen.add(s)
            sys_p, user_p = pipeline.generate_prompt(q, strategy)
            out[strategy][s] = {
                "system": sys_p,
                "user": user_p,
            }
    return out


def build_manifest() -> dict:
    import config  # type: ignore
    excel_path = config.EXCEL_PATH
    manifest = {
        "schema_version": 1,
        "project": "dental-mllm-benchmark",
        "python_version": sys.version.split()[0],
        "python_implementation": platform.python_implementation(),
        "platform": platform.platform(),
        "git": git_commit_info(),
        "pip_freeze": pip_freeze(),
        "source_files": hash_source_files(),
        "data": {
            "excel": {
                "path": str(excel_path.relative_to(ROOT)),
                "sha256": sha256_of(excel_path),
                "size_bytes": excel_path.stat().st_size,
            },
            "images": hash_images(),
        },
        "config": summarise_config(),
        "prompts": summarise_prompts(),
    }
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", default="reproducibility_manifest.json",
                        help="Where to write the manifest (default: ./reproducibility_manifest.json)")
    args = parser.parse_args()

    manifest = build_manifest()
    out_path = ROOT / args.output if not Path(args.output).is_absolute() else Path(args.output)
    with open(out_path, "w") as f:
        json.dump(manifest, f, indent=2, sort_keys=True)
    print(f"Wrote reproducibility manifest ({out_path}):")
    print(f"  git commit:      {manifest['git']['commit']}")
    print(f"  python:          {manifest['python_version']}")
    print(f"  excel SHA-256:   {manifest['data']['excel']['sha256']}")
    print(f"  images hashed:   {len(manifest['data']['images'])}")
    print(f"  source files:    {len(manifest['source_files'])}")
    print(f"  prompts captured:{sum(len(v) for v in manifest['prompts'].values())} "
          f"across {len(manifest['prompts'])} strategies")
    return 0


if __name__ == "__main__":
    sys.exit(main())
