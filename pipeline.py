"""
Dental MLLM Benchmark Pipeline

Reads the benchmark Excel, generates prompts for each query × model × strategy,
submits batch API requests, and collects responses.

Usage:
    # Step 1: Prepare query index (no API key needed)
    python3 pipeline.py prepare

    # Step 2: Submit batches (requires API keys, encodes images on the fly)
    python3 pipeline.py submit

    # Step 3: Check batch status
    python3 pipeline.py status

    # Step 4: Download results
    python3 pipeline.py download

    # Step 5: Parse responses into unified results
    python3 pipeline.py parse
"""

import argparse
import base64
import json
import os
import sys
import tempfile
import time
from pathlib import Path

import config


# ============================================================
# Excel Parser
# ============================================================

def parse_excel():
    """
    Parse the benchmark Excel file into a list of query dicts.

    Extracted columns (per sheet):
      col  1: Image_ID (derived from row position)
      col  2: Category
      col  3: Target_Structure
      col  4: Question_Prompt (TR)
      col  5: OMFR_1            (first specialist ground truth)
      col  6: OMFR_2            (second specialist ground truth)
      col  7: CONSENSUS_Ground_Truth
      col 10–49: Student_1 .. Student_40

    Returns list of query dicts with keys:
        query_id, sheet, image_id, image_path, category, structure,
        landmark_type, landmark_description_en, uses_fdi, original_prompt_tr,
        omfr_1, omfr_2, consensus_gt, students (list of 40 strings-or-None)
    """
    import openpyxl

    wb = openpyxl.load_workbook(str(config.EXCEL_PATH))
    queries = []

    for sheet_name in ["PANORAMIC", "PERIAPICAL", "CEPHALOMETRIC"]:
        ws = wb[sheet_name]
        grid = config.GRID_SPECS[sheet_name]
        landmarks = config.LANDMARKS[sheet_name]
        lm_per_image = grid["landmarks_per_image"]

        for row_idx in range(2, ws.max_row + 1):
            data_row_num = row_idx - 1
            image_num = -(-data_row_num // lm_per_image)  # ceiling division
            image_id = f"{grid['image_prefix']}_{image_num:03d}"

            category = ws.cell(row_idx, 2).value
            structure = ws.cell(row_idx, 3).value
            prompt_tr = ws.cell(row_idx, 4).value
            omfr_1 = ws.cell(row_idx, 5).value
            omfr_2 = ws.cell(row_idx, 6).value
            consensus_gt = ws.cell(row_idx, 7).value
            students = [ws.cell(row_idx, c).value for c in range(10, 50)]

            if not structure:
                continue

            landmark_def = next(
                (lm for lm in landmarks if lm["structure"] == structure), None
            )
            if not landmark_def:
                print(f"WARNING: Unknown structure '{structure}' in {sheet_name} row {row_idx}")
                continue

            image_path = grid["image_dir"] / f"{image_id}.png"

            queries.append({
                "query_id": f"{image_id}_{structure}",
                "sheet": sheet_name,
                "image_id": image_id,
                "image_path": str(image_path),
                "category": category,
                "structure": structure,
                "landmark_type": landmark_def["type"],
                "landmark_description_en": landmark_def["description_en"],
                "uses_fdi": landmark_def.get("uses_fdi", False),
                "original_prompt_tr": prompt_tr,
                "omfr_1": omfr_1,
                "omfr_2": omfr_2,
                "consensus_gt": consensus_gt,
                "students": students,
            })

    wb.close()
    print(f"Parsed {len(queries)} queries from Excel "
          f"({sum(1 for q in queries if q['sheet'] == 'PANORAMIC')} PAN, "
          f"{sum(1 for q in queries if q['sheet'] == 'PERIAPICAL')} PA, "
          f"{sum(1 for q in queries if q['sheet'] == 'CEPHALOMETRIC')} CEPH)")
    return queries


# ============================================================
# Prompt Generation
# ============================================================

def generate_prompt(query, strategy):
    """Generate (system_prompt, user_prompt) for a query + strategy."""
    sheet = query["sheet"]
    grid = config.GRID_SPECS[sheet]
    lm_type = query["landmark_type"]
    desc = query["landmark_description_en"]
    uses_fdi = query.get("uses_fdi", False)

    # FDI clause for landmarks that reference tooth numbers
    fdi_prefix = "Using the FDI two-digit numbering system, " if uses_fdi else ""
    identify = "identify" if uses_fdi else "Identify"

    if strategy == "zero_shot":
        system_prompt = config.SYSTEM_PROMPT
        if lm_type == "point":
            user_prompt = config.ZERO_SHOT_POINT_TEMPLATE.format(
                cols=grid["cols"], rows=grid["rows"],
                landmark_description=desc,
                fdi_prefix=fdi_prefix, Identify=identify,
            )
        else:
            user_prompt = config.ZERO_SHOT_AREA_TEMPLATE.format(
                cols=grid["cols"], rows=grid["rows"],
                landmark_description=desc,
                fdi_prefix=fdi_prefix, Identify=identify,
            )

    elif strategy == "guided":
        modality_clause = config.GUIDED_MODALITY_CLAUSES.get(sheet, "")
        grid_explanation = config.GUIDED_SYSTEM_ADDITION.format(
            cols=grid["cols"],
            max_row=grid["max_row_letter"],
            modality_clause=modality_clause,
        )
        system_prompt = config.SYSTEM_PROMPT + "\n\n" + grid_explanation

        if lm_type == "point":
            user_prompt = config.GUIDED_POINT_TEMPLATE.format(
                landmark_description=desc,
                fdi_prefix=fdi_prefix, Identify=identify,
            )
        else:
            user_prompt = config.GUIDED_AREA_TEMPLATE.format(
                landmark_description=desc,
                fdi_prefix=fdi_prefix, Identify=identify,
            )
    else:
        raise ValueError(f"Unknown strategy: {strategy}")

    return system_prompt, user_prompt


def encode_image_base64(image_path):
    """Read an image file and return its base64 encoding."""
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


# ============================================================
# Image Cache (avoid re-encoding same image for multiple landmarks)
# ============================================================

class ImageCache:
    """Cache base64-encoded images to avoid re-reading from disk."""
    def __init__(self):
        self._cache = {}

    def get(self, image_path):
        if image_path not in self._cache:
            self._cache[image_path] = encode_image_base64(image_path)
        return self._cache[image_path]

    def clear(self):
        self._cache.clear()


# ============================================================
# Request Builders (per provider)
# ============================================================

def build_openai_request(query, strategy, model_cfg, img_b64):
    """Build a single OpenAI batch API request line (dict)."""
    system_prompt, user_prompt = generate_prompt(query, strategy)
    body = {
        "model": model_cfg["model_id"],
        "temperature": model_cfg["temperature"],
        "max_completion_tokens": model_cfg["max_output_tokens"],
        "messages": [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{img_b64}",
                            "detail": "high",
                        },
                    },
                    {"type": "text", "text": user_prompt},
                ],
            },
        ],
    }
    # Pass seed when configured — best-effort reproducibility. The response
    # will include system_fingerprint so we can detect backend changes.
    if "seed" in model_cfg and model_cfg["seed"] is not None:
        body["seed"] = model_cfg["seed"]
    return {
        "custom_id": f"{query['query_id']}_{strategy}",
        "method": "POST",
        "url": "/v1/chat/completions",
        "body": body,
    }


def build_google_request(query, strategy, model_cfg, img_b64):
    """Build a single Google Gemini batch request (dict)."""
    system_prompt, user_prompt = generate_prompt(query, strategy)
    generation_config = {
        "temperature": model_cfg["temperature"],
        "maxOutputTokens": model_cfg["max_output_tokens"],
        "mediaResolution": model_cfg.get(
            "media_resolution", "MEDIA_RESOLUTION_HIGH"
        ),
    }
    # Gemini accepts a seed integer on generationConfig for reproducibility.
    if "seed" in model_cfg and model_cfg["seed"] is not None:
        generation_config["seed"] = model_cfg["seed"]
    return {
        "custom_id": f"{query['query_id']}_{strategy}",
        "request": {
            "model": model_cfg["model_id"],
            "systemInstruction": {"parts": [{"text": system_prompt}]},
            "contents": [{
                "role": "user",
                "parts": [
                    {"inlineData": {"mimeType": "image/png", "data": img_b64}},
                    {"text": user_prompt},
                ],
            }],
            "generationConfig": generation_config,
        },
    }


def build_anthropic_request(query, strategy, model_cfg, img_b64):
    """Build a single Anthropic Message Batches request (dict)."""
    system_prompt, user_prompt = generate_prompt(query, strategy)
    return {
        "custom_id": f"{query['query_id']}_{strategy}",
        "params": {
            "model": model_cfg["model_id"],
            "max_tokens": model_cfg["max_output_tokens"],
            "temperature": model_cfg["temperature"],
            "system": system_prompt,
            "messages": [{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": img_b64,
                        },
                    },
                    {"type": "text", "text": user_prompt},
                ],
            }],
        },
    }


REQUEST_BUILDERS = {
    "openai": build_openai_request,
    "google": build_google_request,
    "anthropic": build_anthropic_request,
}


# ============================================================
# Prepare Command
# ============================================================

def _select_subset_images(queries, subset_images=None, n_per_modality=None):
    """
    Filter the full query list down to a subset specified by image_id.

    * subset_images: explicit set of image_id strings to keep. Wins over n_per_modality.
    * n_per_modality: deterministic stride selection of N images per modality.
      Uses np.linspace-style indexing to cover the full modality range (first
      and last images always included). This is more informative than taking
      the first N because annotators may have drifted over the course of the
      dataset.

    Returns the filtered queries list and a dict describing what was selected.
    """
    if subset_images is None and n_per_modality is None:
        return queries, None

    # Group queries by (sheet, image_id) preserving order
    from collections import OrderedDict
    by_sheet = OrderedDict()
    for q in queries:
        by_sheet.setdefault(q["sheet"], OrderedDict())
        by_sheet[q["sheet"]].setdefault(q["image_id"], []).append(q)

    keep_ids = set()
    plan = {}

    if subset_images:
        keep_ids = {s.strip() for s in subset_images if s.strip()}
        plan["mode"] = "explicit_subset"
        plan["image_ids"] = sorted(keep_ids)
    else:
        plan["mode"] = "n_per_modality"
        plan["n_per_modality"] = n_per_modality
        plan["per_modality_selection"] = {}
        for sheet, images in by_sheet.items():
            img_list = list(images.keys())  # already in insertion order
            n_total = len(img_list)
            if n_per_modality >= n_total:
                chosen = img_list[:]
            else:
                # Evenly-spaced selection, inclusive of first and last image
                step = (n_total - 1) / (n_per_modality - 1) if n_per_modality > 1 else 0
                indices = sorted({round(i * step) for i in range(n_per_modality)})
                # Pad if rounding collisions reduced the count
                while len(indices) < n_per_modality and len(indices) < n_total:
                    for j in range(n_total):
                        if j not in indices:
                            indices.append(j)
                            break
                indices = sorted(indices)[:n_per_modality]
                chosen = [img_list[i] for i in indices]
            plan["per_modality_selection"][sheet] = chosen
            keep_ids.update(chosen)

    filtered = [q for q in queries if q["image_id"] in keep_ids]
    plan["total_queries_kept"] = len(filtered)
    plan["total_images_kept"] = len(keep_ids)
    return filtered, plan


def cmd_prepare(args):
    """Parse Excel, validate images, and save query index."""
    queries = parse_excel()

    # Optional subset filter (for pilot runs / preliminary experiment).
    # Subset is applied BEFORE image validation so we only require the
    # selected subset to be present on disk.
    subset_images = None
    if getattr(args, "subset_images", None):
        subset_images = [s.strip() for s in args.subset_images.split(",") if s.strip()]
    n_per_modality = getattr(args, "n_per_modality", None)

    queries, plan = _select_subset_images(queries, subset_images, n_per_modality)
    if plan is not None:
        print(f"\nSubset filter applied ({plan['mode']}):")
        if plan["mode"] == "explicit_subset":
            print(f"  image_ids: {plan['image_ids']}")
        else:
            for sheet, imgs in plan["per_modality_selection"].items():
                print(f"  {sheet}: {len(imgs)} images → {imgs}")
        print(f"  Total kept: {plan['total_queries_kept']} queries from "
              f"{plan['total_images_kept']} images")

    # Validate that all SELECTED images exist
    missing = [q for q in queries if not Path(q["image_path"]).exists()]
    if missing:
        print(f"ERROR: {len(missing)} images not found:")
        for q in missing[:5]:
            print(f"  {q['image_path']}")
        sys.exit(1)
    print(f"All {len(queries)} images verified.")

    # Save query index
    config.RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    index_path = config.RESULTS_DIR / "query_index.json"
    with open(index_path, "w") as f:
        json.dump(queries, f, indent=2)
    print(f"Query index saved: {index_path}")

    # Show summary of what will be submitted
    active_models = getattr(args, "models", None) or config.ACTIVE_MODELS
    print(f"\nActive models: {active_models}")
    print(f"Strategies: {config.STRATEGIES}")
    print(f"Total API calls: {len(queries)} × {len(config.ACTIVE_MODELS)} × {len(config.STRATEGIES)} "
          f"= {len(queries) * len(config.ACTIVE_MODELS) * len(config.STRATEGIES)}")


# ============================================================
# Submit Command
# ============================================================

# OpenAI batch file size limit ~200MB; each request is ~3MB → ~60 per chunk
OPENAI_CHUNK_SIZE = 50
# Google batchGenerateContent: max 100 requests per call
GOOGLE_CHUNK_SIZE = 100
# Anthropic: 100K requests per batch, but we chunk for memory
ANTHROPIC_CHUNK_SIZE = 100


_OPENAI_TRANSIENT_HTTP = {401, 408, 429, 500, 502, 503, 504}


def _openai_call_with_retries(fn, *args, _label="openai call", _max_attempts=4, **kwargs):
    """
    Retry wrapper for OpenAI HTTP calls. Retries on transient HTTP errors
    (including 401, which we have observed as a transient OpenAI auth flake)
    with exponential backoff: 2s, 4s, 8s. Re-raises after _max_attempts.
    """
    import urllib.error
    last_err = None
    for attempt in range(1, _max_attempts + 1):
        try:
            return fn(*args, **kwargs)
        except urllib.error.HTTPError as e:
            if e.code not in _OPENAI_TRANSIENT_HTTP or attempt == _max_attempts:
                raise
            last_err = e
            wait = 2 ** attempt
            print(f"    {_label}: HTTP {e.code} on attempt {attempt}/{_max_attempts}; "
                  f"retrying in {wait}s")
            time.sleep(wait)
        except (urllib.error.URLError, ConnectionError) as e:
            if attempt == _max_attempts:
                raise
            last_err = e
            wait = 2 ** attempt
            print(f"    {_label}: {type(e).__name__} on attempt {attempt}/{_max_attempts}; "
                  f"retrying in {wait}s")
            time.sleep(wait)
    raise last_err  # pragma: no cover (loop always returns or raises)


def upload_openai_file(jsonl_path, api_key):
    """Upload a JSONL file to OpenAI and return the file ID."""
    import urllib.request

    boundary = "----BatchBoundary"
    file_data = Path(jsonl_path).read_bytes()
    filename = Path(jsonl_path).name

    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="purpose"\r\n\r\nbatch\r\n'
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
        f"Content-Type: application/jsonl\r\n\r\n"
    ).encode() + file_data + f"\r\n--{boundary}--\r\n".encode()

    req = urllib.request.Request(
        "https://api.openai.com/v1/files",
        data=body,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": f"multipart/form-data; boundary={boundary}",
        },
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        return json.loads(resp.read())["id"]


def create_openai_batch(file_id, model_key, api_key):
    """Create an OpenAI batch from an uploaded file."""
    import urllib.request

    body = json.dumps({
        "input_file_id": file_id,
        "endpoint": "/v1/chat/completions",
        "completion_window": "24h",
        "metadata": {"model": model_key},
    }).encode()

    req = urllib.request.Request(
        "https://api.openai.com/v1/batches",
        data=body,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


def submit_openai(queries, model_key, strategy, api_key):
    """Submit queries to OpenAI Batch API in chunks."""
    model_cfg = config.MODELS[model_key]
    img_cache = ImageCache()
    batch_ids = []

    for chunk_idx in range(0, len(queries), OPENAI_CHUNK_SIZE):
        chunk = queries[chunk_idx:chunk_idx + OPENAI_CHUNK_SIZE]
        chunk_num = chunk_idx // OPENAI_CHUNK_SIZE + 1
        total_chunks = -(-len(queries) // OPENAI_CHUNK_SIZE)

        print(f"  Chunk {chunk_num}/{total_chunks} ({len(chunk)} requests)...")

        # Write chunk to temp JSONL file
        tmp_path = None
        try:
            # Assign tmp_path IMMEDIATELY so the finally block can clean up
            # if the write loop raises (e.g., missing image file). The prior
            # version assigned tmp_path only after the loop completed, which
            # leaked temp files on mid-loop exceptions.
            with tempfile.NamedTemporaryFile(mode="w", suffix=".jsonl", delete=False) as tmp:
                tmp_path = tmp.name
                for query in chunk:
                    img_b64 = img_cache.get(query["image_path"])
                    req = build_openai_request(query, strategy, model_cfg, img_b64)
                    tmp.write(json.dumps(req) + "\n")

            file_id = _openai_call_with_retries(
                upload_openai_file, tmp_path, api_key,
                _label=f"upload chunk{chunk_num}")
            print(f"    Uploaded file: {file_id}")
            batch_resp = _openai_call_with_retries(
                create_openai_batch, file_id, model_key, api_key,
                _label=f"create batch chunk{chunk_num}")
            batch_id = batch_resp["id"]
            print(f"    Batch created: {batch_id}")
            batch_ids.append(batch_id)
        finally:
            if tmp_path is not None and os.path.exists(tmp_path):
                os.unlink(tmp_path)

        # Clear image cache after every chunk to bound peak memory.
        img_cache.clear()

    return batch_ids


def submit_google(queries, model_key, strategy, api_key):
    """Submit queries to Google Gemini batchGenerateContent API."""
    import urllib.request

    model_cfg = config.MODELS[model_key]
    img_cache = ImageCache()
    chunk_results = []

    config.RESPONSES_DIR.mkdir(parents=True, exist_ok=True)
    batch_name = f"{model_key}_{strategy}"

    for chunk_idx in range(0, len(queries), GOOGLE_CHUNK_SIZE):
        chunk = queries[chunk_idx:chunk_idx + GOOGLE_CHUNK_SIZE]
        chunk_num = chunk_idx // GOOGLE_CHUNK_SIZE + 1
        total_chunks = -(-len(queries) // GOOGLE_CHUNK_SIZE)

        print(f"  Chunk {chunk_num}/{total_chunks} ({len(chunk)} requests)...")

        # Build requests for this chunk
        custom_ids = []
        batch_requests = []
        for query in chunk:
            img_b64 = img_cache.get(query["image_path"])
            req_data = build_google_request(query, strategy, model_cfg, img_b64)
            custom_ids.append(req_data["custom_id"])
            batch_requests.append(req_data["request"])

        body = json.dumps({"requests": batch_requests}).encode()
        model_id = model_cfg["model_id"]
        url = (
            f"https://generativelanguage.googleapis.com/v1beta/"
            f"models/{model_id}:batchGenerateContent?key={api_key}"
        )

        req = urllib.request.Request(
            url, data=body,
            headers={"Content-Type": "application/json"},
        )

        try:
            with urllib.request.urlopen(req, timeout=600) as resp:
                result = json.loads(resp.read())

            # Pair responses with custom IDs
            responses = result.get("responses", [])
            if len(responses) != len(custom_ids):
                print(f"    WARNING: Google returned {len(responses)} responses "
                      f"for {len(custom_ids)} requests in chunk {chunk_num}. "
                      f"Missing requests will be recorded as None.")
            paired = []
            for i, cid in enumerate(custom_ids):
                paired.append({
                    "custom_id": cid,
                    "response": responses[i] if i < len(responses) else None,
                })

            chunk_path = config.RESPONSES_DIR / f"{batch_name}_chunk{chunk_num:03d}.json"
            with open(chunk_path, "w") as f:
                json.dump(paired, f, indent=2)
            chunk_results.append(str(chunk_path))
            print(f"    Saved: {chunk_path}")

        except Exception as e:
            print(f"    ERROR: {e}")
            chunk_results.append(f"ERROR: {e}")

        # Rate limiting between chunks
        if chunk_num < total_chunks:
            time.sleep(2)

    return chunk_results


def submit_anthropic(queries, model_key, strategy, api_key):
    """Submit queries to Anthropic Message Batches API."""
    import urllib.request

    model_cfg = config.MODELS[model_key]
    img_cache = ImageCache()

    # Build all requests (Anthropic handles batching server-side)
    requests_list = []
    for i, query in enumerate(queries):
        img_b64 = img_cache.get(query["image_path"])
        req = build_anthropic_request(query, strategy, model_cfg, img_b64)
        requests_list.append(req)
        if (i + 1) % 100 == 0:
            print(f"  Prepared {i + 1}/{len(queries)} requests...")
            img_cache.clear()

    body = json.dumps({"requests": requests_list}).encode()

    print(f"  Submitting {len(requests_list)} requests...")
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages/batches",
        data=body,
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
    )
    with urllib.request.urlopen(req, timeout=300) as resp:
        batch_resp = json.loads(resp.read())

    batch_id = batch_resp["id"]
    print(f"  Batch created: {batch_id}")
    return [batch_id]


def cmd_submit(args):
    """Submit batches for all active models × strategies."""
    # Safety gate: .api_lock prevents accidental API spend
    api_lock = Path(__file__).resolve().parent / ".api_lock"
    if api_lock.exists():
        print("ERROR: API calls are locked. Delete .api_lock to unlock.")
        sys.exit(1)

    # Load query index
    index_path = config.RESULTS_DIR / "query_index.json"
    if not index_path.exists():
        print("ERROR: Query index not found. Run 'prepare' first.")
        sys.exit(1)

    with open(index_path) as f:
        queries = json.load(f)

    config.RESPONSES_DIR.mkdir(parents=True, exist_ok=True)

    # Determine which models to submit. --models overrides config.ACTIVE_MODELS
    # for this run only (no file mutation). Every requested model must be
    # defined in config.MODELS; otherwise we refuse to submit so no budget
    # is spent on a typo.
    override_models = None
    if getattr(args, "models", None):
        override_models = [m.strip() for m in args.models.split(",") if m.strip()]
        bad = [m for m in override_models if m not in config.MODELS]
        if bad:
            print(f"ERROR: unknown model(s) in --models: {bad}")
            print(f"Valid models: {sorted(config.MODELS.keys())}")
            sys.exit(1)
    models_to_run = override_models if override_models is not None else config.ACTIVE_MODELS

    # Load/create tracking file
    tracking_path = config.RESULTS_DIR / "batch_tracking.json"
    tracking = {}
    if tracking_path.exists():
        with open(tracking_path) as f:
            tracking = json.load(f)

    print(f"\nSubmitting to models: {models_to_run}")
    print(f"Strategies:          {config.STRATEGIES}")
    print(f"Total API calls:     {len(queries) * len(models_to_run) * len(config.STRATEGIES)}")

    for model_key in models_to_run:
        model_cfg = config.MODELS[model_key]
        provider = model_cfg["provider"]

        try:
            api_key = config.get_api_key(provider)
        except ValueError as e:
            print(f"\nSKIP {model_key}: {e}")
            continue

        for strategy in config.STRATEGIES:
            batch_name = f"{model_key}_{strategy}"

            if batch_name in tracking and tracking[batch_name].get("status") not in ("failed",):
                print(f"\nSKIP: {batch_name} already submitted")
                continue

            print(f"\n{'='*60}")
            print(f"Submitting: {batch_name} ({provider}, {len(queries)} queries)")
            print(f"{'='*60}")

            try:
                if provider == "openai":
                    batch_ids = submit_openai(queries, model_key, strategy, api_key)
                    tracking[batch_name] = {
                        "provider": provider,
                        "batch_ids": batch_ids,
                        "status": "submitted",
                        "model": model_key,
                        "strategy": strategy,
                        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    }
                elif provider == "google":
                    chunk_paths = submit_google(queries, model_key, strategy, api_key)
                    tracking[batch_name] = {
                        "provider": provider,
                        "chunk_paths": chunk_paths,
                        "status": "completed",  # Google is synchronous
                        "model": model_key,
                        "strategy": strategy,
                        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    }
                elif provider == "anthropic":
                    batch_ids = submit_anthropic(queries, model_key, strategy, api_key)
                    tracking[batch_name] = {
                        "provider": provider,
                        "batch_ids": batch_ids,
                        "status": "submitted",
                        "model": model_key,
                        "strategy": strategy,
                        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    }
            except Exception as e:
                print(f"  FATAL ERROR: {e}")
                tracking[batch_name] = {
                    "provider": provider,
                    "status": "failed",
                    "error": str(e),
                    "model": model_key,
                    "strategy": strategy,
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                }

            # Save tracking after each submission
            with open(tracking_path, "w") as f:
                json.dump(tracking, f, indent=2)

    print(f"\nTracking saved: {tracking_path}")


# ============================================================
# Status Command
# ============================================================

def cmd_status(args):
    """Check status of all submitted batches."""
    import urllib.request

    tracking_path = config.RESULTS_DIR / "batch_tracking.json"
    if not tracking_path.exists():
        print("No batches submitted yet. Run 'submit' first.")
        return

    with open(tracking_path) as f:
        tracking = json.load(f)

    print(f"\n{'Batch':<40} {'Provider':<12} {'Status':<15} {'Details'}")
    print("-" * 90)

    for batch_name, info in tracking.items():
        provider = info["provider"]
        status = info.get("status", "unknown")

        if status in ("completed", "failed"):
            details = info.get("error", "")
            print(f"{batch_name:<40} {provider:<12} {status:<15} {details}")
            continue

        # Check OpenAI batch status
        if provider == "openai":
            try:
                api_key = config.get_api_key("openai")
            except Exception as e:
                # Bundle-level error (config / auth) — surface and skip this bundle
                print(f"{batch_name:<40} {provider:<12} {'error':<15} {e}",
                      file=sys.stderr)
                continue

            # Per-chunk progress is preserved across polls so a transient
            # failure on one chunk does not reset the others, and successful
            # chunks are not re-polled unnecessarily.
            chunk_progress = info.setdefault("chunk_progress", {})
            chunk_errors = []
            completed_count = 0
            total_count = 0

            for batch_id in info.get("batch_ids", []):
                if batch_id.startswith("ERROR"):
                    continue

                # If this chunk is already known-completed with an output file,
                # trust the previous record and skip the API call.
                prev = chunk_progress.get(batch_id, {})
                if prev.get("status") == "completed" and prev.get("output_file_id"):
                    completed_count += prev.get("completed_requests", 0)
                    total_count += prev.get("total_requests", 0)
                    continue

                def _poll_openai_batch(_bid=batch_id):
                    req = urllib.request.Request(
                        f"https://api.openai.com/v1/batches/{_bid}",
                        headers={"Authorization": f"Bearer {api_key}"},
                    )
                    with urllib.request.urlopen(req) as resp:
                        return json.loads(resp.read())

                # Per-chunk error handling: log and continue with siblings
                # rather than aborting the whole bundle's status update.
                try:
                    batch_resp = _openai_call_with_retries(
                        _poll_openai_batch,
                        _label=f"poll {batch_name}/{batch_id[:20]}")
                except Exception as e:
                    msg = f"chunk {batch_id[:20]}: {type(e).__name__}: {e}"
                    chunk_errors.append(msg)
                    print(f"{batch_name:<40} {provider:<12} {'chunk-error':<15} {msg}",
                          file=sys.stderr)
                    continue

                b_status = batch_resp.get("status", "unknown")
                counts = batch_resp.get("request_counts", {}) or {}
                cc = counts.get("completed", 0)
                tc = counts.get("total", 0)
                completed_count += cc
                total_count += tc

                # Record per-chunk progress so subsequent polls do not lose
                # this information if the next poll hits a transient failure.
                chunk_progress[batch_id] = {
                    "status": b_status,
                    "output_file_id": batch_resp.get("output_file_id"),
                    "completed_requests": cc,
                    "total_requests": tc,
                }

            # Aggregate bundle status: completed only when every active chunk
            # has been recorded as completed AND has an output_file_id.
            active_ids = [bid for bid in info.get("batch_ids", [])
                          if not bid.startswith("ERROR")]
            all_completed = active_ids and all(
                chunk_progress.get(bid, {}).get("status") == "completed"
                and chunk_progress.get(bid, {}).get("output_file_id")
                for bid in active_ids
            )
            if all_completed:
                info["status"] = "completed"
                # Preserve original chunk order for downstream download logic.
                info["output_file_ids"] = [
                    chunk_progress[bid]["output_file_id"] for bid in active_ids
                ]

            details = f"{completed_count}/{total_count} requests"
            if chunk_errors:
                details += f" ({len(chunk_errors)} chunk error(s) — will retry)"
            status_label = "completed" if all_completed else "in_progress"
            print(f"{batch_name:<40} {provider:<12} {status_label:<15} {details}")

        # Check Anthropic batch status
        elif provider == "anthropic":
            try:
                api_key = config.get_api_key("anthropic")
                for batch_id in info.get("batch_ids", []):
                    def _poll_anthropic_batch():
                        req = urllib.request.Request(
                            f"https://api.anthropic.com/v1/messages/batches/{batch_id}",
                            headers={
                                "x-api-key": api_key,
                                "anthropic-version": "2023-06-01",
                            },
                        )
                        with urllib.request.urlopen(req) as resp:
                            return json.loads(resp.read())

                    batch_resp = _openai_call_with_retries(
                        _poll_anthropic_batch,
                        _label=f"poll {batch_name}/{batch_id[:20]}")
                    b_status = batch_resp.get("processing_status", "unknown")
                    if b_status == "ended":
                        info["status"] = "ended"
                    print(f"{batch_name:<40} {provider:<12} {b_status:<15}")
            except Exception as e:
                print(f"{batch_name:<40} {provider:<12} {'error':<15} {e}")

        elif provider == "google":
            print(f"{batch_name:<40} {provider:<12} {status:<15} sync (already done)")

    with open(tracking_path, "w") as f:
        json.dump(tracking, f, indent=2)


# ============================================================
# Download Command
# ============================================================

def cmd_download(args):
    """Download results from completed batches."""
    import urllib.request

    tracking_path = config.RESULTS_DIR / "batch_tracking.json"
    if not tracking_path.exists():
        print("No batches submitted yet.")
        return

    with open(tracking_path) as f:
        tracking = json.load(f)

    config.RESPONSES_DIR.mkdir(parents=True, exist_ok=True)

    for batch_name, info in tracking.items():
        provider = info["provider"]

        if provider == "openai" and info.get("status") == "completed":
            for i, file_id in enumerate(info.get("output_file_ids", [])):
                out_path = config.RESPONSES_DIR / f"{batch_name}_chunk{i:03d}_results.jsonl"
                if out_path.exists():
                    print(f"{batch_name} chunk {i}: Already downloaded")
                    continue

                def _dl_openai_chunk():
                    api_key = config.get_api_key("openai")
                    req = urllib.request.Request(
                        f"https://api.openai.com/v1/files/{file_id}/content",
                        headers={"Authorization": f"Bearer {api_key}"},
                    )
                    with urllib.request.urlopen(req) as resp:
                        return resp.read()

                try:
                    data = _openai_call_with_retries(
                        _dl_openai_chunk,
                        _label=f"download {batch_name} chunk {i}")
                    with open(out_path, "wb") as f:
                        f.write(data)
                    print(f"{batch_name} chunk {i}: Downloaded to {out_path}")
                except Exception as e:
                    print(f"{batch_name} chunk {i}: ERROR - {e}")

        elif provider == "anthropic" and info.get("status") == "ended":
            for batch_id in info.get("batch_ids", []):
                out_path = config.RESPONSES_DIR / f"{batch_name}_results.jsonl"
                if out_path.exists():
                    print(f"{batch_name}: Already downloaded")
                    continue

                def _dl_anthropic():
                    api_key = config.get_api_key("anthropic")
                    req = urllib.request.Request(
                        f"https://api.anthropic.com/v1/messages/batches/{batch_id}/results",
                        headers={
                            "x-api-key": api_key,
                            "anthropic-version": "2023-06-01",
                        },
                    )
                    with urllib.request.urlopen(req) as resp:
                        return resp.read()

                try:
                    data = _openai_call_with_retries(
                        _dl_anthropic,
                        _label=f"download {batch_name}")
                    with open(out_path, "wb") as f:
                        f.write(data)
                    print(f"{batch_name}: Downloaded to {out_path}")
                except Exception as e:
                    print(f"{batch_name}: ERROR - {e}")

        elif provider == "google":
            print(f"{batch_name}: Results already saved during submission")

        else:
            print(f"{batch_name}: Not ready (status: {info.get('status')})")


# ============================================================
# Parse Command
# ============================================================

def parse_grid_coordinate(text):
    """
    Parse grid coordinates from model response text.

    Handles: "B3", "B 3", "b3", "B-3", "Row B, Column 3", "(B3)", "B3.", etc.

    Returns list of normalized coordinates like ["B3"] or ["C4", "D5", "E6"],
    or None if unparseable.
    """
    import re

    if not text:
        return None

    text = text.strip()

    # Pattern: letter A-H followed by number 1-16
    pattern = r'[A-Ha-h]\s*[-]?\s*(?:1[0-6]|[1-9])\b'
    matches = re.findall(pattern, text)

    if not matches:
        # Try "Row X, Column Y" format
        alt_pattern = r'[Rr]ow\s*([A-Ha-h])\s*,?\s*[Cc]ol(?:umn)?\s*(\d{1,2})'
        alt_matches = re.findall(alt_pattern, text)
        if alt_matches:
            matches = [f"{r}{c}" for r, c in alt_matches]

    if not matches:
        return None

    # Normalize: uppercase, no spaces/dashes
    normalized = []
    for m in matches:
        clean = re.sub(r'[\s\-]', '', m).upper()
        normalized.append(clean)

    # Deduplicate while preserving order
    seen = set()
    unique = []
    for c in normalized:
        if c not in seen:
            seen.add(c)
            unique.append(c)

    return unique


_STRONG_REFUSAL_MARKERS = (
    # Require refusal + action object to avoid matching hedged phrases like
    # "I cannot be sure" (which is ambiguous, not refusal).
    "cannot identify", "can't identify",
    "cannot determine", "can't determine",
    "cannot provide", "can't provide",
    "cannot assist", "can't assist",
    "cannot help", "can't help",
    "cannot locate", "can't locate",
    "cannot answer", "can't answer",
    "cannot give", "can't give",
    "unable to identify", "unable to determine", "unable to provide",
    "unable to assist", "unable to help", "unable to locate", "unable to answer",
    "not able to identify", "not able to determine", "not able to provide",
    "as an ai", "as a language model",
    "i apologize", "i'm sorry", "sorry, i",
    "not qualified", "medical advice",
    "i refuse", "i must decline", "decline to",
)

_AMBIGUOUS_MARKERS = (
    "either", "or possibly", "might be", "could be", "possibly",
    "probably", "likely", "perhaps", "approximately",
    "not sure", "unclear", "hard to tell", "difficult to determine",
    "multiple possible", "several possible", "any of", "one of",
    "without more", "without additional", "cannot be sure", "can't be sure",
    "cannot be certain", "can't be certain", "it's hard", "unclear which",
)


def categorise_unparseable_response(text):
    """
    Classify a response that failed coordinate parsing (or was dropped as
    invalid) into one of four failure modes. Rule-based, deterministic,
    no LLM calls.

    Returns one of:
      - "no_engage" : empty / whitespace / near-empty response
      - "ambiguous" : hedged with multiple candidates or non-committal
                      language (checked BEFORE refusal to avoid the
                      "I cannot be sure" false-positive)
      - "refusal"   : explicit refusal or safety disclaimer
      - "verbose"   : substantive answer but no valid grid coordinates
    """
    if text is None:
        return "no_engage"
    stripped = text.strip()
    if len(stripped) < 5:
        return "no_engage"

    lower = stripped.lower()

    # Ambiguity is checked first: hedged refusal phrases like "I cannot
    # be sure" should land in "ambiguous", not "refusal".
    if any(m in lower for m in _AMBIGUOUS_MARKERS):
        return "ambiguous"
    if any(m in lower for m in _STRONG_REFUSAL_MARKERS):
        return "refusal"
    return "verbose"


def compute_compliance_stats(records):
    """
    Compute instruction compliance rate and failure-mode breakdown from a
    list of parsed response records.

    A record is considered compliant only if parsed_coordinates is non-empty
    AND failure_category is None (i.e., no out-of-range / ambiguous override).
    """
    total = len(records)
    parsed_ok = 0
    categories = {
        "refusal": 0, "verbose": 0, "ambiguous": 0,
        "no_engage": 0, "out_of_range": 0,
    }

    for r in records:
        if r.get("parsed_coordinates") and r.get("failure_category") is None:
            parsed_ok += 1
        else:
            cat = r.get("failure_category")
            if cat in categories:
                categories[cat] += 1

    compliance_rate = (parsed_ok / total) if total else 0.0
    return {
        "total": total,
        "compliant": parsed_ok,
        "non_compliant": total - parsed_ok,
        "compliance_rate": compliance_rate,
        "failure_modes": categories,
    }


def validate_coordinate(coord, modality):
    """
    Check if a coordinate is within the valid grid range for a modality.
    Returns True if valid, False if out of range or malformed.
    """
    import re
    if coord is None:
        return False
    coord = str(coord).strip().upper()
    match = re.match(r'^([A-H])(\d{1,2})$', coord)
    if not match:
        return False
    row_letter = match.group(1)
    col_num = int(match.group(2))
    grid = config.GRID_SPECS.get(modality)
    if not grid:
        return False
    max_row = grid["max_row_letter"]
    max_col = grid["cols"]
    # Explicit membership check — do not rely on ASCII lex order alone.
    valid_rows = "".join(chr(ord('A') + i) for i in range(ord(max_row) - ord('A') + 1))
    return row_letter in valid_rows and 1 <= col_num <= max_col


def parse_response_filename(name):
    """
    Identify (model_key, strategy) from a response filename.

    Expected patterns (all start with `{model_key}_{strategy}`):
      OpenAI    : {model_key}_{strategy}_chunk{N:03d}_results.jsonl
      Google    : {model_key}_{strategy}_chunk{N:03d}.json
      Anthropic : {model_key}_{strategy}_results.jsonl

    Returns (model_key, strategy) or (None, None) if no pattern matches.
    """
    # Longest model_key and strategy first so e.g. "claude-sonnet-4.6" is
    # preferred over any shorter prefix.
    for mk in sorted(config.MODELS.keys(), key=len, reverse=True):
        for strat in sorted(config.STRATEGIES, key=len, reverse=True):
            prefix = f"{mk}_{strat}"
            if name == prefix or name.startswith(prefix + "_") or name.startswith(prefix + "."):
                return mk, strat
    return None, None


def _finalise_record(record, query_modality_map, query_landmark_map):
    """
    Validate parsed_coordinates in-place against the modality grid and enforce
    the single-cell rule for point landmarks. Mutates `record` and returns it.

    Rules (in order):
      1. If parsed_coordinates is None/empty, derive a failure_category from
         the raw response text (refusal / ambiguous / verbose / no_engage).
      2. If the modality is unknown, preserve the parse and flag nothing
         (this shouldn't happen in production — query_index always has the
         sheet name — but is preserved for defensive testing).
      3. POINT LANDMARK: the prompt asks for "exactly one cell coordinate".
         Any response that yields more than one cell — even if some happen
         to be out of grid range — is treated as "ambiguous". The model
         violated the instruction, full stop. This is the strict scientific
         interpretation: we don't second-guess which of the multiple cells
         was the model's "real" answer.
         A single coord that is out of range is "out_of_range".
      4. AREA LANDMARK: drop out-of-range cells silently and keep the valid
         subset. If no valid cells remain, classify as "out_of_range".

    Out-of-range cells (dropped or rejected) are preserved on the record in
    `record["out_of_range"]` for auditability.
    """
    query_id = record["query_id"]
    modality = query_modality_map.get(query_id)
    lm_type = query_landmark_map.get(query_id)
    record["modality"] = modality
    record["landmark_type"] = lm_type

    raw = record.get("raw_response")
    parsed = record.get("parsed_coordinates")

    # Branch 1: nothing parsed
    if not parsed:
        record["parsed_coordinates"] = None
        record["out_of_range"] = []
        record["failure_category"] = categorise_unparseable_response(raw)
        return record

    # Branch 2: modality unknown -> defensive passthrough
    if modality is None:
        record["out_of_range"] = []
        record["failure_category"] = None
        return record

    valid = [c for c in parsed if validate_coordinate(c, modality)]
    invalid = [c for c in parsed if not validate_coordinate(c, modality)]
    record["out_of_range"] = invalid

    # Branch 3: point landmark — strict single-cell rule
    if lm_type == "point":
        # Multi-coord response (any combination of valid+invalid) is
        # an instruction violation -> ambiguous.
        if len(parsed) > 1:
            record["parsed_coordinates"] = None
            record["failure_category"] = "ambiguous"
            return record
        # Exactly one parsed cell — succeeds iff it's in range.
        if not valid:
            record["parsed_coordinates"] = None
            record["failure_category"] = "out_of_range"
            return record
        record["parsed_coordinates"] = valid
        record["failure_category"] = None
        return record

    # Branch 4: area landmark — drop invalid cells, keep the valid subset
    if not valid:
        record["parsed_coordinates"] = None
        record["failure_category"] = "out_of_range"
        return record
    record["parsed_coordinates"] = valid
    record["failure_category"] = None
    return record


def cmd_parse(args):
    """
    Parse all downloaded responses into a single list of per-query records.

    Output file: results/parsed_responses.json

    Output format: a JSON list, one record per (model, strategy, query_id).
    Each record has explicit fields so downstream consumers never need to
    parse custom_ids or filenames:
        {
            "query_id": "PAN_001_Mental_Foramen_L",
            "strategy": "zero_shot",
            "model_key": "gpt-5.4",
            "modality": "PANORAMIC",
            "landmark_type": "point",
            "custom_id": "PAN_001_Mental_Foramen_L_zero_shot",
            "raw_response": "...",
            "parsed_coordinates": ["B3"],        # or None
            "out_of_range": [],                  # dropped cells
            "failure_category": null,            # or a string
            "source_file": "gpt-5.4_zero_shot_chunk000_results.jsonl",
        }

    This format fixes the prior custom_id-collision bug where responses from
    different providers overwrote each other in a dict keyed by custom_id.
    """
    # Build query_id -> modality and query_id -> landmark_type maps from the
    # query index. These are needed for validation and for M1 (single-cell
    # rule on point landmarks).
    index_path = config.RESULTS_DIR / "query_index.json"
    if not index_path.exists():
        print("ERROR: query_index.json not found. Run 'prepare' first.")
        sys.exit(1)
    with open(index_path) as f:
        queries = json.load(f)
    query_modality_map = {q["query_id"]: q["sheet"] for q in queries}
    query_landmark_map = {q["query_id"]: q["landmark_type"] for q in queries}

    records = []

    def add_record(model_key, strategy, custom_id, raw_text, source_name):
        # Strip off the strategy suffix from custom_id to recover query_id
        # (custom_id is built as "{query_id}_{strategy}" by the request builders).
        suffix = f"_{strategy}"
        query_id = custom_id[:-len(suffix)] if custom_id.endswith(suffix) else custom_id
        parsed = parse_grid_coordinate(raw_text)
        record = {
            "query_id": query_id,
            "strategy": strategy,
            "model_key": model_key,
            "custom_id": custom_id,
            "raw_response": raw_text,
            "parsed_coordinates": parsed,
            "source_file": source_name,
        }
        _finalise_record(record, query_modality_map, query_landmark_map)
        records.append(record)

    # ----- OpenAI / Anthropic JSONL files (sorted for determinism) -----
    for resp_file in sorted(config.RESPONSES_DIR.glob("*_results.jsonl")):
        model_key, strategy = parse_response_filename(resp_file.name)
        if model_key is None or strategy is None:
            print(f"WARNING: skipping unrecognised file {resp_file.name}")
            continue
        with open(resp_file) as f:
            for line in f:
                if not line.strip():
                    continue
                entry = json.loads(line)
                custom_id = entry.get("custom_id", "")

                response_text = None
                if "response" in entry and "body" in entry["response"]:
                    # OpenAI format
                    body = entry["response"]["body"]
                    choices = body.get("choices", [])
                    if choices:
                        response_text = choices[0].get("message", {}).get("content", "")
                elif "result" in entry:
                    # Anthropic format — guard against non-text content blocks.
                    result = entry["result"]
                    if result.get("type") == "succeeded":
                        content = result.get("message", {}).get("content", [])
                        if content and content[0].get("type") == "text":
                            response_text = content[0].get("text", "")

                add_record(model_key, strategy, custom_id, response_text, resp_file.name)

    # ----- Google JSON chunk files (sorted for determinism) -----
    for resp_file in sorted(config.RESPONSES_DIR.glob("*_chunk*.json")):
        model_key, strategy = parse_response_filename(resp_file.name)
        if model_key is None or strategy is None:
            print(f"WARNING: skipping unrecognised file {resp_file.name}")
            continue
        with open(resp_file) as f:
            entries = json.load(f)
        for entry in entries:
            custom_id = entry.get("custom_id", "")
            response = entry.get("response", {})
            response_text = None
            candidates = response.get("candidates", []) if response else []
            if candidates:
                parts = candidates[0].get("content", {}).get("parts", [])
                if parts:
                    response_text = parts[0].get("text", "")
            add_record(model_key, strategy, custom_id, response_text, resp_file.name)

    # Save unified results as a list of records (no more dict-keyed overwrites)
    out_path = config.RESULTS_DIR / "parsed_responses.json"
    with open(out_path, "w") as f:
        json.dump(records, f, indent=2)

    # Compute instruction compliance stats (Issues 5 + 8)
    compliance = compute_compliance_stats(records)
    compliance_path = config.RESULTS_DIR / "compliance_stats.json"
    with open(compliance_path, "w") as f:
        json.dump(compliance, f, indent=2)

    # Per-(model, strategy) compliance breakdown
    from collections import defaultdict
    per_batch = defaultdict(list)
    for r in records:
        per_batch[(r["model_key"], r["strategy"])].append(r)

    total = compliance["total"]
    parsed_ok = compliance["compliant"]
    failed = compliance["non_compliant"]
    print(f"\nParsed {total} responses: {parsed_ok} OK, {failed} non-compliant")
    print(f"Instruction compliance rate: {compliance['compliance_rate']:.1%}")
    fm = compliance["failure_modes"]
    print(f"Failure modes: refusal={fm['refusal']}, verbose={fm['verbose']}, "
          f"ambiguous={fm['ambiguous']}, no_engage={fm['no_engage']}, "
          f"out_of_range={fm['out_of_range']}")

    if per_batch:
        print("\nPer-batch compliance:")
        for (mk, strat), rs in sorted(per_batch.items()):
            c = compute_compliance_stats(rs)
            print(f"  {mk} / {strat}: {c['compliant']}/{c['total']} "
                  f"({c['compliance_rate']:.1%})")

    if failed > 0:
        print("\nExample non-compliant entries (first 10):")
        shown = 0
        for r in records:
            if r.get("parsed_coordinates") and r.get("failure_category") is None:
                continue
            raw = (r.get("raw_response") or "")[:100]
            print(f"  [{r['failure_category']}] {r['custom_id']} "
                  f"({r['model_key']}): {raw!r}")
            shown += 1
            if shown >= 10:
                break

    print(f"\nResults saved: {out_path}")


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Dental MLLM Benchmark Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Commands:
  prepare   Parse Excel, validate images, save query index (no API key needed)
  submit    Submit batches to APIs (requires API keys)
  status    Check batch processing status
  download  Download completed batch results
  parse     Parse responses into unified results file
        """,
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    p_prepare = subparsers.add_parser("prepare", help="Parse Excel, validate images, save query index")
    p_prepare.add_argument("--n-per-modality", type=int, default=None,
                           help="Keep only N deterministically-spaced images per modality (for pilot runs)")
    p_prepare.add_argument("--subset-images", type=str, default=None,
                           help="Explicit comma-separated image_ids to keep (overrides --n-per-modality)")
    p_prepare.add_argument("--models", type=str, default=None,
                           help="Comma-separated model_keys to include in the summary (informational)")

    p_submit = subparsers.add_parser("submit", help="Submit batches to APIs")
    p_submit.add_argument("--models", type=str, default=None,
                          help="Comma-separated model_keys to submit (overrides config.ACTIVE_MODELS)")

    subparsers.add_parser("status", help="Check batch processing status")
    subparsers.add_parser("download", help="Download completed batch results")
    subparsers.add_parser("parse", help="Parse responses into unified results file")

    args = parser.parse_args()

    commands = {
        "prepare": cmd_prepare,
        "submit": cmd_submit,
        "status": cmd_status,
        "download": cmd_download,
        "parse": cmd_parse,
    }
    commands[args.command](args)


def _load_dotenv_at_root() -> None:
    """Load .env from project root if present (no python-dotenv dependency)."""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(env_path):
        return
    with open(env_path) as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("export "):
                line = line[len("export "):]
            if "=" not in line:
                continue
            key, _, val = line.partition("=")
            key = key.strip()
            val = val.strip()
            if len(val) >= 2 and val[0] == val[-1] and val[0] in ("'", '"'):
                val = val[1:-1]
            if key and key not in os.environ:
                os.environ[key] = val


if __name__ == "__main__":
    _load_dotenv_at_root()
    main()
