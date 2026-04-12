# Dental MLLM Benchmark — Setup & Cost Guide

## 1. Install Dependencies

```bash
pip3 install -r requirements.txt
```

Required packages: `openpyxl` (Excel parsing), `scipy` (statistical tests), `matplotlib` (plots).

---

## 2. API Key Setup

### OpenAI (for GPT-5.4)

1. Go to **https://platform.openai.com/api-keys**
2. Sign in (or create account)
3. Click **"Create new secret key"**
4. Copy the key (starts with `sk-...`)
5. **Add credits**: Go to **Settings → Billing → Add payment method**, then add credits
   - Estimated need: **~$10-15** for the full experiment
   - Credits are prepaid; you only pay for what you use
6. **Enable Batch API**: No special setup needed — batch API is available by default

### Google (for Gemini 3.1 Pro)

1. Go to **https://aistudio.google.com/apikey**
2. Sign in with Google account
3. Click **"Create API Key"**
4. Select or create a Google Cloud project
5. Copy the key
6. **Billing**: Go to **https://console.cloud.google.com/billing**
   - Link a billing account to your project
   - Estimated need: **~$8-12** for the full experiment
   - Google has a free tier (limited requests/day) — for 900 queries, you'll likely need paid tier
7. **Batch API**: Available through the Gemini Developer API by default

### Anthropic (optional, for Claude Sonnet 4.6)

1. Go to **https://console.anthropic.com/settings/keys**
2. Sign in (or create account)
3. Click **"Create Key"**
4. Copy the key (starts with `sk-ant-...`)
5. **Add credits**: Go to **Settings → Plans & Billing → Add credits**
   - Estimated need: **~$5-8**
6. **Message Batches API**: Available by default

### Set Environment Variables

```bash
# Add to your shell profile (~/.zshrc or ~/.bashrc) or run before each session:
export OPENAI_API_KEY="sk-..."
export GOOGLE_API_KEY="AIza..."
export ANTHROPIC_API_KEY="sk-ant-..."   # optional
```

---

## 3. Final Cost Estimates

### Per-model cost breakdown (using Batch API pricing, 50% off)

| Model | Provider | Batch Input/1M | Batch Output/1M | Est. per strategy (900q) | Est. 2 strategies |
|-------|----------|---------------|-----------------|-------------------------|-------------------|
| GPT-5.4 | OpenAI | $1.25 | $7.50 | ~$3-5 | ~$6-10 |
| Gemini 3.1 Pro | Google | $1.00 | $6.00 | ~$3-5 | ~$6-10 |
| Claude Sonnet 4.6 | Anthropic | $1.50 | $7.50 | ~$3-5 | ~$5-8 |

### Token estimate per query

- **Image tokens**: ~1,000-2,500 (varies by modality and provider)
- **Text input tokens**: ~150 (zero-shot) to ~250 (guided)
- **Output tokens**: ~5-20 (just grid coordinates)
- Total input per query: ~1,200-2,700 tokens

### Total cost scenarios

| Scenario | Models | Total Cost |
|----------|--------|-----------|
| **Core (2 models)** | GPT-5.4 + Gemini 3.1 Pro | **~$10-15** |
| **Recommended (3 models)** | + Claude Sonnet 4.6 | **~$15-20** |

### Important notes

- **Batch API turnaround**: OpenAI batches complete within 24 hours. Google batch is synchronous (immediate). Anthropic batches complete within 24 hours.
- **Rate limits**: Batch APIs handle rate limiting automatically.
- **No subscription needed**: All providers use prepaid credits (pay-as-you-go).
- **Unused credits**: Remain in your account for future use.

---

## 4. Running the Pipeline

```bash
# Step 1: Prepare (validates data, no API key needed)
python3 pipeline.py prepare

# Step 2: Submit batches (requires API keys set as env vars)
python3 pipeline.py submit

# Step 3: Check status (for OpenAI/Anthropic async batches)
python3 pipeline.py status

# Step 4: Download results when batches complete
python3 pipeline.py download

# Step 5: Parse all responses
python3 pipeline.py parse

# Step 6: Run full analysis
python3 analysis.py all
```

---

## 5. File Structure After Running

```
results/
  query_index.json          # 900 queries with metadata
  batch_tracking.json       # Batch IDs and status
  parsed_responses.json     # All parsed model responses
  responses/                # Raw API responses
  analysis/
    all_metrics.json        # ED, NED, Jaccard, Dice per query
    statistical_results.json # Shapiro-Wilk, Kruskal-Wallis, pairwise tests
    summary_table.csv       # Aggregated metrics table
    detailed_results.csv    # Per-query results
    plots/
      boxplot_ed_by_model.png
      sdr_bar_chart.png
      heatmap_ed_by_landmark.png
      scatter_jaccard_dice.png
      boxplot_ed_by_modality.png
```

---

## 6. Known Issues

- **Bug A (fixed in 933b3d4):** GPT-5.4 rejects the `max_tokens` parameter; the pipeline now uses `max_completion_tokens` for OpenAI requests.
- **Bug E (pending fix):** Prompt template examples overlap with 10.1% of ground truths (91/900 queries). Fixed example coordinates (e.g., `B3`, `C4, D5, E6`) in `ZERO_SHOT_POINT_TEMPLATE`, `ZERO_SHOT_AREA_TEMPLATE`, and `GUIDED_SYSTEM_ADDITION` coincidentally match ground-truth values for some queries, creating a potential bias. A proposed fix replaces concrete examples with abstract format descriptions. Pending colleague approval.
- **Note:** All prompt template changes require colleague sign-off before deployment.
