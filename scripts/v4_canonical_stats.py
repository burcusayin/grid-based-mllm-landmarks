"""Canonical statistical computations for v4 paper revision.

Sources:
  - results_consensus/full_run_records.pkl   (GPT-5.4 — per-query mean ED, with
                                              student/OMFR alignment used in
                                              paper's existing analysis)
  - results_full_gemini/full_run_records.pkl (Gemini 3.1 Pro — same schema)
  - results_full/run{1,2,3}/parsed_responses.json (GPT raw — for Table 2 SDR)
  - results_full_gemini/run{1,2,3}/parsed_responses.json (Gemini raw — for
                                              Table 2 SDR and Figure 4 cell
                                              distributions)

Outputs (results_v4_canonical.json):
  - Table 2 (model accuracy) with bootstrap CIs on mean ED + SDR@0/1/2
  - Table 3 (per-landmark zero-shot vs guided GPT) with W + bootstrap CIs
  - Table 4 (model vs student) with W + bootstrap CIs
  - Table 5 (GPT vs Gemini) with W + bootstrap CIs
  - Figure 4 cell prediction distributions (both models)

Convention notes:
  - Pickle uses canonical "letter-digit" parser, so 1 PAN point query with a
    reversed-format student response ('3A' for GT 'A3') is excluded from
    paired model-vs-student analyses (n=299 not 300 for PAN point). This
    matches the colleague's existing analysis files. The Table 2 raw-data
    SDR analysis uses the robust parser so all 300 queries are included
    there.
  - Δ sign convention is documented in the output _meta block.
"""
from __future__ import annotations
import json
import math
import pickle
from pathlib import Path
from collections import Counter

import numpy as np
from scipy import stats

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results_v4_canonical.json"

RNG_SEED = 42
N_BOOT_MEAN = 10_000
N_BOOT_R    = 2_000

POINT_LANDMARKS = {
    "PANORAMIC": ["Mental_Foramen_L", "Condylar_Head_R", "Tooth_33_Apex"],
    "PERIAPICAL": ["Tooth_36_Distal_Apex", "Tooth_36_Distal_CEJ", "Tooth_36_Mesial_CEJ"],
    "CEPHALOMETRIC": ["Sella_S", "Nasion_N", "Menton_Me"],
}
AREA_LANDMARKS = {
    "PANORAMIC": ["Mandibular_Canal_L", "Maxillary_Sinus_R", "External_Oblique_Ridge_R"],
}


def to_rc(c: str) -> tuple[int, int]:
    """Parse a grid cell. Accepts canonical 'A3' or reversed '3A' (rare student typo)."""
    c = c.strip()
    if c[0].isalpha():
        return ord(c[0].upper()) - ord("A"), int(c[1:])
    # reversed: digit(s)-then-letter
    i = 0
    while i < len(c) and c[i].isdigit():
        i += 1
    if i == 0 or i >= len(c):
        raise ValueError(f"unparseable cell: {c!r}")
    return ord(c[i].upper()) - ord("A"), int(c[:i])


def euclid(a: str, b: str) -> float:
    r1, c1 = to_rc(a)
    r2, c2 = to_rc(b)
    return math.sqrt((r1 - r2) ** 2 + (c1 - c2) ** 2)


def parse_cells(s, modality):
    if not s or not isinstance(s, str):
        return []
    return [t.strip() for t in s.split(",") if t.strip()]


def jaccard(a: set, b: set) -> float:
    if not a and not b:
        return 1.0
    return len(a & b) / len(a | b) if (a | b) else 0.0


def load_consensus_and_student():
    qi = json.loads((ROOT / "results_consensus" / "query_index.json").read_text())
    return qi


def load_raw_responses(root_dir: str):
    """Returns dict: {(query_id, strategy): [coord_str_per_rep]}."""
    out = {}
    for rep in (1, 2, 3):
        path = ROOT / root_dir / f"run{rep}" / "parsed_responses.json"
        data = json.loads(path.read_text())
        for entry in data:
            qid = entry.get("query_id", "")
            strat = entry.get("strategy")
            coords = entry.get("parsed_coordinates") or []
            key = (qid, strat)
            out.setdefault(key, []).append(coords)
    return out


def build_per_query_metrics(qi, raw, *, model_label):
    """Build per-query mean ED (point) or per-query mean Jaccard (area)."""
    by_strat = {"zero_shot": {}, "guided": {}}
    for q in qi:
        qid = q["query_id"]
        mod = q["sheet"]
        landmark = q.get("landmark_type")
        gt = q.get("consensus_gt")
        if not gt:
            continue

        for strat in ("zero_shot", "guided"):
            reps = raw.get((qid, strat), [])
            if landmark == "point":
                gt_cell = parse_cells(gt, mod)
                if not gt_cell:
                    continue
                gt_cell = gt_cell[0]
                eds = []
                for r in reps:
                    if r:
                        try:
                            eds.append(euclid(r[0], gt_cell))
                        except Exception:
                            pass
                if eds:
                    by_strat[strat][qid] = {
                        "metric": "ed",
                        "value": sum(eds) / len(eds),
                        "n_reps": len(eds),
                        "mod": mod,
                        "structure": q["structure"],
                    }
            elif landmark == "area":
                gt_set = set(parse_cells(gt, mod))
                if not gt_set:
                    continue
                jaccs = []
                for r in reps:
                    if r:
                        try:
                            jaccs.append(jaccard(set(r), gt_set))
                        except Exception:
                            pass
                if jaccs:
                    by_strat[strat][qid] = {
                        "metric": "jaccard",
                        "value": sum(jaccs) / len(jaccs),
                        "n_reps": len(jaccs),
                        "mod": mod,
                        "structure": q["structure"],
                    }
    return by_strat


def build_student_metrics(qi):
    by_query = {}
    for q in qi:
        qid = q["query_id"]
        mod = q["sheet"]
        landmark = q.get("landmark_type")
        gt = q.get("consensus_gt")
        stu = q.get("student")
        if not gt or not stu:
            continue
        if landmark == "point":
            gt_cell = parse_cells(gt, mod)
            stu_cell = parse_cells(stu, mod)
            if not gt_cell or not stu_cell:
                continue
            by_query[qid] = {
                "metric": "ed",
                "value": euclid(stu_cell[0], gt_cell[0]),
                "mod": mod,
                "structure": q["structure"],
            }
        elif landmark == "area":
            gt_set = set(parse_cells(gt, mod))
            stu_set = set(parse_cells(stu, mod))
            if not gt_set or not stu_set:
                continue
            by_query[qid] = {
                "metric": "jaccard",
                "value": jaccard(stu_set, gt_set),
                "mod": mod,
                "structure": q["structure"],
            }
    return by_query


# ── statistical helpers ────────────────────────────────────────────
def paired_wilcoxon(x: list[float], y: list[float]) -> dict:
    """Paired Wilcoxon signed-rank on (x − y).

    Returns:
        W      : sum of positive ranks (W+)
        p      : two-sided p-value
        r_rb   : rank-biserial r = (W+ − W−) / T,  T = n_nz(n_nz+1)/2
                 (positive r → mean(x) > mean(y), i.e. y is the "winner"
                  when smaller-is-better, x is the "winner" when larger-
                  is-better; sign depends on input order, NOT on metric).
    """
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    d = x - y
    n_total = len(d)
    nz_mask = d != 0
    n_nonzero = int(nz_mask.sum())
    if n_nonzero == 0:
        return {
            "n_total": n_total,
            "n_nonzero": 0,
            "W": 0.0,
            "p": 1.0,
            "mean_delta": float(np.mean(d)),
            "median_delta": float(np.median(d)),
            "rank_biserial_r": 0.0,
        }
    # Compute W+ manually using midrank for ties (scipy default zero_method='wilcox')
    abs_d = np.abs(d[nz_mask])
    ranks = stats.rankdata(abs_d, method="average")
    signs = np.sign(d[nz_mask])
    W_plus = float(np.sum(ranks[signs > 0]))
    W_minus = float(np.sum(ranks[signs < 0]))
    T = n_nonzero * (n_nonzero + 1) / 2.0
    r_rb = (W_plus - W_minus) / T if T > 0 else 0.0
    # Two-sided p from scipy
    res = stats.wilcoxon(x, y, zero_method="wilcox", alternative="two-sided",
                         method="auto")
    return {
        "n_total": n_total,
        "n_nonzero": n_nonzero,
        "W": W_plus,
        "p": float(res.pvalue),
        "mean_delta": float(np.mean(d)),
        "median_delta": float(np.median(d)),
        "rank_biserial_r": float(r_rb),
    }


def bootstrap_ci_mean_delta(x: list[float], y: list[float],
                            n_boot: int = N_BOOT_MEAN,
                            seed: int = RNG_SEED) -> tuple[float, float]:
    rng = np.random.default_rng(seed)
    d = np.asarray(x, dtype=float) - np.asarray(y, dtype=float)
    n = len(d)
    means = np.empty(n_boot, dtype=float)
    for i in range(n_boot):
        idx = rng.integers(0, n, n)
        means[i] = d[idx].mean()
    return float(np.percentile(means, 2.5)), float(np.percentile(means, 97.5))


def bootstrap_ci_rank_biserial(x: list[float], y: list[float],
                               n_boot: int = N_BOOT_R,
                               seed: int = RNG_SEED) -> tuple[float, float]:
    rng = np.random.default_rng(seed)
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    n = len(x)
    rs = np.empty(n_boot, dtype=float)
    for i in range(n_boot):
        idx = rng.integers(0, n, n)
        d = x[idx] - y[idx]
        nz_mask = d != 0
        n_nz = int(nz_mask.sum())
        if n_nz == 0:
            rs[i] = 0.0
            continue
        abs_d = np.abs(d[nz_mask])
        ranks = stats.rankdata(abs_d, method="average")
        signs = np.sign(d[nz_mask])
        W_plus = float(ranks[signs > 0].sum())
        W_minus = float(ranks[signs < 0].sum())
        T = n_nz * (n_nz + 1) / 2.0
        rs[i] = (W_plus - W_minus) / T if T > 0 else 0.0
    return float(np.percentile(rs, 2.5)), float(np.percentile(rs, 97.5))


def bootstrap_ci_mean(values: list[float], n_boot: int = N_BOOT_MEAN,
                      seed: int = RNG_SEED) -> tuple[float, float]:
    rng = np.random.default_rng(seed)
    arr = np.asarray(values, dtype=float)
    n = len(arr)
    means = np.empty(n_boot, dtype=float)
    for i in range(n_boot):
        idx = rng.integers(0, n, n)
        means[i] = arr[idx].mean()
    return float(np.percentile(means, 2.5)), float(np.percentile(means, 97.5))


# ── pull data ──────────────────────────────────────────────────────
print("Loading raw data ...")
qi = load_consensus_and_student()
gpt_raw = load_raw_responses("results_full")
gem_raw = load_raw_responses("results_full_gemini")

print("Building per-query metrics (raw, robust parser — for Table 2 SDR) ...")
gpt_q = build_per_query_metrics(qi, gpt_raw, model_label="GPT-5.4")
gem_q = build_per_query_metrics(qi, gem_raw, model_label="Gemini 3.1 Pro")
stu_q = build_student_metrics(qi)


def load_pickle_metrics(pkl_path: str):
    """Build per-query mean model metric (vs consensus_gt) and student-to-GT
    metric, both keyed by (qid, strat).

    Note: in the pickle's `metrics` dict, each key (consensus_gt / omfr_1 /
    omfr_2 / student) gives the MODEL's per-rep ED measured *against that
    reference*. So `metrics['student']['mean_ed']` is NOT student-vs-GT
    distance — it is model-vs-student distance. To get the student-vs-GT
    distance we recompute from r['student'] and r['consensus_gt'].
    """
    out = {}
    records = pickle.load(open(pkl_path, "rb"))
    for r in records:
        qid = r["query_id"]
        strat = r["strategy"]
        lm = r["landmark_type"]
        mod = r["modality"]
        gt_str = r.get("consensus_gt")
        stu_str = r.get("student")
        if lm == "point":
            metric = "ed"
            model_v = r["metrics"]["consensus_gt"]["mean_ed"]
            stu_v = None
            if gt_str and stu_str:
                try:
                    gt_cells = parse_cells(gt_str, mod)
                    stu_cells = parse_cells(stu_str, mod)
                    if gt_cells and stu_cells:
                        stu_v = euclid(stu_cells[0], gt_cells[0])
                except Exception:
                    pass
        else:
            metric = "jaccard"
            model_v = r["metrics"]["consensus_gt"]["mean_jaccard"]
            stu_v = None
            if gt_str and stu_str:
                gt_set = set(parse_cells(gt_str, mod))
                stu_set = set(parse_cells(stu_str, mod))
                if gt_set and stu_set:
                    stu_v = jaccard(stu_set, gt_set)
        out[(qid, strat)] = {
            "model_val": model_v,
            "student_val": stu_v,
            "modality": mod,
            "structure": r["structure"],
            "landmark_type": lm,
            "metric": metric,
        }
    return out

print("Loading pickle metrics (for Table 3/4/5, matches paper's pairing) ...")
gpt_pkl = load_pickle_metrics("results_consensus/full_run_records.pkl")
gem_pkl = load_pickle_metrics("results_full_gemini/full_run_records.pkl")


# ── Table 2: modality-stratified accuracy ──────────────────────────
def sdr_at(values: list[float], thr: float) -> float:
    if not values: return 0.0
    return 100.0 * sum(1 for v in values if v <= thr) / len(values)


def table2_cells(model_q: dict, model_name: str):
    rows = []
    for mod in ("CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"):
        for strat in ("zero_shot", "guided"):
            entries = [v["value"] for k, v in model_q[strat].items()
                       if v["mod"] == mod and v["metric"] == "ed"]
            entries.sort()
            if not entries:
                continue
            mean = sum(entries) / len(entries)
            ci_low, ci_high = bootstrap_ci_mean(entries)
            median = float(np.median(entries))
            rows.append({
                "model": model_name,
                "modality": mod,
                "strategy": strat,
                "n": len(entries),
                "mean": mean,
                "mean_ci": [ci_low, ci_high],
                "median": median,
                "SDR_0": sdr_at(entries, 0.0),
                "SDR_1": sdr_at(entries, 1.0),
                "SDR_2": sdr_at(entries, 2.0),
            })
    return rows

print("Computing Table 2 ...")
table2 = table2_cells(gpt_q, "GPT-5.4") + table2_cells(gem_q, "Gemini 3.1 Pro")


# ── Table 3: per-landmark GPT zero-shot vs guided (pickle source) ───
print("Computing Table 3 (GPT zero-shot vs guided per landmark) ...")
table3 = []
landmark_order = [
    ("PANORAMIC", "Mental_Foramen_L", "point"),
    ("PANORAMIC", "Condylar_Head_R", "point"),
    ("PANORAMIC", "Tooth_33_Apex", "point"),
    ("CEPHALOMETRIC", "Sella_S", "point"),
    ("CEPHALOMETRIC", "Nasion_N", "point"),
    ("CEPHALOMETRIC", "Menton_Me", "point"),
    ("PERIAPICAL", "Tooth_36_Distal_Apex", "point"),
    ("PERIAPICAL", "Tooth_36_Distal_CEJ", "point"),
    ("PERIAPICAL", "Tooth_36_Mesial_CEJ", "point"),
    ("PANORAMIC", "External_Oblique_Ridge_R", "area"),
    ("PANORAMIC", "Mandibular_Canal_L", "area"),
    ("PANORAMIC", "Maxillary_Sinus_R", "area"),
]
for mod, lm, lm_type in landmark_order:
    zs_vals = []
    g_vals = []
    for (qid, strat), entry in gpt_pkl.items():
        if entry["modality"] != mod or entry["structure"] != lm:
            continue
        # Need both strategies present and non-null
        zs_e = gpt_pkl.get((qid, "zero_shot"))
        g_e  = gpt_pkl.get((qid, "guided"))
        if not (zs_e and g_e):
            continue
        if zs_e["model_val"] is None or g_e["model_val"] is None:
            continue
        if (qid, "zero_shot") == (qid, strat):
            zs_vals.append(zs_e["model_val"])
            g_vals.append(g_e["model_val"])
    if not zs_vals:
        continue
    if lm_type == "point":
        # Δ = zero-shot - guided (positive = guided better)
        res = paired_wilcoxon(zs_vals, g_vals)
        ci_d = bootstrap_ci_mean_delta(zs_vals, g_vals)
        ci_r = bootstrap_ci_rank_biserial(zs_vals, g_vals)
    else:
        # Δ = guided - zero-shot (positive = guided produces higher Jaccard)
        res = paired_wilcoxon(g_vals, zs_vals)
        ci_d = bootstrap_ci_mean_delta(g_vals, zs_vals)
        ci_r = bootstrap_ci_rank_biserial(g_vals, zs_vals)
    table3.append({
        "modality": mod,
        "landmark": lm,
        "type": lm_type,
        "n": res["n_total"],
        "mean_delta": res["mean_delta"],
        "mean_delta_ci": ci_d,
        "W": res["W"],
        "p_raw": res["p"],
        "rank_biserial_r": res["rank_biserial_r"],
        "r_ci": ci_r,
    })


# ── Table 4: model vs student (pickle source) ───────────────────────
print("Computing Table 4 (model vs student) ...")
table4 = []
for model_name, model_pkl in (("GPT-5.4", gpt_pkl), ("Gemini 3.1 Pro", gem_pkl)):
    for strat in ("zero_shot", "guided"):
        for mod in ("CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"):
            mod_eds = []
            stu_eds = []
            for (qid, s), entry in model_pkl.items():
                if s != strat: continue
                if entry["modality"] != mod or entry["landmark_type"] != "point":
                    continue
                if entry["model_val"] is None or entry["student_val"] is None:
                    continue
                mod_eds.append(entry["model_val"])
                stu_eds.append(entry["student_val"])
            if not mod_eds:
                continue
            res = paired_wilcoxon(mod_eds, stu_eds)
            ci_d = bootstrap_ci_mean_delta(mod_eds, stu_eds)
            ci_r = bootstrap_ci_rank_biserial(mod_eds, stu_eds)
            table4.append({
                "model": model_name,
                "strategy": strat,
                "group": f"{mod.capitalize()} (point, ED)",
                "modality": mod,
                "metric": "ed",
                "n": res["n_total"],
                "model_mean": float(np.mean(mod_eds)),
                "student_mean": float(np.mean(stu_eds)),
                "mean_delta": res["mean_delta"],
                "mean_delta_ci": ci_d,
                "W": res["W"],
                "p_raw": res["p"],
                "rank_biserial_r": res["rank_biserial_r"],
                "r_ci": ci_r,
            })
        # Panoramic area (Jaccard)
        mod_js = []
        stu_js = []
        for (qid, s), entry in model_pkl.items():
            if s != strat: continue
            if entry["modality"] != "PANORAMIC" or entry["landmark_type"] != "area":
                continue
            if entry["model_val"] is None or entry["student_val"] is None:
                continue
            mod_js.append(entry["model_val"])
            stu_js.append(entry["student_val"])
        if mod_js:
            res = paired_wilcoxon(mod_js, stu_js)
            ci_d = bootstrap_ci_mean_delta(mod_js, stu_js)
            ci_r = bootstrap_ci_rank_biserial(mod_js, stu_js)
            table4.append({
                "model": model_name,
                "strategy": strat,
                "group": "Panoramic (area, Jaccard)",
                "modality": "PANORAMIC",
                "metric": "jaccard",
                "n": res["n_total"],
                "model_mean": float(np.mean(mod_js)),
                "student_mean": float(np.mean(stu_js)),
                "mean_delta": res["mean_delta"],
                "mean_delta_ci": ci_d,
                "W": res["W"],
                "p_raw": res["p"],
                "rank_biserial_r": res["rank_biserial_r"],
                "r_ci": ci_r,
            })


# ── Table 5: GPT vs Gemini (pickle source) ──────────────────────────
print("Computing Table 5 (GPT vs Gemini) ...")
table5 = []
for strat in ("zero_shot", "guided"):
    for mod in ("CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"):
        gpt_v = []
        gem_v = []
        for (qid, s), entry in gpt_pkl.items():
            if s != strat: continue
            if entry["modality"] != mod or entry["landmark_type"] != "point":
                continue
            gem_entry = gem_pkl.get((qid, strat))
            if gem_entry is None: continue
            if entry["model_val"] is None or gem_entry["model_val"] is None:
                continue
            gpt_v.append(entry["model_val"])
            gem_v.append(gem_entry["model_val"])
        if not gpt_v:
            continue
        # Δ = GPT - Gemini; positive Δ in ED means Gemini closer
        res = paired_wilcoxon(gpt_v, gem_v)
        ci_d = bootstrap_ci_mean_delta(gpt_v, gem_v)
        ci_r = bootstrap_ci_rank_biserial(gpt_v, gem_v)
        table5.append({
            "modality": mod,
            "strategy": strat,
            "metric": "ed",
            "n": res["n_total"],
            "gpt_mean": float(np.mean(gpt_v)),
            "gemini_mean": float(np.mean(gem_v)),
            "mean_delta": res["mean_delta"],
            "mean_delta_ci": ci_d,
            "W": res["W"],
            "p_raw": res["p"],
            "rank_biserial_r": res["rank_biserial_r"],
            "r_ci": ci_r,
        })
    # Panoramic area (Jaccard)
    gpt_v = []
    gem_v = []
    for (qid, s), entry in gpt_pkl.items():
        if s != strat: continue
        if entry["modality"] != "PANORAMIC" or entry["landmark_type"] != "area":
            continue
        gem_entry = gem_pkl.get((qid, strat))
        if gem_entry is None: continue
        if entry["model_val"] is None or gem_entry["model_val"] is None:
            continue
        gpt_v.append(entry["model_val"])
        gem_v.append(gem_entry["model_val"])
    if gpt_v:
        res = paired_wilcoxon(gpt_v, gem_v)
        ci_d = bootstrap_ci_mean_delta(gpt_v, gem_v)
        ci_r = bootstrap_ci_rank_biserial(gpt_v, gem_v)
        table5.append({
            "modality": "PANORAMIC",
            "strategy": strat,
            "metric": "jaccard",
            "n": res["n_total"],
            "gpt_mean": float(np.mean(gpt_v)),
            "gemini_mean": float(np.mean(gem_v)),
            "mean_delta": res["mean_delta"],
            "mean_delta_ci": ci_d,
            "W": res["W"],
            "p_raw": res["p"],
            "rank_biserial_r": res["rank_biserial_r"],
            "r_ci": ci_r,
        })


# ── Figure 4 — Tooth_33_Apex prediction distribution ───────────────
print("Computing Figure 4 distributions ...")
def cell_distribution(raw_dict, landmark="Tooth_33_Apex", strategy="guided"):
    cells = []
    for (qid, strat), reps_list in raw_dict.items():
        if strat != strategy or landmark not in qid:
            continue
        for r in reps_list:
            if r:
                cells.append(r[0])
    return Counter(cells), len(cells)

gpt_dist, gpt_total = cell_distribution(gpt_raw)
gem_dist, gem_total = cell_distribution(gem_raw)

fig4 = {
    "n_per_model": 300,
    "gpt_total": gpt_total,
    "gemini_total": gem_total,
    "gpt_distribution_pct": {
        cell: round(100 * count / gpt_total, 2)
        for cell, count in sorted(gpt_dist.items(), key=lambda x: -x[1])
    },
    "gemini_distribution_pct": {
        cell: round(100 * count / gem_total, 2)
        for cell, count in sorted(gem_dist.items(), key=lambda x: -x[1])
    },
    "gpt_distribution_count": dict(gpt_dist),
    "gemini_distribution_count": dict(gem_dist),
}


# ── Save ────────────────────────────────────────────────────────────
out = {
    "_meta": {
        "seed": RNG_SEED,
        "n_boot_mean": N_BOOT_MEAN,
        "n_boot_r": N_BOOT_R,
        "convention": (
            "Table 2 mean ED: bootstrap percentile CI on per-query mean ED "
            "(point landmarks, aggregated across 3 reps). "
            "SDR@k: per-query mean ED ≤ k threshold. "
            "Table 3 (point): Δ = ED(zero-shot) − ED(guided) so + favours guided. "
            "Table 3 (area): Δ = Jaccard(guided) − Jaccard(zero-shot) so + favours guided. "
            "Table 4: Δ = model − student (ED) so + means student closer; "
            "for area Jaccard Δ = model − student so − means student greater overlap. "
            "Table 5: Δ = GPT − Gemini (ED) so + means Gemini closer; "
            "for area Jaccard Δ = GPT − Gemini so − means Gemini greater overlap. "
            "W = sum of positive signed ranks (scipy.stats.wilcoxon default). "
            "r = rank-biserial r = (2W − T) / T, where T = n_nz·(n_nz+1)/2."
        ),
    },
    "table2": table2,
    "table3": table3,
    "table4": table4,
    "table5": table5,
    "figure4": fig4,
}

OUT.write_text(json.dumps(out, indent=2))
print(f"Wrote {OUT}")
print()
print(f"=== Quick sanity checks ===")
print(f"Table 2: {len(table2)} rows")
for r in table2[:3]:
    print(f"  {r['model']} {r['modality']} {r['strategy']}: mean={r['mean']:.3f} CI={r['mean_ci']}")
print(f"Table 3: {len(table3)} rows")
print(f"Table 4: {len(table4)} rows")
print(f"Table 5: {len(table5)} rows")
print(f"Figure 4 (GPT-5.4 top 3): {dict(list(fig4['gpt_distribution_pct'].items())[:3])}")
print(f"Figure 4 (Gemini top 3): {dict(list(fig4['gemini_distribution_pct'].items())[:3])}")
