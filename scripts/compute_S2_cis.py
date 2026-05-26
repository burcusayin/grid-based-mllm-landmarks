"""Compute 95% bootstrap CIs for Cohen's κ and mean Jaccard in Table S2.

Sources:
  - results_consensus/query_index.json (omfr_1, omfr_2, omfr_1_second,
    omfr_2_second per query)

Writes: results_v4_S2_cis.json
"""
from __future__ import annotations
import json
import math
from pathlib import Path
import numpy as np

ROOT = Path(__file__).resolve().parent.parent

N_BOOT = 10_000
SEED = 42

ROW_LETTERS = "ABCDEFGH"


def to_rc(c: str):
    c = c.strip()
    if c[0].isalpha():
        return ord(c[0].upper()) - ord("A"), int(c[1:])
    i = 0
    while i < len(c) and c[i].isdigit():
        i += 1
    return ord(c[i].upper()) - ord("A"), int(c[:i])


def parse_cells(s):
    if not s or not isinstance(s, str): return []
    return [t.strip() for t in s.split(",") if t.strip()]


def euclid(a, b):
    r1, c1 = to_rc(a); r2, c2 = to_rc(b)
    return math.sqrt((r1-r2)**2 + (c1-c2)**2)


def jaccard(a: set, b: set):
    if not a and not b: return 1.0
    return len(a & b) / len(a | b) if (a | b) else 0.0


def cohens_kappa_pairs(pairs: list[tuple[str, str]]) -> float:
    """Compute Cohen's κ over the *cell* prediction agreement.

    Each pair is (rater1_cell, rater2_cell). Categories are the distinct
    cells. Returns κ as fraction of agreement above chance."""
    if not pairs: return 0.0
    cats = sorted({c for pair in pairs for c in pair})
    n = len(pairs)
    p_obs = sum(1 for a, b in pairs if a == b) / n
    # Marginal probs
    r1 = [p[0] for p in pairs]; r2 = [p[1] for p in pairs]
    p_exp = 0.0
    for c in cats:
        pa = r1.count(c) / n
        pb = r2.count(c) / n
        p_exp += pa * pb
    if p_exp >= 1.0:
        return 1.0
    return (p_obs - p_exp) / (1.0 - p_exp)


def boot_kappa_ci(pairs):
    rng = np.random.default_rng(SEED)
    n = len(pairs)
    kappas = np.empty(N_BOOT)
    arr = np.array(pairs)
    for i in range(N_BOOT):
        idx = rng.integers(0, n, n)
        sub = [tuple(arr[j]) for j in idx]
        kappas[i] = cohens_kappa_pairs(sub)
    return float(np.percentile(kappas, 2.5)), float(np.percentile(kappas, 97.5))


def boot_mean_ci(values):
    rng = np.random.default_rng(SEED)
    arr = np.asarray(values, dtype=float)
    n = len(arr)
    means = np.empty(N_BOOT)
    for i in range(N_BOOT):
        idx = rng.integers(0, n, n)
        means[i] = arr[idx].mean()
    return float(np.percentile(means, 2.5)), float(np.percentile(means, 97.5))


def main():
    qi = json.loads((ROOT / "results_consensus" / "query_index.json").read_text())

    # ── INTER-RATER (OMFR_1 vs OMFR_2) ─────────────────────────────
    inter_results = {}
    for mod, name in (("CEPHALOMETRIC", "Cephalometric (point)"),
                       ("PERIAPICAL", "Periapical (point)"),
                       ("PANORAMIC", "Panoramic (point)")):
        pairs = []
        for q in qi:
            if q["sheet"] != mod or q["landmark_type"] != "point":
                continue
            o1 = q.get("omfr_1"); o2 = q.get("omfr_2")
            if not o1 or not o2: continue
            o1_c = parse_cells(o1); o2_c = parse_cells(o2)
            if o1_c and o2_c:
                pairs.append((o1_c[0], o2_c[0]))
        if not pairs: continue
        kappa = cohens_kappa_pairs(pairs)
        ci = boot_kappa_ci(pairs)
        # Within-1-cell rate
        within_1 = sum(1 for a, b in pairs if euclid(a, b) <= 1.0) / len(pairs)
        inter_results[name] = {
            "n": len(pairs),
            "kappa": kappa,
            "kappa_ci": ci,
            "within_1": within_1 * 100,
        }
        print(f"  Inter {name}: n={len(pairs)}, κ={kappa:.3f} {ci}, within-1={within_1*100:.1f}%")

    # Area landmarks (panoramic, Jaccard / Dice)
    pan_area_jaccs = []
    pan_area_dices = []
    for q in qi:
        if q["sheet"] != "PANORAMIC" or q["landmark_type"] != "area": continue
        o1 = q.get("omfr_1"); o2 = q.get("omfr_2")
        if not o1 or not o2: continue
        a = set(parse_cells(o1)); b = set(parse_cells(o2))
        if not a or not b: continue
        pan_area_jaccs.append(jaccard(a, b))
        d = (2.0 * len(a & b)) / (len(a) + len(b)) if (len(a) + len(b)) > 0 else 1.0
        pan_area_dices.append(d)
    mean_j = float(np.mean(pan_area_jaccs))
    mean_d = float(np.mean(pan_area_dices))
    j_ci = boot_mean_ci(pan_area_jaccs)
    d_ci = boot_mean_ci(pan_area_dices)
    inter_results["Panoramic (area, Jaccard / Dice)"] = {
        "n": len(pan_area_jaccs),
        "jaccard": mean_j,
        "jaccard_ci": j_ci,
        "dice": mean_d,
        "dice_ci": d_ci,
    }
    print(f"  Inter PAN area: n={len(pan_area_jaccs)}, J={mean_j:.3f} {j_ci}, D={mean_d:.3f} {d_ci}")

    # ── INTRA-RATER ─────────────────────────────────────────────────
    intra_results = {}
    for which in ("omfr_1", "omfr_2"):
        # Point landmarks (all modalities pooled)
        pairs = []
        for q in qi:
            if q["landmark_type"] != "point": continue
            v1 = q.get(which); v2 = q.get(f"{which}_second")
            if not v1 or not v2: continue
            c1 = parse_cells(v1); c2 = parse_cells(v2)
            if c1 and c2:
                pairs.append((c1[0], c2[0]))
        if pairs:
            kappa = cohens_kappa_pairs(pairs)
            ci = boot_kappa_ci(pairs)
            within_1 = sum(1 for a, b in pairs if euclid(a, b) <= 1.0) / len(pairs)
            intra_results[f"{which}: point landmarks"] = {
                "n": len(pairs),
                "kappa": kappa,
                "kappa_ci": ci,
                "within_1": within_1 * 100,
            }
            print(f"  Intra {which} point: n={len(pairs)}, κ={kappa:.3f} {ci}")
        # Area landmarks (panoramic only)
        jaccs = []
        for q in qi:
            if q["sheet"] != "PANORAMIC" or q["landmark_type"] != "area": continue
            v1 = q.get(which); v2 = q.get(f"{which}_second")
            if not v1 or not v2: continue
            a = set(parse_cells(v1)); b = set(parse_cells(v2))
            if a and b:
                jaccs.append(jaccard(a, b))
        if jaccs:
            mean_j = float(np.mean(jaccs))
            j_ci = boot_mean_ci(jaccs)
            intra_results[f"{which}: area landmarks (Panoramic, Jaccard)"] = {
                "n": len(jaccs),
                "jaccard": mean_j,
                "jaccard_ci": j_ci,
            }
            print(f"  Intra {which} area: n={len(jaccs)}, J={mean_j:.3f} {j_ci}")

    out = {"inter": inter_results, "intra": intra_results}
    OUT = ROOT / "results_v4_S2_cis.json"
    OUT.write_text(json.dumps(out, indent=2))
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
