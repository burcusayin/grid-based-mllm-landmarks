"""Regenerate Figure 5 (GPT-5.4 per-landmark zero-shot vs guided effect sizes).

v3 figure showed several NS landmarks at r ≈ 0 instead of their actual
rank-biserial r (Nasion +0.19, T36-distal_Apex −0.46, T36-mesial_CEJ +0.42,
Maxillary_Sinus +0.15). This regenerator uses canonical r from
results_v4_canonical.json (which matches results_consensus/analysis.json
byte-for-byte) for all 12 landmarks.

Output: docs/submission/fig5.png
"""
from __future__ import annotations
import json
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

ROOT = Path(__file__).resolve().parent.parent
CANONICAL = ROOT / "results_v4_canonical.json"
OUT_PNG = ROOT / "docs" / "submission" / "fig5.png"


# Canonical landmark display order (matches Figure 6 for direct comparison)
DISPLAY_ORDER = [
    ("Menton (CEPH)",           "Menton_Me",              "CEPHALOMETRIC", "point"),
    ("Nasion (CEPH)",           "Nasion_N",               "CEPHALOMETRIC", "point"),
    ("Sella (CEPH)",            "Sella_S",                "CEPHALOMETRIC", "point"),
    ("T36-distal apex (PA)",    "Tooth_36_Distal_Apex",   "PERIAPICAL",    "point"),
    ("T36-distal CEJ (PA)",     "Tooth_36_Distal_CEJ",    "PERIAPICAL",    "point"),
    ("T36-mesial CEJ (PA)",     "Tooth_36_Mesial_CEJ",    "PERIAPICAL",    "point"),
    ("Mental foramen (PAN)",    "Mental_Foramen_L",       "PANORAMIC",     "point"),
    ("Condylar head (PAN)",     "Condylar_Head_R",        "PANORAMIC",     "point"),
    ("T33 apex (PAN)",          "Tooth_33_Apex",          "PANORAMIC",     "point"),
    ("Mandibular canal (PAN)",  "Mandibular_Canal_L",     "PANORAMIC",     "area"),
    ("Maxillary sinus (PAN)",   "Maxillary_Sinus_R",      "PANORAMIC",     "area"),
    ("Ext. oblique ridge (PAN)","External_Oblique_Ridge_R","PANORAMIC",    "area"),
]

COLORS = {
    "CEPHALOMETRIC": "#1f77b4",
    "PERIAPICAL":    "#ff7f0e",
    "PANORAMIC":     "#2ca02c",
}


def sig_marker(p_raw: float, n: int) -> str:
    p = p_raw * n
    if p < 0.001: return "***"
    if p < 0.01:  return "**"
    if p < 0.05:  return "*"
    return "NS"


canonical = json.loads(CANONICAL.read_text())
t3 = {r["landmark"]: r for r in canonical["table3"]}

fig, ax = plt.subplots(figsize=(13, 6))
xs = list(range(len(DISPLAY_ORDER)))

for i, (label, lm_key, mod, lm_type) in enumerate(DISPLAY_ORDER):
    cr = t3.get(lm_key)
    if cr is None:
        continue
    r = cr["rank_biserial_r"]
    p_raw = cr["p_raw"]
    bonf_n = 3 if lm_type == "area" else 9
    star = sig_marker(p_raw, bonf_n)
    color = COLORS[mod]
    marker = "s" if lm_type == "area" else "o"
    ax.scatter(i, r, s=140, marker=marker, color=color, edgecolors="black",
               linewidths=0.6, zorder=3)
    y_off = 0.08 if r >= 0 else -0.12
    ax.text(i, r + y_off, star, ha="center", va="center",
            fontsize=10, fontweight="bold" if star != "NS" else "normal",
            color="black")

ax.axhline(0, color="black", linestyle="--", linewidth=0.7, alpha=0.6)
ax.set_xticks(xs)
ax.set_xticklabels([d[0] for d in DISPLAY_ORDER], rotation=30, ha="right",
                   fontsize=10)
ax.set_ylabel("Rank-biserial r  (GPT-5.4: zero-shot vs guided)", fontsize=11)
ax.set_ylim(-1.15, 1.15)
ax.yaxis.set_major_locator(plt.MultipleLocator(0.25))
ax.grid(axis="y", linestyle=":", alpha=0.4)
ax.set_axisbelow(True)

# Legend (point vs area, color = modality)
from matplotlib.lines import Line2D
legend_elems = [
    Line2D([0], [0], marker="o", color="w", markerfacecolor="#1f77b4",
            markeredgecolor="black", markersize=10, label="CEPH point"),
    Line2D([0], [0], marker="o", color="w", markerfacecolor="#ff7f0e",
            markeredgecolor="black", markersize=10, label="PA point"),
    Line2D([0], [0], marker="o", color="w", markerfacecolor="#2ca02c",
            markeredgecolor="black", markersize=10, label="PAN point"),
    Line2D([0], [0], marker="s", color="w", markerfacecolor="#2ca02c",
            markeredgecolor="black", markersize=10, label="PAN area"),
]
ax.legend(handles=legend_elems, loc="upper right", fontsize=10, frameon=True)

# Bottom-left annotation box
ax.text(0.01, 0.04,
        "Bonferroni: ×9 (point) / ×3 (area)\n"
        "*** p<0.001  ** p<0.01  * p<0.05  NS = not significant",
        transform=ax.transAxes, fontsize=9,
        bbox=dict(boxstyle="round,pad=0.4", facecolor="white",
                  edgecolor="black", linewidth=0.6),
        verticalalignment="bottom")

ax.set_title("Figure 5. Per-landmark strategy effects in GPT-5.4 "
             "(zero-shot vs guided)", fontsize=12, pad=12)

plt.tight_layout()
plt.savefig(OUT_PNG, dpi=300, bbox_inches="tight")
print(f"Saved {OUT_PNG}")
print()
print("=== Figure 5 values ===")
for label, lm_key, _, lm_type in DISPLAY_ORDER:
    cr = t3[lm_key]
    bonf_n = 3 if lm_type == "area" else 9
    star = sig_marker(cr["p_raw"], bonf_n)
    print(f"  {label:>28}: r={cr['rank_biserial_r']:+.3f}, p_raw={cr['p_raw']:.3e}, mark={star}")
