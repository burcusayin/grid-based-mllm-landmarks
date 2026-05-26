"""Regenerate Figure 3 (SDR@1 bar chart) using corrected SDR@1 values.

v3 figure had: Gemini cephalometric and periapical SDR@1 inconsistent
with the corrected Table 2; student SDR@1 values for all three modalities
also drifted from raw recomputation.

v4 source: results_v4_canonical.json for GPT and Gemini SDR@1; student
SDR@1 computed directly from raw (results_consensus/query_index.json).
"""
from __future__ import annotations
import json
import math
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

ROOT = Path(__file__).resolve().parent.parent
CANONICAL = ROOT / "results_v4_canonical.json"
OUT_PNG = ROOT / "docs" / "submission" / "fig3.png"


def to_rc(c: str):
    c = c.strip()
    if c[0].isalpha():
        return ord(c[0].upper()) - ord("A"), int(c[1:])
    i = 0
    while i < len(c) and c[i].isdigit():
        i += 1
    return ord(c[i].upper()) - ord("A"), int(c[:i])


def euclid(a, b):
    r1, c1 = to_rc(a); r2, c2 = to_rc(b)
    return math.sqrt((r1 - r2) ** 2 + (c1 - c2) ** 2)


# Pull canonical SDR@1 for GPT and Gemini
canonical = json.loads(CANONICAL.read_text())
sdr_by = {(r["model"], r["modality"], r["strategy"]): r["SDR_1"]
          for r in canonical["table2"]}

# Compute student SDR@1 from raw
qi = json.loads((ROOT / "results_consensus" / "query_index.json").read_text())
student_sdr1 = {}
for mod in ("CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"):
    eds = []
    for q in qi:
        if q.get("sheet") != mod or q.get("landmark_type") != "point":
            continue
        stu = q.get("student"); gt = q.get("consensus_gt")
        if not stu or not gt:
            continue
        try:
            stu_cell = stu.split(",")[0].strip()
            gt_cell  = gt.split(",")[0].strip()
            eds.append(euclid(stu_cell, gt_cell))
        except Exception:
            pass
    student_sdr1[mod] = 100.0 * sum(1 for e in eds if e <= 1) / len(eds)

# Figure layout
modalities = ["Cephalometric", "Periapical", "Panoramic\n(point)"]
modality_keys = ["CEPHALOMETRIC", "PERIAPICAL", "PANORAMIC"]
groups = [
    ("GPT-5.4\nZero-shot",  ("GPT-5.4", "zero_shot"),         "#2868b8"),
    ("GPT-5.4\nGuided",     ("GPT-5.4", "guided"),            "#88b3e0"),
    ("Gemini 3.1 Pro\nZero-shot", ("Gemini 3.1 Pro", "zero_shot"), "#b8512a"),
    ("Gemini 3.1 Pro\nGuided",    ("Gemini 3.1 Pro", "guided"),    "#ec9968"),
    ("Student\nConsensus", None,                              "#386b3a"),
]

n_groups = len(groups)
n_mod = len(modalities)
bar_w = 0.15

fig, ax = plt.subplots(figsize=(13, 6.5))
x = np.arange(n_mod)

for i, (label, key, color) in enumerate(groups):
    if key is None:  # Student
        vals = [student_sdr1[m] for m in modality_keys]
    else:
        model, strat = key
        vals = [sdr_by[(model, m, strat)] for m in modality_keys]
    bars = ax.bar(x + (i - 2) * bar_w, vals, bar_w, label=label, color=color,
                   edgecolor="black", linewidth=0.5)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_x() + bar.get_width() / 2, v + 1.0,
                f"{v:.0f}", ha="center", fontsize=9, fontweight="bold")

ax.set_xticks(x)
ax.set_xticklabels(modalities, fontsize=11)
ax.set_ylabel("Successful Detection Rate @ 1 cell (%)", fontsize=11)
ax.set_ylim(0, 105)
ax.yaxis.set_major_locator(plt.MultipleLocator(20))
ax.grid(axis="y", linestyle=":", alpha=0.5)
ax.legend(title="Evaluator", bbox_to_anchor=(1.01, 1), loc="upper left",
          fontsize=9, title_fontsize=10, frameon=True)
ax.set_axisbelow(True)

plt.tight_layout()
plt.savefig(OUT_PNG, dpi=300, bbox_inches="tight")
print(f"Saved {OUT_PNG}")
print()
print("=== Figure 3 SDR@1 values used ===")
for label, key, _ in groups:
    if key is None:
        for m in modality_keys:
            print(f"  Student {m}: {student_sdr1[m]:.1f}%")
    else:
        model, strat = key
        for m in modality_keys:
            print(f"  {model} {strat} {m}: {sdr_by[(model, m, strat)]:.1f}%")
