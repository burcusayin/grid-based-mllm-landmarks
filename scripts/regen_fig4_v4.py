"""Regenerate Figure 4 (Tooth_33_Apex prediction heatmap, GPT vs Gemini, guided)
using real raw data for both models.

Source: results_v4_canonical.json (figure4 block).

Output: docs/submission/fig4.png  (saves to v4 submission folder).

Notes
-----
- Each model has n = 300 guided predictions (100 panoramic images × 3 reps).
- Consensus GT for Tooth_33_Apex on PAN images is G10 (lower-left canine apex).
- GPT-5.4 cluster: F5 (44.7%) and F6 (42.7%) → F5/F6 attractor (87.3%).
- Gemini 3.1 Pro cluster: F10 (72.3%), G10 (22.3% — exact match), tiny
  remainder at F11/G11. No F5/F6 predictions at all.
- Grid: 16 columns × 8 rows (panoramic). Rows A-H top to bottom; cols 1-16
  left to right.
"""
from __future__ import annotations
import json
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

ROOT = Path(__file__).resolve().parent.parent
CANONICAL = ROOT / "results_v4_canonical.json"
OUT_PNG   = ROOT / "docs" / "submission" / "fig4.png"

# Panoramic grid: 16 cols × 8 rows (A-H)
N_ROWS = 8
N_COLS = 16
ROW_LETTERS = "ABCDEFGH"

def cell_to_idx(cell: str):
    row = ROW_LETTERS.index(cell[0].upper())
    col = int(cell[1:]) - 1
    return row, col


def build_heatmap_matrix(distribution_pct: dict[str, float]):
    M = np.zeros((N_ROWS, N_COLS), dtype=float)
    for cell, pct in distribution_pct.items():
        try:
            r, c = cell_to_idx(cell)
            if 0 <= r < N_ROWS and 0 <= c < N_COLS:
                M[r, c] = pct
        except Exception:
            pass
    return M


def main() -> None:
    canonical = json.loads(CANONICAL.read_text())
    fig4_data = canonical["figure4"]

    gpt = build_heatmap_matrix(fig4_data["gpt_distribution_pct"])
    gem = build_heatmap_matrix(fig4_data["gemini_distribution_pct"])

    GT_ROW, GT_COL = cell_to_idx("G10")   # consensus ground truth
    F5_ROW, F5_COL = cell_to_idx("F5")
    F6_ROW, F6_COL = cell_to_idx("F6")

    # Shared colour scale
    vmax = max(gpt.max(), gem.max())

    fig, (ax_gpt, ax_gem) = plt.subplots(1, 2, figsize=(13, 5),
                                          gridspec_kw={"wspace": 0.18})

    # Common helper
    def render(ax, M, title):
        im = ax.imshow(M, cmap="YlOrRd", vmin=0, vmax=vmax, aspect="equal",
                       origin="upper")
        ax.set_xticks(range(N_COLS))
        ax.set_xticklabels([str(i + 1) for i in range(N_COLS)], fontsize=9)
        ax.set_yticks(range(N_ROWS))
        ax.set_yticklabels(list(ROW_LETTERS), fontsize=9)
        ax.set_title(title, fontsize=11, pad=8)
        # Grid lines between cells
        ax.set_xticks(np.arange(-0.5, N_COLS, 1), minor=True)
        ax.set_yticks(np.arange(-0.5, N_ROWS, 1), minor=True)
        ax.grid(which="minor", color="white", linewidth=0.6)
        ax.tick_params(which="minor", length=0)
        # Annotate cells with ≥ 3% of predictions. The "%" is dropped
        # from cell labels (the colorbar already states the unit) and
        # the font size is tuned so even two-digit values fit inside
        # the narrow panoramic grid cells at 300 DPI.
        for rr in range(N_ROWS):
            for cc in range(N_COLS):
                val = M[rr, cc]
                if val >= 3.0:
                    color = "white" if val > vmax * 0.5 else "black"
                    ax.text(cc, rr, f"{val:.0f}", ha="center", va="center",
                            fontsize=6.5, color=color, fontweight="bold")
        # GT cell (G10) green border
        ax.add_patch(mpatches.Rectangle(
            (GT_COL - 0.5, GT_ROW - 0.5), 1, 1,
            fill=False, edgecolor="#2ca02c", linewidth=2.5,
        ))
        # F5 + F6 attractor blue dashed border around the two cells
        ax.add_patch(mpatches.Rectangle(
            (F5_COL - 0.5, F5_ROW - 0.5), 2, 1,
            fill=False, edgecolor="#1f77b4", linewidth=2.0, linestyle="--",
        ))
        return im

    im1 = render(ax_gpt, gpt, "GPT-5.4 (Guided)")
    im2 = render(ax_gem, gem, "Gemini 3.1 Pro (Guided)")

    # Colorbar on the right of the second panel; the colorbar tick
    # format makes the percentage unit explicit so the cell labels
    # can stay un-suffixed.
    cbar = fig.colorbar(im2, ax=[ax_gpt, ax_gem], shrink=0.85,
                        label="Prediction frequency (% of 300)", pad=0.02)
    cbar.ax.tick_params(labelsize=9)

    # Legend bar at bottom (single, shared)
    green_patch = mpatches.Patch(facecolor="none", edgecolor="#2ca02c",
                                  linewidth=2, label="Consensus GT (G10)")
    blue_patch = mpatches.Patch(facecolor="none", edgecolor="#1f77b4",
                                 linewidth=2, linestyle="--",
                                 label="F5/F6 attractor (GPT)")
    fig.legend(handles=[green_patch, blue_patch], loc="lower center", ncol=2,
               frameon=True, fontsize=10, bbox_to_anchor=(0.5, -0.01))

    fig.suptitle("Tooth_33_Apex: Prediction Distribution Under Guided Prompting\n"
                 "(n = 300 predictions per model)",
                 fontsize=12, y=0.98)

    plt.savefig(OUT_PNG, dpi=300, bbox_inches="tight")
    print(f"Saved {OUT_PNG}")

    # Print top-3 sanity for verification
    print("\nGPT-5.4 top cells:")
    for cell, pct in list(fig4_data["gpt_distribution_pct"].items())[:5]:
        print(f"  {cell}: {pct:.1f}%")
    print("Gemini top cells:")
    for cell, pct in list(fig4_data["gemini_distribution_pct"].items())[:5]:
        print(f"  {cell}: {pct:.1f}%")


if __name__ == "__main__":
    main()
