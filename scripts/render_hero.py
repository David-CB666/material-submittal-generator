# -*- coding: utf-8 -*-
"""
Generate hero image for material-submittal-generator README.
Pipeline visualization: Template → Multi-Sheet Excel → PDFs → BQ Merged
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np

# ── Style ──
BG = '#0D1117'
CARD_BG = '#161B22'
ACCENT = '#58A6FF'
ACCENT2 = '#3FB950'
ACCENT3 = '#D29922'
ACCENT4 = '#F78166'
TEXT = '#E6EDF3'
TEXT_DIM = '#8B949E'
BORDER = '#30363D'

fig, ax = plt.subplots(1, 1, figsize=(12, 4.5))
ax.set_xlim(0, 12)
ax.set_ylim(0, 4.5)
ax.set_aspect('equal')
ax.axis('off')
fig.patch.set_facecolor(BG)
ax.set_facecolor(BG)

def draw_card(ax, x, y, w, h, color, title, subtitle, icon):
    """Draw a rounded card with title + subtitle."""
    box = FancyBboxPatch((x, y), w, h,
                         boxstyle="round,pad=0.06",
                         facecolor=CARD_BG, edgecolor=color, linewidth=2)
    ax.add_patch(box)
    ax.text(x + w/2, y + h - 0.28, icon, fontsize=18, ha='center', va='top')
    ax.text(x + w/2, y + h - 0.72, title, fontsize=11, fontweight='bold',
            ha='center', va='top', color=TEXT, fontfamily='monospace')
    ax.text(x + w/2, y + 0.18, subtitle, fontsize=7.5, ha='center', va='bottom',
            color=TEXT_DIM, fontfamily='monospace')

def draw_arrow(ax, x1, y1, x2, y2, color):
    """Draw a thick arrow."""
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color=color, lw=2.5,
                               connectionstyle='arc3,rad=0'))

# ── Pipeline cards ──
# Card 1: Input
draw_card(ax, 0.3, 1.25, 2.2, 2.0, ACCENT,
          'INPUT', 'Template .xlsx\n+ Source Data .json', '')

# Card 2: Generate
draw_card(ax, 3.3, 1.25, 2.2, 2.0, ACCENT2,
          'gen.py', 'Pure ZIP Engine\nPreserves Images & Styles', '')

# Card 3: Export PDF
draw_card(ax, 6.3, 1.25, 2.2, 2.0, ACCENT3,
          'export_pdf.py', 'Excel COM Export\nOne PDF per Sheet', '')

# Card 4: Merge BQ
draw_card(ax, 9.3, 1.25, 2.2, 2.0, ACCENT4,
          'merge_bq.py', 'Auto-match BQ Pages\n→ Final Submittal PDF', '')

# ── Arrows ──
arrow_y = 2.25
ax.annotate('', xy=(3.2, arrow_y), xytext=(2.6, arrow_y),
            arrowprops=dict(arrowstyle='->', color=TEXT_DIM, lw=2))
ax.annotate('', xy=(6.2, arrow_y), xytext=(5.6, arrow_y),
            arrowprops=dict(arrowstyle='->', color=TEXT_DIM, lw=2))
ax.annotate('', xy=(9.2, arrow_y), xytext=(8.6, arrow_y),
            arrowprops=dict(arrowstyle='->', color=TEXT_DIM, lw=2))

# ── Bottom tagline ──
ax.text(6, 3.9, 'Material Submittal Generator',
        fontsize=20, fontweight='bold', ha='center', va='center',
        color=ACCENT, fontfamily='monospace')
ax.text(6, 3.45, 'One-click batch submittal sheets + automatic BQ page merging for MEP construction',
        fontsize=9, ha='center', va='center', color=TEXT_DIM, fontfamily='monospace')

# ── Save ──
plt.tight_layout(pad=0)
plt.savefig('demo/hero.png', dpi=150, facecolor=BG, bbox_inches='tight', pad_inches=0.3)
plt.close()
print('OK demo/hero.png generated')
