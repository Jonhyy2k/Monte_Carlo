"""
Phase 5 — Output and Visualization

Three presentation-quality charts:
1. Price distribution histogram with KDE overlay + vertical markers
2. Percentile table (5th, 25th, 50th, 75th, 95th)
3. Sobol sensitivity bar chart ranked by influence
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

from dcf_engine import PROJECTION_PARAM_LABELS, run_monte_carlo_vectorized
from sobol_analysis import run_sobol


# ---------------------------------------------------------------------------
# Colour palette & style
# ---------------------------------------------------------------------------
DARK_BG = "#1a1a2e"
CARD_BG = "#16213e"
ACCENT_BLUE = "#0f3460"
ACCENT_TEAL = "#00b4d8"
ACCENT_GOLD = "#e2b93d"
ACCENT_RED = "#e74c3c"
ACCENT_GREEN = "#2ecc71"
TEXT_COLOR = "#e0e0e0"
GRID_COLOR = "#2a2a4a"

PARAM_LABELS = {
    "revenue_growth": "Revenue Growth",
    "cogs_pct": "COGS %",
    "sga_pct": "SG&A %",
    "capex_intensity": "CapEx Intensity",
    "tax_rate": "Tax Rate",
    "exit_multiple": "Exit Multiple",
    "terminal_g": "Terminal Growth",
    "wacc": "WACC",
}
PARAM_LABELS.update(PROJECTION_PARAM_LABELS)


def _style_ax(ax: plt.Axes) -> None:
    ax.set_facecolor(CARD_BG)
    ax.tick_params(colors=TEXT_COLOR, labelsize=10)
    ax.xaxis.label.set_color(TEXT_COLOR)
    ax.yaxis.label.set_color(TEXT_COLOR)
    ax.title.set_color(TEXT_COLOR)
    for spine in ax.spines.values():
        spine.set_color(GRID_COLOR)
    ax.grid(axis="y", color=GRID_COLOR, linewidth=0.5, alpha=0.5)


def _estimate_density(samples: np.ndarray, x: np.ndarray, bw_scale: float = 1.0) -> np.ndarray:
    """Lightweight Gaussian KDE without SciPy to keep the pipeline self-contained."""
    samples = np.asarray(samples, dtype=float)
    if samples.size < 2:
        return np.zeros_like(x)

    std = np.std(samples, ddof=1)
    if std <= 0:
        return np.zeros_like(x)

    bandwidth = max(1.06 * std * samples.size ** (-1 / 5) * bw_scale, 1e-6)
    z = (x[:, np.newaxis] - samples[np.newaxis, :]) / bandwidth
    kernel = np.exp(-0.5 * z ** 2) / np.sqrt(2 * np.pi)
    return kernel.mean(axis=1) / bandwidth


def plot_price_distribution(
    prices: np.ndarray,
    current_price: float = 55.0,
    fundamental_target: float = 68.0,
    save_path: str = "price_distribution.png",
) -> None:
    """Histogram + KDE + vertical reference lines."""

    fig, ax = plt.subplots(figsize=(12, 6))
    fig.patch.set_facecolor(DARK_BG)
    _style_ax(ax)

    # Remove extreme outliers for cleaner display
    p1, p99 = np.percentile(prices, [1, 99])
    display = prices[(prices >= p1) & (prices <= p99)]

    # Histogram
    ax.hist(
        display, bins=120, density=True, alpha=0.6,
        color=ACCENT_TEAL, edgecolor="none", label="Simulation",
    )

    # KDE overlay
    x = np.linspace(display.min(), display.max(), 500)
    ax.plot(x, _estimate_density(display, x, bw_scale=0.8),
            color=ACCENT_GOLD, linewidth=2.5, label="KDE")

    # Vertical markers
    sim_mean = np.mean(prices)
    ax.axvline(current_price, color=ACCENT_RED, linestyle="--", linewidth=1.8,
               label=f"Current Price (${current_price:.0f})")
    ax.axvline(sim_mean, color=ACCENT_GREEN, linestyle="-.", linewidth=1.8,
               label=f"Simulation Mean (${sim_mean:.0f})")
    ax.axvline(fundamental_target, color=ACCENT_GOLD, linestyle=":", linewidth=1.8,
               label=f"Fundamental Target (${fundamental_target:.0f})")

    ax.set_xlabel("Implied Share Price ($)", fontsize=12, fontweight="bold")
    ax.set_ylabel("Density", fontsize=12, fontweight="bold")
    ax.set_title("Monte Carlo DCF — Implied Share Price Distribution",
                 fontsize=14, fontweight="bold", pad=15)
    ax.xaxis.set_major_formatter(mticker.StrMethodFormatter("${x:,.0f}"))
    ax.legend(loc="upper right", fontsize=9, facecolor=CARD_BG,
              edgecolor=GRID_COLOR, labelcolor=TEXT_COLOR)

    fig.tight_layout()
    fig.savefig(save_path, dpi=200, facecolor=DARK_BG)
    plt.close(fig)
    print(f"  Saved: {save_path}")


def build_percentile_table(prices: np.ndarray) -> pd.DataFrame:
    """Percentile table: 5th, 25th, 50th, 75th, 95th."""
    pctls = [5, 25, 50, 75, 95]
    vals = np.percentile(prices, pctls)
    df = pd.DataFrame({
        "Percentile": [f"{p}th" for p in pctls],
        "Implied Price ($)": [f"${v:,.2f}" for v in vals],
    })
    return df


def plot_percentile_table(
    prices: np.ndarray,
    save_path: str = "percentile_table.png",
) -> None:
    """Render the percentile table as a styled image."""
    pctls = [5, 25, 50, 75, 95]
    vals = np.percentile(prices, pctls)

    fig, ax = plt.subplots(figsize=(6, 3))
    fig.patch.set_facecolor(DARK_BG)
    ax.set_facecolor(DARK_BG)
    ax.axis("off")

    table_data = [[f"{p}th", f"${v:,.2f}"] for p, v in zip(pctls, vals)]
    table = ax.table(
        cellText=table_data,
        colLabels=["Percentile", "Implied Price"],
        cellLoc="center",
        loc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1, 1.6)

    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor(GRID_COLOR)
        if row == 0:
            cell.set_facecolor(ACCENT_BLUE)
            cell.set_text_props(color=TEXT_COLOR, fontweight="bold")
        else:
            cell.set_facecolor(CARD_BG)
            cell.set_text_props(color=TEXT_COLOR)

    ax.set_title("Implied Price Percentiles", fontsize=13,
                 fontweight="bold", color=TEXT_COLOR, pad=20)

    fig.tight_layout()
    fig.savefig(save_path, dpi=200, facecolor=DARK_BG, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {save_path}")


def plot_sobol_sensitivity(
    sobol_df: pd.DataFrame,
    save_path: str = "sobol_sensitivity.png",
) -> None:
    """Horizontal bar chart of Sobol total-order indices."""
    fig, ax = plt.subplots(figsize=(10, 5))
    fig.patch.set_facecolor(DARK_BG)
    _style_ax(ax)

    df = sobol_df.sort_values("ST", ascending=True)
    labels = [PARAM_LABELS.get(p, p) for p in df["Parameter"]]
    bars = ax.barh(labels, df["ST"], color=ACCENT_TEAL, edgecolor="none",
                   height=0.6, alpha=0.85)
    ax.errorbar(df["ST"], labels, xerr=df["ST_conf"], fmt="none",
                ecolor=TEXT_COLOR, elinewidth=1, capsize=3)

    # Value annotations
    for bar, val in zip(bars, df["ST"]):
        ax.text(bar.get_width() + 0.005, bar.get_y() + bar.get_height() / 2,
                f"{val:.3f}", va="center", ha="left",
                color=TEXT_COLOR, fontsize=9)

    ax.set_xlabel("Total-Order Sobol Index (ST)", fontsize=12, fontweight="bold")
    ax.set_title("Sensitivity Analysis — Which Inputs Drive Valuation Variance?",
                 fontsize=13, fontweight="bold", pad=15)
    ax.grid(axis="x", color=GRID_COLOR, linewidth=0.5, alpha=0.5)
    ax.grid(axis="y", visible=False)

    fig.tight_layout()
    fig.savefig(save_path, dpi=200, facecolor=DARK_BG)
    plt.close(fig)
    print(f"  Saved: {save_path}")


# ---------------------------------------------------------------------------
# Main — generate all outputs with synthetic data
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=== Phase 5: Generating Visualizations ===\n")

    # Run Monte Carlo
    prices = run_monte_carlo_vectorized(n_iter=50_000)

    # 1. Price distribution
    plot_price_distribution(prices, current_price=55.0, fundamental_target=68.0)

    # 2. Percentile table
    print("\n", build_percentile_table(prices).to_string(index=False), "\n")
    plot_percentile_table(prices)

    # 3. Sobol sensitivity
    sobol_df = run_sobol(n_samples=2048)
    plot_sobol_sensitivity(sobol_df)

    print("\n  All charts generated.")
