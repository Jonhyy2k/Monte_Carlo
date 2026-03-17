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
PAPER_BG = "#F6F3EE"
CARD_BG = "#FFFDFC"
ACCENT_NAVY = "#12355B"
ACCENT_BLUE = "#2C7DA0"
ACCENT_TEAL = "#5AA9A7"
ACCENT_GOLD = "#C58B2A"
ACCENT_RED = "#B23A48"
ACCENT_GREEN = "#3D7A5E"
TEXT_COLOR = "#1E2A36"
MUTED_TEXT = "#5B6770"
GRID_COLOR = "#D7D0C8"

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


def _fmt_money_text(value: float, decimals: int = 0) -> str:
    """Return a currency string safe for Matplotlib text rendering."""
    return f"\\${value:,.{decimals}f}"


def _style_ax(ax: plt.Axes) -> None:
    ax.set_facecolor(CARD_BG)
    ax.tick_params(colors=TEXT_COLOR, labelsize=10)
    ax.xaxis.label.set_color(TEXT_COLOR)
    ax.yaxis.label.set_color(TEXT_COLOR)
    ax.title.set_color(TEXT_COLOR)
    for spine in ax.spines.values():
        spine.set_color(GRID_COLOR)
    ax.grid(axis="y", color=GRID_COLOR, linewidth=0.8, alpha=0.7)


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

    fig, ax = plt.subplots(figsize=(13.5, 7.2))
    fig.patch.set_facecolor(PAPER_BG)
    _style_ax(ax)

    p1, p99 = np.percentile(prices, [1, 99])
    display = prices[(prices >= p1) & (prices <= p99)]
    p25, p50, p75 = np.percentile(prices, [25, 50, 75])
    sim_mean = float(np.mean(prices))
    prob_above = float(np.mean(prices > current_price)) if current_price else np.nan

    ax.hist(
        display,
        bins=80,
        density=True,
        alpha=0.55,
        color=ACCENT_BLUE,
        edgecolor=CARD_BG,
        linewidth=0.6,
        label="Simulation density",
    )

    x = np.linspace(display.min(), display.max(), 500)
    density = _estimate_density(display, x, bw_scale=0.85)
    ax.plot(x, density, color=ACCENT_NAVY, linewidth=2.5, label="Smoothed density")
    ax.fill_between(x, density, color=ACCENT_TEAL, alpha=0.12)

    ax.axvspan(p25, p75, color=ACCENT_GOLD, alpha=0.08, label="Middle 50% range")
    ax.axvline(current_price, color=ACCENT_RED, linestyle="--", linewidth=2.0,
               label=f"Current price ({_fmt_money_text(current_price)})")
    ax.axvline(sim_mean, color=ACCENT_GREEN, linestyle="-.", linewidth=2.0,
               label=f"Monte Carlo mean ({_fmt_money_text(sim_mean)})")
    ax.axvline(fundamental_target, color=ACCENT_GOLD, linestyle=":", linewidth=2.2,
               label=f"Base DCF ({_fmt_money_text(fundamental_target)})")
    ax.axvline(p50, color=ACCENT_NAVY, linestyle="-", linewidth=1.5, alpha=0.85,
               label=f"Median ({_fmt_money_text(p50)})")

    ax.set_xlabel("Implied Share Price ($)", fontsize=12, fontweight="bold")
    ax.set_ylabel("Density", fontsize=12, fontweight="bold")
    ax.set_title("Monte Carlo DCF — Implied Share Price Distribution",
                 fontsize=14, fontweight="bold", pad=15)
    ax.xaxis.set_major_formatter(mticker.StrMethodFormatter("${x:,.0f}"))
    ax.text(
        0.015,
        0.96,
        (
            f"P(price > current): {prob_above:>6.1%}\n"
            f"25th-75th range : {_fmt_money_text(p25):>8s} to {_fmt_money_text(p75):>8s}\n"
            "Displayed range : trims 1st/99th percentiles"
        ),
        transform=ax.transAxes,
        va="top",
        ha="left",
        fontsize=10,
        color=TEXT_COLOR,
        family="DejaVu Sans Mono",
        bbox={"boxstyle": "round,pad=0.5", "facecolor": "#F1ECE4", "edgecolor": GRID_COLOR},
    )
    ax.legend(loc="upper right", fontsize=9, facecolor=CARD_BG,
              edgecolor=GRID_COLOR, labelcolor=TEXT_COLOR, framealpha=0.98)

    fig.tight_layout()
    fig.savefig(save_path, dpi=220, facecolor=PAPER_BG)
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
    current_price: float | None = None,
    fundamental_target: float | None = None,
    save_path: str = "percentile_table.png",
) -> None:
    """Render a valuation snapshot card as a styled image."""
    pctls = [5, 25, 50, 75, 95]
    vals = np.percentile(prices, pctls)
    mean_price = float(np.mean(prices))
    prob_above = float(np.mean(prices > current_price)) if current_price else np.nan

    fig, ax = plt.subplots(figsize=(7.5, 4.2))
    fig.patch.set_facecolor(PAPER_BG)
    ax.set_facecolor(PAPER_BG)
    ax.axis("off")

    table_data = [
        ["Current price", f"${current_price:,.2f}" if current_price is not None else "n/a"],
        ["Base DCF", f"${fundamental_target:,.2f}" if fundamental_target is not None else "n/a"],
        ["Monte Carlo mean", f"${mean_price:,.2f}"],
        ["5th / 50th / 95th", f"${vals[0]:,.2f} / ${vals[2]:,.2f} / ${vals[4]:,.2f}"],
        ["P(price > current)", f"{prob_above:.1%}" if current_price is not None else "n/a"],
    ]
    table = ax.table(
        cellText=table_data,
        colLabels=["Metric", "Value"],
        cellLoc="center",
        loc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(10.5)
    table.scale(1, 1.85)

    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor(GRID_COLOR)
        if row == 0:
            cell.set_facecolor(ACCENT_NAVY)
            cell.set_text_props(color="white", fontweight="bold")
        else:
            cell.set_facecolor(CARD_BG if row % 2 else "#F4EEE5")
            cell.set_text_props(color=TEXT_COLOR)

    ax.set_title("Valuation Snapshot", fontsize=13,
                 fontweight="bold", color=TEXT_COLOR, pad=18)
    fig.tight_layout()
    fig.savefig(save_path, dpi=220, facecolor=PAPER_BG, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {save_path}")


def plot_sobol_sensitivity(
    sobol_df: pd.DataFrame,
    save_path: str = "sobol_sensitivity.png",
) -> None:
    """Horizontal bar chart of Sobol total-order indices."""
    df = sobol_df.sort_values("ST", ascending=True)
    fig_height = max(5.8, 0.52 * len(df) + 1.8)
    fig, ax = plt.subplots(figsize=(11.5, fig_height))
    fig.patch.set_facecolor(PAPER_BG)
    _style_ax(ax)

    labels = [PARAM_LABELS.get(p, p) for p in df["Parameter"]]
    colors = plt.cm.Blues(np.linspace(0.45, 0.9, len(df)))
    bars = ax.barh(labels, df["ST"], color=colors, edgecolor="none",
                   height=0.62, alpha=0.9)
    ax.errorbar(df["ST"], labels, xerr=df["ST_conf"], fmt="none",
                ecolor=ACCENT_NAVY, elinewidth=1, capsize=3)

    max_conf = np.nan_to_num(df["ST_conf"].to_numpy(dtype=float), nan=0.0)
    max_extent = float(np.max(df["ST"].to_numpy(dtype=float) + max_conf))
    text_offset = max(0.01, max_extent * 0.02)
    ax.set_xlim(0.0, max_extent + text_offset * 6)

    for bar, val, conf in zip(bars, df["ST"], max_conf):
        ax.text(val + conf + text_offset, bar.get_y() + bar.get_height() / 2,
                f"{val:.3f}", va="center", ha="left",
                color=TEXT_COLOR, fontsize=9, clip_on=False)

    ax.set_xlabel("Total-Order Sobol Index (ST)", fontsize=12, fontweight="bold")
    ax.set_title("Sensitivity Analysis — Which Inputs Drive Valuation Variance?",
                 fontsize=13, fontweight="bold", pad=15)
    ax.grid(axis="x", color=GRID_COLOR, linewidth=0.5, alpha=0.5)
    ax.grid(axis="y", visible=False)
    ax.xaxis.set_major_formatter(mticker.StrMethodFormatter("{x:.2f}"))

    fig.tight_layout()
    fig.savefig(save_path, dpi=220, facecolor=PAPER_BG)
    plt.close(fig)
    print(f"  Saved: {save_path}")


def plot_tornado_chart(
    tornado_df: pd.DataFrame,
    base_price: float,
    save_path: str = "tornado_sensitivity.png",
) -> None:
    """Diverging tornado chart showing dollar target change around the base case."""
    df = tornado_df.sort_values("Max Abs Delta", ascending=True)
    fig_height = max(5.8, 0.52 * len(df) + 1.8)
    fig, ax = plt.subplots(figsize=(12.0, fig_height))
    fig.patch.set_facecolor(PAPER_BG)
    _style_ax(ax)

    labels = []
    for _, row in df.iterrows():
        label = str(row["Label"])
        shock_used = row.get("Shock Used")
        if shock_used:
            label = f"{label} ({shock_used})"
        labels.append(label)
    bear = df["Bear Delta"].to_numpy(dtype=float)
    bull = df["Bull Delta"].to_numpy(dtype=float)
    y = np.arange(len(df))

    ax.barh(y, bear, color=ACCENT_RED, alpha=0.82, height=0.6, label="Downside move")
    ax.barh(y, bull, color=ACCENT_GREEN, alpha=0.82, height=0.6, label="Upside move")
    ax.axvline(0.0, color=ACCENT_NAVY, linewidth=1.4)

    max_abs = float(np.max(np.abs(np.concatenate([bear, bull]))))
    text_offset = max(1.2, max_abs * 0.03)
    ax.set_xlim(-(max_abs * 1.28), max_abs * 1.28)

    for yi, bear_delta, bull_delta in zip(y, bear, bull):
        ax.text(bear_delta - text_offset, yi, f"{bear_delta:+.2f}", va="center", ha="right",
                color=TEXT_COLOR, fontsize=9, clip_on=False)
        ax.text(bull_delta + text_offset, yi, f"{bull_delta:+.2f}", va="center", ha="left",
                color=TEXT_COLOR, fontsize=9, clip_on=False)

    ax.set_yticks(y)
    ax.set_yticklabels(labels)
    ax.set_xlabel("Change In Implied Price Vs. Base DCF ($/share)", fontsize=12, fontweight="bold")
    ax.set_title(
        f"Tornado Analysis — One-Shock Price Impact Around Base DCF ({_fmt_money_text(base_price, 2)})",
        fontsize=13,
        fontweight="bold",
        pad=15,
    )
    ax.xaxis.set_major_formatter(mticker.StrMethodFormatter("${x:,.0f}"))
    ax.grid(axis="x", color=GRID_COLOR, linewidth=0.5, alpha=0.5)
    ax.grid(axis="y", visible=False)
    ax.legend(loc="lower right", fontsize=9, facecolor=CARD_BG,
              edgecolor=GRID_COLOR, labelcolor=TEXT_COLOR, framealpha=0.98)

    fig.tight_layout()
    fig.savefig(save_path, dpi=220, facecolor=PAPER_BG)
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
