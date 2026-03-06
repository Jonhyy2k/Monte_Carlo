"""
Monte Carlo DCF Pipeline — Full End-to-End

Reads assumptions from Excel → runs Monte Carlo + Sobol → writes results
back to Excel → generates presentation-quality charts.
"""

import time
import numpy as np

from dcf_engine import (
    run_monte_carlo_vectorized,
    run_monte_carlo_from_projections,
    DistributionSpec,
    DEFAULT_SPECS,
    PARAM_NAMES,
)
from sobol_analysis import run_sobol, run_sobol_projections
from excel_io import (
    read_inputs_from_excel,
    read_dcf_from_excel,
    has_dcf_sheet,
    write_results_to_excel,
    create_stub_excel,
)
from visualizations import (
    plot_price_distribution,
    plot_percentile_table,
    plot_sobol_sensitivity,
    build_percentile_table,
)
from pathlib import Path


def _run_simple_pipeline(excel_path, n_iter, sobol_samples):
    """Original simple-parameter pipeline."""
    print("Mode: Simple (scalar assumptions)")
    inputs = read_inputs_from_excel(excel_path)

    specs = []
    for spec in DEFAULT_SPECS:
        if spec.name in inputs:
            specs.append(DistributionSpec(
                name=spec.name,
                dist_type=spec.dist_type,
                base=float(inputs[spec.name]),
                spread=spec.spread,
                low=spec.low,
                high=spec.high,
            ))
        else:
            specs.append(spec)

    print(f"Running Monte Carlo ({n_iter:,} iterations)...")
    t0 = time.perf_counter()
    prices = run_monte_carlo_vectorized(
        n_iter=n_iter,
        specs=specs,
        revenue_base=float(inputs["revenue_base"]),
        projection_years=int(inputs["projection_years"]),
        net_debt=float(inputs["net_debt"]),
        shares_outstanding=float(inputs["shares_outstanding"]),
        exit_multiple_weight=float(inputs["exit_multiple_weight"]),
    )
    elapsed = time.perf_counter() - t0
    print(f"  Done in {elapsed:.3f}s  ({len(prices):,} valid draws)")

    print(f"Running Sobol analysis ({sobol_samples} base samples)...")
    sobol_df = run_sobol(n_samples=sobol_samples)

    current_price = float(inputs.get("current_price", 55.0))
    fundamental_target = float(inputs.get("fundamental_target", 68.0))
    return prices, sobol_df, elapsed, current_price, fundamental_target


def _run_projection_pipeline(excel_path, n_iter, sobol_samples):
    """Year-by-year projection pipeline from DCF sheet."""
    print("Mode: Projection (year-by-year DCF from Excel)")
    dcf_data = read_dcf_from_excel(excel_path)
    print(f"  {dcf_data['n_years']} projection years, "
          f"{len(dcf_data['arrays'])} line items from DCF sheet")

    print(f"Running Monte Carlo ({n_iter:,} iterations)...")
    t0 = time.perf_counter()
    prices = run_monte_carlo_from_projections(dcf_data, n_iter=n_iter)
    elapsed = time.perf_counter() - t0
    print(f"  Done in {elapsed:.3f}s  ({len(prices):,} valid draws)")

    print(f"Running Sobol analysis ({sobol_samples} base samples)...")
    sobol_df = run_sobol_projections(dcf_data, n_samples=sobol_samples)

    current_price = dcf_data.get("current_price", 55.0)
    fundamental_target = dcf_data.get("target_price", 68.0)
    return prices, sobol_df, elapsed, current_price, fundamental_target


def run_pipeline(
    excel_path: str = "model_inputs.xlsx",
    n_iter: int = 50_000,
    sobol_samples: int = 2048,
) -> None:
    """Execute the full Monte Carlo DCF pipeline."""

    # ---- Ensure Excel stub exists ----
    if not Path(excel_path).exists():
        print("Creating Excel stub...")
        create_stub_excel(excel_path)

    # ---- Auto-detect mode ----
    print("Reading inputs from Excel...")
    use_projections = has_dcf_sheet(excel_path)

    if use_projections:
        prices, sobol_df, elapsed, current_price, fundamental_target = \
            _run_projection_pipeline(excel_path, n_iter, sobol_samples)
    else:
        prices, sobol_df, elapsed, current_price, fundamental_target = \
            _run_simple_pipeline(excel_path, n_iter, sobol_samples)

    # ---- Collect results ----
    results = {
        "mean": float(np.mean(prices)),
        "median": float(np.median(prices)),
        "std": float(np.std(prices)),
        "p5": float(np.percentile(prices, 5)),
        "p25": float(np.percentile(prices, 25)),
        "p50": float(np.percentile(prices, 50)),
        "p75": float(np.percentile(prices, 75)),
        "p95": float(np.percentile(prices, 95)),
        "n_iterations": n_iter,
        "n_valid": len(prices),
        "elapsed_seconds": round(elapsed, 3),
        "sobol_df": sobol_df,
    }

    # ---- Write results to Excel ----
    print("Writing results to Excel...")
    write_results_to_excel(excel_path, results)

    # ---- Generate charts ----
    print("Generating charts...")
    plot_price_distribution(prices, current_price, fundamental_target)
    plot_percentile_table(prices)
    plot_sobol_sensitivity(sobol_df)

    # ---- Summary ----
    mode_label = "PROJECTION" if use_projections else "SIMPLE"
    print("\n" + "=" * 55)
    print(f"  MONTE CARLO DCF — SUMMARY ({mode_label})")
    print("=" * 55)
    print(f"  Iterations:      {n_iter:>10,}")
    print(f"  Valid draws:     {len(prices):>10,}")
    print(f"  Runtime:         {elapsed:>10.3f}s")
    print(f"  Mean price:      ${results['mean']:>9.2f}")
    print(f"  Median price:    ${results['median']:>9.2f}")
    print(f"  Std dev:         ${results['std']:>9.2f}")
    print(f"  Current price:   ${current_price:>9.2f}")
    print(f"  Target price:    ${fundamental_target:>9.2f}")
    print()
    print(build_percentile_table(prices).to_string(index=False))
    print()
    print("  Top sensitivity drivers (Sobol ST):")
    for _, row in sobol_df.head(4).iterrows():
        print(f"    {row['Parameter']:>20s}  {row['ST']:.3f}")
    print("=" * 55)
    print("  Outputs: model_inputs.xlsx, price_distribution.png,")
    print("           percentile_table.png, sobol_sensitivity.png")
    print("=" * 55)


if __name__ == "__main__":
    run_pipeline()
