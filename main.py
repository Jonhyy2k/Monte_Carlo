"""
Monte Carlo DCF Pipeline — Full End-to-End

Reads assumptions from Excel → runs Monte Carlo + Sobol → writes results
back to Excel → generates presentation-quality charts.
"""

import argparse
import time
from pathlib import Path

import numpy as np

from dcf_engine import (
    DistributionSpec,
    DEFAULT_SPECS,
    run_base_dcf_from_imported_data,
    run_monte_carlo_from_projections,
    run_monte_carlo_vectorized,
)
from sobol_analysis import run_sobol, run_sobol_projections, run_tornado_projections
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
    plot_tornado_chart,
    build_percentile_table,
)


LEGACY_STUB_WORKBOOK = "model_inputs.xlsx"


def _resolve_excel_path(excel_path: str | None) -> Path:
    if excel_path:
        return Path(excel_path)

    workbooks = [
        path for path in sorted(Path.cwd().glob("*.xlsx"))
        if path.is_file() and not path.name.startswith("~$")
    ]
    preferred = [path for path in workbooks if path.name != LEGACY_STUB_WORKBOOK and has_dcf_sheet(str(path))]
    if preferred:
        return preferred[0]
    if Path(LEGACY_STUB_WORKBOOK).exists():
        return Path(LEGACY_STUB_WORKBOOK)
    return Path(LEGACY_STUB_WORKBOOK)


def _build_output_paths(excel_path: Path) -> dict[str, Path]:
    output_dir = Path("outputs")
    output_dir.mkdir(exist_ok=True)
    stem = excel_path.stem.replace(" ", "_")
    return {
        "distribution": output_dir / f"{stem}_price_distribution.png",
        "percentiles": output_dir / f"{stem}_valuation_snapshot.png",
        "sobol": output_dir / f"{stem}_sobol_sensitivity.png",
        "tornado": output_dir / f"{stem}_tornado_sensitivity.png",
    }


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
    return prices, sobol_df, None, elapsed, current_price, fundamental_target


def _run_projection_pipeline(excel_path, n_iter, sobol_samples):
    """Projection pipeline driven by an imported Excel DCF export."""
    print("Mode: Imported DCF (year-by-year export from Excel)")
    dcf_data = read_dcf_from_excel(excel_path)
    print(f"  {dcf_data['n_years']} projection years, "
          f"{len(dcf_data['arrays'])} imported driver series, "
          f"{len(dcf_data['shock_specs'])} stochastic driver families")

    print(f"Running Monte Carlo ({n_iter:,} iterations)...")
    t0 = time.perf_counter()
    prices = run_monte_carlo_from_projections(dcf_data, n_iter=n_iter)
    elapsed = time.perf_counter() - t0
    print(f"  Done in {elapsed:.3f}s  ({len(prices):,} valid draws)")

    print(f"Running Sobol analysis ({sobol_samples} base samples)...")
    sobol_df = run_sobol_projections(dcf_data, n_samples=sobol_samples)
    tornado_df = run_tornado_projections(dcf_data)

    current_price = dcf_data.get("current_price", 55.0)
    fundamental_target = dcf_data.get("target_price", 68.0)
    base_case = run_base_dcf_from_imported_data(dcf_data)
    print(f"  Base imported DCF price: ${base_case:.2f}")
    return prices, sobol_df, tornado_df, elapsed, current_price, fundamental_target


def run_pipeline(
    excel_path: str | None = None,
    n_iter: int = 50_000,
    sobol_samples: int = 2048,
) -> None:
    """Execute the full Monte Carlo DCF pipeline."""
    workbook_path = _resolve_excel_path(excel_path)

    # ---- Ensure Excel stub exists ----
    if not workbook_path.exists():
        print("Creating Excel stub...")
        create_stub_excel(str(workbook_path))

    # ---- Auto-detect mode ----
    print("Reading inputs from Excel...")
    print(f"Workbook: {workbook_path}")
    use_projections = has_dcf_sheet(str(workbook_path))

    if use_projections:
        prices, sobol_df, tornado_df, elapsed, current_price, fundamental_target = \
            _run_projection_pipeline(str(workbook_path), n_iter, sobol_samples)
    else:
        prices, sobol_df, tornado_df, elapsed, current_price, fundamental_target = \
            _run_simple_pipeline(str(workbook_path), n_iter, sobol_samples)

    output_paths = _build_output_paths(workbook_path)
    prob_above_current = float(np.mean(prices > current_price)) if current_price is not None else None
    current_price_percentile = float(np.mean(prices <= current_price)) if current_price is not None else None
    mean_price = float(np.mean(prices))
    mean_upside = (mean_price / current_price - 1.0) if current_price else None
    target_upside = (fundamental_target / current_price - 1.0) if current_price else None

    # ---- Collect results ----
    results = {
        "mean": mean_price,
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
        "sobol_samples": sobol_samples,
        "sobol_df": sobol_df,
        "tornado_df": tornado_df,
        "current_price": float(current_price) if current_price is not None else None,
        "target_price": float(fundamental_target) if fundamental_target is not None else None,
        "prob_above_current": prob_above_current,
        "current_price_percentile": current_price_percentile,
        "mean_upside": mean_upside,
        "target_upside": target_upside,
    }

    # ---- Generate charts ----
    print("Generating charts...")
    plot_price_distribution(
        prices,
        current_price,
        fundamental_target,
        save_path=str(output_paths["distribution"]),
    )
    plot_percentile_table(
        prices,
        current_price=current_price,
        fundamental_target=fundamental_target,
        save_path=str(output_paths["percentiles"]),
    )
    plot_sobol_sensitivity(sobol_df, save_path=str(output_paths["sobol"]))
    if tornado_df is not None:
        plot_tornado_chart(
            tornado_df,
            base_price=fundamental_target,
            save_path=str(output_paths["tornado"]),
        )

    # ---- Write results to Excel ----
    print("Writing results to Excel...")
    write_results_to_excel(
        str(workbook_path),
        results,
        chart_paths=output_paths,
    )

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
    print(f"  Outputs: {workbook_path},")
    print(f"           {output_paths['distribution']},")
    print(f"           {output_paths['percentiles']},")
    print(f"           {output_paths['sobol']},")
    if tornado_df is not None:
        print(f"           {output_paths['tornado']}")
    print("=" * 55)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run Monte Carlo DCF on an Excel workbook.")
    parser.add_argument(
        "--excel",
        default=None,
        help="Path to the Excel workbook. If omitted, the script auto-detects the stock-pitch workbook in this directory.",
    )
    parser.add_argument(
        "--n-iter",
        type=int,
        default=50_000,
        help="Number of Monte Carlo iterations.",
    )
    parser.add_argument(
        "--sobol-samples",
        type=int,
        default=2048,
        help="Base Sobol sample size.",
    )
    args = parser.parse_args()
    run_pipeline(
        excel_path=args.excel,
        n_iter=args.n_iter,
        sobol_samples=args.sobol_samples,
    )
