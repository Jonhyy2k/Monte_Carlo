"""
Phase 4 — Sobol Sensitivity Engine

Runs SALib's Sobol analysis against run_dcf() to rank which inputs
drive the most output variance. Produces a clean DataFrame.
"""

import numpy as np
import pandas as pd
from SALib.sample import saltelli
from SALib.analyze import sobol

from dcf_engine import (
    DEFAULT_PROJECTION_SPECS,
    PARAM_NAMES,
    run_dcf,
    run_projection_family_scenario,
)


# Define the problem: bounds for each stochastic parameter
PROBLEM = {
    "num_vars": len(PARAM_NAMES),
    "names": PARAM_NAMES,
    "bounds": [
        [-0.10, 0.30],    # revenue_growth
        [0.35, 0.75],     # cogs_pct
        [0.05, 0.30],     # sga_pct
        [0.02, 0.12],     # capex_intensity
        [0.15, 0.35],     # tax_rate
        [6.0, 20.0],      # exit_multiple
        [0.005, 0.045],   # terminal_g
        [0.06, 0.16],     # wacc
    ],
}


def run_sobol(n_samples: int = 2048, seed: int = 42) -> pd.DataFrame:
    """
    Run Sobol sensitivity analysis.

    Parameters
    ----------
    n_samples : int
        Base sample size. Total model evaluations = n_samples * (2D + 2).

    Returns
    -------
    pd.DataFrame
        Columns: Parameter, S1, S1_conf, ST, ST_conf
        Sorted by ST descending.
    """
    param_values = saltelli.sample(PROBLEM, n_samples, calc_second_order=False)

    # Evaluate DCF for every Sobol sample row
    n_rows = param_values.shape[0]
    Y = np.empty(n_rows)

    for i in range(n_rows):
        row = param_values[i]
        # Reject draws where terminal_g >= wacc → assign NaN then fill with 0
        if row[6] >= row[7]:
            Y[i] = 0.0
            continue
        Y[i] = run_dcf(
            revenue_growth=row[0],
            cogs_pct=row[1],
            sga_pct=row[2],
            capex_intensity=row[3],
            tax_rate=row[4],
            exit_multiple=row[5],
            terminal_g=row[6],
            wacc=row[7],
        )

    Si = sobol.analyze(PROBLEM, Y, calc_second_order=False)

    df = pd.DataFrame({
        "Parameter": PARAM_NAMES,
        "S1": Si["S1"],
        "S1_conf": Si["S1_conf"],
        "ST": Si["ST"],
        "ST_conf": Si["ST_conf"],
    })
    df = df.sort_values("ST", ascending=False).reset_index(drop=True)
    return df


def run_sobol_projections(dcf_data: dict, n_samples: int = 1024) -> pd.DataFrame:
    """
    Sobol analysis for the imported-driver workflow.

    The dimensions are driver families, not year-by-year cells, so the output
    stays interpretable: Revenue, COGS %, SG&A %, CapEx, WACC, etc.
    """
    specs = dcf_data.get("shock_specs") or DEFAULT_PROJECTION_SPECS
    param_names = [spec.name for spec in specs]
    problem = {
        "num_vars": len(param_names),
        "names": param_names,
        "bounds": [[-2.0, 2.0] for _ in param_names],
    }

    param_values = saltelli.sample(problem, n_samples, calc_second_order=False)
    Y = np.empty(param_values.shape[0])

    for i, row in enumerate(param_values):
        scenario = {name: float(value) for name, value in zip(param_names, row)}
        Y[i] = run_projection_family_scenario(dcf_data, scenario, specs=specs)

    Si = sobol.analyze(problem, Y, calc_second_order=False)

    df = pd.DataFrame({
        "Parameter": param_names,
        "S1": Si["S1"],
        "S1_conf": Si["S1_conf"],
        "ST": Si["ST"],
        "ST_conf": Si["ST_conf"],
    })
    df = df.sort_values("ST", ascending=False).reset_index(drop=True)
    return df


if __name__ == "__main__":
    import time
    t0 = time.perf_counter()
    df = run_sobol(n_samples=2048)
    elapsed = time.perf_counter() - t0
    print("=== Phase 4: Sobol Sensitivity ===")
    print(df.to_string(index=False))
    print(f"\n  Elapsed: {elapsed:.1f}s")
