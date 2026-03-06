"""
Phase 4 — Sobol Sensitivity Engine

Runs SALib's Sobol analysis against run_dcf() to rank which inputs
drive the most output variance. Produces a clean DataFrame.
"""

import numpy as np
import pandas as pd
from SALib.sample import saltelli
from SALib.analyze import sobol

from dcf_engine import run_dcf, run_dcf_from_projections, PARAM_NAMES


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
    param_values = saltelli.sample(PROBLEM, n_samples)

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

    Si = sobol.analyze(PROBLEM, Y)

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
    Sobol analysis for projection mode.

    Defines the problem over per-year shock magnitudes (revenue, cogs, sga,
    capex, nwc for each year) plus scalar params (wacc, exit_multiple, terminal_g).
    """
    arrays = dcf_data["arrays"]
    n_years = dcf_data["n_years"]

    # Back-calculate base margins
    revenue_base = arrays["Revenue"]
    cogs_pct_base = np.abs(arrays["COGS"]) / revenue_base
    sga_pct_base = np.abs(arrays["SG&A"]) / revenue_base
    da_pct_base = np.abs(arrays["D&A"]) / revenue_base
    tax_rate_base = np.abs(arrays["Taxes"]) / np.maximum(np.abs(arrays["EBIT"]), 1e-6)
    capex_base = np.abs(arrays["CapEx"])
    nwc_base = arrays["Chg in NWC"]

    # Parameter names: per-year shocks + scalars
    per_year_items = ["rev_shock", "cogs_shock", "sga_shock", "capex_shock", "nwc_shock"]
    param_names = []
    for item in per_year_items:
        for t in range(n_years):
            param_names.append(f"{item}_y{t+1}")
    param_names.extend(["wacc", "exit_multiple", "terminal_g"])

    # Bounds: shocks are in [-2, 2] std units, scalars have physical bounds
    bounds = []
    for item in per_year_items:
        for _ in range(n_years):
            bounds.append([-2.0, 2.0])
    bounds.append([0.06, 0.16])    # wacc
    bounds.append([6.0, 20.0])     # exit_multiple
    bounds.append([0.005, 0.045])  # terminal_g

    problem = {
        "num_vars": len(param_names),
        "names": param_names,
        "bounds": bounds,
    }

    param_values = saltelli.sample(problem, n_samples)
    n_rows = param_values.shape[0]
    Y = np.empty(n_rows)

    sigma_rev = 0.05
    sigma_cogs = 0.02
    sigma_sga = 0.015
    sigma_capex = 0.08
    sigma_nwc = 0.10
    net_debt = dcf_data["net_debt"]
    shares = dcf_data["shares_outstanding"]
    em_weight = dcf_data["exit_multiple_weight"]

    for i in range(n_rows):
        row = param_values[i]
        idx = 0

        # Decode per-year shocks
        rev_s = row[idx:idx + n_years]; idx += n_years
        cogs_s = row[idx:idx + n_years]; idx += n_years
        sga_s = row[idx:idx + n_years]; idx += n_years
        capex_s = row[idx:idx + n_years]; idx += n_years
        nwc_s = row[idx:idx + n_years]; idx += n_years

        wacc_val = row[idx]; idx += 1
        em_val = row[idx]; idx += 1
        tg_val = row[idx]; idx += 1

        if tg_val >= wacc_val:
            Y[i] = 0.0
            continue

        # Build shocked P&L
        revenue = revenue_base * (1 + sigma_rev * rev_s)
        cogs_pct = np.clip(cogs_pct_base + sigma_cogs * cogs_s, 0.1, 0.95)
        sga_pct = np.clip(sga_pct_base + sigma_sga * sga_s, 0.01, 0.50)

        gross = revenue * (1 - cogs_pct)
        sga = revenue * sga_pct
        ebitda = gross - sga
        da = revenue * da_pct_base
        ebit = ebitda - da
        taxes = np.abs(ebit) * tax_rate_base
        nopat = ebit - taxes

        capex = capex_base * (1 + sigma_capex * capex_s)
        nwc_change = nwc_base * (1 + sigma_nwc * nwc_s)
        fcff = nopat + da - capex - nwc_change

        Y[i] = run_dcf_from_projections(
            fcff, ebitda[-1], fcff[-1],
            wacc=wacc_val, exit_multiple=em_val, terminal_g=tg_val,
            net_debt=net_debt, shares=shares, exit_multiple_weight=em_weight,
        )

    Si = sobol.analyze(problem, Y)

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
