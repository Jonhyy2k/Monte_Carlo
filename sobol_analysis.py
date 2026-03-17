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
    PROJECTION_SPEC_MAP,
    _projection_base_values,
    _projection_dcf_from_draws,
    run_dcf,
    run_base_dcf_from_imported_data,
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

PRESENTATION_TORNADO_STEPS = {
    "revenue": {"delta": 0.02, "mode": "relative", "label": "Revenue", "step_label": "+/-2.0%"},
    "cogs_pct": {"delta": 0.02, "mode": "additive", "label": "COGS %", "step_label": "+/-2.0 pts"},
    "rd_pct": {"delta": 0.01, "mode": "additive", "label": "R&D %", "step_label": "+/-1.0 pt"},
    "sga_pct": {"delta": 0.01, "mode": "additive", "label": "SG&A %", "step_label": "+/-1.0 pt"},
    "da_pct": {"delta": 0.01, "mode": "additive", "label": "D&A %", "step_label": "+/-1.0 pt"},
    "tax_rate": {"delta": 0.01, "mode": "additive", "label": "Tax Rate", "step_label": "+/-1.0 pt"},
    "capex": {"delta": 0.01, "mode": "capex_intensity", "label": "CapEx Intensity", "step_label": "+/-1.0 pt"},
    "nwc": {"delta": 0.10, "mode": "relative", "label": "Change in NWC", "step_label": "+/-10.0%"},
    "wacc": {"delta": 0.01, "mode": "additive", "label": "WACC", "step_label": "+/-100 bps"},
    "exit_multiple": {"delta": 1.0, "mode": "additive", "label": "Exit Multiple", "step_label": "+/-1.0x"},
    "terminal_g": {"delta": 0.005, "mode": "additive", "label": "Terminal Growth", "step_label": "+/-50 bps"},
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


def run_tornado_projections(
    dcf_data: dict,
) -> pd.DataFrame:
    """
    One-at-a-time tornado analysis around the imported DCF base case.

    Each driver is shocked up and down by a fixed, presentation-friendly delta
    expressed in real units: e.g. COGS +/- 2 pts, SG&A +/- 1 pt, WACC +/- 100 bps.
    The output is the price impact in dollars per share, not a percentage sensitivity.
    """
    specs = dcf_data.get("shock_specs") or DEFAULT_PROJECTION_SPECS
    base_price = run_base_dcf_from_imported_data(dcf_data)
    base = _projection_base_values(dcf_data)
    rows = []

    for spec in specs:
        step_cfg = PRESENTATION_TORNADO_STEPS.get(spec.name)
        if step_cfg is None:
            continue

        delta = float(step_cfg["delta"])
        mode = str(step_cfg["mode"])

        def _scenario_price(direction: float) -> float:
            draws = {}
            revenue = np.asarray(base["revenue"], dtype=float).copy()
            cogs_pct = np.asarray(base["cogs_pct"], dtype=float).copy()
            rd_pct = np.asarray(base["rd_pct"], dtype=float).copy()
            sga_pct = np.asarray(base["sga_pct"], dtype=float).copy()
            da_pct = np.asarray(base["da_pct"], dtype=float).copy()
            tax_rate = np.asarray(base["tax_rate"], dtype=float).copy()
            capex = np.asarray(base["capex"], dtype=float).copy()
            nwc = np.asarray(base["nwc"], dtype=float).copy()
            wacc = float(base["wacc"])
            exit_multiple = float(base["exit_multiple"])
            terminal_g = float(base["terminal_g"])

            if spec.name == "revenue":
                revenue = np.maximum(revenue * (1.0 + direction * delta), 1e-9)
            elif spec.name == "cogs_pct":
                cogs_pct = np.clip(cogs_pct + direction * delta, PROJECTION_SPEC_MAP["cogs_pct"].low, PROJECTION_SPEC_MAP["cogs_pct"].high)
            elif spec.name == "rd_pct":
                rd_pct = np.clip(rd_pct + direction * delta, PROJECTION_SPEC_MAP["rd_pct"].low, PROJECTION_SPEC_MAP["rd_pct"].high)
            elif spec.name == "sga_pct":
                sga_pct = np.clip(sga_pct + direction * delta, PROJECTION_SPEC_MAP["sga_pct"].low, PROJECTION_SPEC_MAP["sga_pct"].high)
            elif spec.name == "da_pct":
                da_pct = np.clip(da_pct + direction * delta, PROJECTION_SPEC_MAP["da_pct"].low, PROJECTION_SPEC_MAP["da_pct"].high)
            elif spec.name == "tax_rate":
                tax_rate = np.clip(tax_rate + direction * delta, PROJECTION_SPEC_MAP["tax_rate"].low, PROJECTION_SPEC_MAP["tax_rate"].high)
            elif spec.name == "capex" and mode == "capex_intensity":
                base_capex_pct = capex / np.maximum(np.asarray(base["revenue"], dtype=float), 1e-9)
                capex_pct = np.clip(base_capex_pct + direction * delta, 0.0, 0.60)
                capex = revenue * capex_pct
            elif spec.name == "capex":
                capex = np.maximum(capex * (1.0 + direction * delta), 0.0)
            elif spec.name == "nwc":
                nwc = np.maximum(nwc * (1.0 + direction * delta), 0.0)
            elif spec.name == "wacc":
                wacc = float(np.clip(wacc + direction * delta, PROJECTION_SPEC_MAP["wacc"].low, PROJECTION_SPEC_MAP["wacc"].high))
            elif spec.name == "exit_multiple":
                exit_multiple = float(np.clip(
                    exit_multiple + direction * delta,
                    PROJECTION_SPEC_MAP["exit_multiple"].low,
                    PROJECTION_SPEC_MAP["exit_multiple"].high,
                ))
            elif spec.name == "terminal_g":
                terminal_g = float(np.clip(
                    terminal_g + direction * delta,
                    PROJECTION_SPEC_MAP["terminal_g"].low,
                    PROJECTION_SPEC_MAP["terminal_g"].high,
                ))

            if terminal_g >= wacc:
                return np.nan

            draws["revenue"] = revenue[np.newaxis, :]
            draws["cogs_pct"] = cogs_pct[np.newaxis, :]
            draws["rd_pct"] = rd_pct[np.newaxis, :]
            draws["sga_pct"] = sga_pct[np.newaxis, :]
            draws["da_pct"] = da_pct[np.newaxis, :]
            draws["tax_rate"] = tax_rate[np.newaxis, :]
            draws["capex"] = capex[np.newaxis, :]
            draws["nwc"] = nwc[np.newaxis, :]
            draws["wacc"] = np.array([wacc], dtype=float)
            draws["exit_multiple"] = np.array([exit_multiple], dtype=float)
            draws["terminal_g"] = np.array([terminal_g], dtype=float)
            return float(_projection_dcf_from_draws(dcf_data, draws)[0])

        low_price = _scenario_price(-1.0)
        high_price = _scenario_price(1.0)
        low_delta = low_price - base_price
        high_delta = high_price - base_price
        bear_delta = min(low_delta, high_delta)
        bull_delta = max(low_delta, high_delta)

        rows.append(
            {
                "Parameter": spec.name,
                "Label": str(step_cfg["label"]),
                "Shock Used": str(step_cfg["step_label"]),
                "Bear Delta": bear_delta,
                "Bull Delta": bull_delta,
                "Bear Price": base_price + bear_delta,
                "Bull Price": base_price + bull_delta,
                "Max Abs Delta": max(abs(bear_delta), abs(bull_delta)),
            }
        )

    df = pd.DataFrame(rows)
    return df.sort_values("Max Abs Delta", ascending=False).reset_index(drop=True)


if __name__ == "__main__":
    import time
    t0 = time.perf_counter()
    df = run_sobol(n_samples=2048)
    elapsed = time.perf_counter() - t0
    print("=== Phase 4: Sobol Sensitivity ===")
    print(df.to_string(index=False))
    print(f"\n  Elapsed: {elapsed:.1f}s")
