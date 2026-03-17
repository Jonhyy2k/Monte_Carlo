"""
Phase 1 — DCF Engine
Phase 2 — Distribution Layer
Phase 3 — Monte Carlo Loop with Cholesky Correlation

Revenue → EBITDA (via COGS & SG&A margins) → EBIT (via D&A from CapEx intensity)
→ NOPAT → FCFF → discount at WACC → terminal value (exit multiple & perpetuity growth)
→ enterprise value → equity value → price per share.
"""

from dataclasses import dataclass

import numpy as np
from numpy.typing import NDArray


# ---------------------------------------------------------------------------
# Phase 1 — Core DCF
# ---------------------------------------------------------------------------

def run_dcf(
    # Revenue assumptions
    revenue_base: float = 10_000.0,       # current-year revenue ($M)
    revenue_growth: float = 0.08,         # annual revenue growth rate
    projection_years: int = 5,            # explicit forecast horizon

    # Margin assumptions
    cogs_pct: float = 0.55,              # COGS as % of revenue
    sga_pct: float = 0.15,              # SG&A as % of revenue
    capex_intensity: float = 0.06,       # CapEx as % of revenue (D&A proxy)

    # Tax
    tax_rate: float = 0.25,

    # Terminal value
    exit_multiple: float = 12.0,         # EV/EBITDA exit multiple
    terminal_g: float = 0.025,           # perpetuity growth rate

    # Discount rate
    wacc: float = 0.10,

    # Capital structure / per-share
    net_debt: float = 2_000.0,           # net debt ($M)
    shares_outstanding: float = 500.0,   # millions of shares

    # Weighting between terminal value methods
    exit_multiple_weight: float = 0.5,   # weight on exit-multiple TV
) -> float:
    """
    Deterministic DCF returning implied share price.

    Terminal value is a weighted average of exit-multiple and
    perpetuity-growth approaches.
    """

    # ---- Project free cash flows ----
    revenues = np.empty(projection_years)
    ebitda = np.empty(projection_years)
    fcff = np.empty(projection_years)

    for t in range(projection_years):
        rev = revenue_base * (1 + revenue_growth) ** (t + 1)
        gross_profit = rev * (1 - cogs_pct)
        ebitda_t = gross_profit - rev * sga_pct          # EBITDA
        da = rev * capex_intensity                        # D&A ≈ CapEx intensity
        ebit = ebitda_t - da                              # EBIT
        nopat = ebit * (1 - tax_rate)                     # NOPAT
        # FCFF = NOPAT + D&A − CapEx  (net reinvestment ≈ 0 when D&A ≈ CapEx)
        capex = rev * capex_intensity
        fcff_t = nopat + da - capex
        revenues[t] = rev
        ebitda[t] = ebitda_t
        fcff[t] = fcff_t

    # ---- Discount factors ----
    discount_factors = (1 + wacc) ** np.arange(1, projection_years + 1)
    pv_fcff = np.sum(fcff / discount_factors)

    # ---- Terminal value (year N) ----
    terminal_ebitda = ebitda[-1]
    terminal_fcff = fcff[-1]

    tv_exit = terminal_ebitda * exit_multiple
    tv_perp = terminal_fcff * (1 + terminal_g) / (wacc - terminal_g)
    tv = exit_multiple_weight * tv_exit + (1 - exit_multiple_weight) * tv_perp

    pv_tv = tv / discount_factors[-1]

    # ---- Enterprise → Equity → Per Share ----
    enterprise_value = pv_fcff + pv_tv
    equity_value = enterprise_value - net_debt
    price_per_share = equity_value / shares_outstanding

    return price_per_share


# ---------------------------------------------------------------------------
# Phase 2 — Distribution Layer
# ---------------------------------------------------------------------------

class DistributionSpec:
    """Describes how to sample a single parameter."""

    __slots__ = ("name", "dist_type", "base", "spread", "low", "high")

    def __init__(
        self,
        name: str,
        dist_type: str,   # "normal", "triangular", "lognormal"
        base: float,
        spread: float,    # std-dev for normal/lognormal, half-width for triangular
        low: float = -np.inf,
        high: float = np.inf,
    ):
        self.name = name
        self.dist_type = dist_type
        self.base = base
        self.spread = spread
        self.low = low
        self.high = high

    def sample_uncorrelated(self, z: NDArray) -> NDArray:
        """
        Transform standard-normal draws *z* into the target distribution.

        For correlated sampling we feed Cholesky-transformed z here.
        """
        if self.dist_type == "normal":
            raw = self.base + self.spread * z
        elif self.dist_type == "lognormal":
            # Parameterise so that the *median* equals base.
            # sigma of the underlying normal = spread (in log-space).
            sigma = self.spread
            mu = np.log(1 + self.base) - 0.5 * sigma ** 2
            raw = np.exp(mu + sigma * z) - 1       # growth rate space
        elif self.dist_type == "triangular":
            # Map z → uniform via CDF of standard normal, then inverse-triangular.
            u = _norm_cdf(z)
            lo = self.base - self.spread
            hi = self.base + self.spread
            mid = self.base
            raw = _inv_triangular(u, lo, mid, hi)
        else:
            raise ValueError(f"Unknown distribution type: {self.dist_type}")

        return np.clip(raw, self.low, self.high)


def _norm_cdf(z: NDArray) -> NDArray:
    """Fast approximation of the standard-normal CDF (Abramowitz & Stegun)."""
    return 0.5 * (1 + np.tanh(np.sqrt(2 / np.pi) * (z + 0.044715 * z ** 3)))


def _inv_triangular(u: NDArray, lo: float, mode: float, hi: float) -> NDArray:
    """Inverse CDF of a triangular distribution."""
    fc = (mode - lo) / (hi - lo)
    out = np.empty_like(u)
    mask = u < fc
    out[mask] = lo + np.sqrt(u[mask] * (hi - lo) * (mode - lo))
    out[~mask] = hi - np.sqrt((1 - u[~mask]) * (hi - lo) * (hi - mode))
    return out


# ---- Default parameter specifications (generic company) ----

DEFAULT_SPECS: list[DistributionSpec] = [
    # Revenue growth — log-normal (bounded below −100 %)
    DistributionSpec("revenue_growth",   "lognormal",    0.08,  0.15, low=-0.30, high=0.40),
    # Margins — normal
    DistributionSpec("cogs_pct",         "normal",       0.55,  0.03, low=0.30,  high=0.80),
    DistributionSpec("sga_pct",          "normal",       0.15,  0.02, low=0.05,  high=0.35),
    DistributionSpec("capex_intensity",  "normal",       0.06,  0.01, low=0.02,  high=0.15),
    # Tax rate — normal (tight)
    DistributionSpec("tax_rate",         "normal",       0.25,  0.02, low=0.15,  high=0.35),
    # Terminal value — triangular (natural bounds)
    DistributionSpec("exit_multiple",    "triangular",   12.0,  4.0,  low=5.0,   high=25.0),
    DistributionSpec("terminal_g",       "triangular",   0.025, 0.015, low=0.005, high=0.05),
    # WACC — normal
    DistributionSpec("wacc",             "normal",       0.10,  0.015, low=0.05,  high=0.18),
]

PARAM_NAMES = [s.name for s in DEFAULT_SPECS]


@dataclass(frozen=True)
class ProjectionShockSpec:
    """Configuration for one stochastic driver in the imported DCF model."""

    name: str
    label: str
    kind: str
    spread: float
    low: float
    high: float
    applies_per_year: bool = True


DEFAULT_PROJECTION_SPECS: list[ProjectionShockSpec] = [
    ProjectionShockSpec("revenue", "Revenue", "relative", 0.05, 0.25, 2.50, True),
    ProjectionShockSpec("cogs_pct", "COGS %", "additive", 0.020, 0.05, 0.95, True),
    ProjectionShockSpec("rd_pct", "R&D %", "additive", 0.010, 0.00, 0.25, True),
    ProjectionShockSpec("sga_pct", "SG&A %", "additive", 0.015, 0.01, 0.60, True),
    ProjectionShockSpec("da_pct", "D&A %", "additive", 0.010, 0.00, 0.25, True),
    ProjectionShockSpec("tax_rate", "Tax Rate", "additive", 0.015, 0.00, 0.45, True),
    ProjectionShockSpec("capex", "CapEx", "relative", 0.080, 0.00, 3.00, True),
    ProjectionShockSpec("nwc", "Change in NWC", "relative", 0.100, 0.00, 3.00, True),
    ProjectionShockSpec("wacc", "WACC", "additive", 0.015, 0.04, 0.20, False),
    ProjectionShockSpec("exit_multiple", "Exit Multiple", "additive", 3.000, 4.00, 25.00, False),
    ProjectionShockSpec("terminal_g", "Terminal Growth", "additive", 0.010, 0.005, 0.05, False),
]

PROJECTION_SPEC_MAP = {spec.name: spec for spec in DEFAULT_PROJECTION_SPECS}
PROJECTION_PARAM_NAMES = [spec.name for spec in DEFAULT_PROJECTION_SPECS]
PROJECTION_PARAM_LABELS = {spec.name: spec.label for spec in DEFAULT_PROJECTION_SPECS}


# ---------------------------------------------------------------------------
# Phase 3 — Monte Carlo Loop
# ---------------------------------------------------------------------------

def build_correlation_matrix(n: int) -> NDArray:
    """
    Return an (n × n) correlation matrix for the stochastic parameters.

    Encodes economic intuition:
    - Revenue growth ↔ WACC: mild positive (higher-growth firms → higher risk)
    - COGS ↔ SG&A: mild positive (cost structures move together)
    - Exit multiple ↔ WACC: negative (higher discount → lower multiples)
    - Terminal g ↔ WACC: mild positive (both track macro rates)
    """
    rho = np.eye(n)
    idx = {name: i for i, name in enumerate(PARAM_NAMES)}

    def _set(a: str, b: str, v: float) -> None:
        rho[idx[a], idx[b]] = v
        rho[idx[b], idx[a]] = v

    _set("revenue_growth", "wacc",          0.25)
    _set("cogs_pct",       "sga_pct",       0.20)
    _set("exit_multiple",  "wacc",         -0.30)
    _set("terminal_g",     "wacc",          0.15)

    return rho


def run_monte_carlo(
    n_iter: int = 50_000,
    specs: list[DistributionSpec] | None = None,
    correlation: NDArray | None = None,
    # Fixed (non-stochastic) DCF parameters
    revenue_base: float = 10_000.0,
    projection_years: int = 5,
    net_debt: float = 2_000.0,
    shares_outstanding: float = 500.0,
    exit_multiple_weight: float = 0.5,
    seed: int = 42,
) -> NDArray:
    """
    Run the full Monte Carlo simulation. Returns an array of implied share
    prices, one per iteration.
    """
    if specs is None:
        specs = DEFAULT_SPECS
    if correlation is None:
        correlation = build_correlation_matrix(len(specs))

    rng = np.random.default_rng(seed)
    n_params = len(specs)

    # ---- Cholesky decomposition for correlated draws ----
    L = np.linalg.cholesky(correlation)

    # ---- Draw correlated standard normals ----
    z_indep = rng.standard_normal((n_iter, n_params))
    z_corr = z_indep @ L.T                                # (n_iter, n_params)

    # ---- Transform to target distributions ----
    draws = {}
    for j, spec in enumerate(specs):
        draws[spec.name] = spec.sample_uncorrelated(z_corr[:, j])

    # ---- Hard rejection: terminal_g >= wacc ----
    valid = draws["terminal_g"] < draws["wacc"]
    for key in draws:
        draws[key] = draws[key][valid]
    n_valid = valid.sum()

    # ---- Vectorised DCF evaluation ----
    prices = np.empty(n_valid)
    for i in range(n_valid):
        prices[i] = run_dcf(
            revenue_base=revenue_base,
            revenue_growth=draws["revenue_growth"][i],
            projection_years=projection_years,
            cogs_pct=draws["cogs_pct"][i],
            sga_pct=draws["sga_pct"][i],
            capex_intensity=draws["capex_intensity"][i],
            tax_rate=draws["tax_rate"][i],
            exit_multiple=draws["exit_multiple"][i],
            terminal_g=draws["terminal_g"][i],
            wacc=draws["wacc"][i],
            net_debt=net_debt,
            shares_outstanding=shares_outstanding,
            exit_multiple_weight=exit_multiple_weight,
        )

    return prices


def run_monte_carlo_vectorized(
    n_iter: int = 50_000,
    specs: list[DistributionSpec] | None = None,
    correlation: NDArray | None = None,
    revenue_base: float = 10_000.0,
    projection_years: int = 5,
    net_debt: float = 2_000.0,
    shares_outstanding: float = 500.0,
    exit_multiple_weight: float = 0.5,
    seed: int = 42,
) -> NDArray:
    """
    Fully vectorized Monte Carlo — no Python loop over iterations.
    """
    if specs is None:
        specs = DEFAULT_SPECS
    if correlation is None:
        correlation = build_correlation_matrix(len(specs))

    rng = np.random.default_rng(seed)
    n_params = len(specs)

    L = np.linalg.cholesky(correlation)
    z_indep = rng.standard_normal((n_iter, n_params))
    z_corr = z_indep @ L.T

    draws = {}
    for j, spec in enumerate(specs):
        draws[spec.name] = spec.sample_uncorrelated(z_corr[:, j])

    # Hard rejection: terminal_g >= wacc
    valid = draws["terminal_g"] < draws["wacc"]
    for key in draws:
        draws[key] = draws[key][valid]

    # ---- Vectorised DCF (all iterations at once) ----
    rg = draws["revenue_growth"]
    cogs = draws["cogs_pct"]
    sga = draws["sga_pct"]
    capex_int = draws["capex_intensity"]
    tax = draws["tax_rate"]
    em = draws["exit_multiple"]
    tg = draws["terminal_g"]
    w = draws["wacc"]

    pv_fcff = np.zeros(len(rg))
    last_ebitda = np.zeros(len(rg))
    last_fcff = np.zeros(len(rg))

    for t in range(projection_years):
        rev = revenue_base * (1 + rg) ** (t + 1)
        ebitda_t = rev * (1 - cogs) - rev * sga
        da = rev * capex_int
        ebit = ebitda_t - da
        nopat = ebit * (1 - tax)
        fcff_t = nopat  # + da - capex cancels when D&A ≈ CapEx
        discount = (1 + w) ** (t + 1)
        pv_fcff += fcff_t / discount
        if t == projection_years - 1:
            last_ebitda = ebitda_t
            last_fcff = fcff_t

    tv_exit = last_ebitda * em
    tv_perp = last_fcff * (1 + tg) / (w - tg)
    tv = exit_multiple_weight * tv_exit + (1 - exit_multiple_weight) * tv_perp
    pv_tv = tv / (1 + w) ** projection_years

    ev = pv_fcff + pv_tv
    equity = ev - net_debt
    prices = equity / shares_outstanding

    return prices


# ---------------------------------------------------------------------------
# Phase 7 — Projection-Based DCF (Year-by-Year from Excel)
# ---------------------------------------------------------------------------

def build_projection_correlation_matrix(
    specs: list[ProjectionShockSpec] | None = None,
) -> NDArray:
    """Correlation matrix for imported-driver shocks."""
    if specs is None:
        specs = DEFAULT_PROJECTION_SPECS

    names = [spec.name for spec in specs]
    rho = np.eye(len(names))
    idx = {name: i for i, name in enumerate(names)}

    def _set(a: str, b: str, value: float) -> None:
        if a in idx and b in idx:
            rho[idx[a], idx[b]] = value
            rho[idx[b], idx[a]] = value

    _set("revenue", "wacc", 0.25)
    _set("revenue", "exit_multiple", 0.15)
    _set("cogs_pct", "rd_pct", 0.15)
    _set("cogs_pct", "sga_pct", 0.20)
    _set("rd_pct", "sga_pct", 0.25)
    _set("da_pct", "capex", 0.30)
    _set("wacc", "exit_multiple", -0.35)
    _set("wacc", "terminal_g", 0.15)

    return rho


def _coerce_projection_specs(
    specs: list[ProjectionShockSpec] | None,
) -> list[ProjectionShockSpec]:
    if specs is None:
        return DEFAULT_PROJECTION_SPECS
    return specs


def _projection_base_values(dcf_data: dict) -> dict[str, NDArray | float]:
    arrays = dcf_data["arrays"]
    revenue = np.asarray(arrays["Revenue"], dtype=float)
    denom = np.maximum(revenue, 1e-9)

    return {
        "revenue": revenue,
        "cogs_pct": np.asarray(arrays["COGS"], dtype=float) / denom,
        "rd_pct": np.asarray(arrays["R&D"], dtype=float) / denom,
        "sga_pct": np.asarray(arrays["SG&A"], dtype=float) / denom,
        "da_pct": np.asarray(arrays["D&A"], dtype=float) / denom,
        "tax_rate": np.asarray(arrays["Tax Rate"], dtype=float),
        "capex": np.asarray(arrays["CapEx"], dtype=float),
        "nwc": np.asarray(arrays["Chg in NWC"], dtype=float),
        "wacc": float(dcf_data["wacc"]),
        "exit_multiple": float(dcf_data["exit_multiple"]),
        "terminal_g": float(dcf_data["terminal_g"]),
    }


def _apply_projection_shock(
    base: NDArray | float,
    z: NDArray,
    spec: ProjectionShockSpec,
) -> NDArray:
    if spec.kind == "relative":
        factor = np.clip(1.0 + spec.spread * z, spec.low, spec.high)
        return np.asarray(base, dtype=float) * factor
    if spec.kind == "additive":
        return np.clip(np.asarray(base, dtype=float) + spec.spread * z, spec.low, spec.high)
    raise ValueError(f"Unknown projection shock kind: {spec.kind}")


def _projection_dcf_from_draws(
    dcf_data: dict,
    draws: dict[str, NDArray],
) -> NDArray:
    revenue = draws["revenue"]
    cogs = revenue * draws["cogs_pct"]
    rd = revenue * draws["rd_pct"]
    sga = revenue * draws["sga_pct"]
    da = revenue * draws["da_pct"]

    # In the imported workbook layout, gross margin and opex margins are based on
    # reported operating income structure, so D&A is treated as a cash-flow add-back
    # rather than a separately forecast operating expense line.
    ebit = revenue - cogs - rd - sga
    ebitda = ebit + da
    nopat = ebit * (1 - draws["tax_rate"])
    fcff = nopat + da - draws["capex"] - draws["nwc"]

    n_years = revenue.shape[1]
    discount_periods = np.asarray(
        dcf_data.get("discount_periods", np.arange(1, n_years + 1)),
        dtype=float,
    )
    discount_exp = discount_periods[np.newaxis, :]
    discount_factors = (1 + draws["wacc"][:, np.newaxis]) ** discount_exp
    pv_fcff = np.sum(fcff / discount_factors, axis=1)

    last_ebitda = ebitda[:, -1]
    last_fcff = fcff[:, -1]
    tv_exit = last_ebitda * draws["exit_multiple"]
    tv_perp = last_fcff * (1 + draws["terminal_g"]) / (draws["wacc"] - draws["terminal_g"])
    tv = (
        dcf_data["exit_multiple_weight"] * tv_exit
        + (1 - dcf_data["exit_multiple_weight"]) * tv_perp
    )
    pv_tv = tv / discount_factors[:, -1]

    ev = pv_fcff + pv_tv
    equity = ev - dcf_data["net_debt"]
    return equity / dcf_data["shares_outstanding"]


def run_dcf_from_projections(
    fcff_array: NDArray,
    last_ebitda: float,
    last_fcff: float,
    wacc: float = 0.10,
    exit_multiple: float = 12.0,
    terminal_g: float = 0.025,
    net_debt: float = 2_000.0,
    shares: float = 500.0,
    exit_multiple_weight: float = 0.50,
    discount_periods: NDArray | None = None,
) -> float:
    """
    Deterministic DCF from a pre-computed FCFF array.

    Used as the inner loop for the projection-based Monte Carlo.
    """
    n_years = len(fcff_array)
    if discount_periods is None:
        discount_periods = np.arange(1, n_years + 1)
    discount_factors = (1 + wacc) ** np.asarray(discount_periods, dtype=float)
    pv_fcff = np.sum(fcff_array / discount_factors)

    tv_exit = last_ebitda * exit_multiple
    tv_perp = last_fcff * (1 + terminal_g) / (wacc - terminal_g)
    tv = exit_multiple_weight * tv_exit + (1 - exit_multiple_weight) * tv_perp
    pv_tv = tv / discount_factors[-1]

    ev = pv_fcff + pv_tv
    equity = ev - net_debt
    return equity / shares


def run_base_dcf_from_imported_data(dcf_data: dict) -> float:
    """Deterministic valuation of the imported DCF case with no stochastic shocks."""
    base = _projection_base_values(dcf_data)
    draws = {}
    for name in ("revenue", "cogs_pct", "rd_pct", "sga_pct", "da_pct", "tax_rate", "capex", "nwc"):
        draws[name] = np.asarray(base[name], dtype=float)[np.newaxis, :]
    for name in ("wacc", "exit_multiple", "terminal_g"):
        draws[name] = np.array([float(base[name])], dtype=float)
    return float(_projection_dcf_from_draws(dcf_data, draws)[0])


def run_projection_family_scenario(
    dcf_data: dict,
    shock_vector: dict[str, float],
    specs: list[ProjectionShockSpec] | None = None,
) -> float:
    """Deterministic valuation for one family-level shock scenario."""
    specs = _coerce_projection_specs(specs or dcf_data.get("shock_specs"))
    base = _projection_base_values(dcf_data)
    draws: dict[str, NDArray] = {}

    for spec in specs:
        amplitude = float(shock_vector.get(spec.name, 0.0))
        if spec.applies_per_year:
            shocked = _apply_projection_shock(
                base[spec.name],
                np.full(dcf_data["n_years"], amplitude, dtype=float),
                spec,
            )
            draws[spec.name] = np.asarray(shocked, dtype=float)[np.newaxis, :]
        else:
            shocked = _apply_projection_shock(base[spec.name], np.array([amplitude]), spec)
            draws[spec.name] = np.asarray(shocked, dtype=float)

    if draws["terminal_g"][0] >= draws["wacc"][0]:
        return 0.0

    return float(_projection_dcf_from_draws(dcf_data, draws)[0])


def run_monte_carlo_from_projections(
    dcf_data: dict,
    n_iter: int = 50_000,
    seed: int = 42,
    rho: float = 0.6,
    specs: list[ProjectionShockSpec] | None = None,
    correlation: NDArray | None = None,
) -> NDArray:
    """
    Monte Carlo over an imported DCF export.

    The detailed DCF stays in Excel. This layer perturbs the exported yearly
    drivers and re-runs the valuation.
    """
    specs = _coerce_projection_specs(specs or dcf_data.get("shock_specs"))
    if correlation is None:
        correlation = build_projection_correlation_matrix(specs)

    base = _projection_base_values(dcf_data)
    n_years = dcf_data["n_years"]
    n_params = len(specs)
    path_specs = [spec for spec in specs if spec.applies_per_year]
    scalar_specs = [spec for spec in specs if not spec.applies_per_year]
    n_path = len(path_specs)

    rng = np.random.default_rng(seed)
    rho_decay = np.sqrt(1 - rho ** 2)
    L = np.linalg.cholesky(correlation)
    z_raw = rng.standard_normal((n_iter, n_years, n_params))
    z_corr = z_raw @ L.T

    path_z = np.zeros((n_iter, n_years, n_path))
    for t in range(n_years):
        innovations = z_corr[:, t, :n_path]
        if t == 0:
            path_z[:, t, :] = innovations
        else:
            path_z[:, t, :] = rho * path_z[:, t - 1, :] + rho_decay * innovations

    draws: dict[str, NDArray] = {}
    for j, spec in enumerate(path_specs):
        draws[spec.name] = _apply_projection_shock(base[spec.name], path_z[:, :, j], spec)

    for j, spec in enumerate(scalar_specs):
        scalar_z = z_corr[:, 0, n_path + j]
        draws[spec.name] = _apply_projection_shock(base[spec.name], scalar_z, spec)

    valid = draws["terminal_g"] < draws["wacc"]
    for name, values in draws.items():
        draws[name] = values[valid]

    if not np.any(valid):
        return np.array([], dtype=float)

    return _projection_dcf_from_draws(dcf_data, draws)


# ---------------------------------------------------------------------------
# Quick validation
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # --- Phase 1 edge-case checks ---
    base = run_dcf()
    high_wacc = run_dcf(wacc=0.30)
    near_wacc_g = run_dcf(terminal_g=0.099, wacc=0.10)

    print("=== Phase 1: DCF Edge Cases ===")
    print(f"  Base case price:          ${base:>10.2f}")
    print(f"  WACC = 30% price:         ${high_wacc:>10.2f}")
    print(f"  terminal_g ≈ WACC price:  ${near_wacc_g:>10.2f}")
    assert high_wacc < base, "High WACC should compress price"
    assert near_wacc_g > base * 5, "terminal_g near WACC should explode price"
    print("  ✓ Edge cases passed.\n")

    # --- Phase 3 timing ---
    import time
    t0 = time.perf_counter()
    prices = run_monte_carlo_vectorized(n_iter=50_000)
    elapsed = time.perf_counter() - t0
    print("=== Phase 3: Monte Carlo (vectorized) ===")
    print(f"  Valid iterations: {len(prices):,}")
    print(f"  Mean price:  ${np.mean(prices):>8.2f}")
    print(f"  Median:      ${np.median(prices):>8.2f}")
    print(f"  Std dev:     ${np.std(prices):>8.2f}")
    print(f"  5th pctl:    ${np.percentile(prices, 5):>8.2f}")
    print(f"  95th pctl:   ${np.percentile(prices, 95):>8.2f}")
    print(f"  Elapsed:     {elapsed:.3f}s")
    assert elapsed < 10, f"Too slow: {elapsed:.1f}s"
    print("  ✓ Under 10 s.\n")

    # --- Projection-mode test ---
    print("=== Phase 7: Projection-Based DCF ===")
    # Build a simple 5-year FCFF array from defaults
    from excel_io import _build_dcf_dummy_data, DCF_SCALAR_ASSUMPTIONS
    dummy = _build_dcf_dummy_data(5)
    fcff_arr = np.array(dummy["FCFF"])
    last_eb = dummy["EBITDA"][-1]
    last_fc = dummy["FCFF"][-1]
    det_price = run_dcf_from_projections(
        fcff_arr, last_eb, last_fc,
        wacc=0.10, exit_multiple=12.0, terminal_g=0.025,
        net_debt=2000.0, shares=500.0, exit_multiple_weight=0.50,
    )
    print(f"  Deterministic projection price: ${det_price:.2f}")

    # Quick MC from projections
    from excel_io import create_stub_excel, read_dcf_from_excel
    create_stub_excel("_test_proj.xlsx")
    dcf_data = read_dcf_from_excel("_test_proj.xlsx")
    t0 = time.perf_counter()
    proj_prices = run_monte_carlo_from_projections(dcf_data, n_iter=10_000)
    elapsed_proj = time.perf_counter() - t0
    print(f"  Projection MC: {len(proj_prices):,} valid, mean=${np.mean(proj_prices):.2f}, "
          f"median=${np.median(proj_prices):.2f}, {elapsed_proj:.3f}s")
    import os; os.remove("_test_proj.xlsx")
    print("  ✓ Projection mode passed.\n")
