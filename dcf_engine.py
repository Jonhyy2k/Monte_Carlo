"""
Phase 1 — DCF Engine
Phase 2 — Distribution Layer
Phase 3 — Monte Carlo Loop with Cholesky Correlation

Revenue → EBITDA (via COGS & SG&A margins) → EBIT (via D&A from CapEx intensity)
→ NOPAT → FCFF → discount at WACC → terminal value (exit multiple & perpetuity growth)
→ enterprise value → equity value → price per share.
"""

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
) -> float:
    """
    Deterministic DCF from a pre-computed FCFF array.

    Used as the inner loop for the projection-based Monte Carlo.
    """
    n_years = len(fcff_array)
    discount_factors = (1 + wacc) ** np.arange(1, n_years + 1)
    pv_fcff = np.sum(fcff_array / discount_factors)

    tv_exit = last_ebitda * exit_multiple
    tv_perp = last_fcff * (1 + terminal_g) / (wacc - terminal_g)
    tv = exit_multiple_weight * tv_exit + (1 - exit_multiple_weight) * tv_perp
    pv_tv = tv / discount_factors[-1]

    ev = pv_fcff + pv_tv
    equity = ev - net_debt
    return equity / shares


def run_monte_carlo_from_projections(
    dcf_data: dict,
    n_iter: int = 50_000,
    seed: int = 42,
    rho: float = 0.6,
) -> NDArray:
    """
    Monte Carlo with per-year shocks on the year-by-year P&L from Excel.

    Shock mechanics:
    - revenue_shock[t]: multiplicative on revenue
    - cogs_margin_shock[t]: additive on COGS/Revenue ratio
    - sga_margin_shock[t]: additive on SG&A/Revenue ratio
    - capex_shock[t]: multiplicative on CapEx
    - nwc_shock[t]: multiplicative on NWC change

    AR(1) autocorrelation per line item across years.
    Cross-line-item correlation via Cholesky on innovation terms.
    """
    arrays = dcf_data["arrays"]
    n_years = dcf_data["n_years"]

    # Base arrays (absolute values for negative items)
    revenue_base = arrays["Revenue"]
    cogs_base = np.abs(arrays["COGS"])
    sga_base = np.abs(arrays["SG&A"])
    da_base = np.abs(arrays["D&A"])
    capex_base = np.abs(arrays["CapEx"])
    nwc_base = arrays["Chg in NWC"]  # keep sign

    # Back-calculate implied margins per year
    cogs_pct_base = cogs_base / revenue_base
    sga_pct_base = sga_base / revenue_base
    da_pct_base = da_base / revenue_base
    tax_rate_base = np.abs(arrays["Taxes"]) / np.maximum(np.abs(arrays["EBIT"]), 1e-6)

    # Scalar assumptions
    wacc_base = dcf_data["wacc"]
    em_base = dcf_data["exit_multiple"]
    tg_base = dcf_data["terminal_g"]
    net_debt = dcf_data["net_debt"]
    shares = dcf_data["shares_outstanding"]
    em_weight = dcf_data["exit_multiple_weight"]

    # Shock std devs (reasonable defaults)
    sigma_rev = 0.05       # 5% revenue shock
    sigma_cogs = 0.02      # 2pp COGS margin shock
    sigma_sga = 0.015      # 1.5pp SG&A margin shock
    sigma_capex = 0.08     # 8% CapEx shock
    sigma_nwc = 0.10       # 10% NWC shock
    sigma_wacc = 0.015
    sigma_em = 3.0
    sigma_tg = 0.01

    # Cross-item correlation matrix for innovations (5 per-year items + 3 scalars)
    n_items = 5  # per-year shock channels
    n_scalars = 3  # wacc, exit_multiple, terminal_g
    n_total = n_items + n_scalars

    corr = np.eye(n_total)
    # indices: 0=rev, 1=cogs, 2=sga, 3=capex, 4=nwc, 5=wacc, 6=em, 7=tg
    corr[1, 2] = corr[2, 1] = 0.20   # cogs ↔ sga
    corr[0, 5] = corr[5, 0] = 0.25   # rev ↔ wacc
    corr[5, 6] = corr[6, 5] = -0.30  # wacc ↔ exit multiple
    corr[5, 7] = corr[7, 5] = 0.15   # wacc ↔ terminal g

    L = np.linalg.cholesky(corr)

    rng = np.random.default_rng(seed)
    rho_decay = np.sqrt(1 - rho ** 2)

    # Draw all innovations: (n_iter, n_years, n_items) for per-year + (n_iter, n_scalars) for scalars
    # We draw (n_iter, n_years, n_total) and correlate across the last axis
    z_raw = rng.standard_normal((n_iter, n_years, n_total))
    z_corr = z_raw @ L.T  # (n_iter, n_years, n_total)

    # Apply AR(1) to per-year items (first n_items channels)
    shocks = np.zeros((n_iter, n_years, n_items))
    for t in range(n_years):
        if t == 0:
            shocks[:, t, :] = z_corr[:, t, :n_items]
        else:
            shocks[:, t, :] = rho * shocks[:, t - 1, :] + rho_decay * z_corr[:, t, :n_items]

    # Scalar shocks: average correlated innovations across years for stability
    scalar_z = z_corr[:, 0, n_items:]  # (n_iter, 3) — use first year's draw

    wacc_draw = np.clip(wacc_base + sigma_wacc * scalar_z[:, 0], 0.04, 0.20)
    em_draw = np.clip(em_base + sigma_em * scalar_z[:, 1], 4.0, 25.0)
    tg_draw = np.clip(tg_base + sigma_tg * scalar_z[:, 2], 0.005, 0.05)

    # Hard rejection: terminal_g >= wacc
    valid = tg_draw < wacc_draw
    shocks = shocks[valid]
    wacc_draw = wacc_draw[valid]
    em_draw = em_draw[valid]
    tg_draw = tg_draw[valid]
    n_valid = valid.sum()

    # Extract per-year shocks: (n_valid, n_years)
    rev_shock = shocks[:, :, 0]
    cogs_shock = shocks[:, :, 1]
    sga_shock = shocks[:, :, 2]
    capex_shock = shocks[:, :, 3]
    nwc_shock = shocks[:, :, 4]

    # Build shocked P&L year by year — all vectorized across iterations
    # Arrays: (n_valid, n_years)
    revenue = revenue_base[np.newaxis, :] * (1 + sigma_rev * rev_shock)
    cogs_pct = np.clip(cogs_pct_base[np.newaxis, :] + sigma_cogs * cogs_shock, 0.1, 0.95)
    sga_pct = np.clip(sga_pct_base[np.newaxis, :] + sigma_sga * sga_shock, 0.01, 0.50)

    cogs = revenue * cogs_pct
    gross_profit = revenue - cogs
    sga = revenue * sga_pct
    ebitda = gross_profit - sga

    da = revenue * da_pct_base[np.newaxis, :]
    ebit = ebitda - da
    taxes = np.abs(ebit) * tax_rate_base[np.newaxis, :]
    nopat = ebit - taxes

    capex = capex_base[np.newaxis, :] * (1 + sigma_capex * capex_shock)
    nwc_change = nwc_base[np.newaxis, :] * (1 + sigma_nwc * nwc_shock)

    fcff = nopat + da - capex - nwc_change

    # Discount each year's FCFF
    discount_exp = np.arange(1, n_years + 1)[np.newaxis, :]  # (1, n_years)
    discount_factors = (1 + wacc_draw[:, np.newaxis]) ** discount_exp  # (n_valid, n_years)
    pv_fcff = np.sum(fcff / discount_factors, axis=1)  # (n_valid,)

    # Terminal value from last year
    last_ebitda = ebitda[:, -1]
    last_fcff = fcff[:, -1]

    tv_exit = last_ebitda * em_draw
    tv_perp = last_fcff * (1 + tg_draw) / (wacc_draw - tg_draw)
    tv = em_weight * tv_exit + (1 - em_weight) * tv_perp
    pv_tv = tv / (1 + wacc_draw) ** n_years

    ev = pv_fcff + pv_tv
    equity = ev - net_debt
    prices = equity / shares

    return prices


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
