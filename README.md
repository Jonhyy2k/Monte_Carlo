# Monte Carlo DCF Workbook Layer

This repo runs a Monte Carlo and sensitivity layer on top of a real Excel DCF workbook.

The core idea is:
- keep the full operating model in Excel
- export or maintain the key DCF drivers in the workbook
- let Python shock those drivers, rebuild the valuation, and produce charts plus a `Results` dashboard

This is not a "simulate the price directly" toy model. The simulation perturbs the DCF building blocks and recalculates implied value from fundamentals each iteration.

## What It Does

- Reads the real workbook from Excel, primarily [`Stock Pitch - mine Excel Model-1.xlsx`](./Stock%20Pitch%20-%20mine%20Excel%20Model-1.xlsx)
- Pulls year-by-year DCF drivers from the `DCF Model`, `WACC`, `OP`, and `MC Assumptions` sheets
- Runs Monte Carlo on:
  - Revenue
  - `COGS %`
  - `R&D %`
  - `SG&A %`
  - `D&A %`
  - Tax rate
  - CapEx
  - Change in NWC
  - WACC
  - Exit multiple
  - Terminal growth
- Rebuilds FCFF and implied share price for every iteration
- Runs Sobol sensitivity analysis for global variance attribution
- Runs a tornado analysis for point-estimate dollar sensitivity
- Writes a `Results` dashboard back into the workbook
- Saves pitch-ready charts into `outputs/`

## Main Files

- `main.py`: top-level pipeline and CLI
- `excel_io.py`: workbook parsing and `Results` sheet writing
- `dcf_engine.py`: DCF math, projection shocks, and Monte Carlo engine
- `sobol_analysis.py`: Sobol sensitivity and tornado analysis
- `visualizations.py`: chart generation
- `update_stock_pitch_workbook.py`: workbook patcher for formula and integration fixes

## How To Run

Auto-detect the workbook in the current directory:

```bash
python main.py --n-iter 100000 --sobol-samples 8192
```

Explicit workbook path:

```bash
python main.py --excel "Stock Pitch - mine Excel Model-1.xlsx" --n-iter 100000 --sobol-samples 8192
```

Quick test run:

```bash
python main.py --n-iter 5000 --sobol-samples 256
```

Patch the workbook structure if needed:

```bash
python update_stock_pitch_workbook.py --workbook "Stock Pitch - mine Excel Model-1.xlsx"
```

## Workbook Requirements

The current parser expects the same workbook structure used in the stock pitch model.

Required sheets:
- `DCF Model`
- `WACC`
- `OP`
- `MC Assumptions` for editable shock settings

You can safely change assumptions and operating inputs in the existing layout.

Do not change these without updating the parser:
- sheet names
- row layout on `DCF Model`
- the five-year projection structure
- the capital structure bridge layout

## Key Workbook Inputs

Important `DCF Model` rows:
- Row 8: revenue growth
- Row 11: gross margin
- Row 14: `R&D %`
- Row 17: `SG&A %`
- Row 25: `D&A %`
- Row 31: `CapEx %`
- Row 34: `Change in NWC % of delta revenue`
- Row 42: tax rate
- Row 52: discount periods
- `D57`: WACC
- `D58`: terminal growth
- `D59`: exit multiple
- `D64`: TV method
- `D75:D78`: bridge to equity
- `D82`: shares outstanding
- `D85`: current price

Related workbook inputs:
- `OP!U6`: current share price
- `OP!U11`: market cap
- `OP!U17`: beta
- `WACC!G9`, `G12`, `G15`, `G17`, `G18`, `G22`: WACC stack

## How The Simulation Works

For each iteration:

1. Read the base projection drivers from the workbook.
2. Draw correlated shocks across the driver families.
3. Apply shocks to the DCF inputs, not directly to the final share price.
4. Recompute FCFF:

```text
FCFF = NOPAT + D&A - CapEx - Change in NWC
```

5. Discount explicit cash flows using the workbook discount periods.
6. Compute terminal value using the workbook TV method.
7. Convert enterprise value to equity value and then to implied price per share.

That means the distribution is driven by economics, margins, discount rate, and terminal assumptions, not arbitrary noise around a target price.

## Sensitivity Outputs

### Sobol

Sobol total-order sensitivity answers:

> How much of the total variance in valuation is explained by each driver?

Interpretation:
- higher `ST` = more important for overall valuation uncertainty
- `ST` is not directional
- `ST` does not mean upside or downside
- total-order effects include interactions, so values do not need to sum to 100%

### Tornado

Tornado analysis answers:

> If I move one driver by a fixed amount and hold the rest unchanged, how many dollars does the target price move?

Current presentation steps are defined in `sobol_analysis.py` under `PRESENTATION_TORNADO_STEPS`.

Examples:
- `COGS %`: `+/-2.0 pts`
- `SG&A %`: `+/-1.0 pt`
- `CapEx Intensity`: `+/-1.0 pt`
- `WACC`: `+/-100 bps`
- `Exit Multiple`: `+/-1.0x`

## Outputs

Excel:
- `Results` sheet embedded into the workbook

Charts in `outputs/`:
- `*_price_distribution.png`
- `*_valuation_snapshot.png`
- `*_sobol_sensitivity.png`
- `*_tornado_sensitivity.png`

## Special Features

- Workbook-first workflow: Excel is the input surface
- Driver-level Monte Carlo instead of price-level Monte Carlo
- Correlated shocks across operating and valuation drivers
- Family-level Sobol sensitivity instead of one-cell-per-year noise
- Tornado chart in actual dollar price-target impact
- Results dashboard written back into Excel with embedded charts
- Workbook auto-detection from the current directory

## Technical Notes

- The Monte Carlo engine is vectorized for speed.
- The chart files are workbook-specific and stored under `outputs/`.
- `TV Method = 2` means exit-multiple terminal value; in that case `terminal_g` should have near-zero sensitivity.
- CapEx, tax, and NWC can appear less important than expected when terminal value dominates and the model uses exit multiple.
- Excel recalculates workbook formulas on open; Python does not run Excel's calc engine.

## Practical Workflow

1. Update assumptions in the workbook.
2. Save the workbook.
3. Run:

```bash
python main.py --n-iter 100000 --sobol-samples 8192
```

4. Review:
- the `Results` sheet in Excel
- the charts in `outputs/`
- the console summary

## If You Want To Customize Further

- Change MC distributions in `MC Assumptions` or `dcf_engine.py`
- Change tornado steps in `sobol_analysis.py`
- Change chart style in `visualizations.py`
- Change workbook mappings in `excel_io.py`
