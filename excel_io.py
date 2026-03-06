"""
Phase 6 — Excel I/O Stub

Creates a dummy Excel workbook with an Inputs sheet.
Provides read_inputs_from_excel() and write_results_to_excel().
"""

import numpy as np
import openpyxl
from pathlib import Path


# ---------------------------------------------------------------------------
# Cell map: parameter name → (sheet, cell)
# ---------------------------------------------------------------------------
INPUT_CELL_MAP = {
    "revenue_base":         ("Inputs", "B2"),
    "revenue_growth":       ("Inputs", "B3"),
    "projection_years":     ("Inputs", "B4"),
    "cogs_pct":             ("Inputs", "B5"),
    "sga_pct":              ("Inputs", "B6"),
    "capex_intensity":      ("Inputs", "B7"),
    "tax_rate":             ("Inputs", "B8"),
    "exit_multiple":        ("Inputs", "B9"),
    "terminal_g":           ("Inputs", "B10"),
    "wacc":                 ("Inputs", "B11"),
    "net_debt":             ("Inputs", "B12"),
    "shares_outstanding":   ("Inputs", "B13"),
    "exit_multiple_weight": ("Inputs", "B14"),
    "current_price":        ("Inputs", "B15"),
    "fundamental_target":   ("Inputs", "B16"),
}

# Default placeholder values (generic company)
DEFAULTS = {
    "revenue_base":         10_000.0,
    "revenue_growth":       0.08,
    "projection_years":     5,
    "cogs_pct":             0.55,
    "sga_pct":              0.15,
    "capex_intensity":      0.06,
    "tax_rate":             0.25,
    "exit_multiple":        12.0,
    "terminal_g":           0.025,
    "wacc":                 0.10,
    "net_debt":             2_000.0,
    "shares_outstanding":   500.0,
    "exit_multiple_weight": 0.50,
    "current_price":        55.0,
    "fundamental_target":   68.0,
}


DCF_LINE_ITEMS = [
    "Revenue", "COGS", "Gross Profit", "SG&A", "EBITDA",
    "D&A", "EBIT", "Taxes", "NOPAT", "CapEx", "Chg in NWC", "FCFF",
]

DCF_SCALAR_ASSUMPTIONS = {
    "WACC":                 0.10,
    "Exit Multiple":        12.0,
    "Terminal g":           0.025,
    "Net Debt":             2_000.0,
    "Shares Outstanding":   500.0,
    "Current Price":        55.0,
    "Target Price":         68.0,
    "Exit Multiple Weight": 0.50,
}


def _build_dcf_dummy_data(n_years: int = 5) -> dict[str, list[float]]:
    """Build realistic year-by-year P&L→FCFF from defaults."""
    rev_base = DEFAULTS["revenue_base"]
    g = DEFAULTS["revenue_growth"]
    cogs_pct = DEFAULTS["cogs_pct"]
    sga_pct = DEFAULTS["sga_pct"]
    capex_pct = DEFAULTS["capex_intensity"]
    tax_rate = DEFAULTS["tax_rate"]

    data = {item: [] for item in DCF_LINE_ITEMS}
    for t in range(n_years):
        rev = rev_base * (1 + g) ** (t + 1)
        cogs = -rev * cogs_pct
        gross = rev + cogs
        sga = -rev * sga_pct
        ebitda = gross + sga
        da = -rev * capex_pct
        ebit = ebitda + da
        taxes = -(abs(ebit) * tax_rate)
        nopat = ebit + taxes
        capex = -rev * capex_pct
        nwc = -rev * 0.01  # small NWC change
        fcff = nopat - da + capex + nwc  # NOPAT + D&A - CapEx - Chg NWC

        data["Revenue"].append(round(rev, 1))
        data["COGS"].append(round(cogs, 1))
        data["Gross Profit"].append(round(gross, 1))
        data["SG&A"].append(round(sga, 1))
        data["EBITDA"].append(round(ebitda, 1))
        data["D&A"].append(round(da, 1))
        data["EBIT"].append(round(ebit, 1))
        data["Taxes"].append(round(taxes, 1))
        data["NOPAT"].append(round(nopat, 1))
        data["CapEx"].append(round(capex, 1))
        data["Chg in NWC"].append(round(nwc, 1))
        data["FCFF"].append(round(fcff, 1))
    return data


def create_stub_excel(filepath: str = "model_inputs.xlsx") -> None:
    """Create a dummy Excel file with Inputs and DCF sheets."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inputs"

    # Header
    ws["A1"] = "Parameter"
    ws["B1"] = "Value"
    ws["A1"].font = openpyxl.styles.Font(bold=True)
    ws["B1"].font = openpyxl.styles.Font(bold=True)

    for param, (sheet, cell) in INPUT_CELL_MAP.items():
        row = int(cell[1:])
        ws[f"A{row}"] = param
        ws[cell] = DEFAULTS[param]

    # Adjust column widths
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 15

    # ---- DCF sheet ----
    n_years = int(DEFAULTS["projection_years"])
    dcf_data = _build_dcf_dummy_data(n_years)
    bold = openpyxl.styles.Font(bold=True)

    dcf = wb.create_sheet("DCF")
    # Row 1: headers
    dcf["A1"] = "Line Item"
    dcf["A1"].font = bold
    for c in range(n_years):
        cell = dcf.cell(row=1, column=c + 2, value=f"Year {c + 1}")
        cell.font = bold

    # Rows 2–13: line items
    for r, item in enumerate(DCF_LINE_ITEMS, start=2):
        dcf.cell(row=r, column=1, value=item)
        for c, val in enumerate(dcf_data[item]):
            dcf.cell(row=r, column=c + 2, value=val)

    # Row 15+: scalar assumptions
    scalar_start = len(DCF_LINE_ITEMS) + 3  # row 15
    dcf.cell(row=scalar_start, column=1, value="Assumptions").font = bold
    for i, (key, val) in enumerate(DCF_SCALAR_ASSUMPTIONS.items()):
        r = scalar_start + 1 + i
        dcf.cell(row=r, column=1, value=key)
        dcf.cell(row=r, column=2, value=val)

    dcf.column_dimensions["A"].width = 22
    for c in range(n_years):
        dcf.column_dimensions[openpyxl.utils.get_column_letter(c + 2)].width = 14

    wb.save(filepath)
    print(f"  Created stub: {filepath}")


def read_dcf_from_excel(filepath: str = "model_inputs.xlsx") -> dict:
    """
    Read the DCF sheet → dict with numpy arrays keyed by line item + scalars.

    Returns
    -------
    dict with keys:
        "arrays"  → {line_item_name: np.array of yearly values}
        "n_years" → int
        "wacc", "exit_multiple", "terminal_g", "net_debt",
        "shares_outstanding", "current_price", "target_price",
        "exit_multiple_weight" → floats
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    dcf = wb["DCF"]

    # Detect number of year columns (row 1, starting col B)
    n_years = 0
    for c in range(2, 50):
        if dcf.cell(row=1, column=c).value is None:
            break
        n_years += 1

    # Read line items (rows 2+)
    arrays = {}
    row = 2
    while True:
        label = dcf.cell(row=row, column=1).value
        if label is None or label == "Assumptions":
            break
        if label in DCF_LINE_ITEMS:
            vals = []
            for c in range(2, 2 + n_years):
                v = dcf.cell(row=row, column=c).value
                vals.append(float(v) if v is not None else 0.0)
            arrays[label] = np.array(vals)
        row += 1

    # Read scalar assumptions below the line items
    scalar_map = {
        "WACC": "wacc",
        "Exit Multiple": "exit_multiple",
        "Terminal g": "terminal_g",
        "Net Debt": "net_debt",
        "Shares Outstanding": "shares_outstanding",
        "Current Price": "current_price",
        "Target Price": "target_price",
        "Exit Multiple Weight": "exit_multiple_weight",
    }
    scalars = {}
    for r in range(row, dcf.max_row + 1):
        label = dcf.cell(row=r, column=1).value
        if label in scalar_map:
            val = dcf.cell(row=r, column=2).value
            scalars[scalar_map[label]] = float(val) if val is not None else DCF_SCALAR_ASSUMPTIONS[label]

    # Fill defaults for missing scalars
    for excel_name, key in scalar_map.items():
        if key not in scalars:
            scalars[key] = DCF_SCALAR_ASSUMPTIONS[excel_name]

    wb.close()
    return {"arrays": arrays, "n_years": n_years, **scalars}


def has_dcf_sheet(filepath: str = "model_inputs.xlsx") -> bool:
    """Check if the workbook contains a DCF sheet."""
    wb = openpyxl.load_workbook(filepath, read_only=True)
    result = "DCF" in wb.sheetnames
    wb.close()
    return result


def read_inputs_from_excel(filepath: str = "model_inputs.xlsx") -> dict:
    """
    Read model inputs from Excel and return a dictionary.

    When the real model arrives, just update INPUT_CELL_MAP references.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    inputs = {}
    for param, (sheet, cell) in INPUT_CELL_MAP.items():
        ws = wb[sheet]
        val = ws[cell].value
        if val is None:
            val = DEFAULTS.get(param)
        inputs[param] = val
    wb.close()
    return inputs


def write_results_to_excel(
    filepath: str = "model_inputs.xlsx",
    results: dict | None = None,
) -> None:
    """
    Write Monte Carlo output stats into a Results sheet.

    Expected keys in results dict:
        mean, median, std, p5, p25, p50, p75, p95,
        n_iterations, n_valid, elapsed_seconds,
        sobol_df (optional pandas DataFrame)
    """
    path = Path(filepath)
    if path.exists():
        wb = openpyxl.load_workbook(filepath)
    else:
        wb = openpyxl.Workbook()

    # Remove old Results sheet if present
    if "Results" in wb.sheetnames:
        del wb["Results"]
    ws = wb.create_sheet("Results")

    # Header style
    bold = openpyxl.styles.Font(bold=True)
    money_fmt = '#,##0.00'

    # ---- Summary statistics ----
    ws["A1"] = "Monte Carlo Results"
    ws["A1"].font = openpyxl.styles.Font(bold=True, size=14)

    stats_rows = [
        ("Metric",              "Value"),
        ("Mean Price ($)",      results.get("mean")),
        ("Median Price ($)",    results.get("median")),
        ("Std Dev ($)",         results.get("std")),
        ("5th Percentile ($)",  results.get("p5")),
        ("25th Percentile ($)", results.get("p25")),
        ("50th Percentile ($)", results.get("p50")),
        ("75th Percentile ($)", results.get("p75")),
        ("95th Percentile ($)", results.get("p95")),
        ("Iterations",          results.get("n_iterations")),
        ("Valid Draws",         results.get("n_valid")),
        ("Runtime (s)",         results.get("elapsed_seconds")),
    ]
    for i, (label, val) in enumerate(stats_rows, start=3):
        ws[f"A{i}"] = label
        ws[f"B{i}"] = val
        if i == 3:
            ws[f"A{i}"].font = bold
            ws[f"B{i}"].font = bold
        elif val is not None and isinstance(val, float) and "Price" in label:
            ws[f"B{i}"].number_format = money_fmt

    # ---- Sobol sensitivity (if provided) ----
    sobol_df = results.get("sobol_df")
    if sobol_df is not None:
        start_row = 17
        ws[f"A{start_row}"] = "Sobol Sensitivity"
        ws[f"A{start_row}"].font = openpyxl.styles.Font(bold=True, size=12)
        start_row += 1
        for j, col in enumerate(sobol_df.columns):
            ws.cell(row=start_row, column=j + 1, value=col).font = bold
        for i, row_data in enumerate(sobol_df.itertuples(index=False), start=start_row + 1):
            for j, val in enumerate(row_data):
                ws.cell(row=i, column=j + 1, value=val)

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 15

    wb.save(filepath)
    print(f"  Results written to: {filepath} [Results sheet]")


# ---------------------------------------------------------------------------
# Main — create stub, round-trip test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=== Phase 6: Excel I/O ===\n")

    create_stub_excel("model_inputs.xlsx")
    inputs = read_inputs_from_excel("model_inputs.xlsx")
    print(f"  Read {len(inputs)} parameters:")
    for k, v in inputs.items():
        print(f"    {k:>25s} = {v}")

    # Test DCF sheet round-trip
    assert has_dcf_sheet("model_inputs.xlsx"), "DCF sheet should exist"
    dcf_data = read_dcf_from_excel("model_inputs.xlsx")
    print(f"\n  DCF sheet: {dcf_data['n_years']} years, {len(dcf_data['arrays'])} line items")
    for item, arr in dcf_data["arrays"].items():
        print(f"    {item:>15s}: {arr}")
    print(f"    WACC = {dcf_data['wacc']}, Exit Multiple = {dcf_data['exit_multiple']}")

    # Write dummy results
    dummy_results = {
        "mean": 81.0,
        "median": 64.72,
        "std": 57.40,
        "p5": 19.49,
        "p25": 39.58,
        "p50": 64.72,
        "p75": 105.59,
        "p95": 201.70,
        "n_iterations": 50000,
        "n_valid": 50000,
        "elapsed_seconds": 0.023,
    }
    write_results_to_excel("model_inputs.xlsx", dummy_results)
    print("\n  Round-trip test passed.")
