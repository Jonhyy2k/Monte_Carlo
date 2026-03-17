# Presentation Script: Monte Carlo DCF Stock Pitch

Use this as a speaker script for a full valuation presentation based on the current workbook-driven Monte Carlo setup.

## Slide 1 — What This Is

**Title:** DCF Valuation With Monte Carlo Driver Analysis

**Talk track:**

"This is a DCF-based valuation with a Monte Carlo overlay. The key distinction is that I am not simulating the stock price directly. I am simulating the valuation drivers inside the DCF itself, including revenue, margins, WACC, and terminal assumptions, and then rebuilding implied value from fundamentals each time."

## Slide 2 — Why This Approach

**Title:** Why Not Just Simulate the Price?

**Talk track:**

"A simple price-level Monte Carlo just adds noise around a target price. That is easy to do, but it does not tell us what matters. This model is more useful because it lets us see whether valuation risk comes from the operating model, the discount rate, or terminal assumptions."

"That is especially important in a stock pitch, because the key question is not just what the target is. The key question is what assumptions the target depends on."

## Slide 3 — How The Model Is Built

**Title:** Excel DCF + Python Simulation Layer

**Talk track:**

"The full DCF remains in Excel. Python acts as the Monte Carlo and sensitivity layer on top of it. The workbook stores the forecast and capital structure assumptions. Python reads those drivers, applies shocks, recomputes FCFF and terminal value, and writes the outputs back into a `Results` dashboard."

**Technical callout:**

- `main.py` orchestrates the full run
- `excel_io.py` reads the workbook and writes `Results`
- `dcf_engine.py` handles the DCF math and Monte Carlo engine
- `sobol_analysis.py` handles Sobol sensitivity and tornado analysis
- `visualizations.py` creates the charts

## Slide 4 — What Gets Simulated

**Title:** Simulated Drivers

**Talk track:**

"The simulation perturbs the main DCF value drivers rather than the final share price. Specifically, it shocks revenue, COGS margin, R&D margin, SG&A margin, D&A margin, tax rate, CapEx, change in working capital, WACC, exit multiple, and terminal growth."

"Those shocks are family-level, not one isolated spreadsheet cell at a time, which keeps the model economically interpretable."

## Slide 5 — DCF Mechanics

**Title:** Cash Flow Construction

**Talk track:**

"For each draw, the model rebuilds free cash flow to the firm. The core flow is revenue to operating profit, then to NOPAT, then to FCFF, then to enterprise value, then to equity value, then to implied price per share."

**Technical line to say:**

```text
FCFF = NOPAT + D&A - CapEx - Change in NWC
```

"That means a margin shock affects EBIT, which affects NOPAT, which affects FCFF, which affects value. A WACC shock affects discount factors. An exit multiple shock affects terminal value. So the model preserves the economic chain."

## Slide 6 — Correlation And Why It Matters

**Title:** Correlated Assumptions

**Talk track:**

"The shocks are not independent. The model includes a correlation structure so the valuation reflects more realistic co-movement. For example, WACC and exit multiple are negatively related, and cost structure variables have positive correlation."

"That matters because independent shocks can create unrealistic combinations that make the output less credible."

## Slide 7 — Base Valuation

**Title:** Base Case

**Talk track:**

"The base case is the deterministic DCF implied by the workbook inputs. This is the anchor valuation before simulation. From there, Monte Carlo shows the range of possible fair values around that base under uncertain assumptions."

**Use current numbers from the latest run:**

- Current price: `$252.65`
- Base DCF: `$236.32`
- Monte Carlo mean: about `$235`
- Probability implied value exceeds current price: about `31.9%`

**What to say:**

"The important point is that the base case and the simulation center are very close. That tells us the base valuation is not an outlier created by one weird assumption."

## Slide 8 — Distribution Of Outcomes

**Title:** Monte Carlo Distribution

**Talk track:**

"The distribution shows the spread of implied prices after 100,000 simulated draws. The 5th percentile is around the mid-170s, the median is around the mid-230s, and the 95th percentile is just under 300. The current price sits above the center of the distribution."

"So the stock is not obviously broken or absurdly overvalued, but the market is pricing in a relatively favorable outcome already."

## Slide 9 — Sobol Sensitivity

**Title:** Which Inputs Drive Variance?

**Talk track:**

"Sobol sensitivity measures which assumptions explain the variance of valuation outcomes. This is a global sensitivity measure, not just a local plus-or-minus case."

"In the current model, exit multiple is the biggest driver, followed by WACC, COGS margin, and revenue. That tells us most of the uncertainty is coming from terminal valuation and core operating efficiency rather than small accounting items."

**Important explanation:**

"Sobol numbers are not directional. A higher number does not mean bullish or bearish. It means that variable explains a larger share of the uncertainty in the valuation distribution."

## Slide 10 — Tornado Analysis

**Title:** Dollar Impact On Price Target

**Talk track:**

"The tornado chart complements Sobol by showing actual dollar movement in the price target when one driver changes and the others are held constant."

"For example, with the current presentation shocks, COGS at plus or minus 2 points moves the target by about 13 dollars, WACC at plus or minus 100 basis points moves it by about 9 dollars, and exit multiple at plus or minus 1 turn also moves it by about 9 dollars."

"This is useful because it translates model sensitivity into pitch language: not just what matters statistically, but what actually moves the target price."

## Slide 11 — What The Model Is Really Saying

**Title:** Core Interpretation

**Talk track:**

"The model says the business can still be strong while the stock is already fairly full. The valuation is most exposed to terminal multiple, discount rate, and margin assumptions. In other words, the main debate is not whether the company is good. The debate is whether the current price already embeds a sufficiently optimistic combination of those factors."

## Slide 12 — Recommendation

**Title:** Investment View

**Talk track:**

"Based on the current setup, I would frame the stock as neutral to modestly overvalued. The base case is below the current price, the Monte Carlo center is below the current price, and only about one-third of simulations imply upside from here."

"So unless I have a differentiated view that the market is too conservative on margins, discount rate, or terminal valuation, I would not pitch this as a high-conviction long at the current price."

## Slide 13 — Technical Appendix

**Title:** Why The Model Is Credible

**Talk track:**

"This is not a black box. The implementation is transparent and split into clean layers. The workbook remains the source of operating assumptions. Python handles only the repeatable valuation math, simulation, and charting."

**Code walkthrough talking points:**

- `main.py` runs the pipeline end to end
- `excel_io.py` maps workbook cells and reads the `DCF Model`, `WACC`, `OP`, and `MC Assumptions` sheets
- `dcf_engine.py` applies correlated shocks and rebuilds price from DCF drivers
- `sobol_analysis.py` computes both Sobol and tornado sensitivity
- `visualizations.py` renders the distribution, snapshot, Sobol, and tornado charts

## Slide 14 — Special Features To Highlight

**Title:** What Makes This Better Than A Basic Model

**Talk track:**

"There are four features I would highlight as differentiators."

1. "The Monte Carlo is applied to DCF drivers, not to the final price."
2. "The shocks are correlated rather than unrealistically independent."
3. "The sensitivity output includes both global variance attribution and local dollar target impact."
4. "The whole process is integrated back into Excel, so the model remains usable in a stock-pitch workflow."

## Short Closing Version

If you need a compressed ending:

"My base DCF is about $236 per share against a current price of about $253. The Monte Carlo mean is roughly the same as the base case, and only about 32% of simulated outcomes are above the current price. The biggest valuation risks are exit multiple, WACC, and margin assumptions. So the stock looks fully valued to modestly overvalued rather than obviously mispriced."
