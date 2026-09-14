# Project 5 · Tata Consultancy Services — Football Field and Valuation Summary

**Files:** `TCS_Football_Field.xlsx` (live model, 237 formulas, zero errors, 18 checks) · `TCS_Football_Field.pdf` · `build_model.py`

**As of:** market data 12 Sep 2026 (share price INR 2,200.80) · financials FY2026, year to 31 March 2026 · India 10-year G-Sec 7.02% on 11 Sep 2026.

## What it is

The page a client sees. Every method of valuing TCS expressed as INR per share and drawn as a floating-bar football field, with the weights behind the recommended range exposed as inputs rather than buried in a judgement. Eight bars: the 52-week trading range, analyst consensus targets, a compact rupee DCF built here, three trading-multiple ranges from Project 3, and two precedent ranges from Project 4. The DCF exists because there was no TCS DCF anywhere in the repository — Project 2 valued Apple — and a football field without an intrinsic-value bar is only half an answer.

## Result

| Methodology | Low | High | Midpoint | Weight |
|---|---|---|---|---|
| 52-week trading range | 1,977 | 3,350 | 2,663 | 0% |
| Analyst price targets (43 analysts) | 1,775 | 3,900 | 2,838 | 10% |
| **DCF — WACC and growth sensitivity** | **1,625** | **2,504** | **2,064** | **35%** |
| Trading comps — EV/CY2027 revenue | 1,242 | 1,745 | 1,494 | 15% |
| Trading comps — EV/LTM EBITDA | 1,652 | 2,478 | 2,065 | 15% |
| Trading comps — P/E (CY2027) | 1,874 | 2,689 | 2,282 | 15% |
| Precedents — EV/LTM revenue *(illustrative)* | 1,428 | 2,682 | 2,055 | 5% |
| Precedents — EV/LTM EBITDA *(illustrative)* | 2,862 | 3,532 | 3,197 | 5% |
| **Recommended range (weighted)** | **1,676** | **2,614** | **2,145** | |

Share price 2,200.80 — 2.6% above the recommended midpoint.

**DCF base case:** enterprise value INR 6,777bn, equity value INR 7,065bn, **INR 1,953 per share**. WACC 13.3% in INR (7.02% risk-free + 0.94 Blume-adjusted beta × 6.7% India equity risk premium; no debt tranche, because TCS is net cash). Terminal value is 69.8% of enterprise value. The 11.0× exit multiple implies 6.8% perpetual growth against the 5.5% assumed, and 5.5% implies a 9.1× exit multiple against the 11.0× assumed — close enough that neither assumption contradicts the other.

## Reading it honestly

* **Project 2's beta method was run for TCS and had to be thrown away.** Apple's beta came from regressing peers against the S&P 500, unlevering, taking the median and relevering. Applied to TCS that gives a beta of 0.14 with an R-squared of 0.01; Infosys 0.24 at 0.02; HCLTech 0.06 at 0.00; Wipro 0.33 at 0.04. Indian shares and the S&P 500 trade in different hours and different currencies, so the regression measures nothing. A beta with no explanatory power is not a conservative estimate, it is a wrong one. The regression was re-run where TCS actually trades — against the NIFTY 50, in INR, 254 weekly observations — giving 0.91 with an R-squared of 0.28. The WACC sheet keeps the rejected regressions on the page instead of deleting them, because the rejection is the finding. It is also why the whole DCF is built in INR: local cash flows at a local cost of capital, with no FX forecast buried inside a valuation.
* **The weights are judgement and they are visible.** Trading comps carry 45% across three multiples because they need no forecast. The DCF carries 35% because it is the only method that values TCS rather than its peer group. Precedents carry 10% because of the size and cycle mismatch. Analyst targets carry 10% because they recycle public information and lag the price. The 52-week range carries nothing — it is context, not a valuation. Change a weight on the Ranges sheet and the recommended range moves; that argument is better had in the open than settled silently inside a number.
* **The precedent bars sit above everything else.** Project 4 already flagged that its 2.6× revenue median came from deals struck in 2014–2022, when IT services traded at 20–30× earnings, against a sector since de-rated to 10–15×. Putting those bars on the same axis as September 2026 trading multiples makes the control premium look larger than it is. They are marked illustrative and weighted at 5% each.
* **Seventy per cent of the DCF is terminal value.** Normal for a profitable, asset-light business, and also the honest limit of the exercise: most of the number is an assumption about 2031, not a forecast of the next five years. That is why the DCF bar is the sensitivity range rather than the base-case point.
* **The analyst bar is the widest on the page and that is the point.** Forty-three analysts span 1,775 to 3,900 on the same public information. A consensus that disperses by more than 2× is telling you the sector's AI-disruption debate is unresolved, not that the average of 2,944 is meaningful.
* **The chart cannot draw the current price as a line.** openpyxl will not overlay a scatter series on a bar chart without add-ins, so the price is stated in bold beside the chart and in the conclusion block. In PowerPoint (Project 16) it is drawn properly.

## What I would do with more time

* Pair each precedent with the trading multiples on its own announcement date, separating control premium from cycle.
* Add an LBO floor from Project 7 as a ninth bar — what a sponsor could pay is the practical lower bound in any sale process.
* Build the DCF off a full three-statement TCS model in the Project 1 pattern rather than the compact forecast used here.
* Extend the DCF to the three-case structure used in Project 1, so the bar reflects operating scenarios and not only discount-rate uncertainty.

## Sources

TCS annual financials FY2023–FY2026 (year to 31 March) from the Yahoo Finance fundamentals series for TCS.NS, restating the consolidated annual report. Share price, 52-week range and share count from Yahoo Finance, 12 September 2026 — the same market date as Projects 3 and 4. India 10-year G-Sec at 7.02%, 11 September 2026. Equity risk premium: 4.5% mature-market premium, as used in Project 2, plus a 2.2% India country risk premium (Damodaran). Beta from five years of weekly returns to 12 September 2026. Analyst consensus: 43 published 12-month targets, 12 September 2026. Trading-multiple ranges from Project 3; precedent ranges from Project 4.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
