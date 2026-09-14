# Project 2 · Apple Inc. — DCF Valuation

**Files:** `Apple_DCF.xlsx` (live model) · `Apple_DCF.pdf` · `build_model.py` (pulls the three scenarios out of Project 1, the peer data and market rates, and writes every formula)

**As of:** market data 12 Sep 2026 (share price US$332.27); balance sheet 28 Jun 2026 10-Q; forecasts from Project 1's FY2025-based model.

## What it is

An unlevered free-cash-flow DCF on the Project 1 operating model, with:

* **WACC built bottom-up.** Ten peers (MSFT, GOOGL, AMZN, META, NVDA, SONY, Samsung, DELL, HPQ and Apple itself), five years of weekly returns against the S&P 500, unlevered at each peer's net debt/equity, median relevered at Apple's (zero: Apple is net cash), Blume-adjusted (0.67β + 0.33) as banks do. Result: beta 1.12, cost of equity 10.0% on a 4.95% ten-year Treasury and a 4.5% equity risk premium, WACC 9.9% (equity is 98% of Apple's capital, so WACC ≈ cost of equity).
* **Terminal value both ways** with each method's implied assumption cross-checked against the other, and an implied P/E on LTM net income against the market's.
* **A stub period** for the twelve days left in FY2026 at the valuation date, mid-year convention thereafter.
* **Sensitivity tables** that re-solve the DCF in closed form (no Excel data tables), so they work in Google Sheets and LibreOffice too.
* The three operating cases embedded, so the workbook stands alone; toggle in `Inputs!C5`.

## Result (Base case)

| | Perpetuity growth (3.0%) | Exit multiple (22× EBITDA) |
|---|---|---|
| Enterprise value | US$1,889bn | US$3,240bn |
| Equity value | US$1,951bn | US$3,303bn |
| **Implied price per share** | **US$133** | **US$224** |
| vs market price US$332 | −60% | −32% |
| Terminal value as % of EV | 76% | 86% |
| Implied by the other method | exit multiple **11.4×** | perpetuity growth **6.2%** |
| Implied P/E on LTM earnings | 15× | 26× |
| Market P/E on LTM earnings | 38× | 38× |

## What that means, honestly

The DCF says Apple's share price embeds more than a base-case forecast at a textbook discount rate supports. That is the normal outcome for a DCF of Apple and it is worth understanding rather than "fixing":

1. **The market's implied inputs.** To reach US$332 with a 3% terminal growth rate needs a WACC near 7%; with a 9.9% WACC it needs terminal growth above 6%, higher than nominal GDP. Either the market uses a lower equity risk premium than 4.5% for the safest large-cap balance sheet in the world, or it expects Services (75% gross margin, still growing double digits) to keep lifting the margin structure for far longer than five years, or both.
2. **Terminal value dominates.** At 76–86% of enterprise value, a 1% change in WACC or 0.5% in g moves the answer 10–15%; the sensitivity tables, not the point estimate, are the output. This is a property of every mature-company DCF.
3. **SBC is a real cost.** Treating Apple's ~US$13–16bn a year of stock compensation as an expense (not adding it back) lowers UFCF and value by roughly 8% versus the naïve build. Adding it back would flatter the number and overstate free cash flow to shareholders.
4. **The cross-check does its job.** The perpetuity method implies an 11× exit multiple against peers trading at 18–30×; the exit-multiple method implies 6.2% perpetual growth. The truth is between them, which is why Project 5 draws the DCF as a bar.

In a pitch, this DCF would be presented alongside comparables (Project 3) and precedents (Project 4), and the conversation with the client would be about *which* assumptions the market is making, not about the point estimate.

## What I would do with more time

* A segment-level operating model (Services vs Products) so the terminal margin is an output of mix, not an input.
* A reverse DCF: solve for the growth path the current price implies and judge that instead.
* A 10-year explicit period with fading growth to reduce the terminal-value share.
* Apple's own bond yields (it has 30-year paper) instead of the AAA index for the cost of debt; immaterial here given the weights.

## Sources

Apple Form 10-K FY2025 and Form 10-Q for the quarter ended 28 Jun 2026 (SEC EDGAR CIK 0000320193) · Yahoo Finance: prices, peer balance sheets, 12 Sep 2026 · FRED: DGS10, BAMLC0A1CAAAEY (10 Sep 2026) · Damodaran, implied US equity risk premium (mid-2026) · Blume (1975) on beta adjustment.
