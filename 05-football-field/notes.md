# Project 5 · TCS — Football Field and Valuation Summary

**Files:** `TCS_Football_Field.xlsx` · `TCS_Football_Field.pdf` · `build_model.py`

## What it is

The page a client sees: every valuation method for TCS expressed as INR per share, drawn as a floating-bar football field, with a methodology page saying what moves each bar and a recommended range whose weights are explicit inputs. It combines Project 3 (trading comparables), Project 4 (precedents), a compact rupee DCF built here, the 52-week trading range and analyst consensus.

## The DCF in brief

Five-year unlevered free cash flow in INR crore from FY2026 actuals (revenue INR 2.67 lakh crore, EBIT margin 25%), revenue growth fading from 5% to 4%, 24.5% tax, capex 1.6% of revenue, working capital at 15% of incremental revenue. WACC from Indian inputs: 6.9% ten-year G-sec, 6.5% equity risk premium (mature-market ERP plus India country premium), beta 0.90, negligible debt → **12.6%**. Terminal value by 4% perpetuity growth (implies 8.4× exit EBITDA) and by a 14× exit multiple (implies 7.3% perpetual growth). Value per share INR 1,874 (perpetuity) to INR 2,671 (exit multiple); the football-field bar spans WACC ± 1% across both methods: **INR 1,698 – 2,763**.

## The field

| Method | Low | Mid | High | Weight |
|---|---|---|---|---|
| 52-week trading range | 1,972 | 2,588 | 3,204 | 0% |
| Analyst price targets (41) | 1,800 | 2,400 | 3,480 | 10% |
| Trading comparables (Project 3) | 1,652 | 1,848 | 2,478 | 35% |
| Precedent transactions (Project 4, illustrative) | 1,428 | 2,026 | 2,682 | 15% |
| DCF (WACC ± 1%, both terminal methods) | 1,698 | 2,272 | 2,763 | 40% |
| **Recommended range (weighted)** | **1,652** | **2,100** | **2,723** | |
| Share price, 12 Sep 2026 | | 2,201 | | |

## Reading it

* **The bars overlap the share price.** Comps sit slightly below (the sector de-rated in 2026), the DCF slightly above (TCS's cash generation supports the price at a 12–13% Indian discount rate), and the market is in the middle. The case for calling TCS cheap or expensive rests on disagreeing with a specific assumption: the equity risk premium, the terminal growth rate, or whether the sector's de-rating is permanent.
* **Terminal value is ~70% of the DCF.** The perpetuity and exit methods disagree by 40% because each embeds an assumption the other rejects (an 8× exit multiple is well below where TCS has ever traded; 7% perpetual growth is above India's nominal GDP). The honest bar is wide.
* **Weights are judgement, made visible.** The 52-week range carries no weight (it is context), analyst targets little (they reuse the same information), precedents little (size mismatch, older cycle). Change the weights on the Ranges sheet and the recommended range moves; that is the conversation with the client.
* **What the chart cannot show:** the current price as a vertical line. openpyxl's charts do not support a scatter overlay on a bar chart without add-ins, so the price is stated in the table beside the chart; in PowerPoint (Project 16) it is drawn as a line.

## What I would do with more time

* A full three-statement TCS model (the Project 1 pattern) behind the DCF rather than the compact build here.
* Same-date trading multiples for the precedents to remove the cycle effect from the control premium.
* An LBO floor (Project 7) as a sixth bar: what a sponsor could pay is the practical lower bound in a sale process.

## Sources

TCS FY2026 annual report; NSE prices; Yahoo Finance 52-week range and analyst targets (12 Sep 2026); FRED INDIRLTLT01STM; Damodaran country risk premia; Projects 3 and 4.
