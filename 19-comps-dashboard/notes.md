# Project 19 · Deal-Sourcing and Comps-Refresh Dashboard

**Files:** `dashboard.py` (the tool) · `universe.csv` · `snapshot.json` (vendor fields and FX, 18 Sep 2026) · `Comps_Dashboard.xlsx` (100 formulas, zero errors) · `Comps_Dashboard.pdf` · `dashboard.png`

**Run:** `python dashboard.py` rebuilds everything offline from the snapshot; `python dashboard.py --refresh` pulls live data first.

## What it is

The other projects are models of one company. This one is the tool a coverage desk runs every Monday: 52 names across the DFM, ADX, Tadawul, NSE and US exchanges, pulled from a vendor, normalised to US dollars, with every multiple recomputed from its components and four screens on top — take-private candidates, value against quality, banks on price-to-book against return on equity, and drawdowns. The output is an Excel dashboard in the same design system as every other project, plus a four-chart panel. It is a Python script rather than a workbook because the value is in the refresh, and a refresh has to be reproducible: the snapshot is committed so the build is deterministic, and `--refresh` replaces it.

## Result, 18 September 2026

| | |
|---|---|
| Universe · usable records | 52 · 48 (four ADX tickers return no vendor record) |
| Records flagged for a data problem | **10** — roughly one name in five |
| Currency mismatches (financials in one currency, price in another) | Infosys, HCL Technologies: vendor EV/EBITDA reads **956× and 1,218×**; recomputed **~10×** |
| Vendor multiples disagreeing with components by more than 25% | Emaar Properties and others where cash and debt are stale |
| **Take-private screen** (below sector median, FCF yield > 6%, net debt < 1.5× EBITDA) | **Cognizant, EPAM Systems, Wipro** |
| Situational screen (> 25% below the 52-week high), largest first | EPAM 47%, Trent 46%, Wipro 40%, Infosys 40%, TCS 37%, Emaar Development 35%, Accenture 35% |
| Banks below the justified P/B line (10% cost of equity, 3% growth) | Dubai Islamic Bank, Emirates NBD, Mashreqbank, HDFC Bank |
| Sector medians (US$, recomputed) | Energy 8.8× · IT services 8.9× · Telecom 9.3× · Utilities 10.9× · Infrastructure 18.4× · Technology 24.4× · Retail 27.9× · Consumer 31.8× |

## Reading it honestly

* **The work is in the normalisation, not the screens.** A vendor record for a single company can carry a market cap in rupees and an income statement in dollars, a missing market cap, or an EV that does not reconcile to its own cash and debt. Trusting the vendor's EV/EBITDA would have put two of India's largest companies at a thousand times earnings. The script converts each field in its own currency, rebuilds EV from components, and names the disagreement in a Flags column. The Data Quality sheet exists so that the reader sees what was fixed rather than having to trust that it was.
* **The take-private screen finds the same kind of name it found in Project 7.** Cognizant, EPAM and Wipro are IT-services companies trading below their sector median with high free-cash-flow yields, net cash, and shares 30–47% off their highs. That is a description of what sponsors look for, and it is the reason Cognizant was the LBO case. A screen is a list of places to look; the next step for each is the shareholder register and the reason for the discount, which is what the sector is going through rather than anything company-specific.
* **A metric that works for one sector fails for another, and the script has to know it.** Emaar Development shows an EV/EBITDA below 1× because its customers' advance payments sit in cash and cancel most of its market cap. That is not a bargain; it is the wrong metric. Developers are excluded from the take-private screen and flagged to Project 13's NAV approach; banks are screened on P/B and ROE for the same reason.
* **The bank line is deliberately crude.** A single 10% cost of equity is applied to four markets so the chart is comparable; Project 11 builds Emirates NBD's properly at 10.0% and the Gulf banks sit below the line for the reason found there — the market discounts them at a rate the country-risk build-up does not capture. A production screen would use one cost of equity per market.
* **What is missing from a real desk's version.** Forward estimates (all multiples are trailing), a second vendor to cross-check the first, a shareholder-register pull for the take-private names, sector-specific screens beyond banks and developers, and a scheduler. Each is a day's work on top of this script; none changes its shape.

## What I would do with more time

* Add consensus forward EV/EBITDA and EPS so the value-versus-quality chart is on next year's numbers.
* Pull ownership data for the take-private hits: founder stake, sponsor presence, free float, activist filings.
* Store each week's snapshot and add a "what moved" sheet: names that entered or left a screen, sector medians that shifted.
* Replace the single bank cost of equity with a per-market build-up, reusing Project 11's method.

## Sources

Vendor fields and FX from Yahoo Finance via `yfinance`, snapshot 18 September 2026, committed as `snapshot.json`. Universe, screen thresholds, the vendor-disagreement tolerance and the bank cost of equity are stated in `dashboard.py`. A screen is a list of places to look; nothing here is a recommendation or a suggestion that any company is in a process.
