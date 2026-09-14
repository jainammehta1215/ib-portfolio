# Project 1 · Apple Inc. — Three-Statement Operating Model

**Files:** `Apple_3S_Model.xlsx` (live model, 1,027 formulas, zero errors) · `Apple_3S_Model.pdf` (print) · `build_model.py` (the script that spreads the filings and writes every formula)

**As of:** FY2025 10-K (fiscal year ended 27 Sep 2025). Historicals FY2022–FY2025; forecast FY2026E–FY2030E. US$ billions.

## What it is

An integrated income statement, balance sheet and cash-flow forecast driven from a single assumptions sheet with Base / Upside / Downside cases selected in one cell (`Inputs!C4`). Supporting schedules for PP&E, long-term debt, a revolver that plugs cash shortfalls against a minimum-cash floor, and share count. A Checks sheet confirms the balance sheet balances, closing cash ties, schedules tie and historical net income matches the reported figure, and returns `MODEL OK`.

## How it is built

* Historicals are spread from Apple's Form 10-K figures and reconciled: every subtotal ties to the reported statements. The "other" asset and liability lines are defined as residuals to Apple's reported subtotals so the tie holds in every year regardless of how a data vendor maps leases and non-current payables (this caught a 29bn discrepancy in FY2022 on the first pass).
* Revenue grows at a scenario rate; gross margin, R&D, SG&A, tax, D&A, capex and SBC are percentages of revenue; receivables, inventory and payables are on days; other balance-sheet lines are percentages of revenue.
* Buybacks are a percentage of free cash flow after dividends, which is how Apple manages to a "net cash neutral" position. In the Base case net cash stays flat; in the Upside (110%) cash is spent down to the $25bn floor and the revolver draws from FY2028; in the Downside (80%) cash accumulates.
* Interest is calculated on opening balances to avoid a circular reference. Switching to average balances is a one-line change once iterative calculation is enabled in Excel.
* Share count: buybacks retire shares at an assumed price path; SBC issues shares at 60% of the gross grant (approximating net-of-withholding issuance). EPS uses the average diluted count.

## Key outputs (Base case)

| | FY2025A | FY2026E | FY2028E | FY2030E |
|---|---|---|---|---|
| Revenue | 416.2 | 441.1 | 486.3 | 526.0 |
| EBIT margin | 32.0% | 32.1% | 32.5% | 32.6% |
| Net income | 112.0 | 119.4 | 135.1 | 149.4 |
| Diluted EPS | $7.46 | $8.17 | $9.64 | $11.08 |
| Free cash flow | 98.8 | 131.6 | 148.9 | 164.5 |
| Net cash incl. LT securities | 33.8 | 33.8 | 33.8 | 33.8 |

FY2026E free cash flow steps up because FY2025's operating cash flow was depressed by a $25bn working-capital outflow (mostly tax payments related to the EU State Aid decision); the model's working-capital lines revert to their historical days.

## What I would do differently with more time

* Build revenue bottom-up by product (iPhone, Mac, iPad, Wearables, Services) from the 10-K segment note; Services carries a ~75% gross margin versus ~37% for products, so mix drives the margin path more than any single assumption here.
* Model the debt maturity ladder from the 10-K note rather than a flat repayment assumption.
* Add a quarterly version for the next four quarters to compare against consensus.

## Sources

Apple Inc. Form 10-K for fiscal years 2022–2025, SEC EDGAR CIK 0000320193 (https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0000320193&type=10-K). Figures pulled via Yahoo Finance and reconciled to the reported subtotals. Share price for the buyback share-count assumption: approximate late-September 2025 close.
