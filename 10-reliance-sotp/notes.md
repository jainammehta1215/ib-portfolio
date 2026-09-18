# Project 10 · Reliance Industries — Sum-of-the-Parts Valuation

**Files:** `Reliance_SOTP.xlsx` (live model, 213 formulas, zero errors) · `Reliance_SOTP.pdf` · `build_model.py`

**As of:** audited FY2026 segment information (year to 31 March 2026) · peer multiples and share price 18 Sep 2026 · ₹ crore.

## What it is

Reliance is five businesses behind one share price, and no single multiple fits them all. Each reported segment — Digital Services (Jio Platforms), Retail, Oil to Chemicals, Oil and Gas, Others — is valued on its own peer group at its own EV/EBITDA multiple; RIL's economic stake in Jio (67.03%) and Retail (85.06%) is applied; the adjusted net debt from Project 8 is deducted; a holding-company discount is taken; and the result is expressed per share against the market. The **Peers** sheet shows every comparable, every exclusion with its reason, and the selected multiple against the median so the judgement is visible. The **Market-Implied** sheet then runs the bridge backwards: holding four segments at their base multiples, what multiple is today's share price paying for Jio? With the Jio Platforms listing approved by its board, that is the number the market is about to test.

## Result

| | |
|---|---|
| Jio Platforms — RIL share at 12.0× | ₹615,800 crore (33% of gross value) |
| Retail — RIL share at 25.0× | ₹574,900 crore (31%) |
| Oil to Chemicals at 7.5× | ₹454,100 crore (25%) |
| Oil and Gas at 5.0× | ₹95,300 crore (5%) |
| Others at 10.0× | ₹108,600 crore (6%) |
| Gross asset value | ₹1,848,600 crore |
| Less adjusted net debt · less 10% holding discount | (₹247,800 crore) · (₹160,100 crore) |
| **SOTP value per share** | **₹1,065** (range ₹853–1,277 on low–high multiples) |
| Share price | ₹1,245 — **14.5% above the SOTP** |
| **Market-implied Jio EV / EBITDA** | **17.3×** — a 49% premium to Bharti Airtel, US$158bn EV |
| Market-implied Retail EV / EBITDA (Jio at base) | 36.8× — US$119bn EV |

## Reading it honestly

* **The parts are worth less than the whole, and that is the finding.** At multiples drawn from listed peers the SOTP lands about 15% below the market. A sum-of-the-parts that produces a discount to the price is usually taken as a sell signal; here it is better read as a measurement of what the market is already assuming. Run backwards, the price is consistent with Jio at 17× EBITDA — half again Bharti Airtel's multiple — or, equivalently, Retail at nearly 37×. The market has priced a Jio listing at a substantial premium to the only listed comparable.
* **Two businesses carry two-thirds of the value while generating under half the EBITDA.** Jio and Retail together are 64% of gross asset value on 53% of segment EBITDA; O2C and E&P produce 41% of EBITDA and are worth 30%. That gap between where the cash comes from and where the value sits is the whole story of the group's last decade, and it is why the SOTP method suits Reliance better than any consolidated multiple.
* **The Jio multiple is the whole argument.** Each turn of EBITDA on Jio is worth ₹34 per share after the stake and the discount. The difference between the 12× selected here and the 17× the market implies is ₹170 a share, which is close to the entire gap to the price. Retail is the second lever: each turn is worth ₹15 per share.
* **The holding-company discount is a judgement, stated as one.** Ten percent is at the low end of where Indian conglomerates trade relative to their parts, chosen because Jio and Retail are consolidated subsidiaries rather than associates and a listing narrows the discount further. Each five points is ₹90 per share; the second sensitivity grid shows the full range from zero to 20%.
* **The energy multiples are on trailing EBITDA, not through the cycle.** FY2026 O2C EBITDA was helped by unusually strong middle-distillate cracks. A mid-cycle number would be lower, taking ₹30–40 per share off the total and widening the discount to the price. The E&P multiple is anchored on ONGC and Oil India rather than the global majors, because KG-D6 gas is priced under an Indian formula and production is declining.
* **Net debt is deducted in full at group level.** About a third of the spectrum liabilities belong economically to Jio's minority shareholders, and a purist would deduct Jio's own net debt inside the segment before applying the 67% stake. The results release does not disclose Jio's balance sheet, so the simplification stays; it understates value by roughly ₹25 a share and errs in the conservative direction.

## What I would do with more time

* Value Jio Platforms with its own DCF rather than a multiple, using the subscriber, ARPU and capex disclosures in the quarterly releases — the listing prospectus will make this possible in full.
* Replace the trailing O2C multiple with a mid-cycle EBITDA built from a refining-margin and petrochemical-delta history, which is what a sector analyst would do.
* Pull Jio's and Retail's balance sheets from their standalone filings so net debt is attributed by segment before the stakes are applied.
* Add the listed investments and the New Energy capital employed at cost as separate lines, once the annual report gives a clean breakdown of the ₹1.5 lakh crore of non-current investments.

## Sources

Reliance Industries Limited, audited consolidated segment information for the quarter and year ended 31 March 2026 (media release dated 24 April 2026). Ownership stakes from the annual report. Peer EV/EBITDA multiples from Yahoo Finance as at 18 September 2026, listed individually on the Peers sheet with inclusions and exclusions stated. Share price from the NSE close on 18 September 2026. Adjusted net debt is the same figure as Project 8, and the Checks sheet enforces the tie. The holding-company discount is a stated input.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
