# Project 17 · Kellanova / Mars — Fairness Opinion Replication

**Files:** `Kellanova_Fairness_Opinion.xlsx` (live model, 148 formulas, zero errors) · `Kellanova_Fairness_Opinion.pdf` · `build_model.py`

**As of:** DEFM14A filed 26 September 2024 (SEC accession 0001193125-24-226970) · valuation date 30 June 2024 · US$.

## What it is

When Mars agreed to buy Kellanova for US$83.50 a share in cash in August 2024, Kellanova's board received fairness opinions from Goldman Sachs and Lazard, and the proxy statement disclosed how each was built: the management projections, the discount rates, the multiple and premia ranges, the reference prices, and the per-share range each analysis produced. This project takes only what the proxy discloses and asks whether those ranges can be reproduced. Two inputs the proxy describes but does not quantify — fully diluted shares and adjusted net debt — are back-solved from a single Goldman range and then held fixed for every other analysis, so each of the remaining eleven is a genuine test. The **Football Field** places the offer against all twelve, and the **Checks** sheet demands that each rebuilt range land within a stated tolerance of the proxy's.

## Result

| Analysis (proxy range) | Rebuilt | Deviation |
|---|---|---|
| Back-solved diluted shares · adjusted net debt | 351.2m (2% above the 344.7m basic count) · US$6,389m | — |
| Goldman DCF, 6–7% WACC, 12–14× terminal ($68.42–83.82) | **$67.79–83.19** | $0.63 |
| Goldman implied perpetuity growth (1.0–2.6%) | 1.3–2.9% | brackets |
| Goldman PV of future share price ($60.93–90.01) | $60.65–86.37 | $3.64 (dividend path is a stated input) |
| Goldman premia, undisturbed ($71.80–83.13) and 52-week high ($67.39–77.47) | exact | $0.00 |
| Lazard DCF, 6.25–6.75%, 1.5–2.0% growth ($69.60–79.10) | $61.07–78.76 | $0.34 at the top; **$8.53 at the bottom** |
| Lazard EBITDA comps, 11–13.5× 2025E ($57.20–73.80) | $57.01–74.10 | $0.30 |
| Lazard P/E comps, 14–19× 2025E ($55.90–75.70) | $55.86–75.81 | $0.11 |
| Lazard precedents, 15.5–17.5× LTM ($78.10–90.20) | $75.58–87.68 on back-solved LTM EBITDA of US$2,125m | $2.52 |
| Lazard premia, 20–35% ($75.60–85.10) | $75.58–85.02 | $0.08 |
| **Offer multiples** | **16.0× 2024E · 14.9× 2025E · 16.8× LTM EBITDA** · 32.6% premium | |
| Offer against the twelve ranges | above 7 · inside 5 · below 0 | |

## Reading it honestly

* **The proxy discloses enough to rebuild the opinion, and that is the finding.** Goldman's DCF comes back within 63 cents a share, Lazard's comparables within 30 cents, both sets of premia to the cent, and the P/E range within 11 cents — from projections, rates and ranges printed in a public document plus two numbers back-solved from a third. A reader who wants to know what a board was told does not have to take the banker's word for it.
* **The one miss is instructive rather than embarrassing.** Lazard's DCF low end cannot be reached by pairing its highest WACC with its lowest growth rate on 2027E free cash flow; it needs a terminal cash flow of about US$1,728m, 13% above the projection's US$1,533m and close to 2027E NOPAT. The likely reason is that Lazard normalised the terminal year — removing the working-capital build that depresses 2027E free cash flow — which is standard practice and not stated in the proxy. The sheet back-solves the figure and says so; the check for that analysis tests the high end and records the low-end gap.
* **The offer is a full price by every measure the advisors used.** It sits at the top of Goldman's DCF range and above Lazard's, above the 75th percentile of precedent premia, three to four turns of EBITDA above the best-rated snacking peers, and inside only the precedent-transaction ranges — which run to 21.5× because they include the 2015–2018 food deals and Mars's own Wrigley acquisition at 17.6×. That is what a fairness opinion on a negotiated strategic deal looks like: the intrinsic analyses bracket the price, the market analyses sit beneath it, and the precedents are the one place a higher number can be found.
* **Two advisors agreeing is not two opinions.** Both used the same management projections and the same balance sheet, so their ranges overlap by construction. Where their methods differ — a terminal multiple against a perpetuity growth rate — they land within a few dollars, and the model shows the bridge: Lazard's midpoint terminal value is 12.0× 2027E EBITDA, the bottom of Goldman's range.
* **What cannot be rebuilt is the judgement.** Why 12–14× rather than 11–15×; why 14–32% rather than the full interquartile spread; why 2025E EBITDA for the comparables but FY1 for the precedents. The proxy gives the ranges and the phrase "professional judgment and experience". The replication shows how much of the answer the ranges determine and how little room the judgement actually had.
* **Stated inputs.** The split of 2024E cash flow either side of 30 June (assumed half), the dividend path after 2024 (US$2.28 growing 3.5%), and the mid-year stub. Everything else is quoted from the proxy or derived from it as labelled.

## What I would do with more time

* Take adjusted net debt from Kellanova's 30 June 2024 10-Q — debt, cash, factored receivables, minorities and unconsolidated investments — and confirm the back-solved figure from the balance sheet rather than from a range.
* Rebuild Goldman's present-value-of-future-share-price analysis with the actual declared dividends and the projected year-end share counts the proxy mentions but does not print.
* Add the buyer's side: Mars is private, but the financing (US$29bn of bridge and bond debt) and the implied synergies needed to earn a cost of capital at 16× are public enough to model.
* Compare the two advisors' precedent sets deal by deal, which the proxy lists, to see how the choice of set drives the 12.1–21.5× against 15.5–17.5× ranges.

## Sources

Kellanova definitive proxy statement on Schedule 14A, filed 26 September 2024, retrieved from EDGAR: "Certain Financial Projections", "Opinion of Goldman Sachs & Co. LLC", "Opinion of Lazard Frères & Co. LLC", and the record-date share count. The replication is a learning exercise on a disclosed document; it expresses no view on either advisor's judgement and is not advice.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
