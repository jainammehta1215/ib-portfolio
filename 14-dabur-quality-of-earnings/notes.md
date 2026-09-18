# Project 14 · Dabur India — Quality of Earnings and Working-Capital Peg

**Files:** `Dabur_QoE.xlsx` (live model, 217 formulas, zero errors) · `Dabur_QoE.pdf` · `build_model.py`

**As of:** consolidated accounts FY2023–FY2026 (years to 31 March) · ₹ crore.

## What it is

A buy-side quality-of-earnings review, the report a buyer commissions before agreeing a price. It answers two questions the audited accounts do not. **What is the business earning from operations?** The headline EBITDA a seller's banker quotes — profit before tax plus finance costs plus depreciation — includes everything below the operating line; the **EBITDA Bridge** strips out interest on the treasury book, gains on investments and exceptionals and ties the result to reported operating income. **How much working capital does the business need?** The **Peg** sheet builds a normal level four ways — four-year average, average days on latest revenue, two-year run-rate, and a seasonality-adjusted figure — so the buyer pays for the business and not for whatever sits on the balance sheet on the closing date. Between them, **Cash Conversion** tests whether the adjusted earnings turn into cash and **Working Capital** shows the days trends the negotiation will turn on.

## Result

| | |
|---|---|
| Headline EBITDA, FY2026 | ₹3,015 crore (23.1% margin) |
| Less treasury income · other income incl. investment gains · exceptionals | (₹366) · (₹195) · +₹20 |
| **Adjusted operating EBITDA** | **₹2,473 crore — 82% of headline**, 19.0% margin, ties to reported operating income |
| Value of the gap at a 20× transaction multiple | **₹10,800 crore** — 2.1× the ₹5,148 crore treasury book that generates the income |
| Cash conversion (treasury income removed from both sides), FY2023 → FY2026 | 68% → 87% → 87% → **115%** |
| Capex / depreciation, FY2026 | 0.9× (1.6× in FY2023 — the capex cycle has ended) |
| Accruals ratio | −3.8% and improving: cash exceeds profit |
| Trade working capital, FY2023 → FY2026 | ₹686 → 424 → 363 → **₹83 crore** |
| Payable days · inventory days · receivable days, FY2026 | **157** · 123 · 20 (payables up from 126 in FY2023) |
| Cash conversion cycle | +18 days → **−14 days** |
| Peg by method: four-year average / days / two-year / seasonal | ₹389 / 425 / 223 / 425 crore |
| **Proposed peg (equal weights)** | **₹366 crore** |
| Adjustment if closing were today (closing less peg) | **(₹283 crore)** — buyer receives; 0.4% of market cap |
| Spread between methods — the negotiating range | ₹202 crore |

## Reading it honestly

* **Nearly a fifth of the headline EBITDA is income on cash, and that is the finding.** Dabur holds over ₹5,000 crore of cash and current investments. Interest and investment gains on it run to ₹560 crore a year and sit inside the EBITDA a seller would quote. In a transaction the cash transfers at face value; a buyer who also pays 20× for its income has paid for the same asset twice — and the second payment, ₹10,800 crore, is more than twice the cash pile itself. This is the most common adjustment in an Indian consumer QoE and here it is also the largest.
* **The operating earnings themselves are clean.** Once treasury income is removed from both EBITDA and operating cash flow, conversion has risen from 68% to 115% over four years, capex has fallen back to depreciation after a build-out, accruals are negative and improving, and margins have held near 19% through the FY2025 distributor-inventory correction that the company itself disclosed. There is nothing here to argue about; the arguments are elsewhere.
* **Working capital is where the negotiation is.** Payable days have lengthened by a month over four years while inventory days sat near 120, and trade working capital has fallen to almost nothing. The buyer's question is whether that payables extension is structural — supplier financing, renegotiated terms — or a closing-date effect that unwinds after completion. If it unwinds, the buyer funds around ₹500 crore of working capital in year one. That single diligence question is worth more than the peg arithmetic.
* **The peg spread is the size of the argument, and it is small.** The seller will propose the two-year run-rate (₹223 crore, the lowest peg, the smallest adjustment); the buyer the seasonally adjusted figure (₹425 crore). The ₹200 crore between them is 0.3% of market capitalisation. Relative to the EBITDA adjustment it is a rounding error, which is exactly the sense of proportion a QoE should give a deal team.
* **The other-income line is derived, not read.** The public feed does not carry Dabur's investment gains separately, so other income is the balancing item to reported profit before tax. A check bounds it at under 3% of revenue in every year. In a real engagement it would come straight from note 27 of the accounts.
* **The peg is built from year-ends because that is what Indian filers publish.** Balance sheets appear half-yearly, so the seasonality adjustment is a stated ratio from quarterly revenue rather than a measured monthly working-capital profile. The sheet says so. Twelve to twenty-four monthly closes from a data room would replace methods 1 to 4 with one number.

## What I would do with more time

* Rebuild the bridge from the annual report notes rather than the data feed: other income by category, exceptional items by nature, share-based payments, and any capitalised development or advertising costs.
* Add a proof of revenue — GST filings against reported sales, and the channel-inventory data the company disclosed during the FY2025 correction.
* Take the peg to monthly granularity from the data room and add a net-debt schedule with the debt-like items (provisions, deferred consideration, lease liabilities) that go with a cash-free, debt-free price.
* Extend to a normalised earnings and EPS bridge, so the equity-value effect of the adjustments is shown per share alongside the enterprise-value effect.

## Sources

Dabur India Limited consolidated financial statements FY2023–FY2026 via Yahoo Finance for DABUR.NS, including the quarterly revenue series used for the seasonality ratio. Other income is derived as the balancing item to reported profit before tax and bounded by a check. Total assets for the accruals ratio are approximated and the ratio is directional. The transaction multiple and the peg weights are stated inputs. The review is a portfolio exercise on public information; Dabur is not, to the author's knowledge, in any sale process.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
