# Project 20 · HDFC Bank / HDFC Ltd — Capstone: Live-Deal Replication and Post-Mortem

**Files:** `HDFC_Merger_Capstone.xlsx` (live model, 138 formulas, zero errors) · `HDFC_Merger_Capstone.pdf` · `build_model.py`

**As of:** announcement 4 April 2022 · effective 1 July 2023 · outcome to 18 September 2026 · ₹ crore, per-share values on the pre-bonus basis.

## What it is

India's largest corporate transaction, replicated at announcement and judged three years on. The **Exchange Ratio** sheet reuses Project 6's contribution analysis — profit, net worth, loans, market value — with HDFC Ltd's 21% stake in the bank cancelled and its shareholders' contribution looked through, and sets the implied ratios against the agreed 42-for-25. **Accretion** builds the pro forma FY2022 EPS, book value and ROE. **Regulatory Drag** does what a pitch book puts on a back page: the cash reserve, statutory liquidity and priority-sector requirements that land on a mortgage lender's balance sheet when it becomes a bank's, net of the deposit-funding saving as it arrives. **Post-Mortem** sets the announcement case against reported FY2023–FY2026 results, the share price against Nifty Bank, and the price-to-book de-rating against the cost of equity the price implies, using Project 11's frame in reverse.

## Result

| | |
|---|---|
| Exchange ratio · value per HDFC Ltd share · premium to its close | 1.68× · ₹2,531 · **3.3%** |
| HDFC Ltd holders' ownership of the combined bank | **41.0%** (announced: 41%) |
| Their look-through contribution: profit · net worth · loans · market value | **42.4% · 47.4% · 44.2% · 40.2%** |
| Implied ratio on market value · on profit · on book | 1.63× · 1.78× · 2.17× |
| Pro forma EPS effect before the drag · book value accretion | −1.2% (neutral, as announced) · **+7.7%** |
| ROE: standalone bank → pro forma after the drag | **15.4% → 13.0%** |
| Low-yield assets forced by CRR and SLR on HDFC Ltd's ₹6.2 lakh crore of liabilities | **₹1.40 lakh crore** |
| After-tax drag in year one · as % of HDFC Ltd's profit | ₹4,006 crore · **29%** |
| Deposit-funding saving once half the bonds are refinanced | ₹6,237 crore pre-tax — larger than the drag, but slow |
| Reported ROE FY2023 (last standalone) → FY2026 | **15.7% → 12.8%** |
| EPS · book value per share growth FY2023–FY2026, compound | 5.0% · 12.4% |
| HDFC Bank TSR since 1 April 2022 · Nifty Bank price return | **+3.8% · +50.9%** — relative **−54 points** |
| Price / book: 1 April 2022 → today | **3.48× → 2.05×** |
| Cost of equity the price implies at 7% growth: 2022 → today | 9.4% → 10.1% |

## Reading it honestly

* **The ratio was set on market value, and it favoured the bank.** Once HDFC Ltd shareholders' 21% of the bank is counted as theirs — which it was — they contributed 42% of the combined profit and 47% of the net worth, and received 41% of the shares, which is what 40% of the combined market value implied. The bank's 3.5× book against the lender's roughly 2× did the work. Project 6 made the same point about the ENBD–Mashreq ratio: whoever has the higher multiple sets the terms on price and wins on book.
* **The accretion arithmetic was right and it answered the wrong question.** Book value per share has compounded at 12% since FY2023, as a 7.7% day-one accretion plus retained earnings would predict. EPS held up. But the same pro forma said ROE would fall by two to three points as ₹1.4 lakh crore of reserves and a bond-funded book were bolted on, and it did: 15.7% to 12.8%.
* **The market did what Project 11 says it should.** A bank's price is book value times a multiple set by ROE against the cost of equity. The bank's book grew and its ROE fell; the multiple went from 3.5× to 2.0×; the price is below where it was the day before the announcement while the bank index rose by half. The cost of equity the price implies has barely moved — 9.4% then, 10.1% now. The market re-priced the return, not the franchise. There is no mystery to explain, and no sentiment story needed.
* **The strategic case is not wrong, it is slow.** Replacing HDFC Ltd's ₹4.6 lakh crore of bonds with deposits is worth more than the regulatory drag in steady state — the model shows the saving exceeding the drag once half is refinanced. But it arrives at the pace of deposit growth, and a bank that had to hold loan growth below system to bring its loan-to-deposit ratio down from 110% cannot gather deposits faster than the system. The deal will be judged on FY2028 ROE, not FY2026.
* **What a pitch would have shown and what it would not.** The pitch shows the contribution table, the ownership split and the book accretion. It would not have led with a 29% year-one drag on the acquired profit, or with the ROE line, because those are the numbers that price the stock. The capstone's point is that the back page was the front page.
* **Stated and approximated inputs.** HDFC Ltd's carrying value of its bank stake, the FY2022 borrowings split, the regulatory yield gaps, the priority-sector shortfall, the refinancing share and the funding costs are stated; FY2026 net worth is derived from prior year plus profit less dividends; the Nifty Bank levels at the two intermediate dates are approximate while the endpoints are exact. Everything else is as announced or as reported.

## What I would do with more time

* Rebuild the pro forma from both companies' FY2022 annual reports on an acquisition-accounting basis — fair value of HDFC Ltd's assets, the intangible for the mortgage franchise — rather than the pooling basis the announcement used.
* Track the regulatory drag against the RBI glide path actually granted and the bank's disclosed CRR/SLR and PSL positions each quarter, so the drag is measured rather than modelled.
* Add the loan-to-deposit ratio, deposit growth and net interest margin by quarter since July 2023 — the three lines that will decide whether the ROE recovers.
* Run the Project 11 excess-return valuation on FY2028 consensus ROE to say what the shares are worth if the thesis works, and what it costs if it does not.

## Sources

Merger announcement and investor presentation of 4 April 2022 (exchange ratio, 41% ownership, effective date); FY2022 standalone accounts of HDFC Bank and HDFC Ltd as reported; HDFC Bank reported results FY2023–FY2026, with FY2026 net worth derived as stated. NSE closing prices and Nifty Bank levels via public market data on 18 September 2026; HDFC Bank per-share figures restated to the pre-bonus basis after the August 2025 1:1 bonus. Regulatory parameters as in force in 2022. Nothing here is a recommendation.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
