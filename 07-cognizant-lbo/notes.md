# Project 7 · Cognizant Technology Solutions — Leveraged Buyout

**Files:** `Cognizant_LBO.xlsx` (live model, 281 formulas, zero errors, 16 checks) · `Cognizant_LBO.pdf` · `build_model.py`

**As of:** FY2025 10-K (year to 31 December 2025) · market data 12 Sep 2026 (share price US$62.59) · US$ millions.

**The transaction is hypothetical.** Cognizant is a real, listed company. Nothing here suggests it is or has been in a sale process.

## What it is

A sponsor take-private of Cognizant: sources and uses, a five-year debt schedule with mandatory amortisation and a full cash sweep, returns by money multiple and IRR, a value-creation bridge, two sensitivity grids, and a solve for the highest price a financial buyer could pay and still clear its hurdle.

Cognizant was chosen so that the portfolio's view of IT services is complete. Project 3 says what the public market pays for these businesses, Project 4 says what strategic acquirers have paid, and this says what a financial buyer could pay. Those are three different numbers for the same kind of asset, and being able to set them beside each other is the point.

## Result

| | |
|---|---|
| Offer price per share | US$81.36 (30% premium) |
| Transaction enterprise value | US$38,955m · **entry 9.58× EBITDA** |
| Debt raised | US$21,347m — 3.00× term loan B, 2.25× senior notes, **5.25× total** |
| Sponsor equity cheque | US$18,925m (45% of total capitalisation) |
| Exit (year 5, 9.00× EBITDA) | EV US$47,839m, equity US$37,883m |
| **Money multiple / IRR** | **2.00× / 14.9%** |
| Versus the 20% hurdle | **(5.1pp) — the deal misses** |
| **Maximum price at a 20% return** | **US$73.78, an 18% premium** |

Leverage falls from 5.25× to 1.9× over the hold; interest cover rises from 2.5× to 5.0×. Of the US$18,959m of value created, US$11,391m is debt repaid and US$11,245m is EBITDA growth, against US$(2,361)m lost to the deliberate multiple contraction.

## Reading it honestly

* **The deal misses at a 30% premium, and the useful output is the price that doesn't.** A sponsor could pay US$73.78 — an 18% premium — and still make 20%. That number is the floor under Cognizant's share price in any sale process, and it is the ninth bar Project 5's notes asked for. It also sits 18% above where the stock trades, which is itself the finding: the market is pricing Cognizant below what a leveraged buyer could justify.
* **Cognizant is buyable only because the sector de-rated.** It trades at roughly 7.3× EBITDA unaffected, down from a 52-week high of US$87.03. At the 20× the sector carried in 2021 no sponsor could have levered it at all. The entry multiple is the single most important number in any LBO, and here it is doing most of the work.
* **The exit multiple is set below entry on purpose.** 9.00× out against 9.58× in. Assuming multiple expansion is the standard way an LBO model is made to produce a number someone wants to see; a check in this workbook fails the model if anyone later raises the exit multiple above entry.
* **The return comes from deleveraging and margin, not from financial engineering.** The bridge splits it explicitly: debt paydown and EBITDA growth contribute almost equally, and multiple change subtracts. A sponsor pitching this deal would have to defend 190bp of margin expansion over five years in a business facing an AI-driven pricing debate — which is the real risk, and it is an operating risk rather than a financing one.
* **Interest is charged on opening balances, not average balances.** Average-balance interest makes the model circular: interest drives cash flow, cash flow drives the sweep, the sweep drives interest. Resolving that needs iterative calculation switched on, and a workbook that only computes correctly when a setting is enabled will break on someone else's machine. The cost is a small overstatement of interest in years when debt falls fast, which is conservative.
* **The equity cheque is US$18.9bn, which no single fund writes.** In practice this is a consortium. That changes who signs, not the arithmetic, so the model treats the sponsor as one investor and says so.
* **The leverage sensitivity grid flatters high-leverage cases.** It flexes the opening debt quantum while holding the paydown path at base case, so more debt does not correctly feed through to more interest and a smaller sweep. It shows direction and rough magnitude, not a financeable structure. Stated on the sheet rather than left for a reader to discover.
* **Debt pricing is assumed, not quoted.** SOFR plus 350 on the term loan and an 8.5% coupon on the notes are plausible for a B-rated credit at this leverage in September 2026, but no lender has been asked.

## What I would do with more time

* Add a downside case — flat revenue and margin compression — and test the covenant headroom, which is the case a credit committee actually underwrites.
* Model a revolver for liquidity, and an accordion or delayed-draw tranche for bolt-ons, since services roll-ups are a common sponsor thesis here.
* Add management rollover and an option pool, both of which dilute the sponsor's return by two to three points and are routinely left out of student LBO models.
* Run the exit as a range of dates rather than a fixed year five; IRR is highly sensitive to hold period and a three-year exit at the same multiple looks very different.
* Feed the US$73.78 floor into Project 5's football field as a ninth bar.

## Sources

Cognizant Technology Solutions FY2025 Form 10-K, via the Yahoo Finance fundamentals series for CTSH. Share price and 52-week range from the NASDAQ close on 12 September 2026. SOFR and US Treasury yields, 12 September 2026. Debt pricing, fee levels and the operating case are stated assumptions on the Inputs sheet.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
