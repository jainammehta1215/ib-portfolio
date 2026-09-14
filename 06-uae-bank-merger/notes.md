# Project 6 · Emirates NBD / Mashreqbank — Merger Model

**Files:** `ENBD_Mashreq_Merger_Model.xlsx` (live model, 180 formulas, zero errors) · `ENBD_Mashreq_Merger_Model.pdf` · `build_model.py`

**As of:** FY2025 audited accounts (year to 31 December 2025) · market data 12 Sep 2026 · AED millions.

**The transaction is hypothetical.** Both banks are real, listed on the Dubai Financial Market and file public accounts. Nothing here suggests either is or has been in discussions.

## What it is

A full accretion/dilution model for Emirates NBD acquiring Mashreqbank: offer and consideration mix, purchase accounting with a credit mark and a core deposit intangible, a three-year earnings bridge with phased synergies and integration costs, tangible book dilution with its earnback period, the pro forma CET1 ratio, return on invested capital against a hurdle, and two live sensitivity grids. Bank deals are decided on tangible book dilution and regulatory capital before anyone looks at EPS, so the model is laid out in that order.

## Result

| | |
|---|---|
| Offer price per share | AED 418.63 (25% premium) |
| Equity purchase price | AED 83,980m |
| Consideration | 70% stock / 30% cash · exchange ratio 9.422 |
| Target shareholders' ownership of the combined bank | 23.0% |
| **Price / tangible book paid** | **2.16×** (acquirer trades at 1.41×) |
| Goodwill created | AED 45,143m — 54% of the price paid |
| EPS accretion, years 1 / 2 / 3 | **(8.8%) / (3.0%) / (2.3%)** |
| Synergies required merely to break even | **44.6%** of target opex, against 25% assumed |
| **Tangible book value per share dilution** | **(16.7%)**, AED 3.68 per share |
| Earnback | **Never** — earnings never exceed standalone |
| Pro forma CET1 ratio | 13.0%, down 2.0pp, 1.98pp above the minimum |
| Return on invested capital | 8.6% against a 12% hurdle — **(3.4pp)** |

## Reading it honestly

* **The deal fails, and that is the output.** Emirates NBD trades at 1.41× tangible book. Mashreqbank trades at 1.73×, and a 25% premium takes the price paid to 2.16×. An acquirer issuing its own stock at 1.4× to buy assets at 2.2× transfers value to the seller the moment the deal is announced, and no synergy phasing reverses it. A merger model that only ever produces reasons to proceed is not being used properly; the most valuable thing this one does is quantify how far from working the transaction is.
* **Break-even is the number to look at, not accretion.** Reporting "2.3% dilutive at full synergies" invites the response that synergies could be a bit higher. Solving for the level that holds EPS flat gives 44.6% of target operating expenses — roughly double the 25% assumed and well above anything disclosed in comparable bank mergers. That framing closes the argument instead of inviting a negotiation over an assumption.
* **Revenue synergies are set to zero deliberately.** They are the standard way a bank merger model is made to work: a cross-sell assumption large enough to close the gap, with no mechanism behind it. Leaving the line in at zero makes the choice visible rather than hiding it. Adding revenue synergies would not rescue this deal anyway — the shortfall is close to AED 700m after tax in steady state.
* **The capital ratio is why the consideration is stock-heavy.** At 70% stock the pro forma CET1 lands at 13.0%, about 2 percentage points below standalone and just under 2 points above the regulatory minimum. A cash-heavy structure would not clear it. The second sensitivity grid shows the trade directly: more stock protects capital and worsens tangible book dilution, more cash does the reverse, and no mix makes a price above two times tangible book look cheap.
* **The credit mark is the largest judgement in the model and it is only an assumption.** 1.25% of an estimated loan book, where the loan book itself is approximated as 55% of total assets because the disclosed split was not pulled. In a real process this number comes out of loan-tape diligence and moves goodwill by billions. It is flagged on the Inputs sheet rather than buried.
* **Both banks' earnings are held flat at FY2025.** A model that also forecasts organic growth mixes two questions — whether the businesses grow, and whether the transaction adds value. Holding both flat isolates the second, which is the only one a merger model is competent to answer.
* **Capital inputs are stated, not sourced.** The CET1 ratio and the risk-weight assumptions are inputs set to plausible levels rather than lifted from each bank's Pillar 3 disclosure. Anyone using this externally should replace them first; the cover says so.

## What I would do with more time

* Replace the approximated loan book, deposit split and risk-weighted assets with the disclosed figures from both banks' Pillar 3 reports and note disclosures.
* Add a contribution analysis — each bank's share of pro forma earnings, assets, deposits and equity against its share of the combined market capitalisation — which is the cleanest way to argue an exchange ratio in a merger of this size.
* Model the deal as a merger of equals with no premium and a market-based exchange ratio, which is the structure the arithmetic here actually points toward.
* Add a crossover earnback calculation alongside the simple one, and a purchase-accounting accretion line from unwinding the credit mark, which typically flatters reported EPS in the years after a bank deal closes.

## Sources

Emirates NBD Bank PJSC and Mashreqbank PSC FY2025 annual accounts, via the Yahoo Finance fundamentals series for EMIRATESNBD.AE and MASQ.AE. Share prices from the Dubai Financial Market close on 12 September 2026. UAE federal corporate tax at 9%, in force since June 2023. Capital and risk-weight assumptions are stated inputs.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
