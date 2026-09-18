# Project 13 · Emaar Properties — Real-Estate NAV and Development Appraisal

**Files:** `Emaar_NAV.xlsx` (live model, 135 formulas, zero errors) · `Emaar_NAV.pdf` · `build_model.py`

**As of:** FY2025 results announcement (12 February 2026) and FY2025 balance sheet · DFM prices 18 Sep 2026 · AED millions.

## What it is

A property company is valued on what it owns, so this is a **net asset value** built in four blocks. The UAE development business is the present value of the margin embedded in Emaar Development's AED 125bn sold-but-unrecognised backlog, after tax, plus the 370 million square feet of UAE land at a stated value per square foot, with Emaar Properties' 80% share taken and the block cross-checked against the listed Emaar Development price. International development is treated the same way on the remaining backlog and land. Malls and commercial leasing, and hospitality and leisure, are capitalised at stated yields. Half of the group's cash is treated as customer advances held for construction rather than free cash, borrowings are deducted, and the NAV per share is set against the price to measure the discount the market applies.

A separate **Development Appraisal** runs the calculation a developer does before bidding for a plot: gross development value less construction, soft costs, marketing and finance, less a target profit on total cost including land, leaves the residual land value. It is illustrative, and it exists to show why land is the most geared line in any property NAV.

## Result

| | |
|---|---|
| A. UAE development (backlog PV AED 39.9bn + land AED 16.7bn) × 80% | AED 45.2bn — **4.5% above** the listed Emaar Development value × 80% |
| B. International development | AED 8.0bn |
| C. Malls and commercial leasing at a 7.0% yield (14.3× EBITDA) | AED 82.7bn |
| D. Hospitality, leisure and entertainment at 8.5% | AED 27.2bn |
| Gross asset value | AED 163.1bn — recurring assets are 67% |
| Free cash (50% of AED 52.6bn) less borrowings | AED 15.7bn |
| **Net asset value** | **AED 178.8bn · AED 20.23 per share** |
| Share price · price / book | AED 11.80 · 1.11× |
| **Discount to NAV** | **42%** |
| Market cap as % of recurring assets plus free cash alone | 83% — the price ascribes nothing to development |
| Illustrative tower: GDV · residual land value | AED 1,181m · AED 460m (AED 767 per sq ft GFA, 39% of GDV) |
| Break-even sale price · price cushion | AED 2,000 per sq ft · 17% |

## Reading it honestly

* **The market pays for the malls and the hotels and gets the development business free.** Market capitalisation of AED 104bn is 83% of the recurring-income assets plus free cash alone. At today's price the AED 155bn backlog and 660 million square feet of land are being valued at less than nothing. That is the finding, and it is not unusual: UAE developers have traded at 30–50% discounts to bottom-up NAV for most of their listed history.
* **The external check passes, which matters more than the discount.** The model's value for the UAE development block, built from backlog margin and land, lands within 5% of what the Dubai market pays for Emaar Development's listed shares. A NAV that cannot be tested against anything is an opinion; this block can be, and it holds.
* **The cash haircut is the honest line.** Emaar's AED 52.6bn of cash includes off-plan collections that sit in escrow to fund construction of the backlog. Counting it all as free cash is the most common way a developer NAV is inflated. The 50% haircut is a stated judgement; at zero haircut NAV rises by AED 3 per share, at 100% it falls by the same.
* **The two halves deserve different multiples and get them.** Dubai Mall and the leasing portfolio are near-bond assets at a 7% yield, roughly 14× EBITDA. The backlog is discounted at 11% over three and a half years and taxed, and the margin assumed (42%) sits below the 52% Emaar Development reports, to allow for the cost inflation a late-cycle backlog tends to meet. The first sensitivity grid flexes both.
* **The appraisal explains the discount.** Land is a residual, and the gearing runs both ways: at the base case the plot is 39% of GDV, and a 10% fall in sale prices takes about a fifth off the land value. Apply that gearing to a land bank carried at cost and it becomes clear why the market refuses to capitalise it at full value until the cycle has proved itself.
* **What is not here.** Emaar Misr and the Indian business have their own minorities that block B ignores; the international backlog margin is a round-number assumption; hospitality EBITDA is split from recurring EBITDA using an assumed margin because the full-year release does not give it. None of these moves the per-share value by more than a dirham, and each is flagged on the Inputs sheet.

## What I would do with more time

* Replace the single backlog margin and delivery period with a project-by-project schedule from Emaar Development's disclosures, which would also give a proper cash-flow profile for the escrow balances.
* Value the land bank by master-plan and location rather than at one blended rate per square foot, using recent Dubai Land Department plot transactions.
* Add the investment-property fair values from the annual report notes — Emaar carries investment property at cost under IAS 40 and discloses fair value in the notes, which would replace the yield capitalisation with an audited number.
* Extend the appraisal into a phased cash flow with an IRR, so the profit-on-cost target can be compared with the discount rate used in the NAV.

## Sources

Emaar Properties PJSC results announcements for 9M 2025 (6 November 2025) and FY2025 (12 February 2026): revenue backlog, recurring revenue and EBITDA, the mall and hospitality split, land bank and credit ratings. Emaar Development PJSC FY2025 results for its EBITDA and backlog. Balance sheets and share prices via Yahoo Finance (EMAAR.AE, EMAARDEV.AE), DFM close 18 September 2026. Yields, margins, discount rate, land values, the cash haircut and every appraisal input are stated assumptions on the Inputs sheet.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
