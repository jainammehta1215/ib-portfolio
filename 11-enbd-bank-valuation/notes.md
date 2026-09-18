# Project 11 · Emirates NBD — Bank Valuation: Dividend Discount and Excess Return

**Files:** `ENBD_Bank_Valuation.xlsx` (live model, 239 formulas, zero errors) · `ENBD_Bank_Valuation.pdf` · `build_model.py`

**As of:** FY2025 audited accounts (year to 31 December 2025), the same figures as Project 6 · market data 18 Sep 2026 · AED millions.

## What it is

Banks cannot be valued on enterprise value, because deposits and wholesale funding are the raw material of the business rather than its financing. So this is an equity-side valuation built two ways on one forecast. A five-year equity roll-forward under clean surplus — closing book equals opening book plus earnings less dividends, enforced cell by cell — fades return on equity from the FY2025 level of 17% toward a 14% terminal level and raises the payout so that retained earnings fund exactly the growth assumed. A **dividend discount model** discounts the dividends and a Gordon terminal value. An **excess-return model** starts from book value today and adds the present value of what the bank earns above its cost of equity. The **Justified Multiples** sheet collapses the same logic to one line, (ROE − g)/(ke − g), and runs the share price backwards into the return on equity and the cost of equity it implies.

## Result

| | |
|---|---|
| Return on average equity, FY2025 | 17.3% |
| Book value per share · price / book | AED 22.91 · **1.35×** |
| Cost of equity | **10.0%** (4.1% Treasury + 0.9% country risk + 1.0 × 5.0% equity premium) |
| Terminal ROE · growth · derived payout | 14.0% · 3.0% · 78.6% |
| **Dividend discount value per share** | **AED 41.71** |
| **Excess-return value per share** | **AED 41.71** — identical, by construction |
| Upside to the AED 30.98 price | 34.7% |
| Implied P/B at the model value | 1.82× |
| Share of value that is book already held | 55% |
| Terminal value as % of the DDM value | 86% |
| Justified P/B at the terminal ROE | 1.57× |
| Sustainable ROE the price implies at a 10% cost of equity | 12.5% |
| Cost of equity the price implies at a 14% terminal ROE | 11.1% |

## Reading it honestly

* **Two methods, one number, and that is the point of the exercise.** The DDM and the excess-return model agree to the fils, and the Checks sheet fails the workbook if they ever stop agreeing. They must agree under clean surplus with a consistent terminal state; the usual reason bank models show two different answers is that someone typed a terminal payout that does not match the terminal growth. Here the payout is derived — 1 − g/ROE — not entered.
* **The market is pricing a harder fade than this model assumes.** At a 10% cost of equity the price of 1.35× book is consistent with a sustainable ROE of about 12.5%, against 17% earned last year and 14% assumed in the terminal. Alternatively, if 14% is right, the market is discounting at 11.1%, roughly a point above the country-risk build-up. Both readings say the same thing: investors want compensation for something a textbook cost of equity does not capture — the oil-price linkage of UAE deposit growth, the concentration of a domestic bank, regional risk. That premium is the whole gap between model and price.
* **The excess-return layout is the useful one for a bank.** Just over half of the value is book equity already on the balance sheet; the rest is the capitalised spread of ROE over the cost of equity. The share price sits only a third of the way from book value to the model value, which is a precise statement of how little of the future spread the market is willing to pay for.
* **Terminal value carries 86% of the answer, which is normal and uncomfortable.** A bank that retains most of its earnings for five years and then settles into steady state will always be valued mostly on that steady state. The two sensitivity grids make the dependence explicit: half a point on the cost of equity or a point on the terminal ROE each move the value by around two dirhams.
* **The ROE fade is a judgement, and a benign one.** From 17% to 14% over five years assumes rate cuts compress net interest margins and that capital keeps building. A harsher fade — to 12%, say — takes the value to the low thirties and close to the price. That is a legitimate view; the model makes clear it is the view the market holds.
* **The historical column is not clean surplus.** FY2025 closing equity is the reported figure, so that year absorbs other comprehensive income, the AT1 coupons and any capital movements. From FY2026 the roll-forward is exact. Someone extending the model should treat AT1 distributions explicitly, since they reduce the earnings available to ordinary shareholders.

## What I would do with more time

* Build the earnings forecast bottom-up — loan growth, net interest margin under a rate-cut path, cost of risk, cost-to-income — rather than as a return on opening equity, and let ROE fall out.
* Add a regulatory capital constraint so the payout is bounded by the CET1 target, not only by the growth assumption; ENBD's capital position is strong, so it would likely bind the other way and raise the payout.
* Run the same two models on the UAE bank peer group (FAB, ADCB, DIB, Mashreq) and plot P/B against ROE, which is how the sector is actually traded.
* Treat the AT1 instruments explicitly in both the earnings and the equity base.

## Sources

Emirates NBD Bank PJSC FY2025 annual accounts, via the Yahoo Finance fundamentals series for EMIRATESNBD.AE — the same figures as Project 6, which the Checks sheet enforces. Share price and 52-week range from the Dubai Financial Market close on 18 September 2026. The US Treasury yield, the UAE country risk premium and the mature-market equity risk premium are September 2026 market data and Damodaran's country tables; beta, the ROE path, the terminal ROE and the long-run growth rate are stated inputs.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
