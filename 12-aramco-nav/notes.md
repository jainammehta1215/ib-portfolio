# Project 12 · Saudi Aramco — Energy NAV and Price-Deck Sensitivity

**Files:** `Aramco_NAV.xlsx` (live model, 201 formulas, zero errors) · `Aramco_NAV.pdf` · `build_model.py`

**As of:** Annual Report 2025 and FY2025 results (year to 31 December 2025) · Tadawul close 18 Sep 2026 · US$ billions, per-share values in SAR.

## What it is

An energy company is valued on the cash its reserves will generate, and for Aramco that cash is shaped above all by the state's take. The model builds a per-barrel netback under the Kingdom's published fiscal terms — a royalty on crude that steps from 15% to 45% to 80% as Brent passes US$70 and US$100, and a 50% upstream income tax — applies it to 12.9 million boe a day, and capitalises it over the 52-year proved-reserve life as an annuity. Downstream is valued at a multiple of adjusted EBIT plus depreciation, corporate costs are capitalised, net debt and minorities are deducted, and NAV per share is set against the price. The **Price Deck** sheet then runs the whole company from US$50 to US$110 Brent and solves, in closed form and verified by substitution, for the oil price at which the base dividend is exactly covered and the Brent the share price implies. A calibration check rebuilds Aramco's reported FY2025 free cash flow from the netback.

## Result

| | |
|---|---|
| Blended realisation per boe at US$70 Brent | US$60.62 (liquids at Brent, gas and NGL at US$15) |
| After-tax cash netback per boe | **US$20.21** — Aramco keeps 33% of the barrel |
| Government take on a marginal dollar: below US$70 / US$70–100 / above US$100 | 57.5% / 72.5% / **90%** |
| Upstream value (52-year annuity at 8%) | US$1,168bn — **US$4.73 per boe of proved reserves** |
| Downstream at 7× EBITDA | US$140bn |
| **NAV** | **US$1,165bn · SAR 18.06 per share** |
| Share price · upside/(downside) | SAR 25.56 · **(29%)** |
| Brent the price implies at an 8% discount rate | US$118 |
| **Discount rate the price implies at US$70** | **5.7%** |
| Model rebuild of FY2025 free cash flow vs reported US$85.4bn | within 6% |
| Group free cash flow at US$70 · base dividend cover | US$90bn · **1.03×** |
| **Brent at which the base dividend is exactly covered** | **US$68.3** |
| Brent at which group free cash flow is zero | US$16 |

## Reading it honestly

* **Aramco keeps a third of the barrel and ten cents of a spike.** At US$70 the after-tax netback is US$20 on a US$61 blended realisation. Above US$100 the 80% royalty and the 50% tax leave the equity with ten cents of every incremental dollar. The shares are a claim on volume and cost, not on the oil price — which is exactly why the price-deck grid is so flat at the top and so steep at the bottom.
* **The NAV sits well below the price, and the gap is a discount rate, not an oil forecast.** On a flat US$70 deck at 8% the shares trade 40% above NAV. For that price to be justified at 8%, Brent would need to average close to US$120 in perpetuity. The alternative reading is that the market discounts Aramco's cash at under 6% — plausible for a stock with a 5% base yield, a sovereign owner holding most of the float, and a dividend that has been paid through every cycle. The model reports both readings; the second is the credible one.
* **Reserves are worth a fraction per barrel of a major's, by design.** US$4.73 per boe against the US$8–15 a Western integrated typically commands is not a discount to be closed. It is the arithmetic of who owns the barrel. And the 52-year reserve life adds almost nothing: at 8% the annuity is 98% of a perpetuity, so barrels beyond year thirty are worth close to nothing today.
* **The dividend is the live question, and the model puts a number on it.** After a US$52bn capex programme, group free cash flow covers the US$87.5bn annualised base dividend only above about US$68 Brent. That is why the performance-linked dividend was cut to almost nothing in 2025, why the 2024 payout was partly funded from the balance sheet, and why the company has turned to asset sales such as the Jafurah midstream stake. The Price Deck sheet shows a US$30bn shortfall at US$50 and a US$35bn surplus at US$100.
* **The calibration is the check that earns the rest.** A netback built from stated cost and capex assumptions reproduces the audited FY2025 free cash flow within 6%. Without that, the NAV would be an assertion. With it, the deck and the break-evens inherit the credibility of the reported numbers.
* **What is simplified.** Production is held flat, when gas is guided to grow sharply toward 2030; the royalty is applied to all liquids rather than crude alone; the gas realisation is a blended input; and minorities are taken at book although the pipeline lease vehicles have contractual cash-flow claims that book value understates. Each is flagged on the Inputs sheet and none changes the shape of the answer.

## What I would do with more time

* Split the upstream into crude, gas and NGL streams with their own realisations, costs and growth paths, so the Jafurah and Master Gas System expansion shows up as value rather than being folded into a flat production assumption.
* Replace the annuity with a year-by-year model that lets production, price and capex vary, and reconcile it to the company's own disclosed sensitivity of earnings to a US$1 move in Brent.
* Value the minorities on the pipeline lease vehicles and Jafurah midstream on their contracted tariffs rather than at book.
* Add a dividend-capacity sheet that models the balance sheet under each price deck — gearing, debt issuance and asset-sale proceeds — since the state's cash need, not the NAV, is what actually sets the payout.

## Sources

Saudi Aramco Annual Report 2025 and the fourth-quarter and full-year 2025 results press release: total hydrocarbon production, liquids production, proved reserves, average realised crude price, upstream and downstream adjusted EBIT, capital expenditure, free cash flow, dividends and the Q4 2025 base dividend. Income statement and balance sheet in SAR via Yahoo Finance for 2222.SR, converted at the 3.75 peg. Share price and 52-week range: Tadawul close, 18 September 2026. The royalty schedule and income-tax rates are as published by the Kingdom in 2020. Lifting costs, sustaining capex per boe, the gas realisation, the discount rate, downstream depreciation and multiple, and corporate costs are stated inputs.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
