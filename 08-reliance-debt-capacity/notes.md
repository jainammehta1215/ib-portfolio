# Project 8 · Reliance Industries — Debt Capacity and Capital Structure

**Files:** `Reliance_Debt_Capacity.xlsx` (live model, 253 formulas, zero errors) · `Reliance_Debt_Capacity.pdf` · `build_model.py`

**As of:** audited FY2026 consolidated results (year to 31 March 2026), published 24 April 2026 · market data 18 Sep 2026 · ₹ crore.

## What it is

A capital-structure study in three parts. The **Credit Profile** restates the balance sheet the way a rating agency would, adding lease liabilities and the deferred spectrum instalments owed to the government to the company's own net-debt figure, and stripping the disposal gain out of EBITDA. **Debt Capacity** sizes how much more the company could borrow at each rating category as the tightest of three constraints — leverage, cash interest cover and FFO-to-debt — and names which one binds. **Optimal Structure** unlevers the observed beta, re-levers it across debt weights, looks up the implied rating and cost of debt from the same grid, and reports where WACC bottoms out. A **Liquidity** sheet tests twelve-month sources against uses, and two sensitivity grids rebuild every constraint in each cell.

## Result

| | |
|---|---|
| Net debt / EBITDA, company definition | 0.60× (matches the results release) |
| **Adjusted net debt / adjusted EBITDA** | **1.25×** — leases and spectrum add ₹123,000 crore |
| EBITDA / finance costs (P&L) | 7.4× |
| EBITDA / cash interest paid | 5.0× — interest paid includes the capitalised portion |
| FFO / adjusted net debt | 53% |
| Free cash flow after dividends | ₹40,400 crore on capex of ₹144,300 crore |
| **Incremental capacity at an A profile** | **₹44,300 crore (US$5.3bn)** — FFO-to-debt binds |
| Incremental capacity at BBB | ₹190,300 crore (US$22.8bn) |
| WACC today | 12.88% at 13% market-value debt |
| Minimum WACC in the grid | 12.84% at 20% debt |
| Twelve-month sources / uses | 1.5× (1.9× if short-term lines are repaid rather than rolled) |

## Reading it honestly

* **The headline leverage is half the agency number.** The company reports net debt of ₹124,717 crore and 0.6× EBITDA. Add the ₹23,600 crore of leases and the ₹99,600 crore of spectrum instalments — both interest-bearing, both treated as debt by S&P and Moody's — and the ratio is 1.25×. That is still a clean A-category number, but it is the one to defend, and the model reconciles to the published figure before it restates it so nobody can say the starting point was invented.
* **Cash interest cover is the weak reading, and the reason is instructive.** On finance costs charged to the P&L, cover is 7.4×. On interest actually paid it is 5.0×, because interest on close to ₹2 lakh crore of projects under construction is capitalised and never touches the income statement. A lender looks at the cash figure. The published ratings sit between the two readings, which is what a committee does when the ratios are strong but the capex programme is heavy.
* **Capacity at A is smaller than the leverage ratio implies.** Leverage alone would permit another ₹1 lakh crore of net debt. FFO-to-debt binds first and caps the headroom near ₹44,000 crore. The capacity is real, but with capex running at three-quarters of EBITDA it exists to protect the investment programme through a downturn, not to fund a buyback. Dropping to BBB would release ₹1.9 lakh crore; the company's stated aim of staying two notches above the sovereign says it will not.
* **There is no WACC argument for more debt.** The curve is flat from 10% to 30% debt — the tax shield and the rising cost of equity nearly cancel — and the company is already at 13%. The saving between today's structure and the grid minimum is four basis points. Beyond 30% the implied rating falls to BB, the spread jumps, and WACC rises faster than the shield can compensate. The optimum is a range, and Reliance is inside it.
* **Liquidity is the real story.** Cash and liquid investments of ₹2.5 lakh crore plus one year's operating cash flow cover every use for the next twelve months — maturities, capex held flat, interest, the dividend and the spectrum instalment — with room to spare, and still cover them if every short-term line is repaid rather than rolled. That, more than any ratio, is why the rating sits above India's sovereign.
* **The rating grid is a stated input.** The thresholds are illustrative levels adapted from published methodologies; they are not the agencies' actual matrices. Both the implied-rating rows and the capacity numbers should be read as a screen. The cover says so.

## What I would do with more time

* Replace the single twelve-month maturity bucket with the full debt maturity ladder from the annual report, and add the currency split — a large share of the borrowings are in US dollars, which matters for both refinancing and the cost-of-debt assumption.
* Model FFO-to-debt with the interest feedback loop rather than the one-pass approximation, and use taxes actually paid (₹9,600 crore) rather than the tax charge (₹27,600 crore) once the deferred-tax unwind is understood.
* Run the capacity study on the Jio Platforms and Reliance Retail balance sheets separately. The Jio listing announced in the FY2026 results changes where debt can sit in the group, and a holding-company capacity number blends two very different credit profiles.
* Add the asset-light versus asset-heavy split of EBITDA so the capacity can be stress-tested against an O2C margin cycle, which is what actually moves Reliance's credit metrics year to year.

## Sources

Reliance Industries Limited, audited consolidated financial results and media release for the quarter and year ended 31 March 2026, dated 24 April 2026 — income statement, balance sheet, cash flow statement, segment information and the company's own net-debt reconciliation. Share price from the NSE close on 18 September 2026 via Yahoo Finance. Ratings as disclosed by the company (S&P Global A-, Moody's Baa1). Risk-free rate, equity risk premium, beta, rating thresholds and credit spreads are stated inputs on the Inputs sheet.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
