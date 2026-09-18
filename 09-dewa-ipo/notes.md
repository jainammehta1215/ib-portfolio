# Project 9 · DEWA — IPO Valuation and Pricing Rebuild (April 2022)

**Files:** `DEWA_IPO_Rebuild.xlsx` (live model, 191 formulas, zero errors) · `DEWA_IPO_Rebuild.pdf` · `build_model.py`

**As of:** deal announcements of 24 March and 6 April 2022, FY2021 figures from the prospectus summary · aftermarket to 18 Sep 2026 · AED millions.

## What it is

The largest IPO in UAE history, rebuilt from the public announcements. Three questions, in the order a syndicate desk asks them. **Valuation:** what was the company worth on FY2021 information, by dividend yield on the committed AED 6.2bn payout, by EV/EBITDA, and by a five-year DCF. **Pricing and book:** where the price sat in the AED 2.25–2.48 range, how a deal launched at 6.5% of the company was priced at 18%, how much of it was pre-committed to cornerstones, and whether the raw inputs reproduce the "37 times oversubscribed" headline. **Aftermarket:** the first-day rise, the money left on the table, and what an IPO investor has earned in the four and a half years since against what the business delivered. Two sensitivity grids cover the yield method and the DCF.

## Result

| | |
|---|---|
| Price range · final price | AED 2.25–2.48 · **AED 2.48**, the top |
| Market capitalisation at listing | AED 124bn — reproduced from price × shares |
| Stake sold: launched / priced | 6.5% / **18%** — a 2.8× upsizing |
| Gross proceeds | AED 22.3bn — reproduced |
| Cornerstones and strategics | AED 13.8bn, **62% of the deal** |
| Open-book coverage | **37.0×** — reproduced from AED 315bn demand over AED 8.5bn of open-book shares |
| Coverage of the whole deal, cornerstones included | 14.7× |
| Yield · EV/EBITDA · P/E at the final price | **5.0%** · 12.0× · 18.8× |
| Dividend-yield method | AED 2.25–2.76 per share |
| EV/EBITDA method (10–12.5×) | AED 2.00–2.60 |
| DCF (4.5% growth, 6.5% WACC) | AED 2.11 |
| Final price versus the midpoint of the three methods | **+7.6%** — priced full, not at a discount |
| First-day: open / close | AED 2.98 / **AED 2.87 (+15.7%)** |
| Money left on the table | **AED 3.5bn**, 15.7% of proceeds, ~9× the fees |
| Total return to an IPO investor to date, with dividends | 30.5%, **6.2% a year** |
| EBITDA growth delivered vs assumed | **10.9%** a year vs 4.5% |
| EV/EBITDA then vs now | 12.0× → 9.6× |

## Reading it honestly

* **The price was the dividend.** AED 6.2 billion a year at a 5.0% yield is AED 124 billion, and AED 124 billion over 50 billion shares is AED 2.48. The top of the range was reverse-engineered from a round yield, and the book-building confirmed a number chosen before the book opened. The model's fourth check demands this reproduce exactly, and it does.
* **On fundamentals the deal was priced full.** The multiple and DCF methods on the same FY2021 numbers land below the yield method, and the final price sits about 8% above the midpoint of the three. There was no underpricing in the usual sense. The 16% first-day rise came from a book filled at an average of 3% — investors who wanted a 5% yield in a rising-rate world and could not get enough of it in the allocation — not from a discount to value.
* **The book was three times bigger than the deal it was built for.** Demand of AED 315 billion against a launch size of about AED 8 billion let the seller nearly triple the offering and still report 37× coverage. The model also shows what the original 6.5% deal would have printed: close to 40×. Most of the enlarged deal went to cornerstones and strategics who had committed at the price, so the free float that actually traded on day one was far smaller than 18% suggests.
* **Size, not price, was the lever.** The range was fixed two weeks before pricing and already sat at or above value. Faced with that book, the seller could raise the price only to the range top, so it took size instead. The AED 3.5 billion left on the table is the cost of that decision, and it bought a much larger listed float. Whether that trade was right is a policy question for the Government of Dubai; the model only measures it.
* **Four years on, the story delivered and the multiple did not.** EBITDA has grown at 11% a year, more than double what a cautious DCF assumed, and the committed dividend has been paid on schedule. Yet the shares closed their first anniversary below the offer price and today sit only 10% above it. The multiple has de-rated from 12× to under 10× EBITDA, toward where global utilities trade. An IPO investor's 6% annual return has come almost entirely from the yield that was promised — which is exactly what a regulated monopoly sold on yield should have produced.
* **The peer multiples at April 2022 are stated inputs.** I did not have a reliable historical snapshot of utility multiples for that month, so the 10–12.5× range is a stated assumption; today's readings are shown on the Valuation sheet for context and are not used in the numbers. The fee rate and the DCF assumptions are likewise stated. The first-day prices are Yahoo Finance's unadjusted DFM data and differ slightly from some press reports of the close.

## What I would do with more time

* Rebuild the FY2021 balance sheet and cash flow from the full prospectus rather than the FAQ summary, so net debt is a reported figure instead of one derived from the disclosed 1.5× leverage.
* Source utility multiples and yields as of April 2022 from a data terminal and replace the stated range.
* Add the allocation split between the retail tranche, the qualified-investor tranche and the cornerstones, and model the retail oversubscription separately — the retail book was reportedly filled at an even lower rate.
* Run the same rebuild on the 2019 Saudi Aramco IPO, the other candidate for this project, where the tension between the sovereign's target valuation and the book is the entire story.

## Sources

DEWA PJSC: price-range announcement (24 March 2022), final-price announcement (6 April 2022) and the IPO FAQ published on dewa.gov.ae, for the range, share counts, cornerstone commitments, demand, proceeds, and FY2021 revenue, adjusted EBITDA, net income, leverage and dividend policy. Aftermarket prices from the Dubai Financial Market via Yahoo Finance, unadjusted. FY2022–FY2025 financials via Yahoo Finance for DEWA.AE. Peer multiples at April 2022, target yields, DCF assumptions and the fee rate are stated inputs.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
