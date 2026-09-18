# Project 18 · Dubai World and Nakheel — Restructuring: Liquidity, Recoveries and Debt-for-Equity

**Files:** `Dubai_World_Restructuring.xlsx` (live model, 122 formulas, zero errors) · `Dubai_World_Restructuring.pdf` · `build_model.py`

**As of:** terms announced 20 May 2010 and approved 10 September 2010; Nakheel's plan of 2011 · US$ billions.

## What it is

The largest restructuring the Gulf has seen, rebuilt as the three analyses a restructuring team produces. **Liquidity:** why a group whose assets exceeded its debt had to ask for a standstill on 25 November 2009, and whether the eight-year plan could be met from asset sales. **Recoveries:** what each class of creditor actually received once maturities were pushed out five and eight years and coupons cut to 1% — nominally 100 cents, economically much less — laid out as a waterfall ranked by present-value recovery. **Debt-for-equity:** the Government of Dubai's conversion of US$8.9bn of claims into equity, what that equity was worth on 2010 asset values, and why the banks took paper instead. A sensitivity grid runs bank recovery against the discount rate and the share of Tranche B covered by the sovereign shortfall guarantee.

## Result

| | |
|---|---|
| Perimeter restructured | US$24.9bn: banks 14.4 · government 8.9 · other 1.6 |
| The December 2009 problem | US$3.52bn Nakheel sukuk due against US$0.5bn of cash — a **US$3.0bn** shortfall on one maturity |
| Bank terms | Tranche A US$4.4bn, 5 years, 1% · Tranche B US$10.0bn, 8 years, 1% cash + 1.5% PIK, 60% guaranteed (stated) |
| Cash interest saved by the coupon cut | **US$0.65bn a year** |
| Net realisable assets / bank principal, 2010 forced-sale values | **0.93×** with DP World · **0.47×** without it |
| Tranche A · Tranche B recovery, PV at 8% | 72% · 71% |
| **Bank creditors, blended PV recovery** | **71 cents on the dollar** — a 29% economic haircut on a 0% nominal one |
| Value transferred from banks to the company by the new terms | **US$2.7bn** (against the loans' own 2010 market value; more against par) |
| Nakheel trade creditors (40% cash, 60% 10% sukuk) | **98%** — the best-treated class |
| Nakheel banks (5-year extension at 4%) | 84% |
| Government: claims converted · implied equity value, PV basis | US$8.9bn · US$4.3bn → **48% recovery** on 2010 values |
| Asset value at which the conversion breaks even | US$19.2bn (against US$14.6bn net realisable in 2010) |

## Reading it honestly

* **It was a maturity crisis in a solvent group.** US$3.5bn due in three weeks against half a billion of cash, in a holding company that owned 80% of a global port operator. Every term that followed — bullets at five and eight years, 1% coupons, an asset-realisation programme — was designed to avoid selling DP World, Jafza and the Istithmar portfolio into the 2010 trough. On forced-sale values the assets covered 93% of the bank principal even with DP World; the plan was a bet that values would recover over eight years, and the tenor was the size of the bet. They did.
* **Nobody took a nominal haircut except the government, and the banks took a real one.** Five to eight years at 1% when comparable paper yielded 8% is worth about 71 cents on the dollar. The model measures the transfer conservatively — against what the loans were already worth at 2010 yields, US$2.7bn — rather than against par, where it would be more than US$4bn. Bank regulatory accounting allowed the restructured loans to stay near par because there was no principal forgiveness; that, not valuation, is why the instrument was paper.
* **Trade creditors beat banks despite ranking below them.** Nakheel's contractors received 40% in cash and 60% in a five-year sukuk at 10% that traded close to par: a 98% present-value recovery against the banks' 71%. Contractors could stop work on projects the government needed finished; lenders could only sue. Operational leverage over legal seniority is the recurring lesson of Gulf restructurings, and this is its clearest case.
* **The government's equity was the only real risk capital.** On 2010 values the US$8.9bn converted bought equity worth perhaps US$4.3bn, and it broke even only if group assets reached US$19bn. The conversion made sense because the government's objective was not recovery but survival of the group and of Dubai's name as a borrower. The counterfactual on the Debt-for-Equity sheet — banks taking 30% in equity — would have improved their economics on paper; they declined it because they could not hold Dubai holding-company equity at par and could hold a guaranteed 1% loan at par.
* **The sensitivity grid is the argument between economists and accountants.** At any discount rate between 6% and 10% the banks recovered 65–78 cents whatever the guarantee cover. Only a discount rate near the 1% coupon itself returns the loans to par — which is what a hold-to-maturity book effectively assumed.
* **What is stated rather than sourced.** The guarantee share of Tranche B (lenders chose among three options; 60% is an assumption), the 2010 discount rates, the asset values and haircuts, and the split of the Abu Dhabi support. The announced terms — amounts, tenors, coupons, PIK options, the government conversion, Nakheel's 40/60 split and 10% sukuk — are as publicly reported. A real model would be built from the creditor information memorandum.

## What I would do with more time

* Rebuild the asset-realisation schedule year by year against the Tranche B maturity, with DP World valued on its own trading multiple each year, to show when the plan became fully covered.
* Add the three Tranche B options as separate instruments with their actual take-up, and the AED-denominated variants, from the creditor circular.
* Model Nakheel separately as a developer: project-level cash flows, the AED 30bn of government support, and the 2011 sukuk's trading history as the market's verdict on the trade-creditor terms.
* Compare against the later Dubai restructurings — Dubai Holding, Drydocks World, Limitless — to test whether the ordering of recoveries by class held.

## Sources

Dubai World and Government of Dubai announcements of 20 May 2010 (proposed terms) and 10 September 2010 (approval by creditors representing over 99% of claims); Nakheel announcements of 2010–2011 including the AED 4.8bn sukuk of August 2011; the Abu Dhabi support announcement of 14 December 2009; contemporaneous reporting of tranche amounts, tenors, coupons and PIK options. Discount rates, asset values, haircuts and the guaranteed share are stated inputs on the Inputs sheet. The model is a reconstruction for learning from public information and is not a statement about any creditor's actual position or accounting.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
