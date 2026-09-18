# Project 15 · Salik — Company Profile and Buyer Screen

**Files:** `Salik_Profile_Buyer_Screen.xlsx` (live model, 152 formulas, zero errors) · `Salik_Profile_Buyer_Screen.pdf` · `build_model.py`

**As of:** FY2022–FY2025 accounts · DFM price and peer multiples 18 Sep 2026 · AED millions.

## What it is

The document an M&A team prepares before deciding whether there is a transaction to pitch. **Profile** sets out what Salik is — ten toll gates under a 49-year concession to 2071, a tariff set by decree, no capex, all profit paid out, 75.1% held by the Government of Dubai. **Financials** shows four years since the September 2022 IPO. **Trading** puts the shares against a short, stated set of listed toll roads. **Buyer Screen** scores twelve candidates on five criteria with live weights and a live rank. **Ability to Pay** works out what a financial sponsor and a strategic operator could each pay per share and why the two numbers differ. This is groundwork for Project 16, the pitch book.

## Result

| | |
|---|---|
| FY2025 revenue · EBITDA · margin | AED 3,097m · AED 2,165m · 70% |
| Revenue growth FY2025 | 35% — the variable tariff from 31 January 2025 and two new gates |
| Free cash flow / EBITDA | 96% — capex is AED 3m |
| Net debt / EBITDA | 1.6× (2.2× at the IPO), on AED 4.0bn of unchanged debt |
| Market cap · EV | AED 40.3bn · AED 43.8bn |
| **EV / EBITDA · P/E · yield** | **20.2×** · 25.9× · 3.5% |
| Price return since the IPO at AED 2.00 | +169% |
| Premium to the included peer median (15.5×) | 30% |
| **Buyer screen, top five** | **Vinci Concessions 4.30** · Brookfield 4.20 · KKR 4.20 · Abertis 4.10 · Dubai Holding/ICD 4.10 |
| Sponsor ability to pay (6× debt, 18× exit, 13% IRR) | **AED 4.66** — 13% below the market |
| Strategic ability to pay (7% WACC, 45-year concession) | **AED 6.05** — 13% above the market, 22.6× EBITDA |

## Reading it honestly

* **Salik is a regulated utility that happens to be called a toll road.** The tariff, the gate count and the concession are all government decisions; management's levers are collection efficiency and ancillary revenue. That is why the 2025 step-up in revenue was a policy event, and why a buyer's diligence is about the relationship with the RTA rather than traffic forecasts.
* **The screen's finding is who cannot buy, not who can.** The purest operators — Transurban, Ferrovial — score highest on strategic fit and drop down the list because a controlling stake is not on offer and their balance sheets are stretched. The candidates at the top have already done minority infrastructure deals with Gulf sovereigns: Brookfield, KKR, GIP, the pattern set by the ADNOC pipelines and Jafurah transactions. Vinci leads because it combines operator credibility with the deepest pockets. The realistic transaction is a government sell-down to a sponsor or sovereign anchor, with an operator as a technology partner rather than a buyer.
* **The scores are judgements and are labelled as such.** Twelve buyers, five criteria, 1 to 5, with the weights as inputs. Changing the weight on government consent from 20% to 30% moves the sovereigns to the top. The sheet is built so that the argument about the ranking happens in the weights, where it can be seen, rather than in the prose.
* **The sponsor cannot pay the market price and the strategic can.** With 6× leverage, an exit at 18× and a 13% hurdle, a fund's maximum is AED 4.66 a share, 13% below where the stock trades. A strategic discounting 45 years of concession cash flow at 7% can justify AED 6.05. The gap is structural — a five-year exit against a 45-year hold — and it is why financial buyers anchor Gulf infrastructure deals while strategics set the price. The thirteenth check fails the workbook if that ordering ever reverses, because it would mean the leverage or the exit assumption had been pushed past what the asset supports.
* **The market already prices Salik as a strategic asset.** At 20× EBITDA the shares sit 30% above the included peer median and above what any sponsor could pay. The premium is the concession length, the absence of traffic risk and the government's ownership. It also means there is no obvious arbitrage for a buyer, only a partnership case.
* **What is stated rather than sourced.** The concession fee to the RTA, the exact gate count and the 2071 end date are from the IPO prospectus and public announcements rather than a filing I could reproduce here; the peer multiples are a single-day reading; the buyers' regional presence and precedent are from public deal records. Nothing in the screen suggests any party is or has been in discussions.

## What I would do with more time

* Add traffic and revenue-per-crossing data by gate from the quarterly results, which would let the ability-to-pay case distinguish population growth from tariff and gate effects.
* Build the sponsor case as a full Project-7-style LBO with a cash sweep and covenant tests, rather than the year-by-year paydown table here.
* Add a precedent-transaction sheet for minority stakes in Gulf infrastructure — ADNOC pipelines, Aramco pipelines, Jafurah, Abu Dhabi ports — with the implied multiples, so the strategic and sponsor numbers can be checked against what has actually been paid.
* Extend the screen with a contact map: who at each candidate has done the relevant deal, which is what the pitch book in Project 16 needs.

## Sources

Salik Company PJSC financial statements FY2022–FY2025 via Yahoo Finance for SALIK.AE. Concession terms, gate count, tariff and IPO details from the 2022 prospectus and Roads and Transport Authority announcements. Share price, 52-week range and peer multiples from Yahoo Finance and the DFM close on 18 September 2026. Buyer scores are the author's judgements from public precedent; criteria weights and every ability-to-pay input are stated on the relevant sheet.

All source data is embedded in `build_model.py` with its provenance, so the workbook rebuilds from the repository alone.
