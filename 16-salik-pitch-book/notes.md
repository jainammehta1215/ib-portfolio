# Project 16 · Salik — Sell-Side Pitch Book

**Files:** `Salik_Pitch_Book.pdf` (12 slides, landscape) · `build_deck.py` · `price_history.json`

**As of:** every figure is read from the Project 15 workbook after recalculation · market data 18 Sep 2026.

## What it is

A pitch book for a hypothetical mandate: advising the Government of Dubai on strategic alternatives for its 75.1% stake in Salik. It is the deck a coverage team would take to the Department of Finance — executive summary, situation overview, financial performance, share-price history, trading comparables, a valuation summary in football-field form, four strategic alternatives, the buyer universe, a recommended path with indicative terms, a sixteen-week process, risks and mitigants, and an appendix on sources and basis of preparation.

Unlike the other projects this is a deck rather than a model. That is deliberate: the model already exists in Project 15, and the discipline here is that the deck reads every number from it — the financials, the multiples, the sponsor and strategic ability to pay, the buyer scores and ranks — so the two cannot drift apart. Charts are drawn with matplotlib; pages are laid out with reportlab in Salik's palette from `ibkit/style.py`.

## The argument

| | |
|---|---|
| Codename | Project Gateway |
| Situation | AED 40.3bn market cap, 20.2× EBITDA, +169% since the IPO; government holds 75.1% of a fully valued, scarce asset |
| Constraint | The RTA sets the tariff and holds the concession; a change of control is neither available nor advisable |
| Valuation | Sponsor ceiling AED 4.66 (below market); strategic ceiling AED 6.05; placement range AED 5.64–6.01 (5–12% premium) |
| Alternatives | Status quo · accelerated bookbuild at a discount · **anchor-stake placement** · operator partnership |
| Recommendation | Place 10–15% with a global-sponsor-plus-sovereign consortium at a 5–12% premium, raising AED 4.2–7.2bn; government stays above 60%; operator partnership as a second step |
| Buyers | Vinci Concessions, Brookfield, KKR, Abertis, the incumbent group; Transurban and Ferrovial as partners, not buyers |
| Process | Sixteen weeks: preparation, approach, selection, execution |

## Reading it honestly

* **The deck argues for a minority deal because that is what the model supports.** Project 15 shows that no financial sponsor can reach the market price and that the buyers who score highest on consent are sovereign and sponsor money, not operators. A pitch that recommended a sale of control would contradict its own appendix.
* **The premium is the pitch's one real claim.** A 5–12% premium over market for a 10–15% block is defended on scarcity, governance rights and a three-year lock-up, and it sits inside the strategic ceiling. It is an argument, not a fact; the football field shows exactly how much of the strategic ceiling it consumes.
* **Everything a banker would normally add from a data terminal is absent or stated.** No precedent minority-stake multiples for Gulf infrastructure, no analyst target prices, no traffic data by gate. The appendix says so. A real pitch would have all three, and Project 15's notes list them as the next work.
* **Indicative terms are proposals, not soundings.** Stake size, premium, lock-up and board rights are the author's suggestions for discussion. Nothing in the deck reflects contact with any party.
* **The deck is short by design.** Twelve slides is what a first meeting gets. The Project 15 workbook is the appendix a reader who wants the numbers opens next.

## What I would do with more time

* Add a precedent-transactions slide for minority stakes in Gulf infrastructure (ADNOC and Aramco pipelines, Jafurah, Abu Dhabi ports) with the multiples paid, which is the strongest support for the premium.
* Add analyst consensus and a sum-of-the-parts by gate once traffic disclosure allows it.
* Build the buyer pages out with contact maps and each candidate's recent Gulf activity.
* Produce the same deck as a PowerPoint file for a client that wants to edit it; the PDF is the reviewable version.

## Sources

All numbers from `15-salik-profile/Salik_Profile_Buyer_Screen.xlsx` (read with `openpyxl` after recalculation). Monthly share prices from the DFM via Yahoo Finance, stored in `price_history.json` so the deck rebuilds offline. Concession and IPO facts as in Project 15. Indicative terms are stated proposals.

**This is a hypothetical mandate prepared as a portfolio exercise on public information.** Neither the Government of Dubai nor Salik has, to the author's knowledge, engaged any adviser on such a transaction, and no party named in the buyer universe is or has been in discussions.
