# Project 3 · Tata Consultancy Services — Trading Comparables

**Files:** `TCS_Trading_Comps.xlsx` · `TCS_Trading_Comps.pdf` · `build_model.py` (pulls prices, LTM statements and consensus for the peer set and writes every formula)

**As of:** prices 12 Sep 2026; LTM to the June 2026 quarter; USDINR 95.54.

## What it is

A comparable-companies page for TCS against eight peers in two tiers: Indian large caps (Infosys, HCLTech, Wipro), Indian mid caps (Tech Mahindra, Persistent) and global services names (Accenture, Cognizant, EPAM). Enterprise value built from market cap, debt, minorities and cash; LTM EV/Revenue, EV/EBITDA, EV/EBIT and P/E; forward revenue and EPS calendarised to TCS's March-2027 year; peer statistics; and the implied TCS price range from the interquartile range of each multiple.

## The three things a comps page actually tests

1. **Peer selection, written down.** The criteria, the inclusions and the exclusions (LTIMindtree and Capgemini for data availability, Mphasis and Coforge for size, IBM for mix) are on the Inputs sheet. Anyone can compute a median; the judgement is the list.
2. **Currency and reporting basis.** Infosys files IFRS statements in US dollars while its shares trade in rupees; a naïve pull gives it an EV/Revenue of 200×. The workbook detects the statements currency and converts every component of EV and every metric at its own rate. This is the sort of error that gets a first-year analyst's page thrown back.
3. **Calendarisation.** Consensus is quoted on each company's fiscal year. Accenture's year ends in August, Cognizant's and EPAM's in December, the Indian names' in March. The `Calendarisation` sheet weights each peer's FY0 and FY1 estimates by month overlap with April 2026 – March 2027 (Accenture 0.42 / 0.58, Cognizant 0.75 / 0.25, Indian peers 1.0 / 0.0) so the forward multiples compare like with like.

## Result

| | TCS | Peer median | TCS premium |
|---|---|---|---|
| EV / LTM revenue | 2.8× | 1.8× | +54% |
| EV / LTM EBITDA | 10.4× | 8.7× | +18% |
| EV / LTM EBIT | 11.2× | 10.4× | +7% |
| P / E (LTM) | 15.8× | 15.3× | +3% |
| EV / CY2027 revenue | 2.6× | 1.8× | +47% |
| P / E (CY2027) | 14.2× | 13.4× | +6% |
| EBITDA margin | 27% | 19% | |

Implied TCS price from the peer interquartile range: **INR 1,215 – 2,942**, median across methods **INR 1,848**, against a share price of **INR 2,201**. TCS sits above the peer median on every multiple, at the top of the interquartile range on EBITDA and EBIT, which is where it has traded for most of the last decade: the margin premium is visible in the table (27% EBITDA margin versus 19% for the median peer) and the market pays for it.

Two things stand out in the 2026 numbers. The whole sector has de-rated: 10–14× forward earnings for companies that traded at 25–30× three years ago, reflecting the AI-disruption discount on outsourcing. And Persistent at 29× EBITDA is the growth outlier (16% revenue growth versus 4–6% for the rest); it belongs in the table as context but drags the mean, which is why the median is the statistic used.

## Reading it

The implied range is a market anchor, not an intrinsic value: it says what buyers pay today for businesses like TCS. It is one bar on the football field (Project 5). A premium to peers is neither good nor bad in itself; the question is whether it is larger or smaller than TCS's historical premium, which is the next chart a banker would add.

## What I would do with more time

* Historical premium chart: TCS forward P/E relative to the Nifty IT index over ten years.
* Add LTIMindtree, Capgemini and Globant from a data source that carries them.
* Adjust EBITDA for IFRS 16 leases and stock compensation consistently across US and Indian filers.
* Regress EV/EBITDA on margin and growth across the set to show how much of TCS's premium the fundamentals explain.

## Sources

Yahoo Finance for prices, shares, quarterly statements and consensus revenue and EPS (12 Sep 2026); company annual reports for fiscal-year ends and reporting currency.
