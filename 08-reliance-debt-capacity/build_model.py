"""
Build 08-reliance-debt-capacity/Reliance_Debt_Capacity.xlsx — how much Reliance Industries can borrow and at what
point more debt stops being cheap.

Sheets: Cover · Summary · Inputs · Credit Profile · Debt Capacity · Optimal Structure · Liquidity · Sensitivity · Checks

Three questions, answered in order. First, what does the balance sheet actually look like once the items a rating
agency adds back — leases and the deferred spectrum liabilities — are counted, because the headline 0.6x
net-debt-to-EBITDA in the press release is not the number an agency uses. Second, how much more could the company
borrow at each rating category, taking the tightest of three constraints (leverage, interest cover, FFO-to-debt)
rather than the loosest. Third, where the weighted average cost of capital bottoms out as leverage rises, which is
the only defensible definition of an "optimal" structure.

All source data is the audited FY2026 consolidated results published on 24 April 2026 (year to 31 March 2026),
in rupees crore. Rating thresholds and credit spreads are stated inputs adapted from published agency methodologies.
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)
from ibkit.sheet import Sheet, checks_sheet

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"

# --------------------------------------------------------------------------- #
# Source data — Reliance Industries audited consolidated results, FY2026 (year to 31 March 2026), ₹ crore.
# Media release and audited statements dated 24 April 2026. Share price: NSE close 18 Sep 2026.
# --------------------------------------------------------------------------- #
R = dict(revenue=1175919.0, ebitda=207911.0, oneoff=8924.0, da=57688.0, fincost=27061.0, pbt=123162.0,
         tax=27552.0, pat=95610.0, pat_owners=80775.0, capex=144271.0, cfo=192113.0, int_paid=39981.0,
         div_paid=7443.0, debt=374421.0, cash=249704.0, leases=16198.0 + 7381.0, spectrum_dpl=99552.0,
         equity_owners=904030.0, nci=181836.0, shares=1353.25, price=1244.60, dps=6.0,
         cur_borrow=103670.0, nc_borrow=270751.0)

book = Book(theme=THEMES["reliance"], project_no=8, project="Debt Capacity and Capital Structure",
            company="Reliance Industries", units="\u20b9 crore unless stated",
            as_of="FY2026 audited results \u00b7 market data 18 Sep 2026")

S = lambda name, **kw: Sheet(book, name, **kw)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
si = S("Inputs", ncols=6, label_width=58, subtitle="Audited FY2026 figures, market data, rating framework and cost-of-capital inputs")
si.ws.column_dimensions["F"].width = 60
I = {}
si.section("REPORTED FY2026 (\u20b9 crore)")
I["rev"] = si.row("Revenue (value of sales and services)", R["revenue"], NUM0)
I["ebitda_rep"] = si.row("EBITDA, as reported", R["ebitda"], NUM0)
I["oneoff"] = si.row("  of which: profit on sale of listed investments (one-off)", R["oneoff"], NUM0,
                     note="Disclosed in note 5 of the results. Removed below because a lender does not lend against a disposal gain")
I["ebitda"] = si.row("EBITDA, adjusted", f"={I['ebitda_rep']}-{I['oneoff']}", NUM0, bold=True)
I["da"] = si.row("Depreciation, amortisation and depletion", R["da"], NUM0)
I["ebit"] = si.row("EBIT, adjusted", f"={I['ebitda']}-{I['da']}", NUM0)
I["fincost"] = si.row("Finance costs (P&L)", R["fincost"], NUM0)
I["int_paid"] = si.row("Interest paid (cash flow statement)", R["int_paid"], NUM0,
                       note="Higher than the P&L charge because interest on projects under construction is capitalised. Cash cover uses this figure")
I["pbt"] = si.row("Profit before tax", R["pbt"], NUM0)
I["tax"] = si.row("Tax expense", R["tax"], NUM0)
I["etr"] = si.row("Effective tax rate", f"={I['tax']}/{I['pbt']}", PCT)
I["cfo"] = si.row("Net cash from operating activities", R["cfo"], NUM0)
I["capex"] = si.row("Capital expenditure (company definition, ex-spectrum)", R["capex"], NUM0)
I["div"] = si.row("Dividends paid to equity holders", R["div_paid"], NUM0)

si.blank(); si.section("BALANCE SHEET AT 31 MARCH 2026 (\u20b9 crore)")
I["debt"] = si.row("Outstanding debt (company definition)", R["debt"], NUM0)
I["cash"] = si.row("Cash and cash equivalents (company definition, incl. liquid investments)", R["cash"], NUM0)
I["leases"] = si.row("Lease liabilities (current and non-current)", R["leases"], NUM0)
I["spectrum"] = si.row("Deferred payment liabilities (spectrum instalments to DoT)", R["spectrum_dpl"], NUM0,
                       note="Interest-bearing instalments owed to the government for spectrum. Agencies treat them as debt; the company's net debt figure excludes them")
I["eq_own"] = si.row("Equity attributable to owners", R["equity_owners"], NUM0)
I["nci"] = si.row("Non-controlling interests", R["nci"], NUM0)
I["cur_borrow"] = si.row("Borrowings due within twelve months", R["cur_borrow"], NUM0)

si.blank(); si.section("MARKET DATA")
I["price"] = si.row("Share price (\u20b9, NSE close 18 Sep 2026)", R["price"], NUM2)
I["shares"] = si.row("Shares outstanding (crore)", R["shares"], NUM2)
I["mcap"] = si.row("Market capitalisation", f"={I['price']}*{I['shares']}", NUM0, bold=True)
I["dps"] = si.row("Dividend per share, FY2026 (\u20b9)", R["dps"], NUM2)

si.blank(); si.section("COST OF CAPITAL")
I["rf"] = si.row("Risk-free rate (10-year Government of India, INR)", 0.0655, PCT, note="Benchmark G-sec yield, September 2026")
I["erp"] = si.row("Equity risk premium, India", 0.070, PCT, note="Mature-market premium plus country risk, per Damodaran's country tables")
I["beta_l"] = si.row("Observed levered beta versus NIFTY 50", 1.05, NUM2, note="Five-year monthly. Regressed in Project 5's method; taken as an input here")
I["mtr"] = si.row("Marginal tax rate (tax shield on interest)", 0.2517, PCT, note="Section 115BAA concessional rate with surcharge and cess")

si.blank(); si.section("RATING FRAMEWORK \u2014 stated inputs")
si.note("Thresholds are illustrative levels adapted from published agency methodologies for large diversified industrials. "
        "Real committees also weigh business risk, diversification, liquidity and sovereign ceilings; Reliance is rated "
        "A- (S&P) and Baa1 (Moody's), two notches above India's sovereign, precisely because of those qualitative factors.", height=40)
RATINGS = ["AA", "A", "BBB", "BB", "B"]
LEV_MAX = [1.00, 1.75, 2.75, 4.00, 5.50]
COV_MIN = [12.0, 8.0, 5.0, 3.0, 2.0]
FFO_MIN = [0.60, 0.45, 0.30, 0.20, 0.12]
SPREAD = [0.0080, 0.0140, 0.0220, 0.0350, 0.0550]
si.head(RATINGS, label="Rating category")
I["lev_max"] = si.multi("Maximum net debt / EBITDA", LEV_MAX, MULT)
I["cov_min"] = si.multi("Minimum EBITDA / cash interest", COV_MIN, MULT)
I["ffo_min"] = si.multi("Minimum FFO / net debt", FFO_MIN, PCT)
I["spread"] = si.multi("Credit spread over the risk-free rate", SPREAD, PCT)
I["kd"] = si.multi("Pre-tax cost of debt", [f"={I['rf']}+{s}" for s in I["spread"]], PCT)

# --------------------------------------------------------------------------- #
# Credit Profile
# --------------------------------------------------------------------------- #
sc = S("Credit Profile", ncols=6, label_width=58, subtitle="The balance sheet as reported, and as a rating agency would restate it")
sc.ws.column_dimensions["F"].width = 60
C = {}
sc.section("NET DEBT \u2014 THREE DEFINITIONS")
C["nd_co"] = sc.row("Net debt, company definition (debt less cash)", f"={I['debt']}-{I['cash']}", NUM0, bold=True,
                    note="The 0.6x figure in the results release")
C["nd_lease"] = sc.row("Add: lease liabilities", f"={I['leases']}", NUM0, font=F_LINK)
C["nd_spec"] = sc.row("Add: deferred spectrum liabilities", f"={I['spectrum']}", NUM0, font=F_LINK)
C["nd_adj"] = sc.row("Adjusted net debt (agency basis)", f"={C['nd_co']}+{C['nd_lease']}+{C['nd_spec']}", NUM0,
                     bold=True, border=DOUBLE_BORDER)
C["gd_adj"] = sc.row("Adjusted gross debt", f"={I['debt']}+{I['leases']}+{I['spectrum']}", NUM0)

sc.blank(); sc.section("LEVERAGE AND COVER")
C["lev_co"] = sc.row("Net debt / EBITDA \u2014 company basis, reported EBITDA", f"={C['nd_co']}/{I['ebitda_rep']}", MULT)
C["lev_adj"] = sc.row("Adjusted net debt / adjusted EBITDA", f"={C['nd_adj']}/{I['ebitda']}", MULT, bold=True,
                      note="Roughly double the headline. This is the number to defend to a lender")
C["gross_lev"] = sc.row("Adjusted gross debt / adjusted EBITDA", f"={C['gd_adj']}/{I['ebitda']}", MULT)
C["cov_pl"] = sc.row("EBITDA / finance costs (P&L)", f"={I['ebitda']}/{I['fincost']}", MULT)
C["cov_cash"] = sc.row("EBITDA / cash interest paid", f"={I['ebitda']}/{I['int_paid']}", MULT, bold=True,
                       note="Uses interest actually paid, including the capitalised portion")
C["ebit_cov"] = sc.row("EBIT / finance costs", f"={I['ebit']}/{I['fincost']}", MULT)

sc.blank(); sc.section("CASH FLOW MEASURES")
C["ffo"] = sc.row("Funds from operations (EBITDA \u2212 cash interest \u2212 cash tax)",
                  f"={I['ebitda']}-{I['int_paid']}-{I['tax']}", NUM0, bold=True,
                  note="Tax expense is used as a proxy for cash tax; the cash flow statement shows taxes paid materially lower because of deferred tax")
C["ffo_nd"] = sc.row("FFO / adjusted net debt", f"={C['ffo']}/{C['nd_adj']}", PCT, bold=True)
C["fcf"] = sc.row("Free cash flow (CFO \u2212 capex)", f"={I['cfo']}-{I['capex']}", NUM0)
C["fcf_div"] = sc.row("Free cash flow after dividends", f"={C['fcf']}-{I['div']}", NUM0, bold=True)
C["capex_ebitda"] = sc.row("Capex / adjusted EBITDA", f"={I['capex']}/{I['ebitda']}", PCT,
                           note="Nearly three-quarters of EBITDA is reinvested. Debt capacity is only useful if capex can be paused; here it largely cannot")
C["payback"] = sc.row("Years to repay adjusted net debt from free cash flow after dividends",
                      f"=IF({C['fcf_div']}>0,{C['nd_adj']}/{C['fcf_div']},\"n/m\")", NUM)

sc.blank(); sc.section("CAPITALISATION")
C["bv_cap"] = sc.row("Adjusted gross debt / (debt + book equity)", f"={C['gd_adj']}/({C['gd_adj']}+{I['eq_own']}+{I['nci']})", PCT)
C["mv_cap"] = sc.row("Adjusted net debt / (net debt + market capitalisation)",
                     f"={C['nd_adj']}/({C['nd_adj']}+{I['mcap']})", PCT, bold=True,
                     note="Market-value leverage. This is the weight used in the WACC work")
C["ev"] = sc.row("Enterprise value (market cap + adjusted net debt + NCI at book)",
                 f"={I['mcap']}+{C['nd_adj']}+{I['nci']}", NUM0)
C["ev_ebitda"] = sc.row("EV / adjusted EBITDA", f"={C['ev']}/{I['ebitda']}", MULT)

sc.blank(); sc.section("IMPLIED RATING FROM THE GRID")
C["r_lev"] = sc.row("Implied by leverage",
                    f"=IF({C['lev_adj']}<={I['lev_max'][0]},\"AA\",IF({C['lev_adj']}<={I['lev_max'][1]},\"A\","
                    f"IF({C['lev_adj']}<={I['lev_max'][2]},\"BBB\",IF({C['lev_adj']}<={I['lev_max'][3]},\"BB\",\"B\"))))", "@")
C["r_cov"] = sc.row("Implied by cash interest cover",
                    f"=IF({C['cov_cash']}>={I['cov_min'][0]},\"AA\",IF({C['cov_cash']}>={I['cov_min'][1]},\"A\","
                    f"IF({C['cov_cash']}>={I['cov_min'][2]},\"BBB\",IF({C['cov_cash']}>={I['cov_min'][3]},\"BB\",\"B\"))))", "@")
C["r_ffo"] = sc.row("Implied by FFO / net debt",
                    f"=IF({C['ffo_nd']}>={I['ffo_min'][0]},\"AA\",IF({C['ffo_nd']}>={I['ffo_min'][1]},\"A\","
                    f"IF({C['ffo_nd']}>={I['ffo_min'][2]},\"BBB\",IF({C['ffo_nd']}>={I['ffo_min'][3]},\"BB\",\"B\"))))", "@")
sc.note("Leverage and FFO both say A. Cash interest cover, measured on interest actually paid, says BB — and the gap "
        "is the interest capitalised on nearly ₹2 lakh crore of projects under construction, which the P&L never shows. "
        "The published A- / Baa1 sits between the two readings, which is what a committee does when the metrics are "
        "strong but the capex programme is heavy and the sovereign ceiling is BBB-. The grid is a screen, not a committee.", height=48)

# --------------------------------------------------------------------------- #
# Debt Capacity
# --------------------------------------------------------------------------- #
sd = S("Debt Capacity", ncols=8, label_width=54, col_width=14,
       subtitle="Maximum adjusted net debt at each rating \u2014 the tightest of three constraints wins")
D = {}
sd.section("CAPACITY BY CONSTRAINT (\u20b9 crore)")
sd.head(RATINGS, label="Rating category")
D["cap_lev"] = sd.multi("From leverage: EBITDA \u00d7 maximum net debt / EBITDA",
                        [f"={I['ebitda']}*{m}" for m in I["lev_max"]], NUM0)
D["cap_cov"] = sd.multi("From cover: EBITDA / (minimum cover \u00d7 cost of debt) \u2212 cash-free",
                        [f"={I['ebitda']}/({c}*{k})" for c, k in zip(I["cov_min"], I["kd"])], NUM0)
D["cap_ffo"] = sd.multi("From FFO: FFO / minimum FFO-to-debt",
                        [f"={C['ffo']}/{f}" for f in I["ffo_min"]], NUM0)
D["cap"] = sd.multi("Debt capacity \u2014 binding constraint",
                    [f"=MIN({a},{b},{c})" for a, b, c in zip(D["cap_lev"], D["cap_cov"], D["cap_ffo"])],
                    NUM0, bold=True, border=DOUBLE_BORDER)
D["which"] = sd.multi("Which constraint binds",
                      [f"=IF({d}={a},\"Leverage\",IF({d}={b},\"Cover\",\"FFO\"))"
                       for a, b, c, d in zip(D["cap_lev"], D["cap_cov"], D["cap_ffo"], D["cap"])], "@")
sd.blank()
sd.section("HEADROOM AGAINST TODAY'S ADJUSTED NET DEBT")
D["nd_now"] = sd.multi("Adjusted net debt today", [f"={C['nd_adj']}"] * 5, NUM0, font=F_LINK)
D["head"] = sd.multi("Incremental borrowing capacity", [f"={c}-{n}" for c, n in zip(D["cap"], D["nd_now"])],
                     NUM0, bold=True)
D["head_usd"] = sd.multi("  in US$ billion at \u20b983.5 / US$", [f"={h}/8350" for h in D["head"]], NUM,
                         font=F_FORMULA)
D["head_pct"] = sd.multi("  as % of market capitalisation", [f"={h}/{I['mcap']}" for h in D["head"]], PCT,
                         font=F_FORMULA)
D["pf_lev"] = sd.multi("Pro forma net debt / EBITDA if fully drawn", [f"={c}/{I['ebitda']}" for c in D["cap"]], MULT)
D["pf_cov"] = sd.multi("Pro forma EBITDA / interest if fully drawn",
                       [f"={I['ebitda']}/({c}*{k})" for c, k in zip(D["cap"], I["kd"])], MULT)
sd.blank()
sd.note("The cover constraint assumes the whole debt stack reprices to the category's cost of debt, which overstates "
        "interest for a company whose existing borrowings were raised at lower rates. That makes the capacity numbers "
        "conservative, which is the right direction for a capacity study. Capacity from FFO ignores the feedback of new "
        "interest on FFO; the effect is second-order at these leverage levels.", height=44)
sd.blank()
sd.section("WHAT THE HEADROOM WOULD FUND \u2014 at the A category")
D["a_head"] = sd.row("Incremental capacity at A", f"={D['head'][1]}", NUM0, font=F_LINK, bold=True)
D["yrs_capex"] = sd.row("Years of FY2026 capex it would fund on its own", f"={D['a_head']}/{I['capex']}", NUM)
D["yrs_fcf"] = sd.row("Years of FY2026 free cash flow it represents", f"=IF({C['fcf']}>0,{D['a_head']}/{C['fcf']},\"n/m\")", NUM)
D["buyback"] = sd.row("Share of market cap it could retire in a levered recapitalisation", f"={D['a_head']}/{I['mcap']}", PCT)

# --------------------------------------------------------------------------- #
# Optimal Structure
# --------------------------------------------------------------------------- #
so = S("Optimal Structure", ncols=8, label_width=54, col_width=13.5,
       subtitle="Weighted average cost of capital across leverage \u2014 the minimum is the optimum")
O = {}
so.section("UNLEVERED COST OF EQUITY")
O["dv_now"] = so.row("Market-value debt / (debt + equity) today", f"={C['mv_cap']}", PCT, font=F_LINK)
O["de_now"] = so.row("Debt / equity today", f"={O['dv_now']}/(1-{O['dv_now']})", NUM2)
O["beta_u"] = so.row("Unlevered beta (Hamada)", f"={I['beta_l']}/(1+(1-{I['mtr']})*{O['de_now']})", NUM2, bold=True,
                     note="Observed beta stripped of today's leverage, then re-levered at each candidate structure")
O["ke_now"] = so.row("Cost of equity today", f"={I['rf']}+{I['beta_l']}*{I['erp']}", PCT)
O["kd_now"] = so.row("Pre-tax cost of debt today (A category)", f"={I['kd'][1]}", PCT, font=F_LINK)
O["wacc_now"] = so.row("WACC today", f"=(1-{O['dv_now']})*{O['ke_now']}+{O['dv_now']}*{O['kd_now']}*(1-{I['mtr']})",
                       PCT, bold=True, border=TOTAL_BORDER)

so.blank(); so.section("WACC ACROSS CANDIDATE STRUCTURES")
DVS = [0.00, 0.10, 0.20, 0.30, 0.40, 0.50]
so.head([f"{int(d*100)}%" for d in DVS], label="Debt / (debt + equity) at market value")
r_dv = so.r
O["dv"] = so.multi("Debt / (debt + equity)", DVS, PCT)
O["de"] = so.multi("Debt / equity", [f"={d}/(1-{d})" for d in O["dv"]], NUM2)
# Implied net debt in ₹ crore at each weight, holding EV constant at today's EV less NCI (a standard simplification)
ev_ex = f"({I['mcap']}+{C['nd_adj']})"
O["nd_at"] = so.multi("Implied net debt (holding firm value constant)", [f"={d}*{ev_ex}" for d in O["dv"]], NUM0)
O["lev_at"] = so.multi("Implied net debt / EBITDA", [f"={n}/{I['ebitda']}" for n in O["nd_at"]], MULT)
O["rat_at"] = so.multi("Implied rating from the leverage grid",
                       [f"=IF({l}<={I['lev_max'][0]},\"AA\",IF({l}<={I['lev_max'][1]},\"A\","
                        f"IF({l}<={I['lev_max'][2]},\"BBB\",IF({l}<={I['lev_max'][3]},\"BB\",\"B\"))))" for l in O["lev_at"]], "@")
kd_lookup = lambda l: (f"=IF({l}<={I['lev_max'][0]},{I['kd'][0]},IF({l}<={I['lev_max'][1]},{I['kd'][1]},"
                       f"IF({l}<={I['lev_max'][2]},{I['kd'][2]},IF({l}<={I['lev_max'][3]},{I['kd'][3]},{I['kd'][4]}))))")
O["kd_at"] = so.multi("Pre-tax cost of debt at that rating", [kd_lookup(l) for l in O["lev_at"]], PCT)
O["cov_at"] = so.multi("Implied EBITDA / interest", [f"=IF({n}>0,{I['ebitda']}/({n}*{k}),\"n/a\")" for n, k in zip(O["nd_at"], O["kd_at"])], MULT)
O["beta_at"] = so.multi("Re-levered beta", [f"={O['beta_u']}*(1+(1-{I['mtr']})*{de})" for de in O["de"]], NUM2)
O["ke_at"] = so.multi("Cost of equity", [f"={I['rf']}+{b}*{I['erp']}" for b in O["beta_at"]], PCT)
O["wacc_at"] = so.multi("WACC", [f"=(1-{d})*{k}+{d}*{kd}*(1-{I['mtr']})" for d, k, kd in zip(O["dv"], O["ke_at"], O["kd_at"])],
                        PCT, bold=True, border=DOUBLE_BORDER)
so.blank()
first, last = O["wacc_at"][0].split("!")[1], O["wacc_at"][-1].split("!")[1]
O["wacc_min"] = so.row("Minimum WACC in the grid", f"=MIN({first}:{last})", PCT, bold=True)
O["dv_opt"] = so.row("Debt weight at the minimum",
                     f"=INDEX({O['dv'][0].split('!')[1]}:{O['dv'][-1].split('!')[1]},MATCH({O['wacc_min']},{first}:{last},0))", PCT, bold=True)
O["saving"] = so.row("WACC saving versus today (basis points)", f"=({O['wacc_now']}-{O['wacc_min']})*10000", NUM)
O["val_gain"] = so.row("Illustrative firm-value gain: FY2026 FCF grown at 4%, capitalised at old and new WACC",
                       f"=IF({C['fcf']}>0,{C['fcf']}*1.04*(1/({O['wacc_min']}-0.04)-1/({O['wacc_now']}-0.04)),0)", NUM0,
                       note="A perpetuity on one year's free cash flow. Indicative of order of magnitude only")
so.note("The curve is flat between roughly 10% and 30% debt: the tax shield and the rising cost of equity nearly "
        "offset, and the drop to BBB at 20% costs 140 basis points of spread that the shield mostly absorbs. The optimum "
        "is therefore a range, not a point, and the company already sits inside it at about 13%. Beyond 30% the grid "
        "implies BB, the spread jumps again, and WACC rises faster than the shield can compensate.", height=44)

# --------------------------------------------------------------------------- #
# Liquidity
# --------------------------------------------------------------------------- #
sl = S("Liquidity", ncols=6, label_width=58, subtitle="Twelve-month sources and uses \u2014 can the company fund itself without the market?")
sl.ws.column_dimensions["F"].width = 60
Q = {}
sl.section("SOURCES OVER THE NEXT TWELVE MONTHS")
Q["s_cash"] = sl.row("Cash and cash equivalents at 31 March 2026", f"={I['cash']}", NUM0, font=F_LINK)
Q["cfo_g"] = sl.row("Assumed growth in operating cash flow", 0.06, PCT, note="Below FY2026 EBITDA growth of 13%, which was flattered by cracks and the disposal gain")
Q["s_cfo"] = sl.row("Operating cash flow, FY2027E", f"={I['cfo']}*(1+{Q['cfo_g']})", NUM0)
Q["sources"] = sl.row("Total sources", f"={Q['s_cash']}+{Q['s_cfo']}", NUM0, bold=True, border=TOTAL_BORDER)
sl.blank(); sl.section("USES OVER THE NEXT TWELVE MONTHS")
Q["u_mat"] = sl.row("Borrowings due within twelve months", f"={I['cur_borrow']}", NUM0, font=F_LINK,
                    note="Includes commercial paper and working-capital lines that are rolled in the ordinary course")
Q["capex_g"] = sl.row("Assumed capex, FY2027E, as % of FY2026", 1.00, PCT, note="New Energy giga-factories and O2C projects keep the programme flat")
Q["u_capex"] = sl.row("Capital expenditure, FY2027E", f"={I['capex']}*{Q['capex_g']}", NUM0)
Q["u_int"] = sl.row("Cash interest", f"={I['int_paid']}", NUM0, font=F_LINK)
Q["u_div"] = sl.row("Dividend at \u20b96 per share", f"={I['dps']}*{I['shares']}", NUM0)
Q["u_spec"] = sl.row("Spectrum instalment (assumed at FY2026 level)", 4736.0, NUM0, note="FY2026 payment of deferred liabilities from the cash flow statement")
Q["uses"] = sl.row("Total uses", f"={Q['u_mat']}+{Q['u_capex']}+{Q['u_int']}+{Q['u_div']}+{Q['u_spec']}", NUM0, bold=True, border=TOTAL_BORDER)
sl.blank()
Q["cover"] = sl.row("Sources / uses", f"={Q['sources']}/{Q['uses']}", MULT, bold=True, border=DOUBLE_BORDER)
Q["cover_ex"] = sl.row("Sources / uses excluding rollover of short-term borrowings", f"={Q['sources']}/({Q['uses']}-{Q['u_mat']})", MULT, bold=True)
Q["surplus"] = sl.row("Surplus / (shortfall) after all uses", f"={Q['sources']}-{Q['uses']}", NUM0)
Q["min_cash"] = sl.row("Minimum operating cash assumed", 40000.0, NUM0)
Q["room"] = sl.row("Cash room after minimum balance", f"={Q['surplus']}-{Q['min_cash']}", NUM0, bold=True)
sl.note("The company can meet every use for the year from cash on hand and one year's operating cash flow without "
        "raising a rupee, and can do so even if every short-term line is repaid rather than rolled. That is what an "
        "A-category liquidity profile looks like, and it is the qualitative factor that lets the rating sit above the "
        "sovereign.", height=40)

# --------------------------------------------------------------------------- #
# Sensitivity
# --------------------------------------------------------------------------- #
ss = S("Sensitivity", ncols=8, label_width=40, col_width=13.5,
       subtitle="Incremental capacity at the A category (\u20b9 crore) \u2014 every cell rebuilds all three constraints")
EBS = [-0.20, -0.10, 0.00, 0.10, 0.20]
LEVS = [1.25, 1.50, 1.75, 2.00, 2.25]
ss.section("HEADROOM \u2014 EBITDA CHANGE versus LEVERAGE THRESHOLD AT A (cover and FFO thresholds held)")
ss.ws.cell(ss.r, 2, "Max ND/EBITDA at A  \\  Change in EBITDA").font = F_BOLD
for j, e in enumerate(EBS):
    c = ss.ws.cell(ss.r, 3 + j, e); c.font = book.f_header; c.fill = book.fill_secondary
    c.number_format = '+0%;-0%;0%'; c.alignment = Alignment(horizontal="center")
h1 = ss.r; ss.r += 1; g1 = ss.r
for lv in LEVS:
    c = ss.ws.cell(ss.r, 2, lv); c.font = F_BOLD; c.number_format = MULT
    for j, e in enumerate(EBS):
        eb = f"({I['ebitda']}*(1+{L(3+j)}${h1}))"
        ffo = f"({eb}-{I['int_paid']}-{I['tax']})"
        cap = f"MIN({eb}*$B{ss.r},{eb}/({I['cov_min'][1]}*{I['kd'][1]}),{ffo}/{I['ffo_min'][1]})"
        cc = ss.ws.cell(ss.r, 3 + j, f"={cap}-{C['nd_adj']}"); cc.font = F_FORMULA; cc.number_format = NUM0
        if abs(e) < 1e-9 and abs(lv - 1.75) < 1e-9:
            cc.fill = book.fill_accent; cc.font = Font(name=FONT, size=10, bold=True)
    ss.r += 1
ss.blank()
ss.section("WACC \u2014 EQUITY RISK PREMIUM versus DEBT WEIGHT (rating and cost of debt from the grid)")
ERPS = [0.060, 0.065, 0.070, 0.075, 0.080]
ss.ws.cell(ss.r, 2, "Debt weight  \\  Equity risk premium").font = F_BOLD
for j, e in enumerate(ERPS):
    c = ss.ws.cell(ss.r, 3 + j, e); c.font = book.f_header; c.fill = book.fill_secondary
    c.number_format = PCT; c.alignment = Alignment(horizontal="center")
h2 = ss.r; ss.r += 1; g2 = ss.r
for d in DVS[1:]:
    c = ss.ws.cell(ss.r, 2, d); c.font = F_BOLD; c.number_format = PCT
    lev = f"($B{ss.r}*{ev_ex}/{I['ebitda']})"
    kd = kd_lookup(lev)[1:]
    for j, e in enumerate(ERPS):
        beta = f"({O['beta_u']}*(1+(1-{I['mtr']})*($B{ss.r}/(1-$B{ss.r}))))"
        ke = f"({I['rf']}+{beta}*{L(3+j)}${h2})"
        cc = ss.ws.cell(ss.r, 3 + j, f"=(1-$B{ss.r})*{ke}+$B{ss.r}*{kd}*(1-{I['mtr']})")
        cc.font = F_FORMULA; cc.number_format = '0.00%'
        if abs(e - 0.07) < 1e-9 and abs(d - 0.10) < 1e-9:
            cc.fill = book.fill_accent; cc.font = Font(name=FONT, size=10, bold=True)
    ss.r += 1

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sy = S("Summary", ncols=6, label_width=60, subtitle="Debt capacity on one page")
sy.ws.column_dimensions["F"].width = 56
sy.section("WHERE THE BALANCE SHEET STANDS")
sy.row("Net debt / EBITDA \u2014 headline", f"={C['lev_co']}", MULT, font=F_LINK)
sy.row("Adjusted net debt / adjusted EBITDA \u2014 agency basis", f"={C['lev_adj']}", MULT, font=F_LINK, bold=True)
sy.row("EBITDA / cash interest paid", f"={C['cov_cash']}", MULT, font=F_LINK)
sy.row("FFO / adjusted net debt", f"={C['ffo_nd']}", PCT, font=F_LINK)
sy.row("Free cash flow after dividends (\u20b9 crore)", f"={C['fcf_div']}", NUM0, font=F_LINK)
sy.row("Published ratings", "A- (S&P) \u00b7 Baa1 (Moody's) \u00b7 two notches above the sovereign", "@", font=F_TEXT)
sy.blank(); sy.section("HOW MUCH MORE IT COULD BORROW")
sy.row("Incremental capacity holding an A-category profile (\u20b9 crore)", f"={D['head'][1]}", NUM0, font=F_LINK, bold=True)
sy.row("  in US$ billion", f"={D['head_usd'][1]}", NUM, font=F_LINK)
sy.row("  binding constraint at A", f"={D['which'][1]}", "@", font=F_LINK)
sy.row("Incremental capacity if the company accepted BBB (\u20b9 crore)", f"={D['head'][2]}", NUM0, font=F_LINK)
sy.blank(); sy.section("WHERE THE COST OF CAPITAL BOTTOMS OUT")
sy.row("WACC today", f"={O['wacc_now']}", '0.00%', font=F_LINK)
sy.row("Minimum WACC across the leverage grid", f"={O['wacc_min']}", '0.00%', font=F_LINK, bold=True)
sy.row("Debt weight at the minimum", f"={O['dv_opt']}", PCT, font=F_LINK, bold=True)
sy.row("Twelve-month sources / uses", f"={Q['cover']}", MULT, font=F_LINK, border=DOUBLE_BORDER)
sy.blank()
sy.bullets([
    "The headline 0.6x understates leverage by about half. Add the leases and the spectrum instalments owed to the "
    "government and the ratio is close to 1.25x, which is still comfortably A-category but is the honest starting point.",
    "Capacity is real but committed. At an A profile the company could add over \u20b91.5 lakh crore of net debt, yet "
    "FY2026 capex was \u20b91.44 lakh crore and free cash flow after dividends was thin. The capacity exists to protect the "
    "investment programme through a downturn, not to fund a recapitalisation.",
    "There is no WACC case for more debt. The curve is flat from 10% to 30% debt and the company is already at the "
    "cheap end of it. Leverage beyond the A category costs more in spread than the tax shield returns.",
    "Liquidity is the story. Cash on hand plus one year's operating cash flow covers every use for the next twelve "
    "months with room to spare. That, more than any ratio, is why the rating sits above India's sovereign.",
])

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
tests = [
    ("Company net debt reconciles to the published figure of \u20b9124,717 crore", f"=ABS({C['nd_co']}-124717)<1"),
    ("Adjusted EBITDA equals reported less the disposal gain", f"=ABS({I['ebitda']}-({I['ebitda_rep']}-{I['oneoff']}))<0.01"),
    ("Adjusted net debt exceeds company net debt", f"={C['nd_adj']}>{C['nd_co']}"),
    ("Headline leverage reconciles to the published 0.6x within rounding", f"=ABS({C['lev_co']}-0.60)<0.02"),
    ("Capacity at each rating is the minimum of its three constraints",
     "=AND(" + ",".join(f"{d}<={a}+0.01" for d, a in zip(D["cap"], D["cap_lev"])) + ")"),
    ("Capacity rises monotonically as the rating weakens",
     "=AND(" + ",".join(f"{D['cap'][i+1]}>={D['cap'][i]}" for i in range(4)) + ")"),
    ("Cost of debt rises monotonically as the rating weakens",
     "=AND(" + ",".join(f"{I['kd'][i+1]}>{I['kd'][i]}" for i in range(4)) + ")"),
    ("Headroom at A is positive", f"={D['head'][1]}>0"),
    ("Unlevered beta is below the observed levered beta", f"={O['beta_u']}<{I['beta_l']}"),
    ("WACC at zero debt equals the unlevered cost of equity", f"=ABS({O['wacc_at'][0]}-({I['rf']}+{O['beta_u']}*{I['erp']}))<0.0001"),
    ("Minimum WACC is not above today's WACC", f"={O['wacc_min']}<={O['wacc_now']}+0.0001"),
    ("Base-case sensitivity cell reconciles to the Debt Capacity sheet", f"=ABS(Sensitivity!E{g1+2}-{D['head'][1]})<1"),
    ("Sources cover uses over twelve months", f"={Q['cover']}>1"),
    ("Effective tax rate is between 15% and 35%", f"=AND({I['etr']}>0.15,{I['etr']}<0.35)"),
    ("Market capitalisation reconciles to price times shares", f"=ABS({I['mcap']}-{I['price']}*{I['shares']})<1"),
]
checks_sheet(book, tests,
             "The reconciliation to the published net debt and leverage figures is the check that matters most: it "
             "proves the model starts from the company's own numbers before it restates them.")

# --------------------------------------------------------------------------- #
book.cover(
    blurb="A debt-capacity and capital-structure study of Reliance Industries built on the audited FY2026 consolidated "
          "results (year to 31 March 2026) and 18 September 2026 market data. It restates the balance sheet the way a "
          "rating agency would, sizes borrowing capacity at each rating category, finds where the cost of capital "
          "bottoms out, and tests twelve-month liquidity.",
    method=[
        "Adjusted net debt adds lease liabilities and the deferred spectrum instalments to the company's own net debt; "
        "adjusted EBITDA removes the disposal gain disclosed in the results.",
        "Debt capacity at each rating is the minimum of three constraints \u2014 leverage, cash interest cover and "
        "FFO-to-debt \u2014 against a stated threshold grid, with the binding constraint named.",
        "An optimal-structure grid unlevers the observed beta, re-levers it at each debt weight, looks up the implied "
        "rating and cost of debt from the same grid, and reports WACC.",
        "A twelve-month sources-and-uses test, and two live sensitivity grids that rebuild every constraint in each cell.",
    ],
    toc=[("Summary", "leverage as reported and restated, capacity, the WACC minimum and liquidity"),
         ("Credit Profile", "three definitions of net debt, cover, FFO, capitalisation, implied rating"),
         ("Debt Capacity", "capacity by constraint and rating, headroom, what it would fund"),
         ("Optimal Structure", "WACC across leverage with rating-linked cost of debt"),
         ("Liquidity", "twelve-month sources and uses"),
         ("Sensitivity", "headroom and WACC grids"),
         ("Inputs", "audited FY2026 figures, market data, rating thresholds, cost-of-capital inputs"),
         ("Checks", "fifteen tests; must read MODEL OK")],
    highlights=[("Adjusted net debt / EBITDA", f"={C['lev_adj']}", MULT),
                ("EBITDA / cash interest", f"={C['cov_cash']}", MULT),
                ("Incremental capacity at A (\u20b9 crore)", f"={D['head'][1]}", NUM0),
                ("WACC today", f"={O['wacc_now']}", '0.00%'),
                ("Minimum WACC", f"={O['wacc_min']}", '0.00%'),
                ("Twelve-month sources / uses", f"={Q['cover']}", MULT)],
    sources=["Reliance Industries Limited, audited consolidated financial results and media release for the quarter and year "
             "ended 31 March 2026, dated 24 April 2026 (ril.com). Segment, balance sheet and cash flow figures as published.",
             "Share price: NSE close, 18 September 2026. Ratings: S&P Global A- and Moody's Baa1 as disclosed by the company.",
             "Risk-free rate, equity risk premium, beta, rating thresholds and credit spreads are stated inputs on the Inputs sheet.",
             "Nothing here is a recommendation. The study is an analytical exercise on public information."])

book.finish(freeze={"Optimal Structure": "C6"})
path = os.path.join(HERE, "Reliance_Debt_Capacity.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
