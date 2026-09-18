"""
Build 09-dewa-ipo/DEWA_IPO_Rebuild.xlsx — the April 2022 DEWA IPO, rebuilt.

Sheets: Cover · Summary · Inputs · Valuation · Pricing & Book · Aftermarket · Sensitivity · Checks

An IPO has three questions in it. What is the company worth (Valuation)? Where inside that range should it be
priced, given the book and the discount an issuer must leave for investors (Pricing & Book)? And, with hindsight,
did the pricing get it right (Aftermarket)? DEWA is a clean case: a regulated monopoly with a published five-year
dividend commitment, a secondary sale by one government shareholder, a range of AED 2.25\u20132.48, a book that
allowed the deal to be upsized from 6.5% to 18% of the company, and four years of trading since.

The valuation uses only what was public at the time: FY2021 results from the prospectus summary and the dividend
policy. The peer multiples are stated inputs at approximate April 2022 levels, with today's readings shown beside
them. Aftermarket prices are from the Dubai Financial Market via Yahoo Finance (unadjusted).
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)
from ibkit.sheet import Sheet, checks_sheet

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"
NUM3 = '0.000'

# --------------------------------------------------------------------------- #
# Source data — AED millions unless stated.
# Deal: DEWA price-range (24 Mar 2022) and final-price (6 Apr 2022) announcements; IPO FAQ for FY2021 figures.
# Aftermarket: DFM via Yahoo Finance, unadjusted. Post-IPO financials: FY2022\u2013FY2025 via Yahoo Finance.
# --------------------------------------------------------------------------- #
D = dict(shares=50000.0, range_lo=2.25, range_hi=2.48, final=2.48, offered_init=3250.0, offered_final=9000.0,
         corner_init=4700.0, corner_final=13800.0, demand_ex_corner=315000.0, gross=22320.0,
         rev21=23800.0, ebitda21=12100.0, ni21=6600.0, nd_ebitda21=1.5, nci21=2953.0, div_policy=6200.0,
         d1_open=2.98, d1_high=3.05, d1_close=2.87, y1_close=2.47, now=2.74,
         rev25=32842.0, ebitda25=18283.0, ni25=8347.0, debt25=38630.0, cash25=7233.0, nci25=6753.0,
         divs_since=8 * 0.062)

book = Book(theme=THEMES["dewa"], project_no=9, project="IPO Valuation and Pricing Rebuild (April 2022)",
            company="DEWA", units="AED millions \u00b7 per-share values in AED",
            as_of="FY2021 prospectus figures \u00b7 aftermarket to 18 Sep 2026")
S = lambda name, **kw: Sheet(book, name, **kw)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
si = S("Inputs", ncols=6, label_width=60, subtitle="Deal terms, FY2021 financials, valuation assumptions and the aftermarket record")
si.ws.column_dimensions["F"].width = 60
I = {}
si.section("DEAL TERMS AS ANNOUNCED")
I["shares"] = si.row("Shares in issue (millions)", D["shares"], NUM0)
I["range_lo"] = si.row("Price range \u2014 low (AED)", D["range_lo"], NUM2, note="24 March 2022")
I["range_hi"] = si.row("Price range \u2014 high (AED)", D["range_hi"], NUM2)
I["final"] = si.row("Final offer price (AED)", D["final"], NUM2, note="6 April 2022; the top of the range")
I["off_init"] = si.row("Shares offered at launch (millions)", D["offered_init"], NUM0, note="6.5% of the company")
I["off_final"] = si.row("Shares offered at pricing (millions)", D["offered_final"], NUM0, note="Upsized to 18%")
I["corner_init"] = si.row("Cornerstone commitments at launch", D["corner_init"], NUM0)
I["corner_final"] = si.row("Cornerstone and strategic commitments at pricing", D["corner_final"], NUM0)
I["demand"] = si.row("Total demand excluding cornerstones", D["demand_ex_corner"], NUM0, note="AED 315 billion, per the final-price announcement")
I["fee"] = si.row("Assumed gross underwriting and advisory fees, % of proceeds", 0.0175, PCT, note="Stated assumption; large sovereign deals in the region price below the 3\u20135% global norm")

si.blank(); si.section("FY2021 FINANCIALS FROM THE PROSPECTUS SUMMARY")
I["rev"] = si.row("Revenue", D["rev21"], NUM0, note="Over 95% regulated or contracted")
I["ebitda"] = si.row("Adjusted EBITDA", D["ebitda21"], NUM0)
I["ni"] = si.row("Net income", D["ni21"], NUM0)
I["nd_x"] = si.row("Net debt / EBITDA", D["nd_ebitda21"], MULT)
I["nd"] = si.row("Net debt, derived", f"={I['nd_x']}*{I['ebitda']}", NUM0)
I["nci"] = si.row("Non-controlling interests (Empower minority, FY2022 balance sheet as proxy)", D["nci21"], NUM0)
I["div"] = si.row("Committed annual dividend, first five years", D["div_policy"], NUM0, note="AED 6.2 billion a year, paid in two instalments \u2014 the centre of the equity story")
I["dps"] = si.row("Committed dividend per share (AED)", f"={I['div']}/{I['shares']}", NUM3)

si.blank(); si.section("VALUATION ASSUMPTIONS \u2014 stated inputs at April 2022 levels")
I["y_lo"] = si.row("Target dividend yield \u2014 demanding investor", 0.055, PCT, note="Where Gulf and European integrated utilities traded in spring 2022")
I["y_hi"] = si.row("Target dividend yield \u2014 generous investor", 0.045, PCT)
I["m_lo"] = si.row("EV / EBITDA \u2014 low", 10.0, MULT, note="Regulated networks and integrated utilities; today's peer median is on the Valuation sheet")
I["m_hi"] = si.row("EV / EBITDA \u2014 high", 12.5, MULT)
I["g"] = si.row("EBITDA growth, five years", 0.045, PCT, note="Dubai population and demand growth; FY2021\u2013FY2025 actual was 10.9% a year")
I["capex"] = si.row("Capital expenditure per year", 8000.0, NUM0)
I["tax"] = si.row("UAE corporate tax from 2023 (applied to EBIT)", 0.09, PCT)
I["da_pct"] = si.row("Depreciation as % of EBITDA", 0.45, PCT)
I["wacc"] = si.row("Weighted average cost of capital", 0.065, PCT, note="A dirham-pegged regulated utility with 1.5x leverage")
I["tg"] = si.row("Terminal growth", 0.020, PCT)

si.blank(); si.section("AFTERMARKET (AED per share, unadjusted, DFM)")
I["d1_open"] = si.row("First-day open, 12 April 2022", D["d1_open"], NUM2)
I["d1_high"] = si.row("First-day high", D["d1_high"], NUM2)
I["d1_close"] = si.row("First-day close", D["d1_close"], NUM2)
I["y1"] = si.row("Close one year after listing, 12 April 2023", D["y1_close"], NUM2)
I["now"] = si.row("Close 18 September 2026", D["now"], NUM2)
I["divs_since"] = si.row("Dividends received per share since listing (AED)", D["divs_since"], NUM3, note="Eight semi-annual instalments of AED 0.062 under the policy")

si.blank(); si.section("FY2025 ACTUALS, FOUR YEARS ON")
I["rev25"] = si.row("Revenue", D["rev25"], NUM0)
I["ebitda25"] = si.row("EBITDA", D["ebitda25"], NUM0)
I["ni25"] = si.row("Net income", D["ni25"], NUM0)
I["nd25"] = si.row("Net debt", f"={D['debt25']}-{D['cash25']}", NUM0)
I["nci25"] = si.row("Non-controlling interests", D["nci25"], NUM0)

# --------------------------------------------------------------------------- #
# Valuation
# --------------------------------------------------------------------------- #
sv = S("Valuation", ncols=6, label_width=60, subtitle="Three methods on FY2021 information, against the price range")
sv.ws.column_dimensions["F"].width = 58
V = {}
sv.section("1. DIVIDEND YIELD ON THE COMMITTED PAYOUT")
V["dy_lo"] = sv.row("Equity value at the demanding yield", f"={I['div']}/{I['y_lo']}", NUM0)
V["dy_hi"] = sv.row("Equity value at the generous yield", f"={I['div']}/{I['y_hi']}", NUM0)
V["dy_ps_lo"] = sv.row("  per share, low (AED)", f"={V['dy_lo']}/{I['shares']}", NUM2, bold=True)
V["dy_ps_hi"] = sv.row("  per share, high (AED)", f"={V['dy_hi']}/{I['shares']}", NUM2, bold=True)
V["dy_at_final"] = sv.row("Yield at the final price", f"={I['dps']}/{I['final']}", PCT, bold=True,
                          note="Exactly 5.0%: the price was set so that the committed dividend delivered a round yield")

sv.blank(); sv.section("2. EV / EBITDA ON FY2021 ADJUSTED EBITDA")
V["ev_lo"] = sv.row("Enterprise value, low multiple", f"={I['ebitda']}*{I['m_lo']}", NUM0)
V["ev_hi"] = sv.row("Enterprise value, high multiple", f"={I['ebitda']}*{I['m_hi']}", NUM0)
V["ev_eq_lo"] = sv.row("Equity value, low (EV \u2212 net debt \u2212 minorities)", f"={V['ev_lo']}-{I['nd']}-{I['nci']}", NUM0)
V["ev_eq_hi"] = sv.row("Equity value, high", f"={V['ev_hi']}-{I['nd']}-{I['nci']}", NUM0)
V["ev_ps_lo"] = sv.row("  per share, low (AED)", f"={V['ev_eq_lo']}/{I['shares']}", NUM2, bold=True)
V["ev_ps_hi"] = sv.row("  per share, high (AED)", f"={V['ev_eq_hi']}/{I['shares']}", NUM2, bold=True)
V["ev_at_final"] = sv.row("EV / EBITDA at the final price", f"=({I['final']}*{I['shares']}+{I['nd']}+{I['nci']})/{I['ebitda']}", MULT, bold=True)
V["pe_at_final"] = sv.row("P / E at the final price", f"={I['final']}*{I['shares']}/{I['ni']}", MULT)
sv.subsection("Where utilities trade today (18 Sep 2026, trailing EV / EBITDA) \u2014 context, not an input")
for nm, m in [("Enel", 8.4), ("Engie", 9.3), ("Saudi Electricity", 7.1), ("Duke Energy", 11.3), ("Consolidated Edison", 10.8),
              ("Empower (Dubai)", 11.6), ("Tabreed", 12.1), ("Iberdrola", 13.9), ("National Grid", 14.6)]:
    sv.row("  " + nm, m, MULT, font=F_INPUT)
V["peer_med"] = sv.row("  Median", f"=MEDIAN(C{sv.r-9}:C{sv.r-1})", MULT, bold=True)

sv.blank(); sv.section("3. DISCOUNTED CASH FLOW, FIVE YEARS")
YEARS = ["FY2022E", "FY2023E", "FY2024E", "FY2025E", "FY2026E"]; NY = 5
sv.head(YEARS, first_col=3) if False else None
# Build a compact horizontal DCF on this sheet with its own header row
r_h = sv.r
for j, y in enumerate(YEARS):
    c = sv.ws.cell(r_h, 3 + j, y); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center")
sv.ws.cell(r_h, 2, "Year").font = F_BOLD
# widen this sheet's value columns for the DCF block
for j in range(3, 8): sv.ws.column_dimensions[L(j)].width = 13
sv.ws.column_dimensions["F"].width = 13; sv.ws.column_dimensions["H"].width = 50
sv.r += 1
r_e = sv.r
V["dcf_ebitda"] = sv.multi("EBITDA", [f"={I['ebitda']}*(1+{I['g']})^{j+1}" for j in range(NY)], NUM0)
V["dcf_da"] = sv.multi("Depreciation", [f"={L(3+j)}{r_e}*{I['da_pct']}" for j in range(NY)], NUM0)
V["dcf_tax"] = sv.multi("Tax on EBIT (from 2023)", [f"=-IF({j}>=1,({L(3+j)}{r_e}-{L(3+j)}{r_e+1})*{I['tax']},0)" for j in range(NY)], NUM0)
V["dcf_capex"] = sv.multi("Capital expenditure", [f"=-{I['capex']}"] * NY, NUM0)
V["dcf_fcf"] = sv.multi("Free cash flow to the firm", [f"={L(3+j)}{r_e}+{L(3+j)}{r_e+2}+{L(3+j)}{r_e+3}" for j in range(NY)], NUM0, bold=True, border=TOTAL_BORDER)
V["dcf_df"] = sv.multi("Discount factor", [f"=1/(1+{I['wacc']})^{j+1}" for j in range(NY)], '0.0000')
V["dcf_pv"] = sv.multi("Present value", [f"={L(3+j)}{r_e+4}*{L(3+j)}{r_e+5}" for j in range(NY)], NUM0)
sv.blank()
V["tv"] = sv.row("Terminal value at end FY2026E", f"={L(3+NY-1)}{r_e+4}*(1+{I['tg']})/({I['wacc']}-{I['tg']})", NUM0)
V["tv_pv"] = sv.row("Present value of terminal value", f"={V['tv']}*{L(3+NY-1)}{r_e+5}", NUM0)
V["dcf_ev"] = sv.row("Enterprise value", f"=SUM({L(3)}{r_e+6}:{L(3+NY-1)}{r_e+6})+{V['tv_pv']}", NUM0, bold=True)
V["dcf_eq"] = sv.row("Equity value", f"={V['dcf_ev']}-{I['nd']}-{I['nci']}", NUM0, bold=True, border=TOTAL_BORDER)
V["dcf_ps"] = sv.row("  per share (AED)", f"={V['dcf_eq']}/{I['shares']}", NUM2, bold=True, border=DOUBLE_BORDER)
V["dcf_tv_share"] = sv.row("Terminal value as % of enterprise value", f"={V['tv_pv']}/{V['dcf_ev']}", PCT)

sv.blank(); sv.section("THE THREE METHODS AGAINST THE RANGE (AED per share)")
V["ff_lo"] = sv.row("Lowest of the method lows", f"=MIN({V['dy_ps_lo']},{V['ev_ps_lo']},{V['dcf_ps']})", NUM2)
V["ff_hi"] = sv.row("Highest of the method highs", f"=MAX({V['dy_ps_hi']},{V['ev_ps_hi']},{V['dcf_ps']})", NUM2)
V["ff_mid"] = sv.row("Midpoint of the three method midpoints",
                     f"=AVERAGE(AVERAGE({V['dy_ps_lo']},{V['dy_ps_hi']}),AVERAGE({V['ev_ps_lo']},{V['ev_ps_hi']}),{V['dcf_ps']})", NUM2, bold=True)
V["range_mid"] = sv.row("Midpoint of the announced range", f"=AVERAGE({I['range_lo']},{I['range_hi']})", NUM2, font=F_FORMULA)
V["final_vs_mid"] = sv.row("Final price versus the valuation midpoint", f"={I['final']}/{V['ff_mid']}-1", PCT, bold=True,
                           note="Positive means the deal was priced above a mid-case fair value on FY2021 numbers. The first-day rise was not a fundamental discount")
sv.note("The yield method dominated the marketing: a committed AED 6.2 billion a year at a 5% yield gives AED 124 billion "
        "exactly, and the range top was reverse-engineered from it. The multiple and DCF methods on the same FY2021 numbers "
        "sit lower, and the final price lands about 8% above the midpoint of the three. On fundamentals the deal was priced "
        "full; the first-day rise came from a book that valued the yield above everything else and received 3% allocations, "
        "not from an underpricing discount.", height=48)

# --------------------------------------------------------------------------- #
# Pricing & Book
# --------------------------------------------------------------------------- #
sp = S("Pricing & Book", ncols=6, label_width=60, subtitle="Size, proceeds, cornerstones, coverage and the discount left on the table")
sp.ws.column_dimensions["F"].width = 58
P = {}
sp.section("MARKET CAPITALISATION AT THE RANGE")
P["mcap_lo"] = sp.row("At the bottom of the range", f"={I['range_lo']}*{I['shares']}", NUM0)
P["mcap_hi"] = sp.row("At the top of the range", f"={I['range_hi']}*{I['shares']}", NUM0)
P["mcap"] = sp.row("At the final price", f"={I['final']}*{I['shares']}", NUM0, bold=True, note="AED 124 billion, the largest company on the DFM at listing")
sp.blank(); sp.section("SIZE AND PROCEEDS")
P["stake_init"] = sp.row("Stake offered at launch", f"={I['off_init']}/{I['shares']}", PCT)
P["stake_final"] = sp.row("Stake offered at pricing", f"={I['off_final']}/{I['shares']}", PCT, bold=True)
P["upsize"] = sp.row("Upsizing factor", f"={I['off_final']}/{I['off_init']}", MULT)
P["gross"] = sp.row("Gross proceeds to the selling shareholder", f"={I['final']}*{I['off_final']}", NUM0, bold=True)
P["fees"] = sp.row("Less: fees (assumed)", f"=-{P['gross']}*{I['fee']}", NUM0)
P["net"] = sp.row("Net proceeds", f"={P['gross']}+{P['fees']}", NUM0, border=TOTAL_BORDER)
P["gov_after"] = sp.row("Government of Dubai's stake after the IPO", f"=1-{P['stake_final']}", PCT)
sp.blank(); sp.section("THE BOOK")
P["corner_pct"] = sp.row("Cornerstone and strategic commitments as % of the deal", f"={I['corner_final']}/{P['gross']}", PCT, bold=True,
                         note="Over 60% of the enlarged deal was committed before the book opened")
P["open_book"] = sp.row("Shares placed with the open book (AED)", f"={P['gross']}-{I['corner_final']}", NUM0)
P["coverage"] = sp.row("Coverage of the open book: demand / open-book allocation", f"={I['demand']}/{P['open_book']}", MULT, bold=True, border=DOUBLE_BORDER,
                       note="Reproduces the '37 times oversubscribed' headline")
P["coverage_all"] = sp.row("Coverage of the whole deal, cornerstones included", f"=({I['demand']}+{I['corner_final']})/{P['gross']}", MULT)
P["cov_init"] = sp.row("Coverage the original 6.5% deal would have shown", f"=({I['demand']}+{I['corner_init']})/({I['off_init']}*{I['final']})", MULT,
                       note="The number that made the upsizing possible")
P["fill"] = sp.row("Average fill rate for open-book orders", f"=1/{P['coverage']}", PCT)
sp.blank(); sp.section("THE DISCOUNT")
P["pop"] = sp.row("First-day close versus the offer price", f"={I['d1_close']}/{I['final']}-1", PCT, bold=True)
P["pop_open"] = sp.row("First-day open versus the offer price", f"={I['d1_open']}/{I['final']}-1", PCT)
P["left"] = sp.row("Money left on the table: shares sold \u00d7 (first-day close \u2212 offer)", f"={I['off_final']}*({I['d1_close']}-{I['final']})", NUM0, bold=True)
P["left_pct"] = sp.row("  as % of gross proceeds", f"={P['left']}/{P['gross']}", PCT)
P["left_vs_fee"] = sp.row("  as a multiple of the fees paid", f"={P['left']}/-{P['fees']}", MULT)
P["price_for_10"] = sp.row("Offer price that would have left a 10% first-day discount (AED)", f"={I['d1_close']}/1.10", NUM2,
                           note="Above the range top. The range, not the pricing, was the conservative decision")
sp.note("A 15\u201320% first-day rise is at the upper end of what issuers accept and the lower end of what a 37-times book "
        "would predict. The range was fixed two weeks before pricing and already sat at or above fundamental value, so "
        "the only lever left was size: the seller nearly tripled the deal rather than move the price. The money left on "
        "the table is the cost of that choice, and it bought a much larger free float.", height=44)

# --------------------------------------------------------------------------- #
# Aftermarket
# --------------------------------------------------------------------------- #
sa = S("Aftermarket", ncols=6, label_width=60, subtitle="Four years on: what an IPO investor earned, and what the company delivered")
sa.ws.column_dimensions["F"].width = 58
A = {}
sa.section("RETURNS TO AN IPO INVESTOR (AED per share)")
A["ret_d1"] = sa.row("Day one", f"={P['pop']}", PCT, font=F_LINK)
A["ret_y1"] = sa.row("One year, price only", f"={I['y1']}/{I['final']}-1", PCT, note="Below the offer price after twelve months despite the yield")
A["ret_y1_tr"] = sa.row("One year, with two dividend instalments", f"=({I['y1']}+2*{I['dps']}/2)/{I['final']}-1", PCT)
A["ret_now"] = sa.row("To 18 September 2026, price only", f"={I['now']}/{I['final']}-1", PCT)
A["ret_now_tr"] = sa.row("To 18 September 2026, with dividends received", f"=({I['now']}+{I['divs_since']})/{I['final']}-1", PCT, bold=True)
A["years"] = sa.row("Years since listing", 4.44, NUM2)
A["cagr"] = sa.row("Annualised total return", f"=(1+{A['ret_now_tr']})^(1/{A['years']})-1", PCT, bold=True, border=DOUBLE_BORDER)
A["yield_cost"] = sa.row("Running yield on the IPO price", f"={I['dps']}/{I['final']}", PCT)
sa.blank(); sa.section("WHAT THE COMPANY DELIVERED, FY2021 TO FY2025")
A["g_rev"] = sa.row("Revenue growth, annualised", f"=({I['rev25']}/{I['rev']})^(1/4)-1", PCT)
A["g_ebitda"] = sa.row("EBITDA growth, annualised", f"=({I['ebitda25']}/{I['ebitda']})^(1/4)-1", PCT, bold=True, note="More than double the 4.5% assumed in the DCF")
A["g_ni"] = sa.row("Net income growth, annualised", f"=({I['ni25']}/{I['ni']})^(1/4)-1", PCT)
A["div_cover"] = sa.row("Dividend cover, FY2025 (net income / committed dividend)", f"={I['ni25']}/{I['div']}", MULT)
sa.blank(); sa.section("MULTIPLES THEN AND NOW")
A["ev_then"] = sa.row("EV / EBITDA at the IPO, FY2021", f"={V['ev_at_final']}", MULT, font=F_LINK)
A["ev_now"] = sa.row("EV / EBITDA today on FY2025", f"=({I['now']}*{I['shares']}+{I['nd25']}+{I['nci25']})/{I['ebitda25']}", MULT, bold=True)
A["pe_then"] = sa.row("P / E at the IPO", f"={V['pe_at_final']}", MULT, font=F_LINK)
A["pe_now"] = sa.row("P / E today on FY2025", f"={I['now']}*{I['shares']}/{I['ni25']}", MULT)
A["dy_now"] = sa.row("Dividend yield today", f"={I['dps']}/{I['now']}", PCT)
A["derate"] = sa.row("De-rating: change in EV / EBITDA multiple", f"={A['ev_now']}/{A['ev_then']}-1", PCT, bold=True)
sa.note("The business grew faster than the IPO case and the shares still spent most of four years near the offer price. "
        "The gap is the multiple: DEWA was sold at twelve times EBITDA and has since de-rated toward where global utilities "
        "trade. An IPO investor's return has come almost entirely from the dividend it was promised. That is the shape of "
        "a correctly priced regulated-utility IPO \u2014 the story was the yield, and the yield was delivered.", height=48)

# --------------------------------------------------------------------------- #
# Sensitivity
# --------------------------------------------------------------------------- #
ss = S("Sensitivity", ncols=8, label_width=40, col_width=13.5, subtitle="Value per share (AED) \u2014 each cell rebuilds the method")
YS = [0.040, 0.045, 0.050, 0.055, 0.060]; DIVS = [5600.0, 5900.0, 6200.0, 6500.0, 6800.0]
WS = [0.055, 0.060, 0.065, 0.070, 0.075]; TGS = [0.010, 0.015, 0.020, 0.025, 0.030]
def grid(title, rows, cols, rfmt, cfmt, fn, base):
    ss.section(title)
    ss.ws.cell(ss.r, 2, "\u2193 rows  \\  columns \u2192").font = F_BOLD
    for j, cval in enumerate(cols):
        c = ss.ws.cell(ss.r, 3 + j, cval); c.font = book.f_header; c.fill = book.fill_secondary
        c.number_format = cfmt; c.alignment = Alignment(horizontal="center")
    h = ss.r; ss.r += 1; g0 = ss.r
    for rv in rows:
        c = ss.ws.cell(ss.r, 2, rv); c.font = F_BOLD; c.number_format = rfmt
        for j, cval in enumerate(cols):
            cc = ss.ws.cell(ss.r, 3 + j, fn(f"$B{ss.r}", f"{L(3+j)}${h}")); cc.font = F_FORMULA; cc.number_format = NUM2
            if abs(rv - base[0]) < 1e-9 and abs(cval - base[1]) < 1e-9:
                cc.fill = book.fill_accent; cc.font = Font(name=FONT, size=10, bold=True)
        ss.r += 1
    ss.blank(); return g0
g1 = grid("DIVIDEND YIELD METHOD \u2014 TARGET YIELD (rows) versus ANNUAL DIVIDEND, AED m (columns)", YS, DIVS, PCT, NUM0,
          lambda r, c: f"={c}/{r}/{I['shares']}", (0.050, 6200.0))
pv_exp = "+".join(f"{V['dcf_fcf'][j]}/(1+{{w}})^{j+1}" for j in range(NY))
def dcf_fn(w, g):
    e = pv_exp.replace("{w}", w)
    tv = f"{V['dcf_fcf'][NY-1]}*(1+{g})/({w}-{g})/(1+{w})^{NY}"
    return f"=(({e})+{tv}-{I['nd']}-{I['nci']})/{I['shares']}"
g2 = grid("DCF \u2014 WACC (rows) versus TERMINAL GROWTH (columns)", WS, TGS, PCT, PCT, dcf_fn, (0.065, 0.020))

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sy = S("Summary", ncols=6, label_width=62, subtitle="The DEWA IPO on one page")
sy.ws.column_dimensions["F"].width = 56
sy.section("THE DEAL")
sy.row("Price range (AED)", f"=TEXT({I['range_lo']},\"0.00\")&\" \u2013 \"&TEXT({I['range_hi']},\"0.00\")", "@", font=F_LINK)
sy.row("Final price (AED)", f"={I['final']}", NUM2, font=F_LINK, bold=True)
sy.row("Market capitalisation at listing", f"={P['mcap']}", NUM0, font=F_LINK)
sy.row("Stake sold: launched at / priced at", f"=TEXT({P['stake_init']},\"0.0%\")&\" / \"&TEXT({P['stake_final']},\"0.0%\")", "@", font=F_LINK)
sy.row("Gross proceeds", f"={P['gross']}", NUM0, font=F_LINK, bold=True)
sy.row("Cornerstones as % of the deal", f"={P['corner_pct']}", PCT, font=F_LINK)
sy.row("Open-book coverage", f"={P['coverage']}", MULT, font=F_LINK, bold=True)
sy.blank(); sy.section("THE VALUATION")
sy.row("Dividend-yield method (AED per share)", f"=TEXT({V['dy_ps_lo']},\"0.00\")&\" \u2013 \"&TEXT({V['dy_ps_hi']},\"0.00\")", "@", font=F_LINK)
sy.row("EV / EBITDA method (AED per share)", f"=TEXT({V['ev_ps_lo']},\"0.00\")&\" \u2013 \"&TEXT({V['ev_ps_hi']},\"0.00\")", "@", font=F_LINK)
sy.row("DCF (AED per share)", f"={V['dcf_ps']}", NUM2, font=F_LINK)
sy.row("Final price versus the valuation midpoint", f"={V['final_vs_mid']}", PCT, font=F_LINK, bold=True)
sy.row("Yield / EV-EBITDA / P-E at the final price", f"=TEXT({V['dy_at_final']},\"0.0%\")&\" / \"&TEXT({V['ev_at_final']},\"0.0x\")&\" / \"&TEXT({V['pe_at_final']},\"0.0x\")", "@", font=F_LINK)
sy.blank(); sy.section("THE VERDICT")
sy.row("First-day rise", f"={P['pop']}", PCT, font=F_LINK, bold=True)
sy.row("Money left on the table", f"={P['left']}", NUM0, font=F_LINK)
sy.row("Total return to an IPO investor to date", f"={A['ret_now_tr']}", PCT, font=F_LINK, bold=True)
sy.row("  annualised", f"={A['cagr']}", PCT, font=F_LINK)
sy.row("EBITDA growth delivered versus assumed", f"=TEXT({A['g_ebitda']},\"0.0%\")&\" versus \"&TEXT({I['g']},\"0.0%\")", "@", font=F_LINK)
sy.row("Multiple then versus now", f"=TEXT({A['ev_then']},\"0.0x\")&\" \u2192 \"&TEXT({A['ev_now']},\"0.0x\")", "@", font=F_LINK, border=DOUBLE_BORDER)
sy.blank()
sy.bullets([
    "The price was the dividend. AED 6.2 billion a year at a 5% yield is AED 124 billion, which is the top of the range "
    "to the dirham. The book-building confirmed a number that had been chosen before the book opened.",
    "The book was three times bigger than the deal it was built for. Demand of AED 315 billion against a launch size of "
    "AED 7\u20138 billion let the seller nearly triple the offering to 18% and still report 37 times coverage. Most of the "
    "enlarged deal went to cornerstones who had already agreed the price.",
    "On FY2021 fundamentals the price was full \u2014 about 8% above the midpoint of the yield, multiple and DCF methods \u2014 "
    "yet the shares still rose 16% on day one and AED 3.5 billion was left on the table, nine times the fees. A book "
    "filled at 3% cannot be satisfied by any price inside a range; the seller took size instead.",
    "Four years on the shares are close to the offer price and the business has grown EBITDA at 11% a year. The return "
    "has been the dividend, and the multiple has de-rated toward global utilities. For a regulated monopoly sold on yield, "
    "that is the pricing working as intended.",
])

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
tests = [
    ("Market cap at the final price reproduces the announced AED 124 billion", f"=ABS({P['mcap']}-124000)<1"),
    ("Gross proceeds reproduce the announced AED 22.3 billion", f"=ABS({P['gross']}-22320)<1"),
    ("Stake sold reproduces the announced 18%", f"=ABS({P['stake_final']}-0.18)<0.0001"),
    ("Open-book coverage reproduces the announced 37 times within rounding", f"=ABS({P['coverage']}-37)<1"),
    ("Yield at the final price is 5.0%", f"=ABS({V['dy_at_final']}-0.05)<0.0005"),
    ("Final price sits at the top of the range", f"={I['final']}={I['range_hi']}"),
    ("Dividend-yield value at 5% reproduces the final price", f"=ABS({I['div']}/0.05/{I['shares']}-{I['final']})<0.001"),
    ("DCF free cash flow is positive in every year", "=AND(" + ",".join(f"{c}>0" for c in V["dcf_fcf"]) + ")"),
    ("Terminal growth is below WACC", f"={I['tg']}<{I['wacc']}"),
    ("Base yield sensitivity cell reproduces the final price", f"=ABS(Sensitivity!E{g1+2}-{I['final']})<0.001"),
    ("Base DCF sensitivity cell reconciles to the Valuation sheet", f"=ABS(Sensitivity!E{g2+2}-{V['dcf_ps']})<0.001"),
    ("Money left on the table equals shares sold times the first-day gain", f"=ABS({P['left']}-{I['off_final']}*({I['d1_close']}-{I['final']}))<0.01"),
    ("Net debt derived from the disclosed 1.5x lies between AED 15bn and AED 21bn", f"=AND({I['nd']}>15000,{I['nd']}<21000)"),
    ("Government stake after the IPO is 82%", f"=ABS({P['gov_after']}-0.82)<0.0001"),
]
checks_sheet(book, tests,
             "The first five checks reproduce the headline numbers from the company's own announcements \u2014 AED 124 billion, "
             "AED 22.3 billion, 18%, 37 times, 5% \u2014 from the raw inputs, which is what makes this a rebuild rather than a retelling.")

# --------------------------------------------------------------------------- #
book.cover(
    blurb="A rebuild of the April 2022 DEWA initial public offering from public announcements: the valuation on FY2021 "
          "information by three methods, the pricing inside the range and the mechanics of the book and the upsizing, "
          "the discount measured on the first day of trading, and the four-year record since against what the IPO case assumed.",
    method=[
        "Valuation by dividend yield on the committed AED 6.2 billion payout, EV / EBITDA on FY2021 adjusted EBITDA, and a "
        "five-year DCF \u2014 all on information available at the time, with today's utility multiples shown for context only.",
        "Pricing and book: market capitalisation across the range, proceeds, cornerstone share, coverage of the open book "
        "reproducing the announced 37 times, and the coverage the original 6.5% deal would have shown.",
        "The IPO discount from unadjusted first-day prices, money left on the table against fees, and the price that would "
        "have left a 10% first-day rise.",
        "Aftermarket: total return with dividends, EBITDA growth delivered against assumed, and the multiple then and now.",
    ],
    toc=[("Summary", "deal, valuation, verdict"),
         ("Valuation", "yield, multiple and DCF against the range"),
         ("Pricing & Book", "size, proceeds, cornerstones, coverage, discount"),
         ("Aftermarket", "returns to an IPO investor, delivery against the case, re-rating"),
         ("Sensitivity", "yield method and DCF grids"),
         ("Inputs", "deal terms, FY2021 figures, assumptions, aftermarket data"),
         ("Checks", "fourteen tests; must read MODEL OK")],
    highlights=[("Final price (AED)", f"={I['final']}", NUM2),
                ("Gross proceeds (AED m)", f"={P['gross']}", NUM0),
                ("Open-book coverage", f"={P['coverage']}", MULT),
                ("First-day rise", f"={P['pop']}", PCT),
                ("Money left on the table (AED m)", f"={P['left']}", NUM0),
                ("Total return to date, with dividends", f"={A['ret_now_tr']}", PCT)],
    sources=["DEWA PJSC price-range announcement (24 March 2022), final-price announcement (6 April 2022) and IPO FAQ: deal terms, cornerstone commitments, demand, FY2021 revenue, adjusted EBITDA, net income and leverage.",
             "Aftermarket prices: Dubai Financial Market via Yahoo Finance, unadjusted. FY2022\u2013FY2025 financials via Yahoo Finance for DEWA.AE.",
             "Target yields, peer multiples at April 2022 levels, DCF assumptions and the fee rate are stated inputs; today's peer multiples are shown for context and are not used in the valuation.",
             "Nothing here is a recommendation. The rebuild is an analytical exercise on public information."])

book.finish()
path = os.path.join(HERE, "DEWA_IPO_Rebuild.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
