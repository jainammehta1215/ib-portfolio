"""
Build 13-emaar-nav/Emaar_NAV.xlsx — Emaar Properties valued as a real-estate company: net asset value, and a
residual-value development appraisal for a single tower.

Sheets: Cover · Summary · Inputs · NAV · Development Appraisal · Sensitivity · Checks

Property companies are valued on what they own, not on a multiple of what they earned last year. The NAV here has
four asset blocks: the UAE development business (present value of the margin embedded in the sold-but-unrecognised
backlog, plus the land bank), international development, the mall and commercial-leasing portfolio capitalised at a
yield, and hospitality and leisure capitalised at a yield. Net cash is added with a stated haircut for the share of
cash that is customer advances held for construction, minorities in Emaar Development are removed through the stake,
and the result is compared with the share price to give the discount to NAV that the market is applying.

The Development Appraisal is the calculation a developer runs before buying a plot: gross development value less
construction, soft costs, marketing and finance, less a target profit, leaves the residual land value. It is
illustrative and every input is stated.

Source data: Emaar's FY2025 results announcement (12 February 2026) for backlog, recurring revenue and EBITDA and the
land bank; FY2025 balance sheet via public market data; DFM prices 18 September 2026. Yields, margins and the cash
haircut are stated inputs.
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
# Source data — AED millions unless stated.
# FY2025 results announcement, 12 Feb 2026: revenue backlog AED 155bn (Emaar Development AED 125.2bn); recurring revenue
# AED 10.5bn, recurring EBITDA AED 8.1bn; hospitality/leisure revenue AED 4.2bn; land bank 660m sq ft (370m UAE).
# Balance sheet FY2025 via Yahoo Finance for EMAAR.AE. Prices: DFM close 18 Sep 2026.
# --------------------------------------------------------------------------- #
E = dict(revenue=49557.0, ebitda=25623.0, ni=17599.0, cash=52633.0, debt=10615.0, equity=94278.0, nci=13398.0,
         shares=8838.8, price=11.80, w52lo=10.15, w52hi=17.25,
         backlog_grp=155000.0, backlog_dev=125200.0, rec_rev=10500.0, rec_ebitda=8100.0, hosp_rev=4200.0,
         land_uae=370.0, land_intl=290.0,           # million sq ft
         edev_price=13.52, edev_shares=4000.0, edev_ebitda=14300.0, edev_stake=0.80)

book = Book(theme=THEMES["emaar"], project_no=13, project="Real-Estate NAV and Development Appraisal",
            company="Emaar Properties", units="AED millions \u00b7 per-share values in AED",
            as_of="FY2025 results \u00b7 market data 18 Sep 2026")
S = lambda name, **kw: Sheet(book, name, **kw)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
si = S("Inputs", ncols=6, label_width=60, subtitle="Disclosed operating data, balance sheet, market data and valuation assumptions")
si.ws.column_dimensions["F"].width = 62
I = {}
si.section("DISCLOSED OPERATING DATA, FY2025 (AED millions unless stated)")
I["backlog_grp"] = si.row("Revenue backlog, group", E["backlog_grp"], NUM0, note="Sold but not yet recognised; 12 February 2026 announcement")
I["backlog_dev"] = si.row("  of which Emaar Development (UAE)", E["backlog_dev"], NUM0)
I["backlog_intl"] = si.row("  of which international and other", f"={I['backlog_grp']}-{I['backlog_dev']}", NUM0)
I["rec_rev"] = si.row("Recurring revenue (malls, hospitality, leisure, commercial)", E["rec_rev"], NUM0)
I["rec_ebitda"] = si.row("Recurring EBITDA", E["rec_ebitda"], NUM0)
I["hosp_rev"] = si.row("  of which hospitality, leisure and entertainment revenue", E["hosp_rev"], NUM0)
I["hosp_margin"] = si.row("  assumed hospitality EBITDA margin", 0.55, PCT, note="Emaar's hotels are owned, not leased; disclosed 9M margins run in the mid-50s")
I["hosp_ebitda"] = si.row("Hospitality and leisure EBITDA", f"={I['hosp_rev']}*{I['hosp_margin']}", NUM0)
I["mall_ebitda"] = si.row("Malls and commercial leasing EBITDA (recurring less hospitality)", f"={I['rec_ebitda']}-{I['hosp_ebitda']}", NUM0, bold=True)
I["land_uae"] = si.row("UAE land bank (million sq ft)", E["land_uae"], NUM0)
I["land_intl"] = si.row("International land bank (million sq ft)", E["land_intl"], NUM0)
I["edev_ebitda"] = si.row("Emaar Development EBITDA, FY2025", E["edev_ebitda"], NUM0, note="52% margin on AED 27.5bn revenue")

si.blank(); si.section("BALANCE SHEET, 31 DECEMBER 2025")
I["cash"] = si.row("Cash and bank balances", E["cash"], NUM0)
I["debt"] = si.row("Borrowings and sukuk", E["debt"], NUM0)
I["equity"] = si.row("Equity attributable to owners", E["equity"], NUM0)
I["nci"] = si.row("Non-controlling interests (memo)", E["nci"], NUM0)
I["bvps"] = si.row("Book value per share (AED)", f"={I['equity']}/8838.8", NUM2)

si.blank(); si.section("MARKET DATA")
I["price"] = si.row("Emaar Properties share price (AED, DFM close 18 Sep 2026)", E["price"], NUM2)
I["shares"] = si.row("Shares outstanding (millions)", E["shares"], NUM0)
I["mcap"] = si.row("Market capitalisation", f"={I['price']}*{I['shares']}", NUM0, bold=True)
I["pb"] = si.row("Price / book", f"={I['price']}/{I['bvps']}", MULT)
I["edev_price"] = si.row("Emaar Development share price (AED)", E["edev_price"], NUM2)
I["edev_shares"] = si.row("Emaar Development shares (millions)", E["edev_shares"], NUM0)
I["edev_mcap"] = si.row("Emaar Development market capitalisation", f"={I['edev_price']}*{I['edev_shares']}", NUM0)
I["edev_stake"] = si.row("Emaar Properties' stake in Emaar Development", E["edev_stake"], PCT, note="Approximately 80% since the 2017 listing; stated input")

si.blank(); si.section("VALUATION ASSUMPTIONS \u2014 stated inputs")
I["dev_margin"] = si.row("Development margin on backlog revenue (pre-tax)", 0.42, PCT,
                         note="Emaar Development reports a 52% EBITDA margin; a lower figure allows for the cost inflation typical of a late-cycle backlog")
I["intl_margin"] = si.row("International development margin", 0.25, PCT)
I["dev_years"] = si.row("Average years to recognise the backlog", 3.5, NUM, note="Delivery schedules run three to five years from launch")
I["dev_dr"] = si.row("Discount rate for development cash flows", 0.11, PCT, note="Above the cost of equity for the group: development is the riskiest activity in it")
I["tax"] = si.row("UAE corporate tax", 0.09, PCT)
I["land_uae_v"] = si.row("UAE land bank value (AED per sq ft of plot)", 45.0, NUM, note="Master-plan land carried at historic cost; a conservative blended figure across Dubai and the Northern Emirates")
I["land_intl_v"] = si.row("International land bank value (AED per sq ft of plot)", 8.0, NUM)
I["mall_yield"] = si.row("Capitalisation yield \u2014 malls and commercial", 0.070, PCT, note="Prime Dubai retail; Dubai Mall is the single most valuable asset in the group")
I["hosp_yield"] = si.row("Capitalisation yield \u2014 hospitality and leisure", 0.085, PCT)
I["cash_hair"] = si.row("Share of cash treated as customer advances held for construction", 0.50, PCT,
                        note="Off-plan buyers pay into escrow ahead of delivery. That cash funds the backlog and is not free to shareholders")

# --------------------------------------------------------------------------- #
# NAV
# --------------------------------------------------------------------------- #
sn = S("NAV", ncols=6, label_width=64, subtitle="Net asset value, block by block, and the discount the market applies")
sn.ws.column_dimensions["F"].width = 58
N = {}
sn.section("A. UAE DEVELOPMENT \u2014 EMAAR DEVELOPMENT")
N["a_profit"] = sn.row("Pre-tax profit embedded in the backlog", f"={I['backlog_dev']}*{I['dev_margin']}", NUM0)
N["a_df"] = sn.row("Discount factor at the mid-point of delivery", f"=1/(1+{I['dev_dr']})^({I['dev_years']}/2)", '0.000')
N["a_pv"] = sn.row("Present value of backlog profit, after tax", f"={N['a_profit']}*{N['a_df']}*(1-{I['tax']})", NUM0, bold=True)
N["a_land"] = sn.row("UAE land bank", f"={I['land_uae']}*{I['land_uae_v']}", NUM0)
N["a_gross"] = sn.row("UAE development, 100%", f"={N['a_pv']}+{N['a_land']}", NUM0, bold=True, border=TOTAL_BORDER)
N["a_share"] = sn.row("Emaar Properties' share", f"={N['a_gross']}*{I['edev_stake']}", NUM0, bold=True)
N["a_mkt"] = sn.row("  cross-check: Emaar Development market value \u00d7 stake", f"={I['edev_mcap']}*{I['edev_stake']}", NUM0, font=F_FORMULA,
                    note="The listed subsidiary's price is a second opinion on block A")
N["a_vs"] = sn.row("  NAV block versus the listed value", f"={N['a_share']}/{N['a_mkt']}-1", PCT)

sn.blank(); sn.section("B. INTERNATIONAL DEVELOPMENT")
N["b_pv"] = sn.row("Present value of international backlog profit, after tax", f"={I['backlog_intl']}*{I['intl_margin']}*{N['a_df']}*(1-{I['tax']})", NUM0)
N["b_land"] = sn.row("International land bank", f"={I['land_intl']}*{I['land_intl_v']}", NUM0)
N["b"] = sn.row("International development", f"={N['b_pv']}+{N['b_land']}", NUM0, bold=True, border=TOTAL_BORDER,
                note="Emaar Misr and the Indian business have their own minorities; ignored here, which overstates block B slightly")

sn.blank(); sn.section("C. MALLS AND COMMERCIAL LEASING")
N["c"] = sn.row("EBITDA capitalised at the yield", f"={I['mall_ebitda']}/{I['mall_yield']}", NUM0, bold=True, border=TOTAL_BORDER)
N["c_mult"] = sn.row("  implied EV / EBITDA", f"={N['c']}/{I['mall_ebitda']}", MULT)

sn.blank(); sn.section("D. HOSPITALITY, LEISURE AND ENTERTAINMENT")
N["d"] = sn.row("EBITDA capitalised at the yield", f"={I['hosp_ebitda']}/{I['hosp_yield']}", NUM0, bold=True, border=TOTAL_BORDER)

sn.blank(); sn.section("GROSS ASSET VALUE TO NAV")
N["gav"] = sn.row("Gross asset value (A + B + C + D)", f"={N['a_share']}+{N['b']}+{N['c']}+{N['d']}", NUM0, bold=True)
N["cash_free"] = sn.row("Add: cash not held as customer advances", f"={I['cash']}*(1-{I['cash_hair']})", NUM0)
N["debt"] = sn.row("Less: borrowings", f"=-{I['debt']}", NUM0)
N["nav"] = sn.row("Net asset value", f"={N['gav']}+{N['cash_free']}+{N['debt']}", NUM0, bold=True, border=DOUBLE_BORDER)
sn.blank(); sn.section("PER SHARE AND AGAINST THE MARKET")
N["navps"] = sn.row("NAV per share (AED)", f"={N['nav']}/{I['shares']}", NUM2, bold=True)
N["px"] = sn.row("Share price (AED)", f"={I['price']}", NUM2, font=F_LINK)
N["disc"] = sn.row("Discount of the price to NAV", f"=1-{N['px']}/{N['navps']}", PCT, bold=True, border=DOUBLE_BORDER)
N["p_nav"] = sn.row("Price / NAV", f"={N['px']}/{N['navps']}", MULT)
N["nav_bv"] = sn.row("NAV / book value (the revaluation surplus the accounts do not show)", f"={N['nav']}/{I['equity']}", MULT)
N["rec_share"] = sn.row("Share of gross asset value in recurring-income assets (C + D)", f"=({N['c']}+{N['d']})/{N['gav']}", PCT)
N["impl_disc"] = sn.row("Market capitalisation as a share of recurring assets plus free cash alone", f"={I['mcap']}/({N['c']}+{N['d']}+{N['cash_free']}+{N['debt']})", PCT,
                        note="Below 100% means the market is paying nothing for either development business")
sn.note("UAE developers have traded at discounts of 30\u201350% to bottom-up NAV through most of their listed history. The "
        "discount is the market's judgement on the cycle: a backlog sold at today's prices is only worth its margin if "
        "buyers complete, and the land bank is worth its stated value only if launches continue. The model does not "
        "predict the discount; it measures it.", height=48)

# --------------------------------------------------------------------------- #
# Development Appraisal
# --------------------------------------------------------------------------- #
sa = S("Development Appraisal", ncols=6, label_width=64, subtitle="Residual land value for one illustrative Dubai residential tower \u2014 AED millions")
sa.ws.column_dimensions["F"].width = 58
A = {}
sa.section("THE SCHEME \u2014 stated inputs")
A["gfa"] = sa.row("Gross floor area (thousand sq ft)", 600.0, NUM0)
A["eff"] = sa.row("Saleable efficiency", 0.82, PCT, note="Net saleable area as a share of gross floor area")
A["nsa"] = sa.row("Net saleable area (thousand sq ft)", f"={A['gfa']}*{A['eff']}", NUM0)
A["psf"] = sa.row("Average sale price (AED per sq ft of saleable area)", 2400.0, NUM0, note="A prime but not trophy Dubai location, September 2026")
A["build_psf"] = sa.row("Construction cost (AED per sq ft of gross floor area)", 700.0, NUM0)
A["soft"] = sa.row("Professional and soft costs as % of construction", 0.12, PCT)
A["mkt"] = sa.row("Sales and marketing as % of gross development value", 0.035, PCT)
A["months"] = sa.row("Programme (months from land purchase to completion)", 36, '0')
A["fin_rate"] = sa.row("Finance rate on development costs", 0.070, PCT)
A["fin_draw"] = sa.row("Average share of costs outstanding over the programme", 0.50, PCT, note="S-curve approximation")
A["profit_cost"] = sa.row("Target developer profit on total cost", 0.20, PCT)
A["presale"] = sa.row("Share of GDV collected before completion (off-plan)", 0.60, PCT, note="Reduces the finance charge; Dubai off-plan payment plans typically collect 40\u201370% during construction")

sa.blank(); sa.section("GROSS DEVELOPMENT VALUE")
A["gdv"] = sa.row("Gross development value", f"={A['nsa']}*{A['psf']}/1000", NUM0, bold=True, border=TOTAL_BORDER)

sa.blank(); sa.section("COSTS")
A["c_build"] = sa.row("Construction", f"={A['gfa']}*{A['build_psf']}/1000", NUM0)
A["c_soft"] = sa.row("Professional and soft costs", f"={A['c_build']}*{A['soft']}", NUM0)
A["c_mkt"] = sa.row("Sales and marketing", f"={A['gdv']}*{A['mkt']}", NUM0)
A["c_fin"] = sa.row("Finance on costs net of off-plan receipts",
                    f"=MAX(0,({A['c_build']}+{A['c_soft']})-{A['gdv']}*{A['presale']}*{A['fin_draw']})*{A['fin_rate']}*{A['months']}/12*{A['fin_draw']}", NUM0,
                    note="Off-plan collections fund construction as it proceeds, which is why Dubai developers carry so little debt")
A["c_tot"] = sa.row("Total development cost before land", f"={A['c_build']}+{A['c_soft']}+{A['c_mkt']}+{A['c_fin']}", NUM0, bold=True, border=TOTAL_BORDER)

sa.blank(); sa.section("RESIDUAL LAND VALUE")
A["profit"] = sa.row("Developer profit required: target % of (cost + land), solved", f"={A['gdv']}*{A['profit_cost']}/(1+{A['profit_cost']})", NUM0,
                     note="Profit is a share of total cost including land. Since cost + land = GDV − profit, profit = k × GDV / (1 + k)")
A["rlv"] = sa.row("Residual land value", f"={A['gdv']}-{A['c_tot']}-{A['profit']}", NUM0, bold=True, border=DOUBLE_BORDER)
A["rlv_gfa"] = sa.row("Residual land value per sq ft of gross floor area (AED)", f"={A['rlv']}*1000/{A['gfa']}", NUM0, bold=True)
A["land_gdv"] = sa.row("Land as % of gross development value", f"={A['rlv']}/{A['gdv']}", PCT)
A["profit_gdv"] = sa.row("Developer profit as % of gross development value", f"={A['profit']}/{A['gdv']}", PCT)
A["check_pc"] = sa.row("Check: profit / (total cost + land)", f"={A['profit']}/({A['c_tot']}+{A['rlv']})", PCT)
A["breakeven"] = sa.row("Break-even sale price if land is bought at the residual (AED per sq ft)", f"=({A['c_tot']}+{A['rlv']})*1000/{A['nsa']}", NUM0, bold=True)
A["cushion"] = sa.row("Price cushion before the developer loses money", f"=1-{A['breakeven']}/{A['psf']}", PCT, bold=True)
sa.note("The residual is the most a developer can pay for the plot and still make its target return. Every assumption "
        "flows straight into it: a 10% fall in sale prices takes far more than 10% off the land value, which is why land "
        "is the most volatile asset in property and why Emaar's land bank is carried at cost.", height=40)

# --------------------------------------------------------------------------- #
# Sensitivity
# --------------------------------------------------------------------------- #
ss = S("Sensitivity", ncols=8, label_width=40, col_width=13.5, subtitle="NAV per share (AED) and residual land value \u2014 every cell rebuilds the calculation")
YLDS = [0.060, 0.065, 0.070, 0.075, 0.080]; MARG = [0.34, 0.38, 0.42, 0.46, 0.50]
def nav_formula(yld, marg):
    a = f"(({I['backlog_dev']}*{marg}*{N['a_df']}*(1-{I['tax']})+{N['a_land']})*{I['edev_stake']})"
    return f"=({a}+{N['b']}+{I['mall_ebitda']}/{yld}+{N['d']}+{N['cash_free']}+{N['debt']})/{I['shares']}"
def grid(title, rows, cols, rfmt, cfmt, fn, base, fmt=NUM2):
    ss.section(title)
    ss.ws.cell(ss.r, 2, "\u2193 rows  \\  columns \u2192").font = F_BOLD
    for j, cval in enumerate(cols):
        c = ss.ws.cell(ss.r, 3 + j, cval); c.font = book.f_header; c.fill = book.fill_secondary
        c.number_format = cfmt; c.alignment = Alignment(horizontal="center")
    h = ss.r; ss.r += 1; g0 = ss.r
    for rv in rows:
        c = ss.ws.cell(ss.r, 2, rv); c.font = F_BOLD; c.number_format = rfmt
        for j, cval in enumerate(cols):
            cc = ss.ws.cell(ss.r, 3 + j, fn(f"$B{ss.r}", f"{L(3+j)}${h}")); cc.font = F_FORMULA; cc.number_format = fmt
            if abs(rv - base[0]) < 1e-9 and abs(cval - base[1]) < 1e-9:
                cc.fill = book.fill_accent; cc.font = Font(name=FONT, size=10, bold=True)
        ss.r += 1
    ss.blank(); return g0
g1 = grid("NAV PER SHARE \u2014 MALL YIELD (rows) versus DEVELOPMENT MARGIN (columns)", YLDS, MARG, PCT, PCT,
          lambda r, c: nav_formula(r, c), (0.070, 0.42))
PSF = [2000, 2200, 2400, 2600, 2800]; BLD = [600, 650, 700, 750, 800]
def rlv_formula(psf, bld):
    gdv = f"({A['nsa']}*{psf}/1000)"; cb = f"({A['gfa']}*{bld}/1000)"; cs = f"({cb}*{A['soft']})"; cm = f"({gdv}*{A['mkt']})"
    cf = f"(MAX(0,({cb}+{cs})-{gdv}*{A['presale']}*{A['fin_draw']})*{A['fin_rate']}*{A['months']}/12*{A['fin_draw']})"
    ct = f"({cb}+{cs}+{cm}+{cf})"
    return f"=({gdv}/(1+{A['profit_cost']})-{ct})*1000/{A['gfa']}"
g2 = grid("RESIDUAL LAND VALUE PER SQ FT OF GFA (AED) \u2014 SALE PRICE (rows) versus CONSTRUCTION COST (columns)",
          PSF, BLD, NUM0, NUM0, lambda r, c: rlv_formula(r, c), (2400, 700), fmt=NUM0)

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sy = S("Summary", ncols=6, label_width=62, subtitle="Emaar on one page")
sy.ws.column_dimensions["F"].width = 56
sy.section("NET ASSET VALUE (AED millions)")
sy.row("A. UAE development, Emaar's share", f"={N['a_share']}", NUM0, font=F_LINK)
sy.row("B. International development", f"={N['b']}", NUM0, font=F_LINK)
sy.row("C. Malls and commercial leasing", f"={N['c']}", NUM0, font=F_LINK)
sy.row("D. Hospitality, leisure and entertainment", f"={N['d']}", NUM0, font=F_LINK)
sy.row("Gross asset value", f"={N['gav']}", NUM0, font=F_LINK, bold=True, border=TOTAL_BORDER)
sy.row("Free cash less borrowings", f"={N['cash_free']}+{N['debt']}", NUM0, font=F_LINK)
sy.row("Net asset value", f"={N['nav']}", NUM0, font=F_LINK, bold=True, border=DOUBLE_BORDER)
sy.blank(); sy.section("AGAINST THE MARKET")
sy.row("NAV per share (AED)", f"={N['navps']}", NUM2, font=F_LINK, bold=True)
sy.row("Share price (AED)", f"={N['px']}", NUM2, font=F_LINK)
sy.row("Discount to NAV", f"={N['disc']}", PCT, font=F_LINK, bold=True)
sy.row("Price / book", f"={I['pb']}", MULT, font=F_LINK)
sy.row("Market cap as % of recurring assets plus free cash", f"={N['impl_disc']}", PCT, font=F_LINK, bold=True)
sy.blank(); sy.section("THE DEVELOPMENT APPRAISAL")
sy.row("Gross development value of the illustrative tower", f"={A['gdv']}", NUM0, font=F_LINK)
sy.row("Residual land value", f"={A['rlv']}", NUM0, font=F_LINK, bold=True)
sy.row("  per sq ft of gross floor area (AED)", f"={A['rlv_gfa']}", NUM0, font=F_LINK)
sy.row("Price cushion before the scheme loses money", f"={A['cushion']}", PCT, font=F_LINK, border=DOUBLE_BORDER)
sy.blank()
sy.bullets([
    "Emaar trades at a wide discount to a bottom-up NAV, as UAE developers have for most of their listed history. The "
    "market capitalisation covers the recurring assets and the free cash and little more; the price ascribes almost no "
    "value to a backlog of AED 155bn and 660 million square feet of land.",
    "The two halves of the group deserve different multiples and get them. Dubai Mall and the leasing portfolio are "
    "capitalised at a 7% yield, roughly 14x EBITDA; the development backlog is discounted at 11% over three and a half "
    "years and taxed. The recurring assets are about half of gross value.",
    "Half the cash is treated as belonging to buyers, not shareholders. Off-plan collections sit in escrow to fund "
    "construction; counting them as free cash is the most common way a developer NAV is overstated.",
    "The development appraisal shows why the discount exists. Land is a residual: at the base case it is close to "
    "two-fifths of GDV, and a 10% fall in sale prices removes about a fifth of the land value, twice the gearing. The "
    "market applies that gearing to the whole land bank.",
])

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
tests = [
    ("Backlog split sums to the group figure", f"=ABS({I['backlog_dev']}+{I['backlog_intl']}-{I['backlog_grp']})<0.01"),
    ("Mall EBITDA plus hospitality EBITDA equals disclosed recurring EBITDA", f"=ABS({I['mall_ebitda']}+{I['hosp_ebitda']}-{I['rec_ebitda']})<0.01"),
    ("Recurring EBITDA margin is plausible (60\u201385%)", f"=AND({I['rec_ebitda']}/{I['rec_rev']}>0.60,{I['rec_ebitda']}/{I['rec_rev']}<0.85)"),
    ("Development margin assumed is below Emaar Development's reported EBITDA margin", f"={I['dev_margin']}<0.52"),
    ("Discount factor lies between 0 and 1", f"=AND({N['a_df']}>0,{N['a_df']}<1)"),
    ("Gross asset value equals the sum of the four blocks", f"=ABS({N['gav']}-({N['a_share']}+{N['b']}+{N['c']}+{N['d']}))<0.01"),
    ("NAV per share reconciles to NAV over shares", f"=ABS({N['navps']}-{N['nav']}/{I['shares']})<0.001"),
    ("Block A is within 40% of the listed Emaar Development value", f"=ABS({N['a_vs']})<0.40"),
    ("Base NAV sensitivity cell reconciles to the NAV sheet", f"=ABS(Sensitivity!E{g1+2}-{N['navps']})<0.01"),
    ("Base appraisal sensitivity cell reconciles to the residual land value", f"=ABS(Sensitivity!E{g2+2}-{A['rlv_gfa']})<0.5"),
    ("Developer profit equals the target share of total cost including land", f"=ABS({A['check_pc']}-{A['profit_cost']})<0.0001"),
    ("Residual land value is positive", f"={A['rlv']}>0"),
    ("GDV exceeds total cost plus land plus profit by less than a rounding error", f"=ABS({A['gdv']}-({A['c_tot']}+{A['rlv']}+{A['profit']}))<0.01"),
    ("Market capitalisation reconciles to price times shares", f"=ABS({I['mcap']}-{I['price']}*{I['shares']})<0.01"),
    ("Share price lies within the 52-week range", f"=AND({I['price']}>={E['w52lo']}-0.1,{I['price']}<={E['w52hi']}+0.1)"),
]
checks_sheet(book, tests,
             "The eighth check compares the model's own value for the UAE development block with what the Dubai market pays "
             "for the listed subsidiary that owns it. It is the only external test a NAV can be given, and it is deliberately loose.")

# --------------------------------------------------------------------------- #
book.cover(
    blurb="A net-asset-value model of Emaar Properties on the FY2025 results and 18 September 2026 market data, valuing "
          "the UAE and international development businesses on their disclosed backlog and land bank and the mall and "
          "hospitality portfolios at capitalisation yields, then measuring the discount the market applies. A residual-value "
          "development appraisal for one illustrative Dubai tower shows why land is the most volatile line in the NAV.",
    method=[
        "Development valued as the present value of the margin embedded in the sold-but-unrecognised backlog, after tax, "
        "plus the land bank at a stated value per square foot; Emaar's 80% share of the UAE business taken, and "
        "cross-checked against the listed Emaar Development price.",
        "Malls and hospitality capitalised at stated yields; half of cash treated as customer advances held for construction.",
        "A residual land valuation: gross development value less construction, soft costs, marketing and finance, less a "
        "target profit solved on total cost including land.",
        "Two sensitivity grids and fifteen checks, including an external test of block A against the market.",
    ],
    toc=[("Summary", "NAV, discount, the appraisal"),
         ("NAV", "four asset blocks, cash haircut, NAV per share and the discount"),
         ("Development Appraisal", "residual land value for one tower, break-even and cushion"),
         ("Sensitivity", "NAV against yield and margin; land value against price and cost"),
         ("Inputs", "disclosed operating data, balance sheet, market data, valuation assumptions"),
         ("Checks", "fifteen tests; must read MODEL OK")],
    highlights=[("Net asset value (AED m)", f"={N['nav']}", NUM0),
                ("NAV per share (AED)", f"={N['navps']}", NUM2),
                ("Share price (AED)", f"={N['px']}", NUM2),
                ("Discount to NAV", f"={N['disc']}", PCT),
                ("Residual land value per sq ft GFA (AED)", f"={A['rlv_gfa']}", NUM0),
                ("Price cushion on the scheme", f"={A['cushion']}", PCT)],
    sources=["Emaar Properties PJSC, FY2025 results announcement (12 February 2026): revenue backlog, recurring revenue and EBITDA, hospitality revenue, land bank; 9M 2025 release for the mall and hospitality split.",
             "Emaar Properties and Emaar Development FY2025 balance sheets via Yahoo Finance (EMAAR.AE, EMAARDEV.AE). Share prices: DFM close, 18 September 2026.",
             "Yields, development margins, discount rate, land values per square foot, the cash haircut and every input in the development appraisal are stated assumptions on the Inputs sheet.",
             "Nothing here is a recommendation. The valuation is an analytical exercise on public information."])

book.finish()
path = os.path.join(HERE, "Emaar_NAV.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
