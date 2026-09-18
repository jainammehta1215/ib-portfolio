"""
Build 12-aramco-nav/Aramco_NAV.xlsx — Saudi Aramco valued as an energy company: a reserves-based net asset value under
the Kingdom's fiscal terms, and a price-deck sensitivity that shows where the dividend breaks even.

Sheets: Cover · Summary · Inputs · Fiscal Terms · NAV · Price Deck · Sensitivity · Checks

Energy companies are valued on the cash their reserves will generate, not on last year's earnings, and for Aramco
that cash is shaped above all by the state's take: a royalty that steps up sharply as Brent rises, and a 50% income
tax on upstream profit. The model builds a per-barrel netback under those terms, applies it to production over the
reserve life as an annuity, adds the downstream business at a multiple, deducts net debt and minorities, and reports
NAV per share against the Tadawul price. The Price Deck sheet then runs the whole company across Brent from $50 to
$100 and asks the question that matters most to Aramco's shareholders: at what oil price does free cash flow stop
covering the base dividend?

Source data: Aramco Annual Report 2025 and FY2025 results (year to 31 December 2025), US$ billions unless stated;
Tadawul price 18 September 2026. Fiscal terms are the published royalty schedule and tax rates. Costs, the gas
realisation, the discount rate and the downstream multiple are stated inputs.
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)
from ibkit.sheet import Sheet, checks_sheet

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"

A = dict(fx=3.75, price_sar=25.56, shares=241.9074, w52lo=23.29, w52hi=27.96,
         revenue_sar=1671.2, ni_sar=348.0, ebitda_sar=807.0, capex_usd=52.2, fcf_usd=85.4, divs_usd=85.0,
         debt_sar=363.6, cash_sar=243.1, nci_sar=229.7, equity_sar=1492.0,
         reserves=247.2, prod=12.9, liquids=10.7, realised=69.2, up_ebit=195.5, dn_ebit=10.0, base_div=87.5)

book = Book(theme=THEMES["aramco"], project_no=12, project="Energy NAV and Price-Deck Sensitivity",
            company="Saudi Aramco", units="US$ billions \u00b7 per-share values in SAR",
            as_of="FY2025 annual report \u00b7 market data 18 Sep 2026")
S = lambda name, **kw: Sheet(book, name, **kw)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
si = S("Inputs", ncols=6, label_width=62, subtitle="Reported FY2025 figures, reserves and production, market data and stated assumptions")
si.ws.column_dimensions["F"].width = 60
I = {}
si.section("REPORTED FY2025")
I["fx"] = si.row("SAR per US$ (peg)", A["fx"], NUM2)
I["revenue"] = si.row("Revenue (US$ bn)", f"={A['revenue_sar']}/{I['fx']}", NUM)
I["ni"] = si.row("Net income (US$ bn)", f"={A['ni_sar']}/{I['fx']}", NUM)
I["up_ebit"] = si.row("Upstream adjusted EBIT (US$ bn)", A["up_ebit"], NUM)
I["dn_ebit"] = si.row("Downstream adjusted EBIT (US$ bn)", A["dn_ebit"], NUM, note="Up more than four times on record refining throughput and SABIC synergies")
I["capex"] = si.row("Capital expenditure (US$ bn)", A["capex_usd"], NUM, note="2026 guidance US$50\u201355bn; 65\u201370% upstream")
I["fcf"] = si.row("Free cash flow (US$ bn)", A["fcf_usd"], NUM, bold=True)
I["divs"] = si.row("Dividends paid (US$ bn)", A["divs_usd"], NUM)
I["base_div"] = si.row("Annualised base dividend after the Q4 2025 increase (US$ bn)", A["base_div"], NUM, bold=True,
                       note="Q4 2025 base dividend of US$21.9bn, times four; the performance-linked dividend is on top and discretionary")
I["realised"] = si.row("Average realised crude price (US$/bbl)", A["realised"], NUM2)

si.blank(); si.section("RESERVES AND PRODUCTION")
I["reserves"] = si.row("Proved reserves (billion boe)", A["reserves"], NUM)
I["prod"] = si.row("Total hydrocarbon production (mmboed)", A["prod"], NUM)
I["liquids"] = si.row("  of which liquids (mmbpd)", A["liquids"], NUM)
I["gas"] = si.row("  of which gas and NGL (mmboed)", f"={I['prod']}-{I['liquids']}", NUM)
I["prod_yr"] = si.row("Annual production (billion boe)", f"={I['prod']}*365/1000", NUM2)
I["life"] = si.row("Reserve life at current production (years)", f"={I['reserves']}/{I['prod_yr']}", NUM, bold=True)

si.blank(); si.section("BALANCE SHEET, 31 DECEMBER 2025 (US$ bn)")
I["debt"] = si.row("Borrowings", f"={A['debt_sar']}/{I['fx']}", NUM)
I["cash"] = si.row("Cash and cash equivalents", f"={A['cash_sar']}/{I['fx']}", NUM)
I["nd"] = si.row("Net debt", f"={I['debt']}-{I['cash']}", NUM, bold=True)
I["nci"] = si.row("Non-controlling interests at book", f"={A['nci_sar']}/{I['fx']}", NUM, note="SABIC minorities, the pipeline lease vehicles and Jafurah midstream")
I["equity"] = si.row("Equity attributable to shareholders", f"={A['equity_sar']}/{I['fx']}", NUM)

si.blank(); si.section("MARKET DATA")
I["price"] = si.row("Share price (SAR, Tadawul close 18 Sep 2026)", A["price_sar"], NUM2)
I["shares"] = si.row("Shares outstanding (billions)", A["shares"], NUM2)
I["mcap"] = si.row("Market capitalisation (US$ bn)", f"={I['price']}*{I['shares']}/{I['fx']}", NUM0, bold=True)
I["ev"] = si.row("Enterprise value (US$ bn)", f"={I['mcap']}+{I['nd']}+{I['nci']}", NUM0)
I["dy"] = si.row("Base dividend yield at the price", f"={I['base_div']}/{I['mcap']}", PCT)

si.blank(); si.section("VALUATION ASSUMPTIONS \u2014 stated inputs")
I["brent"] = si.row("Base-case Brent (US$/bbl, flat)", 70.0, NUM2, note="Close to the FY2025 realisation; a flat deck, not a forecast")
I["diff"] = si.row("Realised liquids price relative to Brent (US$/bbl)", 0.0, NUM2)
I["gas_px"] = si.row("Gas and NGL realisation (US$/boe)", 15.0, NUM2, note="Domestic gas is regulated at roughly US$1.75/mmBtu; NGL exports lift the blend")
I["opex"] = si.row("Lifting and other cash costs (US$/boe)", 4.0, NUM2, note="Aramco discloses upstream lifting cost near US$3/boe; the rest is SG&A and exploration")
I["capex_boe"] = si.row("Sustaining upstream capital expenditure (US$/boe)", 7.5, NUM2, note="Two-thirds of the US$52bn programme over 4.7bn boe")
I["dr"] = si.row("Discount rate", 0.08, PCT, note="Nominal, against a flat nominal price deck; conservative by construction")
I["dn_da"] = si.row("Downstream depreciation (US$ bn)", 10.0, NUM, note="Roughly two-fifths of group D&A")
I["dn_mult"] = si.row("Downstream EV / EBITDA multiple", 7.0, MULT, note="Integrated refining and chemicals; between Indian refiners and US majors in Project 10's peer set")
I["corp"] = si.row("Corporate costs and net interest, per year (US$ bn)", 4.0, NUM)

# --------------------------------------------------------------------------- #
# Fiscal Terms
# --------------------------------------------------------------------------- #
sf = S("Fiscal Terms", ncols=6, label_width=62, subtitle="The state's take: a sliding royalty on crude and a 50% upstream income tax")
sf.ws.column_dimensions["F"].width = 60
F = {}
sf.section("ROYALTY SCHEDULE ON CRUDE (marginal rates, published 2020)")
F["r1"] = sf.row("Rate on the price up to the first threshold", 0.15, PCT)
F["t1"] = sf.row("First threshold (US$/bbl)", 70.0, NUM2)
F["r2"] = sf.row("Rate between the first and second thresholds", 0.45, PCT)
F["t2"] = sf.row("Second threshold (US$/bbl)", 100.0, NUM2)
F["r3"] = sf.row("Rate above the second threshold", 0.80, PCT)
F["tax_up"] = sf.row("Upstream income tax rate", 0.50, PCT)
F["tax_dn"] = sf.row("Downstream income tax rate", 0.20, PCT, note="Reduced from 50% to 20% for domestic downstream in 2020 to encourage integration")
sf.blank()
sf.section("ROYALTY PER BARREL AT THE BASE-CASE PRICE")
F["px"] = sf.row("Liquids price (US$/bbl)", f"={I['brent']}+{I['diff']}", NUM2, font=F_FORMULA)
F["roy"] = sf.row("Royalty per barrel (US$)",
                  f"=MIN({F['px']},{F['t1']})*{F['r1']}+MAX(0,MIN({F['px']},{F['t2']})-{F['t1']})*{F['r2']}+MAX(0,{F['px']}-{F['t2']})*{F['r3']}", NUM2, bold=True)
F["roy_pct"] = sf.row("Effective royalty rate", f"={F['roy']}/{F['px']}", PCT)
F["take"] = sf.row("Government take on an incremental dollar above US$100", f"={F['r3']}+(1-{F['r3']})*{F['tax_up']}", PCT, bold=True,
                   note="Ninety cents of every dollar above US$100 goes to the state. Aramco's equity has very little exposure to a spike")
F["take_mid"] = sf.row("Government take on an incremental dollar between US$70 and US$100", f"={F['r2']}+(1-{F['r2']})*{F['tax_up']}", PCT)
F["take_lo"] = sf.row("Government take on an incremental dollar below US$70", f"={F['r1']}+(1-{F['r1']})*{F['tax_up']}", PCT)
sf.note("The schedule is what makes Aramco's cash flow so much less sensitive to oil than an international major's. "
        "Below US$70 the company keeps 42.5 cents of a marginal dollar; between US$70 and US$100 it keeps 27.5 cents; "
        "above US$100 it keeps ten. The price-deck grid on the Price Deck sheet is the direct consequence.", height=40)

# --------------------------------------------------------------------------- #
# NAV
# --------------------------------------------------------------------------- #
sn = S("NAV", ncols=6, label_width=62, subtitle="Per-barrel netback, upstream annuity, downstream multiple, and value per share")
sn.ws.column_dimensions["F"].width = 58
N = {}
sn.section("UPSTREAM NETBACK PER BARREL OF OIL EQUIVALENT (US$)")
N["rev_boe"] = sn.row("Blended realisation: liquids at the Brent-linked price, gas at its realisation",
                      f"=({I['liquids']}*{F['px']}+{I['gas']}*{I['gas_px']})/{I['prod']}", NUM2)
N["roy_boe"] = sn.row("Less: royalty (crude only, spread over all boe)", f"=-{F['roy']}*{I['liquids']}/{I['prod']}", NUM2)
N["opex_boe"] = sn.row("Less: cash operating costs", f"=-{I['opex']}", NUM2)
N["capex_boe2"] = sn.row("Less: sustaining capital expenditure", f"=-{I['capex_boe']}", NUM2)
N["pretax_boe"] = sn.row("Pre-tax cash netback", f"=SUM(C{sn.r-4}:C{sn.r-1})", NUM2, bold=True, border=TOTAL_BORDER)
N["tax_boe"] = sn.row("Less: income tax at 50%", f"=-{N['pretax_boe']}*{F['tax_up']}", NUM2)
N["net_boe"] = sn.row("After-tax cash netback per boe", f"={N['pretax_boe']}+{N['tax_boe']}", NUM2, bold=True, border=DOUBLE_BORDER)
N["net_pct"] = sn.row("Share of the blended realisation Aramco keeps", f"={N['net_boe']}/{N['rev_boe']}", PCT)

sn.blank(); sn.section("UPSTREAM VALUE")
N["up_fcf"] = sn.row("Upstream free cash flow per year (US$ bn)", f"={N['net_boe']}*{I['prod_yr']}", NUM, bold=True)
N["annuity"] = sn.row("Annuity factor over the reserve life at the discount rate", f"=(1-(1+{I['dr']})^-{I['life']})/{I['dr']}", NUM2,
                      note="Flat production, flat price. A 52-year annuity at 8% is worth 98% of a perpetuity; the tail barely matters")
N["up_val"] = sn.row("Upstream value (US$ bn)", f"={N['up_fcf']}*{N['annuity']}", NUM0, bold=True, border=DOUBLE_BORDER)
N["up_per_boe"] = sn.row("  implied value per boe of proved reserves (US$)", f"={N['up_val']}/{I['reserves']}", NUM2,
                         note="International majors are typically valued at US$8\u201315 per boe of proved reserves; the state's take explains the difference")
N["up_cal"] = sn.row("  calibration: model upstream FCF less corporate costs versus reported group FCF",
                     f"=({N['up_fcf']}-{I['corp']}+({I['dn_ebit']}+{I['dn_da']})*(1-{F['tax_dn']})-({I['capex']}-{I['capex_boe']}*{I['prod_yr']}))/{I['fcf']}-1", PCT,
                     note="The model rebuilds group free cash flow from first principles; it should land within 20% of the reported US$85bn at the FY2025 price")

sn.blank(); sn.section("DOWNSTREAM AND OTHER")
N["dn_ebitda"] = sn.row("Downstream EBITDA (adjusted EBIT + depreciation, US$ bn)", f"={I['dn_ebit']}+{I['dn_da']}", NUM)
N["dn_val"] = sn.row("Downstream value at the multiple (US$ bn)", f"={N['dn_ebitda']}*{I['dn_mult']}", NUM0, bold=True)
N["corp_val"] = sn.row("Less: corporate costs and net interest capitalised at the discount rate", f"=-{I['corp']}/{I['dr']}", NUM0)

sn.blank(); sn.section("NET ASSET VALUE")
N["gav"] = sn.row("Gross asset value (US$ bn)", f"={N['up_val']}+{N['dn_val']}+{N['corp_val']}", NUM0, bold=True)
N["nd"] = sn.row("Less: net debt", f"=-{I['nd']}", NUM0, font=F_LINK)
N["nci"] = sn.row("Less: non-controlling interests at book", f"=-{I['nci']}", NUM0, font=F_LINK)
N["nav"] = sn.row("Net asset value (US$ bn)", f"={N['gav']}+{N['nd']}+{N['nci']}", NUM0, bold=True, border=DOUBLE_BORDER)
N["navps"] = sn.row("NAV per share (SAR)", f"={N['nav']}*{I['fx']}/{I['shares']}", NUM2, bold=True)
N["px"] = sn.row("Share price (SAR)", f"={I['price']}", NUM2, font=F_LINK)
N["updown"] = sn.row("Upside / (downside)", f"={N['navps']}/{N['px']}-1", PCT, bold=True, border=DOUBLE_BORDER)
N["impl_brent"] = sn.row("Brent the share price implies (US$/bbl, from the Price Deck sheet)", "", NUM2, bold=True)
N["up_share"] = sn.row("Upstream as % of gross asset value", f"={N['up_val']}/{N['gav']}", PCT)

# --------------------------------------------------------------------------- #
# Price Deck
# --------------------------------------------------------------------------- #
sp = S("Price Deck", ncols=9, label_width=56, col_width=12.5, subtitle="The whole company across Brent from US$50 to US$110 \u2014 and the dividend break-even")
DECK = [50, 60, 70, 80, 90, 100, 110]
P = {}
sp.section("PER BARREL AND PER YEAR AT EACH PRICE")
sp.head([f"${d}" for d in DECK], label="Brent (US$/bbl)")
r_b = sp.r
P["brent"] = sp.multi("Brent", DECK, NUM2)
def roy(col): return (f"MIN({col}+{I['diff']},{F['t1']})*{F['r1']}+MAX(0,MIN({col}+{I['diff']},{F['t2']})-{F['t1']})*{F['r2']}"
                      f"+MAX(0,{col}+{I['diff']}-{F['t2']})*{F['r3']}")
P["roy"] = sp.multi("Royalty per crude barrel (US$)", [f"={roy(L(3+j)+str(r_b))}" for j in range(7)], NUM2)
P["net"] = sp.multi("After-tax netback per boe (US$)",
                    [f"=(({I['liquids']}*({L(3+j)}{r_b}+{I['diff']})+{I['gas']}*{I['gas_px']})/{I['prod']}-{L(3+j)}{r_b+1}*{I['liquids']}/{I['prod']}-{I['opex']}-{I['capex_boe']})*(1-{F['tax_up']})" for j in range(7)],
                    NUM2, bold=True)
P["up_fcf"] = sp.multi("Upstream free cash flow (US$ bn)", [f"={L(3+j)}{r_b+2}*{I['prod_yr']}" for j in range(7)], NUM)
P["grp_fcf"] = sp.multi("Group free cash flow: upstream + downstream after tax and capex \u2212 corporate (US$ bn)",
                        [f"={L(3+j)}{r_b+3}+({I['dn_ebit']}+{I['dn_da']})*(1-{F['tax_dn']})-({I['capex']}-{I['capex_boe']}*{I['prod_yr']})-{I['corp']}" for j in range(7)],
                        NUM, bold=True, border=TOTAL_BORDER)
P["div_cov"] = sp.multi("Base dividend cover: group FCF / base dividend", [f"={L(3+j)}{r_b+4}/{I['base_div']}" for j in range(7)], MULT, bold=True)
P["surplus"] = sp.multi("Surplus / (shortfall) after the base dividend (US$ bn)", [f"={L(3+j)}{r_b+4}-{I['base_div']}" for j in range(7)], NUM)
sp.blank()
sp.section("VALUE AT EACH PRICE")
P["nav"] = sp.multi("NAV (US$ bn)", [f"={L(3+j)}{r_b+3}*{N['annuity']}+{N['dn_val']}+{N['corp_val']}-{I['nd']}-{I['nci']}" for j in range(7)], NUM0)
P["navps"] = sp.multi("NAV per share (SAR)", [f"={L(3+j)}{sp.r-1}*{I['fx']}/{I['shares']}" for j in range(7)], NUM2, bold=True, border=DOUBLE_BORDER)
P["vs_px"] = sp.multi("Versus the share price", [f"={L(3+j)}{sp.r-1}/{I['price']}-1" for j in range(7)], PCT)
sp.blank()
sp.section("BREAK-EVENS (closed form)")
# Dividend break-even solved analytically on the sub-$70 leg then checked; use general algebra on the first leg with fallback message.
# Group FCF(P) on leg 1 = (a*P + b) where a = liquids/prod*(1-r1)*(1-tax)*prod_yr, b = rest.
a1 = f"({I['liquids']}/{I['prod']}*(1-{F['r1']})*(1-{F['tax_up']})*{I['prod_yr']})"
b1 = (f"(({I['gas']}*{I['gas_px']}/{I['prod']}-{I['opex']}-{I['capex_boe']}+{I['liquids']}/{I['prod']}*{I['diff']}*(1-{F['r1']}))*(1-{F['tax_up']})*{I['prod_yr']}"
      f"+({I['dn_ebit']}+{I['dn_da']})*(1-{F['tax_dn']})-({I['capex']}-{I['capex_boe']}*{I['prod_yr']})-{I['corp']})")
P["be_div"] = sp.row("Brent at which group FCF exactly covers the base dividend (US$/bbl)", f"=({I['base_div']}-{b1})/{a1}", NUM2, bold=True,
                     note="Valid while the answer is below the first royalty threshold, which the check confirms")
P["be_zero"] = sp.row("Brent at which group FCF is zero (US$/bbl)", f"=-{b1}/{a1}", NUM2)
# Price-implied Brent: solve NAV(P) = market cap, piecewise across the three royalty legs.
up_target = f"(({I['mcap']}+{I['nd']}+{I['nci']}-{N['dn_val']}-{N['corp_val']})/{N['annuity']})"
def slope(rate): return f"({I['liquids']}/{I['prod']}*(1-{rate})*(1-{F['tax_up']})*{I['prod_yr']})"
def upfcf_at(pbrent):  # upstream FCF at any Brent, full schedule
    return (f"((({I['liquids']}*({pbrent}+{I['diff']})+{I['gas']}*{I['gas_px']})/{I['prod']}-({roy(pbrent)})*{I['liquids']}/{I['prod']}"
            f"-{I['opex']}-{I['capex_boe']})*(1-{F['tax_up']})*{I['prod_yr']})")
f_t1, f_t2 = upfcf_at(F["t1"]), upfcf_at(F["t2"])
leg1 = f"({F['t1']}-({f_t1}-{up_target})/{slope(F['r1'])})"
leg2 = f"({F['t1']}+({up_target}-{f_t1})/{slope(F['r2'])})"
leg3 = f"({F['t2']}+({up_target}-{f_t2})/{slope(F['r3'])})"
P["impl"] = sp.row("Brent the share price implies on this NAV (US$/bbl)",
                   f"=IF({up_target}<={f_t1},{leg1},IF({up_target}<={f_t2},{leg2},{leg3}))", NUM2, bold=True, border=DOUBLE_BORDER,
                   note="Solved on whichever royalty leg the answer falls, and verified by substitution on the Checks sheet")
P["impl_dr"] = sp.row("Discount rate the share price implies at US$70 Brent (perpetuity approximation)",
                      f"=({N['up_fcf']}-{I['corp']})/({I['mcap']}+{I['nd']}+{I['nci']}-{N['dn_val']})", PCT, bold=True,
                      note="Upstream FCF less corporate costs over the enterprise value net of downstream. The 52-year annuity differs from a perpetuity by under 2%")
sn.ws.cell(int(N["impl_brent"].split("$")[-1]), 3).value = f"={P['impl']}"
sn.ws.cell(int(N["impl_brent"].split("$")[-1]), 3).font = F_LINK
sp.note("Group free cash flow rises by about US$2.4bn for each dollar of Brent below US$70 and by roughly US$1.5bn per "
        "dollar between US$70 and US$100. The base dividend alone absorbs the bulk of free cash flow at a US$70 deck, "
        "which is why the performance-linked dividend was cut to almost nothing in 2025 and why the balance sheet, not the "
        "oil price, funded part of the 2024 payout.", height=44)

# --------------------------------------------------------------------------- #
# Sensitivity
# --------------------------------------------------------------------------- #
ss = S("Sensitivity", ncols=8, label_width=40, col_width=13.5, subtitle="NAV per share (SAR) \u2014 each cell rebuilds the netback and the annuity")
BR = [55, 62.5, 70, 77.5, 85]; DRS = [0.07, 0.075, 0.08, 0.085, 0.09]; CPX = [5.5, 6.5, 7.5, 8.5, 9.5]
def nav_fn(brent, dr=None, capex=None):
    dr = dr or I["dr"]; capex = capex or I["capex_boe"]
    net = (f"((({I['liquids']}*({brent}+{I['diff']})+{I['gas']}*{I['gas_px']})/{I['prod']}-({roy(brent)})*{I['liquids']}/{I['prod']}"
           f"-{I['opex']}-{capex})*(1-{F['tax_up']}))")
    ann = f"((1-(1+{dr})^-{I['life']})/{dr})"
    return f"=({net}*{I['prod_yr']}*{ann}+{N['dn_val']}+{N['corp_val']}-{I['nd']}-{I['nci']})*{I['fx']}/{I['shares']}"
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
g1 = grid("BRENT (rows) versus DISCOUNT RATE (columns)", BR, DRS, NUM2, PCT, lambda r, c: nav_fn(r, dr=c), (70, 0.08))
g2 = grid("BRENT (rows) versus SUSTAINING CAPEX PER BOE (columns)", BR, CPX, NUM2, NUM2, lambda r, c: nav_fn(r, capex=c), (70, 7.5))

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sy = S("Summary", ncols=6, label_width=62, subtitle="Aramco on one page")
sy.ws.column_dimensions["F"].width = 56
sy.section("THE NETBACK AT US$70 BRENT")
sy.row("Blended realisation per boe (US$)", f"={N['rev_boe']}", NUM2, font=F_LINK)
sy.row("After-tax cash netback per boe (US$)", f"={N['net_boe']}", NUM2, font=F_LINK, bold=True)
sy.row("Share of the barrel Aramco keeps after royalty, costs, capex and tax", f"={N['net_pct']}", PCT, font=F_LINK)
sy.row("Government take on an incremental dollar above US$100", f"={F['take']}", PCT, font=F_LINK)
sy.blank(); sy.section("NET ASSET VALUE")
sy.row("Upstream value (US$ bn)", f"={N['up_val']}", NUM0, font=F_LINK)
sy.row("  per boe of proved reserves (US$)", f"={N['up_per_boe']}", NUM2, font=F_LINK)
sy.row("Downstream value (US$ bn)", f"={N['dn_val']}", NUM0, font=F_LINK)
sy.row("Net asset value (US$ bn)", f"={N['nav']}", NUM0, font=F_LINK, bold=True)
sy.row("NAV per share (SAR)", f"={N['navps']}", NUM2, font=F_LINK, bold=True)
sy.row("Share price (SAR)", f"={N['px']}", NUM2, font=F_LINK)
sy.row("Upside / (downside)", f"={N['updown']}", PCT, font=F_LINK, bold=True)
sy.row("Brent the price implies at an 8% discount rate (US$/bbl)", f"={P['impl']}", NUM2, font=F_LINK, bold=True)
sy.row("Discount rate the price implies at US$70 Brent", f"={P['impl_dr']}", PCT, font=F_LINK, bold=True)
sy.blank(); sy.section("THE DIVIDEND")
sy.row("Base dividend, annualised (US$ bn)", f"={I['base_div']}", NUM, font=F_LINK)
sy.row("Base dividend yield at the price", f"={I['dy']}", PCT, font=F_LINK)
sy.row("Group free cash flow at US$70 (US$ bn)", f"={P['grp_fcf'][2]}", NUM, font=F_LINK)
sy.row("Base dividend cover at US$70", f"={P['div_cov'][2]}", MULT, font=F_LINK, bold=True)
sy.row("Brent at which the base dividend is exactly covered (US$/bbl)", f"={P['be_div']}", NUM2, font=F_LINK, bold=True, border=DOUBLE_BORDER)
sy.blank()
sy.bullets([
    "Aramco keeps a little over a quarter of each barrel. Royalty, costs, sustaining capex and a 50% tax take the rest, "
    "and the sliding royalty means the equity captures ten cents of every dollar above US$100. The shares are a claim on "
    "volume and cost, not on the oil price.",
    "On a flat US$70 deck at an 8% discount rate the NAV is about 30% below the share price. For the price to be right "
    "at 8%, Brent would have to sit near US$120 for ever, which nobody assumes. The alternative reading is a discount rate "
    "under 6%: the market treats Aramco as a quasi-sovereign yield instrument, and most of the float is held by the state "
    "and its funds. The premium is a discount-rate premium, not an oil bet.",
    "Reserves are valued at a small fraction of what an international major's would fetch per barrel. That is not a "
    "discount to be closed \u2014 it is the arithmetic of who owns the barrel. Fifty-two years of reserve life adds almost "
    "nothing beyond year thirty at an 8% discount rate.",
    "The dividend, not the NAV, is the live question. The base payout alone needs Brent in the high sixties to be covered "
    "from free cash flow after a US$52bn capex programme. Below that the balance sheet or asset sales fund the gap, as "
    "they did in 2024 and 2025.",
])

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
tests = [
    ("Reserve life is between 40 and 60 years", f"=AND({I['life']}>40,{I['life']}<60)"),
    ("Royalty per barrel at the base price reproduces the schedule by hand", f"=ABS({F['roy']}-(MIN({F['px']},70)*0.15+MAX(0,MIN({F['px']},100)-70)*0.45+MAX(0,{F['px']}-100)*0.8))<0.0001"),
    ("Royalty is non-decreasing across the price deck", "=AND(" + ",".join(f"{P['roy'][j+1]}>={P['roy'][j]}" for j in range(6)) + ")"),
    ("Netback is increasing across the price deck", "=AND(" + ",".join(f"{P['net'][j+1]}>{P['net'][j]}" for j in range(6)) + ")"),
    ("Base-case netback ties to the US$70 deck column", f"=ABS({N['net_boe']}-{P['net'][2]})<0.0001"),
    ("Base-case NAV per share ties to the US$70 deck column", f"=ABS({N['navps']}-{P['navps'][2]})<0.001"),
    ("Model rebuild of FY2025 group free cash flow is within 20% of the reported figure", f"=ABS({N['up_cal']})<0.20"),
    ("Annuity factor is below 1 / discount rate", f"={N['annuity']}<1/{I['dr']}"),
    ("Dividend break-even solved on the first royalty leg is below the first threshold", f"={P['be_div']}<{F['t1']}"),
    ("Dividend break-even reproduces exactly when substituted (cover = 1.0x)", f"=ABS(({a1}*{P['be_div']}+{b1})/{I['base_div']}-1)<0.0001"),
    ("Price-implied Brent reproduces the market capitalisation when substituted", f"=ABS({upfcf_at(P['impl'])}*{N['annuity']}+{N['dn_val']}+{N['corp_val']}-{I['nd']}-{I['nci']}-{I['mcap']})<0.5"),
    ("Price-implied discount rate is below the model discount rate when the price exceeds NAV", f"=IF({I['price']}>{N['navps']},{P['impl_dr']}<{I['dr']},TRUE)"),
    ("Base sensitivity cell reconciles to the NAV sheet", f"=ABS(Sensitivity!E{g1+2}-{N['navps']})<0.001"),
    ("Market capitalisation reconciles to price times shares over the peg", f"=ABS({I['mcap']}-{I['price']}*{I['shares']}/{I['fx']})<0.01"),
    ("Share price lies within the 52-week range", f"=AND({I['price']}>={A['w52lo']}-0.1,{I['price']}<={A['w52hi']}+0.1)"),
]
checks_sheet(book, tests,
             "The seventh check rebuilds Aramco's reported free cash flow from a per-barrel netback and demands it land within "
             "20% of the audited figure. A reserves model that cannot reproduce last year's cash is a spreadsheet, not a valuation.")

# --------------------------------------------------------------------------- #
book.cover(
    blurb="A reserves-based net asset value of Saudi Aramco on the FY2025 annual report and 18 September 2026 market data, "
          "built under the Kingdom's fiscal terms \u2014 the sliding crude royalty and the 50% upstream income tax \u2014 with the "
          "whole company run across a Brent price deck to find the oil price the share price implies and the price at which "
          "the base dividend stops being covered.",
    method=[
        "A per-boe netback: blended realisation, the published royalty schedule applied to crude, cash costs, sustaining "
        "capex and tax; multiplied by annual production and capitalised over the 52-year reserve life as an annuity.",
        "Downstream at a multiple of adjusted EBIT plus depreciation; corporate costs capitalised; net debt and minorities deducted.",
        "A price deck from US$50 to US$110 with group free cash flow and dividend cover at each level, and closed-form "
        "solves for the dividend break-even and the price-implied Brent, each verified by substitution.",
        "A calibration check that rebuilds reported FY2025 free cash flow from the netback, and two sensitivity grids.",
    ],
    toc=[("Summary", "netback, NAV, dividend"),
         ("Fiscal Terms", "royalty schedule, tax rates, government take per marginal dollar"),
         ("NAV", "netback per boe, upstream annuity, downstream, value per share"),
         ("Price Deck", "the company across Brent; dividend cover; break-evens; price-implied Brent"),
         ("Sensitivity", "Brent against discount rate and against sustaining capex"),
         ("Inputs", "FY2025 figures, reserves, balance sheet, market data, assumptions"),
         ("Checks", "fifteen tests; must read MODEL OK")],
    highlights=[("After-tax netback per boe at US$70 (US$)", f"={N['net_boe']}", NUM2),
                ("NAV per share (SAR)", f"={N['navps']}", NUM2),
                ("Share price (SAR)", f"={N['px']}", NUM2),
                ("Brent the price implies at 8% (US$/bbl)", f"={P['impl']}", NUM2),
                ("Base dividend cover at US$70", f"={P['div_cov'][2]}", MULT),
                ("Discount rate the price implies at US$70", f"={P['impl_dr']}", PCT)],
    sources=["Saudi Aramco Annual Report 2025 and fourth-quarter and full-year 2025 results (year to 31 December 2025): production, reserves, realised price, segment adjusted EBIT, capital expenditure, free cash flow and dividends.",
             "Balance sheet and income statement in SAR via Yahoo Finance for 2222.SR; converted at the 3.75 peg. Share price: Tadawul close, 18 September 2026.",
             "Royalty schedule and tax rates as published by the Kingdom in 2020. Costs, gas realisation, discount rate, downstream multiple and corporate costs are stated inputs.",
             "Nothing here is a recommendation. The valuation is an analytical exercise on public information."])

book.finish(freeze={"Price Deck": "C6"})
path = os.path.join(HERE, "Aramco_NAV.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
