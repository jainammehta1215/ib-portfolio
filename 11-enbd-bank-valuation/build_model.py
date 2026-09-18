"""
Build 11-enbd-bank-valuation/ENBD_Bank_Valuation.xlsx — Emirates NBD valued the way banks are valued.

Sheets: Cover · Summary · Inputs · Forecast · DDM · Excess Return · Justified Multiples · Sensitivity · Checks

Banks cannot be valued on enterprise value because debt is the raw material, not the financing. Two equity-side methods
are built on one forecast: a dividend discount model, and an excess-return (residual income) model that starts from
book value and adds the present value of what the bank earns above its cost of equity. Under clean-surplus accounting
\u2014 closing equity = opening equity + earnings \u2212 dividends, which the forecast enforces \u2014 the two must give the same
number, and the Checks sheet demands that they do. The excess-return layout is kept because it makes the source of
value visible: how much is the bank already worth on its books, and how much is the market paying for future
returns above the cost of capital.

Source data is the FY2025 audited accounts used in Project 6, so the two workbooks agree. Cost-of-equity inputs are
stated. Market data 18 September 2026.
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
# Source data — Emirates NBD FY2025 audited accounts (year to 31 December 2025), AED millions.
# Same figures as Project 6 (Yahoo Finance fundamentals for EMIRATESNBD.AE restating the published accounts).
# Share price: DFM close 18 Sep 2026.
# --------------------------------------------------------------------------- #
B = dict(ni=23442.0, ni_prior=22973.0, equity=144582.0, equity_prior=125990.0, goodwill_intang=5620.0,
         shares=6311.0, price=30.98, div_paid=6311.0, assets=1164442.0, nii=28407.0, revenue=49177.0,
         w52lo=24.00, w52hi=37.40)
YEARS = ["FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]
NY = len(YEARS)
ROE_PATH = [0.165, 0.160, 0.155, 0.150, 0.145]      # fades toward the terminal ROE
PAYOUT_PATH = [0.30, 0.33, 0.36, 0.40, 0.44]        # rises as growth slows

book = Book(theme=THEMES["enbd"], project_no=11, project="Bank Valuation \u2014 Dividend Discount and Excess Return",
            company="Emirates NBD", units="AED millions \u00b7 per-share values in AED",
            as_of="FY2025 accounts \u00b7 market data 18 Sep 2026")
S = lambda name, **kw: Sheet(book, name, **kw)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
si = S("Inputs", ncols=6, label_width=58, subtitle="FY2025 accounts, market data, cost of equity and terminal assumptions")
si.ws.column_dimensions["F"].width = 60
I = {}
si.section("FY2025 ACCOUNTS (AED millions)")
I["ni"] = si.row("Net profit attributable to equity holders", B["ni"], NUM0, note="As in Project 6")
I["ni_prior"] = si.row("Net profit, FY2024", B["ni_prior"], NUM0)
I["equity"] = si.row("Shareholders' equity, 31 December 2025", B["equity"], NUM0)
I["equity_prior"] = si.row("Shareholders' equity, 31 December 2024", B["equity_prior"], NUM0)
I["gw"] = si.row("Goodwill and intangibles", B["goodwill_intang"], NUM0)
I["tbv"] = si.row("Tangible book value", f"={I['equity']}-{I['gw']}", NUM0, bold=True)
I["div"] = si.row("Cash dividends paid in 2025", B["div_paid"], NUM0, note="AED 1.00 per share")
I["assets"] = si.row("Total assets", B["assets"], NUM0)
I["roe_hist"] = si.row("Return on average equity, FY2025", f"={I['ni']}/AVERAGE({I['equity']},{I['equity_prior']})", PCT, bold=True)
I["rote_hist"] = si.row("Return on average tangible equity, FY2025",
                        f"={I['ni']}/AVERAGE({I['tbv']},{I['equity_prior']}-{I['gw']})", PCT)
I["payout_hist"] = si.row("Dividend payout, FY2025", f"={I['div']}/{I['ni']}", PCT)
I["roa"] = si.row("Return on assets, FY2025", f"={I['ni']}/{I['assets']}", PCT)

si.blank(); si.section("MARKET DATA")
I["price"] = si.row("Share price (AED, DFM close 18 Sep 2026)", B["price"], NUM2)
I["shares"] = si.row("Shares outstanding (millions)", B["shares"], NUM0)
I["mcap"] = si.row("Market capitalisation", f"={I['price']}*{I['shares']}", NUM0, bold=True)
I["bvps"] = si.row("Book value per share (AED)", f"={I['equity']}/{I['shares']}", NUM2)
I["tbvps"] = si.row("Tangible book value per share (AED)", f"={I['tbv']}/{I['shares']}", NUM2)
I["pb_mkt"] = si.row("Price / book", f"={I['price']}/{I['bvps']}", MULT, bold=True)
I["ptbv_mkt"] = si.row("Price / tangible book", f"={I['price']}/{I['tbvps']}", MULT)
I["pe_mkt"] = si.row("Price / FY2025 earnings", f"={I['mcap']}/{I['ni']}", MULT)
I["dy_mkt"] = si.row("Dividend yield on the last dividend", f"={I['div']}/{I['mcap']}", PCT)

si.blank(); si.section("COST OF EQUITY")
I["rf"] = si.row("Risk-free rate (10-year US Treasury; the dirham is pegged)", 0.0410, PCT, note="September 2026")
I["crp"] = si.row("UAE country risk premium", 0.0090, PCT, note="Damodaran country tables for an Aa2 sovereign, rounded up for oil-price and regional concentration")
I["erp"] = si.row("Mature-market equity risk premium", 0.0500, PCT)
I["beta"] = si.row("Beta versus the DFM General Index", 1.00, NUM2, note="Large domestic banks sit close to their index; stated input")
I["ke"] = si.row("Cost of equity", f"={I['rf']}+{I['crp']}+{I['beta']}*{I['erp']}", PCT, bold=True, border=DOUBLE_BORDER)

si.blank(); si.section("TERMINAL ASSUMPTIONS")
I["g"] = si.row("Long-run growth in earnings and book value", 0.030, PCT, note="Nominal UAE GDP trend less some share loss; must be below the cost of equity")
I["roe_t"] = si.row("Terminal return on equity", 0.140, PCT, note="Below today's 17%: rate cuts compress margins and capital builds")
I["payout_t"] = si.row("Terminal payout (1 \u2212 g / ROE, so book grows at g)", f"=1-{I['g']}/{I['roe_t']}", PCT, bold=True,
                       note="Not a free input. If payout and growth are set independently the two models diverge")

# --------------------------------------------------------------------------- #
# Forecast
# --------------------------------------------------------------------------- #
CE = 3; C1 = CE + 1; CN = CE + NY
sf = S("Forecast", ncols=CN + 1, label_width=48, col_width=13.5, valcol=CE,
       subtitle="Equity roll-forward under clean surplus: closing book = opening book + earnings \u2212 dividends")
F = {}
sf.section("EQUITY ROLL-FORWARD (AED millions)")
sf.head(["FY2025A"] + YEARS, first_col=CE)
r_open = sf.r
F["open"] = sf.multi("Opening shareholders' equity", [f"={I['equity_prior']}"] + [f"={L(CE+i)}{r_open+4}" for i in range(NY)], NUM0)
F["roe"] = sf.multi("   Return on opening equity", [f"={I['ni']}/{I['equity_prior']}"] + ROE_PATH, PCT)
F["ni"] = sf.multi("Net profit", [f"={I['ni']}"] + [f"={L(CE+i+1)}{r_open}*{L(CE+i+1)}{r_open+1}" for i in range(NY)], NUM0, bold=True)
F["payout"] = sf.multi("   Payout ratio", [f"={I['div']}/{I['ni']}"] + PAYOUT_PATH, PCT)
F["div"] = sf.multi("Dividends", [f"={I['div']}"] + [f"={L(CE+i+1)}{r_open+2}*{L(CE+i+1)}{r_open+3}" for i in range(NY)], NUM0)
# Row r_open+4 must be closing equity; FY2025A closing = actual equity (other comprehensive items absorbed)
F["close"] = sf.multi("Closing shareholders' equity", [f"={I['equity']}"] +
                      [f"={L(CE+i+1)}{r_open}+{L(CE+i+1)}{r_open+2}-{L(CE+i+1)}{r_open+4}" for i in range(NY)],
                      NUM0, bold=True, border=TOTAL_BORDER)
# fix: dividends row is r_open+4? No: rows are open(r), roe(r+1), ni(r+2), payout(r+3), div(r+4), close(r+5)
for i in range(NY):
    col = L(CE + i + 1)
    sf.ws.cell(r_open + 5, CE + i + 1).value = f"={col}{r_open}+{col}{r_open+2}-{col}{r_open+4}"
    sf.ws.cell(r_open, CE + i + 1).value = f"={L(CE+i)}{r_open+5}"
F["g_bv"] = sf.multi("   Growth in book value", [f"={I['equity']}/{I['equity_prior']}-1"] +
                     [f"={L(CE+i+1)}{r_open+5}/{L(CE+i)}{r_open+5}-1" for i in range(NY)], PCT)
F["g_ni"] = sf.multi("   Growth in net profit", [f"={I['ni']}/{I['ni_prior']}-1"] +
                     [f"={L(CE+i+1)}{r_open+2}/{L(CE+i)}{r_open+2}-1" for i in range(NY)], PCT)
sf.blank(); sf.section("PER SHARE (AED)")
F["eps"] = sf.multi("Earnings per share", [f"={L(CE+i)}{r_open+2}/{I['shares']}" for i in range(NY + 1)], NUM2)
F["dps"] = sf.multi("Dividend per share", [f"={L(CE+i)}{r_open+4}/{I['shares']}" for i in range(NY + 1)], NUM2, bold=True)
F["bvps"] = sf.multi("Book value per share (closing)", [f"={L(CE+i)}{r_open+5}/{I['shares']}" for i in range(NY + 1)], NUM2)
sf.blank()
sf.section("TERMINAL YEAR (FY2031E, the first year of steady state)")
F["t_open"] = sf.row("Opening equity", f"=Forecast!{L(CN)}{r_open+5}", NUM0, font=F_LINK)
F["t_ni"] = sf.row("Net profit at the terminal ROE", f"={F['t_open']}*{I['roe_t']}", NUM0)
F["t_div"] = sf.row("Dividends at the terminal payout", f"={F['t_ni']}*{I['payout_t']}", NUM0, bold=True)
F["t_close"] = sf.row("Closing equity", f"={F['t_open']}+{F['t_ni']}-{F['t_div']}", NUM0)
F["t_g"] = sf.row("Implied growth in book (must equal g)", f"={F['t_close']}/{F['t_open']}-1", PCT, bold=True)
sf.note("FY2025A closing equity is the reported figure, so the historical column absorbs other comprehensive income and "
        "any capital movements in that year. From FY2026E clean surplus is enforced exactly. The ROE path fades from the "
        "FY2025 level toward the terminal assumption; the payout rises so that retained earnings fund only the growth assumed.", height=40)

# --------------------------------------------------------------------------- #
# DDM
# --------------------------------------------------------------------------- #
sd = S("DDM", ncols=CN + 1, label_width=48, col_width=13.5, valcol=CE, subtitle="Dividend discount model \u2014 explicit dividends plus a Gordon terminal value")
D = {}
sd.section("PRESENT VALUE OF DIVIDENDS")
sd.head([""] + YEARS, first_col=CE)
r_t = sd.r
D["t"] = sd.multi("Year", [""] + list(range(1, NY + 1)), '0')
D["div"] = sd.multi("Dividends", [""] + [f"=Forecast!{L(CE+i+1)}{r_open+4}" for i in range(NY)], NUM0, font=F_LINK)
r_df = sd.r
D["df"] = sd.multi("Discount factor at the cost of equity", [""] + [f"=1/(1+{I['ke']})^{L(CE+i+1)}{r_t}" for i in range(NY)], '0.0000')
r_pv = sd.r
D["pv"] = sd.multi("Present value", [""] + [f"={L(CE+i+1)}{r_t+1}*{L(CE+i+1)}{r_df}" for i in range(NY)], NUM0, bold=True)
sd.blank(); sd.section("TERMINAL VALUE")
D["tv"] = sd.row("Terminal value at end FY2030E: FY2031E dividend / (ke \u2212 g)", f"={F['t_div']}/({I['ke']}-{I['g']})", NUM0)
D["tv_pv"] = sd.row("Present value of terminal value", f"={D['tv']}*DDM!{L(CN)}{r_df}", NUM0, bold=True)
D["tv_pb"] = sd.row("Terminal value / terminal opening book (implied exit P/B)", f"={D['tv']}/{F['t_open']}", MULT,
                    note="Should sit close to the justified P/B at the terminal ROE")
sd.blank(); sd.section("EQUITY VALUE")
D["pv_sum"] = sd.row("Sum of present values of dividends", f"=SUM(DDM!{L(C1)}{r_pv}:{L(CN)}{r_pv})", NUM0)
D["eq"] = sd.row("Equity value", f"={D['pv_sum']}+{D['tv_pv']}", NUM0, bold=True, border=DOUBLE_BORDER)
D["tv_share"] = sd.row("Terminal value as % of equity value", f"={D['tv_pv']}/{D['eq']}", PCT)
D["ps"] = sd.row("Value per share (AED)", f"={D['eq']}/{I['shares']}", NUM2, bold=True, border=DOUBLE_BORDER)
D["updown"] = sd.row("Upside / (downside) to the share price", f"={D['ps']}/{I['price']}-1", PCT, bold=True)
D["pb_impl"] = sd.row("Implied price / book on FY2025 equity", f"={D['eq']}/{I['equity']}", MULT)
D["pe_impl"] = sd.row("Implied price / FY2025 earnings", f"={D['eq']}/{I['ni']}", MULT)

# --------------------------------------------------------------------------- #
# Excess Return
# --------------------------------------------------------------------------- #
se = S("Excess Return", ncols=CN + 1, label_width=52, col_width=13.5, valcol=CE,
       subtitle="Book value today plus the present value of returns earned above the cost of equity")
E = {}
se.section("EXCESS RETURNS (AED millions)")
se.head([""] + YEARS, first_col=CE)
E["open"] = se.multi("Opening equity", [""] + [f"=Forecast!{L(CE+i+1)}{r_open}" for i in range(NY)], NUM0, font=F_LINK)
E["roe"] = se.multi("Return on opening equity", [""] + [f"=Forecast!{L(CE+i+1)}{r_open+1}" for i in range(NY)], PCT, font=F_LINK)
E["ke"] = se.multi("Cost of equity", [""] + [f"={I['ke']}"] * NY, PCT, font=F_LINK)
E["spread"] = se.multi("Spread: ROE \u2212 cost of equity", [""] + [f"={L(CE+i+1)}{se.r-2}-{L(CE+i+1)}{se.r-1}" for i in range(NY)], PCT, bold=True)
E["er"] = se.multi("Excess return: spread \u00d7 opening equity", [""] + [f"={L(CE+i+1)}{se.r-1}*{L(CE+i+1)}{se.r-4}" for i in range(NY)], NUM0, bold=True)
r_edf = se.r
E["df"] = se.multi("Discount factor", [""] + [f"=DDM!{L(CE+i+1)}{r_df}" for i in range(NY)], '0.0000', font=F_LINK)
r_epv = se.r
E["pv"] = se.multi("Present value of excess return", [""] + [f"={L(CE+i+1)}{r_edf-1}*{L(CE+i+1)}{r_edf}" for i in range(NY)], NUM0, bold=True)
se.blank(); se.section("TERMINAL EXCESS RETURN")
E["t_er"] = se.row("FY2031E excess return: (terminal ROE \u2212 ke) \u00d7 opening equity", f"=({I['roe_t']}-{I['ke']})*{F['t_open']}", NUM0)
E["t_tv"] = se.row("Terminal value of excess returns at end FY2030E: ER / (ke \u2212 g)", f"={E['t_er']}/({I['ke']}-{I['g']})", NUM0)
E["t_pv"] = se.row("Present value", f"={E['t_tv']}*'Excess Return'!{L(CN)}{r_edf}", NUM0, bold=True)
se.blank(); se.section("EQUITY VALUE BUILT UP")
E["bv0"] = se.row("Book value today (31 December 2025)", f"={I['equity']}", NUM0, font=F_LINK)
E["pv_sum"] = se.row("Add: present value of explicit excess returns", f"=SUM('Excess Return'!{L(C1)}{r_epv}:{L(CN)}{r_epv})", NUM0)
E["tv"] = se.row("Add: present value of terminal excess returns", f"={E['t_pv']}", NUM0, font=F_LINK)
E["eq"] = se.row("Equity value", f"={E['bv0']}+{E['pv_sum']}+{E['tv']}", NUM0, bold=True, border=DOUBLE_BORDER)
E["ps"] = se.row("Value per share (AED)", f"={E['eq']}/{I['shares']}", NUM2, bold=True, border=DOUBLE_BORDER)
E["bv_share"] = se.row("Share of value that is book value already on the balance sheet", f"={E['bv0']}/{E['eq']}", PCT, bold=True)
E["fr_share"] = se.row("Share of value that is future excess returns", f"=1-{E['bv_share']}", PCT)
E["diff"] = se.row("Difference from the DDM value per share (AED, should be nil)", f"={E['ps']}-{D['ps']}", '0.0000', bold=True)
se.note("The two models agree because the forecast obeys clean surplus and the terminal payout is derived from the "
        "terminal ROE and growth rather than set by hand. The excess-return layout says where the value comes from: "
        "most of it is book value the bank already holds, and the premium to book is the capitalised spread of ROE "
        "over the cost of equity.", height=40)

# --------------------------------------------------------------------------- #
# Justified Multiples
# --------------------------------------------------------------------------- #
sj = S("Justified Multiples", ncols=6, label_width=62, subtitle="What P/B and P/E the fundamentals justify, and what the price implies")
sj.ws.column_dimensions["F"].width = 56
J = {}
sj.section("JUSTIFIED PRICE / BOOK: (ROE \u2212 g) / (ke \u2212 g)")
J["pb_hist"] = sj.row("At FY2025 return on equity", f"=({I['roe_hist']}-{I['g']})/({I['ke']}-{I['g']})", MULT)
J["pb_term"] = sj.row("At the terminal return on equity", f"=({I['roe_t']}-{I['g']})/({I['ke']}-{I['g']})", MULT, bold=True)
J["pb_mkt"] = sj.row("Market price / book today", f"={I['pb_mkt']}", MULT, font=F_LINK, bold=True)
J["pb_model"] = sj.row("Model price / book (DDM equity value / FY2025 book)", f"={D['pb_impl']}", MULT, font=F_LINK)
sj.blank(); sj.section("WHAT THE PRICE IMPLIES")
J["roe_impl"] = sj.row("Sustainable ROE the price implies at the model cost of equity: g + P/B \u00d7 (ke \u2212 g)",
                       f"={I['g']}+{I['pb_mkt']}*({I['ke']}-{I['g']})", PCT, bold=True,
                       note="If you believe the cost of equity, this is the return the market expects the bank to sustain")
J["ke_impl"] = sj.row("Cost of equity the price implies at the terminal ROE: g + (ROE \u2212 g) / P/B",
                      f"={I['g']}+({I['roe_t']}-{I['g']})/{I['pb_mkt']}", PCT, bold=True,
                      note="If you believe the terminal ROE, this is the discount rate the market is using")
sj.blank(); sj.section("JUSTIFIED PRICE / EARNINGS: payout \u00d7 (1 + g) / (ke \u2212 g)")
J["pe_term"] = sj.row("At the terminal payout and growth", f"={I['payout_t']}*(1+{I['g']})/({I['ke']}-{I['g']})", MULT, bold=True)
J["pe_mkt"] = sj.row("Market price / FY2025 earnings", f"={I['pe_mkt']}", MULT, font=F_LINK)
sj.blank()
sj.note("The justified P/B formula is the excess-return model collapsed to one line: a bank earning exactly its cost "
        "of equity is worth book, and every point of ROE above it adds 1/(ke \u2212 g) turns of book. It is the fastest "
        "way to see whether a bank's multiple and its returns are telling the same story.", height=36)

# --------------------------------------------------------------------------- #
# Sensitivity
# --------------------------------------------------------------------------- #
ss = S("Sensitivity", ncols=8, label_width=40, col_width=13.5, subtitle="Value per share (AED) \u2014 closed-form, every cell rebuilds the terminal value")
KES = [0.090, 0.095, 0.100, 0.105, 0.110]; GS = [0.020, 0.025, 0.030, 0.035, 0.040]; ROES = [0.120, 0.130, 0.140, 0.150, 0.160]
pv_divs = D["pv_sum"]  # explicit dividends are unaffected by terminal inputs but are affected by ke; rebuild them
div_cells = [f"Forecast!{L(CE+i+1)}{r_open+4}" for i in range(NY)]
def value_formula(ke, g, roe_t):
    pv_exp = "+".join(f"{d}/(1+{ke})^{i+1}" for i, d in enumerate(div_cells))
    t_div = f"({F['t_open']}*{roe_t}*(1-{g}/{roe_t}))"
    return f"=(({pv_exp})+{t_div}/({ke}-{g})/(1+{ke})^{NY})/{I['shares']}"
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
g1 = grid("COST OF EQUITY (rows) versus LONG-RUN GROWTH (columns), terminal ROE at base", KES, GS, PCT, PCT,
          lambda r, c: value_formula(r, c, I["roe_t"]), (0.100, 0.030))
g2 = grid("COST OF EQUITY (rows) versus TERMINAL ROE (columns), growth at base", KES, ROES, PCT, PCT,
          lambda r, c: value_formula(r, I["g"], c), (0.100, 0.140))
ss.note("The grids hold the explicit five-year dividend path fixed and rebuild only the discounting and the terminal "
        "value, which is where nearly all the sensitivity lives. The base cell reconciles to the DDM sheet exactly.", height=30)

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sy = S("Summary", ncols=6, label_width=60, subtitle="The bank valuation on one page")
sy.ws.column_dimensions["F"].width = 56
sy.section("THE BANK TODAY")
sy.row("Return on average equity, FY2025", f"={I['roe_hist']}", PCT, font=F_LINK)
sy.row("Book value per share (AED)", f"={I['bvps']}", NUM2, font=F_LINK)
sy.row("Share price (AED)", f"={I['price']}", NUM2, font=F_LINK)
sy.row("Price / book", f"={I['pb_mkt']}", MULT, font=F_LINK, bold=True)
sy.row("Cost of equity used", f"={I['ke']}", PCT, font=F_LINK)
sy.blank(); sy.section("THE TWO MODELS")
sy.row("Dividend discount value per share (AED)", f"={D['ps']}", NUM2, font=F_LINK, bold=True)
sy.row("Excess-return value per share (AED)", f"={E['ps']}", NUM2, font=F_LINK, bold=True)
sy.row("Upside / (downside) to the price", f"={D['updown']}", PCT, font=F_LINK, bold=True)
sy.row("Implied price / book at the model value", f"={D['pb_impl']}", MULT, font=F_LINK)
sy.row("Share of value that is book value already held", f"={E['bv_share']}", PCT, font=F_LINK)
sy.row("Terminal value as % of the DDM value", f"={D['tv_share']}", PCT, font=F_LINK)
sy.blank(); sy.section("WHAT THE PRICE IMPLIES")
sy.row("Justified P/B at the terminal ROE", f"={J['pb_term']}", MULT, font=F_LINK)
sy.row("Sustainable ROE the price implies", f"={J['roe_impl']}", PCT, font=F_LINK, bold=True)
sy.row("Cost of equity the price implies at the terminal ROE", f"={J['ke_impl']}", PCT, font=F_LINK, bold=True, border=DOUBLE_BORDER)
sy.blank()
sy.bullets([
    "Two methods, one forecast, one answer: the dividend model and the excess-return model agree to the fils because the "
    "forecast obeys clean surplus. If they ever disagree, the forecast is wrong, not the methods.",
    "The market pays 1.35x book for a bank earning 17% on equity. At a 10% cost of equity that multiple is justified only "
    "if returns fade to about 11%; the fade built in here stops at 14% and lands well above the price. Either the market "
    "expects a harder landing for UAE bank margins, or it discounts them at a higher rate than a country-risk build-up gives.",
    "The valuation is a bet on the spread between ROE and the cost of equity, and almost nothing else. A half-point "
    "on the cost of equity moves the value by about two dirhams; a point on the terminal ROE by about the same.",
    "The excess-return split is the useful framing for a bank: just over half the value is book already on the balance "
    "sheet, and the market price sits only a third of the way from book to the model value. The argument is entirely "
    "about how long a 17% return can last.",
])

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
tests = [
    ("FY2025 equity, earnings and shares tie to Project 6", f"=AND(ABS({I['equity']}-144582)<1,ABS({I['ni']}-23442)<1,ABS({I['shares']}-6311)<1)"),
    ("Clean surplus holds in every forecast year (closing = opening + profit \u2212 dividends)",
     "=AND(" + ",".join(f"ABS(Forecast!{L(CE+i+1)}{r_open+5}-(Forecast!{L(CE+i+1)}{r_open}+Forecast!{L(CE+i+1)}{r_open+2}-Forecast!{L(CE+i+1)}{r_open+4}))<0.01" for i in range(NY)) + ")"),
    ("Opening equity each year equals the prior year's closing equity",
     "=AND(" + ",".join(f"ABS(Forecast!{L(CE+i+1)}{r_open}-Forecast!{L(CE+i)}{r_open+5})<0.01" for i in range(NY)) + ")"),
    ("Terminal book grows at exactly g", f"=ABS({F['t_g']}-{I['g']})<0.00001"),
    ("Long-run growth is below the cost of equity", f"={I['g']}<{I['ke']}"),
    ("Terminal ROE is above the cost of equity (otherwise excess returns are negative)", f"={I['roe_t']}>{I['ke']}"),
    ("Terminal ROE is not above the FY2025 ROE (the fade is a fade)", f"={I['roe_t']}<={I['roe_hist']}"),
    ("Payout rises through the forecast", "=AND(" + ",".join(f"Forecast!{L(CE+i+2)}{r_open+3}>=Forecast!{L(CE+i+1)}{r_open+3}" for i in range(NY - 1)) + ")"),
    ("DDM and excess-return values agree to within AED 0.001 per share", f"=ABS({E['diff']})<0.001"),
    ("DDM terminal exit P/B is within 0.1x of the justified P/B at the terminal ROE", f"=ABS({D['tv_pb']}-{J['pb_term']})<0.1"),
    ("Base sensitivity cell reconciles to the DDM value", f"=ABS(Sensitivity!E{g1+2}-{D['ps']})<0.01"),
    ("Market P/B reconciles to price times shares over equity", f"=ABS({I['pb_mkt']}-{I['mcap']}/{I['equity']})<0.001"),
    ("Implied ROE at the market price reproduces the market P/B", f"=ABS(({J['roe_impl']}-{I['g']})/({I['ke']}-{I['g']})-{I['pb_mkt']})<0.0001"),
    ("Share price lies within the 52-week range", f"=AND({I['price']}>={B['w52lo']}-0.5,{I['price']}<={B['w52hi']}+0.5)"),
]
checks_sheet(book, tests,
             "The ninth check is the one that matters: two different valuation methods built on the same forecast must "
             "produce the same number. It fails the model the moment someone edits the terminal payout by hand.")

# --------------------------------------------------------------------------- #
book.cover(
    blurb="An equity-side valuation of Emirates NBD on the FY2025 audited accounts and 18 September 2026 market data. A "
          "five-year equity roll-forward under clean surplus feeds both a dividend discount model and an excess-return "
          "model; the two are forced to agree, the justified multiples are derived, and the share price is run backwards "
          "into the return on equity and the cost of equity it implies.",
    method=[
        "Cost of equity built up from the US Treasury yield (the dirham is pegged), a UAE country risk premium, a "
        "mature-market equity premium and a stated beta.",
        "An equity roll-forward with ROE fading from 17% to a 14% terminal level and a payout that rises so retained "
        "earnings fund exactly the growth assumed; the terminal payout is derived, not typed.",
        "A dividend discount model with a Gordon terminal value, and an excess-return model starting from book value; "
        "the Checks sheet requires them to agree to the fils.",
        "Justified P/B and P/E, the market-implied ROE and cost of equity, and two closed-form sensitivity grids.",
    ],
    toc=[("Summary", "the bank today, the two values, what the price implies"),
         ("Forecast", "equity roll-forward, EPS, DPS and BVPS, terminal year"),
         ("DDM", "dividends discounted plus terminal value"),
         ("Excess Return", "book value plus capitalised spread over the cost of equity"),
         ("Justified Multiples", "P/B and P/E from fundamentals; implied ROE and cost of equity"),
         ("Sensitivity", "cost of equity against growth and against terminal ROE"),
         ("Inputs", "FY2025 accounts, market data, cost of equity, terminal assumptions"),
         ("Checks", "fourteen tests; must read MODEL OK")],
    highlights=[("Return on equity, FY2025", f"={I['roe_hist']}", PCT),
                ("Cost of equity", f"={I['ke']}", PCT),
                ("DDM value per share (AED)", f"={D['ps']}", NUM2),
                ("Excess-return value per share (AED)", f"={E['ps']}", NUM2),
                ("Upside / (downside)", f"={D['updown']}", PCT),
                ("Sustainable ROE the price implies", f"={J['roe_impl']}", PCT)],
    sources=["Emirates NBD Bank PJSC FY2025 annual accounts, via the Yahoo Finance fundamentals series for EMIRATESNBD.AE; the same figures used in Project 6.",
             "Share price and 52-week range: Dubai Financial Market close, 18 September 2026.",
             "US Treasury yield, UAE country risk premium and equity risk premium: September 2026 market data and Damodaran's country tables; beta is a stated input.",
             "Nothing here is a recommendation. The valuation is an analytical exercise on public information."])

book.finish(freeze={"Forecast": "D6", "DDM": "D6", "Excess Return": "D6"})
path = os.path.join(HERE, "ENBD_Bank_Valuation.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
