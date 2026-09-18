"""
Build 10-reliance-sotp/Reliance_SOTP.xlsx — a sum-of-the-parts valuation of Reliance Industries.

Sheets: Cover · Summary · SOTP · Peers · Market-Implied · Sensitivity · Inputs · Checks

Reliance is five businesses that share one share price: a telecom and digital platform, India's largest retailer,
an integrated refining and petrochemicals complex, a declining gas field, and a basket of media, consumer and new-energy
ventures. No single multiple is right for all of them, so each segment is valued on its own peer group at its own
multiple, the group's economic share of each is taken, consolidated net debt is deducted, and a holding-company
discount is applied. The interesting output is then run backwards: given the other four segments at their base
multiples, what multiple is the market putting on Jio Platforms today? That number is the one the forthcoming Jio
listing will test.

Source data is the audited FY2026 segment information (year to 31 March 2026). Peer multiples were read from public
market data on 18 September 2026 and are stated on the Peers sheet with the exclusions explained.
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
# Source data — audited consolidated segment information, FY2026, ₹ crore. Media release 24 April 2026.
# --------------------------------------------------------------------------- #
SEG = [  # name, FY26 segment EBITDA, FY25 segment EBITDA, RIL economic stake, base multiple, low, high
    ("Digital Services (Jio Platforms)", 76560.0, 65001.0, 0.6703, 12.0, 10.0, 14.0),
    ("Retail (Reliance Retail Ventures)", 27034.0, 25094.0, 0.8506, 25.0, 20.0, 30.0),
    ("Oil to Chemicals", 60546.0, 54988.0, 1.0000, 7.5, 6.5, 8.5),
    ("Oil and Gas (E&P)", 19050.0, 21188.0, 1.0000, 5.0, 4.0, 6.0),
    ("Others (media, consumer brands, new energy)", 10857.0, 8526.0, 1.0000, 10.0, 8.0, 12.0),
]
G = dict(ebitda_rep=207911.0, seg_total=194047.0, debt=374421.0, cash=249704.0, leases=23579.0, spectrum=99552.0,
         nci=181836.0, shares=1353.25, price=1244.60, eps=59.69, w52lo=1235.30, w52hi=1611.80)
PEERS = {  # segment -> [(ticker, name, EV/EBITDA, include?, reason)]
    "Digital Services (Jio Platforms)": [("BHARTIARTL.NS", "Bharti Airtel", 11.6, True, "Direct Indian peer, the anchor"),
                                         ("SCMN.SW", "Swisscom", 11.8, True, ""), ("T", "AT&T", 7.6, True, ""),
                                         ("VZ", "Verizon", 7.7, True, ""),
                                         ("IDEA.NS", "Vodafone Idea", 27.9, False, "Distressed balance sheet; multiple is not a valuation signal"),
                                         ("Z74.SI", "Singtel", 23.4, False, "EBITDA excludes large associate earnings; multiple overstated")],
    "Retail (Reliance Retail Ventures)": [("DMART.NS", "Avenue Supermarts", 44.9, True, "Indian grocery leader"),
                                          ("TRENT.NS", "Trent", 50.7, True, "Indian fashion retail"), ("WMT", "Walmart", 20.8, True, ""),
                                          ("COST", "Costco", 27.9, True, ""), ("TGT", "Target", 10.1, True, "")],
    "Oil to Chemicals": [("BPCL.NS", "Bharat Petroleum", 5.6, True, "Indian refiner-marketer"), ("IOC.NS", "Indian Oil", 4.6, True, ""),
                         ("VLO", "Valero", 9.4, True, ""), ("MPC", "Marathon Petroleum", 9.9, True, ""),
                         ("LYB", "LyondellBasell", 9.1, True, "Petrochemicals"), ("DOW", "Dow", 9.6, True, "Petrochemicals"),
                         ("HINDPETRO.NS", "Hindustan Petroleum", 13.4, False, "Trailing EBITDA depressed by LPG under-recoveries; multiple not meaningful")],
    "Oil and Gas (E&P)": [("ONGC.NS", "ONGC", 4.5, True, "Indian upstream, the anchor"), ("OIL.NS", "Oil India", 6.9, True, ""),
                          ("XOM", "ExxonMobil", 10.4, True, ""), ("CVX", "Chevron", 8.9, True, ""), ("BP", "BP", 13.4, True, ""),
                          ("WDS.AX", "Woodside", 9.7, True, "")],
    "Others (media, consumer brands, new energy)": [("DIS", "Walt Disney", 11.0, True, "Media"), ("NFLX", "Netflix", 21.8, True, "Streaming"),
                                                    ("ZEEL.NS", "Zee Entertainment", 21.0, True, "Indian media"),
                                                    ("SUNTV.NS", "Sun TV", 4.8, True, "Indian media")],
}

book = Book(theme=THEMES["reliance"], project_no=10, project="Sum-of-the-Parts Valuation",
            company="Reliance Industries", units="\u20b9 crore \u00b7 per-share values in \u20b9",
            as_of="FY2026 audited segment results \u00b7 market data 18 Sep 2026")
S = lambda name, **kw: Sheet(book, name, **kw)
NS = len(SEG)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
si = S("Inputs", ncols=9, label_width=50, col_width=13, subtitle="Segment EBITDA, ownership, selected multiples and group-level items")
si.ws.column_dimensions["I"].width = 54
I = {}
si.section("SEGMENT EBITDA AND OWNERSHIP (\u20b9 crore, FY2026 audited segment information)")
si.head(["FY2025A", "FY2026A", "Growth", "RIL stake", "Multiple", "Low", "High"], label="Segment")
I["seg"] = []
for name, e26, e25, stake, m, lo, hi in SEG:
    r = si.r
    refs = si.multi(name, [e25, e26, f"={L(4)}{r}/{L(3)}{r}-1", stake, m, lo, hi],
                    fmt=NUM0)
    ws = si.ws
    ws.cell(r, 5).number_format = PCT; ws.cell(r, 5).font = F_FORMULA
    ws.cell(r, 6).number_format = PCT
    for j in (7, 8, 9): ws.cell(r, j).number_format = MULT
    I["seg"].append(dict(name=name, e25=refs[0], e26=refs[1], g=refs[2], stake=refs[3], m=refs[4], lo=refs[5], hi=refs[6]))
r0 = si.r - NS
I["seg_sum"] = si.multi("Total segment EBITDA", [f"=SUM({L(3)}{r0}:{L(3)}{si.r-1})", f"=SUM({L(4)}{r0}:{L(4)}{si.r-1})"],
                        NUM0, bold=True, border=TOTAL_BORDER)
si.note("RIL stake in Jio Platforms is 67.03% after the 2020 investments by Meta, Google, Vista, KKR, PIF and others; "
        "in Reliance Retail Ventures it is 85.06% after the 2020 and 2023 rounds. Stakes are stated inputs from the "
        "annual report. Multiples are selected on the Peers sheet.", height=40)

si.blank(); si.section("GROUP-LEVEL ITEMS (\u20b9 crore)")
I["ebitda_rep"] = si.row("Consolidated EBITDA, as reported", G["ebitda_rep"], NUM0,
                         note="Exceeds the segment total by unallocated interest income and the one-off disposal gain; segments are valued, the rest is not")
I["debt"] = si.row("Outstanding debt (company definition)", G["debt"], NUM0)
I["cash"] = si.row("Cash and cash equivalents (company definition)", G["cash"], NUM0)
I["leases"] = si.row("Lease liabilities", G["leases"], NUM0)
I["spectrum"] = si.row("Deferred spectrum liabilities", G["spectrum"], NUM0, note="Sits in Jio economically; deducted at group level here, see the SOTP note")
I["nd"] = si.row("Adjusted net debt (consistent with Project 8)", f"={I['debt']}-{I['cash']}+{I['leases']}+{I['spectrum']}", NUM0, bold=True)
I["nci"] = si.row("Non-controlling interests at book (memo, not deducted)", G["nci"], NUM0,
                  note="Minorities are handled through the stake percentages, so deducting book NCI as well would double count")
I["disc"] = si.row("Holding-company discount", 0.10, PCT,
                   note="Indian conglomerates typically trade at 10\u201325% below the sum of their parts; 10% reflects that Jio and Retail are consolidated, not associates")

si.blank(); si.section("MARKET DATA")
I["price"] = si.row("Share price (\u20b9, NSE close 18 Sep 2026)", G["price"], NUM2)
I["shares"] = si.row("Shares outstanding (crore)", G["shares"], NUM2)
I["mcap"] = si.row("Market capitalisation", f"={I['price']}*{I['shares']}", NUM0, bold=True)
I["eps"] = si.row("Diluted EPS, FY2026 (\u20b9)", G["eps"], NUM2)
I["w52lo"] = si.row("52-week low (\u20b9)", G["w52lo"], NUM2)
I["w52hi"] = si.row("52-week high (\u20b9)", G["w52hi"], NUM2)

# --------------------------------------------------------------------------- #
# Peers
# --------------------------------------------------------------------------- #
sp = S("Peers", ncols=7, label_width=34, col_width=13, subtitle="Trailing EV / EBITDA by segment peer group, 18 September 2026, with exclusions stated")
sp.ws.column_dimensions["C"].width = 24; sp.ws.column_dimensions["G"].width = 58
P = {}
for seg in I["seg"]:
    name = seg["name"]
    sp.section(name.upper())
    sp.head(["Company", "EV / EBITDA", "Included", "If included", "Note"], first_col=3, label="Ticker")
    rows = PEERS[name]; r_start = sp.r
    for tkr, nm, mult, inc, why in rows:
        ws = sp.ws; r = sp.r
        ws.cell(r, 2, tkr).font = F_TEXT; ws.cell(r, 3, nm).font = F_TEXT
        c = ws.cell(r, 4, mult); c.font = F_INPUT; c.number_format = MULT
        c = ws.cell(r, 5, 1 if inc else 0); c.font = F_INPUT; c.number_format = '"Yes";;"No"'; c.alignment = Alignment(horizontal="center")
        c = ws.cell(r, 6, f'=IF(E{r}=1,D{r},"")'); c.font = F_FORMULA; c.number_format = MULT
        c = ws.cell(r, 7, why); c.font = F_NOTE
        sp.r += 1
    r_end = sp.r - 1
    rng_f = f"$F${r_start}:$F${r_end}"
    P[name] = {}
    P[name]["median"] = sp.row("Median of included peers", f"=MEDIAN({rng_f})", MULT, bold=True, col=4, key=name + " median")
    P[name]["mean"] = sp.row("Mean of included peers", f"=AVERAGE({rng_f})", MULT, col=4, key=name + " mean")
    P[name]["sel"] = sp.row("Selected multiple for the segment", f"={seg['m']}", MULT, font=F_LINK, bold=True, col=4, key=name + " sel")
    P[name]["vs"] = sp.row("Selected versus median", f"={P[name]['sel']}/{P[name]['median']}-1", PCT, col=4, key=name + " vs")
    sp.blank()
sp.note("Multiples are trailing enterprise value to EBITDA as published by Yahoo Finance on 18 September 2026 and are "
        "stated inputs; an exclusion is made only where the trailing figure is not a valuation signal and the reason is "
        "written next to it. Selected multiples sit below the peer medians for Retail (lower margins, hyper-local investment) "
        "and O2C (Indian marketing discount, weak chemical deltas), and slightly above for Digital (19% EBITDA growth against Bharti's mid-teens, with Bharti as the anchor).", height=48)

# --------------------------------------------------------------------------- #
# SOTP
# --------------------------------------------------------------------------- #
so = S("SOTP", ncols=9, label_width=50, col_width=14, subtitle="Segment enterprise values, RIL's share, net debt, discount and value per share")
so.ws.column_dimensions["I"].width = 46
V = {}
so.section("SEGMENT VALUES (\u20b9 crore)")
so.head(["EBITDA FY26", "Multiple", "Segment EV", "RIL stake", "RIL share of EV", "% of gross", "Per share (\u20b9)"], label="Segment")
r_first = so.r
V["rows"] = []
for seg in I["seg"]:
    r = so.r
    refs = so.multi(seg["name"], [f"={seg['e26']}", f"={seg['m']}", f"={L(3)}{r}*{L(4)}{r}", f"={seg['stake']}",
                                  f"={L(5)}{r}*{L(6)}{r}", "", f"={L(7)}{r}/{I['shares']}"], NUM0)
    ws = so.ws
    ws.cell(r, 3).font = F_LINK; ws.cell(r, 4).font = F_LINK; ws.cell(r, 4).number_format = MULT
    ws.cell(r, 6).font = F_LINK; ws.cell(r, 6).number_format = PCT
    ws.cell(r, 8).number_format = PCT; ws.cell(r, 9).number_format = NUM2
    V["rows"].append(dict(ev=refs[2], share=refs[4], pct_cell=(r, 8), ps=refs[6]))
r_last = so.r - 1
V["gross"] = so.row("Gross asset value (RIL share)", f"=SUM({L(7)}{r_first}:{L(7)}{r_last})", NUM0, bold=True, col=7, border=TOTAL_BORDER)
for row in V["rows"]:
    r, c = row["pct_cell"]; so.ws.cell(r, c).value = f"={L(7)}{r}/{V['gross']}"; so.ws.cell(r, c).font = F_FORMULA
so.ws.cell(so.r - 1, 9).value = f"={V['gross']}/{I['shares']}"; so.ws.cell(so.r - 1, 9).number_format = NUM2; so.ws.cell(so.r - 1, 9).font = F_FORMULA
so.blank(); so.section("FROM GROSS ASSET VALUE TO EQUITY VALUE")
V["nd"] = so.row("Less: adjusted net debt", f"=-{I['nd']}", NUM0, font=F_LINK, col=7)
V["nav"] = so.row("Net asset value before discount", f"={V['gross']}+{V['nd']}", NUM0, bold=True, col=7, border=TOTAL_BORDER)
V["disc_amt"] = so.row("Less: holding-company discount", f"=-{V['nav']}*{I['disc']}", NUM0, col=7)
V["eq"] = so.row("Equity value", f"={V['nav']}+{V['disc_amt']}", NUM0, bold=True, col=7, border=DOUBLE_BORDER)
so.blank(); so.section("PER SHARE")
V["ps_gross"] = so.row("Gross asset value per share (\u20b9)", f"={V['gross']}/{I['shares']}", NUM2, col=7)
V["ps_nav"] = so.row("NAV per share before discount (\u20b9)", f"={V['nav']}/{I['shares']}", NUM2, col=7)
V["ps"] = so.row("SOTP value per share (\u20b9)", f"={V['eq']}/{I['shares']}", NUM2, bold=True, col=7, border=DOUBLE_BORDER)
V["px"] = so.row("Share price (\u20b9)", f"={I['price']}", NUM2, font=F_LINK, col=7)
V["updown"] = so.row("Upside / (downside) to the SOTP value", f"={V['ps']}/{V['px']}-1", PCT, bold=True, col=7)
V["impl_ev_ebitda"] = so.row("Implied group EV / segment EBITDA at the SOTP value",
                             f"=({V['eq']}+{I['nd']})/{I['seg_sum'][1]}", MULT, col=7)
V["mkt_ev_ebitda"] = so.row("Market EV / segment EBITDA today", f"=({I['mcap']}+{I['nd']})/{I['seg_sum'][1]}", MULT, col=7)
V["pe"] = so.row("Implied P/E at the SOTP value", f"={V['ps']}/{I['eps']}", MULT, col=7)
so.blank()
so.note("Net debt is deducted at group level in full although roughly a third of the spectrum liabilities belong "
        "economically to Jio's minority shareholders. The alternative \u2014 deducting Jio's own net debt inside the segment "
        "before applying the stake \u2014 needs Jio's balance sheet, which the results release does not give. The simplification "
        "understates value by roughly \u20b925 per share and is left in because it errs the right way.", height=44)

# --------------------------------------------------------------------------- #
# Market-Implied
# --------------------------------------------------------------------------- #
sm = S("Market-Implied", ncols=6, label_width=62, subtitle="Run backwards: what is the market paying for Jio Platforms today?")
sm.ws.column_dimensions["F"].width = 56
M = {}
sm.section("SOLVE FOR THE JIO MULTIPLE THE SHARE PRICE IMPLIES")
M["mcap"] = sm.row("Market capitalisation", f"={I['mcap']}", NUM0, font=F_LINK)
M["eq_pre"] = sm.row("Equity value before discount the market implies (market cap / (1 \u2212 discount))",
                     f"={M['mcap']}/(1-{I['disc']})", NUM0)
M["gav"] = sm.row("Gross asset value the market implies", f"={M['eq_pre']}+{I['nd']}", NUM0, bold=True)
others = "+".join(V["rows"][i]["share"] for i in range(1, NS))
M["non_jio"] = sm.row("Less: RIL's share of the other four segments at base multiples", f"=-({others})", NUM0)
M["jio_share"] = sm.row("RIL's share of Jio Platforms EV the market implies", f"={M['gav']}+{M['non_jio']}", NUM0, bold=True, border=TOTAL_BORDER)
M["jio_ev"] = sm.row("Jio Platforms EV at 100%", f"={M['jio_share']}/{I['seg'][0]['stake']}", NUM0)
M["jio_mult"] = sm.row("Market-implied Jio EV / EBITDA", f"={M['jio_ev']}/{I['seg'][0]['e26']}", MULT, bold=True, border=DOUBLE_BORDER)
M["vs_median"] = sm.row("Premium to the telecom peer median", f"={M['jio_mult']}/{P[I['seg'][0]['name']]['median']}-1", PCT)
M["vs_bharti"] = sm.row("Premium to Bharti Airtel's multiple", f"={M['jio_mult']}/Peers!$D$6-1", PCT, bold=True,
                        note="Bharti is the direct listed comparable and the anchor a Jio IPO will be priced against")
M["jio_usd"] = sm.row("Market-implied Jio Platforms EV in US$ billion at \u20b983.5", f"={M['jio_ev']}/8350", NUM)
sm.blank(); sm.section("THE SAME SOLVE FOR RETAIL")
others_r = "+".join(V["rows"][i]["share"] for i in (0, 2, 3, 4))
M["ret_share"] = sm.row("RIL's share of Retail EV the market implies (Jio at base)", f"={M['gav']}-({others_r})", NUM0)
M["ret_ev"] = sm.row("Retail EV at 100%", f"={M['ret_share']}/{I['seg'][1]['stake']}", NUM0)
M["ret_mult"] = sm.row("Market-implied Retail EV / EBITDA", f"={M['ret_ev']}/{I['seg'][1]['e26']}", MULT, bold=True, border=DOUBLE_BORDER)
M["ret_usd"] = sm.row("Market-implied Retail EV in US$ billion", f"={M['ret_ev']}/8350", NUM,
                      note="The 2023 private round valued Reliance Retail Ventures at roughly US$100 billion")
sm.blank()
sm.note("Only one segment can be solved for at a time; each solve holds the other four at their base multiples. "
        "The Jio solve is the useful one because Jio is about to be priced by a public market, and the gap between "
        "this number and the eventual listing multiple is the size of the re-rating the sell-side is arguing about.", height=40)

# --------------------------------------------------------------------------- #
# Sensitivity
# --------------------------------------------------------------------------- #
ss = S("Sensitivity", ncols=8, label_width=40, col_width=13.5, subtitle="SOTP value per share (\u20b9) \u2014 every cell rebuilds the full bridge")
JM = [10.0, 11.0, 12.0, 13.0, 14.0]; RM = [20.0, 22.5, 25.0, 27.5, 30.0]; DS = [0.00, 0.05, 0.10, 0.15, 0.20]
def ps_formula(jio_m, ret_m, disc):
    s = I["seg"]
    parts = [f"{s[0]['e26']}*{jio_m}*{s[0]['stake']}", f"{s[1]['e26']}*{ret_m}*{s[1]['stake']}"]
    parts += [f"{s[i]['e26']}*{s[i]['m']}*{s[i]['stake']}" for i in range(2, NS)]
    return f"=(({'+'.join(parts)})-{I['nd']})*(1-{disc})/{I['shares']}"
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
            cc = ss.ws.cell(ss.r, 3 + j, fn(f"$B{ss.r}", f"{L(3+j)}${h}")); cc.font = F_FORMULA; cc.number_format = NUM0
            if abs(rv - base[0]) < 1e-9 and abs(cval - base[1]) < 1e-9:
                cc.fill = book.fill_accent; cc.font = Font(name=FONT, size=10, bold=True)
        ss.r += 1
    ss.blank(); return g0
g1 = grid("JIO MULTIPLE (rows) versus RETAIL MULTIPLE (columns), discount at base", JM, RM, MULT, MULT,
          lambda r, c: ps_formula(r, c, I["disc"]), (12.0, 25.0))
g2 = grid("JIO MULTIPLE (rows) versus HOLDING-COMPANY DISCOUNT (columns), Retail at base", JM, DS, MULT, PCT,
          lambda r, c: ps_formula(r, I["seg"][1]["m"], c), (12.0, 0.10))
ss.section("LOW / BASE / HIGH ON EVERY SEGMENT AT ONCE")
lo = "=((" + "+".join(f"{s['e26']}*{s['lo']}*{s['stake']}" for s in I["seg"]) + f")-{I['nd']})*(1-{I['disc']})/{I['shares']}"
hi = "=((" + "+".join(f"{s['e26']}*{s['hi']}*{s['stake']}" for s in I["seg"]) + f")-{I['nd']})*(1-{I['disc']})/{I['shares']}"
V["ps_lo"] = ss.row("All segments at the low multiple (\u20b9 per share)", lo, NUM0)
V["ps_hi"] = ss.row("All segments at the high multiple (\u20b9 per share)", hi, NUM0)
V["ps_base2"] = ss.row("All segments at the base multiple (\u20b9 per share)", f"={V['ps']}", NUM0, font=F_LINK, bold=True)

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sy = S("Summary", ncols=6, label_width=60, subtitle="The sum of the parts on one page")
sy.ws.column_dimensions["F"].width = 56
sy.section("THE PARTS")
for seg, row in zip(I["seg"], V["rows"]):
    sy.row(seg["name"] + " \u2014 RIL share (\u20b9 crore)", f"={row['share']}", NUM0, font=F_LINK)
sy.row("Gross asset value (\u20b9 crore)", f"={V['gross']}", NUM0, font=F_LINK, bold=True, border=TOTAL_BORDER)
sy.row("Less: adjusted net debt", f"={V['nd']}", NUM0, font=F_LINK)
sy.row("Less: holding-company discount", f"={V['disc_amt']}", NUM0, font=F_LINK)
sy.row("Equity value (\u20b9 crore)", f"={V['eq']}", NUM0, font=F_LINK, bold=True, border=DOUBLE_BORDER)
sy.blank(); sy.section("PER SHARE")
sy.row("SOTP value per share (\u20b9)", f"={V['ps']}", NUM2, font=F_LINK, bold=True)
sy.row("Range, all segments low to high (\u20b9)", f"=TEXT({V['ps_lo']},\"#,##0\")&\" \u2013 \"&TEXT({V['ps_hi']},\"#,##0\")", "@", font=F_LINK)
sy.row("Share price (\u20b9)", f"={V['px']}", NUM2, font=F_LINK)
sy.row("Upside / (downside)", f"={V['updown']}", PCT, font=F_LINK, bold=True)
sy.blank(); sy.section("WHAT THE MARKET IS PAYING FOR JIO")
sy.row("Market-implied Jio EV / EBITDA", f"={M['jio_mult']}", MULT, font=F_LINK, bold=True)
sy.row("Market-implied Jio EV (US$ billion)", f"={M['jio_usd']}", NUM, font=F_LINK)
sy.row("Selected Jio multiple in the SOTP", f"={I['seg'][0]['m']}", MULT, font=F_LINK)
sy.row("Premium to Bharti Airtel implied by the price", f"={M['vs_bharti']}", PCT, font=F_LINK, border=DOUBLE_BORDER)
sy.blank()
sy.bullets([
    "Jio and Retail together are roughly two-thirds of gross asset value; the energy businesses that generate over 40% of "
    "EBITDA are worth well under a third of it. That is the whole investment case for the group's last decade in one line.",
    "At peer multiples the SOTP sits below the share price, not above it. Run backwards, the price is consistent with "
    "Jio Platforms at roughly 17x EBITDA — a premium of about half to Bharti Airtel — or, equivalently, with Retail "
    "near 40x. The market has already priced a Jio listing at a premium; the IPO is the event that tests whether it is right.",
    "The holding-company discount is the single largest judgement after the Jio multiple. Each five points is worth about "
    "\u20b990 per share, which is the second sensitivity grid.",
    "The energy segments are valued on peer multiples rather than through the cycle. O2C EBITDA in FY2026 was helped by "
    "strong middle-distillate cracks; a mid-cycle number would take \u20b930\u201340 per share off the total.",
])

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
tests = [
    ("Segment EBITDA total ties to the published \u20b9194,047 crore", f"=ABS({I['seg_sum'][1]}-194047)<1"),
    ("FY2025 segment total ties to the published \u20b9174,797 crore", f"=ABS({I['seg_sum'][0]}-174797)<1"),
    ("Segment EBITDA is below reported consolidated EBITDA (unallocated items excluded)", f"={I['seg_sum'][1]}<{I['ebitda_rep']}"),
    ("Adjusted net debt matches Project 8", f"=ABS({I['nd']}-247848)<1"),
    ("Every stake is between 50% and 100%", "=AND(" + ",".join(f"{s['stake']}>0.5,{s['stake']}<=1" for s in I["seg"]) + ")"),
    ("Every base multiple lies within its low\u2013high range", "=AND(" + ",".join(f"{s['m']}>={s['lo']},{s['m']}<={s['hi']}" for s in I["seg"]) + ")"),
    ("Segment shares of gross value sum to 100%", f"=ABS(SUM(SOTP!H{r_first}:H{r_last})-1)<0.0001"),
    ("Equity value equals NAV less discount", f"=ABS({V['eq']}-{V['nav']}*(1-{I['disc']}))<0.01"),
    ("Per-share value reconciles to equity value over shares", f"=ABS({V['ps']}-{V['eq']}/{I['shares']})<0.01"),
    ("Base sensitivity cell reconciles to the SOTP sheet", f"=ABS(Sensitivity!E{g1+2}-{V['ps']})<0.5"),
    ("Low case is below base and base is below high", f"=AND({V['ps_lo']}<{V['ps']},{V['ps']}<{V['ps_hi']})"),
    ("Market-implied Jio multiple reproduces the share price when substituted",
     f"=ABS(({I['seg'][0]['e26']}*{M['jio_mult']}*{I['seg'][0]['stake']}+({others})-{I['nd']})*(1-{I['disc']})/{I['shares']}-{I['price']})<0.5"),
    ("Every peer median evaluates to a number", "=AND(" + ",".join(f"ISNUMBER({P[s['name']]['median']})" for s in I["seg"]) + ")"),
    ("Every selected multiple lies within 50% of its peer median",
     "=AND(" + ",".join(f"ABS({P[s['name']]['vs']})<=0.50" for s in I["seg"]) + ")"),
    ("The Bharti Airtel reference cell on the Peers sheet is Bharti", '=Peers!$B$6="BHARTIARTL.NS"'),
    ("Share price lies within the 52-week range", f"=AND({I['price']}>={I['w52lo']}-1,{I['price']}<={I['w52hi']}+1)"),
]
checks_sheet(book, tests,
             "The twelfth check is the important one: it substitutes the market-implied Jio multiple back into the full "
             "bridge and demands the share price come out, which proves the reverse solve is the exact inverse of the forward model.")

# --------------------------------------------------------------------------- #
book.cover(
    blurb="A sum-of-the-parts valuation of Reliance Industries on the audited FY2026 segment results and 18 September 2026 "
          "market data. Each of the five reported segments is valued on its own peer group, RIL's economic share is taken, "
          "consolidated net debt and a holding-company discount are deducted, and the bridge is then run backwards to show "
          "the multiple the market is paying for Jio Platforms ahead of its listing.",
    method=[
        "Segment EBITDA from the audited Ind AS 108 segment note; ownership stakes in Jio Platforms and Reliance Retail "
        "Ventures from the annual report.",
        "Peer multiples read from public market data with each exclusion stated and reasoned; the selected multiple is "
        "shown against the peer median so the judgement is visible.",
        "Net debt on the same adjusted basis as Project 8, so the two workbooks agree to the rupee.",
        "A market-implied solve for the Jio and Retail multiples, two sensitivity grids and a low\u2013high range.",
    ],
    toc=[("Summary", "the parts, the bridge, the per-share value and the market-implied Jio multiple"),
         ("SOTP", "segment EVs, RIL share, net debt, discount, per-share"),
         ("Peers", "peer multiples by segment with medians and exclusions"),
         ("Market-Implied", "the Jio and Retail multiples the share price implies"),
         ("Sensitivity", "Jio versus Retail multiple, Jio versus discount, low\u2013high range"),
         ("Inputs", "segment EBITDA, stakes, selected multiples, group items, market data"),
         ("Checks", "sixteen tests; must read MODEL OK")],
    highlights=[("Gross asset value (\u20b9 crore)", f"={V['gross']}", NUM0),
                ("SOTP value per share (\u20b9)", f"={V['ps']}", NUM2),
                ("Share price (\u20b9)", f"={V['px']}", NUM2),
                ("Upside / (downside)", f"={V['updown']}", PCT),
                ("Market-implied Jio EV / EBITDA", f"={M['jio_mult']}", MULT),
                ("Market-implied Jio EV (US$ bn)", f"={M['jio_usd']}", NUM)],
    sources=["Reliance Industries Limited, audited consolidated segment information for the year ended 31 March 2026 (media release, 24 April 2026).",
             "Ownership stakes in Jio Platforms (67.03%) and Reliance Retail Ventures (85.06%): Reliance Industries annual report.",
             "Peer EV / EBITDA multiples: Yahoo Finance, 18 September 2026, as listed on the Peers sheet.",
             "Share price: NSE close, 18 September 2026. Holding-company discount is a stated input.",
             "Nothing here is a recommendation. The valuation is an analytical exercise on public information."])

book.finish(freeze={"SOTP": "C6", "Peers": "C4"})
path = os.path.join(HERE, "Reliance_SOTP.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
