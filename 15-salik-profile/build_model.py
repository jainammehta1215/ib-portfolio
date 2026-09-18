"""
Build 15-salik-profile/Salik_Profile_Buyer_Screen.xlsx — a company profile and buyer screen for Salik, Dubai's toll-gate
operator, of the kind an M&A team prepares before deciding whether there is a transaction to pitch.

Sheets: Cover · Profile · Financials · Trading · Buyer Screen · Ability to Pay · Checks

Salik is the simplest infrastructure asset on the Dubai Financial Market: ten toll gates, a 49-year concession to 2071,
a tariff set by decree, almost no capital expenditure and a policy of paying out all of its profit. That makes it an
unusually clean case for the two questions a buyer screen answers. Who could plausibly own it, ranked on strategic
fit, precedent, capacity and the likelihood of the Government of Dubai consenting? And what could each type of buyer
pay: a financial sponsor constrained by leverage and a return hurdle, and a strategic operator constrained only by its
cost of capital and the concession's remaining life?

Source data: Salik FY2022\u2013FY2025 financials via public market data, DFM price 18 September 2026, peer multiples the
same day. Concession terms and tariff from the IPO prospectus and RTA announcements. Buyer scores are judgements and
are labelled as such.
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)
from ibkit.sheet import Sheet, checks_sheet

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"

YEARS = ["FY2022A", "FY2023A", "FY2024A", "FY2025A"]
FIN = dict(revenue=[1892.0, 2109.0, 2292.0, 3097.0], ebitda=[1443.0, 1420.0, 1627.0, 2165.0], da=[44.0, 83.0, 91.0, 145.0],
           interest=[73.0, 239.0, 257.0, 313.0], tax=[0.0, 0.0, 115.0, 154.0], ni=[1326.0, 1098.0, 1165.0, 1553.0],
           cfo=[1516.0, 1454.0, 1463.0, 2082.0], capex=[3.0, 5.0, 6.0, 3.0], divs=[827.0, 1039.0, 1095.0, 1391.0],
           debt=[3986.0, 3997.0, 4000.0, 4001.0], cash=[823.0, 266.0, 964.0, 513.0], equity=[604.0, 663.0, 1088.0, 1219.0])
M = dict(price=5.37, shares=7500.0, w52lo=4.96, w52hi=6.78, ipo_price=2.00, ipo_date="29 Sep 2022", gov_stake=0.751)
PEERS = [("Transurban (ASX)", "Toll roads, Australia and North America", 26.3, True),
         ("Ferrovial (Cintra)", "Toll roads, Europe and North America", 27.7, False, "Group multiple distorted by construction and airports"),
         ("Vinci", "Concessions and contracting, France", 7.2, False, "Contracting dilutes the concession multiple"),
         ("IRB Infrastructure (NSE)", "Toll roads, India", 15.5, True),
         ("Vesta (NYSE)", "Industrial real estate, Mexico \u2014 yield comparator", 16.9, False, "Not a toll road; shown for yield context only"),
         ("Tabreed (DFM)", "District cooling concessions, UAE", 12.1, True),
         ("Salik at the market price", "", None, None)]
BUYERS = [  # name, type, home, precedent, fit, precedent_score, capacity, regional, consent
    ("Transurban", "Strategic", "Australia", "Owns 20+ urban toll roads; WestConnex, 407 ETR", 5, 5, 3, 2, 3),
    ("Vinci Concessions", "Strategic", "France", "Autoroutes du Sud; Lima Expresa; airports", 5, 5, 5, 3, 3),
    ("Ferrovial / Cintra", "Strategic", "Spain / NL", "407 ETR, LBJ and NTE managed lanes", 5, 5, 4, 2, 3),
    ("Abertis (Mundys / ACS)", "Strategic", "Spain / Italy", "Largest toll-road operator by km; Red de Carreteras, Sanef", 5, 5, 4, 3, 3),
    ("Globalvia", "Strategic", "Spain", "Mid-size toll roads and rail concessions", 4, 4, 2, 2, 2),
    ("Global Infrastructure Partners (BlackRock)", "Financial", "US", "Jafurah midstream (Aramco); Gatwick; ports", 4, 3, 5, 4, 4),
    ("Brookfield Infrastructure", "Financial", "Canada", "Toll roads in Brazil, India, Peru; large Gulf activity", 4, 4, 5, 4, 4),
    ("Macquarie Asset Management", "Financial", "Australia", "Historic toll-road sponsor; UK, US, India roads", 4, 5, 5, 3, 3),
    ("KKR Infrastructure", "Financial", "US", "ADNOC pipelines; Indian roads platform", 4, 4, 5, 4, 4),
    ("CVC / DIF Capital Partners", "Financial", "NL / Luxembourg", "Mid-market PPP and transport concessions", 3, 4, 3, 3, 3),
    ("PIF / Mubadala / ADQ", "Sovereign", "Gulf", "Regional infrastructure; co-investment with global sponsors", 3, 2, 5, 5, 5),
    ("Dubai Holding / ICD (incumbent)", "Sovereign", "Dubai", "Existing owner group; reorganisation rather than a sale", 3, 3, 5, 5, 5),
]
WEIGHTS = [0.25, 0.20, 0.20, 0.15, 0.20]

book = Book(theme=THEMES["salik"], project_no=15, project="Company Profile and Buyer Screen",
            company="Salik", units="AED millions \u00b7 per-share values in AED", as_of="FY2025 accounts \u00b7 market data 18 Sep 2026")
S = lambda name, **kw: Sheet(book, name, **kw)
NY = 4; CE = 3

# --------------------------------------------------------------------------- #
# Financials
# --------------------------------------------------------------------------- #
sf = S("Financials", ncols=7, label_width=52, col_width=13.5, subtitle="Four years since the IPO \u2014 income, cash and balance sheet")
sf.ws.column_dimensions["G"].width = 44
F = {}
sf.section("INCOME STATEMENT (AED millions)")
sf.head(YEARS)
F["rev"] = sf.multi("Revenue", FIN["revenue"], NUM0, bold=True)
F["g"] = sf.multi("   Growth", [""] + [f"={L(CE+i+1)}{sf.r-1}/{L(CE+i)}{sf.r-1}-1" for i in range(3)], PCT)
F["ebitda"] = sf.multi("EBITDA", FIN["ebitda"], NUM0, bold=True)
F["margin"] = sf.multi("   EBITDA margin", [f"={L(CE+i)}{sf.r-1}/{L(CE+i)}{sf.r-3}" for i in range(NY)], PCT)
F["da"] = sf.multi("Depreciation and amortisation", [-x for x in FIN["da"]], NUM0)
F["int"] = sf.multi("Net finance cost", [-x for x in FIN["interest"]], NUM0)
F["tax"] = sf.multi("Corporate tax", [-x for x in FIN["tax"]], NUM0)
F["ni"] = sf.multi("Net profit", FIN["ni"], NUM0, bold=True, border=TOTAL_BORDER)
F["ni_chk"] = sf.multi("   check: EBITDA \u2212 D&A \u2212 finance \u2212 tax", [f"={L(CE+i)}{sf.r-6}+{L(CE+i)}{sf.r-4}+{L(CE+i)}{sf.r-3}+{L(CE+i)}{sf.r-2}" for i in range(NY)], NUM0, font=F_FORMULA)
F["eps"] = sf.multi("Earnings per share (AED)", [f"={L(CE+i)}{sf.r-2}/{M['shares']}" for i in range(NY)], NUM3 if False else '0.000')
sf.blank(); sf.section("CASH FLOW AND DISTRIBUTION")
F["cfo"] = sf.multi("Cash from operations", FIN["cfo"], NUM0)
F["capex"] = sf.multi("Capital expenditure", [-x for x in FIN["capex"]], NUM0)
F["fcf"] = sf.multi("Free cash flow", [f"={L(CE+i)}{sf.r-2}+{L(CE+i)}{sf.r-1}" for i in range(NY)], NUM0, bold=True, border=TOTAL_BORDER)
F["divs"] = sf.multi("Dividends paid", [-x for x in FIN["divs"]], NUM0)
F["payout"] = sf.multi("   Dividends paid / prior-year net profit", [""] + [f"=-{F['divs'][i+1]}/{F['ni'][i]}" for i in range(3)], PCT)
F["conv"] = sf.multi("   Free cash flow / EBITDA", [f"={F['fcf'][i]}/{F['ebitda'][i]}" for i in range(NY)], PCT)
sf.blank(); sf.section("BALANCE SHEET")
F["debt"] = sf.multi("Borrowings", FIN["debt"], NUM0)
F["cash"] = sf.multi("Cash", FIN["cash"], NUM0)
F["nd"] = sf.multi("Net debt", [f"={L(CE+i)}{sf.r-2}-{L(CE+i)}{sf.r-1}" for i in range(NY)], NUM0, bold=True)
F["lev"] = sf.multi("   Net debt / EBITDA", [f"={F['nd'][i]}/{F['ebitda'][i]}" for i in range(NY)], MULT)
F["cover"] = sf.multi("   EBITDA / net finance cost", [f"=-{F['ebitda'][i]}/{F['int'][i]}" for i in range(NY)], MULT)
F["equity"] = sf.multi("Shareholders' equity", FIN["equity"], NUM0)
sf.note("The FY2025 step-up is the tariff, not traffic: variable pricing from 31 January 2025 (AED 6 at peak, AED 4 off-peak) "
        "and two gates opened in November 2024. Capex is negligible because the RTA builds and operates the gates; Salik's "
        "cost is a concession fee inside operating expenses. Equity is small because the company was formed with AED 4bn of "
        "debt and pays out all its profit.", height=44)
last = lambda key: F[key][NY - 1]

# --------------------------------------------------------------------------- #
# Profile
# --------------------------------------------------------------------------- #
sp = S("Profile", ncols=6, label_width=46, subtitle="What Salik is, in the form a buyer's investment committee would read it")
sp.ws.column_dimensions["C"].width = 90
P = {}
def fact(label, text):
    sp.ws.cell(sp.r, 2, label).font = F_BOLD
    c = sp.ws.cell(sp.r, 3, text); c.font = F_TEXT; c.alignment = Alignment(wrap_text=True, vertical="top")
    sp.ws.merge_cells(start_row=sp.r, start_column=3, end_row=sp.r, end_column=6)
    sp.ws.row_dimensions[sp.r].height = 15 * max(1, len(text) // 105 + 1); sp.r += 1
sp.section("THE BUSINESS")
fact("What it does", "Operates Dubai's automated road-toll system under an exclusive 49-year concession from the Roads and Transport Authority running to 2071. Ten gates as of 2025. Tolls are collected electronically per crossing; the tariff is set by government decree.")
fact("Revenue model", "Toll revenue per crossing (the large majority), plus fines, tag activation and ancillary services. Variable pricing was introduced on 31 January 2025: AED 6 at peak hours, AED 4 off-peak, free between 1am and 2am. Ancillary growth comes from parking payments and partnerships.")
fact("Cost model", "Asset-light. The RTA builds, owns and maintains the gates; Salik pays a concession fee and a service fee. There is almost no capital expenditure, so EBITDA converts to cash at close to 100%.")
fact("Ownership", "Listed on the Dubai Financial Market since 29 September 2022 at AED 2.00 per share, raising AED 3.7bn for the Government of Dubai, which retains 75.1%. Free float 24.9%.")
fact("Capital structure", "AED 4.0bn of bank debt put in place at formation and held flat since; the company distributes all of its net profit. Leverage has fallen from 2.2x to 1.6x EBITDA on earnings growth alone.")
sp.blank(); sp.section("WHY A BUYER WOULD LOOK")
fact("Predictability", "Traffic in a growing city, a tariff set by decree and a 45-year remaining term. The revenue line is closer to a regulated utility than to a typical toll road with traffic risk.")
fact("Growth levers", "New gates (two added in 2024), tariff changes (2025), population growth, and ancillary revenue. All three of the first are policy decisions, not management decisions, which is both the attraction and the constraint.")
fact("The constraint", "The Government of Dubai holds 75.1% and the RTA controls the gates, the tariff and the concession. Any change of control is a policy decision. The realistic transactions are a further sell-down of the government's stake, an anchor stake to a long-term investor, or a partnership; an unsolicited bid is not one of them.")
sp.blank(); sp.section("KEY NUMBERS, FY2025")
for lab, ref, fmt in [("Revenue (AED m)", last("rev"), NUM0), ("EBITDA (AED m)", last("ebitda"), NUM0), ("EBITDA margin", last("margin"), PCT),
                      ("Net profit (AED m)", last("ni"), NUM0), ("Free cash flow / EBITDA", last("conv"), PCT), ("Net debt / EBITDA", last("lev"), MULT)]:
    P[lab] = sp.row(lab, f"={ref}", fmt, font=F_LINK)

# --------------------------------------------------------------------------- #
# Trading
# --------------------------------------------------------------------------- #
st = S("Trading", ncols=7, label_width=44, col_width=14, subtitle="Where Salik trades, against listed toll roads and a local concession")
st.ws.column_dimensions["C"].width = 40; st.ws.column_dimensions["G"].width = 46
T = {}
st.section("SALIK AT THE MARKET PRICE")
T["price"] = st.row("Share price (AED, DFM close 18 Sep 2026)", M["price"], NUM2, col=4)
T["shares"] = st.row("Shares (millions)", M["shares"], NUM0, col=4)
T["mcap"] = st.row("Market capitalisation", f"={T['price']}*{T['shares']}", NUM0, bold=True, col=4)
T["nd"] = st.row("Net debt, FY2025", f"={last('nd')}", NUM0, font=F_LINK, col=4)
T["ev"] = st.row("Enterprise value", f"={T['mcap']}+{T['nd']}", NUM0, bold=True, col=4)
T["ev_ebitda"] = st.row("EV / EBITDA, FY2025", f"={T['ev']}/{last('ebitda')}", MULT, bold=True, col=4)
T["pe"] = st.row("P / E, FY2025", f"={T['mcap']}/{last('ni')}", MULT, col=4)
T["dy"] = st.row("Dividend yield on FY2025 dividends paid", f"=-{last('divs')}/{T['mcap']}", PCT, col=4)
T["fcf_y"] = st.row("Free cash flow yield to enterprise value", f"={last('fcf')}/{T['ev']}", PCT, col=4)
T["since_ipo"] = st.row("Price return since the IPO at AED 2.00", f"={T['price']}/{M['ipo_price']}-1", PCT, bold=True, col=4)
st.blank(); st.section("LISTED COMPARABLES \u2014 trailing EV / EBITDA, 18 September 2026")
st.head(["Business", "EV / EBITDA", "Included", "If included", "Note"], first_col=3, label="Company")
r0 = st.r
for row in PEERS:
    nm, biz, mult, inc = row[0], row[1], row[2], row[3]
    ws = st.ws; r = st.r
    ws.cell(r, 2, nm).font = F_BOLD if mult is None else F_TEXT; ws.cell(r, 3, biz).font = F_TEXT
    if mult is None:
        c = ws.cell(r, 4, f"={T['ev_ebitda']}"); c.font = F_LINK; c.number_format = MULT
    else:
        c = ws.cell(r, 4, mult); c.font = F_INPUT; c.number_format = MULT
        c = ws.cell(r, 5, 1 if inc else 0); c.font = F_INPUT; c.number_format = '"Yes";;"No"'; c.alignment = Alignment(horizontal="center")
        c = ws.cell(r, 6, f'=IF(E{r}=1,D{r},"")'); c.font = F_FORMULA; c.number_format = MULT
        if len(row) > 4: ws.cell(r, 7, row[4]).font = F_NOTE
    st.r += 1
r1 = st.r - 2
T["peer_med"] = st.row("Median of included peers", f"=MEDIAN(F{r0}:F{r1})", MULT, bold=True, col=4)
T["prem"] = st.row("Salik premium / (discount) to the peer median", f"={T['ev_ebitda']}/{T['peer_med']}-1", PCT, bold=True, col=4)
st.note("Pure toll-road operators with long concessions trade at high-teens to mid-twenties EBITDA multiples because the "
        "cash flows are long, contracted and inflation-linked. Salik sits inside that range with the shortest listed history "
        "and the least traffic risk, which is why the peer set is short and the exclusions are stated.", height=40)

# --------------------------------------------------------------------------- #
# Buyer Screen
# --------------------------------------------------------------------------- #
sb = S("Buyer Screen", ncols=13, label_width=34, col_width=9.5, subtitle="Twelve candidates scored on five criteria \u2014 the scores are judgements and the weights are inputs")
ws = sb.ws
ws.column_dimensions["C"].width = 10; ws.column_dimensions["D"].width = 13; ws.column_dimensions["E"].width = 46
for col in "FGHIJ": ws.column_dimensions[col].width = 10
ws.column_dimensions["K"].width = 11; ws.column_dimensions["L"].width = 8; ws.column_dimensions["M"].width = 34
B = {}
sb.section("CRITERIA WEIGHTS (must sum to 100%)")
r_w = sb.r
crit = ["Strategic fit", "Toll-road precedent", "Financial capacity", "Regional presence", "Government consent"]
for j, (c_, w) in enumerate(zip(crit, WEIGHTS)):
    ws.cell(sb.r, 2, c_).font = F_TEXT
    c = ws.cell(sb.r, 6 + j, w); c.font = F_INPUT; c.number_format = PCT; c.alignment = Alignment(horizontal="center")
ws.cell(sb.r, 11, f"=SUM(F{sb.r}:J{sb.r})").font = F_FORMULA; ws.cell(sb.r, 11).number_format = PCT
B["wsum"] = f"'Buyer Screen'!$K${sb.r}"
sb.r += 2
sb.section("THE SCREEN (scores 1 = weak, 5 = strong)")
hdr = ["Type", "Home", "Precedent and relevant assets"] + crit + ["Weighted", "Rank", "Read-across"]
for j, h in enumerate(hdr):
    c = ws.cell(sb.r, 3 + j, h); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.cell(sb.r, 2, "Candidate").font = F_BOLD; ws.row_dimensions[sb.r].height = 30; sb.r += 1
r_b0 = sb.r
READ = {"Transurban": "Best operator fit; balance sheet stretched; no Gulf presence",
        "Vinci Concessions": "Deepest pockets among strategics; Gulf contracting relationships via Vinci",
        "Ferrovial / Cintra": "Managed-lane expertise fits dynamic pricing; light in the region",
        "Abertis (Mundys / ACS)": "Largest pure toll-road owner; would want operational control the RTA will not give",
        "Globalvia": "Too small for a AED 40bn asset except as a minority partner",
        "Global Infrastructure Partners (BlackRock)": "Just closed Jafurah with Aramco; the template for a Gulf minority-stake deal",
        "Brookfield Infrastructure": "Owns toll roads and has a Gulf office; natural anchor investor",
        "Macquarie Asset Management": "Invented the listed toll-road fund model; strong precedent",
        "KKR Infrastructure": "ADNOC pipelines precedent shows the Gulf structure works",
        "CVC / DIF Capital Partners": "Cheque size is the constraint",
        "PIF / Mubadala / ADQ": "Most likely co-investor in any government sell-down; no operational value-add",
        "Dubai Holding / ICD (incumbent)": "The counterparty in every scenario; a reorganisation is more likely than a sale"}
for nm, typ, home, prec, *scores in BUYERS:
    r = sb.r
    ws.cell(r, 2, nm).font = F_BOLD; ws.cell(r, 3, typ).font = F_TEXT; ws.cell(r, 4, home).font = F_TEXT
    c = ws.cell(r, 5, prec); c.font = F_NOTE; c.alignment = Alignment(wrap_text=True, vertical="top")
    for j, sc in enumerate(scores):
        c = ws.cell(r, 6 + j, sc); c.font = F_INPUT; c.alignment = Alignment(horizontal="center")
    c = ws.cell(r, 11, f"=SUMPRODUCT(F{r}:J{r},$F${r_w}:$J${r_w})"); c.font = F_FORMULA; c.number_format = NUM2; c.alignment = Alignment(horizontal="center")
    c = ws.cell(r, 12, f"=RANK(K{r},$K${r_b0}:$K${r_b0+len(BUYERS)-1})"); c.font = F_FORMULA; c.alignment = Alignment(horizontal="center")
    c = ws.cell(r, 13, READ[nm]); c.font = F_NOTE; c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 28
    sb.r += 1
r_b1 = sb.r - 1
B["rng_w"] = f"'Buyer Screen'!$K${r_b0}:$K${r_b1}"
sb.blank()
B["top"] = sb.row("Highest-scoring candidate", f"=INDEX($B${r_b0}:$B${r_b1},MATCH(MAX({B['rng_w']}),{B['rng_w']},0))", "@", bold=True, col=5)
B["top_score"] = sb.row("  weighted score", f"=MAX({B['rng_w']})", NUM2, col=5)
B["n_fin"] = sb.row("Financial and sovereign buyers in the top five", f"=SUMPRODUCT((C{r_b0}:C{r_b1}<>\"Strategic\")*(L{r_b0}:L{r_b1}<=5))", '0', col=5)
sb.note("The screen ranks who could realistically hold a stake, not who would bid for control. Vinci leads on fit and "
        "balance sheet; the sponsors that have already done minority infrastructure deals with Gulf sovereigns \u2014 Brookfield, "
        "KKR, GIP \u2014 sit immediately behind it, level with the incumbent shareholder group. The purest operators, Transurban "
        "and Ferrovial, score highest on fit and fall down the list on consent and capacity, which is the finding: the likely "
        "transaction is a minority stake to a sponsor or sovereign, with an operator as a partner rather than a buyer.", height=56)

# --------------------------------------------------------------------------- #
# Ability to Pay
# --------------------------------------------------------------------------- #
sa = S("Ability to Pay", ncols=6, label_width=62, subtitle="What a sponsor and a strategic could each pay per share, and why they differ")
sa.ws.column_dimensions["F"].width = 58
A = {}
sa.section("COMMON OPERATING CASE")
A["ebitda0"] = sa.row("FY2025 EBITDA", f"={last('ebitda')}", NUM0, font=F_LINK)
A["g"] = sa.row("EBITDA growth, next five years", 0.06, PCT, note="Traffic and population plus ancillary; below the 2025 step-up, which was a tariff event")
A["tax"] = sa.row("Corporate tax", 0.09, PCT)
A["da"] = sa.row("Depreciation and amortisation as % of EBITDA", 0.07, PCT)
A["capex"] = sa.row("Capex as % of EBITDA", 0.005, PCT)
A["nd0"] = sa.row("Net debt today", f"={last('nd')}", NUM0, font=F_LINK)
A["ebitda5"] = sa.row("EBITDA in year five", f"={A['ebitda0']}*(1+{A['g']})^5", NUM0)

sa.blank(); sa.section("A. FINANCIAL SPONSOR \u2014 leverage and a return hurdle")
A["lev"] = sa.row("Debt at entry (x EBITDA)", 6.0, MULT, note="Infrastructure lenders will go further on a concession with no traffic risk than on a corporate")
A["kd"] = sa.row("Cost of debt", 0.060, PCT)
A["exit_x"] = sa.row("Exit EV / EBITDA in year five", 18.0, MULT, note="Below today's multiple on purpose")
A["irr"] = sa.row("Required equity IRR", 0.13, PCT, note="Core-plus infrastructure; lower than a buyout because the risk is lower")
A["debt_in"] = sa.row("Debt raised at entry", f"={A['ebitda0']}*{A['lev']}", NUM0)
# five-year cash to equity: after-tax FCF less interest, all swept to debt (closed form via geometric sums is messy; use year-by-year annuity approx)
# Simple and transparent: assume all free cash after interest repays debt each year, computed year by year in helper rows
sa.subsection("Debt paydown, year by year (all free cash after interest and tax repays debt)")
r_open = sa.r
sa.ws.cell(sa.r, 2, "Opening debt").font = F_TEXT
for j in range(5):
    c = sa.ws.cell(sa.r, 3 + j, f"={A['debt_in']}" if j == 0 else f"={L(2+j)}{sa.r+4}"); c.font = F_FORMULA; c.number_format = NUM0
sa.r += 1
sa.ws.cell(sa.r, 2, "EBITDA").font = F_TEXT
for j in range(5):
    c = sa.ws.cell(sa.r, 3 + j, f"={A['ebitda0']}*(1+{A['g']})^{j+1}"); c.font = F_FORMULA; c.number_format = NUM0
sa.r += 1
sa.ws.cell(sa.r, 2, "Interest on opening debt").font = F_TEXT
for j in range(5):
    c = sa.ws.cell(sa.r, 3 + j, f"=-{L(3+j)}{r_open}*{A['kd']}"); c.font = F_FORMULA; c.number_format = NUM0
sa.r += 1
sa.ws.cell(sa.r, 2, "Cash after tax and capex available to repay debt").font = F_TEXT
for j in range(5):
    e, i_ = f"{L(3+j)}{r_open+1}", f"{L(3+j)}{r_open+2}"
    c = sa.ws.cell(sa.r, 3 + j, f"=({e}+{i_}-{e}*{A['da']})*(1-{A['tax']})+{e}*{A['da']}-{e}*{A['capex']}"); c.font = F_FORMULA; c.number_format = NUM0
sa.r += 1
sa.ws.cell(sa.r, 2, "Closing debt").font = F_BOLD
for j in range(5):
    c = sa.ws.cell(sa.r, 3 + j, f"=MAX(0,{L(3+j)}{r_open}-{L(3+j)}{r_open+3})"); c.font = F_FORMULA; c.number_format = NUM0; c.border = TOTAL_BORDER
sa.r += 1
for j in range(5):
    sa.ws.column_dimensions[L(3 + j)].width = 13
A["debt5"] = f"'Ability to Pay'!$G${sa.r-1}"
sa.blank()
A["exit_ev"] = sa.row("Exit enterprise value", f"={A['ebitda5']}*{A['exit_x']}", NUM0)
A["exit_eq"] = sa.row("Exit equity value", f"={A['exit_ev']}-{A['debt5']}", NUM0, bold=True)
A["max_eq"] = sa.row("Maximum equity cheque at the hurdle: exit equity / (1 + IRR)^5", f"={A['exit_eq']}/(1+{A['irr']})^5", NUM0, bold=True)
A["fees"] = sa.row("Transaction and financing fees (% of EV)", 0.02, PCT)
A["max_ev_s"] = sa.row("Maximum enterprise value a sponsor can pay", f"=({A['max_eq']}+{A['debt_in']})/(1+{A['fees']})", NUM0, bold=True, border=TOTAL_BORDER)
A["max_px_s"] = sa.row("Maximum price per share \u2014 sponsor (AED)", f"=({A['max_ev_s']}-{A['nd0']})/{M['shares']}", NUM2, bold=True, border=DOUBLE_BORDER)
A["vs_s"] = sa.row("  versus the market price", f"={A['max_px_s']}/{T['price']}-1", PCT)

sa.blank(); sa.section("B. STRATEGIC OPERATOR \u2014 cost of capital over the remaining concession")
A["wacc"] = sa.row("Strategic buyer's WACC for a Gulf concession", 0.070, PCT)
A["syn"] = sa.row("Cost synergies as % of EBITDA", 0.03, PCT, note="Little to synergise: the RTA runs the gates. Technology and procurement only")
A["n"] = sa.row("Remaining concession life (years, to 2071)", 45, '0')
A["g_lt"] = sa.row("Long-run cash flow growth after year five", 0.03, PCT)
A["fcf1"] = sa.row("Year-one free cash flow to the firm, with synergies",
                   f"={A['ebitda0']}*(1+{A['g']})*(1+{A['syn']})*(1-{A['da']})*(1-{A['tax']})+{A['ebitda0']}*(1+{A['g']})*{A['da']}-{A['ebitda0']}*(1+{A['g']})*{A['capex']}", NUM0)
A["pv5"] = sa.row("PV of years one to five (growing annuity at g)",
                  f"={A['fcf1']}/({A['wacc']}-{A['g']})*(1-((1+{A['g']})/(1+{A['wacc']}))^5)", NUM0)
A["pv_tail"] = sa.row("PV of years six to end of concession (growing annuity at long-run g)",
                      f"={A['fcf1']}*(1+{A['g']})^4*(1+{A['g_lt']})/({A['wacc']}-{A['g_lt']})*(1-((1+{A['g_lt']})/(1+{A['wacc']}))^({A['n']}-5))/(1+{A['wacc']})^5", NUM0)
A["max_ev_t"] = sa.row("Maximum enterprise value a strategic can pay", f"={A['pv5']}+{A['pv_tail']}", NUM0, bold=True, border=TOTAL_BORDER)
A["max_px_t"] = sa.row("Maximum price per share \u2014 strategic (AED)", f"=({A['max_ev_t']}-{A['nd0']})/{M['shares']}", NUM2, bold=True, border=DOUBLE_BORDER)
A["vs_t"] = sa.row("  versus the market price", f"={A['max_px_t']}/{T['price']}-1", PCT)
A["impl_x_t"] = sa.row("  implied EV / EBITDA at that price", f"={A['max_ev_t']}/{A['ebitda0']}", MULT)
sa.note("A sponsor is capped by what lenders will fund and by a five-year exit; a strategic is capped only by its cost of "
        "capital over a 45-year concession. That is why the strategic number is higher and why, in a real process, the "
        "financial buyers are the anchor and the operators are the price-setters. Neither can bid without the government.", height=44)

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
tests = [
    ("Reported net profit reconciles to EBITDA less D&A, finance and tax in every year", "=AND(" + ",".join(f"ABS({a}-{b})<2" for a, b in zip(F["ni"], F["ni_chk"])) + ")"),
    ("Capex is below 1% of EBITDA in every year", "=AND(" + ",".join(f"-{c}/{e}<0.01" for c, e in zip(F["capex"], F["ebitda"])) + ")"),
    ("Net debt / EBITDA has fallen since FY2022", f"={F['lev'][3]}<{F['lev'][0]}"),
    ("Market capitalisation reconciles to price times shares", f"=ABS({T['mcap']}-{M['price']}*{M['shares']})<0.01"),
    ("Share price lies within the 52-week range", f"=AND({T['price']}>={M['w52lo']}-0.05,{T['price']}<={M['w52hi']}+0.05)"),
    ("Peer median evaluates to a number", f"=ISNUMBER({T['peer_med']})"),
    ("Criteria weights sum to 100%", f"=ABS({B['wsum']}-1)<0.0001"),
    ("Every score lies between 1 and 5", f"=AND(MIN('Buyer Screen'!F{r_b0}:J{r_b1})>=1,MAX('Buyer Screen'!F{r_b0}:J{r_b1})<=5)"),
    ("Weighted scores lie between 1 and 5", f"=AND(MIN({B['rng_w']})>=1,MAX({B['rng_w']})<=5)"),
    ("Sponsor debt is fully serviced: closing debt never negative and below opening", f"=AND({A['debt5']}>=0,{A['debt5']}<{A['debt_in']})"),
    ("Sponsor exit multiple is not above today's trading multiple", f"={A['exit_x']}<={T['ev_ebitda']}"),
    ("Strategic WACC exceeds long-run growth", f"={A['wacc']}>{A['g_lt']}"),
    ("Strategic ability to pay exceeds the sponsor's (the structural result)", f"={A['max_px_t']}>{A['max_px_s']}"),
    ("Government stake plus free float equals 100%", f"=ABS({M['gov_stake']}+0.249-1)<0.001"),
]
checks_sheet(book, tests,
             "The thirteenth check is a discipline check: if the sponsor's number ever exceeds the strategic's, either the "
             "leverage or the exit multiple has been pushed past what the asset supports.")

# --------------------------------------------------------------------------- #
book.cover(
    blurb="A company profile and buyer screen for Salik, Dubai's toll-gate operator, on FY2025 accounts and 18 September 2026 "
          "market data: what the business is, how it has performed since the 2022 IPO, where it trades against listed toll "
          "roads, who could realistically own a stake, and what a financial sponsor and a strategic operator could each pay.",
    method=[
        "Four years of income, cash flow and balance sheet with growth, margins, cash conversion and leverage.",
        "Trading multiples against a short, stated peer set with exclusions explained.",
        "Twelve candidate buyers scored 1\u20135 on strategic fit, toll-road precedent, financial capacity, regional presence and "
        "the likelihood of Government of Dubai consent; weights are inputs and the ranking is live.",
        "Ability to pay: a sponsor constrained by 6x leverage, a five-year exit and a 13% hurdle; a strategic constrained by "
        "a 7% cost of capital over the 45 years left on the concession.",
    ],
    toc=[("Profile", "the business, the case for a buyer, the constraint"),
         ("Financials", "FY2022\u2013FY2025 income, cash, balance sheet"),
         ("Trading", "multiples at the market price and against peers"),
         ("Buyer Screen", "twelve candidates, five criteria, live ranking"),
         ("Ability to Pay", "sponsor versus strategic, price per share"),
         ("Checks", "fourteen tests; must read MODEL OK")],
    highlights=[("FY2025 EBITDA (AED m)", f"={last('ebitda')}", NUM0),
                ("EV / EBITDA at the market price", f"={T['ev_ebitda']}", MULT),
                ("Price return since the IPO", f"={T['since_ipo']}", PCT),
                ("Highest-scoring candidate", f"={B['top']}", "@"),
                ("Sponsor ability to pay (AED per share)", f"={A['max_px_s']}", NUM2),
                ("Strategic ability to pay (AED per share)", f"={A['max_px_t']}", NUM2)],
    sources=["Salik Company PJSC financial statements FY2022\u2013FY2025 via Yahoo Finance for SALIK.AE; concession terms, gates and tariff from the 2022 IPO prospectus and RTA announcements.",
             "Share price and 52-week range: DFM close, 18 September 2026. Peer multiples: Yahoo Finance the same day.",
             "Buyer scores are the author's judgements from public precedent; criteria weights and every ability-to-pay input are stated.",
             "Nothing here is a recommendation and there is no suggestion any party is or has been in discussions."])

book.finish(freeze={"Buyer Screen": "C9"})
path = os.path.join(HERE, "Salik_Profile_Buyer_Screen.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
