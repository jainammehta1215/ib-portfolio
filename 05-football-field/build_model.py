"""
Build 05_football_field/TCS_Football_Field.xlsx — valuation summary for TCS.

Sheets: Cover · Football Field (chart + methodology page) · DCF (compact UFCF DCF in INR crore with WACC build) ·
Ranges (all methods normalised to INR per share) · Inputs · Checks
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE, NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"
book = Book(theme=THEMES["tcs"], project_no=5, project="Football Field & Valuation Summary", company="Tata Consultancy Services",
            units="INR crore unless stated · per-share values in INR", as_of="12 Sep 2026 · FY2026 annual report (Mar 2026)")

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
wi = book.sheet("Inputs", ncols=6, label_width=50, subtitle="Market data, DCF assumptions and the ranges brought in from Projects 3 and 4")
wi.column_dimensions["F"].width = 66
book.section(wi, 4, "MARKET DATA", 6)
mkt = [("Share price (INR)", 2200.8, NUM2, "NSE close, 12 Sep 2026"), ("52-week low (INR)", 1971.8, NUM2, "Yahoo Finance, 12 months to 12 Sep 2026"), ("52-week high (INR)", 3204.3, NUM2, ""),
       ("Analyst target low (INR)", 1800.0, NUM2, "Yahoo Finance consensus, 41 analysts"), ("Analyst target median (INR)", 2400.0, NUM2, ""), ("Analyst target high (INR)", 3480.0, NUM2, ""),
       ("Diluted shares (crore)", 361.81, NUM2, "FY2026 annual report; 361.81 crore = 3,618m")]
INP = {}
r = 5
for k, v, fmt, note in mkt:
    wi.cell(r, 2, k).font = F_TEXT; c = wi.cell(r, 3, v); c.font = F_INPUT; c.number_format = fmt; wi.cell(r, 6, note).font = F_NOTE; INP[k] = f"Inputs!$C${r}"; r += 1
r += 1; book.section(wi, r, "DCF ASSUMPTIONS (INR)", 6); r += 1
dcf_in = [("FY2026A revenue (INR crore)", 267021, NUM0, "FY2026 annual report (year to 31 Mar 2026)"), ("FY2026A EBIT (INR crore)", 66714, NUM0, "Reported operating profit before other income"),
          ("FY2026A D&A (INR crore)", 5560, NUM0, ""), ("FY2026A capex (INR crore)", 4146, NUM0, ""), ("Cash & investments (INR crore)", 41373, NUM0, "Cash, equivalents and current investments, 31 Mar 2026"),
          ("Debt incl. leases (INR crore)", 11283, NUM0, "Mostly lease liabilities"), ("Minority interest (INR crore)", 1238, NUM0, ""),
          ("Revenue growth FY2027E", 0.05, PCT, "Consensus FY2027 revenue INR 2.91 lakh crore (+9% in INR, ~5% constant currency); base 5%"),
          ("Revenue growth FY2028E–FY2031E (fading)", 0.045, PCT, "Fades linearly to terminal growth"), ("EBIT margin", 0.245, PCT, "FY2024–26: 24.6%, 25.9%, 25.0%; held at 24.5%"),
          ("Effective tax rate", 0.245, PCT, "FY2026: 24.5%"), ("D&A as % of revenue", 0.021, PCT, "FY2026: 2.1%"), ("Capex as % of revenue", 0.016, PCT, "FY2026: 1.6%"),
          ("Change in NWC as % of incremental revenue", 0.15, PCT, "Receivable-heavy business; FY2024–26 average"),
          ("Risk-free rate (India 10-year G-sec)", 0.069, PCT, "FRED INDIRLTLT01STM, Jun 2026: 6.89%"), ("Equity risk premium (India)", 0.065, PCT, "Damodaran mature-market ERP plus India country risk premium, 2026 (approx.)"),
          ("Beta", 0.90, '0.00', "TCS 5-year weekly beta vs Nifty 50 ≈ 0.85–0.95; low-beta defensive"), ("Pre-tax cost of debt", 0.075, PCT, "Lease-implicit rate; immaterial weight"),
          ("Debt weight", 0.02, PCT, "Book debt / market equity"), ("Terminal growth", 0.04, PCT, "Nominal INR; below long-run Indian nominal GDP"),
          ("Terminal EV / EBITDA exit multiple", 14.0, MULT, "Below TCS's current 10-year average (~18x) and the precedent median (14.7x)"),
          ("Mid-year convention (1 = on)", 1, "General", "")]
for k, v, fmt, note in dcf_in:
    wi.cell(r, 2, k).font = F_TEXT; c = wi.cell(r, 3, v); c.font = F_INPUT; c.number_format = fmt; wi.cell(r, 6, note).font = F_NOTE; INP[k] = f"Inputs!$C${r}"; r += 1
r += 1; book.section(wi, r, "RANGES FROM OTHER PROJECTS (INR per share)", 6); r += 1
ranges_in = [("Trading comps — 25th pct (Project 3, median across multiples)", 1652, NUM0, "Project 3 Summary: EV/EBITDA @25th = 1,652; range across methods 1,215–2,942"),
             ("Trading comps — median", 1848, NUM0, "Project 3 Summary: median across methods"), ("Trading comps — 75th pct", 2478, NUM0, "Project 3 Summary: EV/EBITDA @75th"),
             ("Precedents — 25th pct (Project 4, EV/Revenue)", 1428, NUM0, "Project 4 Summary; illustrative (targets much smaller than TCS)"),
             ("Precedents — median", 2026, NUM0, ""), ("Precedents — 75th pct", 2682, NUM0, "")]
for k, v, fmt, note in ranges_in:
    wi.cell(r, 2, k).font = F_TEXT; c = wi.cell(r, 3, v); c.font = F_INPUT; c.number_format = fmt; wi.cell(r, 6, note).font = F_NOTE; INP[k] = f"Inputs!$C${r}"; r += 1

# --------------------------------------------------------------------------- #
# DCF (compact)
# --------------------------------------------------------------------------- #
wd = book.sheet("DCF", ncols=9, label_width=46, subtitle="Compact unlevered FCF DCF in INR crore · FY2027E–FY2031E · WACC from CAPM with Indian inputs")
years = ["FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
book.section(wd, 4, "WACC", 9)
w_rows = [("Cost of equity = rf + β × ERP", f"={INP['Risk-free rate (India 10-year G-sec)']}+{INP['Beta']}*{INP['Equity risk premium (India)']}", PCT),
          ("After-tax cost of debt", f"={INP['Pre-tax cost of debt']}*(1-{INP['Effective tax rate']})", PCT),
          ("WACC", f"=(1-{INP['Debt weight']})*C5+{INP['Debt weight']}*C6", PCT)]
for i, (lab, f, fmt) in enumerate(w_rows, start=5):
    wd.cell(i, 2, lab).font = F_BOLD if lab == "WACC" else F_TEXT; c = wd.cell(i, 3, f); c.font = F_FORMULA; c.number_format = fmt
    if lab == "WACC": c.border = DOUBLE_BORDER
WACC = "DCF!$C$7"
book.section(wd, 9, "UNLEVERED FREE CASH FLOW (INR crore)", 9)
book.year_header(wd, 10, ["FY2026A"] + years, first_col=3)
D = {}
def row(r, key, label, f_hist, f_fc, fmt=NUM0, bold=False, border=None):
    D[key] = r; wd.cell(r, 2, label).font = F_BOLD if bold else F_TEXT
    c = wd.cell(r, 3, f_hist); c.font = F_LINK if f_hist and f_hist.startswith("=Inputs") else F_FORMULA; c.number_format = fmt
    for j in range(5):
        cl, p = L(4 + j), L(3 + j); c = wd.cell(r, 4 + j, f_fc(cl, p, j)); c.font = F_FORMULA; c.number_format = fmt
        if border: c.border = border
r = 11
row(r, "g", "Revenue growth", "", lambda c, p, j: (f"={INP['Revenue growth FY2027E']}" if j == 0 else f"={INP['Revenue growth FY2028E–FY2031E (fading)']}-({INP['Revenue growth FY2028E–FY2031E (fading)']}-{INP['Terminal growth']})*{j}/4"), fmt=PCT); r += 1
row(r, "rev", "Revenue", f"={INP['FY2026A revenue (INR crore)']}", lambda c, p, j: f"={p}{D['rev']}*(1+{c}{D['g']})"); r += 1
row(r, "ebit", "EBIT", f"={INP['FY2026A EBIT (INR crore)']}", lambda c, p, j: f"={c}{D['rev']}*{INP['EBIT margin']}"); r += 1
row(r, "tax", "  less: tax on EBIT", f"=-C{D['ebit']}*{INP['Effective tax rate']}", lambda c, p, j: f"=-{c}{D['ebit']}*{INP['Effective tax rate']}"); r += 1
row(r, "da", "  plus: D&A", f"={INP['FY2026A D&A (INR crore)']}", lambda c, p, j: f"={c}{D['rev']}*{INP['D&A as % of revenue']}"); r += 1
row(r, "capex", "  less: capex", f"=-{INP['FY2026A capex (INR crore)']}", lambda c, p, j: f"=-{c}{D['rev']}*{INP['Capex as % of revenue']}"); r += 1
row(r, "nwc", "  less: increase in NWC", "", lambda c, p, j: f"=-({c}{D['rev']}-{p}{D['rev']})*{INP['Change in NWC as % of incremental revenue']}"); r += 1
row(r, "ufcf", "Unlevered free cash flow", "", lambda c, p, j: f"={c}{D['ebit']}+{c}{D['tax']}+{c}{D['da']}+{c}{D['capex']}+{c}{D['nwc']}", bold=True, border=DOUBLE_BORDER); r += 1
row(r, "t", "Discount period (years, from 31 Mar 2026; valuation date ≈ 0.45y in)", "", lambda c, p, j: f"={j + 1}-0.45-0.5*{INP['Mid-year convention (1 = on)']}", fmt='0.00'); r += 1
row(r, "df", "Discount factor", "", lambda c, p, j: f"=1/(1+{WACC})^{c}{D['t']}", fmt='0.0000'); r += 1
row(r, "pv", "PV of UFCF", "", lambda c, p, j: f"={c}{D['ufcf']}*{c}{D['df']}", bold=True, border=TOTAL_BORDER); r += 2
book.section(wd, r, "VALUATION (INR crore; per share in INR)", 9); r += 1
for col_, lab in [("C", "Perpetuity growth"), ("E", "Exit multiple")]:
    c = wd[f"{col_}{r}"]; c.value = lab; c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center")
r += 1
V = {}
def vrow(r, key, label, f_pg, f_em, fmt=NUM0, bold=False, border=None):
    V[key] = r; wd.cell(r, 2, label).font = F_BOLD if bold else F_TEXT
    for col_, f in [("C", f_pg), ("E", f_em)]:
        c = wd[f"{col_}{r}"]; c.value = f; c.font = F_FORMULA; c.number_format = fmt
        if border: c.border = border
last = "H"; n_last = f"({last}{D['t']}+0.5*{INP['Mid-year convention (1 = on)']})"
vrow(r, "tv", "Terminal value at FY2031", f"={last}{D['ufcf']}*(1+{INP['Terminal growth']})/({WACC}-{INP['Terminal growth']})", f"=({last}{D['ebit']}+{last}{D['da']})*{INP['Terminal EV / EBITDA exit multiple']}"); r += 1
vrow(r, "pvtv", "PV of terminal value", f"=C{V['tv']}/(1+{WACC})^{n_last}", f"=E{V['tv']}/(1+{WACC})^{n_last}"); r += 1
vrow(r, "pvcf", "PV of forecast cash flows", f"=SUM(D{D['pv']}:H{D['pv']})", f"=C{r}"); r += 1
vrow(r, "ev", "Enterprise value", f"=C{V['pvtv']}+C{V['pvcf']}", f"=E{V['pvtv']}+E{V['pvcf']}", bold=True, border=TOTAL_BORDER); r += 1
vrow(r, "tvpct", "  terminal value % of EV", f"=C{V['pvtv']}/C{V['ev']}", f"=E{V['pvtv']}/E{V['ev']}", fmt=PCT); r += 1
vrow(r, "cash", "  plus: cash & investments", f"={INP['Cash & investments (INR crore)']}", f"={INP['Cash & investments (INR crore)']}"); r += 1
vrow(r, "debt", "  less: debt and minorities", f"=-{INP['Debt incl. leases (INR crore)']}-{INP['Minority interest (INR crore)']}", f"=-{INP['Debt incl. leases (INR crore)']}-{INP['Minority interest (INR crore)']}"); r += 1
vrow(r, "eq", "Equity value", f"=C{V['ev']}+C{V['cash']}+C{V['debt']}", f"=E{V['ev']}+E{V['cash']}+E{V['debt']}", bold=True, border=TOTAL_BORDER); r += 1
vrow(r, "px", "Value per share (INR)", f"=C{V['eq']}/{INP['Diluted shares (crore)']}", f"=E{V['eq']}/{INP['Diluted shares (crore)']}", fmt=NUM0, bold=True, border=DOUBLE_BORDER); r += 1
vrow(r, "impl", "Implied exit multiple / implied growth", f"=C{V['tv']}/({last}{D['ebit']}+{last}{D['da']})", f"=({WACC}*E{V['tv']}-{last}{D['ufcf']})/(E{V['tv']}+{last}{D['ufcf']})", fmt="General"); r += 1
wd[f"C{V['impl']}"].number_format = MULT; wd[f"E{V['impl']}"].number_format = PCT
r += 1; book.section(wd, r, "DCF RANGE FOR THE FOOTBALL FIELD (WACC ± 1%, per-share INR)", 9); r += 1
def px_at(w, method):
    pvcf = f"SUMPRODUCT(DCF!$D${D['ufcf']}:$H${D['ufcf']},1/(1+{w})^DCF!$D${D['t']}:$H${D['t']})"
    tv = f"DCF!$H${D['ufcf']}*(1+{INP['Terminal growth']})/({w}-{INP['Terminal growth']})" if method == "pg" else f"(DCF!$H${D['ebit']}+DCF!$H${D['da']})*{INP['Terminal EV / EBITDA exit multiple']}"
    return f"=({pvcf}+{tv}/(1+{w})^{n_last}+{INP['Cash & investments (INR crore)']}-{INP['Debt incl. leases (INR crore)']}-{INP['Minority interest (INR crore)']})/{INP['Diluted shares (crore)']}"
for j, hname in enumerate(["", "WACC −1%", "WACC base", "WACC +1%"]):
    c = wd.cell(r, 2 + j, hname); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center")
r += 1; R_DCFR = r
for lab, method in [("Perpetuity growth", "pg"), ("Exit multiple", "em")]:
    wd.cell(r, 2, lab).font = F_TEXT
    for j, w in enumerate([f"({WACC}-0.01)", WACC, f"({WACC}+0.01)"]):
        c = wd.cell(r, 3 + j, px_at(w, method)); c.font = F_FORMULA; c.number_format = NUM0
    r += 1
wd.cell(r, 2, "DCF low / mid / high used on the field").font = F_BOLD
wd.cell(r, 3, f"=MIN(C{R_DCFR}:E{R_DCFR+1})").number_format = NUM0; wd.cell(r, 4, f"=AVERAGE(D{R_DCFR}:D{R_DCFR+1})").number_format = NUM0; wd.cell(r, 5, f"=MAX(C{R_DCFR}:E{R_DCFR+1})").number_format = NUM0
for j in (3, 4, 5): wd.cell(r, j).font = F_FORMULA; wd.cell(r, j).border = TOTAL_BORDER
R_DCF_USE = r

# --------------------------------------------------------------------------- #
# Ranges + Football field
# --------------------------------------------------------------------------- #
wr = book.sheet("Ranges", ncols=8, label_width=40, col_width=13, subtitle="Every method expressed as INR per share · low / mid / high · weights for the recommended range")
book.section(wr, 4, "VALUATION RANGES", 8)
for j, hname in enumerate(["Method", "Low", "Mid", "High", "Weight", "Bar base (low)", "Bar length (high − low)"]):
    c = wr.cell(5, 2 + j, hname); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
methods = [("52-week trading range", f"={INP['52-week low (INR)']}", f"=AVERAGE(C6,E6)", f"={INP['52-week high (INR)']}", 0.0),
           ("Analyst price targets", f"={INP['Analyst target low (INR)']}", f"={INP['Analyst target median (INR)']}", f"={INP['Analyst target high (INR)']}", 0.10),
           ("Trading comparables (Project 3)", f"={INP['Trading comps — 25th pct (Project 3, median across multiples)']}", f"={INP['Trading comps — median']}", f"={INP['Trading comps — 75th pct']}", 0.35),
           ("Precedent transactions (Project 4, illustrative)", f"={INP['Precedents — 25th pct (Project 4, EV/Revenue)']}", f"={INP['Precedents — median']}", f"={INP['Precedents — 75th pct']}", 0.15),
           ("DCF (WACC ± 1%, both terminal methods)", f"=DCF!C{R_DCF_USE}", f"=DCF!D{R_DCF_USE}", f"=DCF!E{R_DCF_USE}", 0.40)]
for i, (lab, lo, mid, hi, w) in enumerate(methods, start=6):
    wr.cell(i, 2, lab).font = F_TEXT
    for j, f in enumerate([lo, mid, hi]):
        c = wr.cell(i, 3 + j, f); c.font = F_LINK; c.number_format = NUM0
    c = wr.cell(i, 6, w); c.font = F_INPUT; c.number_format = PCT
    wr.cell(i, 7, f"=C{i}").number_format = NUM0; wr.cell(i, 8, f"=E{i}-C{i}").number_format = NUM0
    wr.cell(i, 7).font = F_FORMULA; wr.cell(i, 8).font = F_FORMULA
wr.cell(11, 2, "Weights sum").font = F_NOTE; c = wr.cell(11, 6, "=SUM(F6:F10)"); c.font = F_FORMULA; c.number_format = PCT
book.section(wr, 13, "RECOMMENDED RANGE", 8)
rec = [("Weighted low", "=SUMPRODUCT(C6:C10,F6:F10)/SUM(F6:F10)"), ("Weighted mid", "=SUMPRODUCT(D6:D10,F6:F10)/SUM(F6:F10)"), ("Weighted high", "=SUMPRODUCT(E6:E10,F6:F10)/SUM(F6:F10)"),
       ("Current share price", f"={INP['Share price (INR)']}"), ("Weighted mid vs price", "=C15/C17-1")]
for i, (lab, f) in enumerate(rec, start=14):
    wr.cell(i, 2, lab).font = F_BOLD if "mid" in lab.lower() else F_TEXT; c = wr.cell(i, 3, f); c.font = F_FORMULA; c.number_format = PCT if "vs" in lab else NUM0
# Football field sheet with the chart
wf = book.sheet("Football Field", ncols=10, label_width=40, subtitle="Valuation summary · INR per share · bars show low to high, marker at mid")
wf.column_dimensions["B"].width = 44
chart = BarChart(); chart.type = "bar"; chart.grouping = "stacked"; chart.overlap = 100; chart.gapWidth = 60
chart.title = "TCS — valuation range by method (INR per share)"; chart.style = 10; chart.height = 9.5; chart.width = 22
cats = Reference(wr, min_col=2, min_row=6, max_row=10)
base = Reference(wr, min_col=7, min_row=5, max_row=10); length = Reference(wr, min_col=8, min_row=5, max_row=10)
chart.add_data(base, titles_from_data=True); chart.add_data(length, titles_from_data=True); chart.set_categories(cats)
s0 = chart.series[0]; s0.graphicalProperties.noFill = True; s0.graphicalProperties.line.noFill = True
s1 = chart.series[1]; s1.graphicalProperties.solidFill = book.theme.primary; s1.graphicalProperties.line.solidFill = book.theme.primary
chart.legend = None; chart.y_axis.title = "INR per share"; chart.y_axis.majorGridlines = None; chart.x_axis.scaling.orientation = "maxMin"
wf.add_chart(chart, "B4")
book.section(wf, 25, "METHODOLOGY AND THE ASSUMPTIONS THAT MOVE EACH BAR", 10)
meth = [("52-week range", "Where the stock has traded in the last year; context, not valuation. Zero weight."),
        ("Analyst targets", "Sell-side consensus (41 analysts, median INR 2,400). Reflects the same information the other bars use; low weight to avoid double-counting."),
        ("Trading comparables", "Peer interquartile range of EV/EBITDA, EV/EBIT and P/E applied to TCS (Project 3). Moves with sector sentiment; the 2026 de-rating is the reason the bar sits near the share price. Highest weight among market methods."),
        ("Precedent transactions", "Control-transaction multiples 2014–2024 (Project 4). Illustrative: targets were a fraction of TCS's size and the deals were struck in a higher-multiple period; low weight."),
        ("DCF", "Intrinsic value at WACC ≈12.6% ± 1% (India G-sec 6.9%, ERP 6.5%, β 0.9), terminal growth 4% or 14× exit. The bar is wide because terminal value is ~70% of EV; the WACC and the terminal assumption dominate. Highest weight.")]
for i, (k, v) in enumerate(meth, start=26):
    wf.cell(i, 2, k).font = F_BOLD; c = wf.cell(i, 3, v); c.font = F_TEXT; c.alignment = Alignment(wrap_text=True, vertical="top")
    wf.merge_cells(start_row=i, start_column=3, end_row=i, end_column=10); wf.row_dimensions[i].height = 30
book.section(wf, 32, "RECOMMENDED RANGE AND READING", 10)
wf.cell(33, 2, "Recommended range (weighted low – high, INR)").font = F_BOLD; c = wf.cell(33, 3, '=TEXT(Ranges!C14,"#,##0")&" – "&TEXT(Ranges!C16,"#,##0")&"   (mid "&TEXT(Ranges!C15,"#,##0")&")"'); c.font = F_FORMULA
wf.cell(34, 2, "Current share price (INR)").font = F_TEXT; c = wf.cell(34, 3, f"={INP['Share price (INR)']}"); c.font = F_LINK; c.number_format = NUM0
wf.cell(35, 2, "Weighted mid vs price").font = F_TEXT; c = wf.cell(35, 3, "=Ranges!C18"); c.font = F_LINK; c.number_format = PCT
read = ["The methods answer different questions. Comps say what the market pays today for businesses like TCS; the DCF says what the cash flows are worth at a required return; precedents say what a buyer paid for control. When the bars overlap the share price, as here, the market is pricing TCS roughly in line with all three, and the case for a large discount or premium rests on disagreeing with a specific assumption (the ERP, the terminal growth, the sector's de-rating).",
        "The DCF and trading-comps bars straddle the price from opposite sides: comps sit slightly below (the sector de-rated in 2026), the DCF slightly above (TCS's cash generation supports the price at a 12–13% Indian discount rate). That is the tension a banker would present to a board.",
        "Weights are judgement and are shown as inputs. Change them and the recommended range moves; the point of the page is to make that judgement explicit."]
for i, m in enumerate(read, start=37):
    c = wf.cell(i, 2, "•  " + m); c.font = F_TEXT; c.alignment = Alignment(wrap_text=True, vertical="top")
    wf.merge_cells(start_row=i, start_column=2, end_row=i, end_column=10); wf.row_dimensions[i].height = 48

wk = book.sheet("Checks", ncols=4, label_width=60, subtitle="Integrity checks · should all be TRUE")
chk = [("Weights sum to 100%", "=ROUND(Ranges!F11,6)=1"), ("Every range is ordered low ≤ mid ≤ high", "=AND(SUMPRODUCT(--(Ranges!C6:C10<=Ranges!D6:D10))=5,SUMPRODUCT(--(Ranges!D6:D10<=Ranges!E6:E10))=5)"),
       ("WACC above terminal growth", f"={WACC}>{INP['Terminal growth']}"), ("DCF terminal value share of EV between 50% and 90%", f"=AND(DCF!C{V['tvpct']}>0.5,DCF!C{V['tvpct']}<0.9)"),
       ("Implied exit multiple from perpetuity method between 8x and 30x", f"=AND(DCF!C{V['impl']}>8,DCF!C{V['impl']}<30)")]
for i, (lab, f) in enumerate(chk, start=4):
    wk.cell(i, 2, lab).font = F_TEXT; c = wk.cell(i, 3, f); c.font = F_FORMULA
book.section(wk, 10, "OVERALL", 4); c = wk.cell(10, 3, '=IF(COUNTIF(C4:C8,FALSE)=0,"MODEL OK","CHECK ERRORS")'); c.font = Font(name=FONT, size=11, bold=True, color=book.theme.text_on_primary)

# order: Football Field, Ranges, DCF, Inputs, Checks (cover is inserted at the front afterwards)
wbk = book.wb; order = ["Football Field", "Ranges", "DCF", "Inputs", "Checks"]
wbk._sheets = [wbk[n] for n in order]
book.cover(
    toc=[("Football Field", "the chart, methodology page and recommended range"), ("Ranges", "every method as INR per share, weights, recommended range"),
         ("DCF", "compact UFCF DCF in INR crore with Indian WACC inputs and a WACC ± 1% range"), ("Inputs", "market data, DCF assumptions, ranges from Projects 3 and 4"), ("Checks", "integrity checks")],
    highlights=[("Recommended range — low (INR)", "=Ranges!C14", NUM0), ("Recommended range — mid (INR)", "=Ranges!C15", NUM0), ("Recommended range — high (INR)", "=Ranges!C16", NUM0),
                ("Share price (INR)", f"={INP['Share price (INR)']}", NUM0), ("DCF WACC", f"={WACC}", PCT), ("Model status", "=Checks!C10", "General")],
    blurb="The valuation summary for TCS that brings Projects 2–4 together: a compact rupee DCF with Indian cost-of-capital inputs, the trading-comparables range, the precedent-transactions range, the 52-week trading range and analyst targets, all expressed per share, drawn as a football field and weighted into a recommended range with the weights shown as inputs.",
    method=["Each method is normalised to INR per share at the same date and share count before it is drawn.",
            "The DCF bar spans WACC ± 1% across both terminal methods; the comps and precedent bars span the peer interquartile range.",
            "Weights are explicit and editable; the recommended range is their weighted low / mid / high.",
            "The floating-bar chart is a stacked bar with an invisible base series, so it renders in Excel, Google Sheets and LibreOffice without add-ins."],
    sources=["TCS FY2026 annual report; NSE prices; Yahoo Finance 52-week range and analyst targets (12 Sep 2026)", "FRED INDIRLTLT01STM (India 10-year), Damodaran country risk premium", "Projects 3 and 4 for the market-based ranges"])
book.finish(freeze={"DCF": "C11"}, repeat_rows={"DCF": "1:2"}, portrait=("Cover",))
out = os.path.join(HERE, "TCS_Football_Field.xlsx"); book.save(out)
rc = recalc(out); print("recalc:", rc.get("status"), rc.get("total_errors"), rc.get("error_summary"))
pdf = export_pdf(out); png = preview_png(pdf, 0, dpi=60); os.rename(png, os.path.join(HERE, "cover.png")); print("done")
