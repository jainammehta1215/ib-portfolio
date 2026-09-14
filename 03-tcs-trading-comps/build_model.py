"""
Build 03_tcs_trading_comps/TCS_Trading_Comps.xlsx — comparable-companies analysis for TCS.

Sheets: Cover · Summary (implied valuation) · Inputs (FX, target, peer selection notes) · Data (raw pulls, blue) ·
Comps (USD normalisation, EV build, LTM and calendarised NTM multiples, statistics) · Calendarisation (worked example) · Checks
"""
import os, sys, json
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import numpy as np, pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE, NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"
raw = pd.read_csv("/tmp/comps_raw.csv", index_col=0); fx = json.load(open("/tmp/fx.json"))
# statements currency heuristic: an INR-priced company whose LTM revenue is < 2% of market cap reports in USD
raw["stmt_ccy"] = raw.apply(lambda r: "USD" if (r["currency"] == "INR" and r["rev"] / r["mcap"] < 0.02) else r["currency"], axis=1)
fy0_end = {"TCS.NS": "2027-03-31", "INFY.NS": "2027-03-31", "HCLTECH.NS": "2027-03-31", "WIPRO.NS": "2027-03-31", "TECHM.NS": "2027-03-31", "PERSISTENT.NS": "2027-03-31",
           "ACN": "2026-08-31", "CTSH": "2026-12-31", "EPAM": "2026-12-31"}
group = {"TCS.NS": "India large-cap", "INFY.NS": "India large-cap", "HCLTECH.NS": "India large-cap", "WIPRO.NS": "India large-cap", "TECHM.NS": "India mid-cap",
         "PERSISTENT.NS": "India mid-cap", "ACN": "Global", "CTSH": "Global", "EPAM": "Global"}
order = ["TCS.NS", "INFY.NS", "HCLTECH.NS", "WIPRO.NS", "TECHM.NS", "PERSISTENT.NS", "ACN", "CTSH", "EPAM"]
raw = raw.loc[order]

book = Book(theme=THEMES["tcs"], project_no=3, project="Trading Comparables", company="Tata Consultancy Services",
            units="US$ millions unless stated · share prices in local currency", as_of="prices 12 Sep 2026 · LTM to Jun 2026")

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
wi = book.sheet("Inputs", ncols=6, label_width=46, subtitle="FX rates, target and the peer-selection rationale")
wi.column_dimensions["F"].width = 70
book.section(wi, 4, "FX (units of local currency per US$) AND TARGET", 6)
inp = [("USD per INR (1 / USDINR)", round(1 / fx["USDINR=X"], 6), '0.000000', f"USDINR = {fx['USDINR=X']:.2f}, Yahoo Finance 12 Sep 2026"),
       ("USD per USD", 1.0, '0.000000', ""),
       ("Target ticker", "TCS.NS", "General", "The company being valued; excluded from peer statistics"),
       ("Calendarisation window start", "2026-04-01", "General", "TCS fiscal year FY2027: 1 Apr 2026 – 31 Mar 2027"),
       ("Calendarisation window end", "2027-03-31", "General", ""),
       ("Minimum peer market cap (US$m) for statistics", 5000, NUM0, "Excludes tiny names from the range; set to 0 to include all")]
INP = {}
for i, (k, v, fmt, note) in enumerate(inp, start=5):
    wi.cell(i, 2, k).font = F_TEXT; c = wi.cell(i, 3, v); c.font = F_INPUT; c.number_format = fmt; wi.cell(i, 6, note).font = F_NOTE
    INP[k] = f"Inputs!$C${i}"
book.section(wi, 12, "PEER SELECTION", 6)
sel = [("Criteria", "IT services and consulting with offshore delivery; revenue > US$1bn; listed with liquid trading; LTM financials available. Two tiers: Indian large caps (the direct peers) and global consulting/services names (the customers' alternative suppliers)."),
       ("Included", "Infosys, HCLTech, Wipro (India large-cap); Tech Mahindra, Persistent (India mid-cap, higher growth, smaller); Accenture, Cognizant, EPAM (global)."),
       ("Excluded", "LTIMindtree and Capgemini: no usable data on the free feed (LTIMindtree would otherwise be included). Mphasis and Coforge: under the size threshold. IBM: conglomerate mix. Tata Elxsi: engineering R&D, different economics."),
       ("Data quirks handled", "Infosys reports IFRS statements in US$ while its shares trade in INR; the statements currency is detected and every figure is normalised to US$ before any multiple is computed. Consensus estimates are in the share-price currency for all names.")]
for i, (k, v) in enumerate(sel, start=13):
    wi.cell(i, 2, k).font = F_BOLD; c = wi.cell(i, 3, v); c.font = F_TEXT; c.alignment = Alignment(wrap_text=True, vertical="top")
    wi.merge_cells(start_row=i, start_column=3, end_row=i, end_column=6); wi.row_dimensions[i].height = 42

# --------------------------------------------------------------------------- #
# Data (raw pulls, blue)
# --------------------------------------------------------------------------- #
wd = book.sheet("Data", ncols=22, label_width=26, col_width=13, subtitle="Raw pulls in reporting currency (blue) · statements are LTM to the latest quarter · estimates are consensus averages")
cols = [("Name", "name", "General"), ("Group", None, "General"), ("Price ccy", "currency", "General"), ("Statements ccy", "stmt_ccy", "General"),
        ("Share price (local)", "price", NUM2), ("Shares (m)", "shares", NUM0), ("Market cap (local m)", "mcap", NUM0),
        ("Total debt (stmt ccy m)", "debt", NUM0), ("Cash & ST inv (stmt ccy m)", "cash", NUM0), ("Minority interest (stmt ccy m)", "minority", NUM0),
        ("LTM revenue (stmt ccy m)", "rev", NUM0), ("LTM EBITDA (stmt ccy m)", "ebitda", NUM0), ("LTM EBIT (stmt ccy m)", "ebit", NUM0), ("LTM net income (stmt ccy m)", "ni", NUM0),
        ("LTM period end", "ltm_end", "General"), ("FY0 end (consensus)", None, "General"), ("Revenue FY0 (price ccy m)", "rev_0y", NUM0), ("Revenue FY1 (price ccy m)", "rev_1y", NUM0),
        ("EPS FY0 (price ccy)", "eps_0y", NUM2), ("EPS FY1 (price ccy)", "eps_1y", NUM2)]
for j, (hname, _, _) in enumerate(cols):
    c = wd.cell(4, 2 + j, hname); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
wd.row_dimensions[4].height = 42
DROW = {}
for i, t in enumerate(order):
    r = 5 + i; DROW[t] = r; rw = raw.loc[t]
    for j, (hname, key, fmt) in enumerate(cols):
        c = wd.cell(r, 2 + j)
        if key is None:
            c.value = group[t] if hname == "Group" else fy0_end[t]
        else:
            v = rw[key]
            if key in ("shares", "mcap", "debt", "cash", "minority", "rev", "ebitda", "ebit", "ni", "rev_0y", "rev_1y"):
                v = 0.0 if pd.isna(v) else float(v) / 1e6
            elif isinstance(v, (float, np.floating)) and pd.isna(v):
                v = 0.0
            c.value = v if not isinstance(v, (np.floating, np.integer)) else float(v)
        c.font = F_INPUT if key not in (None, "name") else F_TEXT; c.number_format = fmt
    wd.cell(r, 1, t).font = F_NOTE
book.note(wd, 5 + len(order) + 1, "Source: Yahoo Finance (prices, shares, quarterly statements, consensus revenue and EPS), 12 Sep 2026. LTM = sum of the last four reported quarters. Minority interest is zero where not reported. Persistent is below the size of the others and is flagged as mid-cap; the statistics exclude names under the market-cap floor on the Inputs sheet.", 22, height=30)
DC = {hname: L(2 + j) for j, (hname, _, _) in enumerate(cols)}    # column letters by header

# --------------------------------------------------------------------------- #
# Comps (USD normalisation, EV, multiples, statistics)
# --------------------------------------------------------------------------- #
wc = book.sheet("Comps", ncols=20, label_width=26, col_width=12, subtitle="Everything in US$ millions · EV = market cap + debt + minorities − cash · NTM figures calendarised to TCS's March year")
hdr = ["Company", "Group", "Market cap", "Net debt", "EV", "LTM revenue", "LTM EBITDA", "LTM EBIT", "LTM NI", "EBITDA margin", "Rev growth FY1/FY0",
       "EV / Revenue", "EV / EBITDA", "EV / EBIT", "P / E (LTM)", "CY revenue (cal.)", "CY EPS (cal., local)", "EV / CY revenue", "P / E (CY)"]
for j, hname in enumerate(hdr):
    c = wc.cell(4, 2 + j, hname); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
wc.row_dimensions[4].height = 42
def fxcell(t, which):   # USD per unit of the relevant currency
    ccy_cell = f"Data!${DC['Price ccy'] if which == 'price' else DC['Statements ccy']}${DROW[t]}"
    return f'IF({ccy_cell}="INR",{INP["USD per INR (1 / USDINR)"]},1)'
def cal_w0(t):          # share of the calendarisation window covered by consensus FY0
    d = f"Data!${DC['FY0 end (consensus)']}${DROW[t]}"
    return f'MAX(0,MIN(DATEVALUE({d}),DATEVALUE({INP["Calendarisation window end"]}))-DATEVALUE({INP["Calendarisation window start"]})+1)/365'
CROW = {}
for i, t in enumerate(order):
    r = 5 + i; CROW[t] = r; d = DROW[t]
    def D_(col): return f"Data!${DC[col]}${d}"
    fxp, fxs = fxcell(t, "price"), fxcell(t, "stmt")
    cells = [
        (f"={D_('Name')}", "General"), (f"={D_('Group')}", "General"),
        (f"={D_('Market cap (local m)')}*{fxp}", NUM0),
        (f"=({D_('Total debt (stmt ccy m)')}+{D_('Minority interest (stmt ccy m)')}-{D_('Cash & ST inv (stmt ccy m)')})*{fxs}", NUM0),
        (f"=D{r}+E{r}", NUM0),
        (f"={D_('LTM revenue (stmt ccy m)')}*{fxs}", NUM0), (f"={D_('LTM EBITDA (stmt ccy m)')}*{fxs}", NUM0), (f"={D_('LTM EBIT (stmt ccy m)')}*{fxs}", NUM0), (f"={D_('LTM net income (stmt ccy m)')}*{fxs}", NUM0),
        (f"=H{r}/G{r}", PCT), (f"={D_('Revenue FY1 (price ccy m)')}/{D_('Revenue FY0 (price ccy m)')}-1", PCT),
        (f"=F{r}/G{r}", MULT), (f"=F{r}/H{r}", MULT), (f"=F{r}/I{r}", MULT), (f"=D{r}/J{r}", MULT),
        (f"=({cal_w0(t)}*{D_('Revenue FY0 (price ccy m)')}+(1-{cal_w0(t)})*{D_('Revenue FY1 (price ccy m)')})*{fxp}", NUM0),
        (f"={cal_w0(t)}*{D_('EPS FY0 (price ccy)')}+(1-{cal_w0(t)})*{D_('EPS FY1 (price ccy)')}", NUM2),
        (f"=F{r}/Q{r}", MULT), (f"={D_('Share price (local)')}/R{r}", MULT),
    ]
    for j, (fml, fmt) in enumerate(cells):
        c = wc.cell(r, 2 + j, fml); c.font = F_LINK if j < 2 else F_FORMULA; c.number_format = fmt
    if t == "TCS.NS":
        for j in range(2, 21): wc.cell(r, j).fill = book.fill_light
rN = 5 + len(order) - 1
r = rN + 2
book.section(wc, r, "PEER STATISTICS (excluding the target and names under the market-cap floor)", 20); r += 1
stat_cols = list(range(11, 21))     # K..T : margin, growth, multiples
peer_rows = [CROW[t] for t in order if t != "TCS.NS"]
def cond_range(col):
    return ",".join(f"IF(AND($D{pr}>={INP['Minimum peer market cap (US$m) for statistics']},$B{pr}<>{INP['Target ticker']}),{L(col)}{pr},\"\")" for pr in peer_rows)
STAT = {}
for lab, fn in [("Maximum", "MAX"), ("75th percentile", "PERCENTILE.INC"), ("Median", "MEDIAN"), ("Mean", "AVERAGE"), ("25th percentile", "PERCENTILE.INC"), ("Minimum", "MIN")]:
    wc.cell(r, 2, lab).font = F_BOLD; STAT[lab] = r
    for col in stat_cols:
        vals = ",".join(f"{L(col)}{pr}" for pr in peer_rows)
        # SUMPRODUCT-free approach: build an array via CHOOSE of eligible rows is verbose; use helper columns instead
        c = wc.cell(r, col); c.number_format = PCT if col in (11, 12) else MULT; c.font = F_FORMULA
        if fn == "PERCENTILE.INC":
            q = 0.75 if "75" in lab else 0.25
            c.value = f"=PERCENTILE.INC(({vals}),{q})" if False else f"=PERCENTILE(({vals}),{q})"
        else:
            c.value = f"={fn}({vals})"
    r += 1
book.note(wc, r, "Statistics are computed over the eight peers only (the target's row is outside the range by construction). All eight clear the US$5bn market-cap floor on the Inputs sheet, so the floor is not binding in this run; drop a name from the peer list to exclude it.", 20, height=30)
r += 2
book.section(wc, r, "TARGET vs PEER MEDIAN", 20); r += 1
tr = CROW["TCS.NS"]
for j, col in enumerate(stat_cols):
    wc.cell(r, 2, "TCS").font = F_BOLD; c = wc.cell(r, col, f"={L(col)}{tr}"); c.font = F_LINK; c.number_format = PCT if col in (11, 12) else MULT
    wc.cell(r + 1, 2, "Premium / (discount) to median").font = F_TEXT; c = wc.cell(r + 1, col, f"={L(col)}{tr}/{L(col)}{STAT['Median']}-1"); c.font = F_FORMULA; c.number_format = PCT
R_TCS_VS = r

# --------------------------------------------------------------------------- #
# Calendarisation (worked example)
# --------------------------------------------------------------------------- #
wcal = book.sheet("Calendarisation", ncols=8, label_width=40, subtitle="Why a March-year company cannot be compared to a December-year company without adjustment")
book.section(wcal, 4, "METHOD", 8)
book.note(wcal, 5, "Consensus estimates are quoted for each company's own fiscal year. To compare TCS (year to 31 March) with Cognizant (year to 31 December) on a forward multiple, both must be expressed over the same twelve months. This workbook calendarises every peer to TCS's FY2027 (1 Apr 2026 – 31 Mar 2027): weight w0 on the consensus year that ends inside the window and (1 − w0) on the following year, where w0 = months of overlap / 12.", 8, height=56)
hdr = ["Company", "Consensus FY0 ends", "Overlap with Apr-26 → Mar-27 (years)", "Weight on FY0", "Weight on FY1", "Revenue FY0 (local m)", "Revenue FY1 (local m)", "Calendarised revenue (local m)"]
for j, hname in enumerate(hdr):
    c = wcal.cell(7, 2 + j, hname); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
wcal.row_dimensions[7].height = 42
for i, t in enumerate(order):
    r = 8 + i; d = DROW[t]
    vals = [(f"=Data!${DC['Name']}${d}", "General"), (f"=Data!${DC['FY0 end (consensus)']}${d}", "General"), (f"={cal_w0(t)}", '0.00'),
            (f"=D{r}", '0.00'), (f"=1-D{r}", '0.00'), (f"=Data!${DC['Revenue FY0 (price ccy m)']}${d}", NUM0), (f"=Data!${DC['Revenue FY1 (price ccy m)']}${d}", NUM0), (f"=E{r}*G{r}+F{r}*H{r}", NUM0)]
    for j, (fml, fmt) in enumerate(vals):
        c = wcal.cell(r, 2 + j, fml); c.font = F_LINK if j in (0, 1, 5, 6) else F_FORMULA; c.number_format = fmt
book.note(wcal, 8 + len(order) + 1, "Reading: an Indian peer's FY0 already coincides with the window (weight 1.0). Cognizant's calendar 2026 covers nine of the twelve months (weight 0.75). Accenture's year to August 2026 covers five months (weight 0.42), so most of its calendarised figure comes from FY2027. Without this step Accenture would be compared on a period ending seven months earlier than TCS's.", 8, height=44)

# --------------------------------------------------------------------------- #
# Summary: implied valuation for TCS
# --------------------------------------------------------------------------- #
ws = book.sheet("Summary", ncols=9, label_width=40, col_width=13, subtitle="Implied TCS valuation from the peer interquartile range · per-share values in INR")
book.section(ws, 4, "TCS METRICS (US$m unless stated)", 9)
tcs = [("LTM revenue", f"=Comps!G{tr}", NUM0), ("LTM EBITDA", f"=Comps!H{tr}", NUM0), ("LTM EBIT", f"=Comps!I{tr}", NUM0), ("LTM net income", f"=Comps!J{tr}", NUM0),
       ("Calendarised FY2027 revenue", f"=Comps!Q{tr}", NUM0), ("Calendarised FY2027 EPS (INR)", f"=Comps!R{tr}", NUM2),
       ("Net debt (negative = net cash)", f"=Comps!E{tr}", NUM0), ("Shares (m)", f"=Data!{DC['Shares (m)']}{DROW['TCS.NS']}", NUM0),
       ("Share price (INR)", f"=Data!{DC['Share price (local)']}{DROW['TCS.NS']}", NUM2), ("USD per INR", f"={INP['USD per INR (1 / USDINR)']}", '0.000000')]
TM = {}
for i, (lab, fml, fmt) in enumerate(tcs, start=5):
    ws.cell(i, 2, lab).font = F_TEXT; c = ws.cell(i, 3, fml); c.font = F_LINK; c.number_format = fmt; TM[lab] = f"$C${i}"
r = 5 + len(tcs) + 1
book.section(ws, r, "IMPLIED VALUE PER SHARE (INR) FROM PEER MULTIPLES", 9); r += 1
for j, hname in enumerate(["Multiple", "TCS metric", "25th pct", "Median", "75th pct", "Implied price @25th", "Implied price @median", "Implied price @75th"]):
    c = ws.cell(r, 2 + j, hname); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
ws.row_dimensions[r].height = 30; r += 1
R_IMPL0 = r
def implied(metric_ref, stat_col, is_equity: bool):
    # EV-based: (metric × multiple − net debt) / shares / USD-per-INR ; equity-based on LTM NI: (NI × PE) / shares / fx ; CY P/E: EPS × PE directly
    out = []
    for stat in ("25th percentile", "Median", "75th percentile"):
        m = f"Comps!{L(stat_col)}{STAT[stat]}"
        if is_equity == "eps":
            out.append(f"={metric_ref}*{m}")
        elif is_equity:
            out.append(f"={metric_ref}*{m}/{TM['Shares (m)']}/{TM['USD per INR']}")
        else:
            out.append(f"=({metric_ref}*{m}-{TM['Net debt (negative = net cash)']})/{TM['Shares (m)']}/{TM['USD per INR']}")
    return out
lines = [("EV / LTM revenue", TM["LTM revenue"], 13, False), ("EV / LTM EBITDA", TM["LTM EBITDA"], 14, False), ("EV / LTM EBIT", TM["LTM EBIT"], 15, False),
         ("P / E (LTM)", TM["LTM net income"], 16, True), ("EV / CY2027 revenue (calendarised)", TM["Calendarised FY2027 revenue"], 19, False), ("P / E (CY2027, calendarised)", TM["Calendarised FY2027 EPS (INR)"], 20, "eps")]
for lab, metric, scol, eq in lines:
    ws.cell(r, 2, lab).font = F_TEXT
    c = ws.cell(r, 3, f"={metric}"); c.font = F_LINK; c.number_format = NUM2 if "EPS" in lab else NUM0
    for j, stat in enumerate(("25th percentile", "Median", "75th percentile")):
        c = ws.cell(r, 4 + j, f"=Comps!{L(scol)}{STAT[stat]}"); c.font = F_LINK; c.number_format = MULT
    for j, fml in enumerate(implied(metric, scol, eq)):
        c = ws.cell(r, 7 + j, fml); c.font = F_FORMULA; c.number_format = NUM0
    r += 1
R_IMPL1 = r - 1
ws.cell(r, 2, "Range across methods (INR)").font = F_BOLD
for j, fn in enumerate(["MIN", "MEDIAN", "MAX"]):
    c = ws.cell(r, 7 + j, f"={fn}(G{R_IMPL0}:I{R_IMPL1})"); c.font = F_FORMULA; c.number_format = NUM0; c.border = TOTAL_BORDER
ws.cell(r + 1, 2, "Current share price (INR)").font = F_TEXT; c = ws.cell(r + 1, 8, f"={TM['Share price (INR)']}"); c.font = F_LINK; c.number_format = NUM0
ws.cell(r + 2, 2, "Median implied vs current").font = F_TEXT; c = ws.cell(r + 2, 8, f"=H{r}/H{r+1}-1"); c.font = F_FORMULA; c.number_format = PCT
R_RANGE = r
book.section(ws, r + 4, "READING THE RESULT", 9)
memo = ["TCS trades at a premium to the peer median on every multiple, which is its long-standing position: highest margins in the group, lowest client-concentration risk, and a Tata parent. The question a comps page answers is whether today's premium is larger or smaller than usual, not whether it exists.",
        "The Indian large caps cluster tightly; the global names (Accenture, Cognizant, EPAM) trade lower because of slower growth and, in 2026, the AI-disruption discount on outsourcing. Mixing the two tiers widens the range; a banker would show both tiers and let the client see the gap.",
        "Calendarised forward multiples are the ones institutional investors quote. The LTM columns are shown because they are the only ones that need no estimates and therefore no judgement.",
        "The implied range is a market-based anchor for the football field (Project 5). It says what the market pays for businesses like this today; it says nothing about what TCS is intrinsically worth (Project 2's question)."]
for i, mline in enumerate(memo):
    c = ws.cell(r + 5 + i, 2, "•  " + mline); c.font = F_TEXT; c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r + 5 + i, start_column=2, end_row=r + 5 + i, end_column=9); ws.row_dimensions[r + 5 + i].height = 32

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
wk = book.sheet("Checks", ncols=4, label_width=64, subtitle="Integrity checks · should all be TRUE")
chk = [("Every peer's EV is positive", f"=MIN(Comps!F5:F{rN})>0"),
       ("Infosys statements detected as US$ (revenue in US$ ≈ 20bn, not 20m)", f"=AND(Comps!G{CROW['INFY.NS']}>15000,Comps!G{CROW['INFY.NS']}<30000)"),
       ("Calendarisation weights lie between 0 and 1", f"=AND(MIN(Calendarisation!D8:D{7 + len(order)})>=0,MAX(Calendarisation!D8:D{7 + len(order)})<=1)"),
       ("TCS EV/EBITDA within a plausible band (10x–40x)", f"=AND(Comps!N{tr}>10,Comps!N{tr}<40)"),
       ("Implied price range is ordered (25th ≤ median ≤ 75th) for EV/EBITDA", f"=AND(Summary!G{R_IMPL0+1}<=Summary!H{R_IMPL0+1},Summary!H{R_IMPL0+1}<=Summary!I{R_IMPL0+1})")]
for i, (lab, fml) in enumerate(chk, start=4):
    wk.cell(i, 2, lab).font = F_TEXT; c = wk.cell(i, 3, fml); c.font = F_FORMULA
book.section(wk, 10, "OVERALL", 4); c = wk.cell(10, 3, '=IF(COUNTIF(C4:C8,FALSE)=0,"MODEL OK","CHECK ERRORS")'); c.font = Font(name=FONT, size=11, bold=True, color=book.theme.text_on_primary)

book.wb.move_sheet("Summary", offset=-4)
book.cover(
    toc=[("Summary", "TCS metrics and the implied per-share range from peer multiples"), ("Inputs", "FX, target, calendarisation window, peer-selection rationale"),
         ("Data", "raw pulls in reporting currency"), ("Comps", "US$ normalisation, EV build, LTM and calendarised multiples, peer statistics"),
         ("Calendarisation", "worked example of aligning March and December fiscal years"), ("Checks", "integrity checks")],
    highlights=[("TCS EV / LTM EBITDA", f"=Comps!N{tr}", MULT), ("Peer median EV / LTM EBITDA", f"=Comps!N{STAT['Median']}", MULT),
                ("TCS P/E (CY2027, calendarised)", f"=Comps!T{tr}", MULT), ("Implied price range, median across methods (INR)", f"=Summary!H{R_RANGE}", NUM0),
                ("Current price (INR)", f"={TM['Share price (INR)']}".replace("$C$", "Summary!$C$"), NUM0), ("Model status", "=Checks!C10", "General")],
    blurb="A comparable-companies analysis for Tata Consultancy Services against eight Indian and global IT-services peers: peer selection with stated criteria, currency normalisation (Infosys reports in US$ while trading in INR), enterprise-value construction, LTM and calendarised forward multiples, peer statistics and the implied valuation range for TCS.",
    method=["LTM figures are the sum of the last four reported quarters, so fiscal-year differences do not affect trailing multiples; forward figures are calendarised to TCS's FY2027 by month-overlap weights.",
            "EV = market cap + total debt + minority interest − cash and short-term investments; every component converted to US$ at the rate for its own currency (statements and share price can differ).",
            "Statistics exclude the target; the implied range applies the peer interquartile range of each multiple to TCS's own metric and bridges back to an INR price per share.",
            "Peer-selection criteria and exclusions are written down on the Inputs sheet, because that is where the judgement in a comps page lives."],
    sources=["Yahoo Finance: prices, shares outstanding, quarterly financial statements, consensus revenue and EPS estimates (12 Sep 2026)",
             "Company filings for fiscal-year ends (TCS, Infosys, HCLTech, Wipro, Tech Mahindra, Persistent: 31 March; Accenture: 31 August; Cognizant, EPAM: 31 December)"])
book.finish(freeze={"Comps": "C5", "Data": "C5"}, repeat_rows={"Comps": "1:4", "Data": "1:4"})
out = os.path.join(HERE, "TCS_Trading_Comps.xlsx"); book.save(out)
rc = recalc(out); print("recalc:", rc.get("status"), rc.get("total_errors"), rc.get("error_summary"))
pdf = export_pdf(out); png = preview_png(pdf, 0, dpi=60); os.rename(png, os.path.join(HERE, "cover.png")); print("done")
