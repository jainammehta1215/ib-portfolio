"""
Build 04_precedent_transactions/IT_Services_Precedents.xlsx — precedent M&A transactions in IT services, 2014–2024.

Deal terms are as disclosed in announcement press releases and filings (US$ millions, converted at the
announcement-date rate where the deal was struck in EUR or INR). Figures are rounded and marked
'approx.' where the source gave a range or where LTM metrics had to be taken from the last annual report.
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE, NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, recalc, export_pdf, preview_png)

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"
book = Book(theme=THEMES["tcs"], project_no=4, project="Precedent Transactions — IT Services", company="Tata Consultancy Services",
            units="US$ millions at announcement", as_of="deals announced Sep 2014 – Jun 2024")

# (announced, acquirer, acquirer type, target, target HQ, EV US$m, LTM revenue, LTM EBITDA (None if n/a), public target?, premium 1-day (None if n/a), consideration, source, note)
DEALS = [
    ("2014-09", "Cognizant", "Strategic", "TriZetto", "US", 2700, 711, 160, False, None, "Cash", "Cognizant press release, 15 Sep 2014", "Healthcare IT platform; EBITDA approx."),
    ("2015-04", "Capgemini", "Strategic", "IGATE", "US", 4040, 1269, 310, True, 0.047, "Cash (US$48/share)", "Capgemini press release, 27 Apr 2015; IGATE 10-K FY2014", "Premium to prior close; ~20% to 3-month average"),
    ("2016-03", "NTT Data", "Strategic", "Dell Services", "US", 3055, 2700, None, False, None, "Cash", "NTT Data / Dell press release, 28 Mar 2016", "Carve-out; revenue approx.; low margin legacy infrastructure"),
    ("2016-04", "Blackstone", "Sponsor", "Mphasis (60.5% stake from HP)", "India", 1400, 920, None, True, None, "Cash (INR 430/share) + open offer", "Blackstone press release, 4 Apr 2016; Mphasis FY16 annual report", "Implied 100% EV approx.; stake purchase with open offer"),
    ("2016-10", "Wipro", "Strategic", "Appirio", "US", 500, 200, None, False, None, "Cash", "Wipro press release, 20 Oct 2016", "Salesforce/Workday cloud services; revenue approx. (2016E)"),
    ("2018-07", "Atos", "Strategic", "Syntel", "US", 3400, 924, 250, True, 0.046, "Cash (US$41/share)", "Atos press release, 22 Jul 2018; Syntel 10-K FY2017", "Premium to prior close after a run-up; ~24% to 3-month VWAP. EBITDA approx. (27% margin)"),
    ("2019-06", "Capgemini", "Strategic", "Altran Technologies", "France", 5650, 3300, None, True, 0.22, "Cash (€14/share)", "Capgemini press release, 24 Jun 2019; Altran 2018 annual report", "EV ≈ €5.0bn incl. debt; converted at 1.13"),
    ("2021-03", "Wipro", "Strategic", "Capco", "UK", 1450, 720, None, False, None, "Cash", "Wipro press release, 4 Mar 2021", "Financial-services consulting; 2020 revenue approx."),
    ("2021-10", "Carlyle", "Sponsor", "Hexaware Technologies", "India", 3000, 850, None, False, None, "Cash", "Carlyle / Baring PE Asia press release, Oct 2021", "Secondary buyout from Baring; EV approx."),
    ("2022-05", "Wipro", "Strategic", "Rizing", "US", 540, 180, None, False, None, "Cash", "Wipro press release, 5 May 2022", "SAP services; revenue approx."),
    ("2022-05", "L&T Infotech", "Strategic", "Mindtree (merger)", "India", 6700, 1410, 290, True, 0.0, "Stock (73 LTI shares per 100 Mindtree)", "LTI / Mindtree exchange filings, 6 May 2022; Mindtree FY22 annual report", "Both under L&T control: not arm's-length; ratio ≈ market, premium ≈ 0"),
    ("2024-01", "Infosys", "Strategic", "in-tech", "Germany", 490, 185, None, False, None, "Cash (€450m)", "Infosys press release, 18 Jan 2024", "Engineering R&D services; converted at 1.09"),
    ("2024-05", "Coforge", "Strategic", "Cigniti Technologies (54% + open offer)", "India", 440, 220, 30, True, None, "Cash (INR 1,415/share)", "Coforge exchange filing, 2 May 2024; Cigniti FY24 annual report", "Implied 100% EV approx.; struck near market"),
    ("2024-06", "Cognizant", "Strategic", "Belcan", "US", 1290, 800, None, False, None, "Cash and stock", "Cognizant press release, 10 Jun 2024", "Engineering services; revenue approx. (2023)"),
]

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
wi = book.sheet("Inputs", ncols=6, label_width=44, subtitle="Screen definition and how to read the deal set")
wi.column_dimensions["F"].width = 72
book.section(wi, 4, "SCREEN", 6)
rows = [("Sector", "IT services, consulting and engineering-services providers with offshore or nearshore delivery"),
        ("Period", "September 2014 – June 2024 (ten years; long enough for a cycle, short enough that business models are comparable)"),
        ("Size", "Enterprise value US$400m – US$7bn at announcement"),
        ("Control", "Majority or 100% acquisitions and mergers; minority stakes excluded"),
        ("Data basis", "Deal value and target LTM metrics as disclosed at announcement; where the press release gave only equity value, EV adds net debt from the target's last balance sheet; EUR/INR deals converted at the announcement-date rate"),
        ("Caveat", "Several private-target revenue figures are approximate (press releases give rounded or 'expected' numbers). Multiples on those rows are indicative; the statistics use medians and quartiles, which are robust to a single imprecise entry. Verify against the cited source before using any figure in a client document.")]
for i, (k, v) in enumerate(rows, start=5):
    wi.cell(i, 2, k).font = F_BOLD; c = wi.cell(i, 3, v); c.font = F_TEXT; c.alignment = Alignment(wrap_text=True, vertical="top")
    wi.merge_cells(start_row=i, start_column=3, end_row=i, end_column=6); wi.row_dimensions[i].height = 30 if len(v) < 130 else 44
book.section(wi, 12, "TARGET FOR THE IMPLIED VALUATION (from Project 3)", 6)
tin = [("TCS LTM revenue (US$m)", 28627, NUM0, "Project 3, LTM to Jun 2026"), ("TCS LTM EBITDA (US$m)", 7748, NUM0, "Project 3"),
       ("TCS net debt (US$m, negative = net cash)", -3020, NUM0, "Project 3"), ("TCS shares (m)", 3618, NUM0, "Project 3"),
       ("USD per INR", 0.010467, '0.000000', "Project 3 (USDINR 95.54)"), ("TCS share price (INR)", 2200.8, NUM2, "12 Sep 2026"),
       ("Trading-comps peer median EV/Revenue (Project 3)", 1.82, MULT, "For the control-premium comparison"), ("Trading-comps peer median EV/EBITDA (Project 3)", 8.75, MULT, "")]
INP = {}
for i, (k, v, fmt, note) in enumerate(tin, start=13):
    wi.cell(i, 2, k).font = F_TEXT; c = wi.cell(i, 3, v); c.font = F_INPUT; c.number_format = fmt; wi.cell(i, 6, note).font = F_NOTE; INP[k] = f"Inputs!$C${i}"

# --------------------------------------------------------------------------- #
# Transactions
# --------------------------------------------------------------------------- #
wt = book.sheet("Transactions", ncols=17, label_width=12, col_width=13, subtitle="Deal database from primary announcements · blue = as disclosed · black = derived")
hdr = ["Announced", "Acquirer", "Acquirer type", "Target", "Target HQ", "EV (US$m)", "LTM revenue", "LTM EBITDA", "EBITDA margin", "EV / Revenue", "EV / EBITDA",
       "Public target", "Premium (1-day)", "Consideration", "Source", "Note"]
for j, hname in enumerate(hdr):
    c = wt.cell(4, 2 + j, hname); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
wt.row_dimensions[4].height = 30
wt.column_dimensions["E"].width = 30; wt.column_dimensions["P"].width = 44; wt.column_dimensions["Q"].width = 44; wt.column_dimensions["O"].width = 22
for i, d in enumerate(DEALS):
    r = 5 + i
    (ann, acq, atype, tgt, hq, ev, rev, ebitda, public, prem, cons, src, note) = d
    vals = [(ann, "General", F_INPUT), (acq, "General", F_INPUT), (atype, "General", F_INPUT), (tgt, "General", F_INPUT), (hq, "General", F_INPUT),
            (ev, NUM0, F_INPUT), (rev, NUM0, F_INPUT), (ebitda if ebitda is not None else "n/a", NUM0, F_INPUT),
            (f'=IF(ISNUMBER(I{r}),I{r}/H{r},"n/a")', PCT, F_FORMULA), (f"=G{r}/H{r}", MULT, F_FORMULA), (f'=IF(ISNUMBER(I{r}),G{r}/I{r},"n/a")', MULT, F_FORMULA),
            ("Yes" if public else "No", "General", F_INPUT), (prem if prem is not None else "n/a", PCT, F_INPUT), (cons, "General", F_INPUT), (src, "General", F_NOTE), (note, "General", F_NOTE)]
    for j, (v, fmt, fnt) in enumerate(vals):
        c = wt.cell(r, 2 + j, v); c.font = fnt; c.number_format = fmt
        if j in (14, 15): c.alignment = Alignment(wrap_text=True, vertical="top")
    wt.row_dimensions[r].height = 30
rN = 5 + len(DEALS) - 1
r = rN + 2
book.section(wt, r, "STATISTICS", 17); r += 1
STAT = {}
for lab, fn in [("Maximum", "MAX"), ("75th percentile", "PERCENTILE"), ("Median", "MEDIAN"), ("Mean", "AVERAGE"), ("25th percentile", "PERCENTILE"), ("Minimum", "MIN"), ("Count", "COUNT")]:
    wt.cell(r, 2, lab).font = F_BOLD; STAT[lab] = r
    for col in (10, 11, 12, 14):    # margin, EV/Rev, EV/EBITDA, premium
        rng = f"{L(col)}5:{L(col)}{rN}"
        if fn == "PERCENTILE":
            q = 0.75 if "75" in lab else 0.25; f = f"=PERCENTILE({rng},{q})"
        else:
            f = f"={fn}({rng})"
        c = wt.cell(r, col, f); c.font = F_FORMULA; c.number_format = "General" if fn == "COUNT" else (PCT if col in (10, 14) else MULT)
    r += 1
book.note(wt, r, "Text entries ('n/a') are ignored by the statistical functions, so EV/EBITDA and premium statistics are computed only over the deals where the figure was disclosed. Count shows how many observations each statistic rests on.", 17, height=28)
r += 2
book.section(wt, r, "BY ACQUIRER TYPE AND TARGET GEOGRAPHY (median EV / Revenue)", 17); r += 1
for i, (lab, col, val) in enumerate([("Strategic acquirers", "D", "Strategic"), ("Financial sponsors", "D", "Sponsor"), ("India-headquartered targets", "F", "India"), ("US / Europe targets", "F", None)]):
    wt.cell(r + i, 2, lab).font = F_TEXT
    if val:
        f = f'=MEDIAN(IF({col}$5:{col}${rN}="{val}",K$5:K${rN}))'
    else:
        f = f'=MEDIAN(IF({col}$5:{col}${rN}<>"India",K$5:K${rN}))'
    c = wt.cell(r + i, 11, f); c.font = F_FORMULA; c.number_format = MULT
    cnt = f'=COUNTIF({col}5:{col}{rN},"{val}")' if val else f'=COUNTIF({col}5:{col}{rN},"<>India")'
    c = wt.cell(r + i, 12, cnt); c.font = F_FORMULA; c.number_format = "General"
    wt.cell(r + i, 13, "deals").font = F_NOTE
R_SEG = r
# array formulas for MEDIAN(IF(...)) — mark as array
from openpyxl.worksheet.formula import ArrayFormula
for i in range(4):
    cell = wt.cell(R_SEG + i, 11); cell.value = ArrayFormula(f"K{R_SEG+i}", cell.value)

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
ws = book.sheet("Summary", ncols=8, label_width=46, col_width=13, subtitle="Precedent multiples, control premium versus trading comps, and the implied TCS range")
book.section(ws, 4, "PRECEDENT MULTIPLES", 8)
for j, hname in enumerate(["", "25th pct", "Median", "75th pct", "Count"]):
    c = ws.cell(5, 2 + j, hname); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center")
for i, (lab, col, fmt) in enumerate([("EV / LTM revenue", 11, MULT), ("EV / LTM EBITDA (disclosed deals)", 12, MULT), ("EBITDA margin of targets", 10, PCT), ("1-day premium (public targets)", 14, PCT)], start=6):
    ws.cell(i, 2, lab).font = F_TEXT
    for j, stat in enumerate(["25th percentile", "Median", "75th percentile", "Count"]):
        c = ws.cell(i, 3 + j, f"=Transactions!{L(col)}{STAT[stat]}"); c.font = F_LINK; c.number_format = "General" if stat == "Count" else fmt
book.section(ws, 11, "CONTROL PREMIUM: PRECEDENTS vs TRADING COMPARABLES", 8)
rows = [("Precedent median EV / Revenue", "=C13*0+Transactions!K{m}".format(m=STAT["Median"]), MULT), ("Trading-comps peer median EV / Revenue (Project 3)", f"={INP['Trading-comps peer median EV/Revenue (Project 3)']}", MULT),
        ("Implied control premium on revenue multiple", "=C12/C13-1", PCT), ("Precedent median EV / EBITDA", f"=Transactions!L{STAT['Median']}", MULT),
        ("Trading-comps peer median EV / EBITDA (Project 3)", f"={INP['Trading-comps peer median EV/EBITDA (Project 3)']}", MULT), ("Implied control premium on EBITDA multiple", "=C15/C16-1", PCT)]
for i, (lab, f, fmt) in enumerate(rows, start=12):
    ws.cell(i, 2, lab).font = F_BOLD if "premium" in lab else F_TEXT; c = ws.cell(i, 3, f); c.font = F_FORMULA if "Project" not in lab else F_LINK; c.number_format = fmt
ws["C12"] = f"=Transactions!K{STAT['Median']}"; ws["C12"].font = F_LINK
book.section(ws, 19, "IMPLIED TCS VALUE PER SHARE (INR) — illustrative", 8)
for j, hname in enumerate(["Multiple", "TCS metric (US$m)", "25th pct", "Median", "75th pct", "Implied @25th", "Implied @median", "Implied @75th"]):
    c = ws.cell(20, 2 + j, hname); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
ws.row_dimensions[20].height = 30
def impl(metric, col):
    return [f"=({metric}*Transactions!{L(col)}{STAT[s]}-{INP['TCS net debt (US$m, negative = net cash)']})/{INP['TCS shares (m)']}/{INP['USD per INR']}" for s in ("25th percentile", "Median", "75th percentile")]
for i, (lab, metric, col) in enumerate([("EV / LTM revenue", INP["TCS LTM revenue (US$m)"], 11), ("EV / LTM EBITDA", INP["TCS LTM EBITDA (US$m)"], 12)], start=21):
    ws.cell(i, 2, lab).font = F_TEXT; c = ws.cell(i, 3, f"={metric}"); c.font = F_LINK; c.number_format = NUM0
    for j, s_ in enumerate(("25th percentile", "Median", "75th percentile")):
        c = ws.cell(i, 4 + j, f"=Transactions!{L(col)}{STAT[s_]}"); c.font = F_LINK; c.number_format = MULT
    for j, f in enumerate(impl(metric, col)):
        c = ws.cell(i, 7 + j, f); c.font = F_FORMULA; c.number_format = NUM0
ws.cell(23, 2, "Current TCS share price (INR)").font = F_TEXT; c = ws.cell(23, 8, f"={INP['TCS share price (INR)']}"); c.font = F_LINK; c.number_format = NUM0
book.note(ws, 25, "Illustrative only: every precedent target is a fraction of TCS's size and none is a like-for-like control transaction for a US$80bn company. The range is shown to complete the football field (Project 5) and to quantify the sector's control premium; it is not a takeover valuation of TCS.", 8, height=42)
book.section(ws, 27, "READING THE RESULT", 8)
memo = ["Precedent multiples sit above trading multiples for the same kind of business: buyers pay a control premium for synergies and for the right to run the asset. The gap between the two medians is the sector's typical premium; the public-target 1-day premia are low here because several targets (IGATE, Syntel) had run up on deal speculation, and one (LTI–Mindtree) was a merger under common control.",
        "Strategic acquirers dominate this sector and paid more than sponsors, consistent with synergies being real in offshore IT services (client cross-sell, pyramid cost savings).",
        "Multiples cluster around 1.5–3.5× revenue for mid-cap services assets; the outliers (TriZetto, Mindtree) had platform or scarcity value. A banker would show this dispersion to a seller before quoting a price.",
        "The data quality caveat on the Inputs sheet is part of the deliverable: precedent analysis from public sources always carries approximation in private-target metrics, and saying so is more credible than false precision."]
for i, m in enumerate(memo):
    c = ws.cell(28 + i, 2, "•  " + m); c.font = F_TEXT; c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=28 + i, start_column=2, end_row=28 + i, end_column=8); ws.row_dimensions[28 + i].height = 44

wk = book.sheet("Checks", ncols=4, label_width=60, subtitle="Integrity checks · should all be TRUE")
chk = [("All EV / Revenue multiples positive", f"=MIN(Transactions!K5:K{rN})>0"), ("EV / Revenue median within 1x–6x", f"=AND(Transactions!K{STAT['Median']}>1,Transactions!K{STAT['Median']}<6)"),
       ("EV / EBITDA statistics computed on ≥ 4 disclosed deals", f"=Transactions!L{STAT['Count']}>=4"), ("Control premium on revenue multiple is positive", "=Summary!C14>0"),
       ("Implied TCS range is ordered", "=AND(Summary!G21<=Summary!H21,Summary!H21<=Summary!I21)")]
for i, (lab, f) in enumerate(chk, start=4):
    wk.cell(i, 2, lab).font = F_TEXT; c = wk.cell(i, 3, f); c.font = F_FORMULA
book.section(wk, 10, "OVERALL", 4); c = wk.cell(10, 3, '=IF(COUNTIF(C4:C8,FALSE)=0,"MODEL OK","CHECK ERRORS")'); c.font = Font(name=FONT, size=11, bold=True, color=book.theme.text_on_primary)

book.wb.move_sheet("Summary", offset=-3)
book.cover(
    toc=[("Summary", "precedent multiples, control premium vs trading comps, illustrative TCS range"), ("Inputs", "screen definition, data caveats, TCS metrics from Project 3"),
         ("Transactions", "fourteen deals 2014–2024 with sources, multiples, premia and statistics"), ("Checks", "integrity checks")],
    highlights=[("Precedent median EV / Revenue", f"=Transactions!K{STAT['Median']}", MULT), ("Precedent median EV / EBITDA", f"=Transactions!L{STAT['Median']}", MULT),
                ("Control premium vs trading comps (revenue)", "=Summary!C14", PCT), ("Deals in the set", f"=Transactions!K{STAT['Count']}", "General"), ("Model status", "=Checks!C10", "General")],
    blurb="Fourteen IT-services acquisitions from 2014 to 2024 (Capgemini–IGATE, Atos–Syntel, Capgemini–Altran, Carlyle–Hexaware, LTI–Mindtree, Cognizant–Belcan and others), built from announcement press releases and filings: enterprise value, target LTM revenue and EBITDA, implied multiples, one-day premia for public targets, consideration and source for every row. Statistics by acquirer type and geography, the control premium relative to Project 3's trading multiples, and an illustrative range for TCS.",
    method=["Each row cites its announcement source; EV adds target net debt where only equity value was disclosed; EUR and INR deals converted at the announcement-date rate.",
            "Private-target metrics are approximate where the source gave rounded figures; medians and quartiles are used because they are robust to a single imprecise entry, and the count of observations is shown for every statistic.",
            "Premia are one-day to the unaffected close; two targets had run up on speculation and one merger was under common control, all noted in the table.",
            "The implied TCS range is labelled illustrative: no precedent target is comparable in size to TCS."],
    sources=["Acquirer press releases and exchange filings at announcement (cited per row)", "Target annual reports / 10-Ks for LTM revenue and EBITDA", "Project 3 for TCS metrics and trading-comps medians"])
book.finish(freeze={"Transactions": "C5"}, repeat_rows={"Transactions": "1:4"})
out = os.path.join(HERE, "IT_Services_Precedents.xlsx"); book.save(out)
rc = recalc(out); print("recalc:", rc.get("status"), rc.get("total_errors"), rc.get("error_summary"))
pdf = export_pdf(out); png = preview_png(pdf, 0, dpi=60); os.rename(png, os.path.join(HERE, "cover.png")); print("done")
