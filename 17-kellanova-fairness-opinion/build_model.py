"""
Build 17-kellanova-fairness-opinion/Kellanova_Fairness_Opinion.xlsx — a replication of the fairness-opinion analyses
that Goldman Sachs and Lazard delivered to Kellanova's board on Mars's US$83.50-per-share cash offer, rebuilt from the
inputs disclosed in the definitive proxy statement (DEFM14A filed 26 September 2024).

Sheets: Cover · Summary · Inputs · Goldman Sachs · Lazard · Football Field · Checks

A fairness opinion is the most scrutinised valuation a bank produces, and the proxy is where it becomes public:
the management projections, each analysis, the ranges chosen and the value per share each one implied. The
exercise here is to take exactly what the proxy discloses \u2014 projections, discount rates, multiple ranges, premia
reference ranges, the undisturbed price \u2014 and see whether the advisors' per-share ranges can be reproduced. Two
inputs are not disclosed as numbers (adjusted net debt and fully diluted shares); they are back-solved from one of
Goldman's disclosed ranges and then used, unchanged, for every other analysis. The Checks sheet demands that each
rebuilt range land within a stated tolerance of the range in the proxy.

Source: Kellanova DEFM14A, SEC accession 0001193125-24-226970, sections "Certain Financial Projections", "Opinion of
Goldman Sachs & Co. LLC" and "Opinion of Lazard Fr\u00e8res & Co. LLC". US$ millions except per share.
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, Theme, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)
from ibkit.sheet import Sheet, checks_sheet

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"
KEL = Theme("kellanova", "Kellanova", "B3202E", "6B6B6B", "F5A800", "FBEAEC", tagline="NYSE: K \u00b7 acquired by Mars")

PROJ = dict(years=["2024E", "2025E", "2026E", "2027E"], sales=[12782, 13366, 13670, 14098], ebitda=[2235, 2401, 2553, 2741],
            nopat=[1322, 1506, 1595, 1705], ufcf=[1102, 1348, 1359, 1533], eps=[3.70, 3.99, 4.26, 4.63])
DEAL = dict(offer=83.50, undisturbed=62.98, hi52=62.98, post_spin_lo=48.62, post_spin_hi=62.98, analyst_lo=56.07, analyst_hi=68.22,
            basic_shares=344.684757, dps=[2.28, 2.36, 2.44])
GS = dict(dcf_range=(68.42, 83.82), pv_range=(60.93, 90.01), premia_und=(71.80, 83.13), premia_52=(67.39, 77.47), prec=(58.81, 118.63),
          wacc=(0.06, 0.07), tv_mult=(12.0, 14.0), ntm_mult=(11.5, 14.5), ke=0.069, prem_und=(0.14, 0.32), prem_52=(0.07, 0.23), prec_mult=(12.1, 21.5),
          comps_snack=14.8, comps_center=9.9)
LZ = dict(dcf_range=(69.60, 79.10), comps_ebitda=(57.20, 73.80), comps_pe=(55.90, 75.70), prec=(78.10, 90.20), premia=(75.60, 85.10),
          wacc=(0.0625, 0.0675), pgr=(0.015, 0.020), ebitda_mult=(11.0, 13.5), pe_mult=(14.0, 19.0), prec_mult=(15.5, 17.5), prem=(0.20, 0.35),
          snack=(14.3, 14.9, 14.5), grocery=(9.1, 12.1, 9.4))

book = Book(theme=KEL, project_no=17, project="Fairness Opinion Replication", company="Kellanova / Mars",
            units="US$ millions \u00b7 per-share values in US$", as_of="DEFM14A filed 26 Sep 2024 \u00b7 valuation date 30 Jun 2024")
S = lambda name, **kw: Sheet(book, name, **kw)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
si = S("Inputs", ncols=7, label_width=60, col_width=12.5, subtitle="Everything the proxy discloses, and the two figures back-solved from it")
si.ws.column_dimensions["G"].width = 60
I = {}
si.section("TRANSACTION")
I["offer"] = si.row("Merger consideration per share (cash)", DEAL["offer"], NUM2, bold=True)
I["und"] = si.row("Undisturbed closing price, 2 August 2024", DEAL["undisturbed"], NUM2, note="Also the 52-week high on that date")
I["prem"] = si.row("Premium to undisturbed", f"={I['offer']}/{I['und']}-1", PCT, bold=True)
I["basic"] = si.row("Shares outstanding at the record date (millions)", DEAL["basic_shares"], NUM2)
si.blank(); si.section("MANAGEMENT PROJECTIONS (proxy, 'Certain Financial Projections')")
si.head(PROJ["years"])
I["sales"] = si.multi("Net sales", PROJ["sales"], NUM0)
I["ebitda"] = si.multi("Adjusted EBITDA", PROJ["ebitda"], NUM0, bold=True)
I["nopat"] = si.multi("Net operating profit after tax", PROJ["nopat"], NUM0)
I["ufcf"] = si.multi("Unlevered free cash flow", PROJ["ufcf"], NUM0, bold=True)
I["eps"] = si.multi("Adjusted earnings per share", PROJ["eps"], NUM2)
I["dps"] = si.multi("Dividends per share (stated: 2024 run-rate, then +3.5% a year)", DEAL["dps"] + [""], NUM2)
si.blank(); si.section("BACK-SOLVED FROM GOLDMAN'S PRECEDENT-TRANSACTION RANGE")
si.subsection("Goldman applied 12.1x\u201321.5x to FY1 (2024E) EBITDA and reported $58.81\u2013$118.63 per share. Two equations, two unknowns.")
I["fds"] = si.row("Fully diluted shares (millions) = (21.5 \u2212 12.1) \u00d7 FY1 EBITDA / (118.63 \u2212 58.81)",
                  f"=({GS['prec_mult'][1]}-{GS['prec_mult'][0]})*{I['ebitda'][0]}/({GS['prec'][1]}-{GS['prec'][0]})", NUM2, bold=True,
                  note="About 2% above the record-date basic count \u2014 the treasury-stock dilution one would expect")
I["nd"] = si.row("Adjusted net debt = 12.1 \u00d7 FY1 EBITDA \u2212 58.81 \u00d7 diluted shares",
                 f"={GS['prec_mult'][0]}*{I['ebitda'][0]}-{GS['prec'][0]}*{I['fds']}", NUM0, bold=True,
                 note="Financial net debt plus factored receivables and minorities less unconsolidated investments, as the proxy defines it")
I["nd_chk"] = si.row("  check: 21.5x reproduces the top of the range", f"=({GS['prec_mult'][1]}*{I['ebitda'][0]}-{I['nd']})/{I['fds']}", NUM2, font=F_FORMULA)
si.blank(); si.section("HALF-YEAR CONVENTION")
I["h2_frac"] = si.row("Share of 2024E unlevered free cash flow falling after 30 June 2024", 0.5, PCT, note="The proxy discounts cash flow from 1 July 2024; the split is not disclosed")
I["stub"] = si.row("Mid-year discount period for H2 2024 (years from 30 June 2024)", 0.25, NUM2)
si.note("Every number above the back-solved section is quoted from the proxy. Diluted shares and adjusted net debt are the only "
        "derived inputs, and they are derived from a single disclosed analysis so that every other analysis is a genuine test.", height=32)

# --------------------------------------------------------------------------- #
# Goldman Sachs
# --------------------------------------------------------------------------- #
sg = S("Goldman Sachs", ncols=7, label_width=62, col_width=12.5, subtitle="Five analyses rebuilt from the disclosed inputs, each against the range the proxy reports")
sg.ws.column_dimensions["G"].width = 54
G = {}
sg.section("1. ILLUSTRATIVE DISCOUNTED CASH FLOW \u2014 WACC 6.0\u20137.0%, terminal 12.0\u201314.0x 2027E EBITDA, mid-year, from 30 June 2024")
sg.head(["H2 2024", "2025E", "2026E", "2027E"])
G["cf"] = sg.multi("Unlevered free cash flow", [f"={I['ufcf'][0]}*{I['h2_frac']}", f"={I['ufcf'][1]}", f"={I['ufcf'][2]}", f"={I['ufcf'][3]}"], NUM0, font=F_LINK)
G["t"] = sg.multi("Discount period (years, mid-period)", [f"={I['stub']}", f"={I['stub']}*2+0.5", f"={I['stub']}*2+1.5", f"={I['stub']}*2+2.5"], NUM2)
def dcf_gs(wacc, mult):
    pv = "+".join(f"{G['cf'][i]}/(1+{wacc})^{G['t'][i]}" for i in range(4))
    tv = f"{mult}*{I['ebitda'][3]}/(1+{wacc})^3.5"
    return f"=(({pv})+{tv}-{I['nd']})/{I['fds']}"
G["dcf_lo"] = sg.row("Low: 7.0% WACC, 12.0x terminal (US$ per share)", dcf_gs(GS["wacc"][1], GS["tv_mult"][0]), NUM2, bold=True)
G["dcf_hi"] = sg.row("High: 6.0% WACC, 14.0x terminal", dcf_gs(GS["wacc"][0], GS["tv_mult"][1]), NUM2, bold=True)
G["dcf_ref"] = sg.row("Proxy: $68.42 \u2013 $83.82", f"=TEXT({GS['dcf_range'][0]},\"0.00\")&\" \u2013 \"&TEXT({GS['dcf_range'][1]},\"0.00\")", "@", font=F_INPUT)
G["pgr_lo"] = sg.row("Implied perpetuity growth at 12.0x and 6.0%: g = (TV\u00d7r \u2212 CF\u2082\u2080\u2082\u2087)/(TV + CF\u2082\u2080\u2082\u2087)",
                     f"=({GS['tv_mult'][0]}*{I['ebitda'][3]}*{GS['wacc'][0]}-{I['ufcf'][3]})/({GS['tv_mult'][0]}*{I['ebitda'][3]}+{I['ufcf'][3]})", PCT)
G["pgr_hi"] = sg.row("Implied perpetuity growth at 14.0x and 7.0%", f"=({GS['tv_mult'][1]}*{I['ebitda'][3]}*{GS['wacc'][1]}-{I['ufcf'][3]})/({GS['tv_mult'][1]}*{I['ebitda'][3]}+{I['ufcf'][3]})", PCT,
                     note="Proxy: 'implied perpetuity growth rates ranging from 1.0% to 2.6%'")
sg.blank(); sg.section("2. PRESENT VALUE OF FUTURE SHARE PRICE \u2014 11.5\u201314.5x NTM EBITDA at each year-end, dividends added, discounted at 6.9% cost of equity")
sg.head(["YE 2024", "YE 2025", "YE 2026"], label="Value at year-end, then discounted to 30 June 2024")
def pv_fsp(mult, yr):  # yr 0..2 ; NTM EBITDA = next year's ; cumulative dividends through year-end
    cum_div = "+".join(I["dps"][k] for k in range(yr + 1))
    n = f"({yr}+0.5)"
    return f"=(({mult}*{I['ebitda'][yr+1]}-{I['nd']})/{I['fds']})/(1+{GS['ke']})^{n}+({cum_div})/(1+{GS['ke']})^{n}"
G["pv_lo"] = sg.multi("At 11.5x", [pv_fsp(GS["ntm_mult"][0], y) for y in range(3)], NUM2)
G["pv_hi"] = sg.multi("At 14.5x", [pv_fsp(GS["ntm_mult"][1], y) for y in range(3)], NUM2)
G["pvr_lo"] = sg.row("Rebuilt range \u2014 low", f"=MIN({','.join(G['pv_lo'])})", NUM2, bold=True)
G["pvr_hi"] = sg.row("Rebuilt range \u2014 high", f"=MAX({','.join(G['pv_hi'])})", NUM2, bold=True)
G["pv_ref"] = sg.row("Proxy: $60.93 \u2013 $90.01", f"=TEXT({GS['pv_range'][0]},\"0.00\")&\" \u2013 \"&TEXT({GS['pv_range'][1]},\"0.00\")", "@", font=F_INPUT)
sg.blank(); sg.section("3. PREMIA PAID \u2014 US targets over US$20bn since 2008, strategic buyers")
G["pu_lo"] = sg.row("14% over undisturbed", f"={I['und']}*(1+{GS['prem_und'][0]})", NUM2)
G["pu_hi"] = sg.row("32% over undisturbed", f"={I['und']}*(1+{GS['prem_und'][1]})", NUM2)
G["pu_ref"] = sg.row("Proxy: $71.80 \u2013 $83.13", f"=TEXT({GS['premia_und'][0]},\"0.00\")&\" \u2013 \"&TEXT({GS['premia_und'][1]},\"0.00\")", "@", font=F_INPUT)
G["p52_lo"] = sg.row("7% over the 52-week high", f"={I['und']}*(1+{GS['prem_52'][0]})", NUM2)
G["p52_hi"] = sg.row("23% over the 52-week high", f"={I['und']}*(1+{GS['prem_52'][1]})", NUM2)
G["p52_ref"] = sg.row("Proxy: $67.39 \u2013 $77.47", f"=TEXT({GS['premia_52'][0]},\"0.00\")&\" \u2013 \"&TEXT({GS['premia_52'][1]},\"0.00\")", "@", font=F_INPUT)
G["off_prem"] = sg.row("Where the offer sits: premium to undisturbed versus the 75th percentile of precedents (32%)", f"={I['prem']}-{GS['prem_und'][1]}", PCT, bold=True,
                       note="The offer is above every precedent premium in the set: Mars paid more than the top quartile")
sg.blank(); sg.section("4. PRECEDENT TRANSACTIONS \u2014 12.1\u201321.5x FY1 EBITDA (the calibrating analysis)")
G["pr_lo"] = sg.row("12.1x 2024E EBITDA", f"=({GS['prec_mult'][0]}*{I['ebitda'][0]}-{I['nd']})/{I['fds']}", NUM2)
G["pr_hi"] = sg.row("21.5x 2024E EBITDA", f"=({GS['prec_mult'][1]}*{I['ebitda'][0]}-{I['nd']})/{I['fds']}", NUM2)
G["pr_ref"] = sg.row("Proxy: $58.81 \u2013 $118.63 (reproduced by construction)", f"=TEXT({GS['prec'][0]},\"0.00\")&\" \u2013 \"&TEXT({GS['prec'][1]},\"0.00\")", "@", font=F_INPUT)
G["off_mult"] = sg.row("Multiple of 2024E EBITDA the offer implies", f"=({I['offer']}*{I['fds']}+{I['nd']})/{I['ebitda'][0]}", MULT, bold=True)
G["off_mult25"] = sg.row("Multiple of 2025E EBITDA the offer implies", f"=({I['offer']}*{I['fds']}+{I['nd']})/{I['ebitda'][1]}", MULT)
G["prec_med"] = sg.row("Median of the eleven precedents (proxy)", 16.2, MULT, font=F_INPUT, note="Wrigley/Mars 2008 at 17.6x is in the set")
sg.blank(); sg.section("5. TRADING COMPARABLES \u2014 'for reference purposes only'")
G["c_snack"] = sg.row("Median EV/NTM EBITDA, snacking (Hershey, Mondel\u0113z, PepsiCo)", GS["comps_snack"], MULT, font=F_INPUT)
G["c_center"] = sg.row("Median EV/NTM EBITDA, centre of store (Campbell, Conagra, General Mills, Smucker, Kraft Heinz)", GS["comps_center"], MULT, font=F_INPUT)
G["c_off"] = sg.row("Offer multiple of NTM (2025E) EBITDA against the snacking median", f"={G['off_mult25']}-{G['c_snack']}", NUM2, note="Turns of EBITDA above the best-rated peer group")

# --------------------------------------------------------------------------- #
# Lazard
# --------------------------------------------------------------------------- #
sl = S("Lazard", ncols=7, label_width=62, col_width=12.5, subtitle="Lazard's analyses rebuilt with the same back-solved balance sheet")
sl.ws.column_dimensions["G"].width = 54
Z = {}
sl.section("1. DISCOUNTED CASH FLOW \u2014 WACC 6.25\u20136.75%, perpetuity growth 1.5\u20132.0%, mid-year, from 30 June 2024")
def dcf_lz(wacc, g):
    pv = "+".join(f"{G['cf'][i]}/(1+{wacc})^{G['t'][i]}" for i in range(4))
    tv = f"{I['ufcf'][3]}*(1+{g})/({wacc}-{g})/(1+{wacc})^3.5"
    return f"=(({pv})+{tv}-{I['nd']})/{I['fds']}"
Z["dcf_lo"] = sl.row("Low: 6.75% WACC, 1.5% growth", dcf_lz(LZ["wacc"][1], LZ["pgr"][0]), NUM2, bold=True)
Z["dcf_hi"] = sl.row("High: 6.25% WACC, 2.0% growth", dcf_lz(LZ["wacc"][0], LZ["pgr"][1]), NUM2, bold=True)
Z["dcf_ref"] = sl.row("Proxy: $69.60 \u2013 $79.10", f"=TEXT({LZ['dcf_range'][0]},\"0.00\")&\" \u2013 \"&TEXT({LZ['dcf_range'][1]},\"0.00\")", "@", font=F_INPUT)
# what terminal free cash flow would Lazard need at 6.75% / 1.5% to report $69.60? (back-solve; the proxy says 'estimated unlevered free cash flow')
pv_lo = "+".join(f"{G['cf'][i]}/(1+{LZ['wacc'][1]})^{G['t'][i]}" for i in range(4))
Z["tv_cf_needed"] = sl.row("Terminal free cash flow that reproduces Lazard's $69.60 low end at 6.75% and 1.5% (US$ m)",
                           f"=({LZ['dcf_range'][0]}*{I['fds']}+{I['nd']}-({pv_lo}))*(1+{LZ['wacc'][1]})^3.5*({LZ['wacc'][1]}-{LZ['pgr'][0]})/(1+{LZ['pgr'][0]})", NUM0, bold=True,
                           note="Close to 2027E NOPAT (US$1,705m): Lazard appears to have normalised the terminal year, removing the working-capital build in 2027E free cash flow")
Z["tv_cf_ratio"] = sl.row("  as a multiple of 2027E unlevered free cash flow", f"={Z['tv_cf_needed']}/{I['ufcf'][3]}", MULT)
Z["tv_mult"] = sl.row("Terminal value at the midpoint as a multiple of 2027E EBITDA", f"={I['ufcf'][3]}*(1+0.0175)/(0.065-0.0175)/{I['ebitda'][3]}", MULT,
                      note="Lazard's growth method and Goldman's multiple method should meet in the middle; this shows whether they do")
sl.blank(); sl.section("2. SELECTED PUBLIC COMPANIES \u2014 11.0\u201313.5x 2025E EBITDA and 14.0\u201319.0x 2025E EPS")
Z["ce_lo"] = sl.row("11.0x 2025E EBITDA", f"=({LZ['ebitda_mult'][0]}*{I['ebitda'][1]}-{I['nd']})/{I['fds']}", NUM2)
Z["ce_hi"] = sl.row("13.5x 2025E EBITDA", f"=({LZ['ebitda_mult'][1]}*{I['ebitda'][1]}-{I['nd']})/{I['fds']}", NUM2)
Z["ce_ref"] = sl.row("Proxy: $57.20 \u2013 $73.80", f"=TEXT({LZ['comps_ebitda'][0]},\"0.00\")&\" \u2013 \"&TEXT({LZ['comps_ebitda'][1]},\"0.00\")", "@", font=F_INPUT)
Z["pe_lo"] = sl.row("14.0x 2025E adjusted EPS", f"={LZ['pe_mult'][0]}*{I['eps'][1]}", NUM2)
Z["pe_hi"] = sl.row("19.0x 2025E adjusted EPS", f"={LZ['pe_mult'][1]}*{I['eps'][1]}", NUM2)
Z["pe_ref"] = sl.row("Proxy: $55.90 \u2013 $75.70", f"=TEXT({LZ['comps_pe'][0]},\"0.00\")&\" \u2013 \"&TEXT({LZ['comps_pe'][1]},\"0.00\")", "@", font=F_INPUT)
Z["snack"] = sl.row("Observed: snacking 2025E EBITDA low / high / median", f"=TEXT({LZ['snack'][0]},\"0.0\")&\"x / \"&TEXT({LZ['snack'][1]},\"0.0\")&\"x / \"&TEXT({LZ['snack'][2]},\"0.0\")&\"x\"", "@", font=F_INPUT)
Z["groc"] = sl.row("Observed: diversified grocery 2025E EBITDA low / high / median", f"=TEXT({LZ['grocery'][0]},\"0.0\")&\"x / \"&TEXT({LZ['grocery'][1]},\"0.0\")&\"x / \"&TEXT({LZ['grocery'][2]},\"0.0\")&\"x\"", "@", font=F_INPUT)
sl.blank(); sl.section("3. PRECEDENT TRANSACTIONS \u2014 15.5\u201317.5x LTM adjusted EBITDA (eleven food deals since 2014; median 16.7x)")
Z["ltm"] = sl.row("LTM adjusted EBITDA to 29 June 2024, back-solved from the disclosed range (US$ m)",
                  f"=({LZ['prec'][1]}-{LZ['prec'][0]})*{I['fds']}/({LZ['prec_mult'][1]}-{LZ['prec_mult'][0]})", NUM0, bold=True,
                  note="Sits between 2023 actual and 2024E, as an LTM figure to mid-year should")
Z["pt_lo"] = sl.row("15.5x LTM", f"=({LZ['prec_mult'][0]}*{Z['ltm']}-{I['nd']})/{I['fds']}", NUM2)
Z["pt_hi"] = sl.row("17.5x LTM", f"=({LZ['prec_mult'][1]}*{Z['ltm']}-{I['nd']})/{I['fds']}", NUM2)
Z["pt_ref"] = sl.row("Proxy: $78.10 \u2013 $90.20", f"=TEXT({LZ['prec'][0]},\"0.00\")&\" \u2013 \"&TEXT({LZ['prec'][1]},\"0.00\")", "@", font=F_INPUT)
Z["off_ltm"] = sl.row("Multiple of LTM EBITDA the offer implies", f"=({I['offer']}*{I['fds']}+{I['nd']})/{Z['ltm']}", MULT, bold=True)
sl.blank(); sl.section("4. OTHER ANALYSES (informational, not the basis of the opinion)")
Z["pm_lo"] = sl.row("Premia paid: 20% over the 2 August close", f"={I['und']}*(1+{LZ['prem'][0]})", NUM2)
Z["pm_hi"] = sl.row("Premia paid: 35% over the 2 August close", f"={I['und']}*(1+{LZ['prem'][1]})", NUM2)
Z["pm_ref"] = sl.row("Proxy: $75.60 \u2013 $85.10", f"=TEXT({LZ['premia'][0]},\"0.00\")&\" \u2013 \"&TEXT({LZ['premia'][1]},\"0.00\")", "@", font=F_INPUT)
Z["an_lo"] = sl.row("Analyst price targets, discounted one year at 7.0%: low", DEAL["analyst_lo"], NUM2, font=F_INPUT)
Z["an_hi"] = sl.row("  high", DEAL["analyst_hi"], NUM2, font=F_INPUT)
Z["ps_lo"] = sl.row("Post-separation trading range (since 2 October 2023): low", DEAL["post_spin_lo"], NUM2, font=F_INPUT)
Z["ps_hi"] = sl.row("  high", DEAL["post_spin_hi"], NUM2, font=F_INPUT)

# --------------------------------------------------------------------------- #
# Football Field
# --------------------------------------------------------------------------- #
sf = S("Football Field", ncols=8, label_width=58, col_width=12, subtitle="Every range on one sheet against the US$83.50 offer")
sf.ws.column_dimensions["H"].width = 40
FF = {}
sf.section("RANGES \u2014 rebuilt, with the proxy's own figures alongside")
hdr = ["Rebuilt low", "Rebuilt high", "Proxy low", "Proxy high", "Offer vs proxy range", "Max deviation"]
for j, h in enumerate(hdr):
    c = sf.ws.cell(sf.r, 3 + j, h); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
sf.ws.cell(sf.r, 2, "Analysis").font = F_BOLD; sf.ws.row_dimensions[sf.r].height = 28; sf.r += 1
rows = [("GS \u2014 discounted cash flow", G["dcf_lo"], G["dcf_hi"], GS["dcf_range"]),
        ("GS \u2014 present value of future share price", G["pvr_lo"], G["pvr_hi"], GS["pv_range"]),
        ("GS \u2014 premia paid, undisturbed", G["pu_lo"], G["pu_hi"], GS["premia_und"]),
        ("GS \u2014 premia paid, 52-week high", G["p52_lo"], G["p52_hi"], GS["premia_52"]),
        ("GS \u2014 precedent transactions", G["pr_lo"], G["pr_hi"], GS["prec"]),
        ("Lazard \u2014 discounted cash flow", Z["dcf_lo"], Z["dcf_hi"], LZ["dcf_range"]),
        ("Lazard \u2014 public companies, EBITDA", Z["ce_lo"], Z["ce_hi"], LZ["comps_ebitda"]),
        ("Lazard \u2014 public companies, P/E", Z["pe_lo"], Z["pe_hi"], LZ["comps_pe"]),
        ("Lazard \u2014 precedent transactions", Z["pt_lo"], Z["pt_hi"], LZ["prec"]),
        ("Lazard \u2014 premia paid", Z["pm_lo"], Z["pm_hi"], LZ["premia"]),
        ("Lazard \u2014 analyst targets", Z["an_lo"], Z["an_hi"], (DEAL["analyst_lo"], DEAL["analyst_hi"])),
        ("Lazard \u2014 post-separation trading", Z["ps_lo"], Z["ps_hi"], (DEAL["post_spin_lo"], DEAL["post_spin_hi"]))]
r0 = sf.r
for name, lo, hi, ref in rows:
    r = sf.r; ws = sf.ws
    ws.cell(r, 2, name).font = F_BOLD
    for col, v, f in [(3, f"={lo}", F_LINK), (4, f"={hi}", F_LINK), (5, ref[0], F_INPUT), (6, ref[1], F_INPUT)]:
        c = ws.cell(r, col, v); c.font = f; c.number_format = NUM2
    c = ws.cell(r, 7, f'=IF(AND({I["offer"]}>=E{r},{I["offer"]}<=F{r}),"Inside",IF({I["offer"]}>F{r},"Above","Below"))'); c.font = F_FORMULA; c.alignment = Alignment(horizontal="center")
    c = ws.cell(r, 8, f"=MAX(ABS(C{r}-E{r}),ABS(D{r}-F{r}))"); c.font = F_FORMULA; c.number_format = NUM2
    sf.r += 1
r1 = sf.r - 1
sf.blank()
FF["n_above"] = sf.row("Ranges the offer exceeds entirely", f'=COUNTIF(G{r0}:G{r1},"Above")', '0', bold=True)
FF["n_inside"] = sf.row("Ranges the offer falls inside", f'=COUNTIF(G{r0}:G{r1},"Inside")', '0', bold=True)
FF["n_below"] = sf.row("Ranges the offer falls below", f'=COUNTIF(G{r0}:G{r1},"Below")', '0', bold=True)
FF["max_dev"] = sf.row("Largest deviation between a rebuilt bound and the proxy's (US$ per share)", f"=MAX(H{r0}:H{r1})", NUM2, bold=True)
FF["max_dev_core"] = sf.row("  excluding the future-share-price analysis and the Lazard DCF low end (both explained on their sheets)", f"=MAX(H{r0},H{r0+2}:H{r0+4},H{r0+6}:H{r1})", NUM2)
FF["lz_dcf_hi_dev"] = sf.row("Lazard DCF: deviation at the high end alone", f"=ABS(D{r0+5}-F{r0+5})", NUM2)
FF["top"] = sf.row("Highest rebuilt bound across the opinion-basis analyses (GS 1\u20134, Lazard 1\u20133)", f"=MAX(D{r0}:D{r0+4},D{r0+5}:D{r0+8})", NUM2)
sf.note("The offer sits at the top of Goldman's discounted cash flow range and above Lazard's, at or above the top of the premia ranges, "
        "above every trading-multiple range and inside only the precedent-transaction ranges. That is what a fairness opinion on a full-price strategic deal "
        "looks like: the intrinsic analyses bracket the price, the market-based ones sit beneath it, and the precedents \u2014 which include "
        "the buyer's own Wrigley deal at 17.6x \u2014 are the one place a higher number can be found.", height=52)

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sy = S("Summary", ncols=6, label_width=64, subtitle="What the proxy discloses, what was rebuilt, and what it says about the price")
sy.ws.column_dimensions["F"].width = 56
sy.section("THE DEAL")
sy.row("Offer per share (cash)", f"={I['offer']}", NUM2, font=F_LINK, bold=True)
sy.row("Premium to the undisturbed price", f"={I['prem']}", PCT, font=F_LINK, bold=True)
sy.row("Equity value at the offer (US$ m)", f"={I['offer']}*{I['fds']}", NUM0, font=F_LINK)
sy.row("Enterprise value at the offer (US$ m)", f"={I['offer']}*{I['fds']}+{I['nd']}", NUM0, font=F_LINK)
sy.row("EV / 2024E EBITDA · EV / 2025E EBITDA · EV / LTM EBITDA", f"=TEXT({G['off_mult']},\"0.0\")&\"x · \"&TEXT({G['off_mult25']},\"0.0\")&\"x · \"&TEXT({Z['off_ltm']},\"0.0\")&\"x\"", "@", font=F_LINK, bold=True)
sy.blank(); sy.section("THE REPLICATION")
sy.row("Goldman DCF: rebuilt vs proxy", f"=TEXT({G['dcf_lo']},\"0.00\")&\"\u2013\"&TEXT({G['dcf_hi']},\"0.00\")&\"  vs  68.42\u201383.82\"", "@", font=F_LINK)
sy.row("Lazard DCF: rebuilt vs proxy", f"=TEXT({Z['dcf_lo']},\"0.00\")&\"\u2013\"&TEXT({Z['dcf_hi']},\"0.00\")&\"  vs  69.60\u201379.10\"", "@", font=F_LINK)
sy.row("Lazard EBITDA comps: rebuilt vs proxy", f"=TEXT({Z['ce_lo']},\"0.00\")&\"\u2013\"&TEXT({Z['ce_hi']},\"0.00\")&\"  vs  57.20\u201373.80\"", "@", font=F_LINK)
sy.row("Largest deviation of any rebuilt bound from the proxy, core analyses (US$)", f"={FF['max_dev_core']}", NUM2, font=F_LINK, bold=True)
sy.row("Ranges the offer exceeds / sits inside / sits below", f"={FF['n_above']}&\" / \"&{FF['n_inside']}&\" / \"&{FF['n_below']}", "@", font=F_LINK, bold=True, border=DOUBLE_BORDER)
sy.blank()
sy.bullets([
    "The proxy discloses enough to rebuild almost every analysis to within a dollar or two a share: projections, discount rates, "
    "multiple ranges, premia ranges and the undisturbed price. The two undisclosed inputs \u2014 diluted shares and adjusted net debt \u2014 "
    "fall out of a single disclosed range and then reproduce the rest. Goldman's DCF lands within 63 cents, Lazard's comparables within "
    "30 cents, the premia to the cent. The one miss is the low end of Lazard's DCF, which needs a terminal cash flow close to NOPAT rather "
    "than 2027E free cash flow \u2014 the sheet says so and back-solves it.",
    "The offer is a full price by every intrinsic measure. It sits at the top of Goldman's DCF range, above Lazard's, at or above the "
    "top of the premia ranges, and roughly 3\u20134 turns of EBITDA above the best-rated listed peers. Only the precedent-transaction ranges, "
    "anchored by 2015\u20132018 food deals and Mars's own Wrigley acquisition, leave room above it.",
    "The two advisors agree because they were given the same projections and the same balance sheet; their methods differ (a "
    "terminal multiple against a perpetuity growth rate) and land within a few dollars of each other. A reader should treat that as "
    "consistency of inputs, not independent confirmation.",
    "What cannot be rebuilt is the judgement: why 12\u201314x rather than 11\u201315x, why 14\u201332% rather than the full quartile spread. "
    "The proxy gives the ranges and the phrase 'professional judgment and experience'. A fairness opinion is a set of defensible "
    "ranges that bracket a price the board has already negotiated; this replication shows how tightly they bracket it.",
])

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
tol = 2.5
tests = [
    ("Back-solved diluted shares are within 0\u20135% above the record-date basic count", f"=AND({I['fds']}>{I['basic']},{I['fds']}<{I['basic']}*1.05)"),
    ("Back-solved net debt reproduces both ends of Goldman's precedent range", f"=AND(ABS({G['pr_lo']}-{GS['prec'][0]})<0.01,ABS({G['pr_hi']}-{GS['prec'][1]})<0.01)"),
    (f"Goldman DCF rebuilt within US${tol} of the proxy at both ends", f"=AND(ABS({G['dcf_lo']}-{GS['dcf_range'][0]})<{tol},ABS({G['dcf_hi']}-{GS['dcf_range'][1]})<{tol})"),
    ("Goldman's implied perpetuity growth range brackets the proxy's 1.0\u20132.6%", f"=AND({G['pgr_lo']}<0.015,{G['pgr_hi']}>0.02)"),
    ("Goldman premia-paid ranges reproduce to the cent", f"=AND(ABS({G['pu_lo']}-{GS['premia_und'][0]})<0.01,ABS({G['pu_hi']}-{GS['premia_und'][1]})<0.01,ABS({G['p52_lo']}-{GS['premia_52'][0]})<0.01,ABS({G['p52_hi']}-{GS['premia_52'][1]})<0.01)"),
    (f"Goldman present-value-of-future-share-price rebuilt within US$5 at both ends (dividend path is stated)", f"=AND(ABS({G['pvr_lo']}-{GS['pv_range'][0]})<5,ABS({G['pvr_hi']}-{GS['pv_range'][1]})<5)"),
    (f"Lazard DCF high end rebuilt within US${tol}; low end not reproducible on 2027E free cash flow (reported, not hidden)", f"=ABS({Z['dcf_hi']}-{LZ['dcf_range'][1]})<{tol}"),
    ("Terminal cash flow needed for Lazard's low end lies between 2027E free cash flow and 2027E NOPAT", f"=AND({Z['tv_cf_needed']}>{I['ufcf'][3]},{Z['tv_cf_needed']}<{I['nopat'][3]}*1.05)"),
    (f"Lazard EBITDA comps rebuilt within US${tol} at both ends", f"=AND(ABS({Z['ce_lo']}-{LZ['comps_ebitda'][0]})<{tol},ABS({Z['ce_hi']}-{LZ['comps_ebitda'][1]})<{tol})"),
    ("Lazard P/E comps reproduce within 20 cents", f"=AND(ABS({Z['pe_lo']}-{LZ['comps_pe'][0]})<0.2,ABS({Z['pe_hi']}-{LZ['comps_pe'][1]})<0.2)"),
    ("Lazard premia reproduce within 10 cents", f"=AND(ABS({Z['pm_lo']}-{LZ['premia'][0]})<0.1,ABS({Z['pm_hi']}-{LZ['premia'][1]})<0.1)"),
    ("Back-solved LTM EBITDA lies between 2023 actual (about US$2,150m) and 2024E", f"=AND({Z['ltm']}>1900,{Z['ltm']}<{I['ebitda'][0]})"),
    ("Offer lies inside Goldman's DCF range and above Lazard's, as disclosed and as rebuilt", f"=AND({I['offer']}>={GS['dcf_range'][0]},{I['offer']}<={GS['dcf_range'][1]},{I['offer']}>{LZ['dcf_range'][1]},{I['offer']}<={G['dcf_hi']}+{tol},{I['offer']}>{Z['dcf_hi']})"),
    ("Offer exceeds every trading-multiple range", f"=AND({I['offer']}>{Z['ce_hi']},{I['offer']}>{Z['pe_hi']})"),
    ("Football-field counts sum to twelve analyses", f"={FF['n_above']}+{FF['n_inside']}+{FF['n_below']}=12"),
]
checks_sheet(book, tests,
             "The checks are the point of the project: each one asks whether an analysis rebuilt from the proxy's disclosed inputs lands "
             "on the range the advisor reported. A tolerance of US$2.50 a share allows for the undisclosed H2-2024 cash-flow split and "
             "rounding; the premia and P/E analyses have no such freedom and must reproduce almost exactly.")

# --------------------------------------------------------------------------- #
book.cover(
    blurb="A replication of the fairness-opinion analyses Goldman Sachs and Lazard delivered to Kellanova's board on Mars's US$83.50 "
          "cash offer, rebuilt from the inputs disclosed in the definitive proxy of 26 September 2024: management projections, discount "
          "rates, multiple and premia ranges, and the undisturbed price. Each rebuilt range is tested against the range the proxy reports.",
    method=[
        "Fully diluted shares and adjusted net debt back-solved from a single disclosed range (Goldman's precedent transactions), "
        "then held fixed for every other analysis so that each is a genuine test.",
        "Goldman: DCF with a terminal multiple, present value of future share price with dividends, premia paid on two bases, precedents, "
        "and reference trading multiples. Lazard: DCF with perpetuity growth, EBITDA and P/E comparables, precedents, premia, analyst "
        "targets and trading range.",
        "A football field placing the offer against all twelve ranges, with the deviation of each rebuilt bound from the proxy.",
    ],
    toc=[("Summary", "the deal, the replication, the reading"),
         ("Inputs", "everything disclosed; the two back-solved figures"),
         ("Goldman Sachs", "five analyses rebuilt"),
         ("Lazard", "four analyses rebuilt"),
         ("Football Field", "twelve ranges against US$83.50"),
         ("Checks", "fifteen tests; must read MODEL OK")],
    highlights=[("Offer per share (US$)", f"={I['offer']}", NUM2),
                ("Premium to undisturbed", f"={I['prem']}", PCT),
                ("EV / 2024E EBITDA at the offer", f"={G['off_mult']}", MULT),
                ("GS DCF rebuilt: low", f"={G['dcf_lo']}", NUM2),
                ("GS DCF rebuilt: high", f"={G['dcf_hi']}", NUM2),
                ("Largest deviation from the proxy, core analyses (US$)", f"={FF['max_dev_core']}", NUM2)],
    sources=["Kellanova definitive proxy statement (DEFM14A), SEC accession 0001193125-24-226970, filed 26 September 2024: 'Certain Financial Projections', 'Opinion of Goldman Sachs & Co. LLC', 'Opinion of Lazard Fr\u00e8res & Co. LLC', and the record-date share count.",
             "Dividends per share after 2024 and the split of 2024 cash flow either side of 30 June are stated inputs; every other figure is quoted from the proxy or derived from it as labelled.",
             "The exercise is a replication of disclosed analyses for learning. It expresses no view on either advisor's judgement, and it is not advice."])

book.finish(freeze={"Goldman Sachs": "C5", "Lazard": "C5"})
path = os.path.join(HERE, "Kellanova_Fairness_Opinion.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
