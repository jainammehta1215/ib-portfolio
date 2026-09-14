"""
Build 01_apple_three_statement_model.xlsx: an integrated 3-statement operating
model for Apple Inc. (FY2021-FY2025 historicals, FY2026E-FY2030E forecasts)
with a scenario selector, supporting schedules, a revolver plug and checks.

Conventions: blue = hard-coded input, black = formula, green = link to another
sheet. Interest is computed on opening balances to avoid circularity (a note on
the Inputs sheet explains how to switch to average balances in Excel).
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.properties import PageSetupProperties

# --------------------------------------------------------------------------- #
# Historical data (US$ billions) from Apple 10-K filings via Yahoo Finance pull
# --------------------------------------------------------------------------- #
IS = pd.read_csv("/tmp/aapl_IS.csv", index_col=0) / 1e9
BS = pd.read_csv("/tmp/aapl_BS.csv", index_col=0) / 1e9
CF = pd.read_csv("/tmp/aapl_CF.csv", index_col=0) / 1e9
HIST = ["2022", "2023", "2024", "2025"]          # four full years available across all three statements

def h(df, row, year):
    v = df.loc[row, year] if row in df.index and year in df.columns else float("nan")
    return 0.0 if pd.isna(v) else round(float(v), 3)

hist = {}
for y in HIST:
    hist[y] = dict(
        revenue=h(IS, "Total Revenue", y), cogs=h(IS, "Cost Of Revenue", y),
        rnd=h(IS, "Research And Development", y), sga=h(IS, "Selling General And Administration", y),
        other_inc=h(IS, "Other Non Operating Income Expenses", y),
        # interest: FY2024-25 not broken out by Yahoo; use net other income as reported. Documented on sheet.
        tax=h(IS, "Tax Provision", y), net_income=h(IS, "Net Income", y),
        dil_shares=h(IS, "Diluted Average Shares", y),
        cash=h(BS, "Cash Cash Equivalents And Short Term Investments", y),
        receivables=h(BS, "Receivables", y), inventory=h(BS, "Inventory", y),
        # "Other" lines are residuals to Apple's reported subtotals so every year ties to the 10-K
        # regardless of how the data vendor maps leases and non-current payables between years.
        other_ca=round(h(BS, "Current Assets", y) - h(BS, "Cash Cash Equivalents And Short Term Investments", y)
                       - h(BS, "Receivables", y) - h(BS, "Inventory", y), 3),
        lt_inv=h(BS, "Investments And Advances", y),
        net_ppe=h(BS, "Net PPE", y),
        other_nca=round(h(BS, "Total Non Current Assets", y) - h(BS, "Investments And Advances", y) - h(BS, "Net PPE", y), 3),
        ap=h(BS, "Accounts Payable", y),
        other_cl=round(h(BS, "Current Liabilities", y) - h(BS, "Accounts Payable", y) - h(BS, "Current Debt", y), 3),
        curr_debt=h(BS, "Current Debt", y), lt_debt=h(BS, "Long Term Debt", y),
        other_ncl=round(h(BS, "Total Non Current Liabilities Net Minority Interest", y) - h(BS, "Long Term Debt", y), 3),
        common=h(BS, "Common Stock", y), re=h(BS, "Retained Earnings", y),
        aoci=h(BS, "Gains Losses Not Affecting Retained Earnings", y),
        da=h(CF, "Depreciation And Amortization", y), sbc=h(CF, "Stock Based Compensation", y),
        capex=-h(CF, "Capital Expenditure", y), dividends=-h(CF, "Cash Dividends Paid", y),
        buybacks=-h(CF, "Repurchase Of Capital Stock", y),
        debt_issued=h(CF, "Issuance Of Debt", y), debt_repaid=-h(CF, "Repayment Of Debt", y),
        cfo=h(CF, "Operating Cash Flow", y), cfi=h(CF, "Investing Cash Flow", y), cff=h(CF, "Financing Cash Flow", y),
        shares_end=h(BS, "Ordinary Shares Number", y),
    )

# --------------------------------------------------------------------------- #
# Styles
# --------------------------------------------------------------------------- #
FONT = "Arial"
BLUE, BLACK, GREEN, GREY = "0000FF", "000000", "008000", "808080"
f_title = Font(name=FONT, size=14, bold=True)
f_hdr = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_sec = Font(name=FONT, size=10, bold=True)
f_in = Font(name=FONT, size=10, color=BLUE)
f_fx = Font(name=FONT, size=10, color=BLACK)
f_ln = Font(name=FONT, size=10, color=GREEN)
f_note = Font(name=FONT, size=9, italic=True, color=GREY)
fill_hdr = PatternFill("solid", fgColor="1F4E79")
fill_sec = PatternFill("solid", fgColor="DDEBF7")
fill_key = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
top_border = Border(top=Side(style="thin", color="000000"))
NUM = '#,##0.0;(#,##0.0);"-"'
PCT = '0.0%;(0.0%);"-"'
INT = '#,##0;(#,##0);"-"'
MULT = '0.00x'

FY_H = HIST                                    # historical columns in the model
FY_F = ["2026", "2027", "2028", "2029", "2030"]
FY = FY_H + FY_F
COL0 = 3                                       # column C = first year
def col(y): return L(COL0 + FY.index(y))
def colF(y): return L(3 + FY_F.index(y))       # Inputs sheet: forecast years start at column C

wb = Workbook()

# --------------------------------------------------------------------------- #
# Cover
# --------------------------------------------------------------------------- #
ws = wb.active; ws.title = "Cover"
ws.column_dimensions["A"].width = 3; ws.column_dimensions["B"].width = 34; ws.column_dimensions["C"].width = 90
ws["B2"] = "Apple Inc. — Three-Statement Operating Model"; ws["B2"].font = f_title
ws["B3"] = "IB portfolio · Project 1 of 20 · US$ billions unless stated · fiscal years ending late September"; ws["B3"].font = f_note
rows = [
    ("Purpose", "Integrated income statement, balance sheet and cash-flow forecast FY2026E–FY2030E driven from an assumptions sheet with Base / Upside / Downside cases."),
    ("How to use", "Change the scenario in Inputs!C4 (1 = Base, 2 = Upside, 3 = Downside). Edit blue cells only. Every other cell is a formula."),
    ("Colour code", "Blue = hard-coded input · Black = formula · Green = link to another sheet · Yellow fill = the scenario selector and key levers."),
    ("Sheets", "Inputs → Historical → Model (IS, BS, CF and schedules) → Checks → Summary."),
    ("Historicals", "FY2022–FY2025 from Apple's Form 10-K filings (SEC EDGAR, CIK 0000320193), pulled via Yahoo Finance and reconciled to reported totals on the Historical sheet."),
    ("Simplifications", "Interest on opening balances (no circularity); receivables include vendor non-trade receivables; other current liabilities include accrued expenses, deferred revenue and taxes payable; long-term marketable securities modelled as a run-off input; buybacks charged to retained earnings."),
    ("Source", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0000320193&type=10-K"),
    ("Author", "Jainam Mehta · github.com/jainammehta1215/ib-portfolio"),
]
for i, (k, v) in enumerate(rows, start=5):
    ws.cell(i, 2, k).font = f_sec; c = ws.cell(i, 3, v); c.font = Font(name=FONT, size=10); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[i].height = 30

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
wi = wb.create_sheet("Inputs")
wi.column_dimensions["A"].width = 3; wi.column_dimensions["B"].width = 44
for j in range(3, 9): wi.column_dimensions[L(j)].width = 12
wi.column_dimensions["I"].width = 60
wi["B2"] = "Assumptions"; wi["B2"].font = f_title
wi["B4"] = "Scenario selector (1 Base · 2 Upside · 3 Downside)"; wi["B4"].font = f_sec
wi["C4"] = 1; wi["C4"].font = Font(name=FONT, size=11, bold=True, color=BLUE); wi["C4"].fill = fill_key
wi["D4"] = '=CHOOSE($C$4,"Base","Upside","Downside")'; wi["D4"].font = f_fx
wi["B5"] = "Interest is computed on opening balances to avoid a circular reference. In Excel you may switch to average balances after enabling File → Options → Formulas → Iterative calculation."; wi["B5"].font = f_note
wi.merge_cells("B5:I5"); wi.row_dimensions[5].height = 28

# single-value inputs
wi["B7"] = "General inputs"; wi["B7"].font = f_sec; wi["B7"].fill = fill_sec
single = [
    ("Interest rate on debt", 0.040, PCT, "Blended coupon on Apple's notes; FY2025 interest expense ≈ 3.9% of average debt (10-K)"),
    ("Yield on cash & investments", 0.040, PCT, "Approximate yield on Apple's short- and long-term marketable securities"),
    ("Revolver interest rate", 0.055, PCT, "Assumed commercial-paper / revolver rate"),
    ("Minimum cash balance (US$bn)", 25.0, NUM, "Operating cash floor; revolver draws when forecast cash falls below this"),
    ("Share price at FY2025 year-end (US$)", 254.0, NUM, "Approximate closing price late September 2025; used only for buyback share count"),
    ("Share price growth p.a. for buyback math", 0.06, PCT, "Assumed; affects share count and EPS, not cash"),
]
SINGLE = {}
for i, (k, v, fmt, note) in enumerate(single, start=8):
    wi.cell(i, 2, k); c = wi.cell(i, 3, v); c.font = f_in; c.number_format = fmt; wi.cell(i, 9, note).font = f_note
    SINGLE[k] = f"Inputs!$C${i}"

# scenario driver blocks
wi["B15"] = "Forecast drivers by scenario (selected row feeds the Model)"; wi["B15"].font = f_sec; wi["B15"].fill = fill_sec
for j, y in enumerate(FY_F):
    c = wi.cell(15, 3 + j, f"FY{y}E"); c.font = f_hdr; c.fill = fill_hdr; c.alignment = Alignment(horizontal="center")
wi.cell(15, 9, "Basis / source").font = f_hdr; wi.cell(15, 9).fill = fill_hdr

drivers = [
    # key, label, fmt, base, upside, downside, note
    ("g_rev", "Revenue growth", PCT, [0.06, 0.05, 0.05, 0.04, 0.04], [0.09, 0.08, 0.07, 0.06, 0.06], [0.02, 0.02, 0.02, 0.02, 0.02],
     "FY2023–25 growth: -2.8%, +2.0%, +6.4%. Base assumes mid-single-digit growth led by services."),
    ("gm", "Gross margin", PCT, [0.470, 0.472, 0.474, 0.475, 0.475], [0.480, 0.485, 0.490, 0.492, 0.495], [0.455, 0.450, 0.450, 0.450, 0.450],
     "FY2025 gross margin 46.9%; services mix lifts it in Base and Upside; tariffs/pricing pressure in Downside."),
    ("rnd", "R&D as % of revenue", PCT, [0.083] * 5, [0.082] * 5, [0.086] * 5, "FY2025: 8.3%"),
    ("sga", "SG&A as % of revenue", PCT, [0.066] * 5, [0.064] * 5, [0.068] * 5, "FY2025: 6.6%"),
    ("tax", "Effective tax rate", PCT, [0.165] * 5, [0.160] * 5, [0.175] * 5, "FY2025: 15.6% (FY2024 24.1% distorted by the EU State Aid charge)"),
    ("da", "D&A as % of revenue", PCT, [0.028] * 5, [0.028] * 5, [0.029] * 5, "FY2023–25: 3.0%, 2.9%, 2.8%"),
    ("capex", "Capex as % of revenue", PCT, [0.031] * 5, [0.033] * 5, [0.028] * 5, "FY2025: 3.1%; Base holds it, Upside adds AI/data-centre spend"),
    ("sbc", "Stock-based compensation as % of revenue", PCT, [0.031] * 5, [0.031] * 5, [0.031] * 5, "FY2025: 3.1%; non-cash, added back in CFO and to equity"),
    ("dso", "Receivable days (on revenue)", NUM, [64] * 5, [62] * 5, [67] * 5, "FY2025: 64 days incl. vendor non-trade receivables"),
    ("dio", "Inventory days (on COGS)", NUM, [10] * 5, [9] * 5, [12] * 5, "FY2025: 9.5 days"),
    ("dpo", "Payable days (on COGS)", NUM, [115] * 5, [117] * 5, [110] * 5, "FY2025: 115 days"),
    ("oca", "Other current assets as % of revenue", PCT, [0.035] * 5, [0.035] * 5, [0.035] * 5, "FY2025: 3.5%"),
    ("ocl", "Other current liabilities as % of revenue", PCT, [0.180] * 5, [0.180] * 5, [0.180] * 5, "FY2025: 18.0% (accrued, deferred revenue, tax payable)"),
    ("onca", "Other non-current assets as % of revenue", PCT, [0.200] * 5, [0.200] * 5, [0.200] * 5, "FY2025: 20.1% incl. deferred tax assets"),
    ("oncl", "Other non-current liabilities as % of revenue", PCT, [0.100] * 5, [0.100] * 5, [0.100] * 5, "FY2025: 10.0%"),
    ("ltinv", "Change in long-term investments (US$bn, negative = sold)", NUM, [-10, -10, -10, -5, -5], [-15, -15, -10, -10, -5], [-5, -5, -5, -5, -5],
     "Apple has run down long-term securities from $121bn (FY22) to $78bn (FY25)"),
    ("div", "Dividends paid (US$bn)", NUM, [15.8, 16.2, 16.6, 17.0, 17.4], [16.0, 16.5, 17.0, 17.5, 18.0], [15.6, 15.8, 16.0, 16.2, 16.4], "FY2025: $15.4bn; ~4% p.a. per-share growth"),
    ("bb", "Share repurchases as % of FCF after dividends", PCT, [1.00] * 5, [1.10] * 5, [0.80] * 5,
     "Apple targets 'net cash neutral'; FY2023–25 buybacks were 92%, 102% and 109% of FCF after dividends"),
    ("rep", "Scheduled debt repayment (US$bn)", NUM, [10, 10, 10, 8, 8], [10, 10, 10, 8, 8], [10, 10, 10, 8, 8], "FY2025 maturities repaid $10.9bn; 10-K maturity schedule"),
    ("iss", "New long-term debt issued (US$bn)", NUM, [5, 5, 5, 5, 5], [0, 0, 0, 0, 0], [8, 8, 8, 8, 8], "FY2025: $4.5bn issued"),
]
DRV = {}                                       # key -> row number of the *selected* line
r = 16
for key, label, fmt, base, up, dn in [(d[0], d[1], d[2], d[3], d[4], d[5]) for d in drivers]:
    note = [d for d in drivers if d[0] == key][0][6]
    wi.cell(r, 2, label).font = f_sec
    wi.cell(r, 9, note).font = f_note
    for k, (nm, vals) in enumerate([("  Base", base), ("  Upside", up), ("  Downside", dn)], start=1):
        wi.cell(r + k, 2, nm)
        for j, v in enumerate(vals):
            c = wi.cell(r + k, 3 + j, v); c.font = f_in; c.number_format = fmt
    wi.cell(r + 4, 2, "  Selected").font = Font(name=FONT, size=10, bold=True)
    for j in range(5):
        cl = L(3 + j)
        c = wi.cell(r + 4, 3 + j, f"=CHOOSE($C$4,{cl}{r+1},{cl}{r+2},{cl}{r+3})"); c.font = f_fx; c.number_format = fmt; c.fill = fill_sec
    DRV[key] = r + 4
    r += 6

def drv(key, y):                               # reference to the selected driver for forecast year y
    return f"Inputs!{colF(y)}${DRV[key]}"

# --------------------------------------------------------------------------- #
# Historical
# --------------------------------------------------------------------------- #
wh = wb.create_sheet("Historical")
wh.column_dimensions["A"].width = 3; wh.column_dimensions["B"].width = 46
for j in range(3, 8): wh.column_dimensions[L(j)].width = 12
wh["B2"] = "Historical financials (US$bn) — Apple 10-K, FY2022–FY2025"; wh["B2"].font = f_title
wh["B3"] = "Reconciled: every subtotal ties to the reported statements. Receivables = trade + vendor non-trade receivables. 'Other' asset and liability lines are residuals to Apple's reported subtotals (so they include accrued expenses, deferred revenue, taxes payable and lease liabilities)."; wh["B3"].font = f_note
wh.merge_cells("B3:H3"); wh.row_dimensions[3].height = 28
for j, y in enumerate(HIST):
    c = wh.cell(5, 3 + j, f"FY{y}"); c.font = f_hdr; c.fill = fill_hdr; c.alignment = Alignment(horizontal="center")
HROW = {}
def hrow(r, label, key=None, fmt=NUM, bold=False, formula=None, section=False):
    c = wh.cell(r, 2, label); c.font = f_sec if (bold or section) else Font(name=FONT, size=10)
    if section: c.fill = fill_sec
    for j, y in enumerate(HIST):
        cell = wh.cell(r, 3 + j)
        if formula:
            cell.value = formula(L(3 + j)); cell.font = f_fx
        elif key:
            cell.value = hist[y][key]; cell.font = f_in
        cell.number_format = fmt
        if bold: cell.border = top_border
    if key: HROW[key] = r
    if formula and label not in HROW: HROW[label] = r
    return r

r = 6
hrow(r, "Income statement", section=True); r += 1
hrow(r, "Revenue", "revenue"); r += 1
hrow(r, "Cost of revenue", "cogs"); r += 1
hrow(r, "Gross profit", bold=True, formula=lambda c: f"={c}{HROW['revenue']}-{c}{HROW['cogs']}"); HROW["gp"] = r; r += 1
hrow(r, "Research & development", "rnd"); r += 1
hrow(r, "Selling, general & administrative", "sga"); r += 1
hrow(r, "Operating income (EBIT)", bold=True, formula=lambda c: f"={c}{HROW['gp']}-{c}{HROW['rnd']}-{c}{HROW['sga']}"); HROW["ebit"] = r; r += 1
hrow(r, "Other income / (expense), net incl. interest", "other_inc"); r += 1
hrow(r, "Pre-tax income", bold=True, formula=lambda c: f"={c}{HROW['ebit']}+{c}{HROW['other_inc']}"); HROW["pbt"] = r; r += 1
hrow(r, "Income tax", "tax"); r += 1
hrow(r, "Net income", bold=True, formula=lambda c: f"={c}{HROW['pbt']}-{c}{HROW['tax']}"); HROW["ni_calc"] = r; r += 1
hrow(r, "Net income as reported (check)", "net_income"); r += 1
hrow(r, "Diluted weighted average shares (bn)", "dil_shares", fmt='#,##0.000'); r += 1
hrow(r, "Diluted EPS (US$)", formula=lambda c: f"={c}{HROW['net_income']}/{c}{HROW['dil_shares']}", fmt='0.00'); HROW["eps"] = r; r += 2

hrow(r, "Balance sheet", section=True); r += 1
for lab, key in [("Cash & short-term investments", "cash"), ("Receivables (trade + vendor non-trade)", "receivables"), ("Inventory", "inventory"),
                 ("Other current assets", "other_ca")]:
    hrow(r, lab, key); r += 1
hrow(r, "Total current assets", bold=True, formula=lambda c: f"=SUM({c}{HROW['cash']}:{c}{HROW['other_ca']})"); HROW["tca"] = r; r += 1
for lab, key in [("Long-term marketable securities", "lt_inv"), ("Property, plant & equipment, net", "net_ppe"), ("Other non-current assets (incl. deferred tax)", "other_nca")]:
    hrow(r, lab, key); r += 1
hrow(r, "Total assets", bold=True, formula=lambda c: f"={c}{HROW['tca']}+{c}{HROW['lt_inv']}+{c}{HROW['net_ppe']}+{c}{HROW['other_nca']}"); HROW["ta"] = r; r += 2
for lab, key in [("Accounts payable", "ap"), ("Other current liabilities", "other_cl"), ("Current debt (incl. commercial paper)", "curr_debt")]:
    hrow(r, lab, key); r += 1
hrow(r, "Total current liabilities", bold=True, formula=lambda c: f"=SUM({c}{HROW['ap']}:{c}{HROW['curr_debt']})"); HROW["tcl"] = r; r += 1
for lab, key in [("Long-term debt", "lt_debt"), ("Other non-current liabilities", "other_ncl")]:
    hrow(r, lab, key); r += 1
hrow(r, "Total liabilities", bold=True, formula=lambda c: f"={c}{HROW['tcl']}+{c}{HROW['lt_debt']}+{c}{HROW['other_ncl']}"); HROW["tl"] = r; r += 1
for lab, key in [("Common stock & APIC", "common"), ("Retained earnings / (deficit)", "re"), ("Accumulated other comprehensive loss", "aoci")]:
    hrow(r, lab, key); r += 1
hrow(r, "Total shareholders' equity", bold=True, formula=lambda c: f"={c}{HROW['common']}+{c}{HROW['re']}+{c}{HROW['aoci']}"); HROW["te"] = r; r += 1
hrow(r, "Total liabilities & equity", bold=True, formula=lambda c: f"={c}{HROW['tl']}+{c}{HROW['te']}"); HROW["tle"] = r; r += 1
hrow(r, "Balance check (assets − liabilities − equity)", formula=lambda c: f"=ROUND({c}{HROW['ta']}-{c}{HROW['tle']},2)"); HROW["bal"] = r; r += 2

hrow(r, "Cash-flow items", section=True); r += 1
for lab, key in [("Depreciation & amortisation", "da"), ("Stock-based compensation", "sbc"), ("Capital expenditure", "capex"), ("Dividends paid", "dividends"),
                 ("Share repurchases", "buybacks"), ("Debt issued", "debt_issued"), ("Debt repaid", "debt_repaid"),
                 ("Cash from operations (reported)", "cfo"), ("Cash from investing (reported)", "cfi"), ("Cash from financing (reported)", "cff"),
                 ("Shares outstanding at year-end (bn)", "shares_end")]:
    hrow(r, lab, key, fmt='#,##0.000' if key == "shares_end" else NUM); r += 1
r += 1
hrow(r, "Historical ratios (calibrate the Inputs sheet)", section=True); r += 1
ratio_rows = [
    ("Revenue growth", lambda c: f'=IF(COLUMN({c}1)=3,"",{c}{HROW["revenue"]}/OFFSET({c}{HROW["revenue"]},0,-1)-1)', PCT),
    ("Gross margin", lambda c: f"={c}{HROW['gp']}/{c}{HROW['revenue']}", PCT),
    ("R&D % revenue", lambda c: f"={c}{HROW['rnd']}/{c}{HROW['revenue']}", PCT),
    ("SG&A % revenue", lambda c: f"={c}{HROW['sga']}/{c}{HROW['revenue']}", PCT),
    ("EBIT margin", lambda c: f"={c}{HROW['ebit']}/{c}{HROW['revenue']}", PCT),
    ("Effective tax rate", lambda c: f"={c}{HROW['tax']}/{c}{HROW['pbt']}", PCT),
    ("D&A % revenue", lambda c: f"={c}{HROW['da']}/{c}{HROW['revenue']}", PCT),
    ("Capex % revenue", lambda c: f"={c}{HROW['capex']}/{c}{HROW['revenue']}", PCT),
    ("SBC % revenue", lambda c: f"={c}{HROW['sbc']}/{c}{HROW['revenue']}", PCT),
    ("Receivable days", lambda c: f"={c}{HROW['receivables']}/{c}{HROW['revenue']}*365", NUM),
    ("Inventory days", lambda c: f"={c}{HROW['inventory']}/{c}{HROW['cogs']}*365", NUM),
    ("Payable days", lambda c: f"={c}{HROW['ap']}/{c}{HROW['cogs']}*365", NUM),
    ("Other current assets % revenue", lambda c: f"={c}{HROW['other_ca']}/{c}{HROW['revenue']}", PCT),
    ("Other current liabilities % revenue", lambda c: f"={c}{HROW['other_cl']}/{c}{HROW['revenue']}", PCT),
    ("Other non-current assets % revenue", lambda c: f"={c}{HROW['other_nca']}/{c}{HROW['revenue']}", PCT),
    ("Other non-current liabilities % revenue", lambda c: f"={c}{HROW['other_ncl']}/{c}{HROW['revenue']}", PCT),
    ("Net debt / (net cash) incl. LT securities", lambda c: f"={c}{HROW['curr_debt']}+{c}{HROW['lt_debt']}-{c}{HROW['cash']}-{c}{HROW['lt_inv']}", NUM),
]
for lab, fn, fmt in ratio_rows:
    hrow(r, lab, formula=fn, fmt=fmt); r += 1

# --------------------------------------------------------------------------- #
# Model
# --------------------------------------------------------------------------- #
wm = wb.create_sheet("Model")
wm.column_dimensions["A"].width = 3; wm.column_dimensions["B"].width = 46
for j in range(3, 3 + len(FY)): wm.column_dimensions[L(j)].width = 12
wm["B2"] = "Integrated model (US$bn)"; wm["B2"].font = f_title
wm["B3"] = '="Scenario: "&Inputs!$D$4&"   ·   FY2022–FY2025 actual (linked), FY2026E–FY2030E forecast"'; wm["B3"].font = f_note
for j, y in enumerate(FY):
    c = wm.cell(5, 3 + j, f"FY{y}" + ("E" if y in FY_F else "A")); c.font = f_hdr; c.fill = fill_hdr; c.alignment = Alignment(horizontal="center")
M = {}

def mrow(r, label, key, hist_key=None, fc=None, fmt=NUM, bold=False, section=False, hist_formula=None):
    """hist_key: link to Historical sheet for actual years; fc(y, c, p): forecast formula for year y in column c with previous column p."""
    M[key] = r                                  # register first so a row can reference itself (growth, margins)
    c = wm.cell(r, 2, label); c.font = f_sec if (bold or section) else Font(name=FONT, size=10)
    if section: c.fill = fill_sec
    for y in FY:
        cl = col(y); prev = L(COL0 + FY.index(y) - 1) if FY.index(y) > 0 else None
        cell = wm[f"{cl}{r}"]; cell.number_format = fmt
        if y in FY_H:
            if hist_formula is not None:
                cell.value = hist_formula(y, cl, prev); cell.font = f_fx
            elif hist_key is not None:
                cell.value = f"=Historical!{L(3 + HIST.index(y))}{HROW[hist_key]}"; cell.font = f_ln
        else:
            if fc is not None:
                cell.value = fc(y, cl, prev); cell.font = f_fx
        if bold: cell.border = top_border
    M[key] = r
    return r

r = 6
mrow(r, "INCOME STATEMENT", "s1", section=True); r += 1
mrow(r, "Revenue", "rev", "revenue", fc=lambda y, c, p: f"={p}{M['rev']}*(1+{drv('g_rev', y)})"); r += 1
mrow(r, "  growth %", "g", fmt=PCT, hist_formula=lambda y, c, p: (f"={c}{M['rev']}/{p}{M['rev']}-1" if p else None), fc=lambda y, c, p: f"={c}{M['rev']}/{p}{M['rev']}-1"); r += 1
mrow(r, "Cost of revenue", "cogs", "cogs", fc=lambda y, c, p: f"={c}{M['rev']}*(1-{drv('gm', y)})"); r += 1
mrow(r, "Gross profit", "gp", bold=True, hist_formula=lambda y, c, p: f"={c}{M['rev']}-{c}{M['cogs']}", fc=lambda y, c, p: f"={c}{M['rev']}-{c}{M['cogs']}"); r += 1
mrow(r, "  gross margin %", "gm", fmt=PCT, hist_formula=lambda y, c, p: f"={c}{M['gp']}/{c}{M['rev']}", fc=lambda y, c, p: f"={c}{M['gp']}/{c}{M['rev']}"); r += 1
mrow(r, "Research & development", "rnd", "rnd", fc=lambda y, c, p: f"={c}{M['rev']}*{drv('rnd', y)}"); r += 1
mrow(r, "Selling, general & administrative", "sga", "sga", fc=lambda y, c, p: f"={c}{M['rev']}*{drv('sga', y)}"); r += 1
mrow(r, "EBIT", "ebit", bold=True, hist_formula=lambda y, c, p: f"={c}{M['gp']}-{c}{M['rnd']}-{c}{M['sga']}", fc=lambda y, c, p: f"={c}{M['gp']}-{c}{M['rnd']}-{c}{M['sga']}"); r += 1
mrow(r, "  EBIT margin %", "ebitm", fmt=PCT, hist_formula=lambda y, c, p: f"={c}{M['ebit']}/{c}{M['rev']}", fc=lambda y, c, p: f"={c}{M['ebit']}/{c}{M['rev']}"); r += 1
mrow(r, "Depreciation & amortisation (memo, in COGS/opex)", "da", "da", fc=lambda y, c, p: f"={c}{M['rev']}*{drv('da', y)}"); r += 1
mrow(r, "EBITDA", "ebitda", hist_formula=lambda y, c, p: f"={c}{M['ebit']}+{c}{M['da']}", fc=lambda y, c, p: f"={c}{M['ebit']}+{c}{M['da']}"); r += 1
# interest rows (forecast only; historical uses reported net other income)
R_INT_INC = r; R_INT_EXP = r + 1; R_OTHER = r + 2
mrow(r, "Interest income on cash & investments (opening balances)", "int_inc",
     fc=lambda y, c, p: f"=({p}{{cash}}+{p}{{ltinv}})*{SINGLE['Yield on cash & investments']}"); r += 1
mrow(r, "Interest expense on debt (opening balances)", "int_exp",
     fc=lambda y, c, p: f"=-(({p}{{cdebt}}+{p}{{ltdebt}})*{SINGLE['Interest rate on debt']}+{p}{{rev_bal}}*{SINGLE['Revolver interest rate']})"); r += 1
mrow(r, "Other income / (expense), net", "other", "other_inc", fc=lambda y, c, p: f"={c}{M['int_inc']}+{c}{M['int_exp']}"); r += 1
mrow(r, "Pre-tax income", "pbt", bold=True, hist_formula=lambda y, c, p: f"={c}{M['ebit']}+{c}{M['other']}", fc=lambda y, c, p: f"={c}{M['ebit']}+{c}{M['other']}"); r += 1
mrow(r, "Income tax", "tax", "tax", fc=lambda y, c, p: f"={c}{M['pbt']}*{drv('tax', y)}"); r += 1
mrow(r, "Net income", "ni", bold=True, hist_formula=lambda y, c, p: f"={c}{M['pbt']}-{c}{M['tax']}", fc=lambda y, c, p: f"={c}{M['pbt']}-{c}{M['tax']}"); r += 1
mrow(r, "  net margin %", "nm", fmt=PCT, hist_formula=lambda y, c, p: f"={c}{M['ni']}/{c}{M['rev']}", fc=lambda y, c, p: f"={c}{M['ni']}/{c}{M['rev']}"); r += 2

mrow(r, "SHARE COUNT & EPS", "s2", section=True); r += 1
mrow(r, "Shares outstanding, year-end (bn)", "sh_end", "shares_end", fmt='#,##0.000',
     fc=lambda y, c, p: f"={p}{M['sh_end']}-{c}{{bb}}/{c}{{price}}+{c}{{sbc}}/{c}{{price}}*0.6"); r += 1
mrow(r, "  assumed share price (US$)", "price", fmt=NUM,
     hist_formula=lambda y, c, p: (f"={SINGLE['Share price at FY2025 year-end (US$)']}" if y == "2025" else None),
     fc=lambda y, c, p: f"={p}{M['price']}*(1+{SINGLE['Share price growth p.a. for buyback math']})"); r += 1
mrow(r, "Diluted weighted average shares (bn)", "sh_avg", "dil_shares", fmt='#,##0.000', fc=lambda y, c, p: f"=AVERAGE({p}{M['sh_end']},{c}{M['sh_end']})*1.003"); r += 1
mrow(r, "Diluted EPS (US$)", "eps", fmt='0.00', hist_formula=lambda y, c, p: f"={c}{M['ni']}/{c}{M['sh_avg']}", fc=lambda y, c, p: f"={c}{M['ni']}/{c}{M['sh_avg']}"); r += 2

mrow(r, "BALANCE SHEET", "s3", section=True); r += 1
mrow(r, "Cash & short-term investments", "cash", "cash", fc=lambda y, c, p: f"={c}{{cf_end}}"); r += 1
mrow(r, "Receivables", "ar", "receivables", fc=lambda y, c, p: f"={c}{M['rev']}*{drv('dso', y)}/365"); r += 1
mrow(r, "Inventory", "inv", "inventory", fc=lambda y, c, p: f"={c}{M['cogs']}*{drv('dio', y)}/365"); r += 1
mrow(r, "Other current assets", "oca", "other_ca", fc=lambda y, c, p: f"={c}{M['rev']}*{drv('oca', y)}"); r += 1
mrow(r, "Total current assets", "tca", bold=True, hist_formula=lambda y, c, p: f"=SUM({c}{M['cash']}:{c}{M['oca']})", fc=lambda y, c, p: f"=SUM({c}{M['cash']}:{c}{M['oca']})"); r += 1
mrow(r, "Long-term marketable securities", "ltinv", "lt_inv", fc=lambda y, c, p: f"={p}{M['ltinv']}+{drv('ltinv', y)}"); r += 1
mrow(r, "PP&E, net", "ppe", "net_ppe", fc=lambda y, c, p: f"={c}{{ppe_end}}"); r += 1
mrow(r, "Other non-current assets", "onca", "other_nca", fc=lambda y, c, p: f"={c}{M['rev']}*{drv('onca', y)}"); r += 1
mrow(r, "Total assets", "ta", bold=True, hist_formula=lambda y, c, p: f"={c}{M['tca']}+{c}{M['ltinv']}+{c}{M['ppe']}+{c}{M['onca']}", fc=lambda y, c, p: f"={c}{M['tca']}+{c}{M['ltinv']}+{c}{M['ppe']}+{c}{M['onca']}"); r += 2
mrow(r, "Accounts payable", "ap", "ap", fc=lambda y, c, p: f"={c}{M['cogs']}*{drv('dpo', y)}/365"); r += 1
mrow(r, "Other current liabilities", "ocl", "other_cl", fc=lambda y, c, p: f"={c}{M['rev']}*{drv('ocl', y)}"); r += 1
mrow(r, "Current debt (incl. commercial paper)", "cdebt", "curr_debt", fc=lambda y, c, p: f"={p}{M['cdebt']}"); r += 1
mrow(r, "Revolver (plug)", "rev_bal", hist_formula=lambda y, c, p: "=0", fc=lambda y, c, p: f"={c}{{revolver_end}}"); r += 1
mrow(r, "Total current liabilities", "tcl", bold=True, hist_formula=lambda y, c, p: f"=SUM({c}{M['ap']}:{c}{M['rev_bal']})", fc=lambda y, c, p: f"=SUM({c}{M['ap']}:{c}{M['rev_bal']})"); r += 1
mrow(r, "Long-term debt", "ltdebt", "lt_debt", fc=lambda y, c, p: f"={c}{{ltdebt_end}}"); r += 1
mrow(r, "Other non-current liabilities", "oncl", "other_ncl", fc=lambda y, c, p: f"={c}{M['rev']}*{drv('oncl', y)}"); r += 1
mrow(r, "Total liabilities", "tl", bold=True, hist_formula=lambda y, c, p: f"={c}{M['tcl']}+{c}{M['ltdebt']}+{c}{M['oncl']}", fc=lambda y, c, p: f"={c}{M['tcl']}+{c}{M['ltdebt']}+{c}{M['oncl']}"); r += 1
mrow(r, "Common stock & APIC", "common", "common", fc=lambda y, c, p: f"={p}{M['common']}+{c}{{sbc}}"); r += 1
mrow(r, "Retained earnings / (deficit)", "re", "re", fc=lambda y, c, p: f"={p}{M['re']}+{c}{M['ni']}-{c}{{div}}-{c}{{bb}}"); r += 1
mrow(r, "Accumulated other comprehensive loss", "aoci", "aoci", fc=lambda y, c, p: f"={p}{M['aoci']}"); r += 1
mrow(r, "Total equity", "te", bold=True, hist_formula=lambda y, c, p: f"={c}{M['common']}+{c}{M['re']}+{c}{M['aoci']}", fc=lambda y, c, p: f"={c}{M['common']}+{c}{M['re']}+{c}{M['aoci']}"); r += 1
mrow(r, "Total liabilities & equity", "tle", bold=True, hist_formula=lambda y, c, p: f"={c}{M['tl']}+{c}{M['te']}", fc=lambda y, c, p: f"={c}{M['tl']}+{c}{M['te']}"); r += 1
mrow(r, "Balance check (should be zero)", "bal", hist_formula=lambda y, c, p: f"=ROUND({c}{M['ta']}-{c}{M['tle']},3)", fc=lambda y, c, p: f"=ROUND({c}{M['ta']}-{c}{M['tle']},3)"); r += 2

mrow(r, "CASH-FLOW STATEMENT", "s4", section=True); r += 1
mrow(r, "Net income", "cf_ni", fc=lambda y, c, p: f"={c}{M['ni']}"); r += 1
mrow(r, "Depreciation & amortisation", "cf_da", fc=lambda y, c, p: f"={c}{M['da']}"); r += 1
mrow(r, "Stock-based compensation", "sbc", "sbc", fc=lambda y, c, p: f"={c}{M['rev']}*{drv('sbc', y)}"); r += 1
mrow(r, "(Increase) / decrease in receivables", "cf_ar", fc=lambda y, c, p: f"={p}{M['ar']}-{c}{M['ar']}"); r += 1
mrow(r, "(Increase) / decrease in inventory", "cf_inv", fc=lambda y, c, p: f"={p}{M['inv']}-{c}{M['inv']}"); r += 1
mrow(r, "(Increase) / decrease in other current assets", "cf_oca", fc=lambda y, c, p: f"={p}{M['oca']}-{c}{M['oca']}"); r += 1
mrow(r, "(Increase) / decrease in other non-current assets", "cf_onca", fc=lambda y, c, p: f"={p}{M['onca']}-{c}{M['onca']}"); r += 1
mrow(r, "Increase / (decrease) in accounts payable", "cf_ap", fc=lambda y, c, p: f"={c}{M['ap']}-{p}{M['ap']}"); r += 1
mrow(r, "Increase / (decrease) in other current liabilities", "cf_ocl", fc=lambda y, c, p: f"={c}{M['ocl']}-{p}{M['ocl']}"); r += 1
mrow(r, "Increase / (decrease) in other non-current liabilities", "cf_oncl", fc=lambda y, c, p: f"={c}{M['oncl']}-{p}{M['oncl']}"); r += 1
mrow(r, "Cash from operations", "cfo", "cfo", bold=True, fc=lambda y, c, p: f"=SUM({c}{M['cf_ni']}:{c}{M['cf_oncl']})"); r += 1
mrow(r, "Capital expenditure", "capex", hist_formula=lambda y, c, p: f"=-Historical!{L(3 + HIST.index(y))}{HROW['capex']}", fc=lambda y, c, p: f"=-{c}{M['rev']}*{drv('capex', y)}"); r += 1
mrow(r, "Net (purchase) / sale of long-term securities", "cf_ltinv", fc=lambda y, c, p: f"={p}{M['ltinv']}-{c}{M['ltinv']}"); r += 1
mrow(r, "Cash from investing", "cfi", "cfi", bold=True, fc=lambda y, c, p: f"={c}{M['capex']}+{c}{M['cf_ltinv']}"); r += 1
mrow(r, "Dividends paid", "div", hist_formula=lambda y, c, p: f"=Historical!{L(3 + HIST.index(y))}{HROW['dividends']}", fc=lambda y, c, p: f"={drv('div', y)}"); r += 1
mrow(r, "Share repurchases", "bb", hist_formula=lambda y, c, p: f"=Historical!{L(3 + HIST.index(y))}{HROW['buybacks']}", fc=lambda y, c, p: f"=MAX(0,({c}{{fcf}}-{c}{M['div']})*{drv('bb', y)})"); r += 1
mrow(r, "Long-term debt issued", "iss", fc=lambda y, c, p: f"={drv('iss', y)}"); r += 1
mrow(r, "Long-term debt repaid", "rep", fc=lambda y, c, p: f"={drv('rep', y)}"); r += 1
mrow(r, "Revolver draw / (repayment)", "rev_flow", fc=lambda y, c, p: f"={c}{{revolver_end}}-{p}{M['rev_bal']}"); r += 1
mrow(r, "Cash from financing", "cff", "cff", bold=True, fc=lambda y, c, p: f"=-{c}{M['div']}-{c}{M['bb']}+{c}{M['iss']}-{c}{M['rep']}+{c}{M['rev_flow']}"); r += 1
mrow(r, "Net change in cash", "dcash", hist_formula=lambda y, c, p: f"={c}{M['cfo']}+{c}{M['cfi']}+{c}{M['cff']}", fc=lambda y, c, p: f"={c}{M['cfo']}+{c}{M['cfi']}+{c}{M['cff']}"); r += 1
mrow(r, "Opening cash & short-term investments", "cf_beg", fc=lambda y, c, p: f"={p}{M['cash']}"); r += 1
mrow(r, "Closing cash & short-term investments", "cf_end", bold=True, fc=lambda y, c, p: f"={c}{M['cf_beg']}+{c}{M['dcash']}"); r += 2

mrow(r, "SCHEDULES", "s5", section=True); r += 1
mrow(r, "PP&E: opening", "ppe_beg", fc=lambda y, c, p: f"={p}{M['ppe']}"); r += 1
mrow(r, "  + capex", "ppe_capex", fc=lambda y, c, p: f"=-{c}{M['capex']}"); r += 1
mrow(r, "  − depreciation & amortisation", "ppe_da", fc=lambda y, c, p: f"=-{c}{M['da']}"); r += 1
mrow(r, "PP&E: closing", "ppe_end", bold=True, fc=lambda y, c, p: f"=SUM({c}{M['ppe_beg']}:{c}{M['ppe_da']})"); r += 2
mrow(r, "Long-term debt: opening", "ltd_beg", fc=lambda y, c, p: f"={p}{M['ltdebt']}"); r += 1
mrow(r, "  + issuance", "ltd_iss", fc=lambda y, c, p: f"={c}{M['iss']}"); r += 1
mrow(r, "  − repayment", "ltd_rep", fc=lambda y, c, p: f"=-{c}{M['rep']}"); r += 1
mrow(r, "Long-term debt: closing", "ltdebt_end", bold=True, fc=lambda y, c, p: f"=SUM({c}{M['ltd_beg']}:{c}{M['ltd_rep']})"); r += 2
mrow(r, "Revolver: opening balance", "rv_beg", fc=lambda y, c, p: f"={p}{M['rev_bal']}"); r += 1
mrow(r, "Cash before revolver", "cash_pre", fc=lambda y, c, p: f"={c}{M['cf_beg']}+{c}{M['cfo']}+{c}{M['cfi']}-{c}{M['div']}-{c}{M['bb']}+{c}{M['iss']}-{c}{M['rep']}"); r += 1
mrow(r, "Minimum cash", "min_cash", fc=lambda y, c, p: f"={SINGLE['Minimum cash balance (US$bn)']}"); r += 1
mrow(r, "Revolver: closing balance", "revolver_end", bold=True, fc=lambda y, c, p: f"=MAX(0,{c}{M['rv_beg']}+{c}{M['min_cash']}-{c}{M['cash_pre']})"); r += 2
mrow(r, "Net debt / (net cash) incl. LT securities", "netdebt", hist_formula=lambda y, c, p: f"={c}{M['cdebt']}+{c}{M['ltdebt']}+{c}{M['rev_bal']}-{c}{M['cash']}-{c}{M['ltinv']}",
     fc=lambda y, c, p: f"={c}{M['cdebt']}+{c}{M['ltdebt']}+{c}{M['rev_bal']}-{c}{M['cash']}-{c}{M['ltinv']}"); r += 1
mrow(r, "Free cash flow (CFO − capex)", "fcf", hist_formula=lambda y, c, p: f"={c}{M['cfo']}+{c}{M['capex']}", fc=lambda y, c, p: f"={c}{M['cfo']}+{c}{M['capex']}"); r += 1
mrow(r, "Shareholder returns (dividends + buybacks)", "payout", hist_formula=lambda y, c, p: f"={c}{M['div']}+{c}{M['bb']}", fc=lambda y, c, p: f"={c}{M['div']}+{c}{M['bb']}"); r += 1
mrow(r, "  as % of free cash flow", "payout_pct", fmt=PCT, hist_formula=lambda y, c, p: f"={c}{M['payout']}/{c}{M['fcf']}", fc=lambda y, c, p: f"={c}{M['payout']}/{c}{M['fcf']}"); r += 1

# resolve the forward references written as {key} placeholders
for row in wm.iter_rows(min_row=6, max_row=r):
    for cell in row:
        if isinstance(cell.value, str) and "{" in cell.value:
            v = cell.value
            for k, rr in M.items():
                v = v.replace("{" + k + "}", str(rr))
            cell.value = v

# freeze panes and print setup
wm.freeze_panes = "C6"
for sheet in (wi, wh, wm):
    sheet.sheet_view.showGridLines = False
    sheet.page_setup.orientation = "landscape"; sheet.page_setup.fitToWidth = 1; sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
wc = wb.create_sheet("Checks")
wc.column_dimensions["A"].width = 3; wc.column_dimensions["B"].width = 52
for j in range(3, 3 + len(FY)): wc.column_dimensions[L(j)].width = 12
wc["B2"] = "Model integrity checks (all should read 0 / TRUE)"; wc["B2"].font = f_title
for j, y in enumerate(FY):
    c = wc.cell(4, 3 + j, f"FY{y}"); c.font = f_hdr; c.fill = fill_hdr; c.alignment = Alignment(horizontal="center")
checks = [
    ("Balance sheet balances (assets − liabilities − equity)", lambda y, c: f"=Model!{c}{M['bal']}", NUM),
    ("Cash-flow closing cash = balance-sheet cash", lambda y, c: (f"=ROUND(Model!{c}{M['cf_end']}-Model!{c}{M['cash']},3)" if y in FY_F else "=0"), NUM),
    ("PP&E schedule closing = balance-sheet PP&E", lambda y, c: (f"=ROUND(Model!{c}{M['ppe_end']}-Model!{c}{M['ppe']},3)" if y in FY_F else "=0"), NUM),
    ("Debt schedule closing = balance-sheet LT debt", lambda y, c: (f"=ROUND(Model!{c}{M['ltdebt_end']}-Model!{c}{M['ltdebt']},3)" if y in FY_F else "=0"), NUM),
    ("Cash never below minimum (TRUE)", lambda y, c: (f"=Model!{c}{M['cash']}>=Model!{c}{M['min_cash']}-0.001" if y in FY_F else "=TRUE"), "General"),
    ("Historical net income ties to reported (0)", lambda y, c: (f"=ROUND(Historical!{L(3 + HIST.index(y))}{HROW['ni_calc']}-Historical!{L(3 + HIST.index(y))}{HROW['net_income']},1)" if y in FY_H else "=0"), NUM),
]
for i, (lab, fn, fmt) in enumerate(checks, start=5):
    wc.cell(i, 2, lab)
    for j, y in enumerate(FY):
        c = wc.cell(i, 3 + j, fn(y, col(y))); c.font = f_fx; c.number_format = fmt
wc.cell(12, 2, "Overall").font = f_sec
wc.cell(12, 3, f'=IF(AND(SUMPRODUCT(ABS(C5:{col(FY[-1])}8))<0.01,COUNTIF(C9:{col(FY[-1])}9,FALSE)=0,SUMPRODUCT(ABS(C10:{col(FY[-1])}10))<0.2),"MODEL OK","CHECK ERRORS")').font = Font(name=FONT, size=11, bold=True)
wc.sheet_view.showGridLines = False

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
wsu = wb.create_sheet("Summary")
wsu.column_dimensions["A"].width = 3; wsu.column_dimensions["B"].width = 40
for j in range(3, 3 + len(FY)): wsu.column_dimensions[L(j)].width = 12
wsu["B2"] = '="Key outputs — "&Inputs!$D$4&" case (US$bn)"'; wsu["B2"].font = f_title
wsu["B3"] = "Toggle Inputs!C4 to compare scenarios. Green = link to Model."; wsu["B3"].font = f_note
for j, y in enumerate(FY):
    c = wsu.cell(5, 3 + j, f"FY{y}" + ("E" if y in FY_F else "A")); c.font = f_hdr; c.fill = fill_hdr; c.alignment = Alignment(horizontal="center")
summ = [("Revenue", "rev", NUM), ("  growth", "g", PCT), ("Gross margin", "gm", PCT), ("EBIT", "ebit", NUM), ("  EBIT margin", "ebitm", PCT),
        ("Net income", "ni", NUM), ("Diluted EPS (US$)", "eps", '0.00'), ("Free cash flow", "fcf", NUM), ("Shareholder returns", "payout", NUM),
        ("Cash & ST investments", "cash", NUM), ("Total debt incl. revolver", None, NUM), ("Net debt / (net cash)", "netdebt", NUM), ("Shares outstanding (bn)", "sh_end", '#,##0.000')]
for i, (lab, key, fmt) in enumerate(summ, start=6):
    wsu.cell(i, 2, lab)
    for j, y in enumerate(FY):
        c = wsu.cell(i, 3 + j); c.number_format = fmt; c.font = f_ln
        if key: c.value = f"=Model!{col(y)}{M[key]}"
        else: c.value = f"=Model!{col(y)}{M['cdebt']}+Model!{col(y)}{M['ltdebt']}+Model!{col(y)}{M['rev_bal']}"
wsu.sheet_view.showGridLines = False
wb.move_sheet("Summary", offset=-3)

out = "/home/claude/ib-portfolio/01-apple-three-statement-model/Apple_3S_Model.xlsx"
import os; os.makedirs(os.path.dirname(out), exist_ok=True)
wb.save(out); print("saved", out, "| model rows", r)
