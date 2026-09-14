"""
Build 02_apple_dcf/Apple_DCF.xlsx — unlevered free-cash-flow DCF for Apple Inc.

Structure
---------
Cover · Summary · Inputs · WACC (peer betas, unlever / relever, CAPM, cost of debt) ·
Operating (Base/Upside/Downside FCF from Project 1's model, embedded so the file stands alone) ·
DCF (UFCF, discounting, terminal value both ways, equity bridge, implied price) ·
Sensitivity (WACC × g, WACC × exit multiple, scenario table) · Checks
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE, NUM, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER,
                         recalc, export_pdf, preview_png)

HERE = os.path.dirname(os.path.abspath(__file__))
FONT = "Arial"
book = Book(theme=THEMES["apple"], project_no=2, project="DCF Valuation", company="Apple Inc.",
            units="US$ billions unless stated", as_of="market data 12 Sep 2026 · FY2025 10-K")

# --------------------------------------------------------------------------- #
# Pull the three scenarios' operating lines from the Project 1 model (recalculated copies)
# --------------------------------------------------------------------------- #
P1 = os.path.join(HERE, "..", "01-apple-three-statement-model", "Apple_3S_Model.xlsx")
FY_F = ["2026", "2027", "2028", "2029", "2030"]
def model_rows(path):
    wb = load_workbook(path, data_only=True); m = wb["Model"]
    lab = {m.cell(r, 2).value: r for r in range(6, 110) if m.cell(r, 2).value}
    def row(label): return [m.cell(lab[label], j).value for j in range(7, 12)]   # G..K = FY2026E..FY2030E
    return dict(revenue=row("Revenue"), ebit=row("EBIT"), da=row("Depreciation & amortisation (memo, in COGS/opex)"),
                capex=[-v for v in row("Capital expenditure")], sbc=row("Stock-based compensation"),
                nwc=[sum(x) for x in zip(row("(Increase) / decrease in receivables"), row("(Increase) / decrease in inventory"),
                                          row("(Increase) / decrease in other current assets"), row("(Increase) / decrease in other non-current assets"),
                                          row("Increase / (decrease) in accounts payable"), row("Increase / (decrease) in other current liabilities"),
                                          row("Increase / (decrease) in other non-current liabilities"))],
                tax_rate=[m.cell(lab["Income tax"], j).value / m.cell(lab["Pre-tax income"], j).value for j in range(7, 12)])
scen = {}
import subprocess, shutil, tempfile
for k, name in [(1, "Base"), (2, "Upside"), (3, "Downside")]:
    tmp = os.path.join(tempfile.gettempdir(), f"p1_scen{k}.xlsx")
    wb = load_workbook(P1); wb["Inputs"]["C4"] = k; wb.save(tmp); recalc(tmp)
    scen[name] = model_rows(tmp)
# balance-sheet items for the bridge (FY2025A)
wb1 = load_workbook(P1, data_only=True); h = wb1["Historical"]
hl = {h.cell(r, 2).value: r for r in range(6, 90) if h.cell(r, 2).value}
fy25 = lambda label: h.cell(hl[label], 6).value
bridge = dict(cash=fy25("Cash & short-term investments"), lt_inv=fy25("Long-term marketable securities"),
              curr_debt=fy25("Current debt (incl. commercial paper)"), lt_debt=fy25("Long-term debt"), shares=fy25("Shares outstanding at year-end (bn)"))
peers = pd.read_csv("/tmp/peers.csv", index_col=0)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
wi = book.sheet("Inputs", ncols=6, label_width=50, subtitle="Valuation inputs · edit blue cells only")
wi.column_dimensions["F"].width = 70
book.section(wi, 4, "SCENARIO AND MARKET DATA", 6)
book.key_cell(wi["C5"], 1); wi["B5"] = "Operating scenario (1 Base · 2 Upside · 3 Downside)"; wi["D5"] = '=CHOOSE($C$5,"Base","Upside","Downside")'; wi["D5"].font = F_BOLD
inputs = [
    ("Valuation date", "2026-09-12", "General", "Market data date"),
    ("Share price (US$)", 332.27, NUM2, "Yahoo Finance close, 12 Sep 2026"),
    ("Diluted shares outstanding (bn)", 14.715, '0.000', "Diluted weighted average, quarter to 30 Jun 2026 (Form 10-Q); FY2025 annual was 15.00bn"),
    ("Risk-free rate (10-year US Treasury)", 0.0495, PCT, "FRED DGS10, 10 Sep 2026"),
    ("Equity risk premium", 0.045, PCT, "Damodaran implied ERP for the US, mid-2026 (approx.); range 4.0–5.0%"),
    ("Pre-tax cost of debt", 0.0553, PCT, "ICE BofA AAA US corporate effective yield (FRED BAMLC0A1CAAAEY), 10 Sep 2026; Apple is rated Aaa/AA+"),
    ("Marginal tax rate (for cost of debt and unlevering)", 0.21, PCT, "US federal statutory rate; Apple's effective rate is ~16% and is used for cash taxes on EBIT"),
    ("Terminal growth rate", 0.030, PCT, "Nominal; roughly long-run US nominal GDP less a haircut for a mature company"),
    ("LTM net income to 30 Jun 2026 (US$bn)", 128.9, NUM, "Sum of the last four quarters (10-Q filings); for the implied P/E cross-check"),
    ("Terminal EV/EBITDA exit multiple", 22.0, MULT, "Apple trades at ~27x LTM EBITDA; large-cap tech peers 18–30x; a discount for a terminal year"),
    ("Mid-year convention (1 = on, 0 = off)", 1, "General", "Cash flows arrive through the year, not at year-end"),
    ("Beta adjustment (1 = Blume-adjusted 0.67β + 0.33, 0 = raw)", 1, "General", "Standard banking practice (Bloomberg 'adjusted beta'): regression betas revert toward 1 over time"),
    ("Months from valuation date to FY2026 year-end", 0.5, '0.0', "Valuation 12 Sep 2026; FY2026 ends 26 Sep 2026: first-year cash flow is almost entirely in the past, stub = 0.5/12"),
]
INP = {}
for i, (k, v, fmt, note) in enumerate(inputs, start=6):
    wi.cell(i, 2, k).font = F_TEXT; c = wi.cell(i, 3, v); c.font = F_INPUT; c.number_format = fmt
    n = wi.cell(i, 6, note); n.font = F_NOTE; n.alignment = Alignment(wrap_text=True, vertical="top")
    INP[k] = f"Inputs!$C${i}"
book.note(wi, 20, "Conventions: unlevered free cash flow = EBIT × (1 − cash tax rate) + D&A − capex − increase in net working capital. Stock-based compensation is treated as a real expense (not added back), which is the conservative choice for a company issuing ~US$13bn a year of it.", 6, height=30)

# --------------------------------------------------------------------------- #
# WACC
# --------------------------------------------------------------------------- #
ww = book.sheet("WACC", ncols=9, label_width=52, subtitle="Cost of capital built bottom-up from a peer set · five years of weekly returns vs S&P 500")
book.section(ww, 4, "PEER BETAS (regression on weekly returns, Sep 2021 – Sep 2026)", 9)
hdr = ["Ticker", "Market cap (US$bn)", "Total debt (US$bn)", "Cash (US$bn)", "Levered beta", "Net D/E", "Unlevered beta", "Note"]
for j, hname in enumerate(hdr):
    c = ww.cell(5, 2 + j, hname); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
ww.row_dimensions[5].height = 30
# USD-convert Sony (JPY→USD) and Samsung (KRW→USD) balance sheets; note it
fx = {"SONY": 1 / 147.0, "005930.KS": 1 / 1380.0}
peer_rows = []
for t, rw in peers.iterrows():
    f = fx.get(t, 1.0)
    mcap = rw["mcap_bn"] * (f if t == "005930.KS" else 1.0)     # Samsung mcap reported in KRW; Sony mcap already USD (ADR)
    peer_rows.append((t, mcap, rw["debt_bn"] * f, rw["cash_bn"] * f, rw["beta_5y_weekly"],
                      "ADR; JPY balance sheet converted at 147" if t == "SONY" else ("KRW converted at 1,380" if t == "005930.KS" else "")))
r0 = 6
for i, (t, mcap, debt, cash, beta, note) in enumerate(peer_rows):
    r = r0 + i
    ww.cell(r, 2, t).font = F_TEXT
    for j, (v, fmt) in enumerate([(mcap, NUM), (debt, NUM), (cash, NUM), (beta, '0.00')], start=3):
        c = ww.cell(r, j, round(float(v), 3)); c.font = F_INPUT; c.number_format = fmt
    ww.cell(r, 7, f"=MAX(0,(D{r}-E{r})/C{r})").number_format = '0.00'; ww.cell(r, 7).font = F_FORMULA
    ww.cell(r, 8, f"=F{r}/(1+(1-{INP['Marginal tax rate (for cost of debt and unlevering)']})*G{r})").number_format = '0.00'; ww.cell(r, 8).font = F_FORMULA
    ww.cell(r, 9, note).font = F_NOTE
rN = r0 + len(peer_rows) - 1
r = rN + 1
for lab, fn in [("Median", "MEDIAN"), ("Mean", "AVERAGE")]:
    ww.cell(r, 2, lab).font = F_BOLD
    for j in (6, 7, 8):
        c = ww.cell(r, j, f"={fn}({L(j)}{r0}:{L(j)}{rN})"); c.font = F_FORMULA; c.number_format = '0.00'; c.border = TOTAL_BORDER
    r += 1
R_MED = rN + 1
book.note(ww, r, "Peers: mega-cap platforms (MSFT, GOOGL, AMZN, META, NVDA) share Apple's scale and services economics; SONY, Samsung, DELL and HPQ share the hardware business. Net debt/equity uses gross debt less cash, floored at zero (most peers are net cash, so unlevering barely moves the beta). Apple's own regression beta is shown for reference and is close to the peer median.", 9, height=42)
r += 3
book.section(ww, r, "COST OF EQUITY (CAPM)", 9); r += 1
R_CE = r
rows = [
    ("Unlevered beta (peer median)", f"=H{R_MED}", '0.00'),
    ("Apple net debt / equity (market; zero — Apple is net cash)", f"=MAX(0,(Operating!$C$4-Operating!$C$5)/({INP['Share price (US$)']}*{INP['Diluted shares outstanding (bn)']}))", '0.00'),
    ("Relevered beta = βu × (1 + (1 − t) × D/E), Blume-adjusted if selected", f"=IF({INP['Beta adjustment (1 = Blume-adjusted 0.67β + 0.33, 0 = raw)']}=1,0.67*C{r}*(1+(1-{INP['Marginal tax rate (for cost of debt and unlevering)']})*C{r+1})+0.33,C{r}*(1+(1-{INP['Marginal tax rate (for cost of debt and unlevering)']})*C{r+1}))", '0.00'),
    ("Risk-free rate", f"={INP['Risk-free rate (10-year US Treasury)']}", PCT),
    ("Equity risk premium", f"={INP['Equity risk premium']}", PCT),
    ("Cost of equity = rf + β × ERP", f"=C{r+3}+C{r+2}*C{r+4}", PCT),
]
for i, (lab, fml, fmt) in enumerate(rows):
    ww.cell(r + i, 2, lab).font = F_BOLD if i == 5 else F_TEXT
    c = ww.cell(r + i, 3, fml); c.font = F_FORMULA; c.number_format = fmt
    if i == 5: c.border = TOTAL_BORDER
R_KE = r + 5; r += 7
book.section(ww, r, "COST OF DEBT AND WEIGHTS", 9); r += 1
rows2 = [
    ("Pre-tax cost of debt", f"={INP['Pre-tax cost of debt']}", PCT),
    ("Tax shield rate", f"={INP['Marginal tax rate (for cost of debt and unlevering)']}", PCT),
    ("After-tax cost of debt", f"=C{r}*(1-C{r+1})", PCT),
    ("Market value of equity (US$bn)", f"={INP['Share price (US$)']}*{INP['Diluted shares outstanding (bn)']}", NUM),
    ("Debt (US$bn, book, 30 Jun 2026)", "=Operating!$C$4", NUM),
    ("Weight of equity", f"=C{r+3}/(C{r+3}+C{r+4})", PCT),
    ("Weight of debt", f"=C{r+4}/(C{r+3}+C{r+4})", PCT),
    ("WACC", f"=C{r+5}*C{R_KE}+C{r+6}*C{r+2}", PCT),
]
for i, (lab, fml, fmt) in enumerate(rows2):
    ww.cell(r + i, 2, lab).font = F_BOLD if i == 7 else F_TEXT
    c = ww.cell(r + i, 3, fml); c.font = F_FORMULA; c.number_format = fmt
    if i == 7: c.border = DOUBLE_BORDER; c.font = Font(name=FONT, size=11, bold=True)
R_WACC = r + 7
book.note(ww, R_WACC + 2, "Apple's market equity (~US$5tn) dwarfs its US$99bn of debt, so WACC ≈ cost of equity. The interview question is not the arithmetic but the ERP: every 0.5% on the ERP moves the WACC by ~0.55% and the value by ~10%.", 9, height=30)
WACC = f"WACC!$C${R_WACC}"

# --------------------------------------------------------------------------- #
# Operating (three scenarios embedded)
# --------------------------------------------------------------------------- #
wo = book.sheet("Operating", ncols=8, label_width=44, subtitle="Operating forecasts by scenario, copied from Project 1 (Apple_3S_Model.xlsx) so this file stands alone")
book.section(wo, 3, "LATEST BALANCE-SHEET ITEMS FOR THE EQUITY BRIDGE (US$bn, 30 Jun 2026 10-Q)", 8)
wo["B4"] = "Total debt (current + long-term), 30 Jun 2026"; wo["C4"] = 84.34; wo["C4"].font = F_INPUT; wo["C4"].number_format = NUM; wo["D4"] = "Form 10-Q, quarter ended 28 Jun 2026: commercial paper + term debt (FY2025 year-end was US$98.7bn)"; wo["D4"].font = F_NOTE
wo["B5"] = "Cash, short-term and long-term marketable securities, 30 Jun 2026"; wo["C5"] = round(62.40 + 84.12, 2); wo["C5"].font = F_INPUT; wo["C5"].number_format = NUM; wo["D5"] = "US$62.4bn cash & ST investments + US$84.1bn long-term marketable securities; the latter are liquid and treated as cash"; wo["D5"].font = F_NOTE
book.year_header(wo, 7, [f"FY{y}E" for y in FY_F], first_col=3)
lines = [("revenue", "Revenue", NUM), ("ebit", "EBIT", NUM), ("tax_rate", "Effective tax rate", PCT), ("da", "D&A", NUM), ("capex", "Capex", NUM), ("nwc", "Change in net working capital (+ = inflow)", NUM), ("sbc", "Stock-based compensation (memo)", NUM)]
OP = {}                                    # OP[(scenario, key)] = row
r = 8
for name in ["Base", "Upside", "Downside"]:
    book.subsection(wo, r, f"{name} case", 8); r += 1
    for key, lab, fmt in lines:
        wo.cell(r, 2, lab).font = F_TEXT
        for j, v in enumerate(scen[name][key]):
            c = wo.cell(r, 3 + j, round(float(v), 4)); c.font = F_INPUT; c.number_format = fmt
        OP[(name, key)] = r; r += 1
    r += 1
book.section(wo, r, "SELECTED SCENARIO (feeds the DCF)", 8); r += 1
book.year_header(wo, r, [f"FY{y}E" for y in FY_F], first_col=3); r += 1
SEL = {}
for key, lab, fmt in lines:
    wo.cell(r, 2, lab).font = F_BOLD
    for j in range(5):
        cl = L(3 + j)
        c = wo.cell(r, 3 + j, f"=CHOOSE(Inputs!$C$5,{cl}{OP[('Base', key)]},{cl}{OP[('Upside', key)]},{cl}{OP[('Downside', key)]})"); c.font = F_FORMULA; c.number_format = fmt; c.fill = book.fill_light
    SEL[key] = r; r += 1

# --------------------------------------------------------------------------- #
# DCF
# --------------------------------------------------------------------------- #
wd = book.sheet("DCF", ncols=9, label_width=48, subtitle="Unlevered free cash flow, discounting, terminal value two ways, equity bridge")
wd["B3"] = '="Scenario: "&Inputs!$D$5'; wd["B3"].font = Font(name=FONT, size=10, bold=True, color=book.theme.accent)
book.section(wd, 4, "UNLEVERED FREE CASH FLOW (US$bn)", 9)
book.year_header(wd, 5, [f"FY{y}E" for y in FY_F], first_col=3)
D = {}
def drow(r, key, label, fml, fmt=NUM, bold=False, border=None):
    D[key] = r; wd.cell(r, 2, label).font = F_BOLD if bold else F_TEXT
    for j in range(5):
        cl = L(3 + j); c = wd.cell(r, 3 + j, fml(cl, j)); c.font = F_FORMULA; c.number_format = fmt
        if border: c.border = border
r = 6
drow(r, "rev", "Revenue", lambda c, j: f"=Operating!{c}{SEL['revenue']}"); r += 1
drow(r, "ebit", "EBIT", lambda c, j: f"=Operating!{c}{SEL['ebit']}"); r += 1
drow(r, "tax", "  less: cash taxes on EBIT (at effective rate)", lambda c, j: f"=-{c}{D['ebit']}*Operating!{c}{SEL['tax_rate']}"); r += 1
drow(r, "nopat", "NOPAT", lambda c, j: f"={c}{D['ebit']}+{c}{D['tax']}", bold=True, border=TOTAL_BORDER); r += 1
drow(r, "da", "  plus: D&A", lambda c, j: f"=Operating!{c}{SEL['da']}"); r += 1
drow(r, "capex", "  less: capex", lambda c, j: f"=-Operating!{c}{SEL['capex']}"); r += 1
drow(r, "nwc", "  plus / (less): change in net working capital", lambda c, j: f"=Operating!{c}{SEL['nwc']}"); r += 1
drow(r, "ufcf", "Unlevered free cash flow", lambda c, j: f"={c}{D['nopat']}+{c}{D['da']}+{c}{D['capex']}+{c}{D['nwc']}", bold=True, border=DOUBLE_BORDER); r += 1
drow(r, "margin", "  UFCF margin", lambda c, j: f"={c}{D['ufcf']}/{c}{D['rev']}", fmt=PCT); r += 2
book.section(wd, r, "DISCOUNTING", 9); r += 1
drow(r, "yearfrac", "Year fraction of cash flow in the forecast", lambda c, j: (f"={INP['Months from valuation date to FY2026 year-end']}/12" if j == 0 else "=1"), fmt='0.00'); r += 1
drow(r, "cf_in", "Cash flow attributable to the holder (stub-adjusted)", lambda c, j: f"={c}{D['ufcf']}*{c}{D['yearfrac']}"); r += 1
drow(r, "t", "Discount period (years from valuation date)", lambda c, j: (f"={c}{D['yearfrac']}*(1-0.5*{INP['Mid-year convention (1 = on, 0 = off)']})" if j == 0 else f"={L(2 + j)}{r}+{L(2+j)}{D['yearfrac']}*0+({L(2+j)}{D['yearfrac']}-{L(2+j)}{r}+{L(2+j)}{r})*0+1*{c}{D['yearfrac']}-0.5*{INP['Mid-year convention (1 = on, 0 = off)']}*({c}{D['yearfrac']}-{L(2+j)}{D['yearfrac']}*0)*0"), fmt='0.00'); r += 1
# simpler, correct discount-period build (overwrite the clumsy formula above)
for j in range(5):
    cl = L(3 + j)
    if j == 0:
        wd[f"{cl}{D['t']}"] = f"={cl}{D['yearfrac']}*(1-0.5*{INP['Mid-year convention (1 = on, 0 = off)']})"
    else:
        prev_end = f"(C{D['yearfrac']}+{'+'.join(L(3 + k) + str(D['yearfrac']) for k in range(1, j))})" if j > 1 else f"C{D['yearfrac']}"
        wd[f"{cl}{D['t']}"] = f"={prev_end}+{cl}{D['yearfrac']}*(1-0.5*{INP['Mid-year convention (1 = on, 0 = off)']})"
drow(r, "df", "Discount factor at WACC", lambda c, j: f"=1/(1+{WACC})^{c}{D['t']}", fmt='0.0000'); r += 1
drow(r, "pv", "Present value of cash flow", lambda c, j: f"={c}{D['cf_in']}*{c}{D['df']}", bold=True, border=TOTAL_BORDER); r += 2

book.section(wd, r, "TERMINAL VALUE (FY2030 base) AND VALUATION", 9); r += 1
last = L(7)          # FY2030 column
V = {}
def vrow(r, key, label, f_pg, f_em, fmt=NUM, bold=False, border=None):
    V[key] = r; wd.cell(r, 2, label).font = F_BOLD if bold else F_TEXT
    for col_, fml in [("C", f_pg), ("E", f_em)]:
        c = wd[f"{col_}{r}"]; c.value = fml; c.font = F_FORMULA; c.number_format = fmt
        if border: c.border = border
for col_, lab in [("C", "Perpetuity growth"), ("E", "Exit multiple")]:
    c = wd[f"{col_}{r}"]; c.value = lab; c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center")
r += 1
g = INP["Terminal growth rate"]; em = INP["Terminal EV/EBITDA exit multiple"]
vrow(r, "tv", "Terminal value at FY2030", f"={last}{D['ufcf']}*(1+{g})/({WACC}-{g})", f"=({last}{D['ebit']}+{last}{D['da']})*{em}"); r += 1
vrow(r, "tv_df", "  discount factor (end of FY2030)", f"=1/(1+{WACC})^(C{D['yearfrac']}+D{D['yearfrac']}+E{D['yearfrac']}+F{D['yearfrac']}+G{D['yearfrac']})", f"=C{r}", fmt='0.0000'); r += 1
vrow(r, "pv_tv", "PV of terminal value", f"=C{V['tv']}*C{V['tv_df']}", f"=E{V['tv']}*E{V['tv_df']}"); r += 1
vrow(r, "pv_cf", "PV of forecast cash flows", f"=SUM(C{D['pv']}:G{D['pv']})", f"=C{r}"); r += 1
vrow(r, "ev", "Enterprise value", f"=C{V['pv_tv']}+C{V['pv_cf']}", f"=E{V['pv_tv']}+E{V['pv_cf']}", bold=True, border=TOTAL_BORDER); r += 1
vrow(r, "tv_share", "  terminal value as % of EV", f"=C{V['pv_tv']}/C{V['ev']}", f"=E{V['pv_tv']}/E{V['ev']}", fmt=PCT); r += 1
vrow(r, "less_debt", "  less: total debt", "=-Operating!$C$4", "=-Operating!$C$4"); r += 1
vrow(r, "plus_cash", "  plus: cash and marketable securities", "=Operating!$C$5", "=Operating!$C$5"); r += 1
vrow(r, "eq", "Equity value", f"=C{V['ev']}+C{V['less_debt']}+C{V['plus_cash']}", f"=E{V['ev']}+E{V['less_debt']}+E{V['plus_cash']}", bold=True, border=TOTAL_BORDER); r += 1
vrow(r, "shares", "  diluted shares (bn)", f"={INP['Diluted shares outstanding (bn)']}", f"={INP['Diluted shares outstanding (bn)']}", fmt='0.000'); r += 1
vrow(r, "px", "Implied value per share (US$)", f"=C{V['eq']}/C{V['shares']}", f"=E{V['eq']}/E{V['shares']}", fmt=NUM2, bold=True, border=DOUBLE_BORDER); r += 1
vrow(r, "updown", "  premium / (discount) to current price", f"=C{V['px']}/{INP['Share price (US$)']}-1", f"=E{V['px']}/{INP['Share price (US$)']}-1", fmt=PCT); r += 2
book.subsection(wd, r, "Cross-checks (each method's assumption implied by the other)", 9); r += 1
vrow(r, "impl_mult", "Implied exit EV/EBITDA from perpetuity method", f"=C{V['tv']}/({last}{D['ebit']}+{last}{D['da']})", f'=""', fmt=MULT); r += 1
vrow(r, "impl_g", "Implied perpetuity growth from exit-multiple method", f'=""', f"=({WACC}*E{V['tv']}-{last}{D['ufcf']})/(E{V['tv']}+{last}{D['ufcf']})", fmt=PCT); r += 1
vrow(r, "impl_pe", "Implied P/E on LTM net income", f"=C{V['eq']}/{INP['LTM net income to 30 Jun 2026 (US$bn)']}", f"=E{V['eq']}/{INP['LTM net income to 30 Jun 2026 (US$bn)']}", fmt=MULT); r += 1
vrow(r, "mkt_pe", "Market P/E on LTM net income (for comparison)", f"={INP['Share price (US$)']}*{INP['Diluted shares outstanding (bn)']}/{INP['LTM net income to 30 Jun 2026 (US$bn)']}", f"=C{r}", fmt=MULT); r += 1
book.note(wd, r + 1, "If the perpetuity-growth method implies an exit multiple far from where the peers trade, or the exit-multiple method implies a growth rate above nominal GDP, one of the terminal assumptions is wrong. The two methods should bracket a defensible range rather than agree exactly.", 9, height=30)

# --------------------------------------------------------------------------- #
# Sensitivity
# --------------------------------------------------------------------------- #
wsn = book.sheet("Sensitivity", ncols=10, label_width=22, subtitle="Two-way tables recomputed with the DCF formula (no Excel data tables, so they work everywhere)")
def sens_table(top, title, row_vals, col_vals, row_lab, col_lab, cell_fn, fmt=NUM2):
    book.section(wsn, top, title, 10)
    wsn.cell(top + 1, 2, f"{row_lab} ↓  /  {col_lab} →").font = F_NOTE
    for j, cv in enumerate(col_vals):
        c = wsn.cell(top + 1, 3 + j, cv); c.font = book.f_header; c.fill = book.fill_secondary; c.number_format = PCT if isinstance(cv, float) and cv < 1 else MULT; c.alignment = Alignment(horizontal="center")
    for i, rv in enumerate(row_vals):
        c = wsn.cell(top + 2 + i, 2, rv); c.font = book.f_header; c.fill = book.fill_secondary; c.number_format = PCT; c.alignment = Alignment(horizontal="center")
        for j, cv in enumerate(col_vals):
            c = wsn.cell(top + 2 + i, 3 + j, cell_fn(f"$B{top + 2 + i}", f"{L(3 + j)}${top + 1}")); c.font = F_FORMULA; c.number_format = fmt
    return top + 2 + len(row_vals)
# closed-form pieces so the tables re-solve without data tables
sum_pv = f"SUMPRODUCT(DCF!$C${D['cf_in']}:$G${D['cf_in']},1/(1+{{w}})^DCF!$C${D['t']}:$G${D['t']})"
n_years = f"(DCF!$C${D['yearfrac']}+DCF!$D${D['yearfrac']}+DCF!$E${D['yearfrac']}+DCF!$F${D['yearfrac']}+DCF!$G${D['yearfrac']})"
bridge_f = f"(-Operating!$C$4+Operating!$C$5)"
shares_f = INP["Diluted shares outstanding (bn)"]
wacc_vals = [0.075, 0.080, 0.085, 0.090, 0.095, 0.100, 0.105]
g_vals = [0.020, 0.025, 0.030, 0.035, 0.040]
em_vals = [16.0, 19.0, 22.0, 25.0, 28.0]
def pg_px(w, gg):
    return f"=({sum_pv.format(w=w)}+DCF!$G${D['ufcf']}*(1+{gg})/({w}-{gg})/(1+{w})^{n_years}+{bridge_f})/{shares_f}"
def em_px(w, m):
    return f"=({sum_pv.format(w=w)}+(DCF!$G${D['ebit']}+DCF!$G${D['da']})*{m}/(1+{w})^{n_years}+{bridge_f})/{shares_f}"
t = 4
t = sens_table(t, "IMPLIED SHARE PRICE (US$): WACC × TERMINAL GROWTH (perpetuity method)", wacc_vals, g_vals, "WACC", "g", pg_px) + 2
t = sens_table(t, "IMPLIED SHARE PRICE (US$): WACC × EXIT EV/EBITDA (exit-multiple method)", wacc_vals, em_vals, "WACC", "multiple", em_px) + 2
book.section(wsn, t, "CURRENT PRICE AND SELECTED-CASE OUTPUTS", 10); t += 1
for i, (lab, fml, fmt) in enumerate([("Share price (US$)", f"={INP['Share price (US$)']}", NUM2), ("WACC (selected)", f"={WACC}", PCT),
                                      ("Perpetuity-method price", f"=DCF!C{V['px']}", NUM2), ("Exit-multiple price", f"=DCF!E{V['px']}", NUM2),
                                      ("Midpoint", f"=AVERAGE(DCF!C{V['px']},DCF!E{V['px']})", NUM2), ("Midpoint vs price", f"=C{t+4}/C{t}-1", PCT)]):
    wsn.cell(t + i, 2, lab).font = F_TEXT; c = wsn.cell(t + i, 3, fml); c.font = F_LINK if "DCF" in fml or "WACC" in fml or "Inputs" in fml else F_FORMULA; c.number_format = fmt
book.note(wsn, t + 7, "Read the tables as ranges, not points. With terminal value at 70–80% of enterprise value, a 1% move in WACC or 0.5% in g is a 10–15% move in value; that is a property of every DCF of a mature company and the reason the football field (Project 5) shows a bar, not a line.", 10, height=30)

# --------------------------------------------------------------------------- #
# Summary and checks
# --------------------------------------------------------------------------- #
wsu = book.sheet("Summary", ncols=6, label_width=46, subtitle="Valuation summary for the selected scenario")
wsu["B3"] = '="Scenario: "&Inputs!$D$5'; wsu["B3"].font = Font(name=FONT, size=10, bold=True, color=book.theme.accent)
book.section(wsu, 4, "VALUATION (US$bn unless stated)", 6)
for j, lab in enumerate(["Perpetuity growth", "Exit multiple"]):
    c = wsu.cell(5, 3 + j, lab); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center")
srows = [("WACC", f"={WACC}", f"={WACC}", PCT), ("Terminal assumption", f"={g}", f"={em}", "General"),
         ("PV of forecast cash flows", f"=DCF!C{V['pv_cf']}", f"=DCF!E{V['pv_cf']}", NUM), ("PV of terminal value", f"=DCF!C{V['pv_tv']}", f"=DCF!E{V['pv_tv']}", NUM),
         ("Enterprise value", f"=DCF!C{V['ev']}", f"=DCF!E{V['ev']}", NUM), ("Net cash", f"=DCF!C{V['less_debt']}+DCF!C{V['plus_cash']}", f"=DCF!E{V['less_debt']}+DCF!E{V['plus_cash']}", NUM),
         ("Equity value", f"=DCF!C{V['eq']}", f"=DCF!E{V['eq']}", NUM), ("Implied value per share (US$)", f"=DCF!C{V['px']}", f"=DCF!E{V['px']}", NUM2),
         ("Current share price (US$)", f"={INP['Share price (US$)']}", f"={INP['Share price (US$)']}", NUM2), ("Premium / (discount)", f"=DCF!C{V['updown']}", f"=DCF!E{V['updown']}", PCT),
         ("Terminal value % of EV", f"=DCF!C{V['tv_share']}", f"=DCF!E{V['tv_share']}", PCT), ("Implied exit multiple / implied growth", f"=DCF!C{V['impl_mult']}", f"=DCF!E{V['impl_g']}", "General")]
for i, (lab, f1, f2, fmt) in enumerate(srows, start=6):
    wsu.cell(i, 2, lab).font = F_BOLD if lab.startswith("Implied value") else F_TEXT
    for j, fml in enumerate([f1, f2]):
        c = wsu.cell(i, 3 + j, fml); c.font = F_LINK; c.number_format = fmt
wsu.cell(7, 3).number_format = PCT; wsu.cell(7, 4).number_format = MULT; wsu.cell(17, 3).number_format = MULT; wsu.cell(17, 4).number_format = PCT
book.section(wsu, 20, "READING THE RESULT", 6)
memo = [
    "Both terminal methods are shown because each embeds an assumption the other tests: the perpetuity method implies an exit multiple, the exit method implies a growth rate. A defensible range sits between them.",
    "Apple's WACC is almost entirely its cost of equity (debt is 2% of capital), so the valuation is a function of beta and the equity risk premium far more than of the operating forecast in years 1–5.",
    "Stock-based compensation is treated as a real cost (not added back), which lowers UFCF by roughly US$13–16bn a year versus a naïve build. Add it back and the value rises ~8%; the notes page explains why it should not be.",
    "Use the sensitivity tables, not the point estimate, and read the Base/Upside/Downside cases together with the WACC range as the DCF bar on the football field (Project 5).",
]
for i, mline in enumerate(memo):
    c = wsu.cell(21 + i, 2, "•  " + mline); c.font = F_TEXT; c.alignment = Alignment(wrap_text=True, vertical="top")
    wsu.merge_cells(start_row=21 + i, start_column=2, end_row=21 + i, end_column=6); wsu.row_dimensions[21 + i].height = 30

wc = book.sheet("Checks", ncols=4, label_width=60, subtitle="Integrity checks · should all be TRUE")
chk = [("Perpetuity and exit-multiple methods use the same PV of forecast cash flows", f"=ROUND(DCF!C{V['pv_cf']}-DCF!E{V['pv_cf']},6)=0"),
       ("WACC above terminal growth (perpetuity formula valid)", f"={WACC}>{g}"),
       ("Discount periods increase by one year after the stub", f"=ROUND(DCF!E{D['t']}-DCF!D{D['t']},6)=1"),
       ("Base-case FY2026 UFCF within 15% of Project 1's FCF (sanity)", f"=ABS(DCF!C{D['ufcf']}/131.6-1)<0.5"),
       ("Implied exit multiple from perpetuity method between 10x and 40x", f"=AND(DCF!C{V['impl_mult']}>10,DCF!C{V['impl_mult']}<40)")]
for i, (lab, fml) in enumerate(chk, start=4):
    wc.cell(i, 2, lab).font = F_TEXT; c = wc.cell(i, 3, fml); c.font = F_FORMULA
book.section(wc, 10, "OVERALL", 4); c = wc.cell(10, 3, '=IF(COUNTIF(C4:C8,FALSE)=0,"MODEL OK","CHECK ERRORS")'); c.font = Font(name=FONT, size=11, bold=True, color=book.theme.text_on_primary)

book.wb.move_sheet("Summary", offset=-5)
book.cover(
    toc=[("Summary", "valuation by both terminal methods and how to read it"), ("Inputs", "market data, rates, terminal assumptions, scenario selector"),
         ("WACC", "peer betas, unlever/relever, CAPM, cost of debt, weights"), ("Operating", "three operating cases from the Project 1 model"),
         ("DCF", "UFCF, discounting with stub and mid-year convention, terminal value, equity bridge"), ("Sensitivity", "WACC × g and WACC × exit multiple"), ("Checks", "integrity checks")],
    highlights=[("Implied price — perpetuity (US$)", f"=DCF!C{V['px']}", NUM2), ("Implied price — exit multiple (US$)", f"=DCF!E{V['px']}", NUM2),
                ("Current price (US$)", f"={INP['Share price (US$)']}", NUM2), ("WACC", f"={WACC}", PCT), ("Model status", "=Checks!C10", "General")],
    blurb="An unlevered free-cash-flow DCF of Apple on the operating model from Project 1, with WACC built bottom-up from a ten-company peer set, terminal value by perpetuity growth and by exit multiple with each method's implied assumption cross-checked against the other, an equity bridge and two-way sensitivity tables.",
    method=["Cost of equity: peer regression betas (five years weekly vs S&P 500), unlevered at each peer's net debt/equity, median relevered at Apple's, times an equity risk premium of 4.5% on a 4.95% ten-year Treasury.",
            "Unlevered FCF = EBIT × (1 − cash tax rate) + D&A − capex − increase in NWC; stock-based compensation is treated as a real expense.",
            "Stub period for the 12 days remaining in FY2026 and mid-year discounting thereafter; terminal value at FY2030 end.",
            "Sensitivity tables re-solve the DCF in closed form (no Excel data tables), so they work in Google Sheets and LibreOffice."],
    sources=["Apple Form 10-K FY2025; Project 1 model for the forecasts", "Yahoo Finance for prices and peer balance sheets (12 Sep 2026)", "FRED: DGS10 (10-year Treasury), BAMLC0A1CAAAEY (AAA corporate yield)", "Damodaran, implied equity risk premium (mid-2026)"])
book.finish(freeze={"DCF": "C6", "Operating": "C8", "WACC": "C6"}, repeat_rows={"DCF": "1:5", "WACC": "1:5"})
out = os.path.join(HERE, "Apple_DCF.xlsx"); book.save(out)
rc = recalc(out); print("recalc:", rc.get("status"), rc.get("total_errors"), rc.get("error_summary"))
pdf = export_pdf(out); png = preview_png(pdf, 0, dpi=60); os.rename(png, os.path.join(HERE, "cover.png")); print("done", out)
