"""
Build 05-tcs-football-field/TCS_Football_Field.xlsx — valuation summary for Tata Consultancy Services.

Sheets: Cover · Football Field (the chart and the range table) · Ranges (every bar, its source and its caveat) ·
DCF (compact INR DCF with terminal value both ways and a live WACC x g sensitivity) · WACC (beta regressions,
including the ones that were run and rejected) · Inputs · Checks

All source data is embedded in this script with its provenance, so the workbook rebuilds from the repository
alone with no scratch files. Cash flows are in INR because TCS reports in INR; discounting local-currency
cash flows at a local-currency WACC avoids putting an FX forecast inside a valuation.
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"

# --------------------------------------------------------------------------- #
# Source data (embedded with provenance)
# --------------------------------------------------------------------------- #
# TCS annual financials, INR billions, fiscal year to 31 March.
# Source: Yahoo Finance fundamentals timeseries for TCS.NS, which restates the consolidated
# annual report. Cross-checked against the LTM figures used in Project 3.
HIST = [  # (label, revenue, EBITDA, EBIT, net income, D&A, capex, tax rate)
    ("FY2023A", 2254.58, 627.08, 543.77, 421.47, 50.22, 31.00, 0.2570),
    ("FY2024A", 2408.93, 677.60, 594.25, 459.08, 49.85, 26.74, 0.2564),
    ("FY2025A", 2553.24, 713.69, 622.93, 485.53, 52.42, 39.37, 0.2530),
    ("FY2026A", 2670.21, 722.74, 670.22, 492.10, 55.60, 41.46, 0.2460),
]
FCST = ["FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
GROWTH = [0.030, 0.040, 0.050, 0.055, 0.060]
MARGIN = [0.248, 0.248, 0.250, 0.250, 0.250]

# Market data, 12 September 2026 (Yahoo Finance), consistent with Projects 3 and 4.
PRICE = 2200.8; W52_HIGH = 3350.0; W52_LOW = 1976.8; SHARES = 3618.087518; USDINR = 95.54
NET_CASH_USD = 3019.93884          # Project 3: net debt US$(3,020)m, i.e. net cash

# Beta regressions: 5 years of weekly total returns to 12 Sep 2026, 254 observations.
BETAS_LOCAL = [("TCS vs NIFTY 50 (INR)", 0.907, 0.28, "Accepted: same market, same currency, same trading hours")]
BETAS_REJECTED = [
    ("TCS vs S&P 500 (USD returns)", 0.140, 0.01, "Rejected: R-squared 0.01 — no explanatory power"),
    ("Infosys vs S&P 500 (USD returns)", 0.240, 0.02, "Rejected: R-squared 0.02"),
    ("HCLTech vs S&P 500 (USD returns)", 0.062, 0.00, "Rejected: R-squared 0.00"),
    ("Wipro vs S&P 500 (USD returns)", 0.326, 0.04, "Rejected: R-squared 0.04"),
    ("Accenture vs S&P 500", 0.975, 0.24, "US-listed peer: regression behaves normally"),
    ("Cognizant vs S&P 500", 0.908, 0.23, "US-listed peer: regression behaves normally"),
    ("EPAM vs S&P 500", 1.485, 0.16, "US-listed peer: regression behaves normally"),
]

# Valuation ranges carried in from earlier projects (per share, INR).
# Project 3 Summary!G:I — implied price at the peer 25th percentile and 75th percentile.
P3 = {"EV/LTM EBITDA": (1652.20, 2478.49), "EV/CY2027 revenue": (1242.05, 1745.32),
      "P/E (CY2027, calendarised)": (1874.11, 2688.94), "P/E (LTM)": (1827.03, 2942.40)}
# Analyst consensus, 12-month price target, 43 analysts (Investing.com, 12 Sep 2026).
ANALYST = (1775.0, 2943.65, 3900.0, 43)
# Project 4 Summary!F:H — illustrative, precedent 25th to 75th percentile.
P4 = {"EV/LTM revenue": (1428.41, 2682.41), "EV/LTM EBITDA": (2862.26, 3532.32)}

book = Book(theme=THEMES["tcs"], project_no=5, project="Football Field and Valuation Summary",
            company="Tata Consultancy Services", units="INR billions · per-share values in INR",
            as_of="market data 12 Sep 2026 · financials FY2026 (year to 31 Mar 2026)")

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
wi = book.sheet("Inputs", ncols=6, label_width=48, subtitle="Market data, capital structure, cost of capital and forecast drivers")
wi.column_dimensions["F"].width = 66
INP = {}


def inp_block(ws, start, title, rows, ncols=6):
    book.section(ws, start, title, ncols)
    r = start + 1
    for label, value, fmt, note in rows:
        ws.cell(r, 2, label).font = F_TEXT
        c = ws.cell(r, 3, value)
        c.font = F_INPUT if not (isinstance(value, str) and value.startswith("=")) else F_FORMULA
        c.number_format = fmt
        n = ws.cell(r, 6, note); n.font = F_NOTE; n.alignment = Alignment(wrap_text=True, vertical="top")
        INP[label] = f"Inputs!$C${r}"
        r += 1
    return r + 1


r = inp_block(wi, 4, "MARKET DATA (12 SEPTEMBER 2026)", [
    ("TCS share price (INR)", PRICE, NUM2, "Yahoo Finance, NSE close 12 Sep 2026. Same price as Projects 3 and 4"),
    ("52-week high (INR)", W52_HIGH, NUM2, "Yahoo Finance"),
    ("52-week low (INR)", W52_LOW, NUM2, "Yahoo Finance"),
    ("Shares outstanding (m)", SHARES, NUM0, "Project 3"),
    ("USD per INR", round(1 / USDINR, 6), '0.000000', f"USDINR {USDINR:.2f}; used only to restate Project 3's US$ net cash"),
    ("Net cash (US$m, from Project 3)", NET_CASH_USD, NUM0, "Project 3 reports net debt of US$(3,020)m"),
    ("Net cash (INR bn)", f"=C10*{USDINR}/1000", NUM, "Restated at the 12 Sep 2026 rate"),
])
R_NETCASH = 11

r = inp_block(wi, r, "COST OF CAPITAL (INR, LOCAL CAPM)", [
    ("Risk-free rate: India 10-year G-Sec", 0.0702, PCT, "7.02% on 11 Sep 2026. Local-currency cash flows require a local-currency risk-free rate"),
    ("Raw beta (TCS vs NIFTY 50, 5y weekly)", 0.907, '0.000', "254 weekly observations to 12 Sep 2026; R-squared 0.28. See the WACC sheet"),
    ("Blume-adjusted beta", "=0.67*C15+0.33", '0.000', "0.67 x raw + 0.33, as banks apply it; pulls the estimate toward 1.0"),
    ("Mature-market equity risk premium", 0.045, PCT, "Same 4.5% used in Project 2, for consistency across the portfolio"),
    ("India country risk premium", 0.022, PCT, "Damodaran country risk premium for India (Baa3/BBB-). Verify against the current table before reuse"),
    ("India equity risk premium", "=C17+C18", PCT, "Mature-market premium plus country risk"),
    ("Cost of equity", "=C14+C16*C19", PCT, "CAPM in INR"),
    ("Debt as % of capital", 0.0, PCT, "TCS carries no meaningful borrowing and is net cash, so WACC collapses to the cost of equity"),
    ("WACC (INR)", "=C20*(1-C21)", PCT, "Used to discount INR cash flows"),
])
R_KE, R_WACC = 20, 22

r = inp_block(wi, r, "FORECAST DRIVERS", [
    ("Tax rate on EBIT", 0.255, PCT, "Between the FY2025 and FY2026 effective rates"),
    ("D&A as % of revenue", 0.021, PCT, "FY2026 actual: 2.08%"),
    ("Capex as % of revenue", 0.016, PCT, "FY2026 actual: 1.55%. An asset-light business"),
    ("Increase in working capital as % of incremental revenue", 0.150, PCT, "Receivables-driven; a growth business absorbs cash here"),
    ("Terminal growth (INR, nominal)", 0.055, PCT, "Roughly Indian long-run inflation plus modest real growth; well below nominal GDP"),
    ("Exit EV/EBITDA (cross-check only)", 11.0, MULT, "Used to test the perpetuity assumption, not to value the company"),
    ("Mid-year discounting convention", 1, '0', "1 = cash flows arrive mid-year (periods 0.5, 1.5, ...); 0 = year-end"),
])
R_TAX, R_DA, R_CAPEX, R_NWC, R_G, R_EXIT, R_MID = 25, 26, 27, 28, 29, 30, 31

r = inp_block(wi, r, "ANALYST CONSENSUS (RESEARCH-TARGET BAR)", [
    ("Analyst target — low (INR)", ANALYST[0], NUM2, "Lowest of 43 published 12-month targets, 12 Sep 2026"),
    ("Analyst target — average (INR)", ANALYST[1], NUM2, "Consensus average; the figure most often quoted back in a meeting"),
    ("Analyst target — high (INR)", ANALYST[2], NUM2, "Highest published target"),
    ("Number of analysts covering", ANALYST[3], '0', "Coverage depth. A thin consensus is worth less than a broad one"),
])

book.note(wi, r, "Everything above is a hard-coded input (blue). Every figure on every other sheet is a formula. "
                 "The DCF is deliberately compact: five forecast years, one case, and a sensitivity grid — its job is to "
                 "produce one honest bar on the football field, not to replace Project 2's full treatment.", 6, height=40)

# --------------------------------------------------------------------------- #
# WACC — the regressions, including the rejected ones
# --------------------------------------------------------------------------- #
ww = book.sheet("WACC", ncols=7, label_width=44, col_width=14,
                subtitle="Beta estimated where the regression is valid, and a record of the regressions that were rejected")
ww.column_dimensions["G"].width = 60
book.section(ww, 4, "BETA ESTIMATE USED", 7)
book.year_header(ww, 5, ["Raw beta", "R-squared", "", ""], first_col=3, label="Regression")
rr = 6
for label, beta, r2, note in BETAS_LOCAL:
    ww.cell(rr, 2, label).font = F_BOLD
    c = ww.cell(rr, 3, beta); c.font = F_INPUT; c.number_format = '0.000'
    c = ww.cell(rr, 4, r2); c.font = F_INPUT; c.number_format = '0.00'
    n = ww.cell(rr, 7, note); n.font = F_NOTE; n.alignment = Alignment(wrap_text=True, vertical="top")
    rr += 1

rr += 1
book.section(ww, rr, "REGRESSIONS RUN AND REJECTED", 7); rr += 1
book.year_header(ww, rr, ["Raw beta", "R-squared", "", ""], first_col=3, label="Regression"); rr += 1
for label, beta, r2, note in BETAS_REJECTED:
    ww.cell(rr, 2, label).font = F_TEXT
    c = ww.cell(rr, 3, beta); c.font = F_INPUT; c.number_format = '0.000'
    c = ww.cell(rr, 4, r2); c.font = F_INPUT; c.number_format = '0.00'
    n = ww.cell(rr, 7, note); n.font = F_NOTE; n.alignment = Alignment(wrap_text=True, vertical="top")
    rr += 1

rr += 1
book.note(ww, rr, "Project 2 built Apple's beta by regressing ten peers against the S&P 500, unlevering, taking the median and "
                  "relevering. The same recipe applied to TCS breaks: Indian shares regressed against the S&P 500 give betas of "
                  "0.06 to 0.33 with R-squared of 0.00 to 0.04, because the two markets trade in different hours and different "
                  "currencies. A beta with no explanatory power is not a conservative beta, it is a wrong one. The regression is "
                  "therefore run in the market where TCS actually trades, against the NIFTY 50, in INR — which is also why the "
                  "whole DCF is built in INR.", 7, height=64); rr += 2

book.section(ww, rr, "COST OF CAPITAL BUILD", 7); rr += 1
wacc_rows = [("Risk-free rate (India 10-year G-Sec)", f"={INP['Risk-free rate: India 10-year G-Sec']}", PCT),
             ("Raw beta", f"={INP['Raw beta (TCS vs NIFTY 50, 5y weekly)']}", '0.000'),
             ("Blume-adjusted beta", f"={INP['Blume-adjusted beta']}", '0.000'),
             ("India equity risk premium", f"={INP['India equity risk premium']}", PCT),
             ("Cost of equity", f"={INP['Cost of equity']}", PCT),
             ("WACC (INR)", f"={INP['WACC (INR)']}", PCT)]
for label, fml, fmt in wacc_rows:
    ww.cell(rr, 2, label).font = F_BOLD if "WACC" in label or "Cost of equity" in label else F_TEXT
    c = ww.cell(rr, 3, fml); c.font = F_LINK; c.number_format = fmt
    if "WACC" in label: c.border = DOUBLE_BORDER
    rr += 1
rr += 1
book.note(ww, rr, "TCS is net cash, so there is no debt tranche to weight and WACC equals the cost of equity. Stating that "
                  "explicitly is better than building a capital-structure table whose debt weight is zero.", 7, height=28)

# --------------------------------------------------------------------------- #
# DCF
# --------------------------------------------------------------------------- #
NH, NF = len(HIST), len(FCST)
C0 = 3                      # first historical column
CF0 = C0 + NH               # first forecast column
CLAST = CF0 + NF - 1
NCOL = CLAST + 1
wd = book.sheet("DCF", ncols=NCOL, label_width=44, col_width=12.5,
                subtitle="INR billions · fiscal year to 31 March · actuals in blue, forecast driven off the Inputs sheet")

book.section(wd, 4, "UNLEVERED FREE CASH FLOW", NCOL)
book.year_header(wd, 5, [h[0] for h in HIST] + FCST, first_col=C0)

ROW = {}
rr = 6


def line(label, bold=False, border=None):
    global rr
    c = wd.cell(rr, 2, label); c.font = F_BOLD if bold else F_TEXT
    ROW[label] = rr
    if border:
        for j in range(C0, NCOL + 1): wd.cell(rr, j).border = border
    rr += 1
    return ROW[label]


# Revenue
r_rev = line("Revenue", bold=True)
for i, h in enumerate(HIST):
    c = wd.cell(r_rev, C0 + i, h[1]); c.font = F_INPUT; c.number_format = NUM
for i in range(NF):
    col = CF0 + i
    c = wd.cell(r_rev, col, f"={L(col-1)}{r_rev}*(1+{L(col)}{r_rev+1})"); c.font = F_FORMULA; c.number_format = NUM

r_g = line("   Growth")
for i in range(1, NH):
    c = wd.cell(r_g, C0 + i, f"={L(C0+i)}{r_rev}/{L(C0+i-1)}{r_rev}-1"); c.font = F_FORMULA; c.number_format = PCT
for i in range(NF):
    c = wd.cell(r_g, CF0 + i, GROWTH[i]); c.font = F_INPUT; c.number_format = PCT

r_ebit = line("EBIT", bold=True)
for i, h in enumerate(HIST):
    c = wd.cell(r_ebit, C0 + i, h[3]); c.font = F_INPUT; c.number_format = NUM
for i in range(NF):
    col = CF0 + i
    c = wd.cell(r_ebit, col, f"={L(col)}{r_rev}*{L(col)}{r_ebit+1}"); c.font = F_FORMULA; c.number_format = NUM

r_m = line("   EBIT margin")
for i in range(NH):
    c = wd.cell(r_m, C0 + i, f"={L(C0+i)}{r_ebit}/{L(C0+i)}{r_rev}"); c.font = F_FORMULA; c.number_format = PCT
for i in range(NF):
    c = wd.cell(r_m, CF0 + i, MARGIN[i]); c.font = F_INPUT; c.number_format = PCT

r_tax = line("Less: tax on EBIT")
for i in range(NH):
    c = wd.cell(r_tax, C0 + i, f"=-{L(C0+i)}{r_ebit}*{HIST[i][7]}"); c.font = F_FORMULA; c.number_format = NUM
for i in range(NF):
    col = CF0 + i
    c = wd.cell(r_tax, col, f"=-{L(col)}{r_ebit}*{INP['Tax rate on EBIT']}"); c.font = F_FORMULA; c.number_format = NUM

r_nopat = line("NOPAT", bold=True, border=TOTAL_BORDER)
for i in range(NH + NF):
    col = C0 + i
    c = wd.cell(r_nopat, col, f"={L(col)}{r_ebit}+{L(col)}{r_tax}"); c.font = F_FORMULA; c.number_format = NUM; c.border = TOTAL_BORDER

r_da = line("Plus: depreciation and amortisation")
for i, h in enumerate(HIST):
    c = wd.cell(r_da, C0 + i, h[5]); c.font = F_INPUT; c.number_format = NUM
for i in range(NF):
    col = CF0 + i
    c = wd.cell(r_da, col, f"={L(col)}{r_rev}*{INP['D&A as % of revenue']}"); c.font = F_FORMULA; c.number_format = NUM

r_capex = line("Less: capital expenditure")
for i, h in enumerate(HIST):
    c = wd.cell(r_capex, C0 + i, -h[6]); c.font = F_INPUT; c.number_format = NUM
for i in range(NF):
    col = CF0 + i
    c = wd.cell(r_capex, col, f"=-{L(col)}{r_rev}*{INP['Capex as % of revenue']}"); c.font = F_FORMULA; c.number_format = NUM

r_nwc = line("Less: increase in working capital")
for i in range(1, NH):
    col = C0 + i
    c = wd.cell(r_nwc, col, f"=-({L(col)}{r_rev}-{L(col-1)}{r_rev})*{INP['Increase in working capital as % of incremental revenue']}")
    c.font = F_FORMULA; c.number_format = NUM
for i in range(NF):
    col = CF0 + i
    c = wd.cell(r_nwc, col, f"=-({L(col)}{r_rev}-{L(col-1)}{r_rev})*{INP['Increase in working capital as % of incremental revenue']}")
    c.font = F_FORMULA; c.number_format = NUM

r_fcf = line("Unlevered free cash flow", bold=True, border=DOUBLE_BORDER)
for i in range(NH + NF):
    col = C0 + i
    c = wd.cell(r_fcf, col, f"={L(col)}{r_nopat}+{L(col)}{r_da}+{L(col)}{r_capex}+{L(col)}{r_nwc}")
    c.font = F_FORMULA; c.number_format = NUM; c.border = DOUBLE_BORDER

r_ebitda = line("Memo: EBITDA")
for i, h in enumerate(HIST):
    c = wd.cell(r_ebitda, C0 + i, h[2]); c.font = F_INPUT; c.number_format = NUM
for i in range(NF):
    col = CF0 + i
    c = wd.cell(r_ebitda, col, f"={L(col)}{r_ebit}+{L(col)}{r_da}"); c.font = F_FORMULA; c.number_format = NUM

rr += 1
book.section(wd, rr, "DISCOUNTING", NCOL); rr += 1
r_per = line("Discount period (years)")
for i in range(NF):
    col = CF0 + i
    c = wd.cell(r_per, col, f"={i+1}-{INP['Mid-year discounting convention']}*0.5"); c.font = F_FORMULA; c.number_format = '0.0'
r_df = line("Discount factor")
for i in range(NF):
    col = CF0 + i
    c = wd.cell(r_df, col, f"=1/(1+{INP['WACC (INR)']})^{L(col)}{r_per}"); c.font = F_FORMULA; c.number_format = '0.000'
r_pv = line("PV of free cash flow")
for i in range(NF):
    col = CF0 + i
    c = wd.cell(r_pv, col, f"={L(col)}{r_fcf}*{L(col)}{r_df}"); c.font = F_FORMULA; c.number_format = NUM

rr += 1
book.section(wd, rr, "TERMINAL VALUE — TWO METHODS, EACH CHECKING THE OTHER", NCOL); rr += 1


def kv(label, fml, fmt, bold=False, note=""):
    global rr
    c = wd.cell(rr, 2, label); c.font = F_BOLD if bold else F_TEXT
    v = wd.cell(rr, C0 + NH + NF - 1, fml); v.font = F_FORMULA; v.number_format = fmt
    ROW[label] = rr
    if note:
        n = wd.cell(rr, NCOL, note); n.font = F_NOTE
    rr += 1
    return ROW[label]


CL = L(CLAST)
r_tvg = kv("Terminal value at FY2031 (perpetuity growth)",
           f"={CL}{r_fcf}*(1+{INP['Terminal growth (INR, nominal)']})/({INP['WACC (INR)']}-{INP['Terminal growth (INR, nominal)']})", NUM, bold=True)
r_pvtvg = kv("PV of terminal value (perpetuity growth)", f"={CL}{r_tvg}*{CL}{r_df}", NUM)
r_tvx = kv("Terminal value at FY2031 (exit multiple)", f"={CL}{r_ebitda}*{INP['Exit EV/EBITDA (cross-check only)']}", NUM)
r_pvtvx = kv("PV of terminal value (exit multiple)", f"={CL}{r_tvx}*{CL}{r_df}", NUM)
r_impg = kv("Growth implied by the exit multiple",
            f"=({INP['WACC (INR)']}*{CL}{r_tvx}-{CL}{r_fcf})/({CL}{r_tvx}+{CL}{r_fcf})", PCT,
            note="If this is far from the assumed terminal growth, one of the two assumptions is wrong")
r_impx = kv("Exit multiple implied by terminal growth", f"={CL}{r_tvg}/{CL}{r_ebitda}", MULT)

rr += 1
book.section(wd, rr, "VALUATION BRIDGE (PERPETUITY-GROWTH METHOD)", NCOL); rr += 1
r_sumpv = kv("Sum of PV of forecast free cash flow", f"=SUM({L(CF0)}{r_pv}:{CL}{r_pv})", NUM)
r_pvtv = kv("PV of terminal value", f"={CL}{r_pvtvg}", NUM)
r_ev = kv("Enterprise value", f"={CL}{r_sumpv}+{CL}{r_pvtv}", NUM, bold=True)
r_tvpct = kv("Terminal value as % of enterprise value", f"={CL}{r_pvtv}/{CL}{r_ev}", PCT,
             note="Above roughly 80% the answer is an assumption about 2031, not a forecast")
r_nc = kv("Plus: net cash", f"={INP['Net cash (INR bn)']}", NUM)
r_eq = kv("Equity value", f"={CL}{r_ev}+{CL}{r_nc}", NUM, bold=True)
r_sh = kv("Shares outstanding (m)", f"={INP['Shares outstanding (m)']}", NUM0)
r_vps = kv("Value per share (INR)", f"={CL}{r_eq}*1000/{CL}{r_sh}", NUM2, bold=True)
wd.cell(r_vps, CLAST).border = DOUBLE_BORDER
r_px = kv("Current share price (INR)", f"={INP['TCS share price (INR)']}", NUM2)
r_prem = kv("Market price vs DCF value", f"={CL}{r_px}/{CL}{r_vps}-1", PCT,
            note="Positive = the market pays more than this DCF supports")

rr += 1
book.section(wd, rr, "SENSITIVITY: VALUE PER SHARE (INR)", NCOL); rr += 1
r_sens_hdr = rr
wd.cell(rr, 2, "Terminal growth  \\  WACC").font = F_BOLD
WACC_OFF = [-0.010, -0.005, 0.0, 0.005, 0.010]
G_OFF = [-0.010, -0.005, 0.0, 0.005, 0.010]
for j, off in enumerate(WACC_OFF):
    c = wd.cell(rr, C0 + j, f"={INP['WACC (INR)']}+{off}")
    c.font = book.f_header; c.fill = book.fill_secondary; c.number_format = PCT
    c.alignment = Alignment(horizontal="center")
rr += 1
r_sens0 = rr
FCF_RNG = f"${L(CF0)}${r_fcf}:${CL}${r_fcf}"
PER_RNG = f"${L(CF0)}${r_per}:${CL}${r_per}"
for i, goff in enumerate(G_OFF):
    c = wd.cell(rr, 2, f"={INP['Terminal growth (INR, nominal)']}+{goff}")
    c.font = F_BOLD; c.number_format = PCT
    for j, woff in enumerate(WACC_OFF):
        w = f"{L(C0+j)}${r_sens_hdr}"
        g = f"$B{rr}"
        fml = (f"=(SUMPRODUCT({FCF_RNG},1/(1+{w})^{PER_RNG})"
               f"+${CL}${r_fcf}*(1+{g})/({w}-{g})*1/(1+{w})^${CL}${r_per}"
               f"+{INP['Net cash (INR bn)']})*1000/{INP['Shares outstanding (m)']}")
        cc = wd.cell(rr, C0 + j, fml); cc.font = F_FORMULA; cc.number_format = NUM0
        if goff == 0.0 and woff == 0.0:
            cc.fill = book.fill_accent; cc.font = Font(name=FONT, size=10, bold=True)
    rr += 1
r_sens1 = rr - 1
rr += 1
book.note(wd, rr, "The shaded cell is the base case and reconciles to the bridge above. The grid is live: every cell "
                  "rediscounts the same five forecast cash flows at its own WACC and regrows the terminal value at its own rate. "
                  "The spread across the grid — not the single base-case number — is what becomes the DCF bar on the football field.",
          NCOL, height=40)

# --------------------------------------------------------------------------- #
# Ranges
# --------------------------------------------------------------------------- #
wr = book.sheet("Ranges", ncols=9, label_width=40, col_width=13,
                subtitle="Every bar, its source, what it is worth trusting, and the weight it carries in the recommended range")
wr.column_dimensions["I"].width = 54
book.year_header(wr, 4, ["Low (INR)", "High (INR)", "Midpoint", "Width", "Weight", "Source"], first_col=3, label="Methodology")

# (name, low, high, weight, source, note)
BARS = [
    ("52-week trading range", f"={INP['52-week low (INR)']}", f"={INP['52-week high (INR)']}", 0.00, "Market",
     "Where the shares have actually changed hands. Context, not a valuation, so it carries no weight."),
    ("Analyst price targets (43)", f"={INP['Analyst target — low (INR)']}", f"={INP['Analyst target — high (INR)']}", 0.10, "Consensus",
     "Published 12-month targets. Weighted lightly: the analysts are reading the same filings as everyone else, and targets lag the price."),
    ("DCF — WACC and growth sensitivity", None, None, 0.35, "Project 5",
     "Low and high are the corners of the sensitivity grid: WACC plus or minus 100bp against terminal growth plus or minus 100bp."),
    ("Trading comps — EV/CY2027 revenue", P3["EV/CY2027 revenue"][0], P3["EV/CY2027 revenue"][1], 0.15, "Project 3",
     "Peer 25th to 75th percentile, calendarised to TCS's March year. The forward multiple institutional investors quote."),
    ("Trading comps — EV/LTM EBITDA", P3["EV/LTM EBITDA"][0], P3["EV/LTM EBITDA"][1], 0.15, "Project 3",
     "Peer 25th to 75th percentile. No estimates involved, so no forecast risk."),
    ("Trading comps — P/E (CY2027)", P3["P/E (CY2027, calendarised)"][0], P3["P/E (CY2027, calendarised)"][1], 0.15, "Project 3",
     "Equity multiple; consistent with TCS's net-cash balance sheet, which flatters EV multiples."),
    ("Precedents — EV/LTM revenue", P4["EV/LTM revenue"][0], P4["EV/LTM revenue"][1], 0.05, "Project 4",
     "Illustrative. Fourteen deals, none remotely TCS's size, struck when the sector traded far higher. Weighted down accordingly."),
    ("Precedents — EV/LTM EBITDA", P4["EV/LTM EBITDA"][0], P4["EV/LTM EBITDA"][1], 0.05, "Project 4",
     "Illustrative and the weakest bar on the page: five disclosed observations, all mid-cap, all pre-2023."),
]
rr = 5
RBAR = {}
for name, lo, hi, wt, src, note in BARS:
    wr.cell(rr, 2, name).font = F_BOLD
    if name.startswith("DCF"):
        clo = wr.cell(rr, 3, f"=MIN(DCF!{L(C0)}{r_sens0}:{L(C0+4)}{r_sens1})")
        chi = wr.cell(rr, 4, f"=MAX(DCF!{L(C0)}{r_sens0}:{L(C0+4)}{r_sens1})")
        clo.font = chi.font = F_LINK
    else:
        clo = wr.cell(rr, 3, lo); chi = wr.cell(rr, 4, hi)
        clo.font = chi.font = F_LINK if isinstance(lo, str) else F_INPUT
    clo.number_format = chi.number_format = NUM0
    c = wr.cell(rr, 5, f"=AVERAGE(C{rr}:D{rr})"); c.font = F_FORMULA; c.number_format = NUM0
    c = wr.cell(rr, 6, f"=D{rr}-C{rr}"); c.font = F_FORMULA; c.number_format = NUM0
    c = wr.cell(rr, 7, wt); c.font = F_INPUT; c.number_format = PCT; c.fill = book.fill_accent
    wr.cell(rr, 8, src).font = F_TEXT
    n = wr.cell(rr, 9, note); n.font = F_NOTE; n.alignment = Alignment(wrap_text=True, vertical="top")
    wr.row_dimensions[rr].height = 30
    RBAR[name] = rr
    rr += 1

R_BAR0, R_BAR1 = 5, rr - 1
wr.cell(rr, 2, "Total weight").font = F_BOLD
c = wr.cell(rr, 7, f"=SUM(G{R_BAR0}:G{R_BAR1})"); c.font = F_FORMULA; c.number_format = PCT; c.border = TOTAL_BORDER
R_WSUM = rr
rr += 2

book.section(wr, rr, "RECOMMENDED RANGE (WEIGHTED)", 9); rr += 1
R_REC = rr
wr.cell(rr, 2, "Recommended value per share (INR)").font = F_BOLD
for j, col in enumerate([3, 4, 5]):
    src_col = {3: "C", 4: "D", 5: "E"}[col]
    c = wr.cell(rr, col, f"=SUMPRODUCT(${src_col}${R_BAR0}:${src_col}${R_BAR1},$G${R_BAR0}:$G${R_BAR1})/$G${R_WSUM}")
    c.font = Font(name=FONT, size=11, bold=True); c.number_format = NUM0
    c.fill = book.fill_light; c.border = DOUBLE_BORDER
n = wr.cell(rr, 9, "Low, high and midpoint, each weighted by the column G weights. Change a weight and this line moves: that is the "
                   "conversation with the client, and it is better had in the open than hidden inside a judgement.")
n.font = F_NOTE; n.alignment = Alignment(wrap_text=True, vertical="top"); wr.row_dimensions[rr].height = 32
rr += 2

book.section(wr, rr, "CONCLUSION", 9); rr += 1
R_CONCL0 = rr
concl = [("Current share price (INR)", f"={INP['TCS share price (INR)']}", NUM2),
         ("Low across all methodologies", f"=MIN(C{R_BAR0}:C{R_BAR1})", NUM0),
         ("High across all methodologies", f"=MAX(D{R_BAR0}:D{R_BAR1})", NUM0),
         ("Median of the methodology midpoints", f"=MEDIAN(E{R_BAR0}:E{R_BAR1})", NUM0),
         ("Recommended midpoint (weighted)", f"=E{R_REC}", NUM0),
         ("Market price vs the recommended midpoint", f"=C{R_CONCL0}/C{R_CONCL0+4}-1", PCT)]
for label, fml, fmt in concl:
    wr.cell(rr, 2, label).font = F_BOLD if "Recommended" in label or "Median" in label else F_TEXT
    c = wr.cell(rr, 3, fml); c.font = F_FORMULA; c.number_format = fmt
    rr += 1
R_PX, R_MIN, R_MAX, R_MED, R_RECMID = R_CONCL0, R_CONCL0 + 1, R_CONCL0 + 2, R_CONCL0 + 3, R_CONCL0 + 4

rr += 1
for txt in [
    "The bars do not agree, and a football field that did agree would be hiding something. The useful reading is which "
    "methods cluster: the DCF and the three trading-comps bars all land between roughly 1,200 and 2,700, and the share "
    "price sits in the upper half of that cluster.",
    "The weights are judgement and they are on the page as inputs rather than buried in a formula. Trading comps carry 45% "
    "between three multiples because they need no forecast; the DCF carries 35% because it is the only method that values "
    "TCS rather than its peer group; precedents carry 10% because of the size and cycle mismatch; analyst targets carry "
    "10% because they recycle public information; the 52-week range carries nothing because it is context.",
    "The two precedent bars sit above everything else. Part of that is a genuine control premium and part is the cycle — "
    "the deals were struck when IT services traded at 20-30x earnings and the sector has since de-rated to 10-15x. "
    "Project 4's notes flag this; it is repeated here because a reader looking only at this page would otherwise take the "
    "precedent bars at face value.",
    "The DCF and the trading comps are built on completely different logic — one discounts TCS's own cash flows at an INR "
    "cost of capital, the other asks what the market pays for comparable businesses today — and they overlap heavily. That "
    "agreement is the most informative thing on the page.",
]:
    book.note(wr, rr, "\u2022  " + txt, 9, height=15 * max(2, len(txt) // 105 + 1)); rr += 1

# --------------------------------------------------------------------------- #
# Football Field
# --------------------------------------------------------------------------- #
wf = book.sheet("Football Field", ncols=10, label_width=40, col_width=13,
                subtitle="Implied value per share (INR) by methodology \u00b7 bar = the range, not a point estimate")
book.section(wf, 4, "CHART DATA (linked from Ranges)", 10)
book.year_header(wf, 5, ["Base (invisible)", "Span", "Low", "High"], first_col=3, label="Methodology")
rr = 6
CHART_ROWS = [(name, RBAR[name]) for name, *_ in BARS] + [("Recommended range (weighted)", R_REC)]
for name, src in CHART_ROWS:
    c = wf.cell(rr, 2, name); c.font = F_BOLD if name.startswith("Recommended") else F_TEXT
    c = wf.cell(rr, 3, f"=Ranges!C{src}"); c.font = F_LINK; c.number_format = NUM0
    c = wf.cell(rr, 4, f"=Ranges!D{src}-Ranges!C{src}"); c.font = F_FORMULA; c.number_format = NUM0
    c = wf.cell(rr, 5, f"=Ranges!C{src}"); c.font = F_LINK; c.number_format = NUM0
    c = wf.cell(rr, 6, f"=Ranges!D{src}"); c.font = F_LINK; c.number_format = NUM0
    rr += 1
R_CH0, R_CH1 = 6, rr - 1

chart = BarChart()
chart.type = "bar"
chart.grouping = "stacked"
chart.overlap = 100
chart.title = "TCS \u2014 implied value per share (INR)"
chart.y_axis.title = "INR per share"
chart.height = 12; chart.width = 26
data = Reference(wf, min_col=3, max_col=4, min_row=5, max_row=R_CH1)
cats = Reference(wf, min_col=2, min_row=R_CH0, max_row=R_CH1)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.series[0].graphicalProperties.noFill = True
chart.series[0].graphicalProperties.line.noFill = True
chart.series[1].graphicalProperties.solidFill = book.theme.primary
chart.series[1].graphicalProperties.line.solidFill = book.theme.primary
from openpyxl.chart.marker import DataPoint
dp = DataPoint(idx=len(CHART_ROWS) - 1)
dp.graphicalProperties.solidFill = book.theme.accent
dp.graphicalProperties.line.solidFill = book.theme.accent
chart.series[1].data_points = [dp]
chart.legend = None
wf.add_chart(chart, f"B{rr + 2}")

c = wf.cell(rr + 1, 2, f"Current share price: INR {PRICE:,.2f} (12 Sep 2026)")
c.font = Font(name=FONT, size=11, bold=True, color=book.theme.accent)
book.note(wf, rr + 28, "The first series is deliberately unfilled: each bar floats from its low to its high, which is what makes "
                       "this a football field rather than a stacked bar chart. The final bar, in the accent colour, is the "
                       "weighted recommended range. Read the overlap between bars, not any single bar.",
          10, height=30)

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
wk = book.sheet("Checks", ncols=5, label_width=68, col_width=16, subtitle="Every check must read TRUE")
book.section(wk, 4, "INTEGRITY CHECKS", 5)
checks = [
    ("WACC exceeds terminal growth", f"={INP['WACC (INR)']}>{INP['Terminal growth (INR, nominal)']}"),
    ("Enterprise value ties to the sum of its parts", f"=ABS(DCF!{CL}{r_ev}-(DCF!{CL}{r_sumpv}+DCF!{CL}{r_pvtv}))<0.01"),
    ("Equity value ties to EV plus net cash", f"=ABS(DCF!{CL}{r_eq}-(DCF!{CL}{r_ev}+DCF!{CL}{r_nc}))<0.01"),
    ("DCF value per share is positive", f"=DCF!{CL}{r_vps}>0"),
    ("Terminal value is between 50% and 90% of EV", f"=AND(DCF!{CL}{r_tvpct}>0.5,DCF!{CL}{r_tvpct}<0.9)"),
    ("The two terminal-value methods are within 35% of each other",
     f"=ABS(DCF!{CL}{r_tvg}/DCF!{CL}{r_tvx}-1)<0.35"),
    ("Base-case sensitivity cell reconciles to the bridge",
     f"=ABS(DCF!{L(C0+2)}{r_sens0+2}-DCF!{CL}{r_vps})<1"),
    ("Every forecast free cash flow is positive", f"=MIN(DCF!{L(CF0)}{r_fcf}:{CL}{r_fcf})>0"),
    ("Forecast revenue growth never exceeds the historical maximum",
     f"=MAX(DCF!{L(CF0)}{r_g}:{CL}{r_g})<=MAX(DCF!{L(C0+1)}{r_g}:{L(C0+NH-1)}{r_g})"),
    ("Every football-field bar is ordered low to high", f"=SUMPRODUCT(--(Ranges!C{R_BAR0}:C{R_BAR1}>Ranges!D{R_BAR0}:D{R_BAR1}))=0"),
    ("Every bar has positive width", f"=MIN(Ranges!F{R_BAR0}:F{R_BAR1})>0"),
    ("Methodology weights sum to 100%", f"=ABS(Ranges!G{R_WSUM}-1)<0.0001"),
    ("Recommended range is ordered low <= mid <= high",
     f"=AND(Ranges!C{R_REC}<=Ranges!E{R_REC},Ranges!E{R_REC}<=Ranges!D{R_REC})"),
    ("Recommended range sits inside the overall range of methods",
     f"=AND(Ranges!C{R_REC}>=Ranges!C{R_MIN},Ranges!D{R_REC}<=Ranges!C{R_MAX})"),
    ("Analyst consensus bar is ordered low <= average <= high",
     f"=AND({INP['Analyst target — low (INR)']}<={INP['Analyst target — average (INR)']},"
     f"{INP['Analyst target — average (INR)']}<={INP['Analyst target — high (INR)']})"),
    ("Chart data ties to the Ranges sheet",
     f"=AND(ABS(SUM('Football Field'!E{R_CH0}:E{R_CH1})-(SUM(Ranges!C{R_BAR0}:C{R_BAR1})+Ranges!C{R_REC}))<0.01,"
     f"ABS(SUM('Football Field'!F{R_CH0}:F{R_CH1})-(SUM(Ranges!D{R_BAR0}:D{R_BAR1})+Ranges!D{R_REC}))<0.01)"),
    ("Share price and share count match Projects 3 and 4",
     f"=AND(ABS({INP['TCS share price (INR)']}-{PRICE})<0.01,ABS({INP['Shares outstanding (m)']}-{SHARES})<0.01)"),
    ("Current price sits inside the overall range",
     f"=AND({INP['TCS share price (INR)']}>=Ranges!C{R_MIN},{INP['TCS share price (INR)']}<=Ranges!C{R_MAX})"),
]
rr = 5
for label, fml in checks:
    wk.cell(rr, 2, label).font = F_TEXT
    c = wk.cell(rr, 3, fml); c.font = F_FORMULA
    c.alignment = Alignment(horizontal="center")
    rr += 1
R_CK0, R_CK1 = 5, rr - 1
rr += 1
wk.cell(rr, 2, "MODEL STATUS").font = Font(name=FONT, size=12, bold=True, color=book.theme.primary)
c = wk.cell(rr, 3, f'=IF(COUNTIF(C{R_CK0}:C{R_CK1},FALSE)=0,"MODEL OK","CHECK FAILED")')
c.font = Font(name=FONT, size=12, bold=True); c.fill = book.fill_accent
c.alignment = Alignment(horizontal="center")
rr += 2
book.note(wk, rr, "The growth-discipline check is the unusual one: it fails if the forecast assumes TCS grows faster than it "
                  "has in any of the last three years. A DCF that quietly solves for the current share price by lifting growth "
                  "is the most common way this analysis goes wrong, and a check is a cheaper defence than a reviewer.", 5, height=40)

# --------------------------------------------------------------------------- #
# Cover, finish, export
# --------------------------------------------------------------------------- #
book.cover(
    blurb="Everything Projects 2 to 4 produced about what a business like TCS is worth, put on one page, plus the DCF the "
          "football field needed and did not have. Seven methodologies, each a range rather than a point, against the "
          "12 September 2026 share price.",
    method=[
        "A compact five-year DCF in INR: local risk-free rate, beta regressed against the NIFTY 50, terminal value by "
        "perpetuity growth and by exit multiple with each checking the other, and a live WACC-by-growth sensitivity grid.",
        "Trading-multiple ranges carried in from Project 3 (peer 25th to 75th percentile) and precedent-transaction ranges "
        "from Project 4, each flagged for how far it can be trusted.",
        "The beta regression Project 2 used for Apple was run for TCS and rejected: Indian shares against the S&P 500 give "
        "R-squared of 0.01. The WACC sheet records the rejected regressions rather than deleting them.",
        "Eighteen integrity checks, including one that fails if the forecast assumes growth above anything TCS has recently "
        "achieved.",
    ],
    toc=[("Football Field", "the chart and the data behind it"),
         ("Ranges", "every bar, its source and its caveat"),
         ("DCF", "five-year forecast, terminal value both ways, sensitivity grid"),
         ("WACC", "beta regressions accepted and rejected, cost of capital build"),
         ("Inputs", "market data, cost of capital and forecast drivers"),
         ("Checks", "eighteen tests; must read MODEL OK")],
    highlights=[("DCF value per share (INR)", f"=DCF!{CL}{r_vps}", NUM0),
                ("Current share price (INR)", f"={INP['TCS share price (INR)']}", NUM0),
                ("Recommended range (INR)", f"=Ranges!C{R_REC}", NUM0),
                ("Recommended midpoint (INR)", f"=Ranges!E{R_REC}", NUM0),
                ("WACC (INR)", f"={INP['WACC (INR)']}", PCT),
                ("Terminal value as % of EV", f"=DCF!{CL}{r_tvpct}", PCT)],
    sources=["TCS annual financials FY2023-FY2026 (year to 31 March), Yahoo Finance fundamentals for TCS.NS, restating the consolidated annual report.",
             "Share price, 52-week range and share count: Yahoo Finance, 12 September 2026 — the same market date as Projects 3 and 4.",
             "Risk-free rate: India 10-year G-Sec at 7.02%, 11 September 2026.",
             "Equity risk premium: 4.5% mature-market premium (as used in Project 2) plus a 2.2% India country risk premium (Damodaran).",
             "Beta: five years of weekly returns to 12 September 2026, TCS against the NIFTY 50, 254 observations.",
             "Trading-multiple ranges from Project 3; precedent-transaction ranges from Project 4."])

book.finish(freeze={"DCF": "C6", "Ranges": "C5", "Football Field": "C6"},
            repeat_rows={"DCF": "1:5"})
path = os.path.join(HERE, "TCS_Football_Field.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
