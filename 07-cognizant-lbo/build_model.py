"""
Build 07-cognizant-lbo/Cognizant_LBO.xlsx — a hypothetical take-private of Cognizant Technology Solutions.

Sheets: Cover · Summary · Sources & Uses · Model · Debt Schedule · Returns · Sensitivity · Inputs · Checks

Cognizant is chosen because Projects 3 and 4 already cover IT services: the trading comps, the precedent
transactions and now the sponsor floor all sit in one sector, so the three answers can be compared directly.
The transaction is hypothetical.

Interest is charged on the opening debt balance of each year. That is a deliberate choice: charging it on the
average balance makes the model circular (interest drives cash flow, cash flow drives paydown, paydown drives
interest), and a circular model that needs iterative calculation switched on is a model that breaks on someone
else's machine. The cost is a small overstatement of interest in years where debt falls quickly.
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"

# --------------------------------------------------------------------------- #
# Source data — Cognizant FY2025 (year to 31 December 2025), US$ millions.
# Yahoo Finance fundamentals for CTSH, restating the 10-K. Price: NASDAQ close 12 Sep 2026.
# --------------------------------------------------------------------------- #
C = dict(price=62.585, shares=488.0, revenue=21108.0, ebitda=4066.0, ebit=3327.0, ni=2230.0,
         da=550.0, capex=288.0, debt=1152.0, cash=1901.0, equity=15015.0,
         w52hi=87.03, w52lo=37.08)
YEARS = ["FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]
GROWTH = [0.030, 0.035, 0.040, 0.040, 0.040]
MARGIN = [0.196, 0.200, 0.204, 0.207, 0.210]

book = Book(theme=THEMES["neutral"], project_no=7, project="Leveraged Buyout — Debt Schedule and Returns",
            company="Cognizant Technology Solutions", units="US$ millions unless stated",
            as_of="FY2025 10-K · market data 12 Sep 2026")


class Sheet:
    def __init__(self, name, ncols, label_width=50, col_width=14, subtitle=None, valcol=3):
        self.ws = book.sheet(name, ncols=ncols, label_width=label_width, col_width=col_width, subtitle=subtitle)
        self.name, self.ncols, self.valcol, self.r = name, ncols, valcol, 4
        self.ref = {}

    def section(self, text):
        book.section(self.ws, self.r, text, self.ncols); self.r += 1

    def blank(self, n=1):
        self.r += n

    def head(self, labels, first_col=None, label=""):
        book.year_header(self.ws, self.r, labels, first_col=first_col or self.valcol, label=label); self.r += 1

    def row(self, label, value, fmt=NUM0, font=None, note="", bold=False, border=None, col=None):
        ws, r = self.ws, self.r
        c = ws.cell(r, 2, label); c.font = F_BOLD if bold else F_TEXT
        col = col or self.valcol
        v = ws.cell(r, col, value)
        if font is None:
            font = F_FORMULA if (isinstance(value, str) and value.startswith("=")) else F_INPUT
        v.font = font; v.number_format = fmt
        if border: v.border = border
        if note:
            n = ws.cell(r, self.ncols, note); n.font = F_NOTE; n.alignment = Alignment(wrap_text=True, vertical="top")
        self.ref[label] = f"'{self.name}'!${L(col)}${r}"
        self.r += 1
        return self.ref[label]

    def multi(self, label, values, fmt=NUM0, font=None, bold=False, border=None, first_col=None, rel=False):
        ws, r = self.ws, self.r
        c = ws.cell(r, 2, label); c.font = F_BOLD if bold else F_TEXT
        fc = first_col or self.valcol
        refs = []
        for i, value in enumerate(values):
            v = ws.cell(r, fc + i, value)
            f = font or (F_FORMULA if (isinstance(value, str) and value.startswith("=")) else F_INPUT)
            v.font = f; v.number_format = fmt
            if border: v.border = border
            refs.append(f"{L(fc+i)}{r}" if rel else f"'{self.name}'!${L(fc+i)}${r}")
        self.ref[label] = refs
        self.r += 1
        return refs

    def note(self, text, height=30):
        book.note(self.ws, self.r, text, self.ncols, height=height); self.r += 1


# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
si = Sheet("Inputs", ncols=6, label_width=54, subtitle="Standalone financials, deal terms, debt structure and operating case")
si.ws.column_dimensions["F"].width = 58
I = {}
si.section("COGNIZANT — FY2025A (US$m)")
for lab, k, f in [("Revenue", "revenue", NUM0), ("EBITDA", "ebitda", NUM0), ("EBIT", "ebit", NUM0),
                  ("Depreciation and amortisation", "da", NUM0), ("Capital expenditure", "capex", NUM0),
                  ("Total debt", "debt", NUM0), ("Cash and equivalents", "cash", NUM0),
                  ("Shares outstanding (m)", "shares", NUM0), ("Share price (US$, 12 Sep 2026)", "price", NUM2)]:
    I[lab] = si.row(lab, C[k], f)
I["mcap"] = si.row("Market capitalisation", f"={I['Share price (US$, 12 Sep 2026)']}*{I['Shares outstanding (m)']}", NUM0, font=F_FORMULA)
I["netdebt"] = si.row("Net debt / (net cash)", f"={I['Total debt']}-{I['Cash and equivalents']}", NUM0, font=F_FORMULA)
I["ev"] = si.row("Enterprise value, unaffected", f"={I['mcap']}+{I['netdebt']}", NUM0, font=F_FORMULA)
I["evebitda"] = si.row("EV / EBITDA, unaffected", f"={I['ev']}/{I['EBITDA']}", MULT, font=F_FORMULA,
                       note="The sector de-rated through 2026; this is the entry point that makes a take-private arithmetically possible at all")

si.blank(); si.section("TRANSACTION")
I["prem"] = si.row("Offer premium to unaffected price", 0.30, PCT, note="Typical control premium for a US public take-private")
I["minqcash"] = si.row("Minimum cash left in the business", 500.0, NUM0)
I["fin_fee"] = si.row("Financing fees as % of debt raised", 0.020, PCT, note="Underwriting and arrangement; capitalised and amortised over the term")
I["adv_fee"] = si.row("Advisory and other fees as % of transaction value", 0.010, PCT, note="Sell-side and buy-side advisory, legal, accounting")
I["tax"] = si.row("Cash tax rate", 0.250, PCT, note="US federal plus state, blended. FY2025 reported rate was distorted by one-off items")

si.blank(); si.section("DEBT STRUCTURE")
I["tlb_x"] = si.row("Term Loan B (x EBITDA)", 3.00, MULT)
I["notes_x"] = si.row("Senior notes (x EBITDA)", 2.25, MULT)
I["sofr"] = si.row("SOFR", 0.0395, PCT, note="Three-month rate, 12 Sep 2026")
I["tlb_spread"] = si.row("Term Loan B spread over SOFR", 0.0350, PCT)
I["notes_cpn"] = si.row("Senior notes coupon (fixed)", 0.0850, PCT)
I["tlb_amort"] = si.row("Term Loan B mandatory amortisation (% of original, per year)", 0.010, PCT)
I["sweep"] = si.row("Cash sweep — % of free cash flow applied to debt", 1.00, PCT,
                    note="100% sweep on the term loan until it is repaid. Notes are bullet and cannot be swept")
I["fee_amort"] = si.row("Financing fee amortisation period (years)", 7, '0')

si.blank(); si.section("OPERATING CASE")
I["nwc"] = si.row("Increase in working capital as % of incremental revenue", 0.120, PCT)
I["capex_pct"] = si.row("Capital expenditure as % of revenue", 0.014, PCT, note="FY2025 actual: 1.4%. Asset-light")
I["da_pct"] = si.row("Depreciation and amortisation as % of revenue", 0.026, PCT)
I["exit_x"] = si.row("Exit EV / EBITDA multiple", 9.00, MULT,
                     note="Below the entry multiple on purpose. Assuming multiple expansion is how an LBO model is made to lie")
I["hurdle"] = si.row("Sponsor target internal rate of return", 0.200, PCT, note="The return the deal has to clear to be worth doing")

si.blank()
si.note("Revenue growth and EBITDA margin by year are set on the Model sheet, where they can be read against the "
        "historical figures. Everything else lives here.", height=26)

# --------------------------------------------------------------------------- #
# Sources & Uses
# --------------------------------------------------------------------------- #
su = Sheet("Sources & Uses", ncols=6, label_width=54, subtitle="What the transaction costs and where the money comes from")
su.ws.column_dimensions["F"].width = 56
S = {}
su.section("OFFER")
S["offer"] = su.row("Offer price per share (US$)", f"={I['Share price (US$, 12 Sep 2026)']}*(1+{I['prem']})", NUM2, bold=True)
S["eqpur"] = su.row("Equity purchase price", f"={S['offer']}*{I['Shares outstanding (m)']}", NUM0, bold=True)
S["tev"] = su.row("Transaction enterprise value", f"={S['eqpur']}+{I['Total debt']}-{I['Cash and equivalents']}", NUM0, bold=True)
S["entry_x"] = su.row("Entry EV / EBITDA", f"={S['tev']}/{I['EBITDA']}", MULT, bold=True,
                      note="The single most important number in the model: it sets the price and, with the exit multiple, most of the return")

su.blank(); su.section("USES")
S["u_eq"] = su.row("Purchase of equity", f"={S['eqpur']}", NUM0, font=F_LINK)
S["u_refi"] = su.row("Refinance existing debt", f"={I['Total debt']}", NUM0, font=F_LINK)
S["tlb"] = su.row("  memo: Term Loan B raised", f"={I['EBITDA']}*{I['tlb_x']}", NUM0, font=F_FORMULA)
S["notes"] = su.row("  memo: Senior notes raised", f"={I['EBITDA']}*{I['notes_x']}", NUM0, font=F_FORMULA)
S["debt_tot"] = su.row("  memo: Total debt raised", f"={S['tlb']}+{S['notes']}", NUM0, font=F_FORMULA, bold=True)
S["u_finfee"] = su.row("Financing fees", f"={S['debt_tot']}*{I['fin_fee']}", NUM0)
S["u_advfee"] = su.row("Advisory and other fees", f"={S['tev']}*{I['adv_fee']}", NUM0)
S["uses"] = su.row("Total uses", f"={S['u_eq']}+{S['u_refi']}+{S['u_finfee']}+{S['u_advfee']}", NUM0,
                   bold=True, border=DOUBLE_BORDER)

su.blank(); su.section("SOURCES")
S["s_tlb"] = su.row("Term Loan B", f"={S['tlb']}", NUM0, font=F_LINK)
S["s_notes"] = su.row("Senior notes", f"={S['notes']}", NUM0, font=F_LINK)
S["s_cash"] = su.row("Cash from the balance sheet",
                     f"=MAX(0,{I['Cash and equivalents']}-{I['minqcash']})", NUM0, font=F_FORMULA,
                     note="Excess cash above the minimum operating balance is used to fund the deal")
S["s_eq"] = su.row("Sponsor equity", f"={S['uses']}-{S['s_tlb']}-{S['s_notes']}-{S['s_cash']}", NUM0,
                   bold=True, note="The cheque the fund writes, and the denominator of every return in this model")
S["sources"] = su.row("Total sources", f"={S['s_tlb']}+{S['s_notes']}+{S['s_cash']}+{S['s_eq']}", NUM0,
                      bold=True, border=DOUBLE_BORDER)

su.blank()
S["lev"] = su.row("Opening total leverage (x EBITDA)", f"={S['debt_tot']}/{I['EBITDA']}", MULT, bold=True)
S["eqpct"] = su.row("Sponsor equity as % of total capitalisation", f"={S['s_eq']}/{S['uses']}", PCT)
su.note("An equity cheque of this size would in practice be a consortium of several funds rather than one. That does not "
        "change the arithmetic — it changes who writes it — so the model treats the sponsor as a single investor.", height=28)

# --------------------------------------------------------------------------- #
# Model — operating forecast and free cash flow
# --------------------------------------------------------------------------- #
NY = len(YEARS)
CE = 3            # entry column (FY2025A)
C1 = CE + 1       # first forecast column
CN = CE + NY
sm = Sheet("Model", ncols=CN + 1, label_width=48, col_width=13.5,
           subtitle="Operating forecast and free cash flow before debt repayment · US$ millions", valcol=CE)
sm.section("OPERATING FORECAST")
sm.head(["FY2025A"] + YEARS, first_col=CE)
M = {}
M["rev"] = sm.multi("Revenue", [f"={I['Revenue']}"] + [f"={L(CE+i)}{sm.r}*(1+{L(CE+i+1)}{sm.r+1})" for i in range(NY)],
                    NUM0, bold=True)
M["g"] = sm.multi("   Growth", [""] + GROWTH, PCT)
M["margin"] = sm.multi("   EBITDA margin", [f"={I['EBITDA']}/{I['Revenue']}"] + MARGIN, PCT)
M["ebitda"] = sm.multi("EBITDA", [f"={I['EBITDA']}"] + [f"={L(CE+i+1)}{sm.r-3}*{L(CE+i+1)}{sm.r-1}" for i in range(NY)],
                       NUM0, bold=True)
M["da"] = sm.multi("Depreciation and amortisation", [f"={I['Depreciation and amortisation']}"] + [f"={L(CE+i+1)}{sm.r-4}*{I['da_pct']}" for i in range(NY)], NUM0)
M["capex"] = sm.multi("Capital expenditure", [f"={I['Capital expenditure']}"] + [f"={L(CE+i+1)}{sm.r-5}*{I['capex_pct']}" for i in range(NY)], NUM0)
M["nwc"] = sm.multi("Increase in working capital", [""] +
                    [f"=({L(CE+i+1)}{sm.r-6}-{L(CE+i)}{sm.r-6})*{I['nwc']}" for i in range(NY)], NUM0)
sm.blank()
sm.section("FREE CASH FLOW BEFORE DEBT SERVICE")
r_ebitda = sm.r - 6
M["fcf_ebitda"] = sm.multi("EBITDA", [""] + [f"={L(CE+i+1)}{r_ebitda}" for i in range(NY)], NUM0, font=F_LINK)
M["fcf_capex"] = sm.multi("Less: capital expenditure", [""] + [f"=-{L(CE+i+1)}{r_ebitda+2}" for i in range(NY)], NUM0)
M["fcf_nwc"] = sm.multi("Less: increase in working capital", [""] + [f"=-{L(CE+i+1)}{r_ebitda+3}" for i in range(NY)], NUM0)
sm.note("Interest and cash tax are charged on the Debt Schedule, where the opening balance for each year is known. "
        "Free cash flow after debt service flows back here only as the sweep amount.", height=26)

# --------------------------------------------------------------------------- #
# Debt Schedule
# --------------------------------------------------------------------------- #
sd = Sheet("Debt Schedule", ncols=CN + 1, label_width=48, col_width=13.5,
           subtitle="Interest charged on opening balances · cash sweep applied to the term loan", valcol=CE)
sd.section("INTEREST AND CASH TAX")
sd.head(["At close"] + YEARS, first_col=CE)
DS = {}
r_tlb_open = sd.r
DS["tlb_open"] = sd.multi("Term Loan B — opening balance",
                          [f"={S['tlb']}"] + [f"={L(CE+i)}{sd.r+3}" for i in range(NY)], NUM0)
DS["tlb_amort"] = sd.multi("   Mandatory amortisation", [""] +
                           [f"=-MIN({L(CE+i+1)}{r_tlb_open},{S['tlb']}*{I['tlb_amort']})" for i in range(NY)], NUM0)
DS["tlb_sweep"] = sd.multi("   Cash sweep", [""] + ["=0"] * NY, NUM0)
DS["tlb_close"] = sd.multi("Term Loan B — closing balance",
                           [f"={S['tlb']}"] + [f"={L(CE+i+1)}{r_tlb_open}+{L(CE+i+1)}{r_tlb_open+1}+{L(CE+i+1)}{r_tlb_open+2}" for i in range(NY)],
                           NUM0, bold=True, border=TOTAL_BORDER)
sd.blank()
r_notes = sd.r
DS["notes_bal"] = sd.multi("Senior notes — balance (bullet)", [f"={S['notes']}"] + [f"={S['notes']}"] * NY, NUM0, font=F_LINK)
DS["tot_debt"] = sd.multi("Total debt",
                          [f"={L(CE)}{r_tlb_open+3}+{L(CE)}{r_notes}"] +
                          [f"={L(CE+i+1)}{r_tlb_open+3}+{L(CE+i+1)}{r_notes}" for i in range(NY)],
                          NUM0, bold=True)
sd.blank()
r_int = sd.r
DS["int_tlb"] = sd.multi("Interest — Term Loan B", [""] +
                         [f"={L(CE+i+1)}{r_tlb_open}*({I['sofr']}+{I['tlb_spread']})" for i in range(NY)], NUM0)
DS["int_notes"] = sd.multi("Interest — senior notes", [""] + [f"={S['notes']}*{I['notes_cpn']}" for i in range(NY)], NUM0)
DS["fee_am"] = sd.multi("Financing fee amortisation", [""] + [f"={S['u_finfee']}/{I['fee_amort']}" for i in range(NY)], NUM0)
DS["int_tot"] = sd.multi("Total interest and fee amortisation", [""] +
                         [f"=SUM({L(CE+i+1)}{r_int}:{L(CE+i+1)}{r_int+2})" for i in range(NY)], NUM0,
                         bold=True, border=TOTAL_BORDER)
sd.blank()
r_tax = sd.r
r_m_ebitda = r_ebitda
DS["ebt"] = sd.multi("Profit before tax", [""] +
                     [f"=Model!{L(CE+i+1)}{r_m_ebitda}-Model!{L(CE+i+1)}{r_m_ebitda+1}-{L(CE+i+1)}{r_int+3}" for i in range(NY)], NUM0)
DS["cashtax"] = sd.multi("Cash tax", [""] + [f"=-MAX(0,{L(CE+i+1)}{r_tax})*{I['tax']}" for i in range(NY)], NUM0)

sd.blank()
sd.section("CASH AVAILABLE FOR DEBT REPAYMENT")
r_cf = sd.r
DS["cf_ebitda"] = sd.multi("EBITDA", [""] + [f"=Model!{L(CE+i+1)}{r_m_ebitda}" for i in range(NY)], NUM0, font=F_LINK)
DS["cf_capex"] = sd.multi("Less: capital expenditure", [""] + [f"=-Model!{L(CE+i+1)}{r_m_ebitda+2}" for i in range(NY)], NUM0)
DS["cf_nwc"] = sd.multi("Less: increase in working capital", [""] + [f"=-Model!{L(CE+i+1)}{r_m_ebitda+3}" for i in range(NY)], NUM0)
DS["cf_int"] = sd.multi("Less: cash interest", [""] +
                        [f"=-({L(CE+i+1)}{r_int}+{L(CE+i+1)}{r_int+1})" for i in range(NY)], NUM0)
DS["cf_tax"] = sd.multi("Less: cash tax", [""] + [f"={L(CE+i+1)}{r_tax+1}" for i in range(NY)], NUM0)
DS["cf_amort"] = sd.multi("Less: mandatory amortisation", [""] + [f"={L(CE+i+1)}{r_tlb_open+1}" for i in range(NY)], NUM0)
DS["cf_free"] = sd.multi("Cash available for the sweep", [""] +
                         [f"=SUM({L(CE+i+1)}{r_cf}:{L(CE+i+1)}{r_cf+5})" for i in range(NY)], NUM0,
                         bold=True, border=DOUBLE_BORDER)

# wire the sweep back into the term loan rows
for i in range(NY):
    col = L(CE + i + 1)
    sd.ws.cell(r_tlb_open + 2, CE + i + 1).value = (
        f"=-MIN(MAX(0,{col}{sd.r-1})*{I['sweep']},{col}{r_tlb_open}+{col}{r_tlb_open+1})")

sd.blank()
sd.section("CREDIT STATISTICS")
r_cs = sd.r
DS["lev_tot"] = sd.multi("Total debt / EBITDA", [f"={L(CE)}{r_notes+1}/{I['EBITDA']}"] +
                         [f"={L(CE+i+1)}{r_notes+1}/Model!{L(CE+i+1)}{r_m_ebitda}" for i in range(NY)], MULT, bold=True)
DS["cov"] = sd.multi("EBITDA / cash interest", [""] +
                     [f"=Model!{L(CE+i+1)}{r_m_ebitda}/({L(CE+i+1)}{r_int}+{L(CE+i+1)}{r_int+1})" for i in range(NY)], MULT, bold=True)
sd.blank()
sd.note("Leverage of over five times EBITDA at close, against interest cover of roughly two and a half times, is an "
        "aggressive but financeable structure for a business with Cognizant's cash conversion. The covenant question a "
        "lender would ask is what happens to cover if revenue falls rather than grows; the sensitivity grids answer it.", height=40)

# --------------------------------------------------------------------------- #
# Returns
# --------------------------------------------------------------------------- #
sr = Sheet("Returns", ncols=6, label_width=56, subtitle="Exit, money multiple, internal rate of return and the price that clears the hurdle")
sr.ws.column_dimensions["F"].width = 56
R = {}
sr.section("EXIT AT THE END OF YEAR 5")
R["ex_ebitda"] = sr.row("Exit-year EBITDA", f"=Model!{L(CN)}{r_m_ebitda}", NUM0, font=F_LINK)
R["ex_x"] = sr.row("Exit EV / EBITDA", f"={I['exit_x']}", MULT, font=F_LINK)
R["ex_ev"] = sr.row("Exit enterprise value", f"={R['ex_ebitda']}*{R['ex_x']}", NUM0, bold=True)
R["ex_debt"] = sr.row("Less: net debt at exit", f"=-'Debt Schedule'!{L(CN)}{r_notes+1}", NUM0, font=F_FORMULA)
R["ex_eq"] = sr.row("Exit equity value", f"={R['ex_ev']}+{R['ex_debt']}", NUM0, bold=True, border=TOTAL_BORDER)
sr.blank()
R["entry_eq"] = sr.row("Sponsor equity invested at close", f"={S['s_eq']}", NUM0, font=F_LINK)
R["moic"] = sr.row("Money multiple (MOIC)", f"={R['ex_eq']}/{R['entry_eq']}", MULT, bold=True)
R["irr"] = sr.row("Internal rate of return", f"={R['moic']}^(1/5)-1", PCT, bold=True, border=DOUBLE_BORDER,
                  note="A single entry and a single exit, so the IRR is the fifth root of the money multiple. No interim dividends are assumed")
R["vs"] = sr.row("Versus the 20% hurdle (percentage points)", f"=({R['irr']}-{I['hurdle']})*100", NUM2, bold=True)

sr.blank(); sr.section("VALUE CREATION BRIDGE")
R["vb_debt"] = sr.row("Debt repaid over the hold", f"='Debt Schedule'!{L(CE)}{r_notes+1}-'Debt Schedule'!{L(CN)}{r_notes+1}", NUM0)
R["vb_ebitda"] = sr.row("EBITDA growth, valued at the exit multiple",
                        f"=({R['ex_ebitda']}-{I['EBITDA']})*{R['ex_x']}", NUM0)
R["vb_mult"] = sr.row("Multiple change, valued on entry EBITDA",
                      f"=({R['ex_x']}-{S['entry_x']})*{I['EBITDA']}", NUM0,
                      note="Negative here because the exit multiple is set below entry on purpose")
R["vb_tot"] = sr.row("Total value created", f"={R['ex_eq']}-{R['entry_eq']}", NUM0, bold=True, border=TOTAL_BORDER)

sr.blank(); sr.section("THE PRICE A SPONSOR COULD ACTUALLY PAY")
R["req_eq"] = sr.row("Sponsor equity consistent with a 20% return",
                     f"={R['ex_eq']}/(1+{I['hurdle']})^5", NUM0, font=F_FORMULA)
R["max_tev"] = sr.row("Implied maximum equity purchase price",
                      f"={R['req_eq']}+{S['s_tlb']}+{S['s_notes']}+{S['s_cash']}-{I['Total debt']}-{S['u_finfee']}-{S['u_advfee']}", NUM0)
R["max_px"] = sr.row("Implied maximum offer price per share (US$)",
                     f"={R['max_tev']}/{I['Shares outstanding (m)']}", NUM2, bold=True, border=DOUBLE_BORDER,
                     note="The LBO floor for Project 5's football field: what a financial buyer could pay and still clear its hurdle")
R["max_prem"] = sr.row("Implied maximum premium to the unaffected price",
                       f"={R['max_px']}/{I['Share price (US$, 12 Sep 2026)']}-1", PCT, bold=True)
sr.note("Advisory fees move slightly with the transaction value, so the maximum price is very marginally circular. The "
        "effect is under a dollar a share and the fees are held at the base case rather than switching on iterative "
        "calculation, which would make the workbook fragile on another machine.", height=34)

# --------------------------------------------------------------------------- #
# Sensitivity
# --------------------------------------------------------------------------- #
ss = Sheet("Sensitivity", ncols=8, label_width=40, col_width=13,
           subtitle="Internal rate of return · every cell rebuilds the equity cheque and the exit from scratch")
PREMS = [0.10, 0.20, 0.30, 0.40, 0.50]
EXITS = [8.00, 8.50, 9.00, 9.50, 10.00]
LEVS = [4.25, 4.75, 5.25, 5.75, 6.25]

ex_eq_fixed = f"'Returns'!{R['ex_eq'].split('!')[1]}"


def irr_formula(prem_ref, exit_ref, lev_ref=None):
    lev = lev_ref or f"({I['tlb_x']}+{I['notes_x']})"
    debt = f"({I['EBITDA']}*{lev})"
    eqpur = f"({I['Share price (US$, 12 Sep 2026)']}*(1+{prem_ref})*{I['Shares outstanding (m)']})"
    tev = f"({eqpur}+{I['Total debt']}-{I['Cash and equivalents']})"
    uses = f"({eqpur}+{I['Total debt']}+{debt}*{I['fin_fee']}+{tev}*{I['adv_fee']})"
    sponsor = f"({uses}-{debt}-{S['s_cash']})"
    exit_eq = f"(Model!{L(CN)}{r_m_ebitda}*{exit_ref}-('Debt Schedule'!{L(CN)}{r_notes+1}+{debt}-{S['debt_tot']}))"
    return f"=({exit_eq}/{sponsor})^(1/5)-1"


ss.section("IRR — OFFER PREMIUM versus EXIT MULTIPLE")
ss.ws.cell(ss.r, 2, "Exit EV/EBITDA  \\  Offer premium").font = F_BOLD
for j, p in enumerate(PREMS):
    c = ss.ws.cell(ss.r, 3 + j, p); c.font = book.f_header; c.fill = book.fill_secondary
    c.number_format = PCT; c.alignment = Alignment(horizontal="center")
h1 = ss.r; ss.r += 1
g1 = ss.r
for x in EXITS:
    c = ss.ws.cell(ss.r, 2, x); c.font = F_BOLD; c.number_format = MULT
    for j, p in enumerate(PREMS):
        cc = ss.ws.cell(ss.r, 3 + j, irr_formula(f"{L(3+j)}${h1}", f"$B{ss.r}"))
        cc.font = F_FORMULA; cc.number_format = PCT
        if abs(p - 0.30) < 1e-9 and abs(x - 9.0) < 1e-9:
            cc.fill = book.fill_accent; cc.font = Font(name=FONT, size=10, bold=True)
    ss.r += 1
g1e = ss.r - 1
ss.blank()

ss.section("IRR — OPENING LEVERAGE versus EXIT MULTIPLE (at the 30% base-case premium)")
ss.ws.cell(ss.r, 2, "Exit EV/EBITDA  \\  Total leverage").font = F_BOLD
for j, lv in enumerate(LEVS):
    c = ss.ws.cell(ss.r, 3 + j, lv); c.font = book.f_header; c.fill = book.fill_secondary
    c.number_format = MULT; c.alignment = Alignment(horizontal="center")
h2 = ss.r; ss.r += 1
g2 = ss.r
for x in EXITS:
    c = ss.ws.cell(ss.r, 2, x); c.font = F_BOLD; c.number_format = MULT
    for j, lv in enumerate(LEVS):
        cc = ss.ws.cell(ss.r, 3 + j, irr_formula(I['prem'], f"$B{ss.r}", f"{L(3+j)}${h2}"))
        cc.font = F_FORMULA; cc.number_format = PCT
        if abs(lv - 5.25) < 1e-9 and abs(x - 9.0) < 1e-9:
            cc.fill = book.fill_accent; cc.font = Font(name=FONT, size=10, bold=True)
    ss.r += 1
g2e = ss.r - 1
ss.blank()
ss.note("The leverage grid holds the debt paydown path at the base case and flexes only the opening quantum, so it "
        "slightly flatters high-leverage cases (more debt would mean more interest and less sweep). It is shown to "
        "illustrate the direction and size of the effect, not to price a financing.", height=34)

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sy = Sheet("Summary", ncols=6, label_width=56, subtitle="The transaction on one page")
sy.ws.column_dimensions["F"].width = 54
sy.section("THE TRANSACTION")
sy.row("Offer price per share (US$)", f"={S['offer']}", NUM2, font=F_LINK)
sy.row("Premium to unaffected price", f"={I['prem']}", PCT, font=F_LINK)
sy.row("Transaction enterprise value", f"={S['tev']}", NUM0, font=F_LINK)
sy.row("Entry EV / EBITDA", f"={S['entry_x']}", MULT, font=F_LINK, bold=True)
sy.row("Opening leverage (x EBITDA)", f"={S['lev']}", MULT, font=F_LINK)
sy.row("Sponsor equity cheque", f"={S['s_eq']}", NUM0, font=F_LINK, bold=True)
sy.blank(); sy.section("THE RETURN")
sy.row("Exit EV / EBITDA assumed", f"={I['exit_x']}", MULT, font=F_LINK)
sy.row("Money multiple", f"={R['moic']}", MULT, font=F_LINK, bold=True)
sy.row("Internal rate of return", f"={R['irr']}", PCT, font=F_LINK, bold=True)
sy.row("Versus the 20% hurdle (pp)", f"={R['vs']}", NUM2, font=F_LINK, bold=True, border=DOUBLE_BORDER)
sy.blank(); sy.section("THE USEFUL OUTPUT")
sy.row("Maximum offer price at a 20% return (US$)", f"={R['max_px']}", NUM2, font=F_LINK, bold=True)
sy.row("Maximum premium to the unaffected price", f"={R['max_prem']}", PCT, font=F_LINK, bold=True)
sy.row("Unaffected share price (US$)", f"={I['Share price (US$, 12 Sep 2026)']}", NUM2, font=F_LINK)
sy.blank()
for t in ["Cognizant is buyable precisely because the sector de-rated: at roughly seven times EBITDA unaffected, a "
          "financial buyer can lever it. At twenty times, as the sector traded in 2021, no sponsor could.",
          "The return is driven by debt paydown and margin expansion, not by multiple expansion — the exit multiple is "
          "set below entry on purpose. Any LBO whose return depends on selling at a higher multiple than it bought is "
          "an assumption dressed as an analysis.",
          "The most useful number here is not the IRR at a 30% premium. It is the maximum price a sponsor could pay and "
          "still clear 20%, because that is the floor under the share price in any sale process, and it is the ninth bar "
          "on Project 5's football field.",
          "The equity cheque is large enough that this would be a consortium deal. The arithmetic is unchanged; the "
          "practicality is not."]:
    sy.note("\\u2022  " + t, height=15 * max(2, len(t) // 100 + 1))

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
sk = Sheet("Checks", ncols=5, label_width=72, col_width=16, subtitle="Every check must read TRUE")
sk.section("INTEGRITY CHECKS")
CHK = [
    ("Sources equal uses", f"=ABS({S['sources']}-{S['uses']})<0.01"),
    ("Sponsor equity is positive", f"={S['s_eq']}>0"),
    ("Opening leverage equals the sum of the tranches", f"=ABS({S['lev']}-({I['tlb_x']}+{I['notes_x']}))<0.001"),
    ("Term loan balance never goes negative", f"=MIN('Debt Schedule'!{L(C1)}{r_tlb_open+3}:{L(CN)}{r_tlb_open+3})>=-0.01"),
    ("Term loan closing balance ties to opening less repayments",
     f"=ABS('Debt Schedule'!{L(CN)}{r_tlb_open+3}-('Debt Schedule'!{L(CN)}{r_tlb_open}"
     f"+'Debt Schedule'!{L(CN)}{r_tlb_open+1}+'Debt Schedule'!{L(CN)}{r_tlb_open+2}))<0.01"),
    ("Debt falls over the hold", f"='Debt Schedule'!{L(CN)}{r_notes+1}<'Debt Schedule'!{L(CE)}{r_notes+1}"),
    ("Cash available for the sweep is positive in every year",
     f"=MIN('Debt Schedule'!{L(C1)}{r_cf+6}:{L(CN)}{r_cf+6})>0"),
    ("Interest cover stays above 1.5x in every year",
     f"=MIN('Debt Schedule'!{L(C1)}{r_cs+1}:{L(CN)}{r_cs+1})>1.5"),
    ("Leverage falls every year", f"=AND('Debt Schedule'!{L(CN)}{r_cs}<'Debt Schedule'!{L(C1)}{r_cs})"),
    ("Exit equity ties to exit enterprise value less net debt",
     f"=ABS({R['ex_eq']}-({R['ex_ev']}+{R['ex_debt']}))<0.01"),
    ("Money multiple and IRR are consistent", f"=ABS((1+{R['irr']})^5-{R['moic']})<0.0001"),
    ("Exit multiple is not above the entry multiple", f"={I['exit_x']}<={S['entry_x']}"),
    ("The maximum price is below the assumed offer price when the deal misses the hurdle",
     f"=IF({R['irr']}<{I['hurdle']},{R['max_px']}<{S['offer']},TRUE)"),
    ("Base-case sensitivity cell reconciles to the Returns sheet",
     f"=ABS(Sensitivity!E{g1+2}-{R['irr']})<0.002"),
    ("Entry EV/EBITDA is above the unaffected multiple", f"={S['entry_x']}>{I['evebitda']}"),
    ("Standalone figures tie to the reported FY2025 accounts",
     f"=AND(ABS({I['EBITDA']}-4066)<1,ABS({I['Revenue']}-21108)<1,ABS({I['Shares outstanding (m)']}-488)<1)"),
]
for label, fml in CHK:
    sk.ws.cell(sk.r, 2, label).font = F_TEXT
    v = sk.ws.cell(sk.r, 3, fml); v.font = F_FORMULA; v.alignment = Alignment(horizontal="center")
    sk.r += 1
K0, K1 = 5, sk.r - 1
sk.blank()
sk.ws.cell(sk.r, 2, "MODEL STATUS").font = Font(name=FONT, size=12, bold=True, color=book.theme.primary)
v = sk.ws.cell(sk.r, 3, f'=IF(COUNTIF(C{K0}:C{K1},FALSE)=0,"MODEL OK","CHECK FAILED")')
v.font = Font(name=FONT, size=12, bold=True); v.fill = book.fill_accent; v.alignment = Alignment(horizontal="center")
sk.r += 2
sk.note("The check that the exit multiple is not above the entry multiple is a discipline check, not an arithmetic one. "
        "It fails the model if someone later raises the exit assumption to make the return work, which is the single "
        "most common way an LBO model is quietly bent.", height=34)

# --------------------------------------------------------------------------- #
book.cover(
    blurb="A hypothetical take-private of Cognizant Technology Solutions by a financial sponsor, built on the FY2025 "
          "10-K and 12 September 2026 market data. Cognizant is used because Projects 3 and 4 already cover IT-services "
          "trading multiples and precedent M&A, so the sponsor's view of value can be set beside both.",
    method=[
        "Sources and uses from an offer premium, with debt sized as a multiple of EBITDA across a term loan and senior "
        "notes, and the sponsor equity falling out as the balancing figure.",
        "A five-year debt schedule with mandatory amortisation and a full cash sweep, interest charged on opening "
        "balances so the model never needs iterative calculation switched on.",
        "Returns by money multiple and IRR, a value-creation bridge splitting the return between debt paydown, EBITDA "
        "growth and multiple change, and two live sensitivity grids.",
        "A solve for the maximum price a sponsor could pay and still clear a 20% return — the LBO floor that belongs on "
        "Project 5's football field.",
    ],
    toc=[("Summary", "the transaction, the return and the maximum price"),
         ("Sources & Uses", "offer, debt quantum, fees and the equity cheque"),
         ("Model", "operating forecast and free cash flow"),
         ("Debt Schedule", "tranches, amortisation, sweep, interest and credit statistics"),
         ("Returns", "exit, MOIC, IRR, value bridge and the hurdle-clearing price"),
         ("Sensitivity", "IRR across premium, exit multiple and leverage"),
         ("Inputs", "standalone financials, deal terms, debt pricing and the operating case"),
         ("Checks", "sixteen tests; must read MODEL OK")],
    highlights=[("Offer price per share (US$)", f"={S['offer']}", NUM2),
                ("Entry EV / EBITDA", f"={S['entry_x']}", MULT),
                ("Sponsor equity cheque (US$m)", f"={S['s_eq']}", NUM0),
                ("Money multiple", f"={R['moic']}", MULT),
                ("Internal rate of return", f"={R['irr']}", PCT),
                ("Maximum price at a 20% return (US$)", f"={R['max_px']}", NUM2)],
    sources=["Cognizant Technology Solutions FY2025 Form 10-K (year to 31 December 2025), via the Yahoo Finance fundamentals series for CTSH.",
             "Share price and 52-week range: NASDAQ close, 12 September 2026.",
             "SOFR and US Treasury yields, 12 September 2026.",
             "Debt pricing, fee levels and the operating case are stated assumptions on the Inputs sheet, not quotations from lenders.",
             "The transaction is hypothetical. Nothing here suggests Cognizant is or has been in a sale process."])

book.finish(freeze={"Model": "D6", "Debt Schedule": "D6"}, repeat_rows={"Model": "1:5", "Debt Schedule": "1:5"})
path = os.path.join(HERE, "Cognizant_LBO.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
