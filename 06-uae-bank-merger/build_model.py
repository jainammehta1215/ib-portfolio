"""
Build 06-uae-bank-merger/ENBD_Mashreq_Merger_Model.xlsx — a hypothetical UAE banking consolidation.

Emirates NBD (DFM: EMIRATESNBD) acquiring Mashreqbank (DFM: MASQ). Both are real, listed and file public
accounts; the transaction is hypothetical and there is no suggestion either bank is in discussions.

Sheets: Cover · Summary · Inputs · Purchase Price · Accretion · Capital & TBV · Sensitivity · Checks

Bank M&A turns on three numbers, in this order: tangible book value per share dilution and how long it takes
to earn back, the pro forma CET1 ratio, and only then EPS accretion. The model is laid out that way.

Every reference is built through a running-row helper rather than hand-counted, so a row inserted anywhere
cannot silently break a formula elsewhere.
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"

# --------------------------------------------------------------------------- #
# Source data — FY2025 audited accounts (year to 31 December 2025), AED millions.
# Yahoo Finance fundamentals for EMIRATESNBD.AE and MASQ.AE, restating the published accounts.
# Share prices: DFM close, 12 September 2026.
# --------------------------------------------------------------------------- #
ACQ = dict(name="Emirates NBD", ticker="DFM: EMIRATESNBD", price=31.10, shares=6311.0,
           ni=23442.0, equity=144582.0, goodwill=5516.0, intang=104.0, assets=1164442.0,
           opex=15035.0, revenue=49177.0, pretax=29838.0, tax=5831.0)
TGT = dict(name="Mashreqbank", ticker="DFM: MASQ", price=334.90, shares=200.60983,
           ni=6615.0, equity=39374.0, goodwill=0.0, intang=423.0, assets=334633.9,
           opex=3871.0, revenue=12594.0, pretax=8261.0, tax=1291.0)

book = Book(theme=THEMES["enbd"], project_no=6, project="Merger Model — Accretion/Dilution, Synergies and Pro Forma",
            company="Emirates NBD / Mashreqbank", units="AED millions · per-share values in AED",
            as_of="FY2025 accounts · market data 12 Sep 2026")


class Sheet:
    """A sheet with a running row cursor and a label -> absolute-reference map."""

    def __init__(self, name, ncols, label_width=52, col_width=15, subtitle=None, valcol=3):
        self.ws = book.sheet(name, ncols=ncols, label_width=label_width, col_width=col_width, subtitle=subtitle)
        self.name, self.ncols, self.valcol, self.r = name, ncols, valcol, 4
        self.ref = {}

    def section(self, text):
        book.section(self.ws, self.r, text, self.ncols); self.r += 1

    def blank(self, n=1):
        self.r += n

    def head(self, labels, first_col=None):
        book.year_header(self.ws, self.r, labels, first_col=first_col or self.valcol)
        self.r += 1

    def row(self, label, value, fmt=NUM, font=None, note="", bold=False, border=None, col=None):
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

    def multi(self, label, values, fmt=NUM, font=None, bold=False, border=None, first_col=None):
        """One label, several columns. Returns list of absolute refs."""
        ws, r = self.ws, self.r
        c = ws.cell(r, 2, label); c.font = F_BOLD if bold else F_TEXT
        fc = first_col or self.valcol
        refs = []
        for i, value in enumerate(values):
            v = ws.cell(r, fc + i, value)
            f = font
            if f is None:
                f = F_FORMULA if (isinstance(value, str) and value.startswith("=")) else F_INPUT
            v.font = f; v.number_format = fmt
            if border: v.border = border
            refs.append(f"'{self.name}'!${L(fc+i)}${r}")
        self.ref[label] = refs
        self.r += 1
        return refs

    def note(self, text, height=30):
        book.note(self.ws, self.r, text, self.ncols, height=height); self.r += 1


# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
si = Sheet("Inputs", ncols=6, label_width=56, subtitle="Standalone financials, deal terms, synergies and purchase-accounting assumptions")
si.ws.column_dimensions["F"].width = 60
si.ws.column_dimensions["D"].width = 15

si.section("STANDALONE FINANCIALS — FY2025A (AED m)")
si.head(["Emirates NBD", "Mashreqbank"])
I = {}
for label, key, fmt in [("Net income attributable to shareholders", "ni", NUM0),
                        ("Shares outstanding (m)", "shares", NUM2),
                        ("Total shareholders' equity", "equity", NUM0),
                        ("Goodwill", "goodwill", NUM0),
                        ("Other intangible assets", "intang", NUM0),
                        ("Total assets", "assets", NUM0),
                        ("Operating expenses", "opex", NUM0),
                        ("Total revenue", "revenue", NUM0),
                        ("Share price (AED, 12 Sep 2026)", "price", NUM2)]:
    I[label] = si.multi(label, [ACQ[key], TGT[key]], fmt=fmt)

si.blank()
si.section("DERIVED STANDALONE METRICS")
for label, fml, fmt in [
    ("Tangible common equity", "={eq}-{gw}-{it}", NUM0),
    ("Market capitalisation", "={px}*{sh}", NUM0),
    ("Earnings per share (AED)", "={ni}/{sh}", NUM2),
    ("Tangible book value per share (AED)", "=TCE/{sh}", NUM2),
    ("Price / earnings", "={px}/EPS", MULT),
    ("Price / tangible book", "={px}/TBVPS", MULT),
]:
    vals = []
    for i in range(2):
        eq, gw, it = I["Total shareholders' equity"][i], I["Goodwill"][i], I["Other intangible assets"][i]
        px, sh, ni = I["Share price (AED, 12 Sep 2026)"][i], I["Shares outstanding (m)"][i], I["Net income attributable to shareholders"][i]
        f = fml.format(eq=eq, gw=gw, it=it, px=px, sh=sh, ni=ni)
        f = f.replace("TCE", si.ref.get("Tangible common equity", ["", ""])[i] if "Tangible common equity" in si.ref else "")
        f = f.replace("EPS", si.ref.get("Earnings per share (AED)", ["", ""])[i] if "Earnings per share (AED)" in si.ref else "")
        f = f.replace("TBVPS", si.ref.get("Tangible book value per share (AED)", ["", ""])[i] if "Tangible book value per share (AED)" in si.ref else "")
        vals.append(f)
    I[label] = si.multi(label, vals, fmt=fmt, font=F_FORMULA, bold=label.startswith("Tangible book"))

si.blank()
si.section("DEAL TERMS")
D = {}
D["prem"] = si.row("Offer premium to unaffected share price", 0.25, PCT,
                   note="Mid-range for a large European or Gulf bank deal. The sensitivity grid tests 10% to 40%")
D["stockpct"] = si.row("Consideration in acquirer stock", 0.70, PCT,
                       note="Stock-heavy, because a cash-heavy deal of this size would breach the CET1 floor")
D["cashpct"] = si.row("Consideration in cash", "=1-" + D["prem"].replace("$C$", "$C$").split("!")[1].join(["'Inputs'!", ""]) if False else "=1-" + D["stockpct"].split("!")[1], PCT, font=F_FORMULA,
                      note="Funded from surplus capital and wholesale issuance")
D["costcash"] = si.row("Pre-tax opportunity cost of cash consideration", 0.040, PCT,
                       note="Yield foregone on the cash paid away; roughly the bank's marginal funding cost")
D["tax"] = si.row("Marginal tax rate on synergies and adjustments", 0.09, PCT,
                  note="UAE federal corporate tax, in force since June 2023")

si.blank()
si.section("SYNERGIES AND INTEGRATION")
D["syn"] = si.row("Run-rate cost synergies as % of target operating expenses", 0.25, PCT,
                  note="Branch and head-office overlap in a single emirate. Ambitious but within the range disclosed in comparable bank mergers")
D["ph1"] = si.row("Synergy phasing — year 1", 0.25, PCT)
D["ph2"] = si.row("Synergy phasing — year 2", 0.75, PCT)
D["ph3"] = si.row("Synergy phasing — year 3", 1.00, PCT)
D["integ"] = si.row("One-time integration costs as a multiple of run-rate synergies", 1.50, MULT,
                    note="Charged in year 1. Systems migration, redundancy and rebranding")
D["revsyn"] = si.row("Revenue synergies", 0.0, NUM0,
                     note="Set to zero on purpose. Revenue synergies in bank mergers are routinely promised and rarely delivered; the notes explain the choice")

si.blank()
si.section("PURCHASE ACCOUNTING")
D["loanpct"] = si.row("Target gross loans as % of total assets", 0.55, PCT, note="Approximation pending the disclosed loan book")
D["creditmark"] = si.row("Credit mark on the acquired loan book", 0.0125, PCT,
                         note="Fair-value write-down for expected credit losses not yet provisioned. The single largest judgement in the model")
D["deppct"] = si.row("Target core deposits as % of total assets", 0.65, PCT)
D["cdi"] = si.row("Core deposit intangible as % of core deposits", 0.010, PCT,
                  note="Value of a cheap, sticky funding base; recognised as an identifiable intangible")
D["cdiyrs"] = si.row("Core deposit intangible amortisation period (years)", 8, '0')

si.blank()
si.section("CAPITAL")
D["rwa_a"] = si.row("Acquirer risk-weighted assets as % of total assets", 0.62, PCT)
D["rwa_t"] = si.row("Target risk-weighted assets as % of total assets", 0.60, PCT)
D["cet1_a"] = si.row("Acquirer CET1 ratio, standalone", 0.150, PCT, note="Reported Basel III common equity tier 1 ratio")
D["cet1_min"] = si.row("CET1 regulatory minimum including buffers", 0.110, PCT, note="UAE Central Bank minimum plus conservation and D-SIB buffers")
D["hurdle"] = si.row("Required return on invested capital (hurdle)", 0.120, PCT, note="Approximate cost of equity for a Gulf bank; the test the deal must clear")

si.blank()
si.note("Every figure on this sheet is an input (blue) except the derived metrics block. Everything on every other sheet "
        "is a formula. The two banks' FY2025 accounts are the only historical data in the model.", height=28)

ACQ_I, TGT_I = 0, 1

# --------------------------------------------------------------------------- #
# Purchase Price and allocation
# --------------------------------------------------------------------------- #
sp = Sheet("Purchase Price", ncols=6, label_width=58, subtitle="Offer, consideration mix and the allocation of the price paid")
sp.ws.column_dimensions["F"].width = 58

sp.section("OFFER")
P = {}
P["offer"] = sp.row("Offer price per Mashreqbank share (AED)",
                    f"={I['Share price (AED, 12 Sep 2026)'][TGT_I]}*(1+{D['prem']})", NUM2, bold=True)
P["pp"] = sp.row("Equity purchase price", f"={P['offer']}*{I['Shares outstanding (m)'][TGT_I]}", NUM0, bold=True)
P["exch"] = sp.row("Exchange ratio (acquirer shares per target share)",
                   f"={P['offer']}*{D['stockpct']}/{I['Share price (AED, 12 Sep 2026)'][ACQ_I]}", '0.000',
                   note="What a Mashreqbank holder receives in Emirates NBD stock")
P["stockval"] = sp.row("Consideration in stock", f"={P['pp']}*{D['stockpct']}", NUM0)
P["cashval"] = sp.row("Consideration in cash", f"={P['pp']}*{D['cashpct']}", NUM0)
P["newsh"] = sp.row("New acquirer shares issued (m)",
                    f"={P['stockval']}/{I['Share price (AED, 12 Sep 2026)'][ACQ_I]}", NUM2)
P["pfsh"] = sp.row("Pro forma shares outstanding (m)",
                   f"={I['Shares outstanding (m)'][ACQ_I]}+{P['newsh']}", NUM2, bold=True)
P["ownership"] = sp.row("Target shareholders' ownership of the combined bank",
                        f"={P['newsh']}/{P['pfsh']}", PCT,
                        note="At this mix the acquired shareholders own roughly a quarter of the group, which is a governance question as much as a financial one")
sp.blank()
P["pe_paid"] = sp.row("Price / earnings paid", f"={P['pp']}/{I['Net income attributable to shareholders'][TGT_I]}", MULT)
P["ptbv_paid"] = sp.row("Price / tangible book paid", f"={P['pp']}/{I['Tangible common equity'][TGT_I]}", MULT, bold=True,
                        note="Compare with the acquirer's own price/tangible book on the Inputs sheet. Issuing stock below the multiple being paid transfers value to the seller")

sp.blank()
sp.section("FAIR VALUE OF IDENTIFIABLE NET ASSETS ACQUIRED")
P["tce_t"] = sp.row("Target tangible common equity (book)", f"={I['Tangible common equity'][TGT_I]}", NUM0)
P["cm"] = sp.row("Less: credit mark on the loan book",
                 f"=-{I['Total assets'][TGT_I]}*{D['loanpct']}*{D['creditmark']}", NUM0)
P["cdi_amt"] = sp.row("Plus: core deposit intangible recognised",
                      f"={I['Total assets'][TGT_I]}*{D['deppct']}*{D['cdi']}", NUM0)
P["dt"] = sp.row("Deferred tax on fair-value adjustments",
                 f"=-({P['cm']}+{P['cdi_amt']})*{D['tax']}", NUM0)
P["fvna"] = sp.row("Fair value of identifiable net assets", f"={P['tce_t']}+{P['cm']}+{P['cdi_amt']}+{P['dt']}", NUM0,
                   bold=True, border=TOTAL_BORDER)
P["gw"] = sp.row("Goodwill created", f"={P['pp']}-{P['fvna']}", NUM0, bold=True, border=DOUBLE_BORDER,
                 note="The part of the price with no identifiable asset behind it. It is deducted in full from regulatory capital")
P["gwpct"] = sp.row("Goodwill as % of the price paid", f"={P['gw']}/{P['pp']}", PCT)
sp.blank()
P["cdi_am"] = sp.row("Annual core deposit intangible amortisation", f"={P['cdi_amt']}/{D['cdiyrs']}", NUM0)
P["synrun"] = sp.row("Run-rate pre-tax cost synergies", f"={I['Operating expenses'][TGT_I]}*{D['syn']}", NUM0)
P["integ_amt"] = sp.row("One-time integration costs", f"={P['synrun']}*{D['integ']}", NUM0)

sp.blank()
sp.note("The credit mark and the core deposit intangible largely offset each other here, so goodwill is close to the "
        "premium over tangible book. That is the honest shape of a bank deal: almost everything paid above tangible "
        "book becomes goodwill, and goodwill is worthless for regulatory capital purposes.", height=40)

# --------------------------------------------------------------------------- #
# Accretion / dilution
# --------------------------------------------------------------------------- #
sa = Sheet("Accretion", ncols=7, label_width=56, col_width=15,
           subtitle="Pro forma earnings bridge · AED millions except per-share")
sa.ws.column_dimensions["G"].width = 46
YR = ["Year 1", "Year 2", "Year 3"]
sa.section("PRO FORMA NET INCOME")
sa.head(YR)
ph = [D["ph1"], D["ph2"], D["ph3"]]
A = {}
A["acq"] = sa.multi("Acquirer net income (held flat)", [f"={I['Net income attributable to shareholders'][ACQ_I]}"] * 3, NUM0, font=F_FORMULA)
A["tgt"] = sa.multi("Target net income (held flat)", [f"={I['Net income attributable to shareholders'][TGT_I]}"] * 3, NUM0, font=F_FORMULA)
A["syn"] = sa.multi("Plus: after-tax cost synergies",
                    [f"={P['synrun']}*{p}*(1-{D['tax']})" for p in ph], NUM0, font=F_FORMULA)
A["cdi"] = sa.multi("Less: core deposit intangible amortisation, after tax",
                    [f"=-{P['cdi_am']}*(1-{D['tax']})"] * 3, NUM0, font=F_FORMULA)
A["fc"] = sa.multi("Less: opportunity cost of cash consideration, after tax",
                   [f"=-{P['cashval']}*{D['costcash']}*(1-{D['tax']})"] * 3, NUM0, font=F_FORMULA)
A["int"] = sa.multi("Less: one-time integration costs, after tax",
                    [f"=-{P['integ_amt']}*(1-{D['tax']})", "=0", "=0"], NUM0, font=F_FORMULA)
A["pfni"] = sa.multi("Pro forma net income", [f"=SUM({L(3+i)}{sa.r-6}:{L(3+i)}{sa.r-1})" for i in range(3)],
                     NUM0, font=F_FORMULA, bold=True, border=TOTAL_BORDER)
sa.blank()
sa.section("PER SHARE")
A["pfsh"] = sa.multi("Pro forma shares (m)", [f"={P['pfsh']}"] * 3, NUM2, font=F_LINK)
A["pfeps"] = sa.multi("Pro forma earnings per share (AED)",
                      [f"={A['pfni'][i]}/{A['pfsh'][i]}" for i in range(3)], NUM2, font=F_FORMULA, bold=True)
A["sa_eps"] = sa.multi("Standalone acquirer earnings per share (AED)",
                       [f"={I['Earnings per share (AED)'][ACQ_I]}"] * 3, NUM2, font=F_LINK)
A["accr"] = sa.multi("Accretion / (dilution)", [f"={A['pfeps'][i]}/{A['sa_eps'][i]}-1" for i in range(3)],
                     PCT, font=F_FORMULA, bold=True, border=DOUBLE_BORDER)
A["accr_aed"] = sa.multi("Accretion / (dilution) per share (AED)",
                         [f"={A['pfeps'][i]}-{A['sa_eps'][i]}" for i in range(3)], NUM2, font=F_FORMULA)

sa.blank()
sa.section("WHAT WOULD HAVE TO BE TRUE FOR THE DEAL TO BREAK EVEN ON EPS")
A["be_ni"] = sa.row("Pro forma net income required for EPS neutrality (steady state)",
                    f"={I['Earnings per share (AED)'][ACQ_I]}*{P['pfsh']}", NUM0, font=F_FORMULA)
A["be_gap"] = sa.row("Shortfall at full run-rate synergies",
                     f"={A['be_ni']}-{A['pfni'][2]}", NUM0, font=F_FORMULA)
A["be_syn"] = sa.row("Cost synergies required, as % of target operating expenses",
                     f"=({A['be_ni']}-({A['acq'][0]}+{A['tgt'][0]}+{A['cdi'][0]}+{A['fc'][0]}))"
                     f"/(1-{D['tax']})/{I['Operating expenses'][TGT_I]}", PCT, font=F_FORMULA, bold=True,
                     note="Against the 25% assumed. If this number is not achievable, the deal is not EPS-neutral at any credible synergy level")
sa.blank()
sa.note("Acquirer and target earnings are held flat at FY2025. A merger model that also forecasts both banks' "
        "organic growth mixes two questions together: whether the businesses grow, and whether the transaction adds "
        "value. Holding both flat isolates the transaction, which is the question the model exists to answer.", height=40)

# --------------------------------------------------------------------------- #
# Capital and tangible book
# --------------------------------------------------------------------------- #
sc = Sheet("Capital & TBV", ncols=6, label_width=58, subtitle="Tangible book dilution, earnback, and the regulatory capital constraint")
sc.ws.column_dimensions["F"].width = 56

sc.section("TANGIBLE COMMON EQUITY BRIDGE")
B = {}
B["tce_a"] = sc.row("Acquirer tangible common equity, standalone", f"={I['Tangible common equity'][ACQ_I]}", NUM0)
B["stock"] = sc.row("Plus: stock issued as consideration", f"={P['stockval']}", NUM0)
B["gw"] = sc.row("Less: goodwill created", f"=-{P['gw']}", NUM0)
B["cdi"] = sc.row("Less: core deposit intangible created", f"=-{P['cdi_amt']}", NUM0)
B["tce_pf"] = sc.row("Pro forma tangible common equity", f"={B['tce_a']}+{B['stock']}+{B['gw']}+{B['cdi']}", NUM0,
                     bold=True, border=TOTAL_BORDER,
                     note="The target's tangible net assets are exactly offset by the price paid for them, so the bridge moves only by stock issued less intangibles created")
sc.blank()
B["tbv_a"] = sc.row("Acquirer tangible book value per share, standalone (AED)",
                    f"={I['Tangible book value per share (AED)'][ACQ_I]}", NUM2)
B["tbv_pf"] = sc.row("Pro forma tangible book value per share (AED)", f"={B['tce_pf']}/{P['pfsh']}", NUM2, bold=True)
B["tbv_dil"] = sc.row("Tangible book value per share dilution", f"={B['tbv_pf']}/{B['tbv_a']}-1", PCT, bold=True,
                      border=DOUBLE_BORDER,
                      note="The number a bank's shareholders react to first. Anything beyond roughly 5% needs a short earnback to be defensible")
B["tbv_aed"] = sc.row("Dilution per share (AED)", f"={B['tbv_a']}-{B['tbv_pf']}", NUM2)
B["earnback"] = sc.row("Tangible book earnback (years)",
                       f'=IF({A["accr_aed"][2]}<=0,"Never — the deal is dilutive to EPS at full synergies",'
                       f'{B["tbv_aed"]}/{A["accr_aed"][2]})', '0.0', font=F_FORMULA, bold=True,
                       note="Dilution per share divided by steady-state accretion per share. If earnings never exceed standalone, the book value is never earned back")

sc.blank()
sc.section("REGULATORY CAPITAL")
B["rwa_a"] = sc.row("Acquirer risk-weighted assets", f"={I['Total assets'][ACQ_I]}*{D['rwa_a']}", NUM0)
B["rwa_t"] = sc.row("Target risk-weighted assets", f"={I['Total assets'][TGT_I]}*{D['rwa_t']}", NUM0)
B["rwa_pf"] = sc.row("Pro forma risk-weighted assets", f"={B['rwa_a']}+{B['rwa_t']}", NUM0, bold=True)
B["cet1_a"] = sc.row("Acquirer CET1 capital, standalone", f"={B['rwa_a']}*{D['cet1_a']}", NUM0)
B["dcet1"] = sc.row("Change in CET1 capital from the transaction",
                    f"={B['stock']}+{B['gw']}+{B['cdi']}", NUM0,
                    note="Stock issued less the goodwill and intangibles created, both of which are deducted in full from CET1")
B["cet1_pf_cap"] = sc.row("Pro forma CET1 capital", f"={B['cet1_a']}+{B['dcet1']}", NUM0)
B["cet1_pf"] = sc.row("Pro forma CET1 ratio", f"={B['cet1_pf_cap']}/{B['rwa_pf']}", PCT, bold=True, border=DOUBLE_BORDER)
B["cet1_delta"] = sc.row("Change versus standalone (percentage points)",
                         f"=({B['cet1_pf']}-{D['cet1_a']})*100", NUM2, font=F_FORMULA)
B["headroom"] = sc.row("Headroom above the regulatory minimum (percentage points)",
                       f"=({B['cet1_pf']}-{D['cet1_min']})*100", NUM2, font=F_FORMULA, bold=True,
                       note="Negative here would mean the deal cannot be done as structured at any price")

sc.blank()
sc.section("RETURN ON THE CAPITAL INVESTED")
B["roic"] = sc.row("Return on invested capital at full synergies",
                   f"=({I['Net income attributable to shareholders'][TGT_I]}+{P['synrun']}*(1-{D['tax']})"
                   f"-{P['cdi_am']}*(1-{D['tax']}))/{P['pp']}", PCT, font=F_FORMULA, bold=True)
B["spread"] = sc.row("Spread over the required return (percentage points)",
                     f"=({B['roic']}-{D['hurdle']})*100", NUM2, font=F_FORMULA, bold=True,
                     note="The cleanest single test of whether the price is justified, and the one least sensitive to the consideration mix")

# --------------------------------------------------------------------------- #
# Sensitivity
# --------------------------------------------------------------------------- #
ss = Sheet("Sensitivity", ncols=8, label_width=44, col_width=13,
           subtitle="Both grids are live: every cell rebuilds the offer, the share count and the purchase accounting from scratch")
PREMS = [0.10, 0.175, 0.25, 0.325, 0.40]
SYNS = [0.15, 0.20, 0.25, 0.30, 0.35]
MIXES = [0.50, 0.60, 0.70, 0.80, 0.90]

ss.section("EPS ACCRETION / (DILUTION) AT FULL RUN-RATE SYNERGIES")
ss.ws.cell(ss.r, 2, "Cost synergies (% of target opex)  \\  Offer premium").font = F_BOLD
for j, p in enumerate(PREMS):
    c = ss.ws.cell(ss.r, 3 + j, p); c.font = book.f_header; c.fill = book.fill_secondary
    c.number_format = PCT; c.alignment = Alignment(horizontal="center")
r_hdr1 = ss.r; ss.r += 1
r_g1 = ss.r
tgt_px, tgt_sh, acq_px, acq_sh = (I['Share price (AED, 12 Sep 2026)'][TGT_I], I['Shares outstanding (m)'][TGT_I],
                                  I['Share price (AED, 12 Sep 2026)'][ACQ_I], I['Shares outstanding (m)'][ACQ_I])
for i, s in enumerate(SYNS):
    c = ss.ws.cell(ss.r, 2, s); c.font = F_BOLD; c.number_format = PCT
    for j, p in enumerate(PREMS):
        prem = f"{L(3+j)}${r_hdr1}"; syn = f"$B{ss.r}"
        pp = f"({tgt_px}*(1+{prem})*{tgt_sh})"
        newsh = f"({pp}*{D['stockpct']}/{acq_px})"
        cash = f"({pp}*{D['cashpct']})"
        ni = (f"({I['Net income attributable to shareholders'][ACQ_I]}+{I['Net income attributable to shareholders'][TGT_I]}"
              f"+{I['Operating expenses'][TGT_I]}*{syn}*(1-{D['tax']})"
              f"-{P['cdi_am']}*(1-{D['tax']})-{cash}*{D['costcash']}*(1-{D['tax']}))")
        fml = f"={ni}/({acq_sh}+{newsh})/{I['Earnings per share (AED)'][ACQ_I]}-1"
        cc = ss.ws.cell(ss.r, 3 + j, fml); cc.font = F_FORMULA; cc.number_format = PCT
        if abs(p - 0.25) < 1e-9 and abs(s - 0.25) < 1e-9:
            cc.fill = book.fill_accent; cc.font = Font(name=FONT, size=10, bold=True)
    ss.r += 1
r_g1e = ss.r - 1
ss.blank()

ss.section("TANGIBLE BOOK VALUE PER SHARE DILUTION")
ss.ws.cell(ss.r, 2, "Consideration in stock  \\  Offer premium").font = F_BOLD
for j, p in enumerate(PREMS):
    c = ss.ws.cell(ss.r, 3 + j, p); c.font = book.f_header; c.fill = book.fill_secondary
    c.number_format = PCT; c.alignment = Alignment(horizontal="center")
r_hdr2 = ss.r; ss.r += 1
r_g2 = ss.r
for i, m in enumerate(MIXES):
    c = ss.ws.cell(ss.r, 2, m); c.font = F_BOLD; c.number_format = PCT
    for j, p in enumerate(PREMS):
        prem = f"{L(3+j)}${r_hdr2}"; mix = f"$B{ss.r}"
        pp = f"({tgt_px}*(1+{prem})*{tgt_sh})"
        stock = f"({pp}*{mix})"
        newsh = f"({stock}/{acq_px})"
        gw = f"({pp}-{P['fvna']})"
        tce = f"({I['Tangible common equity'][ACQ_I]}+{stock}-{gw}-{P['cdi_amt']})"
        fml = f"={tce}/({acq_sh}+{newsh})/{I['Tangible book value per share (AED)'][ACQ_I]}-1"
        cc = ss.ws.cell(ss.r, 3 + j, fml); cc.font = F_FORMULA; cc.number_format = PCT
        if abs(p - 0.25) < 1e-9 and abs(m - 0.70) < 1e-9:
            cc.fill = book.fill_accent; cc.font = Font(name=FONT, size=10, bold=True)
    ss.r += 1
r_g2e = ss.r - 1
ss.blank()
ss.note("Shaded cells are the base case and reconcile to the Accretion and Capital sheets. Read the second grid against "
        "the first: paying more in stock protects the CET1 ratio and worsens tangible book dilution, paying more in cash "
        "does the reverse. There is no mix that makes a price above two times tangible book look cheap.", height=40)

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sm = Sheet("Summary", ncols=6, label_width=58, subtitle="The transaction on one page")
sm.ws.column_dimensions["F"].width = 56
sm.section("THE TRANSACTION")
sm.row("Acquirer", f'="{ACQ["name"]}  ({ACQ["ticker"]})"', "General", font=F_TEXT)
sm.row("Target", f'="{TGT["name"]}  ({TGT["ticker"]})"', "General", font=F_TEXT)
sm.row("Offer price per share (AED)", f"={P['offer']}", NUM2, font=F_LINK)
sm.row("Premium to unaffected price", f"={D['prem']}", PCT, font=F_LINK)
sm.row("Equity purchase price (AED m)", f"={P['pp']}", NUM0, font=F_LINK)
sm.row("Consideration", f'=TEXT({D["stockpct"]},"0%")&" stock / "&TEXT({D["cashpct"]},"0%")&" cash"', "General", font=F_TEXT)
sm.row("Exchange ratio", f"={P['exch']}", '0.000', font=F_LINK)
sm.row("Target shareholders' ownership of the combined bank", f"={P['ownership']}", PCT, font=F_LINK)
sm.blank()
sm.section("THE THREE TESTS")
sm.row("1 · Tangible book value per share dilution", f"={B['tbv_dil']}", PCT, font=F_LINK, bold=True)
sm.row("      Earnback", f"={B['earnback']}", "General", font=F_LINK)
sm.row("2 · Pro forma CET1 ratio", f"={B['cet1_pf']}", PCT, font=F_LINK, bold=True)
sm.row("      Headroom above the minimum (pp)", f"={B['headroom']}", NUM2, font=F_LINK)
sm.row("3 · EPS accretion at full run-rate synergies", f"={A['accr'][2]}", PCT, font=F_LINK, bold=True)
sm.row("      Synergies required merely to break even", f"={A['be_syn']}", PCT, font=F_LINK)
sm.blank()
sm.row("Return on invested capital", f"={B['roic']}", PCT, font=F_LINK, bold=True)
sm.row("Spread over the 12% required return (pp)", f"={B['spread']}", NUM2, font=F_LINK, bold=True,
       border=DOUBLE_BORDER)
sm.blank()
sm.section("VERDICT")
for t in ["Emirates NBD trades at roughly 1.4 times tangible book. Mashreqbank trades at roughly 1.7 times, and a 25% "
          "premium takes the price paid above 2.1 times. An acquirer issuing its own stock at 1.4 times to buy assets at "
          "2.1 times transfers value to the seller on the day the deal is announced; no amount of synergy phasing undoes that.",
          "The arithmetic is not close. The cost synergies required just to hold earnings per share flat are far above the "
          "25% of target operating expenses assumed here, and above anything disclosed in comparable bank mergers.",
          "The capital ratio survives, which is why the consideration is stock-heavy. A cash-heavy structure would not.",
          "The useful output of a merger model is often the reason not to do the deal. This one says Emirates NBD would "
          "have to buy Mashreqbank close to its unaffected price, or in stock at a materially higher acquirer rating, for "
          "the transaction to make sense for its own shareholders."]:
    sm.note("•  " + t, height=15 * max(2, len(t) // 100 + 1))

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
sk = Sheet("Checks", ncols=5, label_width=72, col_width=16, subtitle="Every check must read TRUE")
sk.section("INTEGRITY CHECKS")
CHK = [
    ("Consideration percentages sum to 100%", f"=ABS({D['stockpct']}+{D['cashpct']}-1)<0.0001"),
    ("Stock plus cash consideration equals the equity purchase price",
     f"=ABS({P['stockval']}+{P['cashval']}-{P['pp']})<0.01"),
    ("Purchase price equals net assets acquired plus goodwill",
     f"=ABS({P['pp']}-({P['fvna']}+{P['gw']}))<0.01"),
    ("Goodwill is positive (the price exceeds fair value of net assets)", f"={P['gw']}>0"),
    ("Pro forma share count equals acquirer shares plus new shares",
     f"=ABS({P['pfsh']}-({I['Shares outstanding (m)'][ACQ_I]}+{P['newsh']}))<0.001"),
    ("Tangible equity bridge ties to stock issued less intangibles created",
     f"=ABS({B['tce_pf']}-({I['Tangible common equity'][ACQ_I]}+{P['stockval']}-{P['gw']}-{P['cdi_amt']}))<0.01"),
    ("Pro forma CET1 stays above the regulatory minimum", f"={B['cet1_pf']}>{D['cet1_min']}"),
    ("Synergy phasing is non-decreasing and ends at 100%",
     f"=AND({D['ph1']}<={D['ph2']},{D['ph2']}<={D['ph3']},{D['ph3']}=1)"),
    ("Year 3 pro forma net income exceeds year 1 (integration costs drop out)",
     f"={A['pfni'][2]}>{A['pfni'][0]}"),
    ("Base-case sensitivity cell reconciles to the Accretion sheet",
     f"=ABS(Sensitivity!E{r_g1+2}-{A['accr'][2]})<0.0005"),
    ("Base-case TBV sensitivity cell reconciles to the Capital sheet",
     f"=ABS(Sensitivity!E{r_g2+2}-{B['tbv_dil']})<0.0005"),
    ("Price paid per share exceeds the target's unaffected price",
     f"={P['offer']}>{I['Share price (AED, 12 Sep 2026)'][TGT_I]}"),
    ("Standalone EPS reconciles to reported diluted EPS within 1%",
     f"=AND(ABS({I['Earnings per share (AED)'][ACQ_I]}/3.71-1)<0.01,ABS({I['Earnings per share (AED)'][TGT_I]}/32.98-1)<0.01)"),
    ("Tangible common equity is below total equity for both banks",
     f"=AND({I['Tangible common equity'][ACQ_I]}<{I[chr(84)+'otal shareholders'+chr(39)+' equity'][ACQ_I]},"
     f"{I['Tangible common equity'][TGT_I]}<{I[chr(84)+'otal shareholders'+chr(39)+' equity'][TGT_I]})"),
]
for label, fml in CHK:
    c = sk.ws.cell(sk.r, 2, label); c.font = F_TEXT
    v = sk.ws.cell(sk.r, 3, fml); v.font = F_FORMULA; v.alignment = Alignment(horizontal="center")
    sk.r += 1
R_CK0, R_CK1 = 5, sk.r - 1
sk.blank()
c = sk.ws.cell(sk.r, 2, "MODEL STATUS"); c.font = Font(name=FONT, size=12, bold=True, color=book.theme.primary)
v = sk.ws.cell(sk.r, 3, f'=IF(COUNTIF(C{R_CK0}:C{R_CK1},FALSE)=0,"MODEL OK","CHECK FAILED")')
v.font = Font(name=FONT, size=12, bold=True); v.fill = book.fill_accent; v.alignment = Alignment(horizontal="center")
sk.r += 2
sk.note("The EPS reconciliation check is the one that matters most: it ties the model's computed earnings per share back "
        "to the diluted EPS each bank actually reported, so a mistyped share count or net income figure fails the model "
        "rather than quietly changing the answer.", height=34)

# --------------------------------------------------------------------------- #
# Cover, finish, export
# --------------------------------------------------------------------------- #
book.cover(
    blurb="A hypothetical acquisition of Mashreqbank by Emirates NBD, built on both banks' FY2025 audited accounts and "
          "12 September 2026 market prices. Neither bank is in discussions; the transaction is a modelling exercise, and "
          "the model's conclusion is that it should not be done at this price.",
    method=[
        "Purchase accounting in full: credit mark on the acquired loan book, core deposit intangible, deferred tax on "
        "both, and the goodwill that falls out of the difference.",
        "A three-year earnings bridge with synergies phased 25 / 75 / 100 and integration costs charged in year one, "
        "plus the synergy level that would be required merely to hold earnings per share flat.",
        "The two tests that actually decide bank deals: tangible book value per share dilution with its earnback period, "
        "and the pro forma CET1 ratio against the regulatory minimum.",
        "Two live sensitivity grids — premium against synergies for EPS, premium against consideration mix for tangible "
        "book — with every cell rebuilding the offer and the purchase accounting from scratch.",
    ],
    toc=[("Summary", "the transaction and the three tests on one page"),
         ("Purchase Price", "offer, consideration mix, allocation and goodwill"),
         ("Accretion", "three-year pro forma earnings bridge and the break-even synergy level"),
         ("Capital & TBV", "tangible book dilution, earnback, CET1 and return on invested capital"),
         ("Sensitivity", "EPS and tangible book across premium, synergies and mix"),
         ("Inputs", "standalone financials, deal terms and purchase-accounting assumptions"),
         ("Checks", "fourteen tests; must read MODEL OK")],
    highlights=[("Offer price per share (AED)", f"={P['offer']}", NUM2),
                ("Price / tangible book paid", f"={P['ptbv_paid']}", MULT),
                ("EPS accretion at full synergies", f"={A['accr'][2]}", PCT),
                ("Tangible book value dilution", f"={B['tbv_dil']}", PCT),
                ("Pro forma CET1 ratio", f"={B['cet1_pf']}", PCT),
                ("Return on invested capital", f"={B['roic']}", PCT)],
    sources=["Emirates NBD Bank PJSC and Mashreqbank PSC, FY2025 annual accounts (year to 31 December 2025), via the Yahoo Finance fundamentals series for EMIRATESNBD.AE and MASQ.AE.",
             "Share prices: Dubai Financial Market close, 12 September 2026.",
             "Tax rate: UAE federal corporate tax at 9%, in force since June 2023.",
             "Capital ratios and risk-weight assumptions are stated inputs on the Inputs sheet and should be replaced with each bank's Pillar 3 disclosures before any external use.",
             "The transaction is hypothetical. Nothing here is a suggestion that either bank is or has been in discussions."])

book.finish(freeze={"Accretion": "C6", "Inputs": "C5"})
path = os.path.join(HERE, "ENBD_Mashreq_Merger_Model.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
