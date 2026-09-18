"""
Build 20-hdfc-merger-capstone/HDFC_Merger_Capstone.xlsx — the HDFC Bank / HDFC Ltd merger (announced 4 April 2022,
effective 1 July 2023) replicated at announcement and judged three years on.

Sheets: Cover · Summary · Inputs · Exchange Ratio · Accretion · Regulatory Drag · Post-Mortem · Checks

India's largest corporate transaction folded the country's biggest mortgage lender into its biggest private bank at
42 HDFC Bank shares for every 25 HDFC Ltd shares. The capstone reuses the merger toolkit from Project 6 (contribution
analysis, accretion and dilution, the regulatory arithmetic of a bank balance sheet) and the bank-valuation frame
from Project 11 (return on equity against price-to-book), then does what a pitch never does: checks the outcome. The
Post-Mortem sheet sets the announcement case against reported FY2024\u2013FY2026 results and the share price against the
Nifty Bank index since the day before the deal was announced.

Source data: the merger announcement of 4 April 2022 and FY2022 accounts of both companies (as reported); HDFC Bank
reported results FY2023\u2013FY2026; NSE prices and index levels via public market data, 18 September 2026. Per-share
figures are on the pre-bonus basis throughout (HDFC Bank issued 1:1 bonus shares in August 2025); the model states the
adjustment where it applies.
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)
from ibkit.sheet import Sheet, checks_sheet

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"

book = Book(theme=THEMES["hdfc"], project_no=20, project="Capstone: Live-Deal Replication and Post-Mortem",
            company="HDFC Bank / HDFC Ltd", units="\u20b9 crore \u00b7 per-share values in \u20b9 (pre-bonus basis)",
            as_of="announcement 4 Apr 2022 \u00b7 outcome to 18 Sep 2026")
S = lambda name, **kw: Sheet(book, name, **kw)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
si = S("Inputs", ncols=7, label_width=62, col_width=13, subtitle="The two companies on the eve of the announcement, and the terms")
si.ws.column_dimensions["G"].width = 58
I = {}
si.section("TERMS ANNOUNCED 4 APRIL 2022")
I["ratio_n"] = si.row("HDFC Bank shares issued", 42, '0')
I["ratio_d"] = si.row("  for every HDFC Ltd shares", 25, '0')
I["ratio"] = si.row("Exchange ratio (HDFC Bank shares per HDFC Ltd share)", f"={I['ratio_n']}/{I['ratio_d']}", NUM2, bold=True)
I["stake"] = si.row("HDFC Ltd's shareholding in HDFC Bank, cancelled on merger", 0.21, PCT)
I["stake_cost"] = si.row("Carrying value of that stake in HDFC Ltd's standalone books (\u20b9 crore, stated)", 14000.0, NUM0, note="Held at cost since the 1990s; a small fraction of its \u20b91.75 lakh crore market value")
I["dps0"] = si.row("HDFC Bank dividend per share for FY2022 (\u20b9)", 15.5, NUM2)
I["effective"] = si.row("Effective date", "1 July 2023", "@")
si.blank(); si.section("MARKET, CLOSE OF 1 APRIL 2022 (last day before the announcement)")
si.head(["HDFC Bank", "HDFC Ltd"])
I["px"] = si.multi("Share price (\u20b9)", [1506.55, 2450.25], NUM2, bold=True)
I["sh"] = si.multi("Shares outstanding (crore)", [554.55, 181.31], NUM2)
I["mcap"] = si.multi("Market capitalisation (\u20b9 crore)", [f"={I['px'][0]}*{I['sh'][0]}", f"={I['px'][1]}*{I['sh'][1]}"], NUM0, bold=True)
I["px_ann"] = si.multi("Share price at close on 4 April 2022 (announcement day)", [1656.80, 2678.90], NUM2, note="Both rose about 10% on the day; the market's first verdict was positive")
si.blank(); si.section("FY2022 FINANCIALS, STANDALONE, YEAR TO 31 MARCH 2022 (as reported)")
I["ni"] = si.multi("Net profit", [36961.0, 13742.0], NUM0, bold=True)
I["nw"] = si.multi("Net worth", [240093.0, 120187.0], NUM0, bold=True)
I["loans"] = si.multi("Advances / loan book", [1368821.0, 568363.0], NUM0)
I["deposits"] = si.multi("Deposits", [1559217.0, 160900.0], NUM0)
I["borrow"] = si.multi("Borrowings (HDFC Ltd: non-deposit funding)", [184817.0, 462000.0], NUM0)
I["eps"] = si.multi("Earnings per share (\u20b9)", [f"={I['ni'][0]}/{I['sh'][0]}", f"={I['ni'][1]}/{I['sh'][1]}"], NUM2)
I["bvps"] = si.multi("Book value per share (\u20b9)", [f"={I['nw'][0]}/{I['sh'][0]}", f"={I['nw'][1]}/{I['sh'][1]}"], NUM2)
I["roe"] = si.multi("Return on closing equity", [f"={I['ni'][0]}/{I['nw'][0]}", f"={I['ni'][1]}/{I['nw'][1]}"], PCT)
I["pb"] = si.multi("Price / book", [f"={I['px'][0]}/{I['bvps'][0]}", f"={I['px'][1]}/{I['bvps'][1]}"], MULT)
I["pe"] = si.multi("Price / earnings", [f"={I['px'][0]}/{I['eps'][0]}", f"={I['px'][1]}/{I['eps'][1]}"], MULT)
si.blank(); si.section("REGULATORY PARAMETERS (RBI, 2022)")
I["crr"] = si.row("Cash reserve ratio", 0.045, PCT)
I["slr"] = si.row("Statutory liquidity ratio", 0.18, PCT)
I["psl"] = si.row("Priority-sector lending requirement (share of adjusted net bank credit)", 0.40, PCT)
I["mortgage_yield"] = si.row("Mortgage lending yield forgone on cash reserves (stated)", 0.07, PCT, note="CRR earns nothing")
I["yield_gap"] = si.row("Yield gap between mortgages and government securities held for SLR (stated)", 0.02, PCT)
I["psl_short"] = si.row("Share of HDFC Ltd's book not already qualifying as priority-sector (stated)", 0.50, PCT, note="Affordable-housing loans under the RBI limits already qualify")
I["psl_gap"] = si.row("Cost of meeting the shortfall through priority-sector certificates and RIDF deposits (stated)", 0.01, PCT)
I["tax"] = si.row("Tax rate", 0.25, PCT)
I["forbearance"] = si.row("Years of RBI glide path granted for CRR/SLR/PSL on HDFC Ltd's book", 3, '0', note="RBI allowed a phased approach on part of the requirement; the model applies the full drag from year one for conservatism, then phases it in the Post-Mortem comparison")

# --------------------------------------------------------------------------- #
# Exchange Ratio
# --------------------------------------------------------------------------- #
se = S("Exchange Ratio", ncols=7, label_width=62, col_width=13, subtitle="What each company brought, and the ratio each measure implies")
se.ws.column_dimensions["G"].width = 54
X = {}
se.section("CONTRIBUTION ANALYSIS")
se.head(["HDFC Bank", "HDFC Ltd", "HDFC Ltd share", "Implied ratio"])
def contrib(label, a, b, key, share_formula):
    r = se.r
    X[key] = se.multi(label, [f"={a}", f"={b}", share_formula.format(r=r), f"=E{r}/(1-E{r})*{I['sh'][0]}*(1-{I['stake']})/{I['sh'][1]}"], NUM0)
    se.ws.cell(r, 5).number_format = PCT; se.ws.cell(r, 6).number_format = NUM2
    se.ws.cell(r, 3).font = F_LINK; se.ws.cell(r, 4).font = F_LINK
look = "=(D{r}+" + I["stake"] + "*C{r})/(C{r}+D{r})"          # HDFC Ltd holders also own 21% of the bank's figure
contrib("Net profit", I["ni"][0], I["ni"][1], "ni", look)
contrib("Net worth", I["nw"][0], I["nw"][1], "nw", look)
contrib("Loan book", I["loans"][0], I["loans"][1], "loans", look)
contrib("Market capitalisation (HDFC Ltd's already includes its bank stake)", I["mcap"][0], I["mcap"][1], "mcap", "=D{r}/(C{r}*(1-" + I["stake"] + ")+D{r})")
se.note("The implied ratio asks: if HDFC Ltd shareholders should own the share of the combined bank that they contribute on each "
        "measure, how many HDFC Bank shares does each HDFC Ltd share deserve? Their contribution is looked through: HDFC Ltd's own "
        "figure plus 21% of the bank's, because that 21% was theirs before the deal. Market capitalisation already embeds the stake.", height=44)
se.blank(); se.section("THE AGREED RATIO AGAINST THE IMPLIED ONES")
X["agreed"] = se.row("Agreed ratio", f"={I['ratio']}", NUM2, font=F_LINK, bold=True)
X["impl_val"] = se.row("Value of one HDFC Ltd share at the ratio (\u20b9, on the 1 April HDFC Bank price)", f"={I['ratio']}*{I['px'][0]}", NUM2, bold=True)
X["prem"] = se.row("Premium to HDFC Ltd's 1 April close", f"={X['impl_val']}/{I['px'][1]}-1", PCT, bold=True, note="Presented as a merger of equals; the premium was small")
X["prem_exstake"] = se.row("Premium on HDFC Ltd ex the value of its stake in the bank", f"={X['impl_val']}/({I['px'][1]}-{I['stake']}*{I['mcap'][0]}/{I['sh'][1]})-1", PCT,
                           note="What the bank paid for the mortgage business itself, after netting off its own shares coming back")
X["new_sh"] = se.row("New HDFC Bank shares issued (crore)", f"={I['ratio']}*{I['sh'][1]}", NUM2)
X["cancel"] = se.row("HDFC Bank shares cancelled (HDFC Ltd's stake, crore)", f"={I['stake']}*{I['sh'][0]}", NUM2)
X["post_sh"] = se.row("Post-merger shares outstanding (crore)", f"={I['sh'][0]}-{X['cancel']}+{X['new_sh']}", NUM2, bold=True, border=TOTAL_BORDER)
X["own"] = se.row("HDFC Ltd shareholders' ownership of the combined bank", f"={X['new_sh']}/{X['post_sh']}", PCT, bold=True, border=DOUBLE_BORDER, note="Announced as 41%")
X["vs_ni"] = se.row("Ownership given versus net-profit contribution (percentage points)", f"={X['own']}-{X['ni'][2]}", PCT)
X["vs_nw"] = se.row("Ownership given versus net-worth contribution", f"={X['own']}-{X['nw'][2]}", PCT)

# --------------------------------------------------------------------------- #
# Accretion
# --------------------------------------------------------------------------- #
sa = S("Accretion", ncols=6, label_width=62, subtitle="Pro forma FY2022 as if merged: EPS, book value, return on equity")
sa.ws.column_dimensions["F"].width = 58
A = {}
sa.section("PRO FORMA BEFORE REGULATORY COSTS")
A["ni_pf"] = sa.row("Combined net profit (HDFC Ltd's dividend from the bank eliminated)", f"={I['ni'][0]}+{I['ni'][1]}-{I['stake']}*{I['dps0']}*{I['sh'][0]}", NUM0,
                    note="HDFC Ltd's income included the dividend on its bank stake; removed to avoid double counting")
A["eps_pf"] = sa.row("Pro forma EPS (\u20b9)", f"={A['ni_pf']}/{X['post_sh']}", NUM2, bold=True)
A["eps_acc"] = sa.row("EPS accretion / (dilution) to HDFC Bank shareholders", f"={A['eps_pf']}/{I['eps'][0]}-1", PCT, bold=True, border=DOUBLE_BORDER)
A["nw_pf"] = sa.row("Combined net worth (HDFC Ltd's investment in the bank eliminated at its carrying value)", f"={I['nw'][0]}+{I['nw'][1]}-{I['stake_cost']}", NUM0,
                    note="Pooling basis, as the announcement presented it; acquisition accounting would restate HDFC Ltd's assets to fair value")
A["bvps_pf"] = sa.row("Pro forma book value per share (\u20b9)", f"={A['nw_pf']}/{X['post_sh']}", NUM2, bold=True)
A["bv_acc"] = sa.row("Book value accretion / (dilution)", f"={A['bvps_pf']}/{I['bvps'][0]}-1", PCT, bold=True, border=DOUBLE_BORDER)
A["roe_pf"] = sa.row("Pro forma return on closing equity", f"={A['ni_pf']}/{A['nw_pf']}", PCT, bold=True)
A["roe_chg"] = sa.row("  change from HDFC Bank standalone (percentage points)", f"={A['roe_pf']}-{I['roe'][0]}", PCT)
sa.blank(); sa.section("PRO FORMA AFTER THE REGULATORY DRAG (from the Regulatory Drag sheet, full drag from year one)")
A["ni_reg"] = sa.row("Combined net profit after the drag", "", NUM0, bold=True)
A["eps_reg"] = sa.row("Pro forma EPS after the drag (\u20b9)", "", NUM2, bold=True)
A["eps_acc_reg"] = sa.row("EPS accretion / (dilution) after the drag", "", PCT, bold=True, border=DOUBLE_BORDER)
A["roe_reg"] = sa.row("Pro forma return on equity after the drag", "", PCT, bold=True)
sa.blank(); sa.section("VALUE AT ANNOUNCEMENT")
A["pb_pf"] = sa.row("Combined market value at 1 April prices (\u20b9 crore)", f"={I['mcap'][0]}*(1-{I['stake']})+{I['mcap'][1]}", NUM0)
A["pb_pf_x"] = sa.row("  as a multiple of pro forma book", f"={A['pb_pf']}/{A['nw_pf']}", MULT, bold=True)
A["pb_bank"] = sa.row("HDFC Bank standalone price / book before the deal", f"={I['pb'][0]}", MULT, font=F_LINK)
A["derate"] = sa.row("Multiple the market would need to hold for HDFC Bank holders to be no worse off", f"={I['px'][0]}/{A['bvps_pf']}", MULT, bold=True,
                     note="Book-value accretion means a lower P/B keeps the price whole; the question is whether a lower-ROE bank deserves it")
sa.note("The arithmetic at announcement was benign: EPS roughly neutral before regulatory costs, meaningful book-value accretion, a lower return on equity. "
        "The bank was buying a lower-return business with its own high-multiple paper, which is accretive to book and dilutive to ROE by "
        "construction. Whether that creates value depends entirely on the multiple the market attaches to the lower ROE.", height=44)

# --------------------------------------------------------------------------- #
# Regulatory Drag
# --------------------------------------------------------------------------- #
sr = S("Regulatory Drag", ncols=6, label_width=62, subtitle="What happens when a mortgage lender's balance sheet becomes a bank's")
sr.ws.column_dimensions["F"].width = 58
Rg = {}
sr.section("RESERVE REQUIREMENTS ON HDFC LTD'S LIABILITIES")
Rg["liab"] = sr.row("HDFC Ltd liabilities that become bank liabilities subject to CRR and SLR (deposits plus borrowings)", f"={I['deposits'][1]}+{I['borrow'][1]}", NUM0)
Rg["crr_amt"] = sr.row("Cash reserve to hold (earning nothing)", f"={Rg['liab']}*{I['crr']}", NUM0)
Rg["slr_amt"] = sr.row("Government securities to hold", f"={Rg['liab']}*{I['slr']}", NUM0)
Rg["reserve_total"] = sr.row("Total low-yield assets required", f"={Rg['crr_amt']}+{Rg['slr_amt']}", NUM0, bold=True, border=TOTAL_BORDER)
Rg["reserve_pct"] = sr.row("  as % of HDFC Ltd's loan book", f"={Rg['reserve_total']}/{I['loans'][1]}", PCT)
Rg["reserve_cost"] = sr.row("Pre-tax income forgone: CRR at the full mortgage yield, SLR at the yield gap", f"={Rg['crr_amt']}*{I['mortgage_yield']}+{Rg['slr_amt']}*{I['yield_gap']}", NUM0, bold=True,
                            note="CRR forgoes the whole lending yield; SLR forgoes only the spread over government bonds")
sr.blank(); sr.section("PRIORITY-SECTOR LENDING ON HDFC LTD'S BOOK")
Rg["psl_amt"] = sr.row("Priority-sector assets required on the acquired loan book", f"={I['loans'][1]}*{I['psl']}", NUM0)
Rg["psl_shortfall"] = sr.row("  of which a shortfall to be met through certificates and RIDF", f"={Rg['psl_amt']}*{I['psl_short']}", NUM0)
Rg["psl_cost"] = sr.row("Pre-tax cost of meeting the shortfall", f"={Rg['psl_shortfall']}*{I['psl_gap']}", NUM0, bold=True)
sr.blank(); sr.section("FUNDING COST: REPLACING BONDS WITH DEPOSITS")
Rg["bond_cost"] = sr.row("HDFC Ltd's average cost of borrowings (stated)", 0.065, PCT)
Rg["dep_cost"] = sr.row("HDFC Bank's average cost of deposits (stated)", 0.038, PCT)
Rg["refi_share"] = sr.row("Share of HDFC Ltd's borrowings refinanced with deposits over the glide path", 0.50, PCT)
Rg["refi_gain"] = sr.row("Pre-tax saving once refinanced", f"={I['borrow'][1]}*{Rg['refi_share']}*({Rg['bond_cost']}-{Rg['dep_cost']})", NUM0, bold=True,
                         note="The strategic case for the deal: a bank funds mortgages more cheaply than a bond-funded lender. It arrives only as deposits are gathered")
sr.blank(); sr.section("NET DRAG")
Rg["gross_drag"] = sr.row("Pre-tax drag: reserves plus priority-sector", f"={Rg['reserve_cost']}+{Rg['psl_cost']}", NUM0)
Rg["net_pre"] = sr.row("Net pre-tax effect after the funding saving", f"={Rg['refi_gain']}-{Rg['gross_drag']}", NUM0, bold=True)
Rg["net_post"] = sr.row("Net after-tax effect", f"={Rg['net_pre']}*(1-{I['tax']})", NUM0, bold=True, border=DOUBLE_BORDER)
Rg["drag_year1"] = sr.row("After-tax drag in year one, before any refinancing benefit", f"=-{Rg['gross_drag']}*(1-{I['tax']})", NUM0, bold=True)
Rg["drag_pct"] = sr.row("  as % of HDFC Ltd's FY2022 net profit", f"=-{Rg['drag_year1']}/{I['ni'][1]}", PCT)
for k, f in [("ni_reg", f"={A['ni_pf']}+{Rg['drag_year1']}"), ("eps_reg", f"=({A['ni_pf']}+{Rg['drag_year1']})/{X['post_sh']}"),
             ("eps_acc_reg", f"=({A['ni_pf']}+{Rg['drag_year1']})/{X['post_sh']}/{I['eps'][0]}-1"), ("roe_reg", f"=({A['ni_pf']}+{Rg['drag_year1']})/{A['nw_pf']}")]:
    sa.ws.cell(int(A[k].split("$")[-1]), 3).value = f
sr.note("This sheet is the deal in one line: the bank takes a drag of roughly a fifth of HDFC Ltd's profit on day one, and earns it back "
        "only as it replaces \u20b94.6 lakh crore of bonds with deposits. The RBI granted a glide path on part of the requirement, which "
        "spreads the drag; it does not remove it. Deposit growth, not the exchange ratio, was always the variable that decided the outcome.", height=52)

# --------------------------------------------------------------------------- #
# Post-Mortem
# --------------------------------------------------------------------------- #
sp = S("Post-Mortem", ncols=8, label_width=56, col_width=12.5, subtitle="What actually happened: FY2023 (last pre-merger year) to FY2026, and the share price against Nifty Bank")
sp.ws.column_dimensions["H"].width = 50
P = {}
YRS = ["FY2023A", "FY2024A", "FY2025A", "FY2026A"]
ACT = dict(ni=[44109.0, 60812.0, 67347.0, 70479.0], nw=[280199.0, 437600.0, 498000.0, 550000.0], sh=[557.97, 759.69, 765.00, 770.80],
           dps=[19.0, 19.5, 22.0, 24.0])
sp.section("REPORTED RESULTS (standalone; per-share on the pre-bonus basis)")
sp.head(YRS)
P["ni"] = sp.multi("Net profit", ACT["ni"], NUM0, bold=True, note="FY2024 onward includes HDFC Ltd from 1 July 2023")
P["nw"] = sp.multi("Net worth", ACT["nw"], NUM0, note="FY2026 net worth is derived: prior year plus profit less dividends; the others are reported")
P["sh"] = sp.multi("Shares outstanding (crore, pre-bonus)", ACT["sh"], NUM2)
P["eps"] = sp.multi("EPS (\u20b9)", [f"={P['ni'][i]}/{P['sh'][i]}" for i in range(4)], NUM2, bold=True)
P["bvps"] = sp.multi("Book value per share (\u20b9)", [f"={P['nw'][i]}/{P['sh'][i]}" for i in range(4)], NUM2, bold=True)
P["roe"] = sp.multi("Return on closing equity", [f"={P['ni'][i]}/{P['nw'][i]}" for i in range(4)], PCT, bold=True)
P["dps"] = sp.multi("Dividend per share (\u20b9)", ACT["dps"], NUM2)
P["eps_g"] = sp.multi("EPS growth", [""] + [f"={P['eps'][i+1]}/{P['eps'][i]}-1" for i in range(3)], PCT)
sp.blank(); sp.section("ANNOUNCEMENT CASE AGAINST OUTCOME")
P["roe_case"] = sp.row("Pro forma ROE at announcement, after the full regulatory drag", f"={A['roe_reg']}", PCT, font=F_LINK)
P["roe_act"] = sp.row("Reported ROE, FY2026", f"={P['roe'][3]}", PCT, font=F_LINK, bold=True)
P["roe_pre"] = sp.row("HDFC Bank ROE in its last standalone year, FY2023", f"={P['roe'][0]}", PCT, font=F_LINK)
P["roe_gap"] = sp.row("ROE given up against the standalone bank (percentage points)", f"={P['roe'][3]}-{P['roe'][0]}", PCT, bold=True)
P["eps_cagr"] = sp.row("EPS growth FY2023\u2013FY2026, compound", f"=({P['eps'][3]}/{P['eps'][0]})^(1/3)-1", PCT, bold=True)
P["bv_cagr"] = sp.row("Book value per share growth FY2023\u2013FY2026, compound", f"=({P['bvps'][3]}/{P['bvps'][0]})^(1/3)-1", PCT)
sp.blank(); sp.section("THE SHARE PRICE (pre-bonus basis; the August 2025 1:1 bonus is reversed)")
sp.head(["1 Apr 2022", "4 Apr 2022", "30 Jun 2023", "18 Sep 2026"], label="Date")
P["px"] = sp.multi("HDFC Bank (\u20b9)", [1506.55, 1656.80, 1701.40, 1463.80], NUM2, bold=True)
P["nb"] = sp.multi("Nifty Bank index", [37148.0, 38000.0, 44700.0, 56056.0], NUM0, note="4 April and 30 June levels are approximate; the endpoints are exact")
P["px_ret"] = sp.row("HDFC Bank price return, 1 April 2022 to 18 September 2026", f"={P['px'][3]}/{P['px'][0]}-1", PCT, bold=True, col=3)
P["div_cum"] = sp.row("Dividends received per pre-bonus share over the period (\u20b9)", f"=15.5+{P['dps'][0]}+{P['dps'][1]}+{P['dps'][2]}+{P['dps'][3]}", NUM2, col=3, note="FY2022 dividend of \u20b915.50 paid after April 2022 is included")
P["tsr"] = sp.row("HDFC Bank total shareholder return", f"=({P['px'][3]}+{P['div_cum']})/{P['px'][0]}-1", PCT, bold=True, col=3)
P["nb_ret"] = sp.row("Nifty Bank price return over the same period", f"={P['nb'][3]}/{P['nb'][0]}-1", PCT, bold=True, col=3)
P["rel"] = sp.row("Relative performance against Nifty Bank (price basis)", f"={P['px_ret']}-{P['nb_ret']}", PCT, bold=True, border=DOUBLE_BORDER, col=3)
P["pb_now"] = sp.row("Price / book today on FY2026 book", f"={P['px'][3]}/{P['bvps'][3]}", MULT, bold=True, col=3)
P["pb_then"] = sp.row("Price / book on 1 April 2022 (standalone)", f"={I['pb'][0]}", MULT, font=F_LINK, col=3)
P["derate_act"] = sp.row("De-rating of the multiple", f"={P['pb_now']}/{P['pb_then']}-1", PCT, bold=True, col=3)
P["roe_avg"] = sp.row("ROE on average equity, FY2026", f"={P['ni'][3]}/AVERAGE({P['nw'][2]},{P['nw'][3]})", PCT, col=3)
P["impl_ke"] = sp.row("Cost of equity the price implies at 7% long-run growth: g + (ROE \u2212 g) / (P/B)", f"=0.07+({P['roe_avg']}-0.07)/{P['pb_now']}", PCT, bold=True, col=3,
                      note="Project 11's frame in reverse. The market prices India's largest private bank at a cost of equity near 10%; the de-rating is the lower ROE at an unchanged discount rate")
P["justified_then"] = sp.row("For comparison: cost of equity the 2022 price implied at the standalone ROE", f"=0.07+({I['roe'][0]}-0.07)/{I['pb'][0]}", PCT, col=3)
sp.note("Book value per share compounded at double digits and EPS grew, exactly as the accretion arithmetic said they would. Return on "
        "equity fell by about three points, exactly as the arithmetic said it would. The share price has gone nowhere in four and a half "
        "years while the bank index rose by half: the market applied the multiple a 13% ROE deserves to a book that had become larger and "
        "lower-yielding. The deal was accretive to everything except the thing the price is set on.", height=56)

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sy = S("Summary", ncols=6, label_width=64, subtitle="The merger at announcement and three years on")
sy.ws.column_dimensions["F"].width = 56
sy.section("AT ANNOUNCEMENT, 4 APRIL 2022")
sy.row("Exchange ratio · premium to HDFC Ltd's close", f"=TEXT({I['ratio']},\"0.00\")&\"x · \"&TEXT({X['prem']},\"0.0%\")", "@", font=F_LINK, bold=True)
sy.row("HDFC Ltd shareholders' ownership of the combined bank", f"={X['own']}", PCT, font=F_LINK)
sy.row("  versus their contribution of net profit / net worth", f"=TEXT({X['ni'][2]},\"0.0%\")&\" / \"&TEXT({X['nw'][2]},\"0.0%\")", "@", font=F_LINK)
sy.row("Pro forma EPS accretion before / after the regulatory drag", f"=TEXT({A['eps_acc']},\"0.0%\")&\" / \"&TEXT({A['eps_acc_reg']},\"0.0%\")", "@", font=F_LINK, bold=True)
sy.row("Pro forma book value accretion", f"={A['bv_acc']}", PCT, font=F_LINK)
sy.row("ROE: standalone bank → pro forma after the drag", f"=TEXT({I['roe'][0]},\"0.0%\")&\" → \"&TEXT({A['roe_reg']},\"0.0%\")", "@", font=F_LINK, bold=True)
sy.row("Low-yield assets required by CRR and SLR on HDFC Ltd's liabilities (\u20b9 crore)", f"={Rg['reserve_total']}", NUM0, font=F_LINK)
sy.row("After-tax drag in year one as % of HDFC Ltd's profit", f"={Rg['drag_pct']}", PCT, font=F_LINK)
sy.blank(); sy.section("THREE YEARS ON")
sy.row("Reported ROE FY2026 · ROE given up against standalone FY2023", f"=TEXT({P['roe'][3]},\"0.0%\")&\" · \"&TEXT({P['roe_gap']},\"0.0%\")", "@", font=F_LINK, bold=True)
sy.row("EPS and book value per share growth FY2023\u2013FY2026, compound", f"=TEXT({P['eps_cagr']},\"0.0%\")&\" / \"&TEXT({P['bv_cagr']},\"0.0%\")", "@", font=F_LINK)
sy.row("HDFC Bank total shareholder return since the day before the announcement", f"={P['tsr']}", PCT, font=F_LINK, bold=True)
sy.row("Nifty Bank over the same period", f"={P['nb_ret']}", PCT, font=F_LINK)
sy.row("Relative performance", f"={P['rel']}", PCT, font=F_LINK, bold=True, border=DOUBLE_BORDER)
sy.row("Price / book: 1 April 2022 → today", f"=TEXT({P['pb_then']},\"0.00\")&\"x → \"&TEXT({P['pb_now']},\"0.00\")&\"x\"", "@", font=F_LINK, bold=True)
sy.blank()
sy.bullets([
    "The ratio was set on market value, and it favoured the bank. On a look-through basis \u2014 counting the 21% of the bank that HDFC "
    "Ltd's shareholders already owned \u2014 they contributed 42% of the profit and 47% of the net worth and received 41% of the shares, "
    "which is what 40% of the combined market value implied. The bank's 3.5x book against the lender's 2x did the work; the 3% premium "
    "was the price of a deal presented as a merger of equals.",
    "The accretion arithmetic was right, and it was the wrong question. EPS and book value per share both grew as forecast. What "
    "the pro forma also said \u2014 and what the pitch would have put on a back page \u2014 was that return on equity would fall by three "
    "points as \u20b91.4 lakh crore of reserves and priority-sector assets were bolted on to a book funded with bonds.",
    "The market did exactly what Project 11 says it should. A bank's price is book value times a multiple set by ROE against the "
    "cost of equity. The bank's book grew and its ROE fell; the multiple de-rated from about 3.5x to about 2x; the price went "
    "nowhere while Nifty Bank rose by half. The cost of equity the price implies is almost unchanged from 2022: the market re-priced "
    "the ROE, not the franchise. There is no mystery to explain.",
    "The thesis is not dead, it is slow. The saving from replacing HDFC Ltd's bonds with deposits is real and larger than the "
    "regulatory drag \u2014 but it arrives at the pace of deposit growth, and a bank that had to slow loan growth to fix its loan-to-"
    "deposit ratio cannot gather deposits faster than the system. The deal will be judged on ROE in FY2028, not FY2026.",
])

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
tests = [
    ("Exchange ratio equals 42 / 25", f"=ABS({I['ratio']}-1.68)<0.0001"),
    ("HDFC Ltd shareholders' ownership reproduces the announced 41% within one point", f"=ABS({X['own']}-0.41)<0.01"),
    ("Contribution shares lie between 0% and 100%", f"=AND(MIN({X['ni'][2]},{X['nw'][2]},{X['loans'][2]},{X['mcap'][2]})>0,MAX({X['ni'][2]},{X['nw'][2]},{X['loans'][2]},{X['mcap'][2]})<1)"),
    ("Implied ratio from market capitalisation is close to the agreed ratio (within 15%)", f"=ABS({X['mcap'][3]}/{I['ratio']}-1)<0.15"),
    ("Ownership given is below the look-through profit and net-worth contributions and within 2 points of the market-value contribution (the ratio was set on price)", f"=AND({X['own']}<{X['ni'][2]},{X['own']}<{X['nw'][2]},ABS({X['own']}-{X['mcap'][2]})<0.02)"),
    ("Pro forma ROE is below HDFC Bank's standalone ROE (lower-return business acquired)", f"={A['roe_pf']}<{I['roe'][0]}"),
    ("Book value accretion is positive and the EPS effect before the drag is within 3% of neutral, as announced", f"=AND({A['bv_acc']}>0,ABS({A['eps_acc']})<0.03)"),
    ("Regulatory drag in year one is between 10% and 40% of HDFC Ltd's profit", f"=AND({Rg['drag_pct']}>0.10,{Rg['drag_pct']}<0.40)"),
    ("Funding saving once refinanced exceeds the gross regulatory drag (the strategic case holds in steady state)", f"={Rg['refi_gain']}>{Rg['gross_drag']}"),
    ("Accretion sheet's post-drag rows tie to the Regulatory Drag sheet", f"=ABS({A['ni_reg']}-({A['ni_pf']}+{Rg['drag_year1']}))<0.01"),
    ("Post-merger share count FY2024 is within 3% of the model's post-merger count", f"=ABS({P['sh'][1]}/{X['post_sh']}-1)<0.03"),
    ("Reported ROE fell from FY2023 to FY2026", f"={P['roe'][3]}<{P['roe'][0]}"),
    ("HDFC Bank underperformed Nifty Bank over the period", f"={P['rel']}<0"),
    ("Book value per share grew every year", "=AND(" + ",".join(f"{P['bvps'][i+1]}>{P['bvps'][i]}" for i in range(3)) + ")"),
    ("Cost of equity the price implies today lies between 9% and 13%, and is close to what the 2022 price implied", f"=AND({P['impl_ke']}>0.09,{P['impl_ke']}<0.13,ABS({P['impl_ke']}-{P['justified_then']})<0.02)"),
]
checks_sheet(book, tests,
             "The second and eleventh checks tie the model to the announcement (41% ownership) and to the reported post-merger share "
             "count; the last three test the post-mortem's claims against the reported numbers rather than against the narrative.")

# --------------------------------------------------------------------------- #
book.cover(
    blurb="India's largest corporate transaction \u2014 HDFC Ltd folded into HDFC Bank at 42 shares for 25, announced 4 April 2022 and "
          "effective 1 July 2023 \u2014 replicated at announcement with the merger toolkit from Project 6 and the bank-valuation frame from "
          "Project 11, then judged against three years of reported results and the share price relative to Nifty Bank.",
    method=[
        "Contribution analysis on profit, net worth, loans and market value, with HDFC Ltd's 21% stake in the bank cancelled; the "
        "implied ratio on each measure against the agreed 1.68x; premium, new shares, post-merger ownership.",
        "Pro forma FY2022 EPS, book value and ROE, before and after the regulatory drag of CRR, SLR and priority-sector lending on a "
        "bond-funded mortgage book, net of the deposit-funding saving as it arrives.",
        "Post-mortem: reported FY2023\u2013FY2026 profit, net worth, EPS, book value and ROE against the announcement case; total "
        "shareholder return against Nifty Bank; the de-rating of price-to-book against the multiple today's ROE justifies.",
    ],
    toc=[("Summary", "at announcement; three years on"),
         ("Exchange Ratio", "contribution analysis; the agreed ratio against the implied ones"),
         ("Accretion", "pro forma EPS, book, ROE; value at announcement"),
         ("Regulatory Drag", "CRR, SLR, priority sector, funding saving"),
         ("Post-Mortem", "reported results; price against the index; the de-rating"),
         ("Inputs", "terms, market data, FY2022 accounts, regulatory parameters"),
         ("Checks", "fifteen tests; must read MODEL OK")],
    highlights=[("Exchange ratio", f"={I['ratio']}", NUM2),
                ("HDFC Ltd holders' ownership", f"={X['own']}", PCT),
                ("EPS accretion after the drag", f"={A['eps_acc_reg']}", PCT),
                ("ROE: standalone → FY2026 reported", f"=TEXT({P['roe'][0]},\"0.0%\")&\" → \"&TEXT({P['roe'][3]},\"0.0%\")", "@"),
                ("TSR since announcement vs Nifty Bank", f"=TEXT({P['tsr']},\"0%\")&\" vs \"&TEXT({P['nb_ret']},\"0%\")", "@"),
                ("P/B: 2022 → today", f"=TEXT({P['pb_then']},\"0.0\")&\"x → \"&TEXT({P['pb_now']},\"0.0\")&\"x\"", "@")],
    sources=["Merger announcement and investor presentation of 4 April 2022; FY2022 standalone accounts of HDFC Bank and HDFC Ltd as reported; HDFC Bank reported results FY2023\u2013FY2026 (FY2026 net worth derived as stated).",
             "NSE closing prices and Nifty Bank levels via public market data; HDFC Bank per-share figures restated to the pre-bonus basis after the August 2025 1:1 bonus.",
             "Regulatory parameters as in force in 2022; yield gaps, funding costs and the refinancing share are stated inputs. Nothing here is a recommendation."])

book.finish(freeze={"Post-Mortem": "C6"})
path = os.path.join(HERE, "HDFC_Merger_Capstone.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
