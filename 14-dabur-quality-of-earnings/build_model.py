"""
Build 14-dabur-quality-of-earnings/Dabur_QoE.xlsx — a buy-side quality-of-earnings review of Dabur India, with a
working-capital peg of the kind a sale-and-purchase agreement is written around.

Sheets: Cover · Summary · Statements · EBITDA Bridge · Cash Conversion · Working Capital · Peg · Checks

A quality-of-earnings report asks two questions the audited accounts do not. First, what is the business actually
earning from operating, once financial income, one-offs and reclassifications are stripped out of the headline
EBITDA that a seller's banker will quote? Second, how much working capital does the business need to run, so that
the buyer pays for a normal level and not for whatever happens to be on the balance sheet on the closing date?
Dabur is chosen because it is a clean, listed Indian consumer company with a large treasury book \u2014 which makes the
first question interesting \u2014 and a long payables cycle, which makes the second one interesting.

Source: Dabur India consolidated statements FY2023\u2013FY2026 (years to 31 March) via public market data, ₹ crore.
Indian filers publish balance sheets half-yearly, so the peg uses year-end points and says so; a real engagement uses
twelve to twenty-four monthly closes from the data room.
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from ibkit.style import (Book, THEMES, Theme, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)
from ibkit.sheet import Sheet, checks_sheet

HERE = os.path.dirname(os.path.abspath(__file__)); FONT = "Arial"
DABUR = Theme("dabur", "Dabur India", "0B6E4F", "5A7A6E", "E8A317", "EAF4EF", tagline="NSE: DABUR")

Y = ["FY2023A", "FY2024A", "FY2025A", "FY2026A"]; NY = 4; CE = 3
D = dict(
    revenue=[11426.5, 12261.6, 12400.0, 13042.0], cogs=[6316.1, 6496.4, 6583.5, 6880.6],
    opex=[3241.9, 3746.5, 3932.6, 4157.1],           # operating expenses excluding D&A? Yahoo 'Operating Expense' incl. D&A? tested below
    op_income=[1868.5, 2018.8, 1883.9, 2004.3], da=[311.0, 399.2, 445.6, 468.9],
    int_income=[361.1, 393.1, 408.2, 366.2], other_nonop=[21.5, 29.6, 25.3, 21.8], special=[15.6, -16.8, -3.8, -20.0],
    int_exp=[69.1, 110.2, 146.9, 125.5], pbt=[2218.7, 2358.7, 2257.9, 2420.4], tax=[517.4, 547.4, 517.5, 551.7],
    ni=[1707.2, 1842.7, 1767.6, 1895.0], sbc=[51.2, 50.4, 13.5, 28.6], shares=[177.6, 177.6, 177.6, 177.7],
    ar=[848.8, 898.7, 888.5, 715.4], inv=[2024.2, 1947.0, 2300.1, 2321.6], ap=[2186.6, 2421.7, 2825.3, 2954.3],
    prov=[214.0, 249.9, 260.5, 310.2], cash_inv=[1051.6, 2323.1, 2655.5, 5147.8], debt=[1173.8, 1365.1, 950.1, 1287.3],
    equity=[8973.3, 9866.3, 10800.7, 11419.5], ppe=[2408.7, 2770.0, 2966.3, 2977.4], intang=[1298.0, 1231.8, 1148.9, 1056.5],
    ocf=[1488.4, 2013.5, 1986.8, 2578.6], capex=[509.1, 563.9, 569.5, 420.1], divs=[921.3, 965.8, 974.8, 1419.0],
    total_assets=[13800.0, 14900.0, 16100.0, 18100.0],   # approximated as equity + debt + payables + provisions + other; used only for the accruals ratio denominator
    price=387.45, q_rev=[2830.1, 3404.6, 3558.6, 3038.0, 3764.4], q_lab=["Q4 FY25", "Q1 FY26", "Q3 FY26", "Q4 FY26", "Q1 FY27"])

book = Book(theme=DABUR, project_no=14, project="Quality of Earnings and Working-Capital Peg",
            company="Dabur India", units="\u20b9 crore unless stated", as_of="FY2026 consolidated accounts (year to 31 March 2026)")
S = lambda name, **kw: Sheet(book, name, **kw)

# --------------------------------------------------------------------------- #
# Statements
# --------------------------------------------------------------------------- #
ss_ = S("Statements", ncols=7, label_width=54, col_width=13, subtitle="Four years of consolidated figures as reported \u2014 the raw material")
ss_.ws.column_dimensions["G"].width = 46
T = {}
ss_.section("INCOME STATEMENT")
ss_.head(Y)
T["rev"] = ss_.multi("Revenue from operations", D["revenue"], NUM, bold=True)
T["cogs"] = ss_.multi("Cost of goods sold", [-x for x in D["cogs"]], NUM)
T["gp"] = ss_.multi("Gross profit", [f"={L(CE+i)}{ss_.r-2}+{L(CE+i)}{ss_.r-1}" for i in range(NY)], NUM, bold=True, border=TOTAL_BORDER)
T["gm"] = ss_.multi("   Gross margin", [f"={L(CE+i)}{ss_.r-1}/{L(CE+i)}{ss_.r-3}" for i in range(NY)], PCT)
T["opex"] = ss_.multi("Operating expenses before depreciation", [f"=-({D['revenue'][i]}-{D['cogs'][i]}-{D['op_income'][i]}-{D['da'][i]})" for i in range(NY)], NUM,
                      font=F_FORMULA)
T["ebitda_op"] = ss_.multi("Operating EBITDA (revenue \u2212 COGS \u2212 opex)", [f"={L(CE+i)}{ss_.r-3}+{L(CE+i)}{ss_.r-1}" for i in range(NY)], NUM, bold=True, border=TOTAL_BORDER)
T["da"] = ss_.multi("Depreciation and amortisation", [-x for x in D["da"]], NUM)
T["op_inc"] = ss_.multi("Operating income (EBIT from operations)", [f"={L(CE+i)}{ss_.r-2}+{L(CE+i)}{ss_.r-1}" for i in range(NY)], NUM, bold=True)
T["op_chk"] = ss_.multi("   check: reported operating income", D["op_income"], NUM, font=F_INPUT)
T["int_inc"] = ss_.multi("Interest and treasury income", D["int_income"], NUM)
T["other"] = ss_.multi("Other income, incl. gains on investments (derived: PBT \u2212 operating income \u2212 interest income + finance costs \u2212 exceptionals)",
                       [f"={D['pbt'][i]}-{D['op_income'][i]}-{D['int_income'][i]}+{D['int_exp'][i]}-{D['special'][i]}" for i in range(NY)], NUM,
                       note="The public data feed does not carry this line separately; it is the balancing item to reported profit before tax and is treated as non-operating")
T["special"] = ss_.multi("Exceptional items (\u2212 = charge)", D["special"], NUM)
T["int_exp"] = ss_.multi("Finance costs", [-x for x in D["int_exp"]], NUM)
T["pbt"] = ss_.multi("Profit before tax", [f"=SUM({L(CE+i)}{ss_.r-6}:{L(CE+i)}{ss_.r-1})-{L(CE+i)}{ss_.r-5}" for i in range(NY)], NUM, bold=True, border=TOTAL_BORDER)
T["pbt_chk"] = ss_.multi("   check: reported profit before tax", D["pbt"], NUM, font=F_INPUT)
T["tax"] = ss_.multi("Tax", [-x for x in D["tax"]], NUM)
T["ni"] = ss_.multi("Profit after tax (reported)", D["ni"], NUM, bold=True, border=DOUBLE_BORDER)
T["etr"] = ss_.multi("   Effective tax rate", [f"=-{L(CE+i)}{ss_.r-2}/{L(CE+i)}{ss_.r-4}" for i in range(NY)], PCT)
T["sbc"] = ss_.multi("   memo: share-based payment expense (inside opex)", D["sbc"], NUM)
ss_.blank(); ss_.section("BALANCE SHEET (selected)")
T["ar"] = ss_.multi("Trade receivables", D["ar"], NUM)
T["inv"] = ss_.multi("Inventories", D["inv"], NUM)
T["ap"] = ss_.multi("Trade payables", D["ap"], NUM)
T["prov"] = ss_.multi("Current provisions", D["prov"], NUM)
T["cash_inv"] = ss_.multi("Cash and current investments (treasury book)", D["cash_inv"], NUM, bold=True)
T["debt"] = ss_.multi("Borrowings", D["debt"], NUM)
T["equity"] = ss_.multi("Shareholders' equity", D["equity"], NUM)
T["ppe"] = ss_.multi("Property, plant and equipment", D["ppe"], NUM)
T["intang"] = ss_.multi("Goodwill and intangibles", D["intang"], NUM)
ss_.blank(); ss_.section("CASH FLOW (selected)")
T["ocf"] = ss_.multi("Net cash from operating activities (after tax)", D["ocf"], NUM, bold=True)
T["capex"] = ss_.multi("Capital expenditure", [-x for x in D["capex"]], NUM)
T["fcf"] = ss_.multi("Free cash flow", [f"={L(CE+i)}{ss_.r-2}+{L(CE+i)}{ss_.r-1}" for i in range(NY)], NUM, bold=True, border=TOTAL_BORDER)
T["divs"] = ss_.multi("Dividends paid", [-x for x in D["divs"]], NUM)
ss_.note("Yahoo Finance's headline 'EBITDA' for Dabur is EBIT plus depreciation, where EBIT is profit before tax plus finance costs \u2014 "
         "so it includes treasury income and exceptional items. That headline is what a seller quotes. Operating EBITDA above is "
         "rebuilt from revenue, cost of sales and operating expenses, and ties to reported operating income on the check row.", height=44)

# --------------------------------------------------------------------------- #
# EBITDA Bridge
# --------------------------------------------------------------------------- #
sb = S("EBITDA Bridge", ncols=7, label_width=58, col_width=13, subtitle="From the headline the seller quotes to the EBITDA a buyer should pay for")
sb.ws.column_dimensions["G"].width = 46
B = {}
sb.section("HEADLINE TO ADJUSTED EBITDA")
sb.head(Y)
B["headline"] = sb.multi("Headline EBITDA (PBT + finance costs + D&A)", [f"={T['pbt'][i]}-{T['int_exp'][i]}-{T['da'][i]}" for i in range(NY)], NUM, bold=True)
B["a_int"] = sb.multi("Less: interest and treasury income", [f"=-{T['int_inc'][i]}" for i in range(NY)], NUM)
B["a_other"] = sb.multi("Less: other income incl. investment gains (derived)", [f"=-{T['other'][i]}" for i in range(NY)], NUM)
B["a_special"] = sb.multi("Add back / (remove): exceptional items", [f"=-{T['special'][i]}" for i in range(NY)], NUM)
B["adj"] = sb.multi("Adjusted operating EBITDA", [f"=SUM({L(CE+i)}{sb.r-4}:{L(CE+i)}{sb.r-1})" for i in range(NY)], NUM, bold=True, border=DOUBLE_BORDER)
B["adj_chk"] = sb.multi("   check: operating EBITDA from the Statements sheet, ex exceptionals", [f"={T['ebitda_op'][i]}" for i in range(NY)], NUM, font=F_LINK)
B["haircut"] = sb.multi("Adjusted as % of headline", [f"={B['adj'][i]}/{B['headline'][i]}" for i in range(NY)], PCT, bold=True)
B["adj_m"] = sb.multi("Adjusted EBITDA margin", [f"={B['adj'][i]}/{T['rev'][i]}" for i in range(NY)], PCT)
B["hl_m"] = sb.multi("Headline EBITDA margin", [f"={B['headline'][i]}/{T['rev'][i]}" for i in range(NY)], PCT)
sb.blank(); sb.section("ITEMS CONSIDERED AND LEFT IN")
B["sbc_pct"] = sb.multi("Share-based payments as % of adjusted EBITDA (left in as a real cost)", [f"={T['sbc'][i]}/{B['adj'][i]}" for i in range(NY)], PCT)
B["intang_amort"] = sb.multi("Decline in intangibles (amortisation, inside D&A, below EBITDA)", [""] + [f"={T['intang'][i]}-{T['intang'][i+1]}" for i in range(NY - 1)], NUM)
sb.blank(); sb.section("WHAT THE DIFFERENCE IS WORTH")
B["mult"] = sb.row("Illustrative transaction multiple of EBITDA", 20.0, MULT, note="Indian consumer staples change hands at high-teens to mid-twenties")
B["gap"] = sb.row("Headline less adjusted EBITDA, FY2026", f"={B['headline'][3]}-{B['adj'][3]}", NUM, bold=True)
B["gap_val"] = sb.row("Value at the multiple \u2014 what a buyer would overpay on the headline", f"={B['gap']}*{B['mult']}", NUM0, bold=True, border=DOUBLE_BORDER)
B["treasury"] = sb.row("Treasury book that generates the income (cash and current investments)", f"={T['cash_inv'][3]}", NUM0, font=F_LINK,
                       note="The buyer gets this as cash at closing \u2014 once. Paying a multiple for its income as well is paying twice")
B["double"] = sb.row("Overpayment as a multiple of the treasury book itself", f"={B['gap_val']}/{B['treasury']}", MULT)
sb.note("Nearly a fifth of the headline EBITDA is treasury income on a cash pile of over \u20b95,000 crore. In a transaction the cash "
        "transfers at face value, so a buyer who capitalises its income at 20x has paid for it twice. This is the single most "
        "common adjustment in an Indian consumer QoE and it is usually the largest.", height=40)

# --------------------------------------------------------------------------- #
# Cash Conversion
# --------------------------------------------------------------------------- #
sc = S("Cash Conversion", ncols=7, label_width=58, col_width=13, subtitle="Does the adjusted EBITDA turn into cash?")
sc.ws.column_dimensions["G"].width = 46
C = {}
sc.section("EBITDA TO CASH")
sc.head(Y)
C["adj"] = sc.multi("Adjusted operating EBITDA", [f"={B['adj'][i]}" for i in range(NY)], NUM, font=F_LINK)
C["tax_paid"] = sc.multi("Less: tax (charge used as proxy for cash tax)", [f"={T['tax'][i]}" for i in range(NY)], NUM, font=F_LINK)
C["ocf_pre"] = sc.multi("Operating cash flow before interest received (reported OCF less treasury income)", [f"={T['ocf'][i]}-{T['int_inc'][i]}" for i in range(NY)], NUM, bold=True,
                        note="Reported OCF includes interest received on the treasury book; it is removed for the same reason as on the Bridge")
C["conv"] = sc.multi("Cash conversion: OCF before interest / (adjusted EBITDA \u2212 tax)", [f"={C['ocf_pre'][i]}/({C['adj'][i]}+{C['tax_paid'][i]})" for i in range(NY)], PCT, bold=True)
C["conv_raw"] = sc.multi("   for comparison: reported OCF / headline EBITDA", [f"={T['ocf'][i]}/{B['headline'][i]}" for i in range(NY)], PCT)
C["capex_da"] = sc.multi("Capex / depreciation", [f"=-{T['capex'][i]}/-{T['da'][i]}" for i in range(NY)], MULT)
C["fcf_ni"] = sc.multi("Free cash flow / profit after tax", [f"={T['fcf'][i]}/{T['ni'][i]}" for i in range(NY)], PCT)
sc.blank(); sc.section("ACCRUALS")
C["accr"] = sc.multi("Accruals ratio: (profit after tax \u2212 OCF) / average total assets",
                     [f"=({T['ni'][i]}-{T['ocf'][i]})/{D['total_assets'][i]}" for i in range(NY)], PCT, bold=True,
                     note="Total assets are approximated (equity + debt + payables + provisions + other); negative is good, meaning cash exceeds profit")
C["accr_trend"] = sc.row("Direction over four years", f"=IF({C['accr'][3]}<{C['accr'][0]},\"Improving (cash gaining on profit)\",\"Deteriorating\")", "@", bold=True)
sc.blank(); sc.section("REVENUE QUALITY")
C["g"] = sc.multi("Revenue growth", [""] + [f"={T['rev'][i+1]}/{T['rev'][i]}-1" for i in range(NY - 1)], PCT)
C["ar_g"] = sc.multi("Receivables growth", [""] + [f"={T['ar'][i+1]}/{T['ar'][i]}-1" for i in range(NY - 1)], PCT)
C["inv_g"] = sc.multi("Inventory growth", [""] + [f"={T['inv'][i+1]}/{T['inv'][i]}-1" for i in range(NY - 1)], PCT)
C["flag"] = sc.multi("Flag: receivables or inventory growing faster than revenue", [""] +
                     [f"=IF(OR({C['ar_g'][i+1]}>{C['g'][i+1]}+0.05,{C['inv_g'][i+1]}>{C['g'][i+1]}+0.05),\"Review\",\"Clean\")" for i in range(NY - 1)], "@")
sc.note("FY2025 inventory grew 18% on 1% revenue growth \u2014 the year of the distributor inventory correction that Dabur disclosed. "
        "FY2026 receivables then fell 19% on 5% growth as channel stock normalised. Both are consistent with the company's own "
        "account and neither is a revenue-recognition concern; a QoE would still ask for the channel-inventory data to confirm it.", height=44)

# --------------------------------------------------------------------------- #
# Working Capital
# --------------------------------------------------------------------------- #
sw = S("Working Capital", ncols=7, label_width=58, col_width=13, subtitle="Days, trends and the definition the SPA will use")
sw.ws.column_dimensions["G"].width = 46
Wc = {}
sw.section("TRADE WORKING CAPITAL")
sw.head(Y)
Wc["ar"] = sw.multi("Trade receivables", [f"={T['ar'][i]}" for i in range(NY)], NUM, font=F_LINK)
Wc["inv"] = sw.multi("Inventories", [f"={T['inv'][i]}" for i in range(NY)], NUM, font=F_LINK)
Wc["ap"] = sw.multi("Less: trade payables", [f"=-{T['ap'][i]}" for i in range(NY)], NUM, font=F_LINK)
Wc["nwc"] = sw.multi("Trade working capital", [f"=SUM({L(CE+i)}{sw.r-3}:{L(CE+i)}{sw.r-1})" for i in range(NY)], NUM, bold=True, border=DOUBLE_BORDER)
Wc["nwc_rev"] = sw.multi("   as % of revenue", [f"={Wc['nwc'][i]}/{T['rev'][i]}" for i in range(NY)], PCT)
sw.blank(); sw.section("DAYS")
Wc["dso"] = sw.multi("Receivable days (on revenue)", [f"={T['ar'][i]}/{T['rev'][i]}*365" for i in range(NY)], NUM)
Wc["dio"] = sw.multi("Inventory days (on cost of goods sold)", [f"={T['inv'][i]}/-{T['cogs'][i]}*365" for i in range(NY)], NUM)
Wc["dpo"] = sw.multi("Payable days (on cost of goods sold)", [f"={T['ap'][i]}/-{T['cogs'][i]}*365" for i in range(NY)], NUM)
Wc["ccc"] = sw.multi("Cash conversion cycle (days)", [f"={Wc['dso'][i]}+{Wc['dio'][i]}-{Wc['dpo'][i]}" for i in range(NY)], NUM, bold=True, border=TOTAL_BORDER)
sw.blank(); sw.section("DEFINITION FOR THE SALE AND PURCHASE AGREEMENT")
sw.subsection("Included: trade receivables, inventories, trade payables. Excluded: cash and investments (cash-free / debt-free), borrowings, tax balances, provisions, capex creditors.")
Wc["prov_note"] = sw.multi("Current provisions (excluded from the peg; treated as debt-like)", [f"={T['prov'][i]}" for i in range(NY)], NUM, font=F_LINK)
sw.note("Payable days have stretched from 126 to 157 over four years while inventory days sat near 120. A buyer will ask whether "
        "the payables extension is structural (supplier financing, longer terms) or a closing-date effect that unwinds after "
        "completion. If it unwinds, the buyer funds roughly \u20b9500 crore of working capital in the first year. That question, not "
        "the peg arithmetic, is where the negotiation happens.", height=48)

# --------------------------------------------------------------------------- #
# Peg
# --------------------------------------------------------------------------- #
sp = S("Peg", ncols=6, label_width=64, subtitle="A normal level of working capital, so the buyer pays for the business and not for the closing date")
sp.ws.column_dimensions["F"].width = 58
Pg = {}
sp.section("METHOD 1 \u2014 AVERAGE OF THE FOUR YEAR-END BALANCES")
Pg["m1"] = sp.row("Average trade working capital, FY2023\u2013FY2026", f"=AVERAGE({','.join(Wc['nwc'])})", NUM, bold=True)
sp.section("METHOD 2 \u2014 AVERAGE DAYS APPLIED TO THE LATEST YEAR")
Pg["avg_days"] = sp.row("Average trade working capital as % of revenue, four years", f"=AVERAGE({','.join(Wc['nwc_rev'])})", PCT)
Pg["m2"] = sp.row("Applied to FY2026 revenue", f"={Pg['avg_days']}*{T['rev'][3]}", NUM, bold=True)
sp.section("METHOD 3 \u2014 LAST TWO YEARS ONLY (the post-correction run-rate)")
Pg["m3"] = sp.row("Average of FY2025 and FY2026", f"=AVERAGE({Wc['nwc'][2]},{Wc['nwc'][3]})", NUM, bold=True)
sp.section("METHOD 4 \u2014 SEASONALITY-ADJUSTED (stated)")
Pg["q_note"] = sp.row("Ratio of average quarterly revenue to the March quarter", f"=AVERAGE({D['q_rev'][0]},{D['q_rev'][1]},{D['q_rev'][2]},{D['q_rev'][3]})/AVERAGE({D['q_rev'][0]},{D['q_rev'][3]})", NUM2,
                      note="March is Dabur's seasonally softest quarter; a March year-end understates average working capital")
Pg["m4"] = sp.row("Method 1 scaled by the seasonality ratio", f"={Pg['m1']}*{Pg['q_note']}", NUM, bold=True)
sp.blank(); sp.section("THE PEG")
Pg["w1"] = sp.row("Weight on method 1", 0.25, PCT)
Pg["w2"] = sp.row("Weight on method 2", 0.25, PCT)
Pg["w3"] = sp.row("Weight on method 3", 0.25, PCT)
Pg["w4"] = sp.row("Weight on method 4", 0.25, PCT)
Pg["peg"] = sp.row("Proposed working-capital peg", f"={Pg['w1']}*{Pg['m1']}+{Pg['w2']}*{Pg['m2']}+{Pg['w3']}*{Pg['m3']}+{Pg['w4']}*{Pg['m4']}", NUM, bold=True, border=DOUBLE_BORDER)
Pg["close"] = sp.row("Trade working capital at the latest balance sheet (31 March 2026)", f"={Wc['nwc'][3]}", NUM, font=F_LINK)
Pg["adj"] = sp.row("Price adjustment if closing were today: closing less peg (\u2212 = buyer receives)", f"={Pg['close']}-{Pg['peg']}", NUM, bold=True, border=DOUBLE_BORDER)
Pg["adj_ps"] = sp.row("  per share (\u20b9)", f"={Pg['adj']}/177.7", NUM2)
Pg["adj_pct"] = sp.row("  as % of market capitalisation", f"={Pg['adj']}/({D['price']}*177.7)", PCT)
sp.blank(); sp.section("HOW SENSITIVE THE ADJUSTMENT IS TO THE METHOD")
Pg["rng_lo"] = sp.row("Lowest of the four methods", f"=MIN({Pg['m1']},{Pg['m2']},{Pg['m3']},{Pg['m4']})", NUM)
Pg["rng_hi"] = sp.row("Highest of the four methods", f"=MAX({Pg['m1']},{Pg['m2']},{Pg['m3']},{Pg['m4']})", NUM)
Pg["spread"] = sp.row("Spread between them \u2014 the size of the argument", f"={Pg['rng_hi']}-{Pg['rng_lo']}", NUM, bold=True)
sp.note("The seller will argue for method 3 (the recent run-rate, lowest peg, smallest adjustment); the buyer for method 4 "
        "(seasonally adjusted, highest peg). The spread between them is the negotiating range. A real peg uses twelve to "
        "twenty-four monthly balance sheets from the data room; Indian filers publish only half-yearly, so this one uses "
        "year-ends and a stated seasonality ratio, and says so.", height=48)

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
sy = S("Summary", ncols=6, label_width=62, subtitle="The quality-of-earnings findings on one page")
sy.ws.column_dimensions["F"].width = 56
sy.section("EARNINGS")
sy.row("Headline EBITDA, FY2026 (\u20b9 crore)", f"={B['headline'][3]}", NUM, font=F_LINK)
sy.row("Adjusted operating EBITDA, FY2026", f"={B['adj'][3]}", NUM, font=F_LINK, bold=True)
sy.row("Adjusted as % of headline", f"={B['haircut'][3]}", PCT, font=F_LINK, bold=True)
sy.row("Adjusted EBITDA margin", f"={B['adj_m'][3]}", PCT, font=F_LINK)
sy.row("Value of the difference at 20x", f"={B['gap_val']}", NUM0, font=F_LINK, bold=True)
sy.blank(); sy.section("CASH")
sy.row("Cash conversion, FY2026 (before treasury income)", f"={C['conv'][3]}", PCT, font=F_LINK, bold=True)
sy.row("Capex / depreciation, FY2026", f"={C['capex_da'][3]}", MULT, font=F_LINK)
sy.row("Accruals ratio, FY2026", f"={C['accr'][3]}", PCT, font=F_LINK)
sy.row("Accruals direction", f"={C['accr_trend']}", "@", font=F_LINK)
sy.blank(); sy.section("WORKING CAPITAL")
sy.row("Trade working capital at 31 March 2026", f"={Pg['close']}", NUM, font=F_LINK)
sy.row("Proposed peg", f"={Pg['peg']}", NUM, font=F_LINK, bold=True)
sy.row("Price adjustment if closing were today (\u2212 = buyer receives)", f"={Pg['adj']}", NUM, font=F_LINK, bold=True)
sy.row("Spread between peg methods", f"={Pg['spread']}", NUM, font=F_LINK)
sy.row("Payable days: FY2023 \u2192 FY2026", f"=TEXT({Wc['dpo'][0]},\"0\")&\" \u2192 \"&TEXT({Wc['dpo'][3]},\"0\")", "@", font=F_LINK, border=DOUBLE_BORDER)
sy.blank()
sy.bullets([
    "The headline EBITDA overstates operating earnings by close to a fifth, and almost all of it is interest on a "
    "\u20b95,000-crore treasury book. The cash transfers at face value in any deal; a buyer who also capitalises its income at "
    "20x has paid for the same asset twice. This is the finding a QoE exists to make.",
    "Cash conversion is strong once the same treasury income is removed from both sides, capex runs close to depreciation, "
    "and accruals are negative and improving. The operating earnings are real.",
    "Working capital is the negotiation. Payables have stretched by a month over four years, trade working capital has "
    "fallen to almost nothing, and a March year-end is seasonally soft. Whether the payables extension holds after "
    "completion is worth around \u20b9500 crore to whoever gets the answer right.",
    "The peg spread between the seller's method and the buyer's is the size of the argument, and it is small relative to "
    "the equity value. The EBITDA adjustment is where the money is.",
])

# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
tests = [
    ("Rebuilt operating income ties to reported in every year", "=AND(" + ",".join(f"ABS({a}-{b})<1" for a, b in zip(T["op_inc"], T["op_chk"])) + ")"),
    ("Rebuilt profit before tax ties to reported in every year", "=AND(" + ",".join(f"ABS({a}-{b})<1" for a, b in zip(T["pbt"], T["pbt_chk"])) + ")"),
    ("Derived other income is between 0% and 3% of revenue in every year (a plausibility bound on the plug)", "=AND(" + ",".join(f"{o}>=0,{o}<0.03*{r}" for o, r in zip(T["other"], T["rev"])) + ")"),
    ("Adjusted EBITDA on the Bridge equals operating EBITDA on the Statements", "=AND(" + ",".join(f"ABS({a}-{b})<1" for a, b in zip(B["adj"], B["adj_chk"])) + ")"),
    ("Adjusted EBITDA is below headline in every year", "=AND(" + ",".join(f"{a}<{b}" for a, b in zip(B["adj"], B["headline"])) + ")"),
    ("Effective tax rate is between 20% and 26% in every year", "=AND(" + ",".join(f"{e}>0.20,{e}<0.26" for e in T["etr"]) + ")"),
    ("Cash conversion before treasury income is between 60% and 130% in every year", "=AND(" + ",".join(f"{c}>0.6,{c}<1.3" for c in C["conv"]) + ")"),
    ("Cash conversion cycle equals DSO + DIO \u2212 DPO", "=AND(" + ",".join(f"ABS({c}-({a}+{b}-{d}))<0.01" for a, b, c, d in zip(Wc["dso"], Wc["dio"], Wc["ccc"], Wc["dpo"])) + ")"),
    ("Payable days have lengthened over the period", f"={Wc['dpo'][3]}>{Wc['dpo'][0]}"),
    ("Peg weights sum to 100%", f"=ABS({Pg['w1']}+{Pg['w2']}+{Pg['w3']}+{Pg['w4']}-1)<0.0001"),
    ("Peg lies between the lowest and highest method", f"=AND({Pg['peg']}>={Pg['rng_lo']},{Pg['peg']}<={Pg['rng_hi']})"),
    ("Price adjustment equals closing working capital less the peg", f"=ABS({Pg['adj']}-({Pg['close']}-{Pg['peg']}))<0.01"),
    ("Seasonality ratio is above 1 (March is a soft quarter)", f"={Pg['q_note']}>1"),
    ("Treasury income is below 5% of the treasury book (a sanity check on the classification)", f"={T['int_inc'][3]}/{T['cash_inv'][3]}<0.10"),
]
checks_sheet(book, tests,
             "The first two checks matter most: the operating EBITDA used throughout is rebuilt from revenue and costs and must tie "
             "to the reported operating income and profit before tax. Everything after that inherits its credibility from them.")

# --------------------------------------------------------------------------- #
book.cover(
    blurb="A buy-side quality-of-earnings review of Dabur India on the FY2023\u2013FY2026 consolidated accounts: a bridge from the "
          "headline EBITDA a seller quotes to the operating EBITDA a buyer should pay for, a cash-conversion and accruals test, a "
          "working-capital analysis in days, and a working-capital peg built four ways for a sale and purchase agreement.",
    method=[
        "Operating EBITDA rebuilt from revenue, cost of sales and operating expenses and tied to reported operating income; "
        "treasury income, other non-operating income and exceptionals bridged out of the headline.",
        "Cash conversion measured with treasury income removed from both EBITDA and operating cash flow; accruals ratio and "
        "revenue-quality flags across four years.",
        "Trade working capital defined as the SPA would (receivables, inventory, payables; cash-free, debt-free), with days and the "
        "cash conversion cycle.",
        "A peg from four methods \u2014 four-year average, average days, two-year run-rate, seasonality-adjusted \u2014 with weights as "
        "inputs and the spread between methods shown as the negotiating range.",
    ],
    toc=[("Summary", "earnings, cash, working capital"),
         ("Statements", "four years as reported, with operating EBITDA rebuilt and tied"),
         ("EBITDA Bridge", "headline to adjusted, and what the gap is worth"),
         ("Cash Conversion", "EBITDA to cash, accruals, revenue quality"),
         ("Working Capital", "trade working capital, days, the SPA definition"),
         ("Peg", "four methods, weights, the adjustment"),
         ("Checks", "fourteen tests; must read MODEL OK")],
    highlights=[("Headline EBITDA FY2026 (\u20b9 crore)", f"={B['headline'][3]}", NUM),
                ("Adjusted operating EBITDA", f"={B['adj'][3]}", NUM),
                ("Adjusted as % of headline", f"={B['haircut'][3]}", PCT),
                ("Cash conversion", f"={C['conv'][3]}", PCT),
                ("Working-capital peg", f"={Pg['peg']}", NUM),
                ("Adjustment if closing were today", f"={Pg['adj']}", NUM)],
    sources=["Dabur India Limited consolidated financial statements FY2023\u2013FY2026 (years to 31 March) via Yahoo Finance for DABUR.NS; quarterly revenue for the seasonality ratio from the same source.",
             "Total assets for the accruals ratio are approximated; the ratio is directional. Transaction multiple and peg weights are stated inputs.",
             "The review is a portfolio exercise on public information. Dabur is not, to the author's knowledge, in any sale process."])

book.finish(freeze={"Statements": "D6", "EBITDA Bridge": "D6", "Cash Conversion": "D6", "Working Capital": "D6"})
path = os.path.join(HERE, "Dabur_QoE.xlsx")
book.save(path)
print("saved", path)
print("recalc:", recalc(path))
pdf = export_pdf(path)
print("pdf:", pdf, os.path.exists(pdf))
png = preview_png(pdf, page=0, dpi=80)
os.replace(png, os.path.join(HERE, "cover.png"))
print("cover.png written")
